#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
达梦数据库结构差异对比工具
用于对比两个达梦数据库之间的表结构与索引差异，支持任意环境方向的同步
"""

import configparser
import fnmatch
import logging
import os
import re
import socket
import sys
import time
from datetime import datetime
from typing import Callable, Dict, List, Set, Tuple, Optional, Any
import dataclasses

logger = logging.getLogger('dm_compare')

NO_ARGUMENT_DATA_TYPES = {
    'BIGINT', 'INTEGER', 'INT', 'SMALLINT', 'TINYINT', 'BYTE',
    'DATE', 'CLOB', 'BLOB', 'TEXT', 'IMAGE'
}
PRECISION_DATA_TYPES = {'NUMBER', 'NUMERIC', 'DECIMAL', 'DEC', 'FLOAT'}
SCALE_DATA_TYPES = {'NUMBER', 'NUMERIC', 'DECIMAL', 'DEC'}
LENGTH_DATA_TYPES = {
    'CHAR', 'CHARACTER', 'VARCHAR', 'VARCHAR2',
    'NCHAR', 'NVARCHAR', 'NVARCHAR2',
    'BINARY', 'VARBINARY', 'RAW'
}


def _normalize_dm_data_type(data_type: Optional[str]) -> str:
    """Normalize equivalent Dameng type names for comparison and DDL."""
    normalized = (data_type or '').strip().upper()
    if normalized == 'VARCHAR2':
        return 'VARCHAR'
    if normalized == 'DATETIME':
        return 'TIMESTAMP'
    if normalized == 'INT':
        return 'INTEGER'
    return normalized


def _format_dm_column_type(data_type: Optional[str], length: Optional[int],
                           precision: Optional[int], scale: Optional[int],
                           normalize_aliases: bool = False) -> str:
    """Format a Dameng column type without adding invalid length modifiers."""
    source_type = (data_type or '').strip().upper()
    normalized_type = _normalize_dm_data_type(source_type)
    ddl_type = normalized_type if normalize_aliases else source_type

    if normalized_type in NO_ARGUMENT_DATA_TYPES:
        return ddl_type

    if precision is not None and normalized_type in PRECISION_DATA_TYPES:
        if scale is not None and normalized_type in SCALE_DATA_TYPES:
            return f"{ddl_type}({precision},{scale})"
        return f"{ddl_type}({precision})"

    if length and normalized_type in LENGTH_DATA_TYPES:
        return f"{ddl_type}({length})"

    return ddl_type

try:
    import dmPython
except ImportError:
    print("错误: 未安装 dmPython 驱动")
    print("请运行: pip install dmPython")
    sys.exit(1)

try:
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter as _openpyxl_get_column_letter
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    print("错误: 未安装 openpyxl")
    print("请运行: pip install openpyxl")
    sys.exit(1)


@dataclasses.dataclass
class ColumnInfo:
    """字段信息"""
    name: str
    data_type: str
    length: Optional[int]
    precision: Optional[int]
    scale: Optional[int]
    nullable: str
    default: Optional[str]
    comment: Optional[str] = ""
    is_primary_key: bool = False
    is_identity: bool = False
    identity_start: Optional[int] = None
    identity_increment: Optional[int] = None

    def __hash__(self):
        return hash(self.name)

    def __eq__(self, other):
        if not isinstance(other, ColumnInfo):
            return False
        return self.name == other.name

    def get_type_string(self) -> str:
        """获取类型字符串表示"""
        return _format_dm_column_type(
            self.data_type,
            self.length,
            self.precision,
            self.scale,
            normalize_aliases=False
        )


@dataclasses.dataclass
class IndexColumnInfo:
    """索引列信息"""
    name: str
    position: int
    sort_order: str = 'ASC'

    def get_display_string(self) -> str:
        """获取列展示字符串"""
        if (self.sort_order or '').upper() == 'DESC':
            return f"{self.name} DESC"
        return self.name


@dataclasses.dataclass
class IndexInfo:
    """索引信息"""
    name: str
    schema: str
    table_name: str
    uniqueness: Optional[str]
    index_type: Optional[str] = ""
    constraint_name: Optional[str] = None
    constraint_type: Optional[str] = None
    columns: List[IndexColumnInfo] = dataclasses.field(default_factory=list)

    def __hash__(self):
        return hash((self.schema, self.table_name, self.name))

    def __eq__(self, other):
        if not isinstance(other, IndexInfo):
            return False
        return (
            self.schema == other.schema
            and self.table_name == other.table_name
            and self.name == other.name
        )

    def get_columns_string(self) -> str:
        """获取索引列定义字符串"""
        ordered_columns = sorted(self.columns, key=lambda col: (col.position, col.name))
        return ', '.join(column.get_display_string() for column in ordered_columns)

    def get_definition_string(self) -> str:
        """获取索引定义字符串"""
        parts = []
        if self.uniqueness:
            parts.append(self.uniqueness)
        if self.index_type:
            parts.append(self.index_type)

        columns_str = self.get_columns_string()
        if columns_str:
            parts.append(f"({columns_str})")

        return ' '.join(parts) if parts else '-'

    def is_constraint_backed(self) -> bool:
        """是否为主键或唯一约束背后的索引。"""
        return (self.constraint_type or '').upper() in {'P', 'U'}


@dataclasses.dataclass
class SequenceInfo:
    """序列信息。schema 级对象，不挂在某张表下。"""
    name: str
    schema: str
    min_value: Optional[int]
    max_value: Optional[int]
    increment_by: Optional[int]
    cycle_flag: Optional[str]   # 'Y' / 'N'
    order_flag: Optional[str]   # 'Y' / 'N'
    cache_size: Optional[int]   # 0 表示 NOCACHE
    start_with: Optional[int] = None  # 建序列起始值，仅存在于 DDL

    def __hash__(self):
        return hash((self.schema, self.name))

    def __eq__(self, other):
        if not isinstance(other, SequenceInfo):
            return False
        return self.schema == other.schema and self.name == other.name

    def get_attribute_string(self) -> str:
        """紧凑属性串，用于对比展示。"""
        cache = 'NOCACHE' if (self.cache_size or 0) == 0 else f'CACHE {self.cache_size}'
        return (
            f"START_WITH={self.start_with}, "
            f"INCREMENT_BY={self.increment_by}, "
            f"MIN_VALUE={self.min_value}, "
            f"MAX_VALUE={self.max_value}, "
            f"CYCLE={self.cycle_flag}, "
            f"{cache}, "
            f"ORDER={self.order_flag}"
        )


@dataclasses.dataclass
class SchemaConnectionConfig:
    """单个模式的数据库连接配置"""
    schema: str
    user: str
    password: str
    host: str
    port: int


@dataclasses.dataclass
class ForeignKeyColumn:
    """外键列映射"""
    column_name: str
    position: int
    ref_column_name: str


@dataclasses.dataclass
class ForeignKeyInfo:
    """外键约束信息"""
    name: str
    schema: str
    table_name: str
    ref_schema: str
    ref_table_name: str
    delete_rule: Optional[str] = ""
    columns: List[ForeignKeyColumn] = dataclasses.field(default_factory=list)

    def __hash__(self):
        return hash((self.schema, self.table_name, self.name))

    def __eq__(self, other):
        if not isinstance(other, ForeignKeyInfo):
            return False
        return (
            self.schema == other.schema
            and self.table_name == other.table_name
            and self.name == other.name
        )

    def get_columns_string(self) -> str:
        ordered = sorted(self.columns, key=lambda c: (c.position, c.column_name))
        return ', '.join(c.column_name for c in ordered)

    def get_ref_columns_string(self) -> str:
        ordered = sorted(self.columns, key=lambda c: (c.position, c.column_name))
        return ', '.join(c.ref_column_name for c in ordered)

    def get_definition_string(self) -> str:
        cols = self.get_columns_string()
        ref_cols = self.get_ref_columns_string()
        ref = f"{self.ref_schema}.{self.ref_table_name}"
        parts = [f"FOREIGN KEY ({cols}) REFERENCES {ref} ({ref_cols})"]
        if self.delete_rule and self.delete_rule != 'NO ACTION':
            parts.append(f"ON DELETE {self.delete_rule}")
        return ' '.join(parts)


@dataclasses.dataclass
class ViewInfo:
    """视图信息"""
    name: str
    schema: str
    text: str = ""

    def __hash__(self):
        return hash((self.schema, self.name))

    def __eq__(self, other):
        if not isinstance(other, ViewInfo):
            return False
        return self.schema == other.schema and self.name == other.name


@dataclasses.dataclass
class ProcedureInfo:
    """存储过程/函数信息"""
    name: str
    schema: str
    object_type: str  # 'PROCEDURE' / 'FUNCTION' / 'TRIGGER'
    text: str = ""

    def __hash__(self):
        return hash((self.schema, self.name, self.object_type))

    def __eq__(self, other):
        if not isinstance(other, ProcedureInfo):
            return False
        return (
            self.schema == other.schema
            and self.name == other.name
            and self.object_type == other.object_type
        )


@dataclasses.dataclass
class TableInfo:
    """表信息"""
    name: str
    schema: str
    table_type: str
    comment: Optional[str] = ""
    columns: Dict[str, ColumnInfo] = dataclasses.field(default_factory=dict)
    indexes: Dict[str, IndexInfo] = dataclasses.field(default_factory=dict)
    primary_key_columns: List[str] = dataclasses.field(default_factory=list)
    foreign_keys: Dict[str, ForeignKeyInfo] = dataclasses.field(default_factory=dict)

    def __hash__(self):
        return hash((self.schema, self.name))

    def __eq__(self, other):
        if not isinstance(other, TableInfo):
            return False
        return self.schema == other.schema and self.name == other.name


@dataclasses.dataclass
class SchemaInfo:
    """模式信息"""
    name: str
    tables: Dict[str, TableInfo] = dataclasses.field(default_factory=dict)
    sequences: Dict[str, SequenceInfo] = dataclasses.field(default_factory=dict)
    views: Dict[str, ViewInfo] = dataclasses.field(default_factory=dict)
    procedures: Dict[str, ProcedureInfo] = dataclasses.field(default_factory=dict)


class DMConnection:
    """达梦数据库连接封装"""

    # 连接超时秒数（覆盖 TCP 建连 + 达梦登录握手）
    CONNECT_TIMEOUT = 30

    def __init__(self, host: str, port: int, user: str, password: str):
        self.host = host
        self.port = port
        self.user = user
        self.password = password
        self.conn = None

    def connect(self) -> bool:
        """连接数据库。通过 socket 默认超时兜底，避免网络不通时无限等待。"""
        old_timeout = socket.getdefaulttimeout()
        socket.setdefaulttimeout(self.CONNECT_TIMEOUT)
        try:
            self.conn = dmPython.connect(
                user=self.user,
                password=self.password,
                server=self.host,
                port=self.port
            )
        except Exception as e:
            logger.error("连接失败 %s:%s - %s: %s", self.host, self.port, type(e).__name__, e)
            return False
        finally:
            socket.setdefaulttimeout(old_timeout)

        logger.info("  连接成功 (%s:%s)", self.host, self.port)
        self.password = '***'
        return True

    def is_connected(self) -> bool:
        """检查是否已连接"""
        return self.conn is not None

    def close(self):
        """关闭连接"""
        if self.conn:
            self.conn.close()
            self.conn = None

    def execute_query(self, sql: str, params: tuple = None) -> List[tuple]:
        """执行查询并返回结果"""
        # 格式化SQL用于日志（压缩空白）
        log_sql = ' '.join(sql.split())
        if len(log_sql) > 200:
            log_sql = log_sql[:200] + '...'
        logger.debug("执行SQL: %s  参数: %s", log_sql, params)

        start_time = time.time()
        cursor = self.conn.cursor()
        try:
            if params:
                cursor.execute(sql, params)
            else:
                cursor.execute(sql)
            rows = cursor.fetchall()
            elapsed = time.time() - start_time
            logger.debug("SQL完成: 返回%d行, 耗时%.3fs", len(rows), elapsed)
            return rows
        except Exception as e:
            elapsed = time.time() - start_time
            logger.error("SQL执行失败(%.3fs): %s  SQL: %s", elapsed, e, log_sql)
            raise
        finally:
            cursor.close()


class DMMetadataExtractor:
    """达梦数据库元数据提取器"""

    def __init__(self, conn: DMConnection):
        self.conn = conn
        self._schemas_cache: Optional[List[str]] = None

    def get_all_schemas(self) -> List[str]:
        """获取所有模式"""
        if self._schemas_cache is not None:
            return self._schemas_cache

        # 使用 ALL_USERS 代替 DBA_USERS，不需要DBA权限
        sql = "SELECT USERNAME FROM ALL_USERS ORDER BY USERNAME"
        result = self.conn.execute_query(sql)
        self._schemas_cache = [row[0] for row in result]
        return self._schemas_cache

    def get_all_tables(self, schema: str) -> Dict[str, TableInfo]:
        """获取指定模式下的所有表"""
        # 使用 ALL_TABLES 代替 DBA_TABLES，不需要DBA权限
        # ALL_TABLES 没有 TABLE_TYPE 字段，统一设置为 'TABLE'
        sql = """
            SELECT T.TABLE_NAME, C.COMMENTS
            FROM ALL_TABLES T
            LEFT JOIN ALL_TAB_COMMENTS C
                ON T.TABLE_NAME = C.TABLE_NAME AND T.OWNER = C.OWNER
            WHERE T.OWNER = ?
            ORDER BY T.TABLE_NAME
        """
        result = self.conn.execute_query(sql, (schema,))
        tables = {}
        for row in result:
            table_name, comment = row
            tables[table_name] = TableInfo(
                name=table_name,
                schema=schema,
                table_type='TABLE',  # 统一设置为 TABLE
                comment=comment or ""
            )
        return tables

    def get_schema_columns(self, schema: str) -> Dict[str, Dict[str, ColumnInfo]]:
        """一次查询获取模式下所有表的所有字段，按表名分组。

        批量替代逐表查询，消除 N+1 网络往返；与 get_schema_indexes 等 schema 级
        查询保持一致的批量风格。dict 保持插入顺序，列按 COLUMN_ID 排序。
        """
        sql = """
            SELECT
                T.TABLE_NAME,
                T.COLUMN_NAME,
                T.DATA_TYPE,
                T.DATA_LENGTH,
                T.DATA_PRECISION,
                T.DATA_SCALE,
                T.NULLABLE,
                T.DATA_DEFAULT,
                C.COMMENTS
            FROM ALL_TAB_COLUMNS T
            LEFT JOIN ALL_COL_COMMENTS C
                ON T.COLUMN_NAME = C.COLUMN_NAME AND T.OWNER = C.OWNER
                AND T.TABLE_NAME = C.TABLE_NAME
            WHERE T.OWNER = ?
            ORDER BY T.TABLE_NAME, T.COLUMN_ID
        """
        result = self.conn.execute_query(sql, (schema,))
        table_columns: Dict[str, Dict[str, ColumnInfo]] = {}
        for row in result:
            table_name, col_name, data_type, length, precision, scale, nullable, default, comment = row
            table_columns.setdefault(table_name, {})[col_name] = ColumnInfo(
                name=col_name,
                data_type=data_type,
                length=length,
                precision=precision,
                scale=scale,
                nullable=nullable,
                default=str(default) if default else None,
                comment=comment or ""
            )
        return table_columns

    def get_schema_indexes(self, schema: str) -> Dict[str, Dict[str, IndexInfo]]:
        """获取指定模式下所有表的索引信息"""
        sql = """
            SELECT
                C.TABLE_NAME,
                C.INDEX_NAME,
                I.UNIQUENESS,
                I.INDEX_TYPE,
                AC.CONSTRAINT_NAME,
                AC.CONSTRAINT_TYPE,
                C.COLUMN_NAME,
                C.COLUMN_POSITION,
                CASE C.DESCEND
                    WHEN 'Y' THEN 'DESC'
                    WHEN 'DESC' THEN 'DESC'
                    ELSE 'ASC'
                END AS SORT_ORDER
            FROM ALL_IND_COLUMNS C
            JOIN ALL_INDEXES I
                ON I.OWNER = C.INDEX_OWNER
                AND I.INDEX_NAME = C.INDEX_NAME
                AND I.TABLE_OWNER = C.TABLE_OWNER
                AND I.TABLE_NAME = C.TABLE_NAME
            LEFT JOIN ALL_CONSTRAINTS AC
                ON AC.OWNER = I.TABLE_OWNER
                AND AC.TABLE_NAME = I.TABLE_NAME
                AND AC.INDEX_OWNER = I.OWNER
                AND AC.INDEX_NAME = I.INDEX_NAME
                AND AC.CONSTRAINT_TYPE IN ('P', 'U')
            WHERE C.TABLE_OWNER = ?
            ORDER BY C.TABLE_NAME, C.INDEX_NAME, C.COLUMN_POSITION
        """
        result = self.conn.execute_query(sql, (schema,))
        table_indexes: Dict[str, Dict[str, IndexInfo]] = {}

        for row in result:
            (
                table_name,
                index_name,
                uniqueness,
                index_type,
                constraint_name,
                constraint_type,
                column_name,
                column_position,
                sort_order
            ) = row
            indexes = table_indexes.setdefault(table_name, {})

            if index_name not in indexes:
                indexes[index_name] = IndexInfo(
                    name=index_name,
                    schema=schema,
                    table_name=table_name,
                    uniqueness=uniqueness,
                    index_type=index_type or "",
                    constraint_name=constraint_name,
                    constraint_type=constraint_type,
                )

            indexes[index_name].columns.append(IndexColumnInfo(
                name=column_name or "<表达式>",
                position=column_position or len(indexes[index_name].columns) + 1,
                sort_order=sort_order or 'ASC'
            ))

        return table_indexes

    def get_schema_primary_key_columns(self, schema: str) -> Dict[str, List[str]]:
        """获取指定模式下所有主键列，按 POSITION 保留源端定义顺序。"""
        sql = """
            SELECT C.TABLE_NAME, CC.COLUMN_NAME
            FROM ALL_CONSTRAINTS C
            JOIN ALL_CONS_COLUMNS CC
                ON C.OWNER = CC.OWNER
                AND C.CONSTRAINT_NAME = CC.CONSTRAINT_NAME
                AND C.TABLE_NAME = CC.TABLE_NAME
            WHERE C.OWNER = ? AND C.CONSTRAINT_TYPE = 'P'
            ORDER BY C.TABLE_NAME, CC.POSITION
        """
        result = self.conn.execute_query(sql, (schema,))
        table_columns: Dict[str, List[str]] = {}

        for table_name, column_name in result:
            if table_name and column_name:
                lst = table_columns.setdefault(table_name, [])
                if column_name not in lst:
                    lst.append(column_name)

        return table_columns

    def get_schema_foreign_keys(self, schema: str) -> Dict[str, Dict[str, ForeignKeyInfo]]:
        """获取指定模式下所有外键约束，按表名分组。"""
        sql = """
            SELECT
                C.TABLE_NAME,
                C.CONSTRAINT_NAME,
                R.OWNER       AS REF_OWNER,
                R.TABLE_NAME  AS REF_TABLE_NAME,
                C.DELETE_RULE,
                CC.COLUMN_NAME,
                CC.POSITION,
                RCC.COLUMN_NAME AS REF_COLUMN_NAME
            FROM ALL_CONSTRAINTS C
            JOIN ALL_CONS_COLUMNS CC
                ON C.OWNER = CC.OWNER
                AND C.CONSTRAINT_NAME = CC.CONSTRAINT_NAME
                AND C.TABLE_NAME = CC.TABLE_NAME
            JOIN ALL_CONSTRAINTS R
                ON C.R_OWNER = R.OWNER
                AND C.R_CONSTRAINT_NAME = R.CONSTRAINT_NAME
            JOIN ALL_CONS_COLUMNS RCC
                ON R.OWNER = RCC.OWNER
                AND R.CONSTRAINT_NAME = RCC.CONSTRAINT_NAME
                AND CC.POSITION = RCC.POSITION
            WHERE C.OWNER = ? AND C.CONSTRAINT_TYPE = 'R'
            ORDER BY C.TABLE_NAME, C.CONSTRAINT_NAME, CC.POSITION
        """
        result = self.conn.execute_query(sql, (schema,))
        table_fks: Dict[str, Dict[str, ForeignKeyInfo]] = {}

        for row in result:
            (
                table_name, constraint_name,
                ref_owner, ref_table, delete_rule,
                col_name, position, ref_col_name
            ) = row
            fks = table_fks.setdefault(table_name, {})

            if constraint_name not in fks:
                fks[constraint_name] = ForeignKeyInfo(
                    name=constraint_name,
                    schema=schema,
                    table_name=table_name,
                    ref_schema=ref_owner or '',
                    ref_table_name=ref_table or '',
                    delete_rule=(delete_rule or '').upper(),
                )

            fks[constraint_name].columns.append(ForeignKeyColumn(
                column_name=col_name or '',
                position=position or len(fks[constraint_name].columns) + 1,
                ref_column_name=ref_col_name or '',
            ))

        return table_fks

    def get_schema_views(self, schema: str) -> Dict[str, ViewInfo]:
        """获取指定模式下所有视图及其定义文本。"""
        sql = """
            SELECT VIEW_NAME, TEXT
            FROM ALL_VIEWS
            WHERE OWNER = ?
            ORDER BY VIEW_NAME
        """
        result = self.conn.execute_query(sql, (schema,))
        views: Dict[str, ViewInfo] = {}
        for name, text in result:
            if hasattr(text, 'read'):
                text = text.read()
            views[name] = ViewInfo(name=name, schema=schema, text=str(text or ''))
        return views

    def get_schema_procedures(self, schema: str) -> Dict[str, ProcedureInfo]:
        """获取指定模式下所有存储过程、函数、触发器及其源码。

        达梦 ALL_SOURCE 按 LINE 分行存储源码，需要拼接。
        """
        sql = """
            SELECT NAME, TYPE, LINE, TEXT
            FROM ALL_SOURCE
            WHERE OWNER = ?
            ORDER BY TYPE, NAME, LINE
        """
        result = self.conn.execute_query(sql, (schema,))
        procs: Dict[str, ProcedureInfo] = {}
        for name, obj_type, _line, text in result:
            if hasattr(text, 'read'):
                text = text.read()
            key = f"{obj_type}::{name}"
            if key not in procs:
                procs[key] = ProcedureInfo(
                    name=name,
                    schema=schema,
                    object_type=(obj_type or '').upper(),
                    text='',
                )
            procs[key].text += str(text or '')
        return procs

    def get_schema_sequences(self, schema: str,
                             exclude_patterns: Optional[List[str]] = None) -> Dict[str, SequenceInfo]:
        """一次查询获取模式下所有序列，再逐个解析 START WITH。

        ALL_SEQUENCES 不存 START WITH（只存 LAST_NUMBER 当前值），起始值只在
        DDL 里，必须 DBMS_METADATA.GET_DDL 取。序列数通常几十个，N+1 可接受，
        与 get_table_identity_params 按对象调 GET_DDL 的风格一致。
        exclude_patterns 命名的序列跳过（含 GET_DDL），用于挡框架元数据噪音。
        """
        sql = """
            SELECT SEQUENCE_NAME, MIN_VALUE, MAX_VALUE, INCREMENT_BY,
                   CYCLE_FLAG, ORDER_FLAG, CACHE_SIZE, LAST_NUMBER
            FROM ALL_SEQUENCES
            WHERE SEQUENCE_OWNER = ?
            ORDER BY SEQUENCE_NAME
        """
        result = self.conn.execute_query(sql, (schema,))
        sequences: Dict[str, SequenceInfo] = {}
        for row in result:
            name, min_v, max_v, incr, cycle, order, cache, _last = row
            if not name:
                continue
            if _is_name_excluded(name, exclude_patterns):
                continue
            sequences[name] = SequenceInfo(
                name=name,
                schema=schema,
                min_value=min_v,
                max_value=max_v,
                increment_by=incr,
                cycle_flag=cycle,
                order_flag=order,
                cache_size=cache,
            )

        for name, seq in sequences.items():
            seq.start_with = self._get_sequence_start_with(schema, name)

        return sequences

    def _get_sequence_start_with(self, schema: str, sequence_name: str) -> Optional[int]:
        """从建序列 DDL 解析 START WITH。ALL_SEQUENCES 不暴露此属性。"""
        sql = "SELECT DBMS_METADATA.GET_DDL('SEQUENCE', ?, ?) FROM DUAL"
        result = self.conn.execute_query(sql, (sequence_name, schema))
        if not result:
            return None

        ddl = result[0][0]
        if hasattr(ddl, 'read'):
            ddl = ddl.read()
        ddl_text = str(ddl) or ''

        match = re.search(r'START\s+WITH\s+(-?\d+)', ddl_text, re.IGNORECASE)
        return int(match.group(1)) if match else None

    @staticmethod
    def _mark_primary_key_columns(columns: Dict[str, ColumnInfo], pk_columns: List[str]):
        """标记字段是否属于主键。"""
        if not columns or not pk_columns:
            return

        column_map = {name.upper(): column for name, column in columns.items()}
        for pk_column_name in pk_columns:
            column = column_map.get(pk_column_name.upper())
            if column:
                column.is_primary_key = True

    def get_schema_identity_flags(self, schema: str) -> Dict[str, Set[str]]:
        """识别模式下所有 IDENTITY 自增列。

        达梦 SYSCOLUMNS.INFO2=1 标识物理表的 IDENTITY 列；INFO1 不存此信息。
        INFO2 语义随对象类型变化（系统视图列可能为其他值），仅物理表视为 IDENTITY。
        """
        sql = """
            SELECT O.NAME AS TABLE_NAME, C.NAME AS COLUMN_NAME
            FROM SYS.SYSCOLUMNS C
            JOIN SYS.SYSOBJECTS O ON C.ID = O.ID
            JOIN SYS.SYSOBJECTS S ON O.SCHID = S.ID
            WHERE C.INFO2 = 1 AND S.NAME = ?
        """
        result = self.conn.execute_query(sql, (schema,))
        flags: Dict[str, Set[str]] = {}
        for table_name, column_name in result:
            if table_name and column_name:
                # 表名/列名统一大写，规避 ALL_TABLES 与 SYSCOLUMNS 间的大小写差异
                flags.setdefault(table_name.upper(), set()).add(column_name.upper())
        return flags

    def get_table_identity_params(self, schema: str, table_name: str) -> Dict[str, Tuple[int, int]]:
        """解析建表 DDL，提取每个 IDENTITY 列的 (start, increment)。

        SYSCOLUMNS 不存 start/increment，达梦也无 DBA_TAB_IDENTITY_COLS 视图，
        唯一精确来源是 DBMS_METADATA.GET_DDL。调用方应保证只在含 IDENTITY 列的表上调用。
        """
        sql = "SELECT DBMS_METADATA.GET_DDL('TABLE', ?, ?) FROM DUAL"
        result = self.conn.execute_query(sql, (table_name, schema))
        if not result:
            return {}

        ddl = result[0][0]
        if hasattr(ddl, 'read'):
            ddl = ddl.read()
        ddl_text = str(ddl) or ''

        params: Dict[str, Tuple[int, int]] = {}
        pattern = re.compile(
            r'"(?P<col>[^"]+)"\s+\w+(?:\s*\([^)]*\))?\s+IDENTITY\s*\(\s*(?P<start>-?\d+)\s*,\s*(?P<incr>-?\d+)\s*\)',
            re.IGNORECASE,
        )
        for match in pattern.finditer(ddl_text):
            params[match.group('col').upper()] = (int(match.group('start')), int(match.group('incr')))
        return params

    def _mark_identity_columns(self, schema: str, table_name: str,
                               columns: Dict[str, ColumnInfo],
                               identity_flags: Dict[str, Set[str]]):
        """标记 IDENTITY 列并填充 (start, increment)。同表只调一次 GET_DDL。"""
        identity_cols = identity_flags.get(table_name.upper())
        if not identity_cols or not columns:
            return

        column_map = {name.upper(): column for name, column in columns.items()}
        for col_name in identity_cols:
            column = column_map.get(col_name.upper())
            if column:
                column.is_identity = True

        params = self.get_table_identity_params(schema, table_name)
        for col_name in identity_cols:
            column = column_map.get(col_name.upper())
            if column and column.is_identity:
                start, incr = params.get(col_name.upper(), (1, 1))
                column.identity_start = start
                column.identity_increment = incr

    def extract_all_metadata(self, exclude_schemas: Set[str] = None,
                             exclude_sequence_patterns: Optional[List[str]] = None) -> Dict[str, SchemaInfo]:
        """提取所有元数据"""
        if exclude_schemas is None:
            exclude_schemas = set()

        logger.info("  正在提取模式列表...")
        schemas = self.get_all_schemas()

        metadata = {}
        total = len(schemas)
        for i, schema in enumerate(schemas, 1):
            if schema in exclude_schemas:
                continue

            logger.info("  [%d/%d] 提取模式 %s...", i, total, schema)
            schema_info = SchemaInfo(name=schema)
            schema_info.tables = self.get_all_tables(schema)
            schema_indexes = self.get_schema_indexes(schema)
            schema_pk_columns = self.get_schema_primary_key_columns(schema)
            schema_fks = self.get_schema_foreign_keys(schema)
            identity_flags = self.get_schema_identity_flags(schema)
            schema_info.sequences = self.get_schema_sequences(schema, exclude_sequence_patterns)
            schema_info.views = self.get_schema_views(schema)
            schema_info.procedures = self.get_schema_procedures(schema)

            # 一次性批量提取字段，再按表分发，消除逐表 N+1 查询
            schema_columns = self.get_schema_columns(schema)
            for j, (table_name, table_info) in enumerate(schema_info.tables.items(), 1):
                logger.debug("    [%d/%d] 表 %s", j, len(schema_info.tables), table_name)
                table_info.columns = schema_columns.get(table_name, {})
                table_info.indexes = schema_indexes.get(table_name, {})
                table_info.primary_key_columns = schema_pk_columns.get(table_name, [])
                table_info.foreign_keys = schema_fks.get(table_name, {})
                self._mark_primary_key_columns(table_info.columns, table_info.primary_key_columns)
                self._mark_identity_columns(schema, table_name, table_info.columns, identity_flags)

            metadata[schema] = schema_info

        return metadata

    def extract_single_schema(self, schema_name: str,
                              exclude_sequence_patterns: Optional[List[str]] = None) -> SchemaInfo:
        """只提取指定模式的元数据（用于多连接场景）"""
        logger.info("  提取模式 %s...", schema_name)
        schema_info = SchemaInfo(name=schema_name)
        schema_info.tables = self.get_all_tables(schema_name)
        schema_indexes = self.get_schema_indexes(schema_name)
        schema_pk_columns = self.get_schema_primary_key_columns(schema_name)
        schema_fks = self.get_schema_foreign_keys(schema_name)
        identity_flags = self.get_schema_identity_flags(schema_name)
        schema_info.sequences = self.get_schema_sequences(schema_name, exclude_sequence_patterns)
        schema_info.views = self.get_schema_views(schema_name)
        schema_info.procedures = self.get_schema_procedures(schema_name)

        # 一次性批量提取字段，再按表分发，消除逐表 N+1 查询
        schema_columns = self.get_schema_columns(schema_name)
        for j, (table_name, table_info) in enumerate(schema_info.tables.items(), 1):
            logger.debug("    [%d/%d] 表 %s", j, len(schema_info.tables), table_name)
            table_info.columns = schema_columns.get(table_name, {})
            table_info.indexes = schema_indexes.get(table_name, {})
            table_info.primary_key_columns = schema_pk_columns.get(table_name, [])
            table_info.foreign_keys = schema_fks.get(table_name, {})
            self._mark_primary_key_columns(table_info.columns, table_info.primary_key_columns)
            self._mark_identity_columns(schema_name, table_name, table_info.columns, identity_flags)

        return schema_info


@dataclasses.dataclass
class ColumnDiff:
    """字段差异"""
    schema: str
    table: str
    column: str
    diff_type: str  # '源端缺失', '目标端缺失', '类型变更', '长度变更', '精度变更', '可空性变更', '默认值变更', '注释变更'
    test_value: str
    prod_value: str
    test_column: Optional[ColumnInfo] = None
    prod_column: Optional[ColumnInfo] = None


@dataclasses.dataclass
class TableDiff:
    """表差异"""
    schema: str
    table: str
    diff_type: str  # '目标端缺失', '源端缺失'
    comment: str = ""
    test_table: Optional[TableInfo] = None
    prod_table: Optional[TableInfo] = None


@dataclasses.dataclass
class IndexDiff:
    """索引差异"""
    schema: str
    table: str
    index: str
    diff_type: str  # '目标端缺失', '源端缺失', '唯一性变更', '索引类型变更', '索引列变更'
    test_value: str
    prod_value: str
    test_index: Optional[IndexInfo] = None
    prod_index: Optional[IndexInfo] = None


@dataclasses.dataclass
class PrimaryKeyDiff:
    """主键差异"""
    schema: str
    table: str
    diff_type: str  # '目标端缺失', '源端缺失', '主键列变更'
    test_value: str
    prod_value: str


@dataclasses.dataclass
class SequenceDiff:
    """序列差异"""
    schema: str
    sequence: str
    diff_type: str  # '目标端缺失', '源端缺失', '属性变更'
    test_value: str
    prod_value: str
    test_sequence: Optional[SequenceInfo] = None
    prod_sequence: Optional[SequenceInfo] = None


@dataclasses.dataclass
class ForeignKeyDiff:
    """外键差异"""
    schema: str
    table: str
    fk_name: str
    diff_type: str  # '目标端缺失', '源端缺失', '定义变更'
    test_value: str
    prod_value: str
    test_fk: Optional[ForeignKeyInfo] = None
    prod_fk: Optional[ForeignKeyInfo] = None


@dataclasses.dataclass
class ViewDiff:
    """视图差异"""
    schema: str
    view: str
    diff_type: str  # '目标端缺失', '源端缺失', '定义变更'
    test_value: str
    prod_value: str


@dataclasses.dataclass
class ProcedureDiff:
    """存储过程/函数/触发器差异"""
    schema: str
    name: str
    object_type: str
    diff_type: str  # '目标端缺失', '源端缺失', '定义变更'
    test_value: str
    prod_value: str


@dataclasses.dataclass
class ComparisonResult:
    """对比结果"""
    # 连接信息
    test_host: str = ''
    test_port: int = 0
    prod_hosts: List[str] = dataclasses.field(default_factory=list)
    source_hosts: List[str] = dataclasses.field(default_factory=list)
    target_hosts: List[str] = dataclasses.field(default_factory=list)
    source_label: str = '源端'
    target_label: str = '目标端'

    # 差异列表
    table_diffs: List[TableDiff] = dataclasses.field(default_factory=list)
    column_diffs: List[ColumnDiff] = dataclasses.field(default_factory=list)
    index_diffs: List[IndexDiff] = dataclasses.field(default_factory=list)
    primary_key_diffs: List[PrimaryKeyDiff] = dataclasses.field(default_factory=list)
    sequence_diffs: List[SequenceDiff] = dataclasses.field(default_factory=list)
    foreign_key_diffs: List[ForeignKeyDiff] = dataclasses.field(default_factory=list)
    view_diffs: List[ViewDiff] = dataclasses.field(default_factory=list)
    procedure_diffs: List[ProcedureDiff] = dataclasses.field(default_factory=list)

    # 模式映射信息
    schema_mappings: Dict[str, str] = dataclasses.field(default_factory=dict)

    # 统计信息
    test_schema_count: int = 0
    prod_schema_count: int = 0
    common_schema_count: int = 0
    test_table_count: int = 0
    prod_table_count: int = 0
    test_unique_table_count: int = 0  # 源端独有模式中的表数量
    prod_unique_table_count: int = 0  # 目标端独有模式中的表数量
    common_schema_table_count: int = 0  # 共同模式中的总表数
    common_table_count: int = 0
    test_index_count: int = 0
    prod_index_count: int = 0
    common_index_count: int = 0
    test_sequence_count: int = 0
    prod_sequence_count: int = 0
    common_sequence_count: int = 0
    test_fk_count: int = 0
    prod_fk_count: int = 0
    test_view_count: int = 0
    prod_view_count: int = 0
    test_procedure_count: int = 0
    prod_procedure_count: int = 0


@dataclasses.dataclass
class CompareRunResult:
    """一次比对运行结果"""
    result: ComparisonResult
    report_path: str
    output_dir: str


class SchemaComparator:
    """模式对比器"""

    def __init__(self, test_metadata: Dict[str, SchemaInfo], prod_metadata: Dict[str, SchemaInfo]):
        self.test_metadata = test_metadata
        self.prod_metadata = prod_metadata

    def compare_full(self, test_host: str = '', test_port: int = 0,
                     prod_hosts: List[str] = None,
                     source_hosts: List[str] = None,
                     target_hosts: List[str] = None,
                     source_label: str = '源端',
                     target_label: str = '目标端') -> ComparisonResult:
        """完整对比"""
        result = ComparisonResult(
            test_host=test_host,
            test_port=test_port,
            prod_hosts=prod_hosts or [],
            source_hosts=source_hosts or [],
            target_hosts=target_hosts or [],
            source_label=source_label,
            target_label=target_label,
        )

        logger.info("  对比表差异...")
        result.table_diffs = self._compare_tables()

        logger.info("  对比字段差异...")
        result.column_diffs = self._compare_columns()

        logger.info("  对比索引差异...")
        result.index_diffs = self._compare_indexes()

        logger.info("  对比主键差异...")
        result.primary_key_diffs = self._compare_primary_keys()

        logger.info("  对比序列差异...")
        result.sequence_diffs = self._compare_sequences()

        logger.info("  对比外键差异...")
        result.foreign_key_diffs = self._compare_foreign_keys()

        logger.info("  对比视图差异...")
        result.view_diffs = self._compare_views()

        logger.info("  对比存储过程/函数/触发器差异...")
        result.procedure_diffs = self._compare_procedures()

        # 详细的表统计
        result.test_table_count = sum(len(s.tables) for s in self.test_metadata.values())
        result.prod_table_count = sum(len(s.tables) for s in self.prod_metadata.values())
        result.test_index_count = self._count_indexes(self.test_metadata)
        result.prod_index_count = self._count_indexes(self.prod_metadata)
        result.test_sequence_count = sum(len(s.sequences) for s in self.test_metadata.values())
        result.prod_sequence_count = sum(len(s.sequences) for s in self.prod_metadata.values())
        result.common_sequence_count = self._get_common_sequence_count()
        result.test_fk_count = sum(
            sum(len(t.foreign_keys) for t in s.tables.values())
            for s in self.test_metadata.values()
        )
        result.prod_fk_count = sum(
            sum(len(t.foreign_keys) for t in s.tables.values())
            for s in self.prod_metadata.values()
        )
        result.test_view_count = sum(len(s.views) for s in self.test_metadata.values())
        result.prod_view_count = sum(len(s.views) for s in self.prod_metadata.values())
        result.test_procedure_count = sum(len(s.procedures) for s in self.test_metadata.values())
        result.prod_procedure_count = sum(len(s.procedures) for s in self.prod_metadata.values())

        # 计算独有模式中的表数量
        test_unique_schemas = set(self.test_metadata.keys()) - set(self.prod_metadata.keys())
        prod_unique_schemas = set(self.prod_metadata.keys()) - set(self.test_metadata.keys())
        result.test_unique_table_count = sum(len(self.test_metadata[s].tables) for s in test_unique_schemas)
        result.prod_unique_table_count = sum(len(self.prod_metadata[s].tables) for s in prod_unique_schemas)

        # 共同模式中的表数量
        common_schemas = set(self.test_metadata.keys()) & set(self.prod_metadata.keys())
        result.common_schema_table_count = sum(len(self.test_metadata[s].tables) for s in common_schemas)

        result.common_table_count = self._get_common_table_count()
        result.common_index_count = self._get_common_index_count()

        return result

    @staticmethod
    def _count_indexes(metadata: Dict[str, SchemaInfo]) -> int:
        """统计普通索引数量，排除主键/唯一约束背后的索引。"""
        return sum(
            sum(1 for index in table.indexes.values() if not index.is_constraint_backed())
            for schema_info in metadata.values()
            for table in schema_info.tables.values()
        )

    @staticmethod
    def _normalize_name(value: str) -> str:
        """规范化对象名，用于大小写无关匹配。"""
        return (value or '').strip().upper()

    @classmethod
    def _build_name_map(cls, objects: Dict[str, Any]) -> Dict[str, str]:
        """构建大小写无关名称映射，值保留原始对象名。"""
        name_map = {}
        for name in objects.keys():
            normalized = cls._normalize_name(name)
            if normalized not in name_map:
                name_map[normalized] = name
        return name_map

    def _compare_tables(self) -> List[TableDiff]:
        """对比表差异"""
        diffs = []
        common_schemas = set(self.test_metadata.keys()) & set(self.prod_metadata.keys())

        for schema in sorted(common_schemas):
            test_tables = self.test_metadata[schema].tables
            prod_tables = self.prod_metadata[schema].tables

            test_table_map = self._build_name_map(test_tables)
            prod_table_map = self._build_name_map(prod_tables)
            test_table_names = set(test_table_map.keys())
            prod_table_names = set(prod_table_map.keys())

            # 仅在源端的表（目标端缺失）
            for normalized_table_name in test_table_names - prod_table_names:
                table_name = test_table_map[normalized_table_name]
                table = test_tables[table_name]
                diffs.append(TableDiff(
                    schema=schema,
                    table=table_name,
                    diff_type='目标端缺失',
                    comment=table.comment,
                    test_table=table
                ))

            # 仅在目标端的表（源端缺失）
            for normalized_table_name in prod_table_names - test_table_names:
                table_name = prod_table_map[normalized_table_name]
                table = prod_tables[table_name]
                diffs.append(TableDiff(
                    schema=schema,
                    table=table_name,
                    diff_type='源端缺失',
                    comment=table.comment,
                    prod_table=table
                ))

        return diffs

    def _compare_columns(self) -> List[ColumnDiff]:
        """对比字段差异"""
        diffs = []
        common_schemas = set(self.test_metadata.keys()) & set(self.prod_metadata.keys())

        for schema in sorted(common_schemas):
            test_tables = self.test_metadata[schema].tables
            prod_tables = self.prod_metadata[schema].tables

            test_table_map = self._build_name_map(test_tables)
            prod_table_map = self._build_name_map(prod_tables)
            common_table_names = set(test_table_map.keys()) & set(prod_table_map.keys())

            for normalized_table_name in sorted(common_table_names):
                test_table_name = test_table_map[normalized_table_name]
                prod_table_name = prod_table_map[normalized_table_name]
                test_columns = test_tables[test_table_name].columns
                prod_columns = prod_tables[prod_table_name].columns

                test_col_map = self._build_name_map(test_columns)
                prod_col_map = self._build_name_map(prod_columns)
                test_col_names = set(test_col_map.keys())
                prod_col_names = set(prod_col_map.keys())

                # 仅在源端的字段（目标端缺失）
                for normalized_col_name in test_col_names - prod_col_names:
                    col_name = test_col_map[normalized_col_name]
                    col = test_columns[col_name]
                    diffs.append(ColumnDiff(
                        schema=schema,
                        table=prod_table_name,
                        column=col_name,
                        diff_type='目标端缺失',
                        test_value=col.get_type_string(),
                        prod_value='-',
                        test_column=col
                    ))

                # 仅在目标端的字段（源端缺失）
                for normalized_col_name in prod_col_names - test_col_names:
                    col_name = prod_col_map[normalized_col_name]
                    col = prod_columns[col_name]
                    diffs.append(ColumnDiff(
                        schema=schema,
                        table=prod_table_name,
                        column=col_name,
                        diff_type='源端缺失',
                        test_value='-',
                        prod_value=col.get_type_string(),
                        prod_column=col
                    ))

                # 共有字段对比
                for normalized_col_name in sorted(test_col_names & prod_col_names):
                    test_col = test_columns[test_col_map[normalized_col_name]]
                    prod_col = prod_columns[prod_col_map[normalized_col_name]]
                    self._compare_column_detail(schema, prod_table_name, prod_col.name, test_col, prod_col, diffs)

        return diffs

    def _compare_column_detail(self, schema: str, table: str, col_name: str,
                               test_col: ColumnInfo, prod_col: ColumnInfo, diffs: List[ColumnDiff]):
        """详细对比字段"""
        # 类型对比
        if _normalize_dm_data_type(test_col.data_type) != _normalize_dm_data_type(prod_col.data_type):
            diffs.append(ColumnDiff(
                schema=schema,
                table=table,
                column=col_name,
                diff_type='类型变更',
                test_value=test_col.get_type_string(),
                prod_value=prod_col.get_type_string(),
                test_column=test_col,
                prod_column=prod_col
            ))

        # 长度对比
        if (_normalize_dm_data_type(test_col.data_type) in LENGTH_DATA_TYPES
                and test_col.length != prod_col.length
                and test_col.length and prod_col.length):
            diffs.append(ColumnDiff(
                schema=schema,
                table=table,
                column=col_name,
                diff_type='长度变更',
                test_value=str(test_col.length),
                prod_value=str(prod_col.length),
                test_column=test_col,
                prod_column=prod_col
            ))

        # 精度对比
        if (_normalize_dm_data_type(test_col.data_type) in PRECISION_DATA_TYPES
                and (test_col.precision != prod_col.precision or test_col.scale != prod_col.scale)):
            test_prec = f"{test_col.precision},{test_col.scale}" if test_col.precision else "-"
            prod_prec = f"{prod_col.precision},{prod_col.scale}" if prod_col.precision else "-"
            diffs.append(ColumnDiff(
                schema=schema,
                table=table,
                column=col_name,
                diff_type='精度变更',
                test_value=test_prec,
                prod_value=prod_prec,
                test_column=test_col,
                prod_column=prod_col
            ))

        # 可空性对比
        if test_col.nullable != prod_col.nullable:
            if not (test_col.is_primary_key or prod_col.is_primary_key):
                diffs.append(ColumnDiff(
                    schema=schema,
                    table=table,
                    column=col_name,
                    diff_type='可空性变更',
                    test_value='Y' if test_col.nullable == 'Y' else 'N',
                    prod_value='Y' if prod_col.nullable == 'Y' else 'N',
                    test_column=test_col,
                    prod_column=prod_col
                ))

        # 默认值对比
        test_default = self._normalize_default_for_compare(test_col.default)
        prod_default = self._normalize_default_for_compare(prod_col.default)
        if test_default != prod_default:
            diffs.append(ColumnDiff(
                schema=schema,
                table=table,
                column=col_name,
                diff_type='默认值变更',
                test_value=self._format_default_display(test_col.default),
                prod_value=self._format_default_display(prod_col.default),
                test_column=test_col,
                prod_column=prod_col
            ))

        # 注释对比（双向：源端无注释但目标端有注释也报差异）
        test_comment = (test_col.comment or '').strip()
        prod_comment = (prod_col.comment or '').strip()
        if test_comment != prod_comment:
            diffs.append(ColumnDiff(
                schema=schema,
                table=table,
                column=col_name,
                diff_type='注释变更',
                test_value=test_comment or '-',
                prod_value=prod_comment or '-',
                test_column=test_col,
                prod_column=prod_col
            ))

    def _compare_indexes(self) -> List[IndexDiff]:
        """对比索引差异"""
        diffs = []
        common_schemas = set(self.test_metadata.keys()) & set(self.prod_metadata.keys())

        for schema in sorted(common_schemas):
            test_tables = self.test_metadata[schema].tables
            prod_tables = self.prod_metadata[schema].tables

            test_table_map = self._build_name_map(test_tables)
            prod_table_map = self._build_name_map(prod_tables)
            common_table_names = set(test_table_map.keys()) & set(prod_table_map.keys())

            for normalized_table_name in sorted(common_table_names):
                test_table_name = test_table_map[normalized_table_name]
                prod_table_name = prod_table_map[normalized_table_name]
                test_indexes = self._filter_comparable_indexes(test_tables[test_table_name].indexes)
                prod_indexes = self._filter_comparable_indexes(prod_tables[prod_table_name].indexes)

                test_index_map = self._build_name_map(test_indexes)
                prod_index_map = self._build_name_map(prod_indexes)
                test_index_names = set(test_index_map.keys())
                prod_index_names = set(prod_index_map.keys())

                test_only_index_names = test_index_names - prod_index_names
                prod_only_index_names = prod_index_names - test_index_names
                test_only_actual_names = {test_index_map[name] for name in test_only_index_names}
                prod_only_actual_names = {prod_index_map[name] for name in prod_only_index_names}
                matched_equivalent_test_names, matched_equivalent_prod_names = self._match_equivalent_indexes(
                    test_indexes,
                    prod_indexes,
                    test_only_actual_names,
                    prod_only_actual_names
                )

                # 仅在源端的索引（目标端缺失）
                for normalized_index_name in sorted(test_only_index_names):
                    index_name = test_index_map[normalized_index_name]
                    if index_name in matched_equivalent_test_names:
                        continue
                    index = test_indexes[index_name]
                    diffs.append(IndexDiff(
                        schema=schema,
                        table=prod_table_name,
                        index=index_name,
                        diff_type='目标端缺失',
                        test_value=index.get_definition_string(),
                        prod_value='-',
                        test_index=index
                    ))

                # 仅在目标端的索引（源端缺失）
                for normalized_index_name in sorted(prod_only_index_names):
                    index_name = prod_index_map[normalized_index_name]
                    if index_name in matched_equivalent_prod_names:
                        continue
                    index = prod_indexes[index_name]
                    diffs.append(IndexDiff(
                        schema=schema,
                        table=prod_table_name,
                        index=index_name,
                        diff_type='源端缺失',
                        test_value='-',
                        prod_value=index.get_definition_string(),
                        prod_index=index
                    ))

                # 共有索引详细对比
                for normalized_index_name in sorted(test_index_names & prod_index_names):
                    test_index = test_indexes[test_index_map[normalized_index_name]]
                    prod_index = prod_indexes[prod_index_map[normalized_index_name]]
                    self._compare_index_detail(schema, prod_table_name, prod_index.name, test_index, prod_index, diffs)

        return diffs

    def _compare_primary_keys(self) -> List[PrimaryKeyDiff]:
        """对比主键差异"""
        diffs = []
        common_schemas = set(self.test_metadata.keys()) & set(self.prod_metadata.keys())

        for schema in sorted(common_schemas):
            test_tables = self.test_metadata[schema].tables
            prod_tables = self.prod_metadata[schema].tables

            test_table_map = self._build_name_map(test_tables)
            prod_table_map = self._build_name_map(prod_tables)
            common_table_names = set(test_table_map.keys()) & set(prod_table_map.keys())

            for normalized_table_name in sorted(common_table_names):
                test_table_name = test_table_map[normalized_table_name]
                prod_table_name = prod_table_map[normalized_table_name]
                test_pk = test_tables[test_table_name].primary_key_columns
                prod_pk = prod_tables[prod_table_name].primary_key_columns

                # 主键列名按大写归一化对比，但保留源端 POSITION 顺序用于 SQL 生成，
                # 避免 id 与 ID 被误判为差异，同时保持主键索引列顺序与源端一致
                test_pk_norm = [self._normalize_name(c) for c in test_pk]
                prod_pk_norm = [self._normalize_name(c) for c in prod_pk]
                test_pk_str = ', '.join(test_pk_norm)
                prod_pk_str = ', '.join(prod_pk_norm)

                if not test_pk and not prod_pk:
                    continue

                # 集合相同即无差异（顺序无关）；集合不同时再细化判断
                test_pk_set = set(test_pk_norm)
                prod_pk_set = set(prod_pk_norm)

                if test_pk and not prod_pk:
                    diffs.append(PrimaryKeyDiff(
                        schema=schema,
                        table=prod_table_name,
                        diff_type='目标端缺失',
                        test_value=test_pk_str,
                        prod_value='-',
                    ))
                elif not test_pk and prod_pk:
                    diffs.append(PrimaryKeyDiff(
                        schema=schema,
                        table=prod_table_name,
                        diff_type='源端缺失',
                        test_value='-',
                        prod_value=prod_pk_str,
                    ))
                elif test_pk_set != prod_pk_set:
                    diffs.append(PrimaryKeyDiff(
                        schema=schema,
                        table=prod_table_name,
                        diff_type='主键列变更',
                        test_value=test_pk_str,
                        prod_value=prod_pk_str,
                    ))

        return diffs

    def _compare_sequences(self) -> List[SequenceDiff]:
        """对比序列差异。序列是 schema 级对象，直接按 schema 名比对序列集合。"""
        diffs = []
        common_schemas = set(self.test_metadata.keys()) & set(self.prod_metadata.keys())

        for schema in sorted(common_schemas):
            test_seqs = self.test_metadata[schema].sequences
            prod_seqs = self.prod_metadata[schema].sequences

            test_map = self._build_name_map(test_seqs)
            prod_map = self._build_name_map(prod_seqs)
            test_names = set(test_map.keys())
            prod_names = set(prod_map.keys())

            # 仅在源端（目标端缺失）→ CREATE
            for normalized_name in sorted(test_names - prod_names):
                name = test_map[normalized_name]
                seq = test_seqs[name]
                diffs.append(SequenceDiff(
                    schema=schema,
                    sequence=name,
                    diff_type='目标端缺失',
                    test_value=seq.get_attribute_string(),
                    prod_value='-',
                    test_sequence=seq,
                ))

            # 仅在目标端（源端缺失）→ 默认不生成 DROP，需人工确认
            for normalized_name in sorted(prod_names - test_names):
                name = prod_map[normalized_name]
                seq = prod_seqs[name]
                diffs.append(SequenceDiff(
                    schema=schema,
                    sequence=name,
                    diff_type='源端缺失',
                    test_value='-',
                    prod_value=seq.get_attribute_string(),
                    prod_sequence=seq,
                ))

            # 共有序列属性对比
            for normalized_name in sorted(test_names & prod_names):
                test_seq = test_seqs[test_map[normalized_name]]
                prod_seq = prod_seqs[prod_map[normalized_name]]
                if self._sequence_attributes_equal(test_seq, prod_seq):
                    continue
                diffs.append(SequenceDiff(
                    schema=schema,
                    sequence=prod_seq.name,
                    diff_type='属性变更',
                    test_value=test_seq.get_attribute_string(),
                    prod_value=prod_seq.get_attribute_string(),
                    test_sequence=test_seq,
                    prod_sequence=prod_seq,
                ))

        return diffs

    @staticmethod
    def _sequence_attributes_equal(a: SequenceInfo, b: SequenceInfo) -> bool:
        """序列属性是否一致。LAST_NUMBER 是运行时状态，不参与结构对比。"""
        return (
            (a.start_with or 0) == (b.start_with or 0)
            and (a.increment_by or 0) == (b.increment_by or 0)
            and (a.min_value or 0) == (b.min_value or 0)
            and (a.max_value or 0) == (b.max_value or 0)
            and (a.cycle_flag or '').upper() == (b.cycle_flag or '').upper()
            and (a.order_flag or '').upper() == (b.order_flag or '').upper()
            and (a.cache_size or 0) == (b.cache_size or 0)
        )

    def _get_common_sequence_count(self) -> int:
        """共同序列数量（同名即算）。"""
        count = 0
        common_schemas = set(self.test_metadata.keys()) & set(self.prod_metadata.keys())
        for schema in common_schemas:
            test_names = set(self._build_name_map(self.test_metadata[schema].sequences).keys())
            prod_names = set(self._build_name_map(self.prod_metadata[schema].sequences).keys())
            count += len(test_names & prod_names)
        return count

    # ---------- 外键对比 ----------

    def _compare_foreign_keys(self) -> List[ForeignKeyDiff]:
        """对比外键差异。"""
        diffs = []
        common_schemas = set(self.test_metadata.keys()) & set(self.prod_metadata.keys())

        for schema in sorted(common_schemas):
            test_tables = self.test_metadata[schema].tables
            prod_tables = self.prod_metadata[schema].tables

            test_table_map = self._build_name_map(test_tables)
            prod_table_map = self._build_name_map(prod_tables)
            common_table_names = set(test_table_map.keys()) & set(prod_table_map.keys())

            for normalized_table_name in sorted(common_table_names):
                test_table_name = test_table_map[normalized_table_name]
                prod_table_name = prod_table_map[normalized_table_name]
                test_fks = test_tables[test_table_name].foreign_keys
                prod_fks = prod_tables[prod_table_name].foreign_keys

                test_map = self._build_name_map(test_fks)
                prod_map = self._build_name_map(prod_fks)
                test_names = set(test_map.keys())
                prod_names = set(prod_map.keys())

                for normalized_name in sorted(test_names - prod_names):
                    name = test_map[normalized_name]
                    fk = test_fks[name]
                    diffs.append(ForeignKeyDiff(
                        schema=schema, table=prod_table_name, fk_name=name,
                        diff_type='目标端缺失',
                        test_value=fk.get_definition_string(), prod_value='-',
                        test_fk=fk,
                    ))

                for normalized_name in sorted(prod_names - test_names):
                    name = prod_map[normalized_name]
                    fk = prod_fks[name]
                    diffs.append(ForeignKeyDiff(
                        schema=schema, table=prod_table_name, fk_name=name,
                        diff_type='源端缺失',
                        test_value='-', prod_value=fk.get_definition_string(),
                        prod_fk=fk,
                    ))

                for normalized_name in sorted(test_names & prod_names):
                    test_fk = test_fks[test_map[normalized_name]]
                    prod_fk = prod_fks[prod_map[normalized_name]]
                    if self._fk_definition_equal(test_fk, prod_fk):
                        continue
                    diffs.append(ForeignKeyDiff(
                        schema=schema, table=prod_table_name, fk_name=prod_fk.name,
                        diff_type='定义变更',
                        test_value=test_fk.get_definition_string(),
                        prod_value=prod_fk.get_definition_string(),
                        test_fk=test_fk, prod_fk=prod_fk,
                    ))

        return diffs

    @staticmethod
    def _fk_definition_equal(a: ForeignKeyInfo, b: ForeignKeyInfo) -> bool:
        """外键定义是否一致（列、引用表、删除规则）。"""
        a_cols = sorted((c.column_name.upper(), c.ref_column_name.upper()) for c in a.columns)
        b_cols = sorted((c.column_name.upper(), c.ref_column_name.upper()) for c in b.columns)
        return (
            a_cols == b_cols
            and a.ref_schema.upper() == b.ref_schema.upper()
            and a.ref_table_name.upper() == b.ref_table_name.upper()
            and (a.delete_rule or 'NO ACTION') == (b.delete_rule or 'NO ACTION')
        )

    # ---------- 视图对比 ----------

    def _compare_views(self) -> List[ViewDiff]:
        """对比视图差异。"""
        diffs = []
        common_schemas = set(self.test_metadata.keys()) & set(self.prod_metadata.keys())

        for schema in sorted(common_schemas):
            test_views = self.test_metadata[schema].views
            prod_views = self.prod_metadata[schema].views

            test_map = self._build_name_map(test_views)
            prod_map = self._build_name_map(prod_views)
            test_names = set(test_map.keys())
            prod_names = set(prod_map.keys())

            for normalized_name in sorted(test_names - prod_names):
                name = test_map[normalized_name]
                diffs.append(ViewDiff(
                    schema=schema, view=name, diff_type='目标端缺失',
                    test_value=test_views[name].text or '-', prod_value='-',
                ))

            for normalized_name in sorted(prod_names - test_names):
                name = prod_map[normalized_name]
                diffs.append(ViewDiff(
                    schema=schema, view=name, diff_type='源端缺失',
                    test_value='-', prod_value=prod_views[name].text or '-',
                ))

            for normalized_name in sorted(test_names & prod_names):
                test_view = test_views[test_map[normalized_name]]
                prod_view = prod_views[prod_map[normalized_name]]
                if self._normalize_view_text(test_view.text) != self._normalize_view_text(prod_view.text):
                    diffs.append(ViewDiff(
                        schema=schema, view=prod_view.name, diff_type='定义变更',
                        test_value=test_view.text or '-', prod_value=prod_view.text or '-',
                    ))

        return diffs

    @staticmethod
    def _normalize_view_text(text: str) -> str:
        """规范化视图定义文本用于比较（压缩空白、统一大小写）。"""
        return ' '.join((text or '').split()).upper()

    # ---------- 存储过程/函数/触发器对比 ----------

    def _compare_procedures(self) -> List[ProcedureDiff]:
        """对比存储过程/函数/触发器差异。"""
        diffs = []
        common_schemas = set(self.test_metadata.keys()) & set(self.prod_metadata.keys())

        for schema in sorted(common_schemas):
            test_procs = self.test_metadata[schema].procedures
            prod_procs = self.prod_metadata[schema].procedures

            test_map = self._build_name_map(test_procs)
            prod_map = self._build_name_map(prod_procs)
            test_names = set(test_map.keys())
            prod_names = set(prod_map.keys())

            for normalized_name in sorted(test_names - prod_names):
                key = test_map[normalized_name]
                proc = test_procs[key]
                diffs.append(ProcedureDiff(
                    schema=schema, name=proc.name, object_type=proc.object_type,
                    diff_type='目标端缺失',
                    test_value=proc.text or '-', prod_value='-',
                ))

            for normalized_name in sorted(prod_names - test_names):
                key = prod_map[normalized_name]
                proc = prod_procs[key]
                diffs.append(ProcedureDiff(
                    schema=schema, name=proc.name, object_type=proc.object_type,
                    diff_type='源端缺失',
                    test_value='-', prod_value=proc.text or '-',
                ))

            for normalized_name in sorted(test_names & prod_names):
                test_key = test_map[normalized_name]
                prod_key = prod_map[normalized_name]
                test_proc = test_procs[test_key]
                prod_proc = prod_procs[prod_key]
                if self._normalize_view_text(test_proc.text) != self._normalize_view_text(prod_proc.text):
                    diffs.append(ProcedureDiff(
                        schema=schema, name=prod_proc.name, object_type=prod_proc.object_type,
                        diff_type='定义变更',
                        test_value=test_proc.text or '-', prod_value=prod_proc.text or '-',
                    ))

        return diffs

    @staticmethod
    def _filter_comparable_indexes(indexes: Dict[str, IndexInfo]) -> Dict[str, IndexInfo]:
        """过滤掉主键/唯一约束背后的索引，这类对象应按约束处理。"""
        return {
            index_name: index
            for index_name, index in indexes.items()
            if not index.is_constraint_backed()
        }

    def _match_equivalent_indexes(self, test_indexes: Dict[str, IndexInfo], prod_indexes: Dict[str, IndexInfo],
                                  test_only_names: Set[str], prod_only_names: Set[str]) -> Tuple[Set[str], Set[str]]:
        """匹配两边定义一致但名称不同的普通索引。"""
        matched_test_names: Set[str] = set()
        matched_prod_names: Set[str] = set()

        prod_signature_map: Dict[Tuple[str, str, Tuple[Tuple[str, str], ...]], List[str]] = {}
        for prod_name in prod_only_names:
            prod_signature_map.setdefault(self._get_index_signature(prod_indexes[prod_name]), []).append(prod_name)

        for test_name in sorted(test_only_names):
            signature = self._get_index_signature(test_indexes[test_name])
            candidate_names = prod_signature_map.get(signature)
            if not candidate_names:
                continue

            prod_name = candidate_names.pop(0)
            matched_test_names.add(test_name)
            matched_prod_names.add(prod_name)

        return matched_test_names, matched_prod_names

    @staticmethod
    def _is_auto_index_name(index_name: str) -> bool:
        """识别达梦常见自动索引名，如 INDEX33560575。"""
        return re.fullmatch(r'INDEX\d+', index_name or '', flags=re.IGNORECASE) is not None

    def _get_index_signature(self, index: IndexInfo) -> Tuple[str, str, Tuple[Tuple[str, str], ...]]:
        """获取不包含索引名的索引定义签名。"""
        return (
            self._normalize_text(index.uniqueness),
            self._normalize_text(index.index_type),
            self._get_index_columns_signature(index)
        )

    def _compare_index_detail(self, schema: str, table: str, index_name: str,
                              test_index: IndexInfo, prod_index: IndexInfo, diffs: List[IndexDiff]):
        """详细对比索引定义"""
        if self._normalize_text(test_index.uniqueness) != self._normalize_text(prod_index.uniqueness):
            diffs.append(IndexDiff(
                schema=schema,
                table=table,
                index=index_name,
                diff_type='唯一性变更',
                test_value=test_index.uniqueness or '-',
                prod_value=prod_index.uniqueness or '-',
                test_index=test_index,
                prod_index=prod_index
            ))

        if self._normalize_text(test_index.index_type) != self._normalize_text(prod_index.index_type):
            diffs.append(IndexDiff(
                schema=schema,
                table=table,
                index=index_name,
                diff_type='索引类型变更',
                test_value=test_index.index_type or '-',
                prod_value=prod_index.index_type or '-',
                test_index=test_index,
                prod_index=prod_index
            ))

        test_columns = test_index.get_columns_string() or '-'
        prod_columns = prod_index.get_columns_string() or '-'
        if self._get_index_columns_signature(test_index) != self._get_index_columns_signature(prod_index):
            diffs.append(IndexDiff(
                schema=schema,
                table=table,
                index=index_name,
                diff_type='索引列变更',
                test_value=test_columns,
                prod_value=prod_columns,
                test_index=test_index,
                prod_index=prod_index
            ))

    def _get_index_columns_signature(self, index: IndexInfo) -> Tuple[Tuple[str, str], ...]:
        """获取不区分大小写的索引列签名。"""
        return tuple(
            (self._normalize_text(column.name), self._normalize_text(column.sort_order or 'ASC'))
            for column in sorted(index.columns, key=lambda col: (col.position, col.name))
        )

    @staticmethod
    def _normalize_text(value: Optional[str]) -> str:
        """规范化对比文本"""
        return (value or '').strip().upper()

    @staticmethod
    def _normalize_default_for_compare(value: Optional[str]) -> Optional[str]:
        """规范化默认值用于比较，DEFAULT NULL 与无默认值视为一致。"""
        if value is None:
            return None

        default = str(value).strip()
        if not default or default.upper() == 'NULL':
            return None
        return default

    @staticmethod
    def _format_default_display(value: Optional[str]) -> str:
        """格式化默认值展示文本。"""
        if value is None:
            return '-'

        default = str(value).strip()
        return default if default else '-'

    def _get_common_table_count(self) -> int:
        """获取共同表数量"""
        count = 0
        common_schemas = set(self.test_metadata.keys()) & set(self.prod_metadata.keys())
        for schema in common_schemas:
            test_tables = set(self._build_name_map(self.test_metadata[schema].tables).keys())
            prod_tables = set(self._build_name_map(self.prod_metadata[schema].tables).keys())
            count += len(test_tables & prod_tables)
        return count

    def _get_common_index_count(self) -> int:
        """获取共同索引数量"""
        count = 0
        common_schemas = set(self.test_metadata.keys()) & set(self.prod_metadata.keys())

        for schema in common_schemas:
            test_tables = self.test_metadata[schema].tables
            prod_tables = self.prod_metadata[schema].tables
            test_table_map = self._build_name_map(test_tables)
            prod_table_map = self._build_name_map(prod_tables)
            common_tables = set(test_table_map.keys()) & set(prod_table_map.keys())

            for normalized_table_name in common_tables:
                test_table_name = test_table_map[normalized_table_name]
                prod_table_name = prod_table_map[normalized_table_name]
                test_signatures = [
                    self._get_index_signature(index)
                    for index in self._filter_comparable_indexes(test_tables[test_table_name].indexes).values()
                ]
                prod_signatures = [
                    self._get_index_signature(index)
                    for index in self._filter_comparable_indexes(prod_tables[prod_table_name].indexes).values()
                ]

                for signature in test_signatures:
                    if signature in prod_signatures:
                        count += 1
                        prod_signatures.remove(signature)

        return count


class ReportGenerator:
    """报告生成器"""

    def __init__(self, result: ComparisonResult):
        self.result = result

    def _create_style(self):
        """创建样式"""
        # 标题样式
        title_font = Font(name='微软雅黑', size=12, bold=True, color='FFFFFF')
        title_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        title_alignment = Alignment(horizontal='center', vertical='center')

        # 边框
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        return {
            'title_font': title_font,
            'title_fill': title_fill,
            'title_alignment': title_alignment,
            'thin_border': thin_border
        }

    @staticmethod
    def _apply_table_features(ws, max_col: int = None):
        """给差异 sheet 加冻结首行和自动筛选。"""
        ws.freeze_panes = 'A2'
        if ws.max_row >= 2:
            last_col = max_col or ws.max_column
            col_letter = _openpyxl_get_column_letter(last_col)
            ws.auto_filter.ref = f"A1:{col_letter}{ws.max_row}"

    def generate_excel(self, output_path: str):
        """生成Excel报告"""
        wb = Workbook()
        style = self._create_style()

        # 删除默认sheet
        if 'Sheet' in wb.sheetnames:
            del wb['Sheet']

        # Sheet1: 汇总概览
        self._create_summary_sheet(wb, style)

        # Sheet2: 表差异
        self._create_table_diff_sheet(wb, style)

        # Sheet4: 字段差异
        self._create_column_diff_sheet(wb, style)

        # Sheet5: 索引差异
        self._create_index_diff_sheet(wb, style)

        # Sheet6: 主键差异
        self._create_primary_key_diff_sheet(wb, style)

        # Sheet7: 序列差异
        self._create_sequence_diff_sheet(wb, style)

        # Sheet8: 外键差异
        self._create_foreign_key_diff_sheet(wb, style)

        # Sheet9: 视图差异
        self._create_view_diff_sheet(wb, style)

        # Sheet10: 存储过程/函数/触发器差异
        self._create_procedure_diff_sheet(wb, style)

        wb.save(output_path)

    def _create_summary_sheet(self, wb: Workbook, style: dict):
        """创建汇总sheet"""
        ws = wb.create_sheet('汇总概览')

        # 标题
        ws['A1'] = '达梦数据库结构差异对比报告'
        ws['A1'].font = Font(name='微软雅黑', size=14, bold=True)
        ws.merge_cells('A1:B1')

        # 连接信息
        source_label = self.result.source_label
        target_label = self.result.target_label
        row = 3
        ws[f'A{row}'] = source_label
        ws[f'A{row}'].font = Font(bold=True)
        source_display = self.result.test_host and f"{self.result.test_host}:{self.result.test_port}"
        if not source_display and self.result.source_hosts:
            source_display = ', '.join(self.result.source_hosts)
        ws[f'B{row}'] = source_display or '-'

        row += 1
        ws[f'A{row}'] = target_label
        ws[f'A{row}'].font = Font(bold=True)
        ws[f'B{row}'] = ', '.join(self.result.prod_hosts) if self.result.prod_hosts else '-'

        row += 1
        ws[f'A{row}'] = '对比时间'
        ws[f'A{row}'].font = Font(bold=True)
        ws[f'B{row}'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # 模式映射信息
        if self.result.schema_mappings:
            row += 2
            ws[f'A{row}'] = '模式映射'
            ws[f'A{row}'].font = Font(bold=True, size=12)

            row += 1
            ws.cell(row=row, column=1, value='源端模式').font = Font(bold=True)
            ws.cell(row=row, column=2, value='目标端模式').font = Font(bold=True)

            for test_schema, prod_schema in self.result.schema_mappings.items():
                row += 1
                ws.cell(row=row, column=1, value=test_schema)
                ws.cell(row=row, column=2, value=prod_schema)

        # 统计信息
        row += 2
        ws[f'A{row}'] = '差异统计'
        ws[f'A{row}'].font = Font(bold=True, size=12)

        row += 1
        headers = ['项目', source_label, target_label, '差异数量']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col, value=header)
            cell.font = style['title_font']
            cell.fill = style['title_fill']
            cell.alignment = style['title_alignment']
            cell.border = style['thin_border']

        # 数据行
        data = [
            ('模式数量', self.result.test_schema_count, self.result.prod_schema_count, '-'),
            ('  - 共同模式数', self.result.common_schema_count, self.result.common_schema_count, '-'),
            ('表数量', self.result.test_table_count, self.result.prod_table_count, '-'),
            ('  - 共同模式中的表', self.result.common_schema_table_count, self.result.common_schema_table_count, len(self.result.table_diffs)),
            ('共同表数量', self.result.common_table_count, self.result.common_table_count, '-'),
            ('字段差异', '-', '-', len(self.result.column_diffs)),
            ('索引数量', self.result.test_index_count, self.result.prod_index_count, '-'),
            ('共同索引数量', self.result.common_index_count, self.result.common_index_count, '-'),
            ('索引差异', '-', '-', len(self.result.index_diffs)),
            ('主键差异', '-', '-', len(self.result.primary_key_diffs)),
            ('序列数量', self.result.test_sequence_count, self.result.prod_sequence_count, '-'),
            ('共同序列数量', self.result.common_sequence_count, self.result.common_sequence_count, '-'),
            ('序列差异', '-', '-', len(self.result.sequence_diffs)),
            ('外键数量', self.result.test_fk_count, self.result.prod_fk_count, '-'),
            ('外键差异', '-', '-', len(self.result.foreign_key_diffs)),
            ('视图数量', self.result.test_view_count, self.result.prod_view_count, '-'),
            ('视图差异', '-', '-', len(self.result.view_diffs)),
            ('过程/函数/触发器数量', self.result.test_procedure_count, self.result.prod_procedure_count, '-'),
            ('过程/函数/触发器差异', '-', '-', len(self.result.procedure_diffs)),
        ]

        for r, (label, test_val, prod_val, diff_count) in enumerate(data, row + 1):
            ws.cell(row=r, column=1, value=label).border = style['thin_border']
            ws.cell(row=r, column=2, value=test_val).border = style['thin_border']
            ws.cell(row=r, column=3, value=prod_val).border = style['thin_border']
            ws.cell(row=r, column=4, value=diff_count).border = style['thin_border']

        # 调整列宽
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 15

    def _create_table_diff_sheet(self, wb: Workbook, style: dict):
        """创建表差异sheet"""
        ws = wb.create_sheet('表差异')

        # 标题
        headers = ['模式名', '表名', '所在环境', '表注释', '同步SQL']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = style['title_font']
            cell.fill = style['title_fill']
            cell.alignment = style['title_alignment']
            cell.border = style['thin_border']

        # 数据行
        for row, diff in enumerate(self.result.table_diffs, 2):
            ws.cell(row=row, column=1, value=diff.schema).border = style['thin_border']
            ws.cell(row=row, column=2, value=diff.table).border = style['thin_border']
            ws.cell(row=row, column=3, value=diff.diff_type).border = style['thin_border']
            ws.cell(row=row, column=4, value=diff.comment).border = style['thin_border']
            sql = self._generate_table_sync_sql(diff)
            sql_cell = ws.cell(row=row, column=5, value=sql or '')
            sql_cell.border = style['thin_border']
            sql_cell.alignment = Alignment(wrap_text=True, vertical='top')

        # 调整列宽
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 40
        ws.column_dimensions['E'].width = 120
        self._apply_table_features(ws)

    def _generate_table_sync_sql(self, diff: TableDiff) -> Optional[str]:
        """生成表级同步SQL。"""
        if diff.diff_type == '目标端缺失':
            if not diff.test_table:
                return None
            return self._generate_create_table_sql(diff.test_table, diff.schema)

        if diff.diff_type == '源端缺失':
            return f"DROP TABLE {self._qualified_name(diff.schema, diff.table)};"

        return None

    def _generate_create_table_sql(self, table: TableInfo, target_schema: str) -> Optional[str]:
        """根据表元数据生成达梦 CREATE TABLE 语句。"""
        if not table.columns:
            return None

        table_name = self._qualified_name(target_schema, table.name)
        column_lines = []
        for column in table.columns.values():
            column_sql = f"    {self._quote_identifier(column.name)} {self._get_column_type_sql(column)}"
            if column.is_identity:
                # IDENTITY 列不接受 DEFAULT，且隐含 NOT NULL
                column_sql += f" IDENTITY({self._identity_args(column)})"
            else:
                default_sql = self._format_default_sql(column.default)
                if default_sql:
                    column_sql += f" DEFAULT {default_sql}"
            if column.nullable == 'N' or column.is_identity:
                column_sql += " NOT NULL"
            column_lines.append(column_sql)

        # PRIMARY KEY 内联到 CREATE TABLE，避免单独 ALTER TABLE ADD PRIMARY KEY
        # 产生多余的自动命名索引
        if table.primary_key_columns:
            pk_cols = ', '.join(self._quote_identifier(c) for c in table.primary_key_columns)
            column_lines.append(f"    PRIMARY KEY ({pk_cols})")

        columns_sql = ",\n".join(column_lines)
        statements = [
            f"CREATE TABLE {table_name} (\n{columns_sql}\n);"
        ]

        if table.comment:
            statements.append(f"COMMENT ON TABLE {table_name} IS '{self._escape_sql_literal(table.comment)}';")

        for column in table.columns.values():
            if column.comment:
                column_name = f"{table_name}.{self._quote_identifier(column.name)}"
                statements.append(f"COMMENT ON COLUMN {column_name} IS '{self._escape_sql_literal(column.comment)}';")

        # 只为非约束索引生成 CREATE INDEX，约束索引由 PK/UK 约束自动创建
        for index in table.indexes.values():
            if index.is_constraint_backed():
                continue
            index_sql = self._generate_create_index_sql(index, target_schema=target_schema)
            if index_sql:
                statements.append(index_sql)

        return "\n".join(statements)

    def _get_column_type_sql(self, column: ColumnInfo) -> str:
        """生成达梦字段类型定义。"""
        return _format_dm_column_type(
            column.data_type,
            column.length,
            column.precision,
            column.scale,
            normalize_aliases=True
        )

    @staticmethod
    def _identity_args(column: ColumnInfo) -> str:
        """生成 IDENTITY(start, increment) 的参数串，缺省回落 (1, 1)。"""
        start = column.identity_start if column.identity_start is not None else 1
        incr = column.identity_increment if column.identity_increment is not None else 1
        return f"{start}, {incr}"

    def _normalize_type_string(self, type_string: str) -> str:
        """将展示用类型字符串转换为达梦DDL可执行类型。"""
        if not type_string:
            return type_string

        normalized = type_string.strip().upper()
        match = re.match(r'^([A-Z0-9_]+)(?:\((.*)\))?$', normalized)
        if not match:
            return normalized

        data_type, args = match.groups()
        normalized_data_type = _normalize_dm_data_type(data_type)
        if normalized_data_type in NO_ARGUMENT_DATA_TYPES:
            return normalized_data_type

        if normalized_data_type == 'BIT' and args == '1':
            return 'BIT'

        if data_type != normalized_data_type:
            return f"{normalized_data_type}({args})" if args else normalized_data_type

        return normalized

    def _get_column_sync_type_sql(self, test_column: Optional[ColumnInfo], fallback_type: str) -> str:
        """获取字段同步目标类型。"""
        if test_column:
            return self._get_column_type_sql(test_column)
        return self._normalize_type_string(fallback_type)

    def _generate_sync_sql(self, schema: str, table: str, column: str,
                           diff_type: str, test_value: str, prod_value: str,
                           test_column: Optional[ColumnInfo] = None) -> Optional[str]:
        """生成同步SQL语句

        Args:
            schema: 模式名
            table: 表名
            column: 字段名
            diff_type: 差异类型
            test_value: 源端值
            prod_value: 目标端值

        Returns:
            SQL语句字符串，如果不需要SQL则返回None
        """
        sql = None
        table_name = self._qualified_name(schema, table)
        column_name = self._quote_identifier(column)
        column_full_name = f"{table_name}.{column_name}"

        if diff_type == '类型变更':
            test_type = self._get_column_sync_type_sql(test_column, test_value)
            sql = f"ALTER TABLE {table_name} MODIFY {column_name} {test_type};"

        elif diff_type == '长度变更':
            test_type = self._get_column_sync_type_sql(test_column, f"VARCHAR({test_value})")
            sql = f"ALTER TABLE {table_name} MODIFY {column_name} {test_type};"

        elif diff_type == '精度变更':
            if test_column:
                test_type = self._get_column_type_sql(test_column)
                sql = f"ALTER TABLE {table_name} MODIFY {column_name} {test_type};"

        elif diff_type == '注释变更':
            # 双向：源端有注释→设置；源端无注释→清除目标端注释
            if test_value and test_value != '-':
                comment = test_value.replace("'", "''")
                sql = f"COMMENT ON COLUMN {column_full_name} IS '{comment}';"
            else:
                sql = f"COMMENT ON COLUMN {column_full_name} IS '';"

        elif diff_type == '默认值变更':
            default_val = self._format_default_sql(test_value)
            if default_val:
                sql = f"ALTER TABLE {table_name} ALTER COLUMN {column_name} SET DEFAULT {default_val};"
            else:
                sql = f"ALTER TABLE {table_name} ALTER COLUMN {column_name} DROP DEFAULT;"

        elif diff_type == '目标端缺失':
            # 目标端缺失该字段，需要添加
            col_type = self._get_column_sync_type_sql(test_column, test_value)
            if test_column and test_column.is_identity:
                col_type += f" IDENTITY({self._identity_args(test_column)}) NOT NULL"
            sql = f"ALTER TABLE {table_name} ADD {column_name} {col_type};"

        elif diff_type == '源端缺失':
            # 目标端独有字段，不自动生成删除SQL
            return None

        elif diff_type == '可空性变更':
            if test_column and test_column.is_primary_key:
                return None
            if test_value == 'Y':
                sql = f"ALTER TABLE {table_name} MODIFY {column_name} NULL;"
            elif test_value == 'N':
                sql = f"ALTER TABLE {table_name} MODIFY {column_name} NOT NULL;"

        return sql

    @staticmethod
    def _quote_identifier(identifier: str) -> str:
        """必要时为达梦标识符加双引号。"""
        if not identifier:
            return identifier

        normalized = identifier.upper()
        if identifier == normalized and all(ch.isascii() and (ch.isalnum() or ch in ('_', '$', '#')) for ch in identifier):
            return identifier

        return f'"{identifier.replace(chr(34), chr(34) + chr(34))}"'

    @staticmethod
    def _normalize_text(value: Optional[str]) -> str:
        """规范化SQL生成使用的元数据文本。"""
        return (value or '').strip().upper()

    @staticmethod
    def _escape_sql_literal(value: str) -> str:
        """转义SQL字符串字面量。"""
        return value.replace("'", "''")

    def _format_default_sql(self, value: Optional[str]) -> str:
        """格式化达梦字段默认值，避免给已有 SQL 字面量重复加引号。"""
        if value is None:
            return ''

        default = str(value).strip()
        if not default or default == '-':
            return ''

        upper = default.upper()
        if self._is_quoted_sql_literal(default):
            return default

        if re.fullmatch(r'[-+]?\d+(?:\.\d+)?', default):
            return default

        if upper in {
            'NULL', 'SYSDATE', 'NOW()', 'CURRENT_TIMESTAMP', 'CURRENT_DATE',
            'CURRENT_TIME', 'SYSTIMESTAMP', 'TRUE', 'FALSE'
        }:
            return default

        if '(' in default or upper.endswith('.NEXTVAL'):
            return default

        return f"'{self._escape_sql_literal(default)}'"

    @staticmethod
    def _is_quoted_sql_literal(value: str) -> bool:
        """判断字符串是否已经是 SQL 字符串字面量。"""
        return re.fullmatch(r"(?is)N?'.*'", value.strip()) is not None

    def _qualified_name(self, schema: str, name: str) -> str:
        return f"{self._quote_identifier(schema)}.{self._quote_identifier(name)}"

    def _generate_create_index_sql(self, index: IndexInfo, target_schema: Optional[str] = None) -> Optional[str]:
        """根据索引元数据生成达梦 CREATE INDEX 语句。"""
        columns = []
        for column in sorted(index.columns, key=lambda col: (col.position, col.name)):
            if column.name == "<表达式>":
                return None
            column_sql = self._quote_identifier(column.name)
            if (column.sort_order or '').upper() == 'DESC':
                column_sql += ' DESC'
            columns.append(column_sql)

        if not columns:
            return None

        index_type = self._normalize_text(index.index_type)
        uniqueness = self._normalize_text(index.uniqueness)

        index_keywords = ['CREATE']
        if uniqueness == 'UNIQUE':
            index_keywords.append('UNIQUE')
        if 'BITMAP' in index_type:
            index_keywords.append('BITMAP')
        index_keywords.append('INDEX')

        schema = target_schema or index.schema
        index_name = self._qualified_name(schema, index.name)
        table_name = self._qualified_name(schema, index.table_name)
        return f"{' '.join(index_keywords)} {index_name} ON {table_name} ({', '.join(columns)});"

    def _generate_drop_index_sql(self, schema: str, index_name: str) -> str:
        """生成达梦 DROP INDEX 语句。"""
        return f"DROP INDEX {self._qualified_name(schema, index_name)};"

    def _generate_index_sync_sql(self, diff: IndexDiff) -> Optional[str]:
        """生成索引同步SQL，以源端索引定义为目标同步目标端。"""
        if (diff.test_index and diff.test_index.is_constraint_backed()
                or diff.prod_index and diff.prod_index.is_constraint_backed()):
            return None

        if diff.diff_type == '目标端缺失':
            return self._generate_create_index_sql(diff.test_index, target_schema=diff.schema) if diff.test_index else None

        if diff.diff_type == '源端缺失':
            return None

        if diff.diff_type in ('唯一性变更', '索引类型变更', '索引列变更'):
            if not diff.test_index:
                return None
            create_sql = self._generate_create_index_sql(diff.test_index, target_schema=diff.schema)
            if not create_sql:
                return None
            drop_sql = self._generate_drop_index_sql(diff.schema, diff.index)
            return f"{drop_sql}\n{create_sql}"

        return None

    def _get_index_sync_advice(self, diff: IndexDiff) -> str:
        """给索引差异输出执行建议，避免把目标端独有索引误当成必须删除。"""
        if (diff.test_index and diff.test_index.is_constraint_backed()
                or diff.prod_index and diff.prod_index.is_constraint_backed()):
            return '约束索引，需通过约束处理'

        if diff.diff_type == '目标端缺失':
            return '目标端缺失普通索引，可按源端创建'

        if diff.diff_type == '源端缺失':
            return '目标端独有普通索引，需人工确认；默认不生成DROP SQL'

        if diff.diff_type in ('唯一性变更', '索引类型变更', '索引列变更'):
            return '普通索引定义不一致，确认后按源端重建'

        return ''

    def _create_column_diff_sheet(self, wb: Workbook, style: dict):
        """创建字段差异sheet"""
        ws = wb.create_sheet('字段差异')

        # 标题
        headers = ['模式名', '表名', '字段名', '差异类型',
                    f'{self.result.source_label}值', f'{self.result.target_label}值', '同步SQL']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = style['title_font']
            cell.fill = style['title_fill']
            cell.alignment = style['title_alignment']
            cell.border = style['thin_border']

        # 数据行
        for row, diff in enumerate(self.result.column_diffs, 2):
            ws.cell(row=row, column=1, value=diff.schema).border = style['thin_border']
            ws.cell(row=row, column=2, value=diff.table).border = style['thin_border']
            ws.cell(row=row, column=3, value=diff.column).border = style['thin_border']
            cell = ws.cell(row=row, column=4, value=diff.diff_type)
            cell.border = style['thin_border']
            # 根据差异类型设置颜色
            if diff.diff_type == '源端缺失':
                cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
            elif diff.diff_type == '目标端缺失':
                cell.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
            ws.cell(row=row, column=5, value=diff.test_value).border = style['thin_border']
            ws.cell(row=row, column=6, value=diff.prod_value).border = style['thin_border']

            # 生成并写入SQL
            sql = self._generate_sync_sql(diff.schema, diff.table, diff.column,
                                          diff.diff_type, diff.test_value, diff.prod_value,
                                          diff.test_column)
            if sql:
                sql_cell = ws.cell(row=row, column=7, value=sql)
                sql_cell.border = style['thin_border']
                sql_cell.alignment = Alignment(wrap_text=True, vertical='top')

        # 调整列宽
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 25
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 30
        ws.column_dimensions['F'].width = 30
        ws.column_dimensions['G'].width = 120
        self._apply_table_features(ws)

    def _create_index_diff_sheet(self, wb: Workbook, style: dict):
        """创建索引差异sheet"""
        ws = wb.create_sheet('索引差异')

        headers = ['模式名', '表名', '索引名', '差异类型',
                    f'{self.result.source_label}值', f'{self.result.target_label}值', '同步建议', '同步SQL']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = style['title_font']
            cell.fill = style['title_fill']
            cell.alignment = style['title_alignment']
            cell.border = style['thin_border']

        for row, diff in enumerate(self.result.index_diffs, 2):
            ws.cell(row=row, column=1, value=diff.schema).border = style['thin_border']
            ws.cell(row=row, column=2, value=diff.table).border = style['thin_border']
            ws.cell(row=row, column=3, value=diff.index).border = style['thin_border']
            cell = ws.cell(row=row, column=4, value=diff.diff_type)
            cell.border = style['thin_border']
            if diff.diff_type == '源端缺失':
                cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
            elif diff.diff_type == '目标端缺失':
                cell.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
            ws.cell(row=row, column=5, value=diff.test_value).border = style['thin_border']
            ws.cell(row=row, column=6, value=diff.prod_value).border = style['thin_border']
            ws.cell(row=row, column=7, value=self._get_index_sync_advice(diff)).border = style['thin_border']
            sql = self._generate_index_sync_sql(diff)
            sql_cell = ws.cell(row=row, column=8, value=sql or '')
            sql_cell.border = style['thin_border']
            sql_cell.alignment = Alignment(wrap_text=True, vertical='top')

        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 35
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 45
        ws.column_dimensions['F'].width = 45
        ws.column_dimensions['G'].width = 45
        ws.column_dimensions['H'].width = 120
        self._apply_table_features(ws)

    def _generate_pk_sync_sql(self, diff: PrimaryKeyDiff) -> Optional[str]:
        """生成主键同步 SQL，以源端主键为目标。"""
        table_name = self._qualified_name(diff.schema, diff.table)
        if diff.diff_type == '目标端缺失':
            cols = self._format_pk_columns(diff.test_value)
            return f"ALTER TABLE {table_name} ADD PRIMARY KEY ({cols});"
        if diff.diff_type == '主键列变更':
            cols = self._format_pk_columns(diff.test_value)
            return (
                f"ALTER TABLE {table_name} DROP PRIMARY KEY;\n"
                f"ALTER TABLE {table_name} ADD PRIMARY KEY ({cols});"
            )
        # 源端缺失：目标端独有主键，默认不生成 DROP，需人工确认
        return None

    @staticmethod
    def _format_pk_columns(pk_value: str) -> str:
        """从主键差异值（逗号分隔列名）还原引号包裹的列定义。

        注：primary_key_columns 按 ALL_CONS_COLUMNS.POSITION 保留源端顺序。
        """
        cols = [c.strip() for c in (pk_value or '').split(',') if c.strip()]
        return ', '.join(ReportGenerator._quote_identifier(c) for c in cols)

    def _get_pk_sync_advice(self, diff: PrimaryKeyDiff) -> str:
        if diff.diff_type == '目标端缺失':
            return '目标端缺失主键，可按源端添加'
        if diff.diff_type == '主键列变更':
            return '主键列不一致，确认后按源端重建（DROP+ADD）'
        if diff.diff_type == '源端缺失':
            return '目标端独有主键，需人工确认；默认不生成DROP SQL'
        return ''

    def _create_primary_key_diff_sheet(self, wb: Workbook, style: dict):
        """创建主键差异sheet"""
        ws = wb.create_sheet('主键差异')

        sl = self.result.source_label
        tl = self.result.target_label
        headers = ['模式名', '表名', '差异类型', f'{sl}主键', f'{tl}主键', '同步建议', '同步SQL']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = style['title_font']
            cell.fill = style['title_fill']
            cell.alignment = style['title_alignment']
            cell.border = style['thin_border']

        for row, diff in enumerate(self.result.primary_key_diffs, 2):
            ws.cell(row=row, column=1, value=diff.schema).border = style['thin_border']
            ws.cell(row=row, column=2, value=diff.table).border = style['thin_border']
            cell = ws.cell(row=row, column=3, value=diff.diff_type)
            cell.border = style['thin_border']
            if diff.diff_type == '源端缺失':
                cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
            elif diff.diff_type == '目标端缺失':
                cell.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
            elif diff.diff_type == '主键列变更':
                cell.fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
            ws.cell(row=row, column=4, value=diff.test_value).border = style['thin_border']
            ws.cell(row=row, column=5, value=diff.prod_value).border = style['thin_border']
            ws.cell(row=row, column=6, value=self._get_pk_sync_advice(diff)).border = style['thin_border']
            sql_cell = ws.cell(row=row, column=7, value=self._generate_pk_sync_sql(diff) or '')
            sql_cell.border = style['thin_border']
            sql_cell.alignment = Alignment(wrap_text=True, vertical='top')

        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 35
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 45
        ws.column_dimensions['E'].width = 45
        ws.column_dimensions['F'].width = 45
        ws.column_dimensions['G'].width = 120
        self._apply_table_features(ws)

    # ---------- 序列差异 ----------

    def _generate_sequence_create_sql(self, seq: SequenceInfo, target_schema: str) -> str:
        """按序列属性生成 CREATE SEQUENCE 语句。

        START WITH 取建序列起始值（来自 GET_DDL，非 LAST_NUMBER 当前值）；
        同步的是定义，不是运行时状态。
        """
        cache_clause = 'NOCACHE' if (seq.cache_size or 0) == 0 else f'CACHE {seq.cache_size}'
        cycle_clause = 'CYCLE' if (seq.cycle_flag or '').upper() == 'Y' else 'NOCYCLE'
        order_clause = 'ORDER' if (seq.order_flag or '').upper() == 'Y' else 'NOORDER'
        qualified = self._qualified_name(target_schema, seq.name)
        start_clause = f'START WITH {seq.start_with}' if seq.start_with is not None else ''
        parts = [
            f"CREATE SEQUENCE {qualified}",
            f"INCREMENT BY {seq.increment_by or 1}",
            start_clause,
            f"MINVALUE {seq.min_value or 1}",
            f"MAXVALUE {seq.max_value or 999999999999999999999999999}",
            cycle_clause,
            cache_clause,
            order_clause,
        ]
        return ' '.join(part for part in parts if part) + ';'

    def _generate_sequence_alter_sql(self, seq: SequenceInfo, target_schema: str) -> str:
        """按源端属性生成 ALTER SEQUENCE 语句。ALTER 不能改 START WITH，与我们的用法一致。"""
        cache_clause = 'NOCACHE' if (seq.cache_size or 0) == 0 else f'CACHE {seq.cache_size}'
        cycle_clause = 'CYCLE' if (seq.cycle_flag or '').upper() == 'Y' else 'NOCYCLE'
        order_clause = 'ORDER' if (seq.order_flag or '').upper() == 'Y' else 'NOORDER'
        qualified = self._qualified_name(target_schema, seq.name)
        return (
            f"ALTER SEQUENCE {qualified} "
            f"INCREMENT BY {seq.increment_by or 1} "
            f"MINVALUE {seq.min_value or 1} "
            f"MAXVALUE {seq.max_value or 999999999999999999999999999} "
            f"{cycle_clause} {cache_clause} {order_clause};"
        )

    def _generate_sequence_sync_sql(self, diff: SequenceDiff) -> Optional[str]:
        """生成序列同步 SQL，以源端定义为目标。"""
        if diff.diff_type == '目标端缺失':
            return self._generate_sequence_create_sql(diff.test_sequence, diff.schema) if diff.test_sequence else None
        if diff.diff_type == '属性变更':
            return self._generate_sequence_alter_sql(diff.test_sequence, diff.schema) if diff.test_sequence else None
        # 源端缺失：目标端独有序列，默认不生成 DROP，需人工确认
        return None

    def _get_sequence_sync_advice(self, diff: SequenceDiff) -> str:
        if diff.diff_type == '目标端缺失':
            return '目标端缺失序列，可按源端创建（含 START WITH 起始值；当前值不同步）'
        if diff.diff_type == '属性变更':
            extra = ''
            if (diff.test_sequence and diff.prod_sequence
                    and (diff.test_sequence.start_with or 0) != (diff.prod_sequence.start_with or 0)):
                extra = '；START WITH 不同，ALTER 改不了，需手动 DROP+CREATE'
            return '序列属性不一致，确认后按源端 ALTER' + extra
        if diff.diff_type == '源端缺失':
            return '目标端独有序列，需人工确认；默认不生成 DROP SQL'
        return ''

    def _create_sequence_diff_sheet(self, wb: Workbook, style: dict):
        """创建序列差异 sheet"""
        ws = wb.create_sheet('序列差异')

        sl = self.result.source_label
        tl = self.result.target_label
        headers = ['模式名', '序列名', '差异类型',
                   f'{sl}属性', f'{tl}属性', '同步建议', '同步SQL']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = style['title_font']
            cell.fill = style['title_fill']
            cell.alignment = style['title_alignment']
            cell.border = style['thin_border']

        for row, diff in enumerate(self.result.sequence_diffs, 2):
            ws.cell(row=row, column=1, value=diff.schema).border = style['thin_border']
            ws.cell(row=row, column=2, value=diff.sequence).border = style['thin_border']
            cell = ws.cell(row=row, column=3, value=diff.diff_type)
            cell.border = style['thin_border']
            if diff.diff_type == '源端缺失':
                cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
            elif diff.diff_type == '目标端缺失':
                cell.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
            elif diff.diff_type == '属性变更':
                cell.fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
            ws.cell(row=row, column=4, value=diff.test_value).border = style['thin_border']
            ws.cell(row=row, column=5, value=diff.prod_value).border = style['thin_border']
            ws.cell(row=row, column=6, value=self._get_sequence_sync_advice(diff)).border = style['thin_border']
            sql_cell = ws.cell(row=row, column=7, value=self._generate_sequence_sync_sql(diff) or '')
            sql_cell.border = style['thin_border']
            sql_cell.alignment = Alignment(wrap_text=True, vertical='top')

        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 35
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 55
        ws.column_dimensions['E'].width = 55
        ws.column_dimensions['F'].width = 45
        ws.column_dimensions['G'].width = 120
        self._apply_table_features(ws)

    # ---------- 外键差异 ----------

    def _generate_fk_create_sql(self, fk: ForeignKeyInfo, target_schema: str) -> str:
        """生成 ALTER TABLE ADD CONSTRAINT FOREIGN KEY 语句。"""
        table_name = self._qualified_name(target_schema, fk.table_name)
        cols = ', '.join(self._quote_identifier(c.column_name) for c in sorted(fk.columns, key=lambda x: x.position))
        ref_cols = ', '.join(self._quote_identifier(c.ref_column_name) for c in sorted(fk.columns, key=lambda x: x.position))
        ref_table = self._qualified_name(fk.ref_schema, fk.ref_table_name)
        constraint_name = self._quote_identifier(fk.name)
        sql = f"ALTER TABLE {table_name} ADD CONSTRAINT {constraint_name} FOREIGN KEY ({cols}) REFERENCES {ref_table} ({ref_cols})"
        if fk.delete_rule and fk.delete_rule != 'NO ACTION':
            sql += f" ON DELETE {fk.delete_rule}"
        return sql + ';'

    def _generate_fk_sync_sql(self, diff: ForeignKeyDiff) -> Optional[str]:
        """生成外键同步 SQL。"""
        if diff.diff_type == '目标端缺失' and diff.test_fk:
            return self._generate_fk_create_sql(diff.test_fk, diff.schema)
        if diff.diff_type == '定义变更' and diff.test_fk:
            drop = f"ALTER TABLE {self._qualified_name(diff.schema, diff.table)} DROP CONSTRAINT {self._quote_identifier(diff.fk_name)};"
            create = self._generate_fk_create_sql(diff.test_fk, diff.schema)
            return f"{drop}\n{create}"
        return None

    @staticmethod
    def _get_fk_sync_advice(diff: ForeignKeyDiff) -> str:
        if diff.diff_type == '目标端缺失':
            return '目标端缺失外键，可按源端添加'
        if diff.diff_type == '定义变更':
            return '外键定义不一致，确认后按源端重建（DROP+ADD）'
        if diff.diff_type == '源端缺失':
            return '目标端独有外键，需人工确认；默认不生成 DROP SQL'
        return ''

    def _create_foreign_key_diff_sheet(self, wb: Workbook, style: dict):
        """创建外键差异 sheet"""
        ws = wb.create_sheet('外键差异')

        sl = self.result.source_label
        tl = self.result.target_label
        headers = ['模式名', '表名', '外键名', '差异类型',
                   f'{sl}定义', f'{tl}定义', '同步建议', '同步SQL']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = style['title_font']
            cell.fill = style['title_fill']
            cell.alignment = style['title_alignment']
            cell.border = style['thin_border']

        for row, diff in enumerate(self.result.foreign_key_diffs, 2):
            ws.cell(row=row, column=1, value=diff.schema).border = style['thin_border']
            ws.cell(row=row, column=2, value=diff.table).border = style['thin_border']
            ws.cell(row=row, column=3, value=diff.fk_name).border = style['thin_border']
            cell = ws.cell(row=row, column=4, value=diff.diff_type)
            cell.border = style['thin_border']
            if diff.diff_type == '源端缺失':
                cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
            elif diff.diff_type == '目标端缺失':
                cell.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
            elif diff.diff_type == '定义变更':
                cell.fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
            ws.cell(row=row, column=5, value=diff.test_value).border = style['thin_border']
            ws.cell(row=row, column=6, value=diff.prod_value).border = style['thin_border']
            ws.cell(row=row, column=7, value=self._get_fk_sync_advice(diff)).border = style['thin_border']
            sql_cell = ws.cell(row=row, column=8, value=self._generate_fk_sync_sql(diff) or '')
            sql_cell.border = style['thin_border']
            sql_cell.alignment = Alignment(wrap_text=True, vertical='top')

        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 35
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 55
        ws.column_dimensions['F'].width = 55
        ws.column_dimensions['G'].width = 45
        ws.column_dimensions['H'].width = 120
        self._apply_table_features(ws)

    # ---------- 视图差异 ----------

    def _create_view_diff_sheet(self, wb: Workbook, style: dict):
        """创建视图差异 sheet"""
        ws = wb.create_sheet('视图差异')

        sl = self.result.source_label
        tl = self.result.target_label
        headers = ['模式名', '视图名', '差异类型', f'{sl}定义', f'{tl}定义']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = style['title_font']
            cell.fill = style['title_fill']
            cell.alignment = style['title_alignment']
            cell.border = style['thin_border']

        for row, diff in enumerate(self.result.view_diffs, 2):
            ws.cell(row=row, column=1, value=diff.schema).border = style['thin_border']
            ws.cell(row=row, column=2, value=diff.view).border = style['thin_border']
            cell = ws.cell(row=row, column=3, value=diff.diff_type)
            cell.border = style['thin_border']
            if diff.diff_type == '源端缺失':
                cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
            elif diff.diff_type == '目标端缺失':
                cell.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
            elif diff.diff_type == '定义变更':
                cell.fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
            ws.cell(row=row, column=4, value=diff.test_value).border = style['thin_border']
            ws.cell(row=row, column=4).alignment = Alignment(wrap_text=True, vertical='top')
            ws.cell(row=row, column=5, value=diff.prod_value).border = style['thin_border']
            ws.cell(row=row, column=5).alignment = Alignment(wrap_text=True, vertical='top')

        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 35
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 80
        ws.column_dimensions['E'].width = 80
        self._apply_table_features(ws)

    # ---------- 存储过程/函数/触发器差异 ----------

    def _create_procedure_diff_sheet(self, wb: Workbook, style: dict):
        """创建存储过程/函数/触发器差异 sheet"""
        ws = wb.create_sheet('过程函数触发器差异')

        sl = self.result.source_label
        tl = self.result.target_label
        headers = ['模式名', '对象类型', '对象名', '差异类型', f'{sl}定义', f'{tl}定义']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = style['title_font']
            cell.fill = style['title_fill']
            cell.alignment = style['title_alignment']
            cell.border = style['thin_border']

        for row, diff in enumerate(self.result.procedure_diffs, 2):
            ws.cell(row=row, column=1, value=diff.schema).border = style['thin_border']
            ws.cell(row=row, column=2, value=diff.object_type).border = style['thin_border']
            ws.cell(row=row, column=3, value=diff.name).border = style['thin_border']
            cell = ws.cell(row=row, column=4, value=diff.diff_type)
            cell.border = style['thin_border']
            if diff.diff_type == '源端缺失':
                cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
            elif diff.diff_type == '目标端缺失':
                cell.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
            elif diff.diff_type == '定义变更':
                cell.fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
            ws.cell(row=row, column=5, value=diff.test_value).border = style['thin_border']
            ws.cell(row=row, column=5).alignment = Alignment(wrap_text=True, vertical='top')
            ws.cell(row=row, column=6, value=diff.prod_value).border = style['thin_border']
            ws.cell(row=row, column=6).alignment = Alignment(wrap_text=True, vertical='top')

        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 35
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 80
        ws.column_dimensions['F'].width = 80
        self._apply_table_features(ws)


def load_config(config_path: str) -> configparser.ConfigParser:
    """加载配置文件"""
    config = configparser.ConfigParser(interpolation=None)
    with open(config_path, encoding='utf-8') as f:
        config.read_file(f)
    return config


def parse_exclude_schemas(schemas_str: str) -> Set[str]:
    """解析排除的模式列表"""
    if not schemas_str:
        return set()
    return set(s.strip() for s in schemas_str.split(',') if s.strip())


# 框架元数据序列默认排除：Liquibase 建表时附带、但实际无列引用的悬空序列。
# 见 ALL_TAB_COLUMNS 无 .NEXTVAL 引用，属纯噪音。
DEFAULT_EXCLUDE_SEQUENCES: List[str] = ['DATABASECHANGELOG*', 'DATABASECHANGELOGLOCK*']


def parse_exclude_sequences(patterns_str: str) -> List[str]:
    """解析排除的序列名模式列表，支持 glob 通配符(* ?)，大小写不敏感。"""
    if not patterns_str:
        return []
    return [p.strip().upper() for p in patterns_str.split(',') if p.strip()]


def _is_name_excluded(name: str, patterns: Optional[List[str]]) -> bool:
    """按 glob 模式判断对象名是否应排除（大小写不敏感，跨平台一致）。"""
    if not patterns:
        return False
    upper_name = (name or '').upper()
    return any(fnmatch.fnmatchcase(upper_name, p) for p in patterns)


def parse_schema_mapping(mapping_str: str) -> Dict[str, str]:
    """解析模式名映射配置
    格式: 源端模式名1:目标端模式名1,源端模式名2:目标端模式名2
    """
    mapping = {}
    if not mapping_str:
        return mapping

    for pair in mapping_str.split(','):
        pair = pair.strip()
        if ':' not in pair:
            logger.warning("无效的映射配置 '%s'，格式应为 '源端名:目标端名'", pair)
            continue
        test_name, prod_name = pair.split(':', 1)
        test_name = test_name.strip()
        prod_name = prod_name.strip()
        if test_name and prod_name:
            mapping[test_name] = prod_name
        else:
            logger.warning("映射配置 '%s' 中有空名称，已忽略", pair)

    return mapping


def load_prod_schema_configs(config: configparser.ConfigParser) -> List[SchemaConnectionConfig]:
    """加载目标端的多个模式配置

    从配置文件中读取所有 prod_env_* 开头的配置段
    每个配置段包含: host, port, schema, user, password
    每个模式都是独立的数据库实例
    """
    configs = []

    # 遍历所有配置段，找出 prod_env_* 开头的段
    for section_name in config.sections():
        if section_name.startswith('prod_env_'):
            try:
                host = config.get(section_name, 'host')
                port = config.getint(section_name, 'port')
                schema = config.get(section_name, 'schema')
                user = config.get(section_name, 'user')
                password = config.get(section_name, 'password')

                configs.append(SchemaConnectionConfig(
                    schema=schema,
                    user=user,
                    password=password,
                    host=host,
                    port=port
                ))
            except Exception as e:
                logger.warning("配置段 [%s] 读取失败: %s，已跳过", section_name, e)
                continue

    if not configs:
        raise ValueError("未找到任何目标端配置（prod_env_* 开头的配置段）")

    logger.info("  加载了 %d 个目标端模式配置", len(configs))
    for cfg in configs:
        logger.info("    - %s (%s:%d, 用户: %s)", cfg.schema, cfg.host, cfg.port, cfg.user)

    return configs


def load_schema_configs(config: configparser.ConfigParser, prefix: str) -> List[SchemaConnectionConfig]:
    """通用：加载以 prefix 开头的配置段作为多模式连接。

    Args:
        config: 配置解析器
        prefix: section 前缀，如 'source_env_' 或 'target_env_'
    """
    configs = []
    for section_name in config.sections():
        if section_name.startswith(prefix):
            try:
                host = config.get(section_name, 'host')
                port = config.getint(section_name, 'port')
                schema = config.get(section_name, 'schema')
                user = config.get(section_name, 'user')
                password = config.get(section_name, 'password')
                configs.append(SchemaConnectionConfig(
                    schema=schema, user=user, password=password, host=host, port=port
                ))
            except Exception as e:
                logger.warning("配置段 [%s] 读取失败: %s，已跳过", section_name, e)
                continue
    logger.info("  加载了 %d 个 %s 模式配置", len(configs), prefix.rstrip('_'))
    for cfg in configs:
        logger.info("    - %s (%s:%d, 用户: %s)", cfg.schema, cfg.host, cfg.port, cfg.user)
    return configs


def _detect_config_mode(config: configparser.ConfigParser) -> str:
    """检测配置文件格式：'legacy'（test_env + prod_env_*）或 'dual'（source_env_* + target_env_*）。"""
    has_source = any(s.startswith('source_env_') for s in config.sections())
    has_target = any(s.startswith('target_env_') for s in config.sections())
    if has_source and has_target:
        return 'dual'
    return 'legacy'


def apply_schema_mapping(metadata: Dict[str, SchemaInfo],
                         mapping: Dict[str, str],
                         env_name: str = "源端") -> Tuple[Dict[str, SchemaInfo], Dict[str, str]]:
    """应用模式名映射

    Args:
        metadata: 原始元数据
        mapping: 映射关系 {源端模式名: 目标端模式名}
        env_name: 环境名称（用于日志）

    Returns:
        (重命名后的元数据, 实际应用的映射记录)
    """
    if not mapping:
        return metadata, {}

    new_metadata = {}
    applied_mapping = {}

    for old_schema_name, schema_info in metadata.items():
        if old_schema_name in mapping:
            # 应用映射
            new_schema_name = mapping[old_schema_name]
            new_metadata[new_schema_name] = schema_info
            schema_info.name = new_schema_name
            applied_mapping[old_schema_name] = new_schema_name
            logger.info("  [%s] 模式 '%s' 映射为 '%s'", env_name, old_schema_name, new_schema_name)
        else:
            # 保持原名
            new_metadata[old_schema_name] = schema_info

    # 重映射外键引用的模式名（ref_schema），避免跨环境 schema 不同名时误报差异
    for schema_info in new_metadata.values():
        for table_info in schema_info.tables.values():
            for fk in table_info.foreign_keys.values():
                if fk.ref_schema in mapping:
                    fk.ref_schema = mapping[fk.ref_schema]

    if applied_mapping:
        logger.info("  [%s] 共应用了 %d 个模式映射", env_name, len(applied_mapping))

    return new_metadata, applied_mapping


def _ensure_logging(script_dir: str):
    """确保日志已配置（幂等，多次调用不会重复添加 handler）。"""
    if logger.handlers:
        return

    log_dir = os.path.join(script_dir, 'logs')
    os.makedirs(log_dir, exist_ok=True)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = os.path.join(log_dir, f'dm_compare_{timestamp}.log')

    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S'))

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S'))

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    logger.setLevel(logging.DEBUG)

    logger.info("日志文件: %s", log_file)


def run_compare(config_path: str, progress_callback: Optional[Callable[[str], None]] = None) -> CompareRunResult:
    """执行数据库结构比对，返回比对结果和Excel报告路径。

    支持两种配置格式:
    - legacy: [test_env] 单连接 + [prod_env_*] 多连接
    - dual:   [source_env_*] 多连接 + [target_env_*] 多连接
    """
    script_dir = os.path.dirname(os.path.abspath(config_path))
    _ensure_logging(script_dir)

    def emit(message: str):
        print(message)
        logger.info(message)
        if progress_callback:
            progress_callback(message)

    _step = [0]

    def step(message: str):
        """输出带动态步骤编号的消息"""
        _step[0] += 1
        emit(f"[{_step[0]}] {message}")

    step("加载配置文件...")
    config = load_config(config_path)
    mode = _detect_config_mode(config)
    emit(f"  配置模式: {mode}")

    # 解析公共配置
    exclude_schemas = parse_exclude_schemas(config.get('options', 'exclude_schemas', fallback=''))
    exclude_sequence_patterns = list(dict.fromkeys(
        DEFAULT_EXCLUDE_SEQUENCES
        + parse_exclude_sequences(config.get('options', 'exclude_sequences', fallback=''))
    ))
    schema_mapping = parse_schema_mapping(config.get('options', 'schema_mapping', fallback=''))
    output_dir = config.get('options', 'output_dir', fallback='./reports')
    source_label = config.get('options', 'source_label', fallback='源端')
    target_label = config.get('options', 'target_label', fallback='目标端')

    if schema_mapping:
        emit(f"  配置的模式映射: {len(schema_mapping)} 条")
        for test_name, prod_name in schema_mapping.items():
            emit(f"    {test_name} -> {prod_name}")

    # ---------- 源端连接 & 元数据 ----------
    source_conns: List[DMConnection] = []
    all_conns: List[DMConnection] = []  # 所有需要关闭的连接

    if mode == 'dual':
        source_configs = load_schema_configs(config, 'source_env_')
        if not source_configs:
            raise ValueError("未找到任何源端配置（source_env_* 开头的配置段）")

        step(f"连接 {source_label} 数据库（{len(source_configs)} 个连接）...")
        source_hosts = []
        for i, schema_cfg in enumerate(source_configs, 1):
            emit(f"  [{i}/{len(source_configs)}] 连接模式 {schema_cfg.schema} ({schema_cfg.host}:{schema_cfg.port})...")
            conn = DMConnection(schema_cfg.host, schema_cfg.port, schema_cfg.user, schema_cfg.password)
            if not conn.connect():
                for c in all_conns:
                    c.close()
                raise RuntimeError(f"{source_label} 模式 {schema_cfg.schema} 连接失败: {schema_cfg.host}:{schema_cfg.port}")
            source_conns.append(conn)
            all_conns.append(conn)
            source_hosts.append(f"{schema_cfg.schema}@{schema_cfg.host}:{schema_cfg.port}")
            emit(f"    模式 {schema_cfg.schema} 连接成功")

        step(f"提取 {source_label} 元数据...")
        test_metadata = {}
        for i, (conn, schema_cfg) in enumerate(zip(source_conns, source_configs), 1):
            emit(f"  [{i}/{len(source_conns)}] 从连接提取模式 {schema_cfg.schema}...")
            extractor = DMMetadataExtractor(conn)
            schema_info = extractor.extract_single_schema(schema_cfg.schema, exclude_sequence_patterns)
            test_metadata[schema_cfg.schema] = schema_info

        test_host, test_port = '', 0
    else:
        test_host = config.get('test_env', 'host')
        test_port = config.getint('test_env', 'port')
        test_user = config.get('test_env', 'user')
        test_password = config.get('test_env', 'password')

        step(f"连接 {source_label} 数据库 ({test_host}:{test_port})...")
        test_conn = DMConnection(test_host, test_port, test_user, test_password)
        if not test_conn.connect():
            raise RuntimeError(f"{source_label} 数据库连接失败: {test_host}:{test_port}")
        all_conns.append(test_conn)
        source_hosts = []
        emit(f"  {source_label} 连接成功")

        step(f"提取 {source_label} 元数据...")
        test_extractor = DMMetadataExtractor(test_conn)
        test_metadata = test_extractor.extract_all_metadata(exclude_schemas, exclude_sequence_patterns)

    # ---------- 目标端连接 & 元数据 ----------
    if mode == 'dual':
        prod_schema_configs = load_schema_configs(config, 'target_env_')
    else:
        prod_schema_configs = load_prod_schema_configs(config)

    step(f"连接 {target_label} 数据库（{len(prod_schema_configs)} 个连接）...")
    prod_conns = []
    prod_hosts = []

    for i, schema_cfg in enumerate(prod_schema_configs, 1):
        emit(f"  [{i}/{len(prod_schema_configs)}] 连接模式 {schema_cfg.schema} ({schema_cfg.host}:{schema_cfg.port})...")
        conn = DMConnection(schema_cfg.host, schema_cfg.port, schema_cfg.user, schema_cfg.password)
        if not conn.connect():
            for c in all_conns:
                c.close()
            raise RuntimeError(f"{target_label} 模式 {schema_cfg.schema} 连接失败: {schema_cfg.host}:{schema_cfg.port}")
        prod_conns.append(conn)
        all_conns.append(conn)
        prod_hosts.append(f"{schema_cfg.schema}@{schema_cfg.host}:{schema_cfg.port}")
        emit(f"    模式 {schema_cfg.schema} 连接成功")

    try:
        step(f"提取 {target_label} 元数据...")
        prod_metadata = {}
        for i, (conn, schema_cfg) in enumerate(zip(prod_conns, prod_schema_configs), 1):
            emit(f"  [{i}/{len(prod_conns)}] 从连接提取模式 {schema_cfg.schema}...")
            extractor = DMMetadataExtractor(conn)
            schema_info = extractor.extract_single_schema(schema_cfg.schema, exclude_sequence_patterns)
            prod_metadata[schema_cfg.schema] = schema_info

        # 应用模式映射
        applied_mappings = {}
        if schema_mapping:
            emit("[应用模式映射]")
            test_metadata, applied_mappings = apply_schema_mapping(test_metadata, schema_mapping, source_label)

        # 执行对比
        step("执行差异对比...")
        comparator = SchemaComparator(test_metadata, prod_metadata)
        result = comparator.compare_full(
            test_host, test_port, prod_hosts,
            source_hosts=source_hosts, target_hosts=prod_hosts,
            source_label=source_label, target_label=target_label,
        )
        result.schema_mappings = applied_mappings

        # 生成报告
        step("生成Excel报告...")
        if not os.path.isabs(output_dir):
            output_dir = os.path.join(script_dir, output_dir)
        output_dir = os.path.normpath(output_dir)
        os.makedirs(output_dir, exist_ok=True)

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        report_path = os.path.join(output_dir, f'dm_compare_report_{timestamp}.xlsx')

        generator = ReportGenerator(result)
        generator.generate_excel(report_path)

        emit(f"完成！报告已保存至: {report_path}")
        emit(f"  - 表差异: {len(result.table_diffs)}个")
        emit(f"  - 字段差异: {len(result.column_diffs)}个")
        emit(f"  - 字段差异: {len(result.column_diffs)}个")
        emit(f"  - 索引差异: {len(result.index_diffs)}个")
        emit(f"  - 主键差异: {len(result.primary_key_diffs)}个")
        emit(f"  - 序列差异: {len(result.sequence_diffs)}个")
        emit(f"  - 外键差异: {len(result.foreign_key_diffs)}个")
        emit(f"  - 视图差异: {len(result.view_diffs)}个")
        emit(f"  - 过程/函数/触发器差异: {len(result.procedure_diffs)}个")
        return CompareRunResult(result=result, report_path=report_path, output_dir=output_dir)

    finally:
        for conn in all_conns:
            conn.close()
        emit("数据库连接已关闭")


def build_summary_rows(result: ComparisonResult) -> List[Dict[str, Any]]:
    """构建页面展示用的汇总行。"""
    sl = result.source_label
    tl = result.target_label
    return [
        {'项目': '模式数量', sl: result.test_schema_count, tl: result.prod_schema_count, '差异数量': '-'},
        {'项目': '共同模式数', sl: result.common_schema_count, tl: result.common_schema_count, '差异数量': '-'},
        {'项目': '表数量', sl: result.test_table_count, tl: result.prod_table_count, '差异数量': len(result.table_diffs)},
        {'项目': '共同表数量', sl: result.common_table_count, tl: result.common_table_count, '差异数量': '-'},
        {'项目': '字段差异', sl: '-', tl: '-', '差异数量': len(result.column_diffs)},
        {'项目': '索引数量', sl: result.test_index_count, tl: result.prod_index_count, '差异数量': '-'},
        {'项目': '共同索引数量', sl: result.common_index_count, tl: result.common_index_count, '差异数量': '-'},
        {'项目': '索引差异', sl: '-', tl: '-', '差异数量': len(result.index_diffs)},
        {'项目': '主键差异', sl: '-', tl: '-', '差异数量': len(result.primary_key_diffs)},
        {'项目': '序列数量', sl: result.test_sequence_count, tl: result.prod_sequence_count, '差异数量': '-'},
        {'项目': '共同序列数量', sl: result.common_sequence_count, tl: result.common_sequence_count, '差异数量': '-'},
        {'项目': '序列差异', sl: '-', tl: '-', '差异数量': len(result.sequence_diffs)},
        {'项目': '外键数量', sl: result.test_fk_count, tl: result.prod_fk_count, '差异数量': '-'},
        {'项目': '外键差异', sl: '-', tl: '-', '差异数量': len(result.foreign_key_diffs)},
        {'项目': '视图数量', sl: result.test_view_count, tl: result.prod_view_count, '差异数量': '-'},
        {'项目': '视图差异', sl: '-', tl: '-', '差异数量': len(result.view_diffs)},
        {'项目': '过程/函数/触发器数量', sl: result.test_procedure_count, tl: result.prod_procedure_count, '差异数量': '-'},
        {'项目': '过程/函数/触发器差异', sl: '-', tl: '-', '差异数量': len(result.procedure_diffs)},
    ]


def build_table_diff_rows(result: ComparisonResult) -> List[Dict[str, Any]]:
    generator = ReportGenerator(result)
    return [
        {
            '模式名': diff.schema,
            '表名': diff.table,
            '差异类型': diff.diff_type,
            '表注释': diff.comment,
            '同步SQL': generator._generate_table_sync_sql(diff) or ''
        }
        for diff in result.table_diffs
    ]


def build_column_diff_rows(result: ComparisonResult) -> List[Dict[str, Any]]:
    sl = result.source_label
    tl = result.target_label
    generator = ReportGenerator(result)
    return [
        {
            '模式名': diff.schema,
            '表名': diff.table,
            '字段名': diff.column,
            '差异类型': diff.diff_type,
            f'{sl}值': diff.test_value,
            f'{tl}值': diff.prod_value,
            '同步SQL': generator._generate_sync_sql(diff.schema, diff.table, diff.column,
                                                   diff.diff_type, diff.test_value, diff.prod_value,
                                                   diff.test_column) or ''
        }
        for diff in result.column_diffs
    ]


def build_index_diff_rows(result: ComparisonResult) -> List[Dict[str, Any]]:
    sl = result.source_label
    tl = result.target_label
    generator = ReportGenerator(result)
    return [
        {
            '模式名': diff.schema,
            '表名': diff.table,
            '索引名': diff.index,
            '差异类型': diff.diff_type,
            f'{sl}值': diff.test_value,
            f'{tl}值': diff.prod_value,
            '同步建议': generator._get_index_sync_advice(diff),
            '同步SQL': generator._generate_index_sync_sql(diff) or ''
        }
        for diff in result.index_diffs
    ]


def build_primary_key_diff_rows(result: ComparisonResult) -> List[Dict[str, Any]]:
    sl = result.source_label
    tl = result.target_label
    generator = ReportGenerator(result)
    return [
        {
            '模式名': diff.schema,
            '表名': diff.table,
            '差异类型': diff.diff_type,
            f'{sl}主键': diff.test_value,
            f'{tl}主键': diff.prod_value,
            '同步建议': generator._get_pk_sync_advice(diff),
            '同步SQL': generator._generate_pk_sync_sql(diff) or '',
        }
        for diff in result.primary_key_diffs
    ]


def build_sequence_diff_rows(result: ComparisonResult) -> List[Dict[str, Any]]:
    sl = result.source_label
    tl = result.target_label
    generator = ReportGenerator(result)
    return [
        {
            '模式名': diff.schema,
            '序列名': diff.sequence,
            '差异类型': diff.diff_type,
            f'{sl}属性': diff.test_value,
            f'{tl}属性': diff.prod_value,
            '同步建议': generator._get_sequence_sync_advice(diff),
            '同步SQL': generator._generate_sequence_sync_sql(diff) or '',
        }
        for diff in result.sequence_diffs
    ]


def build_foreign_key_diff_rows(result: ComparisonResult) -> List[Dict[str, Any]]:
    sl = result.source_label
    tl = result.target_label
    generator = ReportGenerator(result)
    return [
        {
            '模式名': diff.schema,
            '表名': diff.table,
            '外键名': diff.fk_name,
            '差异类型': diff.diff_type,
            f'{sl}定义': diff.test_value,
            f'{tl}定义': diff.prod_value,
            '同步建议': generator._get_fk_sync_advice(diff),
            '同步SQL': generator._generate_fk_sync_sql(diff) or '',
        }
        for diff in result.foreign_key_diffs
    ]


def build_view_diff_rows(result: ComparisonResult) -> List[Dict[str, Any]]:
    sl = result.source_label
    tl = result.target_label
    return [
        {
            '模式名': diff.schema,
            '视图名': diff.view,
            '差异类型': diff.diff_type,
            f'{sl}定义': diff.test_value,
            f'{tl}定义': diff.prod_value,
        }
        for diff in result.view_diffs
    ]


def build_procedure_diff_rows(result: ComparisonResult) -> List[Dict[str, Any]]:
    sl = result.source_label
    tl = result.target_label
    return [
        {
            '模式名': diff.schema,
            '对象类型': diff.object_type,
            '对象名': diff.name,
            '差异类型': diff.diff_type,
            f'{sl}定义': diff.test_value,
            f'{tl}定义': diff.prod_value,
        }
        for diff in result.procedure_diffs
    ]


def generate_deployment_script(result: ComparisonResult) -> str:
    """生成按依赖关系排序的完整部署脚本。

    排序规则（phase 越小越先执行）：
      1. CREATE TABLE  — 被引用的表必须先存在
      2. ADD COLUMN    — 字段必须先存在才能建约束/索引
      3. MODIFY COLUMN — 类型/长度/精度/可空性/默认值变更
      4. PRIMARY KEY   — 主键依赖列就绪
      5. INDEX         — 索引依赖列就绪
      6. SEQUENCE      — 独立对象，靠后
      7. FOREIGN KEY   — 最后，引用表和列都已就绪
      8. COMMENT       — 不影响结构，放最后
    """
    generator = ReportGenerator(result)
    # (phase, schema, table, sort_key, sql)
    items: List[Tuple[int, str, str, str, str]] = []

    # Phase 1: CREATE TABLE
    for diff in result.table_diffs:
        if diff.diff_type == '目标端缺失':
            sql = generator._generate_table_sync_sql(diff)
            if sql:
                items.append((1, diff.schema, diff.table, diff.table, sql))

    # Phase 2-3: 字段差异（ADD 先于 MODIFY）
    for diff in result.column_diffs:
        sql = generator._generate_sync_sql(
            diff.schema, diff.table, diff.column,
            diff.diff_type, diff.test_value, diff.prod_value,
            diff.test_column,
        )
        if not sql:
            continue
        if diff.diff_type == '目标端缺失':
            items.append((2, diff.schema, diff.table, diff.column, sql))
        elif diff.diff_type == '注释变更':
            items.append((8, diff.schema, diff.table, diff.column, sql))
        else:
            # 类型/长度/精度/可空性/默认值变更
            items.append((3, diff.schema, diff.table, diff.column, sql))

    # Phase 4: PRIMARY KEY
    for diff in result.primary_key_diffs:
        sql = generator._generate_pk_sync_sql(diff)
        if sql:
            items.append((4, diff.schema, diff.table, diff.table, sql))

    # Phase 5: INDEX
    for diff in result.index_diffs:
        sql = generator._generate_index_sync_sql(diff)
        if sql:
            items.append((5, diff.schema, diff.table, diff.index, sql))

    # Phase 6: SEQUENCE
    for diff in result.sequence_diffs:
        sql = generator._generate_sequence_sync_sql(diff)
        if sql:
            items.append((6, diff.schema, '', diff.sequence, sql))

    # Phase 7: FOREIGN KEY
    for diff in result.foreign_key_diffs:
        sql = generator._generate_fk_sync_sql(diff)
        if sql:
            items.append((7, diff.schema, diff.table, diff.fk_name, sql))

    items.sort(key=lambda x: (x[0], x[1], x[2], x[3]))

    phase_labels = {
        1: '-- 1. 建表（CREATE TABLE）',
        2: '-- 2. 加字段（ADD COLUMN）',
        3: '-- 3. 改字段（MODIFY COLUMN）',
        4: '-- 4. 主键（PRIMARY KEY）',
        5: '-- 5. 索引（INDEX）',
        6: '-- 6. 序列（SEQUENCE）',
        7: '-- 7. 外键（FOREIGN KEY）',
        8: '-- 8. 注释（COMMENT）',
    }

    lines = [
        '-- ============================================',
        '-- 达梦数据库结构同步部署脚本（按依赖排序）',
        f'-- 源端: {result.source_label}',
        f'-- 目标端: {result.target_label}',
        f'-- 生成时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}',
        '-- ============================================',
        '',
    ]

    current_phase = 0
    for phase, _schema, _table, _key, sql in items:
        if phase != current_phase:
            current_phase = phase
            lines.append('')
            lines.append(phase_labels.get(phase, f'-- {phase}'))
            lines.append('')
        lines.append(sql)
        if not sql.endswith('\n'):
            lines.append('')

    if not items:
        lines.append('-- 无差异，空脚本')

    return '\n'.join(lines)


def main():
    """主函数"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, 'config.ini')

    print("=" * 60)
    print("达梦数据库结构差异对比工具")
    print("=" * 60)
    try:
        run_compare(config_path)
    except Exception as e:
        logger.error("执行失败: %s", e, exc_info=True)
        print(f"执行失败: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
