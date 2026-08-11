#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""dm_compare 纯函数自检测试。

不连数据库，只测类型格式化、默认值归一化、标识符引用、SQL 拆分等核心逻辑。
运行：python test_dm_compare.py
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from dm_compare import (
    _normalize_dm_data_type,
    _format_dm_column_type,
    ColumnInfo,
    IndexInfo,
    IndexColumnInfo,
    ForeignKeyInfo,
    ForeignKeyColumn,
)
from streamlit_app import split_sql_statements


def test_normalize_dm_data_type():
    assert _normalize_dm_data_type('VARCHAR2') == 'VARCHAR'
    assert _normalize_dm_data_type('varchar2') == 'VARCHAR'
    assert _normalize_dm_data_type('DATETIME') == 'TIMESTAMP'
    assert _normalize_dm_data_type('INT') == 'INTEGER'
    assert _normalize_dm_data_type('NUMBER') == 'NUMBER'
    assert _normalize_dm_data_type(None) == ''
    assert _normalize_dm_data_type('') == ''


def test_format_dm_column_type():
    # 无参数类型
    assert _format_dm_column_type('DATE', None, None, None) == 'DATE'
    assert _format_dm_column_type('BLOB', None, None, None) == 'BLOB'
    assert _format_dm_column_type('INT', None, None, None) == 'INT'

    # 精度类型
    assert _format_dm_column_type('NUMBER', None, 10, 2) == 'NUMBER(10,2)'
    assert _format_dm_column_type('NUMBER', None, 10, None) == 'NUMBER(10)'
    assert _format_dm_column_type('DECIMAL', None, 5, 0) == 'DECIMAL(5,0)'

    # 长度类型
    assert _format_dm_column_type('VARCHAR', 255, None, None) == 'VARCHAR(255)'
    assert _format_dm_column_type('CHAR', 10, None, None) == 'CHAR(10)'

    # 别名归一化
    assert _format_dm_column_type('VARCHAR2', 100, None, None, normalize_aliases=True) == 'VARCHAR(100)'
    assert _format_dm_column_type('VARCHAR2', 100, None, None, normalize_aliases=False) == 'VARCHAR2(100)'

    # 不给长度不该追加 ()
    assert _format_dm_column_type('VARCHAR', None, None, None) == 'VARCHAR'


def test_column_info_type_string():
    col = ColumnInfo(name='id', data_type='NUMBER', length=None, precision=10, scale=0, nullable='N', default=None)
    assert col.get_type_string() == 'NUMBER(10,0)'

    col2 = ColumnInfo(name='name', data_type='VARCHAR', length=200, precision=None, scale=None, nullable='Y', default=None)
    assert col2.get_type_string() == 'VARCHAR(200)'


def test_index_info_columns_string():
    idx = IndexInfo(name='idx_test', schema='S', table_name='T', uniqueness='UNIQUE')
    idx.columns = [
        IndexColumnInfo(name='col_b', position=2),
        IndexColumnInfo(name='col_a', position=1, sort_order='DESC'),
    ]
    s = idx.get_columns_string()
    assert 'col_a DESC' in s
    assert 'col_b' in s
    assert s.index('col_a') < s.index('col_b')


def test_fk_definition_string():
    fk = ForeignKeyInfo(name='fk_1', schema='S', table_name='T', ref_schema='S', ref_table_name='REF_T',
                        delete_rule='CASCADE')
    fk.columns = [
        ForeignKeyColumn(column_name='ref_id', position=1, ref_column_name='id'),
    ]
    s = fk.get_definition_string()
    assert 'FOREIGN KEY (ref_id)' in s
    assert 'REFERENCES S.REF_T (id)' in s
    assert 'ON DELETE CASCADE' in s


def test_fk_no_delete_rule():
    fk = ForeignKeyInfo(name='fk_2', schema='S', table_name='T', ref_schema='S', ref_table_name='REF_T',
                        delete_rule='NO ACTION')
    fk.columns = [ForeignKeyColumn(column_name='x', position=1, ref_column_name='y')]
    s = fk.get_definition_string()
    assert 'ON DELETE' not in s


def test_split_sql_statements():
    # 基本分号拆分
    stmts = split_sql_statements("SELECT 1; SELECT 2;")
    assert len(stmts) == 2
    assert stmts[0] == 'SELECT 1'
    assert stmts[1] == 'SELECT 2'

    # 单引号内的分号不拆
    stmts = split_sql_statements("INSERT INTO t VALUES('a;b'); SELECT 1;")
    assert len(stmts) == 2

    # 双引号标识符内的分号不拆
    stmts = split_sql_statements('ALTER TABLE "my;table" ADD c INT; SELECT 1;')
    assert len(stmts) == 2

    # 行注释
    stmts = split_sql_statements("SELECT 1 -- this; is comment\n; SELECT 2;")
    assert len(stmts) == 2

    # 块注释
    stmts = split_sql_statements("SELECT 1 /* ; ; */; SELECT 2;")
    assert len(stmts) == 2

    # 转义单引号 ''
    stmts = split_sql_statements("INSERT INTO t VALUES('it''s; ok'); SELECT 1;")
    assert len(stmts) == 2

    # 空字符串
    assert split_sql_statements('') == []
    assert split_sql_statements('   ') == []


def test_split_sql_create_table():
    sql = """
        CREATE TABLE schema.t (
            id INT NOT NULL,
            name VARCHAR(100)
        );
        COMMENT ON TABLE schema.t IS 'test;comment';
    """
    stmts = split_sql_statements(sql)
    assert len(stmts) == 2
    assert 'CREATE TABLE' in stmts[0]
    assert 'COMMENT ON TABLE' in stmts[1]


def test_normalize_default_for_compare():
    from dm_compare import SchemaComparator
    norm = SchemaComparator._normalize_default_for_compare
    assert norm(None) is None
    assert norm('NULL') is None
    assert norm('null') is None
    assert norm('') is None
    assert norm('0') == '0'
    assert norm('abc') == 'abc'


def test_quote_identifier():
    from dm_compare import ReportGenerator
    q = ReportGenerator._quote_identifier
    # 纯大写字母数字下划线 → 不加引号
    assert q('MY_TABLE') == 'MY_TABLE'
    assert q('COL_123') == 'COL_123'
    # 含小写 → 加引号
    assert q('myTable') == '"myTable"'
    # 含特殊字符 → 加引号
    assert q('my table') == '"my table"'
    # 双引号转义
    assert q('col"quote') == '"col""quote"'
    # 空字符串
    assert q('') == ''


def test_is_name_excluded():
    from dm_compare import _is_name_excluded
    patterns = ['DATABASECHANGELOG*', 'TEMP_*']
    assert _is_name_excluded('DATABASECHANGELOG', patterns)
    assert _is_name_excluded('databasechangeloglock', patterns)
    assert _is_name_excluded('TEMP_001', patterns)
    assert not _is_name_excluded('MY_TABLE', patterns)
    assert not _is_name_excluded('MY_TABLE', None)


def test_deployment_script_ordering():
    """部署脚本必须按依赖顺序排列：建表→加字段→改字段→主键→索引→序列→外键→注释。"""
    from dm_compare import (
        ComparisonResult, ColumnDiff, TableDiff, TableInfo, ColumnInfo,
        generate_deployment_script,
    )
    result = ComparisonResult()
    # 构造一张有字段的表，让 _generate_table_sync_sql 能输出 CREATE TABLE
    new_table = TableInfo(name='NEW_T', schema='S', table_type='TABLE')
    new_table.columns = {
        'id': ColumnInfo('id', 'INT', None, None, None, 'N', None),
    }
    # 模拟混合差异：故意打乱插入顺序
    result.column_diffs = [
        ColumnDiff(schema='S', table='T', column='c2', diff_type='注释变更',
                   test_value='注释', prod_value='-', test_column=ColumnInfo('c2', 'INT', None, None, None, 'Y', None)),
        ColumnDiff(schema='S', table='T', column='c1', diff_type='目标端缺失',
                   test_value='INT', prod_value='-', test_column=ColumnInfo('c1', 'INT', None, None, None, 'Y', None)),
    ]
    result.table_diffs = [
        TableDiff(schema='S', table='NEW_T', diff_type='目标端缺失',
                  test_table=new_table),
    ]

    script = generate_deployment_script(result)
    # Phase 标题必须按序出现
    assert '-- 1. 建表' in script
    assert '-- 2. 加字段' in script
    assert '-- 8. 注释' in script
    # 建表 section 必须在加字段 section 之前
    assert script.index('-- 1. 建表') < script.index('-- 2. 加字段')
    # 加字段 section 必须在注释 section 之前
    assert script.index('-- 2. 加字段') < script.index('-- 8. 注释')
    # CREATE TABLE 语句确实生成
    assert 'CREATE TABLE' in script


def run_all():
    tests = [
        test_normalize_dm_data_type,
        test_format_dm_column_type,
        test_column_info_type_string,
        test_index_info_columns_string,
        test_fk_definition_string,
        test_fk_no_delete_rule,
        test_split_sql_statements,
        test_split_sql_create_table,
        test_normalize_default_for_compare,
        test_quote_identifier,
        test_is_name_excluded,
        test_deployment_script_ordering,
    ]
    passed = 0
    failed = 0
    for test in tests:
        try:
            test()
            print(f"  PASS  {test.__name__}")
            passed += 1
        except AssertionError as e:
            print(f"  FAIL  {test.__name__}: {e}")
            failed += 1
        except Exception as e:
            print(f"  ERROR {test.__name__}: {type(e).__name__}: {e}")
            failed += 1

    print(f"\n{'='*40}")
    print(f"结果: {passed} 通过, {failed} 失败 (共 {len(tests)} 个)")
    return 0 if failed == 0 else 1


if __name__ == '__main__':
    sys.exit(run_all())
