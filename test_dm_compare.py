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
    DMConnection,
)
from streamlit_app import split_sql_statements, load_ddl_tuning_config, is_lock_timeout_error


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


def test_missing_table_sql_includes_fk():
    """目标端缺失的表：同步SQL应一次带全（序列+建表+注释+索引+外键）。"""
    from dm_compare import (
        ComparisonResult, TableDiff, TableInfo, ColumnInfo,
        ForeignKeyInfo, ForeignKeyColumn, ReportGenerator,
        SequenceDiff, SequenceInfo,
    )
    new_table = TableInfo(name='NEW_T', schema='S', table_type='TABLE')
    new_table.columns = {
        'id': ColumnInfo('id', 'INT', None, None, None, 'N', None),
    }
    fk = ForeignKeyInfo(name='fk_new', schema='S', table_name='NEW_T',
                        ref_schema='S', ref_table_name='REF_T', delete_rule='CASCADE')
    fk.columns = [ForeignKeyColumn(column_name='ref_id', position=1, ref_column_name='id')]
    new_table.foreign_keys = {'fk_new': fk}

    result = ComparisonResult()
    result.sequence_diffs = [
        SequenceDiff(schema='S', sequence='NEW_T_SEQ', diff_type='目标端缺失',
                     test_value='START_WITH=1', prod_value='-',
                     test_sequence=SequenceInfo(name='NEW_T_SEQ', schema='S', min_value=1,
                                                max_value=99999999, increment_by=1,
                                                cycle_flag='N', order_flag='N', cache_size=20)),
    ]
    diff = TableDiff(schema='S', table='NEW_T', diff_type='目标端缺失', test_table=new_table)
    sql = ReportGenerator(result)._generate_table_sync_sql(diff)
    # 序列在前，建表在后
    assert sql.index('CREATE SEQUENCE') < sql.index('CREATE TABLE')
    assert 'FOREIGN KEY ("ref_id")' in sql
    assert 'REFERENCES S.REF_T ("id")' in sql
    assert 'ON DELETE CASCADE' in sql


def _make_length_diff(test_len, prod_len, schema='S',
                      test_type='VARCHAR', prod_type='VARCHAR'):
    """构造一条长度变更的 ColumnDiff 及其 ComparisonResult。"""
    from dm_compare import ComparisonResult, ColumnDiff, ColumnInfo
    result = ComparisonResult()
    diff = ColumnDiff(
        schema=schema, table='T', column='C', diff_type='长度变更',
        test_value=str(test_len), prod_value=str(prod_len),
        test_column=ColumnInfo('C', test_type, test_len, None, None, 'Y', None),
        prod_column=ColumnInfo('C', prod_type, prod_len, None, None, 'Y', None),
    )
    return result, diff


def test_reverse_length_diff_targets_source():
    """目标端更长（生产临时扩长）：SQL 改为在源端执行、按目标端长度扩源端。"""
    from dm_compare import ReportGenerator
    result, diff = _make_length_diff(100, 200)
    generator = ReportGenerator(result)
    assert generator._is_reverse_length_diff(diff)
    assert generator._get_column_execution_side(diff) == '源端'
    assert generator._generate_column_sync_sql(diff) == 'ALTER TABLE S.T MODIFY C VARCHAR(200);'
    assert '反向同步' in generator._get_column_sync_advice(diff)


def test_forward_length_diff_still_targets_prod():
    """源端更长：维持正向，在目标端执行、把目标端扩到源端长度。"""
    from dm_compare import ReportGenerator
    result, diff = _make_length_diff(200, 100)
    generator = ReportGenerator(result)
    assert not generator._is_reverse_length_diff(diff)
    assert generator._get_column_execution_side(diff) == '目标端'
    assert generator._generate_column_sync_sql(diff) == 'ALTER TABLE S.T MODIFY C VARCHAR(200);'
    assert generator._get_column_sync_advice(diff) == ''


def test_reverse_length_diff_type_mismatch_keeps_forward():
    """类型也不一致时不视为临时扩长，维持正向（源端为准）。"""
    from dm_compare import ReportGenerator
    result, diff = _make_length_diff(100, 200, test_type='VARCHAR', prod_type='CHAR')
    generator = ReportGenerator(result)
    assert not generator._is_reverse_length_diff(diff)
    assert generator._get_column_execution_side(diff) == '目标端'


def test_reverse_length_diff_restores_source_schema_name():
    """模式映射 SRC_APP→APP：反向 SQL 的模式名需还原为源端实际名 SRC_APP。"""
    from dm_compare import ReportGenerator
    result, diff = _make_length_diff(100, 200, schema='APP')
    result.schema_mappings = {'SRC_APP': 'APP'}
    generator = ReportGenerator(result)
    sql = generator._generate_column_reverse_sql(diff)
    assert 'SRC_APP.T' in sql
    assert 'ALTER TABLE SRC_APP.T MODIFY C VARCHAR(200);' == sql


def test_build_column_diff_rows_marks_execution_side():
    from dm_compare import build_column_diff_rows
    result, diff = _make_length_diff(100, 200)
    result.column_diffs = [diff]
    rows = build_column_diff_rows(result)
    assert rows[0]['执行端'] == '源端'
    assert rows[0]['同步SQL'] == 'ALTER TABLE S.T MODIFY C VARCHAR(200);'

    result2, diff2 = _make_length_diff(200, 100)
    result2.column_diffs = [diff2]
    rows2 = build_column_diff_rows(result2)
    assert rows2[0]['执行端'] == '目标端'


def test_deployment_script_excludes_reverse_diff():
    """目标端部署脚本必须剔除反向项；反向脚本包含它。"""
    from dm_compare import generate_deployment_script, generate_reverse_sync_script
    result, diff = _make_length_diff(100, 200)
    result.column_diffs = [diff]

    script = generate_deployment_script(result)
    assert 'MODIFY' not in script
    assert '源端' in script  # 头部有提示

    reverse = generate_reverse_sync_script(result)
    assert 'ALTER TABLE S.T MODIFY C VARCHAR(200);' in reverse
    assert '在源端执行' in reverse


def test_get_inverse_schema_mapping():
    """目标端模式名反查源端模式名，用于反向 SQL 的连接路由。"""
    import configparser
    from streamlit_app import get_inverse_schema_mapping
    config = configparser.ConfigParser(interpolation=None)
    config.read_string('[options]\nschema_mapping = SRC_APP:APP, SAME:SAME\n')
    assert get_inverse_schema_mapping(config) == {'APP': 'SRC_APP', 'SAME': 'SAME'}
    empty_config = configparser.ConfigParser(interpolation=None)
    empty_config.read_string('[options]\n')
    assert get_inverse_schema_mapping(empty_config) == {}


class _FakeCursor:
    """记录 execute 的 SQL，可配置为抛异常，模拟 dmPython 游标。"""

    def __init__(self, conn):
        self.conn = conn

    def execute(self, sql):
        if self.conn.fail:
            raise RuntimeError('unsupported session parameter')
        self.conn.executed.append(sql)

    def close(self):
        pass


class _FakeDmConn:
    def __init__(self, fail=False):
        self.executed = []
        self.fail = fail

    def cursor(self):
        return _FakeCursor(self)


def _make_dm_connection(fail=False):
    conn = DMConnection('mockhost', 5236, 'user', 'pwd')
    conn.conn = _FakeDmConn(fail=fail)
    return conn


def test_apply_ddl_session_tuning_success():
    """会话调优成功：两条 ALTER SESSION SET 按序执行并返回参数描述。"""
    conn = _make_dm_connection()
    applied = conn.apply_ddl_session_tuning()
    assert applied == ['ALTER_TABLE_OPT=3', 'DDL_WAIT_TIME=60']
    assert conn.conn.executed == [
        "ALTER SESSION SET 'ALTER_TABLE_OPT' = 3;",
        "ALTER SESSION SET 'DDL_WAIT_TIME' = 60;",
    ]


def test_apply_ddl_session_tuning_clamps_wait_time():
    """DDL_WAIT_TIME 超过达梦上限 60 秒时收敛到 60。"""
    conn = _make_dm_connection()
    applied = conn.apply_ddl_session_tuning(ddl_wait_time=120)
    assert applied == ['ALTER_TABLE_OPT=3', 'DDL_WAIT_TIME=60']
    assert "ALTER SESSION SET 'DDL_WAIT_TIME' = 60;" in conn.conn.executed


def test_apply_ddl_session_tuning_failure_degrades():
    """设置失败时降级：不抛异常、返回空列表，同步流程可继续。"""
    conn = _make_dm_connection(fail=True)
    applied = conn.apply_ddl_session_tuning()
    assert applied == []


def test_load_ddl_tuning_config():
    """[ddl_tuning] 缺段回退默认；有段取覆盖值；非法值回退默认。"""
    import configparser

    # 缺段 → 内置默认
    config = configparser.ConfigParser(interpolation=None)
    config.read_string('[options]\noutput_dir = ./reports\n')
    assert load_ddl_tuning_config(config) == {'enabled': True, 'alter_table_opt': 3, 'ddl_wait_time': 60}

    # 显式关闭
    config.read_string('[ddl_tuning]\nenabled = false\nalter_table_opt = 2\nddl_wait_time = 30\n')
    assert load_ddl_tuning_config(config) == {'enabled': False, 'alter_table_opt': 2, 'ddl_wait_time': 30}

    # 非法值 → 逐项回退默认
    config.read_string('[ddl_tuning]\nenabled = maybe\nalter_table_opt = fast\nddl_wait_time = \n')
    assert load_ddl_tuning_config(config) == {'enabled': True, 'alter_table_opt': 3, 'ddl_wait_time': 60}


def test_is_lock_timeout_error():
    """识别达梦 -6407 锁超时的各种报错文本形态。"""
    assert is_lock_timeout_error(Exception('[-6407]: Lock request timeout'))
    assert is_lock_timeout_error(Exception('code: -6407 锁超时'))
    assert is_lock_timeout_error(RuntimeError('锁请求超时: 锁超时'))
    assert is_lock_timeout_error(Exception('Lock request timeout occurred'))
    assert not is_lock_timeout_error(Exception('[-6406]: other error'))
    assert not is_lock_timeout_error(Exception('network unreachable'))


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
        test_missing_table_sql_includes_fk,
        test_reverse_length_diff_targets_source,
        test_forward_length_diff_still_targets_prod,
        test_reverse_length_diff_type_mismatch_keeps_forward,
        test_reverse_length_diff_restores_source_schema_name,
        test_build_column_diff_rows_marks_execution_side,
        test_deployment_script_excludes_reverse_diff,
        test_get_inverse_schema_mapping,
        test_apply_ddl_session_tuning_success,
        test_apply_ddl_session_tuning_clamps_wait_time,
        test_apply_ddl_session_tuning_failure_degrades,
        test_load_ddl_tuning_config,
        test_is_lock_timeout_error,
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
