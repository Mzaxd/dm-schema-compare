#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""达梦数据库结构差异对比 Streamlit 页面。"""

import configparser
import hashlib
import json
import logging
import os
from datetime import datetime
from typing import Dict, List

import pandas as pd
import streamlit as st

logger = logging.getLogger('dm_compare')

from dm_compare import (
    DMConnection,
    SchemaConnectionConfig,
    build_column_diff_rows,
    build_foreign_key_diff_rows,
    build_index_diff_rows,
    build_primary_key_diff_rows,
    build_procedure_diff_rows,
    build_sequence_diff_rows,
    build_summary_rows,
    build_table_diff_rows,
    build_view_diff_rows,
    generate_deployment_script,
    load_config,
    load_prod_schema_configs,
    load_schema_configs,
    _detect_config_mode,
    parse_exclude_schemas,
    parse_schema_mapping,
    run_compare,
)


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXECUTION_STATE_PATH = os.path.join(SCRIPT_DIR, 'reports', 'sql_execution_state.json')


def _scan_config_files() -> List[str]:
    """扫描脚本目录下的 config*.ini 文件。"""
    import glob as _glob
    return sorted(_glob.glob(os.path.join(SCRIPT_DIR, 'config*.ini')))


def mask_secret(value: str) -> str:
    """脱敏展示密码。"""
    if not value:
        return ''
    if len(value) <= 4:
        return '*' * len(value)
    return f"{value[:2]}{'*' * max(len(value) - 4, 4)}{value[-2:]}"


def _mask_schema_configs(items: List[SchemaConnectionConfig]) -> List[Dict[str, object]]:
    """将 SchemaConnectionConfig 列表转为脱敏字典列表。"""
    return [
        {
            'schema': item.schema,
            'host': item.host,
            'port': item.port,
            'user': item.user,
            'password': mask_secret(item.password),
        }
        for item in items
    ]


def build_config_summary(config: configparser.ConfigParser) -> Dict[str, object]:
    """构建页面展示用配置摘要。"""
    mode = _detect_config_mode(config)
    source_label = config.get('options', 'source_label', fallback='源端')
    target_label = config.get('options', 'target_label', fallback='目标端')

    if mode == 'dual':
        source_envs = _mask_schema_configs(load_schema_configs(config, 'source_env_'))
        target_envs = _mask_schema_configs(load_schema_configs(config, 'target_env_'))
        test_env = {}
    else:
        test_env = {
            'host': config.get('test_env', 'host', fallback=''),
            'port': config.get('test_env', 'port', fallback=''),
            'user': config.get('test_env', 'user', fallback=''),
            'password': mask_secret(config.get('test_env', 'password', fallback='')),
        }
        source_envs = []
        target_envs = _mask_schema_configs(load_prod_schema_configs(config))

    exclude_schemas = sorted(parse_exclude_schemas(config.get('options', 'exclude_schemas', fallback='')))
    schema_mapping = parse_schema_mapping(config.get('options', 'schema_mapping', fallback=''))

    return {
        'mode': mode,
        'source_label': source_label,
        'target_label': target_label,
        'test_env': test_env,
        'source_envs': source_envs,
        'prod_envs': target_envs,
        'exclude_schemas': exclude_schemas,
        'schema_mapping': schema_mapping,
        'output_dir': config.get('options', 'output_dir', fallback='./reports'),
    }


def to_dataframe(rows: List[Dict[str, object]]) -> pd.DataFrame:
    """转换为 DataFrame，空数据也保留表头。"""
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    return df.where(pd.notna(df), '').astype(str)


def build_row_label(row: pd.Series) -> str:
    """构建差异行显示标签。"""
    schema = row.get('模式名', '')
    table = row.get('表名', '')
    name = row.get('字段名') or row.get('索引名') or ''
    diff_type = row.get('差异类型') or ''
    return ' / '.join(str(part) for part in (schema, table, name, diff_type) if str(part))


def split_sql_statements(sql: str) -> List[str]:
    """按分号拆分SQL，忽略字符串、双引号标识符、注释内的分号。"""
    statements = []
    current = []
    in_single_quote = False
    in_double_quote = False
    in_line_comment = False
    in_block_comment = False
    index = 0

    while index < len(sql):
        char = sql[index]
        next_char = sql[index + 1] if index + 1 < len(sql) else ''

        # 块注释结束
        if in_block_comment:
            if char == '*' and next_char == '/':
                in_block_comment = False
                index += 2
                continue
            index += 1
            continue

        # 行注释结束（换行）
        if in_line_comment:
            if char == '\n':
                in_line_comment = False
                current.append(char)
            index += 1
            continue

        # 块注释开始 /*
        if char == '/' and next_char == '*' and not in_single_quote and not in_double_quote:
            in_block_comment = True
            index += 2
            continue

        # 行注释开始 --
        if char == '-' and next_char == '-' and not in_single_quote and not in_double_quote:
            in_line_comment = True
            index += 2
            continue

        if char == "'" and not in_double_quote:
            current.append(char)
            if in_single_quote and next_char == "'":
                current.append(next_char)
                index += 2
                continue
            in_single_quote = not in_single_quote
            index += 1
            continue

        if char == '"' and not in_single_quote:
            current.append(char)
            if in_double_quote and next_char == '"':
                current.append(next_char)
                index += 2
                continue
            in_double_quote = not in_double_quote
            index += 1
            continue

        if char == ';' and not in_single_quote and not in_double_quote:
            statement = ''.join(current).strip()
            if statement:
                statements.append(statement)
            current = []
            index += 1
            continue

        current.append(char)
        index += 1

    statement = ''.join(current).strip()
    if statement:
        statements.append(statement)
    return statements


def normalize_sql_for_key(sql: str) -> str:
    """规范化 SQL 文本，用于识别同一条同步 SQL。"""
    statements = split_sql_statements(sql or '')
    normalized = [' '.join(statement.split()) for statement in statements if statement.strip()]
    return ';\n'.join(normalized)


def get_sql_fingerprint(sql: str) -> str:
    """生成 SQL 指纹。"""
    normalized_sql = normalize_sql_for_key(sql)
    return hashlib.sha256(normalized_sql.encode('utf-8')).hexdigest()


def load_execution_state() -> Dict[str, Dict[str, Dict[str, str]]]:
    """读取已成功执行 SQL 的本地状态。"""
    if not os.path.exists(EXECUTION_STATE_PATH):
        return {'executed': {}}

    try:
        with open(EXECUTION_STATE_PATH, 'r', encoding='utf-8') as state_file:
            state = json.load(state_file)
        if not isinstance(state, dict):
            return {'executed': {}}
        executed = state.get('executed')
        if not isinstance(executed, dict):
            state['executed'] = {}
        return state
    except (OSError, json.JSONDecodeError):
        return {'executed': {}}


def save_execution_state(state: Dict[str, Dict[str, Dict[str, str]]]):
    """保存 SQL 执行成功状态。"""
    os.makedirs(os.path.dirname(EXECUTION_STATE_PATH), exist_ok=True)
    with open(EXECUTION_STATE_PATH, 'w', encoding='utf-8') as state_file:
        json.dump(state, state_file, ensure_ascii=False, indent=2)


def mark_sql_success(state: Dict[str, Dict[str, Dict[str, str]]],
                     row: pd.Series,
                     sql_column: str):
    """将一行同步 SQL 标记为已成功执行。"""
    sql_text = str(row.get(sql_column, '')).strip()
    if not sql_text:
        return

    fingerprint = get_sql_fingerprint(sql_text)
    state.setdefault('executed', {})[fingerprint] = {
        'executed_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'schema': str(row.get('模式名', '')).strip(),
        'object': build_row_label(row),
        'sql': sql_text,
    }


def add_execution_status(df: pd.DataFrame, sql_column: str) -> pd.DataFrame:
    """给差异表增加执行状态列。"""
    if sql_column not in df.columns or df.empty:
        return df

    state = load_execution_state()
    executed = state.get('executed', {})
    df = df.copy()
    statuses = []
    for sql_text in df[sql_column].fillna('').astype(str):
        sql_text = sql_text.strip()
        if not sql_text:
            statuses.append('-')
            continue
        statuses.append('已成功' if get_sql_fingerprint(sql_text) in executed else '未执行')

    insert_at = df.columns.get_loc(sql_column)
    df.insert(insert_at, '执行状态', statuses)
    return df


def style_execution_status(df: pd.DataFrame):
    """已执行成功的 SQL 行显示绿色底色。"""
    if '执行状态' not in df.columns:
        return df

    def highlight(row: pd.Series):
        color = 'background-color: #dcfce7' if row.get('执行状态') == '已成功' else ''
        return [color] * len(row)

    return df.style.apply(highlight, axis=1)


def get_target_config_map(config: configparser.ConfigParser):
    """按目标端模式名索引连接配置，兼容 legacy 和 dual 模式。"""
    mode = _detect_config_mode(config)
    if mode == 'dual':
        items = load_schema_configs(config, 'target_env_')
    else:
        items = load_prod_schema_configs(config)
    return {item.schema: item for item in items}


DEFAULT_DDL_TUNING = {'enabled': True, 'alter_table_opt': 3, 'ddl_wait_time': 60}

# -6407 锁超时的可操作提示：引导定位阻塞 DDL 的未提交事务
LOCK_TIMEOUT_HINT = (
    '。提示: 表上存在未提交事务阻塞 DDL，可执行 SELECT * FROM V$TRXWAIT; '
    '定位阻塞源（配合 V$LOCK / V$SESSIONS 查会话详情），待事务结束或低峰期重试；'
    '确认无用会话可 sp_close_session(sess_id) 关闭'
)


def load_ddl_tuning_config(config: configparser.ConfigParser) -> Dict[str, object]:
    """读取 [ddl_tuning] 会话级 DDL 优化配置；缺段/缺项/非法值回退内置默认。"""
    tuning = dict(DEFAULT_DDL_TUNING)
    if not config.has_section('ddl_tuning'):
        return tuning
    section = config['ddl_tuning']
    try:
        tuning['enabled'] = section.getboolean('enabled', fallback=tuning['enabled'])
    except ValueError:
        logger.warning("[DDL优化] enabled 配置非法，回退默认 true")
    for key in ('alter_table_opt', 'ddl_wait_time'):
        try:
            tuning[key] = section.getint(key, fallback=tuning[key])
        except ValueError:
            logger.warning("[DDL优化] %s 配置非法，回退默认 %s", key, tuning[key])
    return tuning


def is_lock_timeout_error(exc: BaseException) -> bool:
    """判断异常是否为达梦 -6407 锁超时（DDL 等表排它锁超时）。"""
    text = str(exc)
    return any(marker in text for marker in ('6407', '锁超时', 'Lock request timeout'))


def execute_selected_sql(selected_df: pd.DataFrame,
                         sql_column: str,
                         allow_repeat: bool = False,
                         progress_callback=None,
                         config_path: str = None) -> List[Dict[str, str]]:
    """按模式路由到目标端数据库执行选中SQL。"""
    config = load_config(config_path or os.path.join(SCRIPT_DIR, 'config.ini'))
    prod_config_map = get_target_config_map(config)
    ddl_tuning = load_ddl_tuning_config(config)
    connections: Dict[str, DMConnection] = {}
    results = []
    execution_state = load_execution_state()
    total = len(selected_df)

    try:
        for row_index, (_, row) in enumerate(selected_df.iterrows()):
            if progress_callback:
                progress_callback(row_index + 1, total, row)
            schema = str(row.get('模式名', '')).strip()
            sql_text = str(row.get(sql_column, '')).strip()
            if not schema or not sql_text:
                continue

            if not allow_repeat and get_sql_fingerprint(sql_text) in execution_state.get('executed', {}):
                results.append({
                    '模式名': schema,
                    '对象': build_row_label(row),
                    '执行结果': '跳过: 已标记成功',
                    'SQL': sql_text
                })
                continue

            if schema not in prod_config_map:
                results.append({
                    '模式名': schema,
                    '对象': build_row_label(row),
                    '执行结果': f'失败: 未找到目标端模式 {schema} 的连接配置',
                    'SQL': sql_text
                })
                continue

            if schema not in connections:
                schema_config = prod_config_map[schema]
                connection = DMConnection(
                    schema_config.host,
                    schema_config.port,
                    schema_config.user,
                    schema_config.password
                )
                if not connection.connect():
                    results.append({
                        '模式名': schema,
                        '对象': build_row_label(row),
                        '执行结果': f'失败: 目标端模式 {schema} 连接失败',
                        'SQL': sql_text
                    })
                    continue
                if ddl_tuning['enabled']:
                    applied = connection.apply_ddl_session_tuning(
                        ddl_tuning['alter_table_opt'],
                        ddl_tuning['ddl_wait_time']
                    )
                    if applied:
                        logger.info("[审计] DDL会话优化 | 模式=%s | %s（仅本会话，断开自动失效）",
                                    schema, ', '.join(applied))
                connections[schema] = connection

            statements = split_sql_statements(sql_text)
            all_succeeded = True
            for statement in statements:
                cursor = connections[schema].conn.cursor()
                try:
                    cursor.execute(statement)
                    results.append({
                        '模式名': schema,
                        '对象': build_row_label(row),
                        '执行结果': '成功',
                        'SQL': statement
                    })
                    log_sql = ' '.join(statement.split())[:200]
                    logger.info("[审计] SQL执行成功 | 模式=%s | SQL=%s", schema, log_sql)
                except Exception as exc:
                    if hasattr(connections[schema].conn, 'rollback'):
                        connections[schema].conn.rollback()
                    fail_message = f'失败: {exc}'
                    if is_lock_timeout_error(exc):
                        fail_message += LOCK_TIMEOUT_HINT
                    results.append({
                        '模式名': schema,
                        '对象': build_row_label(row),
                        '执行结果': fail_message,
                        'SQL': statement
                    })
                    log_sql = ' '.join(statement.split())[:200]
                    logger.error("[审计] SQL执行失败 | 模式=%s | 错误=%s | SQL=%s", schema, exc, log_sql)
                    all_succeeded = False
                    break
                finally:
                    cursor.close()

            if all_succeeded:
                if hasattr(connections[schema].conn, 'commit'):
                    connections[schema].conn.commit()
                mark_sql_success(execution_state, row, sql_column)

        save_execution_state(execution_state)
        return results
    finally:
        for connection in connections.values():
            connection.close()


def render_dataframe(title: str, rows: List[Dict[str, object]], sql_column: str = '同步SQL',
                     config_path: str = None):
    """渲染差异表格和完整SQL查看区。"""
    df = add_execution_status(to_dataframe(rows), sql_column)

    if df.empty:
        st.info('没有差异')
        return

    table_col, sql_col = st.columns([3, 2], gap='large')

    with table_col:
        st.subheader(title)
        if '执行状态' in df.columns:
            hide_executed = st.checkbox(
                '隐藏已成功 SQL',
                key=f'hide_executed_{title}',
                help='仅隐藏本工具已记录为执行成功的同步 SQL。'
            )
            if hide_executed:
                df = df[df['执行状态'] != '已成功']

        filter_columns = [column for column in ('差异类型', '所在环境', '模式名', '同步建议', '执行状态') if column in df.columns]
        if filter_columns:
            filter_cols = st.columns(len(filter_columns))
            for index, column in enumerate(filter_columns):
                values = sorted(str(value) for value in df[column].dropna().unique())
                selected = filter_cols[index].multiselect(
                    column,
                    values,
                    key=f'filter_{title}_{column}'
                )
                if selected:
                    df = df[df[column].astype(str).isin(selected)]

        if df.empty:
            st.info('当前筛选条件下没有差异')
            return

        st.caption('点击表格最左侧选择框可多选；右侧会自动拼接所选行的同步 SQL。')
        event = st.dataframe(
            style_execution_status(df),
            width='stretch',
            height=620,
            hide_index=True,
            key=f'diff_table_{title}',
            on_select='rerun',
            selection_mode='multi-row'
        )

    with sql_col:
        st.subheader('选中 SQL')
        if sql_column not in df.columns or df.empty:
            st.info('当前表格没有同步SQL列')
            return

        selected_positions = event.selection.rows if hasattr(event, 'selection') else []
        selected_df = df.iloc[selected_positions] if selected_positions else df.iloc[0:0]
        sql_list = [
            str(sql).strip()
            for sql in selected_df[sql_column].fillna('').tolist()
            if str(sql).strip()
        ]

        if not selected_positions:
            st.info('请在左侧表格勾选一行或多行')
            sql_df = df[df[sql_column].fillna('').astype(str).str.strip() != '']
            if not sql_df.empty:
                preview_index = st.selectbox(
                    '未选中时可单条预览',
                    options=list(sql_df.index),
                    format_func=lambda idx: build_row_label(sql_df.loc[idx]),
                    key=f'sql_preview_{title}'
                )
                st.code(str(sql_df.loc[preview_index, sql_column]), language='sql')
            return

        if not sql_list:
            st.warning('选中的行没有可展示的同步SQL')
            return

        combined_sql = '\n\n'.join(sql_list)
        st.metric('已选行数', len(selected_positions))
        selected_success_count = int((selected_df.get('执行状态', pd.Series(dtype=str)) == '已成功').sum())
        if selected_success_count:
            st.warning(f'已选中 {selected_success_count} 条已成功 SQL，默认执行时会跳过。')
        st.download_button(
            '下载选中SQL',
            combined_sql.encode('utf-8'),
            file_name=f'{title}_selected.sql',
            mime='text/plain',
            key=f'download_sql_{title}'
        )
        st.code(combined_sql, language='sql')

        st.divider()
        st.warning('执行会直接连接目标端数据库。请先确认 SQL 内容、目标模式和对象。')
        allow_repeat = st.checkbox(
            '允许重复执行已成功 SQL',
            key=f'allow_repeat_execute_{title}',
            help='默认不会重复执行已经标记成功的 SQL。'
        )
        pending_key = f'pending_execute_{title}'
        nonce_key = f'execute_nonce_{title}'
        result_key = f'execute_result_{title}'
        if nonce_key not in st.session_state:
            st.session_state[nonce_key] = 0

        if st.button('准备执行选中SQL', key=f'prepare_execute_{title}', type='primary'):
            st.session_state[nonce_key] += 1
            st.session_state[pending_key] = True

        if st.session_state.get(pending_key):
            confirm_key = f'confirm_execute_{title}_{st.session_state[nonce_key]}'
            st.checkbox(
                '我确认要在目标端执行当前选中的 SQL',
                key=confirm_key
            )
            if st.button(
                '确认执行',
                key=f'execute_sql_{title}',
                disabled=not st.session_state.get(confirm_key, False)
            ):
                progress_bar = st.progress(0, text='准备执行...')
                try:

                    def on_execute_progress(current, total, row):
                        label = build_row_label(row) if row is not None else ''
                        progress_bar.progress(
                            current / total,
                            text=f'正在执行 {current}/{total} — {label}'
                        )

                    st.session_state[result_key] = execute_selected_sql(
                        selected_df, sql_column, allow_repeat,
                        progress_callback=on_execute_progress,
                        config_path=config_path
                    )
                    progress_bar.progress(1.0, text='执行完成')
                except Exception as exc:
                    st.session_state[result_key] = [{
                        '模式名': '',
                        '对象': '',
                        '执行结果': f'执行异常: {exc}',
                        'SQL': ''
                    }]
                    logger.error("[审计] SQL执行异常: %s", exc, exc_info=True)
                st.session_state[pending_key] = False
                st.session_state[nonce_key] += 1
                st.rerun()

        if st.session_state.get(result_key):
            st.caption('最近一次执行结果')
            st.dataframe(
                to_dataframe(st.session_state[result_key]),
                width='stretch',
                height=260,
                hide_index=True
            )


def render_summary_table(rows: List[Dict[str, object]]):
    """渲染汇总表。"""
    st.subheader('汇总概览')
    df = to_dataframe(rows)
    if df.empty:
        st.info('暂无汇总数据')
        return
    st.dataframe(df, width='stretch', hide_index=True)


def render_config_summary(config_summary: Dict[str, object]):
    """渲染脱敏配置摘要。"""
    source_label = config_summary.get('source_label', '源端')
    target_label = config_summary.get('target_label', '目标端')
    mode = config_summary.get('mode', 'legacy')

    with st.sidebar.expander('配置摘要', expanded=True):
        # 源端
        st.caption(source_label)
        if mode == 'dual' and config_summary.get('source_envs'):
            st.dataframe(pd.DataFrame(config_summary['source_envs']), width='stretch', hide_index=True)
        else:
            st.dataframe(pd.DataFrame([config_summary['test_env']]), width='stretch', hide_index=True)

        # 目标端
        st.caption(target_label)
        st.dataframe(pd.DataFrame(config_summary['prod_envs']), width='stretch', hide_index=True)

        st.caption('模式映射')
        mapping = config_summary['schema_mapping']
        mapping_rows = [{f'{source_label}模式': key, f'{target_label}模式': value} for key, value in mapping.items()]
        st.dataframe(pd.DataFrame(mapping_rows), width='stretch', hide_index=True)

        st.caption('排除模式')
        st.write(', '.join(config_summary['exclude_schemas']) or '-')


def render_result(config_path: str):
    """渲染比对结果。"""
    run_result = st.session_state.get('run_result')
    if not run_result:
        st.info('点击左侧「开始比对」后，结果会显示在这里。')
        return

    result = run_result.result
    with st.sidebar.expander('本次结果', expanded=True):
        st.success(f"完成：{datetime.now().strftime('%H:%M:%S')}")
        st.metric('表差异', len(result.table_diffs))
        st.metric('字段差异', len(result.column_diffs))
        st.metric('索引差异', len(result.index_diffs))
        st.metric('主键差异', len(result.primary_key_diffs))
        st.metric('序列差异', len(result.sequence_diffs))
        st.metric('外键差异', len(result.foreign_key_diffs))
        st.metric('视图差异', len(result.view_diffs))
        st.metric('过程/函数/触发器差异', len(result.procedure_diffs))

        if os.path.exists(run_result.report_path):
            with open(run_result.report_path, 'rb') as report_file:
                st.download_button(
                    '下载Excel报告',
                    report_file,
                    file_name=os.path.basename(run_result.report_path),
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

        st.divider()
        deploy_script = generate_deployment_script(result)
        diff_total = (len(result.table_diffs) + len(result.column_diffs)
                      + len(result.index_diffs) + len(result.primary_key_diffs)
                      + len(result.sequence_diffs) + len(result.foreign_key_diffs))
        st.download_button(
            '⬇ 下载完整部署脚本（已排序）',
            deploy_script.encode('utf-8'),
            file_name=f'deploy_{datetime.now().strftime("%Y%m%d_%H%M%S")}.sql',
            mime='text/plain',
            help=f'按依赖关系排序：建表→加字段→改字段→主键→索引→序列→外键→注释。共 {diff_total} 条差异。',
            disabled=(diff_total == 0),
        )

    tabs = st.tabs(['汇总概览', '表差异', '字段差异', '索引差异', '主键差异', '序列差异',
                     '外键差异', '视图差异', '过程函数触发器差异'])
    with tabs[0]:
        render_summary_table(build_summary_rows(result))
    with tabs[1]:
        render_dataframe('表差异', build_table_diff_rows(result), config_path=config_path)
    with tabs[2]:
        render_dataframe('字段差异', build_column_diff_rows(result), config_path=config_path)
    with tabs[3]:
        render_dataframe('索引差异', build_index_diff_rows(result), config_path=config_path)
    with tabs[4]:
        render_dataframe('主键差异', build_primary_key_diff_rows(result), config_path=config_path)
    with tabs[5]:
        render_dataframe('序列差异', build_sequence_diff_rows(result), config_path=config_path)
    with tabs[6]:
        render_dataframe('外键差异', build_foreign_key_diff_rows(result), config_path=config_path)
    with tabs[7]:
        render_dataframe('视图差异', build_view_diff_rows(result), sql_column='', config_path=config_path)
    with tabs[8]:
        render_dataframe('过程函数触发器差异', build_procedure_diff_rows(result), sql_column='', config_path=config_path)


def main():
    st.set_page_config(page_title='达梦数据库结构差异对比', layout='wide')
    st.title('达梦数据库结构差异对比')
    st.caption('本页面展示差异和同步 SQL；执行 SQL 需要手动勾选确认。')

    config_files = _scan_config_files()
    if not config_files:
        st.error('未找到任何 config*.ini 配置文件')
        return

    config_path = st.sidebar.selectbox(
        '选择配置文件',
        config_files,
        format_func=lambda p: os.path.basename(p),
        index=0,
    )

    config = load_config(config_path)
    render_config_summary(build_config_summary(config))

    st.sidebar.divider()
    st.sidebar.subheader('运行控制')
    progress_box = st.sidebar.empty()
    if st.sidebar.button('开始比对', type='primary', use_container_width=True):
        if os.path.exists(EXECUTION_STATE_PATH):
            os.remove(EXECUTION_STATE_PATH)
        st.session_state.pop('run_result', None)
        progress_messages: List[str] = []

        def on_progress(message: str):
            progress_messages.append(message)
            progress_box.text('\n'.join(progress_messages[-12:]))

        with st.spinner('正在执行数据库结构比对...'):
            try:
                st.session_state['run_result'] = run_compare(config_path, progress_callback=on_progress)
            except Exception as exc:
                st.session_state.pop('run_result', None)
                st.error(f'比对失败: {exc}')
                return

    render_result(config_path)


if __name__ == '__main__':
    main()
