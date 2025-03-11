import pandas as pd
import pymysql

# 数据库连接参数
host = ''
user = ''
password = ''
database = ''

# 创建数据库连接
connection = pymysql.connect(host=host, user=user, password=password, database=database)

# 创建游标
cursor = connection.cursor()

# SQL 查询语句
sql = '''
SELECT 
    COLUMN_NAME AS '字段名',
    DATA_TYPE AS '类型',
    IS_NULLABLE AS '是否为空',
    COLUMN_DEFAULT AS '默认值',
    COLUMN_COMMENT AS '注释'
FROM 
    INFORMATION_SCHEMA.COLUMNS
WHERE 
    TABLE_SCHEMA = %s
    AND TABLE_NAME = %s
'''

# 定义要查询的表名列表
tables = ['blade_application',
'blade_attach',
'blade_client',
'blade_code',
'blade_datasource',
'blade_dept',
'blade_dept_bak1',
'blade_dept_new',
'blade_dept_old',
'blade_dept_region_backup',
'blade_dict',
'blade_dict_biz',
'blade_error_turn_message',
'blade_log_api',
'blade_log_error',
'blade_log_usual',
'blade_menu',
'blade_notice',
'blade_notice_record',
'blade_oss',
'blade_param',
'blade_post',
'blade_process_leave',
'blade_region',
'blade_region_temp',
'blade_report_file',
'blade_role',
'blade_role_application',
'blade_role_bak',
'blade_role_copy1',
'blade_role_dept',
'blade_role_menu',
'blade_role_scope',
'blade_scope_api',
'blade_scope_data',
'blade_sms',
'blade_tenant',
'blade_top_menu',
'blade_top_menu_setting',
'blade_user',
'blade_user_app',
'blade_user_application',
'blade_user_dept',
'blade_user_oauth',
'blade_user_other',
'blade_user_web',
'branch_table',
'distributed_lock',
'global_table',
'lock_table',
'tb_anjgl_cwcl',
'tb_anjgl_cwcl_anj',
'tb_anjgl_info',
'tb_anjgl_ry',
'tb_anjgl_wp',
'tb_anjgl_xzdw',
'tb_cass_boat',
'tb_cass_car',
'tb_cass_info',
'tb_cass_items',
'tb_cass_person',
'tb_cass_sazz',
'tb_grid',
'tb_grid_user',
'tb_gzdx_key_places',
'tb_gzdx_key_places_record',
'tb_gzdx_key_units',
'tb_gzdx_key_units_person',
'tb_gzdx_key_units_record',
'tb_gzdx_ryfcjl',
'tb_gzdx_rykhjl',
'tb_gzdx_rylgjl',
'tb_gzdx_ryxx',
'tb_kp_config',
'tb_kp_config_dept',
'tb_kp_config_dept_person',
'tb_kp_config_item',
'tb_kp_file',
'tb_kp_plan',
'tb_kp_plan_dept',
'tb_kp_plan_dept_item',
'tb_kp_plan_dept_person',
'tb_qrcode_address',
'tb_sazz',
'tb_search_catalog',
'tb_search_catalog_resource',
'tb_search_resource',
'tb_sgdanj_xsaj',
'tb_sgdanj_xsaj_ryxx',
'tb_sgdanj_xsaj_wpxx',
'tb_sgdanj_xzaj',
'tb_sgdanj_xzaj_ryxx',
'tb_sgdanj_xzaj_wpxx',
'tb_shfxd_bgjl',
'tb_shfxd_bgjl_log',
'tb_shfxd_info',
'tb_shfxd_info_nw',
'tb_shfxd_info_ww',
'tb_spe_work',
'tb_spe_work_approve',
'tb_spe_work_feedback',
'tb_sync_resource',
'tb_test_financial',
'tb_warn',
'tb_warn_approve',
'tb_warn_person',
'tb_warn_turn',
'tb_warn_turn_feedback',
'undo_log',
'w_application',
'w_message',
'w_task']

# 初始化一个空的DataFrame列表，用于存储每个表的数据块
dfs = []

for i, table in enumerate(tables):
    # 执行查询
    cursor.execute(sql, (database, table))
    result = cursor.fetchall()

    # 将结果转换为列表
    result_list = [list(row) for row in result]

    # 创建当前表的DataFrame
    df = pd.DataFrame(result_list, columns=['字段名', '类型', '是否为空', '默认值', '注释'])

    df_11 = pd.DataFrame([['字段名', '类型', '是否为空', '默认值', '注释']], columns=df.columns)


    # 添加表名行
    table_row = pd.DataFrame([[f'表名: {table}', '', '', '', '']], columns=df.columns)
    dfs.append(table_row)
    dfs.append(df_11)
    dfs.append(df)

    # 如果不是最后一个表，添加空行
    if i != len(tables) - 1:
        empty_row = pd.DataFrame([['', '', '', '', '']], columns=df.columns)
        dfs.append(empty_row)

# 合并所有数据块到一个总的DataFrame
df_all = pd.concat(dfs, ignore_index=True)

# 保存到Excel文件
df_all.to_excel('数据库字段信息.xlsx', index=False)

# 关闭游标和连接
cursor.close()
connection.close()

print("数据已成功保存到 Excel 文件。")