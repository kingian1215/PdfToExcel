import os
import time
import cx_Oracle
from prettytable import from_db_cursor
import prettytable as pt

# 定义颜色
# Color
R = "\033[0;31;40m"  # RED
G = "\033[0;32;40m"  # GREEN
Y = "\033[0;33;40m"  # Yellow
B = "\033[0;34;40m"  # Blue
N = "\033[0m"  # Reset
a = "使用率正常"
b = "使用率异常"

# 当前文件的路径信息
current_path = os.path.abspath(__file__)  # 获取当前文件路径
father_path = os.path.abspath(os.path.dirname(current_path) + os.path.sep + ".")  # 获取当前文件的父路径

# 获取时间
checktime = time.strftime("%Y-%m-%d", time.localtime())

# 通过配置文件或者oracle数据库配置信息
ora_info = open(father_path + '/ora.cfg', 'r')
db_info = ora_info.readlines()
ora_info.close()

for db_list in db_info:
    info = db_list.split()
    ora_ip = info[0].strip()
    ora_user = info[1].strip()
    ora_passwd = info[2].strip()
    ora_port = info[3].strip()
    ora_sid = info[4].strip()
    ora_use = info[5].strip()
    filename = father_path + '/ora_tspCheck' + checktime + '.log'
    error_filename = father_path + '/Error_ora_tspCheck' + checktime + '.log'
    out_file = open(filename, 'a')
    err_file = open(error_filename, 'a')
    try:
        ora_db = cx_Oracle.connect(ora_user + '/' + ora_passwd + '@' + ora_ip + ':' + ora_port + '/' + ora_sid)
    except cx_Oracle.DatabaseError as e1:
        print('IP:%s 用途为:%s 无法正常连接，请检查，报错代码为:%s' % (ora_ip, ora_use, e1), file=err_file)
    except KeyboardInterrupt as e2:
        print('其他错误，请检查 %s' % e2, file=err_file)
    else:
        print('开始检查IP %s 用途为 %s的数据库' % (ora_ip, ora_use), file=out_file)
        curs = ora_db.cursor()
        curs.execute("""
    SELECT /*+ NO_CPU_COSTING */
            F.HOST_NAME,
             d.dbid,
             F.INSTANCE_NAME,
             SYSDATE,
             D.LOG_MODE,
             a.tablespace_name,
             E1.file_count,
             TRUNC (a.total) allocated_space_mb,
             TRUNC (a.total - b.free) Used_mb,
             TRUNC (b.free) free_space_mb,
             ROUND (1 - b.free / a.total, 4) * 100 "USAGE_%",
             c.AUTOSIZE AUTOSIZE_mb,
             ROUND ( (a.total - b.free) / c.AUTOSIZE, 4) * 100 "AUTOUSAGE_%"
        FROM (  SELECT tablespace_name, SUM (NVL (bytes, 2)) / 1024 / 1024 total
                  FROM dba_data_files DBA_DATA_FILES1
              GROUP BY tablespace_name) a,
             (  SELECT tablespace_name, SUM (NVL (bytes, 2)) / 1024 / 1024 free
                  FROM dba_free_space
              GROUP BY tablespace_name) b,
             (  SELECT tablespace_name, COUNT (file_name) file_count
                  FROM dba_data_files DBA_DATA_FILES2
              GROUP BY tablespace_name) E1,
             (  SELECT x.TABLESPACE_NAME, SUM (x.AUTOSIZE) AUTOSIZE
                  FROM (SELECT TABLESPACE_NAME,
                               CASE
                                  WHEN MAXBYTES / 1024 / 1024 = 0
                                  THEN
                                     BYTES / 1024 / 1024
                                  ELSE
                                     MAXBYTES / 1024 / 1024
                               END
                                  AUTOSIZE
                          FROM DBA_DATA_FILES DBA_DATA_FILES3) x
              GROUP BY x.tablespace_name) c,
             v$database d,
             (SELECT UTL_INADDR.get_host_address () ip FROM DUAL) E2,
             v$instance f
       WHERE     b.tablespace_name = a.tablespace_name
             AND c.TABLESPACE_NAME = b.TABLESPACE_NAME
             AND E1.tablespace_name = a.tablespace_name
             AND a.tablespace_name = c.TABLESPACE_NAME
             AND E1.tablespace_name = b.tablespace_name
             AND E1.tablespace_name = c.TABLESPACE_NAME
    ORDER BY 13 DESC
                """)
        data = curs.fetchall()
        row = list(range(len(data)))
        tb = pt.PrettyTable()
        tb.field_names = ['IP','Hostname', 'DBID', 'Sid', 'Date', 'Archive_Mode', 'TableSpace', 'Db_File', 'Total(Mb)',
                          'Use(Mb)', 'Free(Mb)', 'Usage%', 'Max_Total(Mb)', 'Max_Usage%', 'Status']
        for l in row:
            for i in data:
                usage = int(data[l][-1])
                if usage >= 85:
                    tb.add_row(
                        [ora_ip,data[l][0], data[l][1], data[l][2], data[l][3], data[l][4], data[l][5], data[l][6], data[l][7],
                         data[l][8], data[l][9], data[l][10], data[l][11], data[l][12], R + b + N])
                    break
                else:
                    tb.add_row(
                        [ora_ip,data[l][0], data[l][1], data[l][2], data[l][3], data[l][4], data[l][5], data[l][6], data[l][7],
                         data[l][8], data[l][9], data[l][10], data[l][11], data[l][12], G + a + N])
                    break
        print(tb,file=out_file)
        curs.close()
        ora_db.close()
out_file.close()
err_file.close()