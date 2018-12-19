(1)前置条件及环境变量
导出数据库实例名：dghdb
导出数据库用户："SDE","FM","PLATFORM_APP","PLATFORM","GHSJ_PWGL","GHSJ_JCSJ","GHSJ_XM","GHSJ_GHJH"
操作系统:RHEL
ORACLE_BASE路径：/home/u01/app/oracle
ORACLE_HOME路径：/home/u01/app/oracle/product/11.2.0/db_1
操作系统环境变量LANG：en_US.UTF-8 #可通过echo $LANG获知，NLS_LANG须与LANG编码一致



(2)通过expdp在本地导出文件，详细导出步骤如下：

[1]登录系统后，切换到ORACLE用户下,创建DMP_DIRECTORY
#su - oracle
$sqlplus / as sysdba
sql>CREATE DIRECTORY DMP_DIRECTORY AS '/home/oracle';

[2]创建expdper用户，并授予权限
-- Create the user 
create user EXPDPER
	identified by expdper
  default tablespace USERS
  temporary tablespace TEMP
  profile DEFAULT;
-- Grant/Revoke object privileges 
grant read, write on directory DMP_DIRECTORY to EXPDPER;
-- Grant/Revoke role privileges 
grant connect to EXPDPER;
grant datapump_exp_full_database to EXPDPER;
grant datapump_imp_full_database to EXPDPER;
grant resource to EXPDPER;
-- Grant/Revoke system privileges 
grant unlimited tablespace to EXPDPER;


[3]通过vi创建以下脚本内容,并执行
$vi exp_sh


export ORACLE_BASE=/home/u01/app/oracle
export ORACLE_HOME=$ORACLE_BASE/product/11.2.0/db_1
export ORACLE_SID=dghdb
export ORACLE_TERM=xterm
export PATH=/usr/sbin:$PATH
export PATH=$ORACLE_HOME/bin:$PATH
export LANG=en_US.UTF-8
export NLS_LANG=AMERICAN_AMERICA.UTF8

dmp_owner="SDE","FM","PLATFORM_APP","PLATFORM","GHSJ_PWGL","GHSJ_JCSJ","GHSJ_XM","GHSJ_GHJH"
dmp_file=dmp20130809.dmp
log_exp_file="exp.log"

expdp expdper/expdper@$ORACLE_SID SCHEMAS=$dmp_owner dumpfile=exp%U.dmp DIRECTORY=DMP_DIRECTORY  PARALLEL=8  logfile=$log_exp_file

expdp expdper/expdper@migration FULL=y dumpfile=migration.dmp DIRECTORY=DMP_DIRECTORY  logfile=migration.log

sql>CREATE DIRECTORY DMP_DIRECTORY AS 'd:/exp_dmp/';

-- Create the user 
create user EXPDPER
	identified by expdper
  default tablespace USERS
  temporary tablespace TEMP
  profile DEFAULT;
-- Grant/Revoke object privileges 
grant read, write on directory DMP_DIRECTORY to EXPDPER;
-- Grant/Revoke role privileges 
grant connect to EXPDPER;
grant datapump_exp_full_database to EXPDPER;
grant datapump_imp_full_database to EXPDPER;
grant resource to EXPDPER;
-- Grant/Revoke system privileges 
grant unlimited tablespace to EXPDPER;

expdp expdper/expdper@$ORACLE_SID TABLES=table_name dumpfile=exp%U.dmp DIRECTORY=DMP_DIRECTORY  PARALLEL=32  logfile=exp.log STATUS=120
