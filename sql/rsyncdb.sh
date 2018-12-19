export TMP=/tmp
export TMPDIR=$TMP
export ORACLE_BASE=/home/u01/app/oracle
export ORACLE_HOME=$ORACLE_BASE/product/11.2.0/db_1
export ORACLE_SID=ghsjdb
export ORACLE_TERM=xterm
export PATH=/usr/sbin:$PATH
export PATH=$ORACLE_HOME/bin:$PATH
# user locale variable
#export LANG=zh_CN.GB18030
export LANG=en_US.UTF-8
#export NLS_LANG=simplified chinese_china.ZHS16GBK
#export NLS_LANG=AMERICAN_AMERICA.ZHS32GB18030
export NLS_LANG=AMERICAN_AMERICA.UTF8

dmp_owner="SDE","FM","PLATFORM_APP","PLATFORM","GHSJ_PWGL","GHSJ_JCSJ","GHSJ_XM"
curtime=`date +%Y%m%d%H%M`
dmp_directory=DMP_DIRECTORY
dmp_file=dmp.dmp
log_exp_file="log_$curtime_exp.log"
job_exp_name="expdp_$curtime_job"
log_imp_file="log_$curtime_exp.log"
job_imp_name="expdp_$curtime_job"

echo " "
echo "============expdp impdp ==============="

expdp expdper/expdper@ghsjdb SCHEMAS=$dmp_owner dumpfile=$dmp_file DIRECTORY=$dmp_directory  PARALLEL=16  logfile=$log_exp_file job_name=$job_exp_name
#impdp impdper/impdper@ghsjdb_test SCHEMAS=$dmp_owner DIRECTORY=$dmp_directory DUMPFILE=$dmp_file TABLE_EXISTS_ACTION=replace LOGFILE=$log_imp_file JOB_NAME=$job_imp_name PARALLEL=16
