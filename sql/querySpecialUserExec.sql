-- 查询特定用户和特定客户端程序执行过的sql
-- 本例用户是ghsj_jcsj,客户端程序是TableaU
select 
sqltext.FIRST_LOAD_TIME,
sqltext.LAST_ACTIVE_TIME,
sqltext.*
  from v$sql sqltext, v$session sesion
 where 
sesion.PREV_HASH_VALUE = sqltext.hash_value
and sesion.PREV_SQL_ADDR = sqltext.address
   and sesion.username='GHSJ_JCSJ'
   and sesion.PROGRAM = 'tabprotosrv.exe'
   order by sqltext.LAST_ACTIVE_TIME
