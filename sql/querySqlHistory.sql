--query special table be accessed history info in 7 day
select t.sql_id
       ,sql.sql_text
       ,sql.sql_fulltext
       ,t.optimizer_cost
       ,t.optimizer_mode
       ,t.optimized_physical_reads_total
       ,t.optimized_physical_reads_delta
       ,t.sharable_mem
       ,t.fetches_total
       ,t.sorts_total
       ,t.executions_total
       ,t.parse_calls_total
       ,t.disk_reads_total
       ,t.direct_writes_total
       ,t.buffer_gets_total
       ,t.cpu_time_total
       ,t.elapsed_time_total
       ,t.iowait_total
       ,t.clwait_total
       ,t.apwait_total
       ,t.ccwait_total
       ,t.plsexec_time_total
       ,t.io_interconnect_bytes_total
       ,t.physical_read_requests_total
       ,t.physical_read_bytes_total
       ,s.begin_interval_time
from   dba_hist_sqlstat t, dba_hist_snapshot s,(select q.sql_id,q.SQL_TEXT,q.SQL_FULLTEXT
from v$sql q
where q.SQL_TEXT like '%PL_CALC_LLR_TG%') sql
where  t.snap_id = s.snap_id
and    t.dbid = s.dbid
and    t.instance_number = s.instance_number
and    s.begin_interval_time between trunc(sysdate)-7 and trunc(sysdate)
and    t.sql_id = sql.sql_id;
/

--query performance of sql of special table be accessed  in 7 day
select sql.sql_id
       ,min(sql.sql_text)
       ,count(sql.sql_id) timers
       ,sum(t.elapsed_time_total) elapsed_time_sum
       ,sum(t.iowait_total)
       ,(sum(t.iowait_total)/sum(t.elapsed_time_total) ) iowait_rato
       ,sum(t.clwait_total)
       ,sum(t.apwait_total)
       ,sum(t.ccwait_total)
       ,sum(t.executions_total)
       ,sum(t.parse_calls_total)
       ,sum(t.physical_read_requests_total)
from   dba_hist_sqlstat t, dba_hist_snapshot s,(select q.sql_id,q.SQL_TEXT,q.SQL_FULLTEXT
from v$sql q
where q.SQL_TEXT like '%PL_CALC_LLR_TG%') sql
where  t.snap_id = s.snap_id
and    t.dbid = s.dbid
and    t.instance_number = s.instance_number
and    s.begin_interval_time between trunc(sysdate)-7 and trunc(sysdate)
and    t.sql_id = sql.sql_id
group by sql.sql_id
order  by  elapsed_time_sum desc,timers desc


-- query performance of sql run by speical user  last 7 day
-- get user_id from DBA_USERS
select sql.sql_id
       ,min(his_sess.user_id)
       ,min(sql.sql_text)
       ,count(sql.sql_id) timers
       ,sum(t.elapsed_time_total) elapsed_time_sum
       ,sum(t.iowait_total)
       ,(sum(t.iowait_total)/sum(t.elapsed_time_total) ) iowait_rato
       ,sum(t.clwait_total)
       ,sum(t.apwait_total)
       ,sum(t.ccwait_total)
       ,sum(t.executions_total)
       ,sum(t.parse_calls_total)
       ,sum(t.physical_read_requests_total)
from   dba_hist_sqlstat t, dba_hist_snapshot s,(select q.sql_id,q.SQL_TEXT,q.SQL_FULLTEXT from v$sql q ) sql ,
  (select sql_id, user_id 
  from dba_hist_active_sess_history h 
  where  (h.user_id ='105' or h.user_id = '106')) his_sess
where  t.snap_id = s.snap_id
and    t.dbid = s.dbid
and    t.instance_number = s.instance_number
and    s.begin_interval_time between trunc(sysdate)-7 and trunc(sysdate)
and    t.sql_id = sql.sql_id
and    t.sql_id = his_sess.sql_id
group by sql.sql_id
order  by  elapsed_time_sum desc,timers desc
