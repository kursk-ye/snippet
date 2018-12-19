run following from sysdba:
select owner||'.'||object_name obj
,oracle_username||' ('||s.status||')' oruser
,os_user_name osuser
,machine computer
,l.process unix
,''''||s.sid||','||s.serial#||'''' ss
,r.name rs
,to_char(s.logon_time,'yyyy/mm/dd hh24:mi:ss') time
from v$locked_object l
,dba_objects o
,v$session s
,v$transaction t
,v$rollname r
where l.object_id = o.object_id
and s.sid=l.session_id
and s.taddr=t.addr
and t.xidusn=r.usn
order by osuser, ss, obj


To kill a perticular session :
alter system kill session '<>';
e.g.
alter system kill session '101,1120';
