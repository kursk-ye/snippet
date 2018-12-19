SET ECHO off
REM NAME:   TFSMLOCK.SQL
REM USAGE:"@path/tfsmlock"
REM ------------------------------------------------------------------------
REM REQUIREMENTS:
REM    SELECT on V_$LOCK, V_$SESSION, USER$, OBJ$
REM ------------------------------------------------------------------------
REM PURPOSE:
REM    The following locking information script provides the following
REM    items: SessionID, Serial #, Username, Terminal, Resource
REM    Type, Table, Owner, ID1, ID2, Lock Held, Lock Requested.
REM
REM    The TFTS series contains scripts to provide (less detailed) lock
REM    information in a simpler format (TFSLLOCK.SQL), and a script which
REM    breaks the details out further (TFSCLOCK.SQL).
REM ------------------------------------------------------------------------
REM EXAMPLE:
REM    Too complex to fit a representative sample here.
REM ------------------------------------------------------------------------
REM DISCLAIMER:
REM    This script is provided for educational purposes only. It is NOT
REM   supported by Oracle World Wide Technical Support.
REM    The script has been tested and appears to work as intended.
REM You should always run new scripts on a test instance initially.
REM ------------------------------------------------------------------------
REM Main text of script follows:  set pagesize 60
set linesize 200 feedback off heading on
column sid format 999
column res heading 'Resource Type' format a20
column id1 format 9999999
column id2 format 9999999
column lmode heading 'Lock Held' format a14
column request heading 'Lock Req.' format a14
column serial# format 99999
column username  format a10
column terminal heading Term format a6
column tab format a10
column owner format a8
select  l.sid,s.serial#,s.username,s.terminal,
        decode(l.type,'RW','RW - Row Wait Enqueue',
                    'TM','TM - DML Enqueue',
                    'TX','TX - Trans Enqueue',
                    'UL','UL - User',l.type||' - System') res,
        substr(t.name,1,10) tab,u.name owner,
        l.id1,l.id2,
        decode(l.lmode,1,'No Lock',
                2,'Row Share',
                3,'Row Exclusive',
                4,'Share',
                5,'Shr Row Excl',
                6,'Exclusive',null) lmode,
        decode(l.request,1,'No Lock',
                2,'Row Share',
                3,'Row Excl',
                4,'Share',
                5,'Shr Row Excl',
                6,'Exclusive',null) request
from v$lock l, v$session s,
sys.user$ u,sys.obj$ t
where l.sid = s.sid
and s.type != 'BACKGROUND'
and t.obj# = l.id1
and u.user# = t.owner#
/
set feedback on
