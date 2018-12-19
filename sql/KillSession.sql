SET LINESIZE 100
COLUMN spid FORMAT A10
COLUMN username FORMAT A10
COLUMN program FORMAT A45

SELECT s.inst_id,
       s.sid,
       s.serial#,
       p.spid,
       s.username,
       s.program
FROM   gv$session s
       JOIN gv$process p ON p.addr = s.paddr AND p.inst_id = s.inst_id
WHERE  s.type != 'BACKGROUND';

--Single
ALTER SYSTEM KILL SESSION 'sid,serial#'  IMMEDIATE;

ALTER SYSTEM DISCONNECT SESSION 'sid,serial#' POST_TRANSACTION;
ALTER SYSTEM DISCONNECT SESSION 'sid,serial#' IMMEDIATE;

-- Rac
  ALTER SYSTEM KILL SESSION 'sid,serial#,@inst_id'  IMMEDIATE;
