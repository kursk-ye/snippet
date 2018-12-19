-- Grant/Revoke object privileges
-- Grant/Revoke system privileges
grant administer database trigger to ecs;
grant alter any index to ecs;
grant alter any table to ecs;
grant analyze any to ecs;
grant create any index to ecs;
grant create any procedure to ecs;
grant create any sequence to ecs;
grant create any trigger to ecs;
grant create indextype to ecs;
grant create library to ecs;
grant create operator to ecs;
grant create procedure to ecs;
grant create public synonym to ecs;
grant create sequence to ecs;
grant create session to ecs;
grant create table to ecs;
grant create trigger to ecs;
grant create type to ecs;
grant create view to ecs;
grant create database link to ecs;
grant create materialized view to ecs;
grant drop any index to ecs;
grant drop any procedure to ecs;
grant drop any sequence to ecs;
grant drop any table to ecs;
grant drop any view to ecs;
grant drop public synonym to ecs;
grant execute any procedure to ecs;
grant select any sequence to ecs;
grant select any table to ecs;
grant unlimited tablespace to ecs;
GRANT exp_full_database TO ecs;
GRANT imp_full_database TO ecs;
GRANT debug connect session to ecs; 
GRANT debug any procedure to ecs;

select s.sid, s.serial#, s.status, p.spid
from v$session s, v$process p
where s.username = 'ecs'
and p.addr  = s.paddr;

   INST_ID        SID    SERIAL#
---------- ---------- ----------
         1         11       8911
         1         12      21355
         1         16       9025
         1         17        409
         1         18         33
         1         19          7
         1         20         23
         1         21          9
         1         22          7
         1         23          7
         1         71      20177

   INST_ID        SID    SERIAL#
---------- ---------- ----------
         1         72       8463
         1         73      19157
         1         77      21823
         1         79          7
         1         80          7
         1         81          9
         1         82         41
         1         83          7
         1         84          7
         1        130       2751
         1        134       5097

   INST_ID        SID    SERIAL#
---------- ---------- ----------
         1        137         93
         1        138        429
         1        141        233
         1        142          7
         1        143          7
         1        144          5
         1        145          7
         1        146          7
         1        193        567
         1        201       8455
         1        204      49491

   INST_ID        SID    SERIAL#
---------- ---------- ----------
         1        205         15
         1        206        441
         1        207         13
         1        208          7
         1        209          7
         1        210         63
         1        211          7
