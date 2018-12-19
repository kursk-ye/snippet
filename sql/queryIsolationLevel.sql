SELECT s.sid, s.serial#,
   CASE BITAND(t.flag, POWER(2, 28))
      WHEN 0 THEN 'READ COMMITTED'
      ELSE 'SERIALIZABLE'
   END AS isolation_level
FROM v$transaction t 
JOIN v$session s ON t.addr = s.taddr AND s.sid = sys_context('USERENV', 'SID');

#If you are not already in a transaction you can start one with the following:

declare 
   trans_id Varchar2(100);
begin
   trans_id := dbms_transaction.local_transaction_id( TRUE );
end;
/
