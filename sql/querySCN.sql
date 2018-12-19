-- query scn max value
col maxscn for 999,999,999,999,999,999
select 
(
    (
        (
            (
                (   
                    (
                        to_char(sysdate,'YYYY')-1988
                    )*12+
                to_char(sysdate,'mm')-1
                )*31+to_char(sysdate,'dd')-1
            )*24+to_char(sysdate,'hh24')
        )*60+to_char(sysdate,'mi')
    )*60+to_char(sysdate,'ss')
) * to_number('ffff','XXXXXXXX')/4 as maxscn
from dual;

-- query retain days
col maxscn for 999,999,999,999,999,999
col retain_days for 999,999,999,999,999,999
select 
(
    (
        (
            (
                (   
                    (
                        to_char(sysdate,'YYYY')-1988
                    )*12+
                to_char(sysdate,'mm')-1
                )*31+to_char(sysdate,'dd')-1
            )*24+to_char(sysdate,'hh24')
        )*60+to_char(sysdate,'mi')
    )*60+to_char(sysdate,'ss')
) * to_number('ffff','XXXXXXXX')/4 maxscn,
( maxscn - dbms_flashback.get_system_change_number ) / 16 / 1024 / 3600 / 24 / 365 as retain_days
from dual;


