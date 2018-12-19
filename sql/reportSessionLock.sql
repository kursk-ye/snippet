SET ECHO off
REM NAME:   TFSLCKWT.SQL
REM USAGE:"@path/tfslckwt"
REM ------------------------------------------------------------------------
REM REQUIREMENTS:
REM    SELECT on V$SESSION, V$LOCK
REM ------------------------------------------------------------------------
REM PURPOSE:
REM    Reports users waiting for locks.
REM ------------------------------------------------------------------------
REM EXAMPLE:
REM    USERNAME          SID TYPE HELD        REQ           ID1        ID2
REM    --------------- ----- ---- ----------- ----------- -------- --------
REM    SYSTEM             12 TX   Exclusive   None          131087     2328
REM    SCOTT               7 TX   None        Exclusive     131087     2328
REM    SCOTT               8 TX   Exclusive   None          131099     2332
REM    SYSTEM             10 TX   None        Exclusive     131099     2332
REM    SYSTEM             12 TX   None        Exclusive     131099     2332
REM
REM ------------------------------------------------------------------------
REM Main text of script follows:

column username format  A15
column sid      format  9990    heading SID
column type     format  A4
column lmode    format  990     heading 'HELD'
column request  format  990     heading 'REQ'
column id1      format  9999990
column id2 format  9999990

SELECT sn.username, m.sid, m.type,
        DECODE(m.lmode, 0, 'None',
                        1, 'Null',
                        2, 'Row Share',
                        3, 'Row Excl.',
                        4, 'Share',
                        5, 'S/Row Excl.',
                        6, 'Exclusive',
                lmode, ltrim(to_char(lmode,'990'))) lmode,
        DECODE(m.request,0, 'None',
                         1, 'Null',
                         2, 'Row Share',
                         3, 'Row Excl.',
                         4, 'Share',
                         5, 'S/Row Excl.',
                         6, 'Exclusive',
                         request, ltrim(to_char(m.request,
                '990'))) request, m.id1, m.id2
FROM v$session sn, v$lock m
WHERE (sn.sid = m.sid AND m.request != 0)
        OR (sn.sid = m.sid
                AND m.request = 0 AND lmode != 4
                AND (id1, id2) IN (SELECT s.id1, s.id2
     FROM v$lock s
                        WHERE request != 0
              AND s.id1 = m.id1
                                AND s.id2 = m.id2)
                )
ORDER BY id1, id2, m.request;

