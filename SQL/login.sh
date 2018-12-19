set termout off
define _editor=vi
set serveroutput on size 1000000 format wrapped
col object_name for a30
col segment_name for a30
col file_name for a40
col name for a30
col what for a30 word_wrapped
col plan_plus_exp for a100 

set trimspool on
set long 5000
set linesize 131
set pagesize 0

define gname=idle
col global_name new_value gname

select lower(user) || '@' || 	substr(global_name, 1, decode(dot, 0, length(global_name), dot-1)) global_name
	from (select global_name, instr(global_name, ' . ') dot from global_name );
	
set termout on



REDEM set appinfo ON
REDEM set appinfo "SQL*Plus"
REDEM set arraysize 15
REDEM set autocommit OFF
REDEM set autoprint OFF
REDEM set autorecovery OFF
REDEM set autotrace off
REDEM set blockterminator "."
REDEM set cmdsep OFF
REDEM set colsep " "
REDEM set compatibility NATIVE
REDEM set concat "."
REDEM set copycommit 0
REDEM set copytypecheck ON
REDEM set define "&"
REDEM set describe DEPTH 1 LINENUM OFF INDENT ON
REDEM set markup HTML OFF SPOOL OFF ENTMAP ON PRE OFF
REDEM set echo OFF
REDEM set editfile "afiedt.buf"
REDEM set embedded OFF
REDEM set endbuftoken ""
REDEM set escape OFF
REDEM set feedback 6
REDEM set flagger OFF
REDEM set flush ON
REDEM SET HEADING ON
REDEM set headsep "|"
REDEM SET LINESIZE 250
REDEM set logsource ""
REDEM set long 80
REDEM set longchunksize 80
REDEM set newpage 1
REDEM set null ""
REDEM set numformat ""
REDEM set numwidth 10
REDEM SET PAGESIZE 24
REDEM set pause OFF
REDEM set recsep WRAP
REDEM set recsepchar " "
REDEM SET SERVEROUTPUT ON size 1000000 format WORD_WRAPPED
REDEM set shiftinout invisible
REDEM set showmode OFF
REDEM set sqlblanklines OFF
REDEM set sqlcase MIXED
REDEM set sqlcontinue "> "
REDEM set sqlnumber ON
REDEM set sqlprefix "#"
REDEM set sqlprompt "SQL> "
REDEM set sqlterminator ";"
REDEM set suffix "sql"
REDEM set tab ON
REDEM set termout ON
REDEM SET TIME ON
REDEM SET TIMING ON
REDEM SET TRIMOUT ON
REDEM set trimspool OFF
REDEM set underline "-"
REDEM set verify ON
REDEM set wrap ON
