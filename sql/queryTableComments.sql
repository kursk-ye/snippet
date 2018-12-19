select * from all_tab_comments
where substr(table_name,1,4) != 'BIN$'
and table_name = 'B_GE_T_GT_C'
/
