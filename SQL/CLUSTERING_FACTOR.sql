select t.TABLE_NAME ,i.LEAF_BLOCKS, i.CLUSTERING_FACTOR, t.NUM_ROWS, (t.BLOCKS - t.EMPTY_BLOCKS) Used_Blocks
 from user_indexes i, user_tables t
where i.TABLE_NAME = t.TABLE_NAME 
and i.TABLE_NAME = 'B_GE_T_GT'
