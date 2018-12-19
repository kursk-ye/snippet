select t1.index_name, t1.index_type, t2.COLUMN_NAME
from user_indexes t1, user_ind_columns t2
where t1.table_name = 'C_CONS' and t1.index_name = t2.index_name
