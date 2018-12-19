select 
	'alter table ' || OWNER || '.' || TABLE_NAME || ' disable constraint ' || CONSTRAINT_NAME || ';' noprimarykey  
from dba_constraints 
where owner in ('FM','GHSJ_XM','GHSJ_JCSJ','PLATFORM') and CONSTRAINT_TYPE <> 'P';


select 
	'alter table ' || OWNER || '.' || TABLE_NAME || ' disable constraint ' || CONSTRAINT_NAME || ';' noprimarykey  
from dba_constraints 
where owner in ('SDE') and CONSTRAINT_TYPE <> 'P';

select 
	'alter table ' || OWNER || '.' || TABLE_NAME || ' enable constraint ' || CONSTRAINT_NAME || ';' noprimarykey  
from dba_constraints 
where owner in ('FM','GHSJ_XM','GHSJ_JCSJ','PLATFORM','SDE') and CONSTRAINT_TYPE <> 'P';select 
	'alter table ' || OWNER || '.' || TABLE_NAME || ' disable constraint ' || CONSTRAINT_NAME || ';' noprimarykey  
from dba_constraints 
where owner in ('FM','GHSJ_XM','GHSJ_JCSJ','PLATFORM') and CONSTRAINT_TYPE <> 'P';


select 
	'alter table ' || OWNER || '.' || TABLE_NAME || ' disable constraint ' || CONSTRAINT_NAME || ';' noprimarykey  
from dba_constraints 
where owner in ('SDE') and CONSTRAINT_TYPE <> 'P';

select 
	'alter table ' || OWNER || '.' || TABLE_NAME || ' enable constraint ' || CONSTRAINT_NAME || ';' noprimarykey  
from dba_constraints 
where owner in ('FM','GHSJ_XM','GHSJ_JCSJ','PLATFORM','SDE') and CONSTRAINT_TYPE <> 'P';
