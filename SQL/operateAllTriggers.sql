select 
	'alter table ' || TABLE_OWNER || '.' || TABLE_NAME || ' enable all triggers;' 
from all_triggers 
where TABLE_OWNER in ('FM','GHSJ_JCSJ','PLATFORM','GHSJ_XM');
