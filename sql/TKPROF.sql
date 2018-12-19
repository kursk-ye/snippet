--get tkporf file path
select rtrim(c.value,'/') || '/' || d.instance_name || '_ora_' || ltrim(to_char(a.spid)) || '.trc' 
from v$process a, v$session b, v$parameter c, v$instance d 
where a.addr = b.paddr 
and b.audsid = sys_context( 'userenv', 'sessionid') 
and c.name = 'user_dump_dest';
