-- usage:
-- exec  runstats_pkg.rs_start;
-- run produces1
-- exec runstats_pkg.rs_middle
-- run produces2
-- exec runstats_pkg.rs_stop(100);

create or replace view stats
as select 'STAT...' || a.name name, b.value
      from v$statname a, v$mystat b
     where a.statistic# = b.statistic#
    union all
    select 'LATCH.' || name,  gets
      from v$latch;

create global temporary table run_stats
( runid varchar2(15),
  name varchar2(80),
  value int )
on commit preserve rows;


create or replace package runstats_pkg
as
    procedure rs_start;
    procedure rs_middle;
    procedure rs_stop( p_difference_threshold in number default 0 );
end;
/

create or replace package body runstats_pkg
as

g_start      number;
g_run1     number;
g_run2     number;

procedure rs_start
is
begin
    delete from run_stats;

    insert into run_stats
    select 'before', stats.* from stats;

    g_start := dbms_utility.get_time;
end;

procedure rs_middle
is
begin
    g_run1 := (dbms_utility.get_time-g_start);

    insert into run_stats
    select 'after 1', stats.* from stats;
    g_start := dbms_utility.get_time;

end;

procedure rs_stop(p_difference_threshold in number default 0)
is
begin
    g_run2 := (dbms_utility.get_time-g_start);

    dbms_output.put_line
        ( 'Run1 ran in ' || g_run1 || ' hsecs' );
    dbms_output.put_line
        ( 'Run2 ran in ' || g_run2 || ' hsecs' );
    dbms_output.put_line
    ( 'run1:run2  rato is ' || round(g_run1/g_run2,2)  );
        dbms_output.put_line( chr(9) );

    insert into run_stats
    select 'after 2', stats.* from stats;

    dbms_output.put_line
    ( rpad( 'Name', 30 ) || lpad( 'Run1', 10 ) ||
      lpad( 'Run2', 10 ) || lpad( 'Diff', 10 ) );

    for x in
    ( select rpad( a.name, 30 ) ||
             to_char( b.value-a.value, '9,999,999' ) ||
             to_char( c.value-b.value, '9,999,999' ) ||
             to_char( ( (c.value-b.value)-(b.value-a.value)), '9,999,999' ) data
        from run_stats a, run_stats b, run_stats c
       where a.name = b.name
         and b.name = c.name
         and a.runid = 'before'
         and b.runid = 'after 1'
         and c.runid = 'after 2'
         and (c.value-a.value) > 0
         and abs( (c.value-b.value) - (b.value-a.value) )
               > p_difference_threshold
       order by abs( (c.value-b.value)-(b.value-a.value))
    ) loop
        dbms_output.put_line( x.data );
    end loop;

        dbms_output.put_line( chr(9) );
    dbms_output.put_line( 'Run1 latches total versus runs -- difference and pct' );
    dbms_output.put_line( lpad( 'Run1', 10 ) || lpad( 'Run2', 10 ) ||  lpad( 'Diff', 10 ) || lpad( '   run1:run2', 10 ) );

    for x in
    ( select to_char( run1, '9,999,999' ) ||
             to_char( run2, '9,999,999' ) ||
             to_char( diff, '9,999,999' ) ||
             to_char( round( run1/run2 ,2 ), '999.99' ) data
        from ( select sum(b.value-a.value) run1, sum(c.value-b.value) run2,
                      sum( (c.value-b.value)-(b.value-a.value)) diff
                 from run_stats a, run_stats b, run_stats c
                where a.name = b.name
                  and b.name = c.name
                  and a.runid = 'before'
                  and b.runid = 'after 1'
                  and c.runid = 'after 2'
                  and a.name like 'LATCH%'
                )
    ) loop
        dbms_output.put_line( x.data );
    end loop;
end;

end;
/


