CREATE TABLE PL_CALC_LLR_TG_NEW
  (
      tg_id      VARCHAR2(32),
      tg_no      VARCHAR2(20),
      d_prov_pq  NUMBER(16,4),
      d_spq      NUMBER(16,4),
      d_ll       NUMBER(16,4),
      d_llr      NUMBER(16,4),
      stat_day   DATE,
      send_time  DATE,
      city_no    VARCHAR2(50) default '42406',
      cc_cnt     NUMBER,
      read_ratio NUMBER
  )
  PARTITION BY RANGE (stat_day)    INTERVAL (NUMTOYMINTERVAL(1,'MONTH'))
        SUBPARTITION BY LIST (city_no)
        SUBPARTITION TEMPLATE (
        subpartition GROUP1 values ('42401') tablespace TOP_MON_OWNER_NEW,
        subpartition GROUP2 values ('42402') tablespace TOP_MON_OWNER_NEW,
        subpartition GROUP3 values ('42403') tablespace TOP_MON_OWNER_NEW,
        subpartition GROUP4 values ('42404') tablespace TOP_MON_OWNER_NEW,
        subpartition GROUP5 values ('42405') tablespace TOP_MON_OWNER_NEW,
        subpartition GROUP6 values ('42406') tablespace TOP_MON_OWNER_NEW,
        subpartition GROUP7 values ('42407') tablespace TOP_MON_OWNER_NEW,
        subpartition GROUP8 values ('42408') tablespace TOP_MON_OWNER_NEW,
        subpartition GROUP9 values ('42409') tablespace TOP_MON_OWNER_NEW,
        subpartition GROUP10 values ('42410') tablespace TOP_MON_OWNER_NEW,
        subpartition GROUP11 values ('42411') tablespace TOP_MON_OWNER_NEW,
        subpartition GROUP12 values ('42412') tablespace TOP_MON_OWNER_NEW,
        subpartition GROUP13 values ('42413') tablespace TOP_MON_OWNER_NEW,
        subpartition GROUP14 values ('42414') tablespace TOP_MON_OWNER_NEW
        )
  (partition values less than  (TO_DATE(' 2015-02-01 00:00:00', 'SYYYY-MM-DD HH24:MI:SS') ) );
