run
{
allocate channel ch1 device type disk; #one disk one channel
backup  as COMPRESSED backupset database format '/home/oracle/rmanback/db_%U_%d';
backup archivelog all format '/home/oracle/rmanback/arch_%t_%s' delete all input;
backup format '/home/oracle/rmanback/con_%s_%p' current controlfile;
crosscheck backup;
crosscheck archivelog all;
delete noprompt expired backup;
delete noprompt obsolete;
delete noprompt archivelog all;
release channel ch1;
}
