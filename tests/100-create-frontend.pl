use warnings;
use strict;

use lib '..';
use Access;

my $db_backend = 'c:\temp\backend.accdb';
my $db_name    = 'c:\temp\frontend.accdb';

my $access = new Access;

#  TODO: This is sort of stupid, first to create the db, then
#        to open it...
my $db     = $access -> create_db($db_name);
my $db     = $access -> open_db  ($db_name);


$access -> link_mdb_table($db_backend, 'table_parent');
$access -> link_mdb_table($db_backend, 'table_child' );

$db     -> close;
$access -> close;
$access -> quit;
