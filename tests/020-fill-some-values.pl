use warnings;
use strict;

use lib '..';
use Access;

my $db_name = 'c:\temp\backend.accdb';

my $access = new Access;
my $db     = $access->open_db($db_name) or die "Could not open $db_name";

my $id = $access->exec_with_identity($db, "insert into table_parent (col_p) values ('first record')");
$db -> execute("insert into table_child(id_parent, col_c) values ($id, 'foo')");
$db -> execute("insert into table_child(id_parent, col_c) values ($id, 'bar')");
$db -> execute("insert into table_child(id_parent, col_c) values ($id, 'baz')");

$id = $access->exec_with_identity($db, "insert into table_parent (col_p) values ('second record')");
$db -> execute("insert into table_child(id_parent, col_c) values ($id, 'abc')");
$db -> execute("insert into table_child(id_parent, col_c) values ($id, 'def')");
$db -> execute("insert into table_child(id_parent, col_c) values ($id, 'ghi')");
