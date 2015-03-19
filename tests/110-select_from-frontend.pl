use warnings;
use strict;

use lib '..';
use Access;

my $db_frontend = 'c:\temp\frontend.accdb';

my $access = new Access;

my $db = $access -> open_db($db_frontend);

my $rs = $db->OpenRecordSet('
             select
               p.col_p,
               c.col_c
             from
               table_parent p inner join
               table_child  c on p.id = c.id_parent
             order by
               p.id,
               c.id
   ');

while (! $rs->{EOF}) {

  printf "%-15s %-5s\n", $rs->{Fields}->{col_p}->{Value}, $rs->{Fields}->{col_c}->{Value};
  $rs->MoveNext();
}
