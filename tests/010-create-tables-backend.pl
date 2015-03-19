use warnings;
use strict;

use lib '..';
use Access;

my $db_name = 'c:\temp\backend.accdb';


my $access = new Access;
my $db     = $access -> create_db($db_name);

create_schema();

$db     -> close;
$access -> close;
$access -> quit;


sub create_schema { # {{{
    
    $db -> execute('
      create table table_parent (
        id     autoincrement primary key,
        col_p  text
      )
    ');

    $db -> execute('
      create table table_child (
        id         autoincrement primary key,
        id_parent  int,
        col_c      text,
        constraint fk_parent_child foreign key (id_parent) references table_parent
      )
    ');

} # }}}
