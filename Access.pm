use warnings;
use strict;

use Win32::OLE;

use Access::Form;

package Access;

#   For acForm, acSaveYes etc
use Win32::OLE::Const 'Microsoft.Access';
#   After a Windows-Update, the 
#     previous use Win32::OLE::Const...
#     resulted in a "No type library matching "Microsoft.Access" found at ..."
#     But replacing it with a ...
#   use Win32::OLE::Const 'DAO.DBEngine';
#      ... seemed then to work
#      The respective GUID is 4AFFC9A0-5F99-101B-AF4E-00AA003F0F07 (?)
# -------------------------------------------------------------------------

#   For dbLangGeneral etc:
use Win32::OLE::Const 'Microsoft.DAO';

# The following line is needed for the constant vbext_ct_StdModule:
# BTW, The respective GUID is: 0002E157-0000-0000-C000-000000000046
use Win32::OLE::Const 'Microsoft Visual Basic for Applications Extensibility 5.3';

sub scale_x {
  my $x = shift;

  return $x * 15;
}

sub scale_y {
  my $y = shift;
  return $y * 15;
}

sub new {

  my $self = {};

  bless $self, shift;

  $self->{access} = new Win32::OLE 'Access.Application' or die 'Access.Application';

    return $self;
}

sub create_db {
  my $self              = shift;
  my $mdb_name_and_path = shift;

  unlink $mdb_name_and_path if -e $mdb_name_and_path;

  my $db  = $self->{access}->DBEngine->Workspaces(0)->CreateDatabase($mdb_name_and_path, dbLangGeneral, 0) or die "CreateDatabase $!";

  return $db;
}

sub open_db {

  my $self              = shift;
  my $mdb_name_and_path = shift;

   $self->{access} -> OpenCurrentDatabase($mdb_name_and_path);

  my $db = $self->{access}->DBEngine->Workspaces(0)->Databases(0);

  return $db; 
}

# { misc

sub apply_properties_on_obj {
  my $obj               = shift;
  my $property_hash_ref = shift;

  for my $key (keys %$property_hash_ref) {
    $obj->{$key} = $property_hash_ref->{$key};
  }
}

sub apply_events_on_form {
  my $self           = shift;
  my $form_name      = shift;
  my $event_hash_ref = shift;

  for my $key (keys %$event_hash_ref) {
    $self -> create_event_proc_for_form($form_name, $key, $event_hash_ref->{$key});
  }
}

sub link_mdb_table {
  my $self       = shift;
  my $mdb_name   = shift;
  my $table_name = shift;

  $self->{access}->DoCmd->TransferDatabase(acLink, 'Microsoft Access', $mdb_name, acTable, $table_name, $table_name);
}

# }

# {{{ Database related

sub select_identity { # {{{
  my $self = shift;
  my $db   = shift;

  my $rs = $db -> OpenRecordSet('select @@identity');

  return $rs->{Fields}->Item(0)->{Value};

# Alternatively, use
#    select @@identity AS FOO 
#    reutrn $rs->{'FOO'}->{Value}

} # }}}

sub exec_with_identity { # {{{
  my $self = shift;
  my $db   = shift;
  my $sql  = shift;

  $db -> execute($sql);
  return $self->select_identity($db);

    
} # }}}

# }}}


# { Form related
sub create_form {
  # usually, after calling create_form, you'll want to call open_form(), 
  # then manipulate the form, then call close_form().
  my $self  = shift;
  my $opts  = shift;

  my $form = new Form($self);

  my $form_name_orig = $form->{form}->{Name}; 
  my $form_name      = $opts    ->{name};

  apply_properties_on_obj($form -> {form}, $opts->{property}) if exists $opts->{property};

  $form -> close;
  $self -> {access} -> DoCmd -> Rename($form_name, acForm, $form_name_orig);

  $form -> {name} = $form_name;

  return $form;
}

sub do_cmd {
  my $self = shift;
  return $self->{access}->DoCmd;
}

sub create_form_deprecated {
  # usually, after calling create_form, you'll want to call open_form(), 
  # then manipulate the form, then call close_form().

  my $self   = shift;
  my $opts   = shift;

  my $form_obj = $self->{access}->CreateForm();

  die "couldn't create form, is current database opened? [\$access->OpenCurrentDatabase()]" unless $form_obj;

  my $form_name_orig = $form_obj->{Name}; 
  my $form_name      = $opts    ->{name};

  if (exists $opts->{caption}) {
    $form_obj->{Caption} = $opts->{caption};
  }

  $self->{access}->DoCmd->Close (acForm, $form_name_orig, acSaveYes);
  $self->{access}->DoCmd->Rename($form_name, acForm, $form_name_orig);

}

sub close_form {
  my $self      = shift;
  my $form_name = shift;

  $self->{access}->DoCmd->Close(acForm, $form_name, acSaveYes);
}

sub form_obj_from_form_name {
  my $self      = shift;
  my $form_name = shift;

  my $form_obj = $self->{access}->Forms($form_name);
  # last line could also be written as:
  # my $form_obj = $access->Forms->Item($form_name);

  die "form_obj_from_form_name (form_name: $form_name was not found). Is the form currently opened?" unless $form_obj;

  return $form_obj;
}

sub set_startup_form {
# TODO: Does this even work?
  my $self        = shift;
  my $db          = shift;
  my $form        = shift;

  $self -> create_X_property($db, "StartupForm", dbText, $form->{name});
}


sub create_event_proc_for_form {
# see also create_event_proc_for_control_on_form
  my $self         = shift;
  my $form_name    = shift;
  my $event_name   = shift;
  my $event_code   = shift;

  my $form_obj    =  self -> form_obj_from_form_name($form_name);
  my $form_module = $form_obj->Module();

  my $first_line_of_event_proc = $form_module->CreateEventProc($event_name, 'Form');

  die "first_line_of_event_proc not defined for from: $form_name, event_name: $event_name" unless defined $first_line_of_event_proc;
  $form_module->InsertLines($first_line_of_event_proc + 1, $event_code);
}


# form related }

# { control related

sub create_control_on_form_use_this {
  my $self = shift;
  my $opts = shift;

  my $section;
  my $parent  = "";
  my $ac_section;
  if (defined $opts->{section}) {
      $section = $opts->{section};
  }
  else {
      $section = 'detail';
  }

  if ($opts->{parent}) {
      $parent = $opts->{parent};
  }

  if    ($section eq 'header') { $ac_section = acHeader; }
  elsif ($section eq 'detail') { $ac_section = acDetail; }
  elsif ($section eq 'footer') { $ac_section = acFooter; }
  else {die "unknown section $section in create_control_on_form"};

  my $ctrl_obj = $self->{access}->CreateControl({
                     FormName    =>  $opts->{form_name} , 
                     ControlType =>  $opts->{control_type},
                     Left        =>  scale_x($opts->{rect}->{x}),
                     Top         =>  scale_y($opts->{rect}->{y}),
                     Width       =>  scale_x($opts->{rect}->{w}),
                     Height      =>  scale_y($opts->{rect}->{h}),
                     Section     =>  $ac_section,
                     parent      =>  $parent,
                   });

  apply_properties_on_obj($ctrl_obj, $opts->{property});

  
  $self -> apply_events_on_ctrl($opts->{form_name}, $ctrl_obj, $opts->{event}) if exists $opts->{event};

  return $ctrl_obj;
}

sub create_control_on_form { # TODO: DEPRECATED

  my $self         = shift;
  my $form_name    = shift;
  my $control_name = shift;
  my $control_type = shift;
  my $x            = shift;
  my $y            = shift;
  my $w            = shift;
  my $h            = shift;
  my $section      = shift;

  $section = 'detail' unless defined $section;

  my $ac_section;

  if    ($section eq 'header') { $ac_section = acHeader; }
  elsif ($section eq 'detail') { $ac_section = acDetail; }
  elsif ($section eq 'footer') { $ac_section = acFooter; }
  else {die "unknown section $section in create_control_on_form"};

  my $control_obj = $self->{access}->CreateControl({
                      FormName    =>  $form_name , 
                      ControlType =>  $control_type,
                      Left        =>  scale_x($x),
                      Top         =>  scale_y($y),
                      Width       =>  scale_x($w),
                      Height      =>  scale_y($h),
                      Section     =>  $ac_section,
                    });

  $control_obj->{Name} = $control_name;

  return $control_obj;
}

sub ctrl_obj_from_ctrl_name {
  my $self      = shift;
  my $form_name = shift;
  my $ctrl_name = shift;

  my $form_obj = $self -> form_obj_from_form_name($form_name);

  my $ctrl_obj = $form_obj->{$ctrl_name};
  return $ctrl_obj;
}

sub create_label_on_form { # TODO: DEPRECATED

  my $self          = shift;
  my $form_name     = shift;
  my $lbl_name      = shift;
  my $text          = shift;
  my $x             = shift;
  my $y             = shift;
  my $w             = shift;
  my $h             = shift;
  my $section       = shift;

  $section = 'detail' unless $section;

  my $label_obj = $self -> create_control_on_form ($form_name, $lbl_name , acLabel, $x, $y, $w, $h, $section);

  $label_obj->{Caption} = $text;
  return $label_obj;
}

# }

# module related  {
# 
#   # see also http://www.cpearson.com/excel/vbe.aspx
# 
#  TODO: Should this not return the "module object"?
sub insert_module {
  my $self        = shift;
  my $module_name = shift;

  # The application's VBE property represents the Microsoft Visual Basic for Applications editor.
  my $module_obj = $self -> {access}->VBE()->ActiveVBProject()->VBComponents->Add(vbext_ct_StdModule);
  $module_obj->{Name} = $module_name;
# $self -> {access} ->DoCmd->Close(acModule, $module_name, acSaveYes);
  $self -> do_cmd -> Close (acModule, $module_name, acSaveYes);
}
# 
sub write_to_module {
  # see also write_to_form_module()
# my $access      = shift;
  my $self        = shift;

  my $module_name = shift;
  my $text        = shift;

# my $module_obj      = $access->VBE()->ActiveVBProject()->VBComponents($module_name);
  my $module_obj      = $self  ->{access}->VBE()->ActiveVBProject()->VBComponents($module_name);

  my $module_code_obj = $module_obj -> CodeModule();

  $module_code_obj -> AddFromString($text);

# $access->DoCmd->Close(acModule, $module_name, acSaveYes);
  $self -> do_cmd -> Close(acModule, $module_name, acSaveYes);

}

sub add_typelib_guid {
  my $self     = shift;
  my $guid     = shift;

  my $opening_curly_braces = chr(123);

  if (substr($guid,0,1) ne $opening_curly_braces) {
     $guid = "{$guid}";
  }

  my $ref = $self->{access}->VBE()->ActiveVBProject()->References();

  $ref->AddFromGuid($guid, 0, 0);
}

sub create_X_property {
  my $self           = shift;  # TODO: Unused!
  my $X              = shift;  # can be db, fld...
  my $property_name  = shift;
  my $property_type  = shift;
  my $property_value = shift;

  my $property = $X->CreateProperty($property_name, $property_type, $property_value);
  $X->Properties->Append($property);
}


# }

sub close {
  my $self = shift;
  $self->{access} -> Close();
}

sub quit {
  my $self = shift;
  $self->{access} -> Quit();
}

1;
