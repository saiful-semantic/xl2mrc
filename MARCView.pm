# Name: MARCView.pm
# Version: 0.1 (Nov. 2008)
# Description: Converts data passed through param into different MARC outputs

# TODO: UTF-8 encoding
# TODO: Add punctuation (if desired by user) - Custom mode
# TODO: Sort MARC output by tag number (preserving field order)

package MARCView;
use CGI::Pretty qw/:html3/;
use base 'CGI::Application';
use strict;
use warnings;

require Exporter;

use vars qw(@ISA @EXPORT_OK);
@ISA       = qw(Exporter);
@EXPORT_OK = qw(marc_wizard marc_custom);

my ( $d, $m, $y ) = (gmtime)[ 3 .. 5 ];
my $current_time = ( $y + 1900 ) . ( $m + 1 ) . $d;

my $date = substr( $current_time, 2 );

# View MARC in Wizard mode
sub marc_wizard {
    my $self = shift;

    # Create CGI object
    my $q = $self->query();

    my $mode = $q->param('mode');

    my $data_view = '';
    use MARC::File::XML;
    use MARC::File::MARCMaker;
    use MARC::File::ISIS;

    $data_view .= $q->start_html(
        -title   => 'Excel to MARC - Wizard',
        -Style   => { 'src' => '/marc/marc.css' },
        -BGCOLOR => '#94B6D8'
    );
    $data_view .= $q->h2('Excel to MARC - Wizard');

    if ( $mode eq 'Preview' ) {
        $data_view .= "Preview records";
        $data_view .= "<br /><br />\n";
        $data_view .= "<div align=\"left\" class=\"preview\">\n<pre>\n";
        $data_view .= $q->button(
            -class   => 'button',
            -value   => 'Back',
            -onClick => "history.back()"
        );
        $data_view .= "<br /><br />\n";
    }
    else {
        $data_view .= "MARC21 format"    if $mode eq 'MARC21';
        $data_view .= "MARCXML format"   if $mode eq 'MARCXML';
        $data_view .= "MARCMaker format" if $mode eq 'MARCMaker';
        $data_view .= "CDS/ISIS format"  if $mode eq 'ISIS';

        $data_view .= "<br /><br />\n";
        $data_view .= "<div align=\"left\" class=\"preview\">\n";
        $data_view .= '<form name="isoview" action="#">' . "\n";
        $data_view .= '<textarea name="marc" rows="20" cols="82">' . "\n";

        # Add the MARCXML headers
        $data_view .= MARC::File::XML::header() if $mode eq 'MARCXML';
    }

    my @cols = $q->param('col_ids');
    my ( @used_data_fields, %field_map ) = ();
    my @labels = qw/tag repeatable/;

    # Get the MARC mapping table to find used columns
    for my $col (@cols) {

        # Map column number with array position
        my $data_col = $col - 1;

        for my $label (@labels) {
            if ( $q->param( 'tag' . $col ) ) {

                # Create the field map
                $field_map{ $label . $data_col } = $q->param( $label . $col );
            }
        }

        # Identify used fields
        push @used_data_fields, $data_col if $q->param( 'tag' . $col );
    }

    # Access param() values
    my @data       = $q->param('data');
    my $rec_status = $q->param('rec_status');
    my $gmd        = $q->param('gmd');
    my $bib_level  = $q->param('biblevel');
    my $p_medium   = $q->param('p_medium');
    my $item_type  = $q->param('item_type');
    $item_type = '[' . $item_type . ']' if $item_type;
    my $lng    = $q->param('lng');
    my $prefix = $q->param('prefix');
    my $agency = $q->param('agency');

    # Now create MARC records
    use MARC::Record;

    my $rec_count = 0;
    foreach (@data) {
        my @data_row = split /\|/, $_;
        $rec_count++;

        my $record = MARC::Record->new();

        my ( $title, $author, $tag, $sub );

        # Prepare the leader
        my $leader = '00054nam#a22002891a 4500';
        substr( $leader, 5, 1 ) = $rec_status if $rec_status;
        substr( $leader, 6, 1 ) = $gmd        if $gmd;
        substr( $leader, 7, 1 ) = $bib_level  if $bib_level;

        $record->leader($leader);

        # Control Number
        my $control_no = '';
        if ($prefix) {
            $control_no = $prefix . sprintf( "%06d", $rec_count );
        }
        else {
            $control_no = sprintf( "%06d", $rec_count );
        }

        # Append control number field
        $record->append_fields( MARC::Field->new( '001', $control_no ) );

        # Create and append mandatory tag_008
        my $data_008 = "$date|                r|    ||111|eng||||";
        substr( $data_008, 23, 1 ) = $p_medium if $p_medium;

        # Do padding (lang code must be 3 chars)
        substr( $data_008, 35, 4 ) = sprintf( "%-3s", $lng ) if $lng;

        my $tag_008 = MARC::Field->new( '008', $data_008 );
        $record->append_fields($tag_008) if $mode ne 'ISIS';

        if ($agency) {
            my $tag_040 = MARC::Field->new( '040', '', '', 'a' => $agency );
            $record->append_fields($tag_040);
        }

        if ($lng) {
            my $tag_041 = MARC::Field->new( '041', '', '', 'a' => $lng );
            $record->append_fields($tag_041);
        }

        # Create fields data from each row
        foreach my $i (@used_data_fields) {

            # Normalize text into sentence case
            my $this_field = sentence_case( $data_row[$i] );

            # Find author field
            if ( $field_map{ 'tag' . $i } == 100 ) {
                $author = $data_row[$i];

                # Normalize author into title case
                $author = title_case($author);

                # Check for multiple authors
                if ( $author =~ /;/ ) {
                    my @authors = split( /;/, $author );

                    # Retrieve the first author
                    my $first_author = shift @authors;
                    my $primary_author = MARC::Field->new( '100', '1', '',
                        'a' => "$first_author" );
                    $record->append_fields($primary_author);

                    foreach (@authors) {

                        # Create repeatable Author field
                        my $tag_700 =
                          MARC::Field->new( '700', '1', '', 'a' => "$_" );
                        $record->append_fields($tag_700);
                    }
                }
                else {
                    # Create single Author field
                    my $tag_100 =
                      MARC::Field->new( '100', '1', '', 'a' => "$author" );
                    $record->append_fields($tag_100);
                }

                # Look for title field
            }
            elsif ( $field_map{ 'tag' . $i } == 245 ) {
                $title = $this_field;

                my $title_ind1 = 1;
                my $title_ind2 = 0;
                $title_ind2 = &get_title_ind2($title);

                # Create Title field
                if ( $title =~ /:/ ) {

                    # Separate title and subtitle
                    my ( $title_proper, $subtitle ) = split( /:/, $title );
                    $subtitle =~ s/^\s*//;
                    my $tag_245 = MARC::Field->new(
                        '245', $title_ind1, $title_ind2,
                        'a' => "$title_proper",
                        'b' => "$subtitle"
                    );
                    $tag_245->update( 'h' => $item_type ) if $item_type;
                    $record->append_fields($tag_245);
                }
                else {
                    my $tag_245 =
                      MARC::Field->new( '245', $title_ind1, $title_ind2,
                        'a' => "$title" );
                    $tag_245->update( 'h' => $item_type ) if $item_type;
                    $record->append_fields($tag_245);
                }
            }
            elsif ( $field_map{ 'tag' . $i } =~ /0[589]?/ ) {

                # Create Call number field
                my $tag_no = $field_map{ 'tag' . $i };
                my ( $class_no, $book_no ) = ();
                if ( $data_row[$i] =~ /(.+)\s+(.+)/ ) {
                    ( $class_no, $book_no ) = ( $1, $2 );
                }
                else {
                    $class_no = $data_row[$i];
                }

                my $call_no =
                  MARC::Field->new( $tag_no, '', '', 'a' => "$class_no" );
                $call_no->update( 'b' => "$book_no" ) if $book_no;
                $record->append_fields($call_no);

                # Electronic resource (URL)
            }
            elsif ( $field_map{ 'tag' . $i } == 856 ) {

                #
                my $tag_856 =
                  MARC::Field->new( '856', 4, 2, 'u' => $data_row[$i] );
                $record->append_fields($tag_856);

                # All other fields
            }
            else {
                # Create other fields
                my $tag_no = $field_map{ 'tag' . $i };

                #				my $before = find_tag($record, $tag_no);
                my $other_tag =
                  MARC::Field->new( $tag_no, '', '', 'a' => "$this_field" );

                #				$record->insert_fields_after($before, $other_tag);
                $record->append_fields($other_tag);
            }

        }

        # Display the records
        if ( $mode eq 'Preview' ) {

            $data_view .= $record->as_formatted() . "\n\n";

        }
        elsif ( $mode eq 'MARCMaker' ) {

            $data_view .= MARC::File::MARCMaker->encode($record);

        }
        elsif ( $mode eq 'ISIS' ) {

            #			$data_view .= MARC::File::ISIS->encode( $record );
            $data_view .= $record->as_isis();

        }
        elsif ( $mode eq 'MARC21' ) {

            $data_view .= $record->as_usmarc();

        }
        elsif ( $mode eq 'MARCXML' ) {

            $data_view .= MARC::File::XML::record($record);

        }
    }

    # Add the footer display
    if ( $mode eq 'Preview' ) {

        $data_view .= "</pre>\n";
        $data_view .= $q->button(
            -class   => 'button',
            -value   => 'Back',
            -onClick => "history.back()"
        );

    }
    else {

        $data_view .= MARC::File::XML::footer() if $mode eq 'MARCXML';
        $data_view .= '</textarea><br/><br/>' . "\n";

        $data_view .= $q->button(
            -class => 'button',
            -value => 'Select All',
            -onClick =>
              "javascript:this.form.marc.focus();this.form.marc.select();"
        );

        $data_view .= '&nbsp;';
        $data_view .= $q->button(
            -class   => 'button',
            -value   => 'Back',
            -onClick => "history.back()"
        );
        $data_view .= '</form>';
    }

    $data_view .= "</div>\n";
    $data_view .= $self->footer();
    $data_view .= $q->end_html;

    # Return the form
    return $data_view;
}

# View MARC in Custom mode
sub marc_custom {
    my $self = shift;

    # Create CGI object
    my $q = $self->query();

    my $mode = $q->param('mode');

    my $data_view = '';
    use MARC::File::XML;
    use MARC::File::MARCMaker;
    use MARC::File::ISIS;

    $data_view .= $q->start_html(
        -title   => 'Excel to MARC - Custom',
        -Style   => { 'src' => '/marc/marc.css' },
        -BGCOLOR => '#94B6D8'
    );
    $data_view .= $q->h2('Excel to MARC - Wizard');

    if ( $mode eq 'Preview' ) {
        $data_view .= "Preview records";
        $data_view .= "<br /><br />\n";
        $data_view .= "<div align=\"left\" class=\"preview\">\n<pre>\n";
        $data_view .= $q->button(
            -class   => 'button',
            -value   => 'Back',
            -onClick => "history.back()"
        );
        $data_view .= "<br /><br />\n";
    }
    else {
        $data_view .= "MARC21 format"    if $mode eq 'MARC21';
        $data_view .= "MARCXML format"   if $mode eq 'MARCXML';
        $data_view .= "MARCMaker format" if $mode eq 'MARCMaker';
        $data_view .= "CDS/ISIS format"  if $mode eq 'ISIS';

        $data_view .= "<br /><br />\n";
        $data_view .= "<div align=\"left\" class=\"preview\">\n";
        $data_view .= '<form name="isoview" action="#">' . "\n";
        $data_view .= '<textarea name="marc" rows="20" cols="82">' . "\n";

        # Add the MARCXML headers
        $data_view .= MARC::File::XML::header() if $mode eq 'MARCXML';
    }

    my @cols = $q->param('col_ids');
    my ( @used_data_fields, %field_map ) = ();
    my @labels = qw/tag repeatable/;

    # Get the MARC mapping table to find used columns
    for my $col (@cols) {

        # Map column number with array position
        my $data_col = $col - 1;

        for my $label (@labels) {
            if ( $q->param( 'tag' . $col ) ) {

                # Create the field map
                $field_map{ $label . $data_col } = $q->param( $label . $col );
            }
        }

        # Identify used fields
        push @used_data_fields, $data_col if $q->param( 'tag' . $col );
    }

    # Access param() values
    my @data       = $q->param('data');
    my $rec_status = $q->param('rec_status');
    my $gmd        = $q->param('gmd');
    my $bib_level  = $q->param('biblevel');
    my $p_medium   = $q->param('p_medium');
    my $item_type  = $q->param('item_type');
    $item_type = '[' . $item_type . ']' if $item_type;
    my $lng    = $q->param('lng');
    my $prefix = $q->param('prefix');
    my $agency = $q->param('agency');

    # Now create MARC records
    use MARC::Record;

    my $rec_count = 0;
    foreach (@data) {
        my @data_row = split /\|/, $_;
        $rec_count++;

        my $record = MARC::Record->new();

        my ( $title, $author, $tag, $sub );

        # Prepare the leader
        my $leader = '00054nam#a22002891a 4500';
        substr( $leader, 5, 1 ) = $rec_status if $rec_status;
        substr( $leader, 6, 1 ) = $gmd        if $gmd;
        substr( $leader, 7, 1 ) = $bib_level  if $bib_level;

        $record->leader($leader);

        # Control Number
        my $control_no = '';
        if ($prefix) {
            $control_no = $prefix . sprintf( "%06d", $rec_count );
        }
        else {
            $control_no = sprintf( "%06d", $rec_count );
        }

        # Append control number field
        $record->append_fields( MARC::Field->new( '001', $control_no ) );

        # Create and append mandatory tag_008
        my $data_008 = "$date|                r|    ||111|eng||||";
        substr( $data_008, 23, 1 ) = $p_medium if $p_medium;

        # Do padding (lang code must be 3 chars)
        substr( $data_008, 35, 4 ) = sprintf( "%-3s", $lng ) if $lng;

        my $tag_008 = MARC::Field->new( '008', $data_008 );
        $record->append_fields($tag_008) if $mode ne 'ISIS';

        if ($agency) {
            my $tag_040 = MARC::Field->new( '040', '', '', 'a' => $agency );
            $record->append_fields($tag_040);
        }

        if ($lng) {
            my $tag_041 = MARC::Field->new( '041', '', '', 'a' => $lng );
            $record->append_fields($tag_041);
        }

        # Create fields data from each row
        foreach my $i (@used_data_fields) {

            # Normalize text into sentence case
            my $this_field = sentence_case( $data_row[$i] );

            # Find author field
            if ( $field_map{ 'tag' . $i } == 100 ) {
                $author = $data_row[$i];

                # Normalize author into title case
                $author = title_case($author);

                # Check for multiple authors
                if ( $author =~ /;/ ) {
                    my @authors = split( /;/, $author );

                    # Retrieve the first author
                    my $first_author = shift @authors;
                    my $primary_author = MARC::Field->new( '100', '1', '',
                        'a' => "$first_author" );
                    $record->append_fields($primary_author);

                    foreach (@authors) {

                        # Create repeatable Author field
                        my $tag_700 =
                          MARC::Field->new( '700', '1', '', 'a' => "$_" );
                        $record->append_fields($tag_700);
                    }
                }
                else {
                    # Create single Author field
                    my $tag_100 =
                      MARC::Field->new( '100', '1', '', 'a' => "$author" );
                    $record->append_fields($tag_100);
                }

                # Look for title field
            }
            elsif ( $field_map{ 'tag' . $i } == 245 ) {
                $title = $this_field;

                my $title_ind1 = 1;
                my $title_ind2 = 0;
                $title_ind2 = &get_title_ind2($title);

                # Create Title field
                if ( $title =~ /:/ ) {

                    # Separate title and subtitle
                    my ( $title_proper, $subtitle ) = split( /:/, $title );
                    $subtitle =~ s/^\s*//;
                    my $tag_245 = MARC::Field->new(
                        '245', $title_ind1, $title_ind2,
                        'a' => "$title_proper",
                        'b' => "$subtitle"
                    );
                    $tag_245->update( 'h' => $item_type ) if $item_type;
                    $record->append_fields($tag_245);
                }
                else {
                    my $tag_245 =
                      MARC::Field->new( '245', $title_ind1, $title_ind2,
                        'a' => "$title" );
                    $tag_245->update( 'h' => $item_type ) if $item_type;
                    $record->append_fields($tag_245);
                }
            }
            elsif ( $field_map{ 'tag' . $i } =~ /0[589]?/ ) {

                # Create Call number field
                my $tag_no = $field_map{ 'tag' . $i };
                my ( $class_no, $book_no ) = ();
                if ( $data_row[$i] =~ /(.+)\s+(.+)/ ) {
                    ( $class_no, $book_no ) = ( $1, $2 );
                }
                else {
                    $class_no = $data_row[$i];
                }

                my $call_no =
                  MARC::Field->new( $tag_no, '', '', 'a' => "$class_no" );
                $call_no->update( 'b' => "$book_no" ) if $book_no;
                $record->append_fields($call_no);

                # Electronic resource (URL)
            }
            elsif ( $field_map{ 'tag' . $i } == 856 ) {

                #
                my $tag_856 =
                  MARC::Field->new( '856', 4, 2, 'u' => $data_row[$i] );
                $record->append_fields($tag_856);

                # All other fields
            }
            else {
                # Create other fields
                my $tag_no = $field_map{ 'tag' . $i };

                #				my $before = find_tag($record, $tag_no);
                my $other_tag =
                  MARC::Field->new( $tag_no, '', '', 'a' => "$this_field" );

                #				$record->insert_fields_after($before, $other_tag);
                $record->append_fields($other_tag);
            }

        }

        # Display the records
        if ( $mode eq 'Preview' ) {

            $data_view .= $record->as_formatted() . "\n\n";

        }
        elsif ( $mode eq 'MARCMaker' ) {

            $data_view .= MARC::File::MARCMaker->encode($record);

        }
        elsif ( $mode eq 'ISIS' ) {

            #			$data_view .= MARC::File::ISIS->encode( $record );
            $data_view .= $record->as_isis();

        }
        elsif ( $mode eq 'MARC21' ) {

            $data_view .= $record->as_usmarc();

        }
        elsif ( $mode eq 'MARCXML' ) {

            $data_view .= MARC::File::XML::record($record);

        }
    }

    # Add the footer display
    if ( $mode eq 'Preview' ) {

        $data_view .= "</pre>\n";
        $data_view .= $q->button(
            -class   => 'button',
            -value   => 'Back',
            -onClick => "history.back()"
        );

    }
    else {

        $data_view .= MARC::File::XML::footer() if $mode eq 'MARCXML';
        $data_view .= '</textarea><br/><br/>' . "\n";

        $data_view .= $q->button(
            -class => 'button',
            -value => 'Select All',
            -onClick =>
              "javascript:this.form.marc.focus();this.form.marc.select();"
        );

        $data_view .= '&nbsp;';
        $data_view .= $q->button(
            -class   => 'button',
            -value   => 'Back',
            -onClick => "history.back()"
        );
        $data_view .= '</form>';
    }

    $data_view .= "</div>\n";
    $data_view .= $self->footer();
    $data_view .= $q->end_html;

    # Return the form
    return $data_view;
}

# Other internal functions
# Find tag number for sorting
sub find_tag {

    # Get the MARC record object and tag ref
    my ( $rec, $num ) = @_;

    my $id;
    foreach ( $rec->fields() ) {
        $id = $_;
        last if $_->tag() >= $num;
    }
    return $id;
}

# Normalize title case
sub title_case {
    my $string = shift;
    $string = join ' ', map { ucfirst lc } split / /, $string;
    return $string;
}

# Normalize sentence case
sub sentence_case {
    my $string = shift;
    $string = lc($string);
    $string = ucfirst($string);
    return $string;
}

# Get title indicator
sub get_title_ind2 {
    my $given_title = shift;

    # Define the second indicator for title
    my $title_ind2 = 0;

    if ( $given_title =~ /^A / ) {
        $title_ind2 = 2;
    }
    elsif ( $given_title =~ /^An / ) {
        $title_ind2 = 3;
    }
    elsif ( $given_title =~ /^The / ) {
        $title_ind2 = 4;
    }
    return $title_ind2;
}

1;
