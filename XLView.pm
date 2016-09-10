# Name: XLView.pm
# Version: 0.1 (Nov. 2008)
#
# Description: Converts MS-Excel data into CGI-param for conversion into MARC output

# TODO: OnSubmit 'fields' form validation for MARC tag value (numeric)
# TODO: Ask for encoding (UTF-8) - (to be implemented in MARCView.pm)
# TODO: Filesize warning
# TODO: Add punctuation (if desired by user) - Custom mode
# TODO: Link to MARC help pages (http://www.loc.gov/marc/bibliographic/bdsummary.html)

package XLView;
use CGI::Pretty qw/:html3/;
use base 'CGI::Application';
use strict;
use warnings;

use vars qw(@ISA @EXPORT_OK);
require Exporter;
@ISA       = qw(Exporter);
@EXPORT_OK = qw(wizard_conv custom_conv);

# Set master variables
my %marc_labels = (
    '020' => 'ISBN',
    '022' => 'ISSN',
    '050' => 'Call No (LC)',
    '082' => 'Call No (Dewey)',
    '080' => 'Call No (UDC)',
    '090' => 'Call No (Local)',
    '100' => 'Main Entry (Personal name)',
    '110' => 'Main Entry (Corporate name)',
    '111' => 'Main Entry (Meeting name)',
    '245' => 'Title',
    '250' => 'Edition',
    '260' => 'Publication (Imprint)',
    '300' => 'Extent (pages)',
    '310' => 'Frequency',
    '440' => 'Series title',
    '650' => 'Subject',
    '500' => 'Notes',
    '856' => 'Electronic Location (URL)'
);

# MARC tag 'values' for selecting field from fields.tag_value
# - manage blank tag value on 'Select'
my @tags =
  ( '', sort { $marc_labels{$a} cmp $marc_labels{$b} } keys %marc_labels );
$marc_labels{''} = 'Select';

my %rec_status_labels = (
    'a' => 'Encoding increased',
    'c' => 'Corrected/Revised',
    'd' => 'Deleted',
    'n' => 'New'
);
my @rec_status_keys = sort keys %rec_status_labels;

my %gmd_labels = (
    'a' => 'Language material',
    'c' => 'Noteted music',
    'e' => 'Cartographic material',
    'g' => 'Project medium',
    'i' => 'Lecture recording',
    'j' => 'Music recording',
    'm' => 'Computer file',
    'o' => 'Kit',
    'p' => 'Mixed material',
    'r' => '3D artifact',
    't' => 'Manuscript'
);
my @gmd_label_keys = sort keys %gmd_labels;

my %bib_level = (
    'a' => 'Monographic component',
    'b' => 'Serial component',
    'c' => 'Collection',
    'd' => 'Subunit',
    'i' => 'Integrating resource',
    'm' => 'Monograph/Item',
    's' => 'Serial'
);
my @bib_label_keys = sort keys %bib_level;

my %p_medium_labels = (
    '#' => 'None of the following',
    'a' => 'Microfilm',
    'b' => 'Microfiche',
    'c' => 'Microopaque',
    'd' => 'Large print',
    'f' => 'Braille',
    'r' => 'Regular print',
    's' => 'Electronic',
    '|' => 'No attempt to code'
);
my @p_medium_keys = sort keys %p_medium_labels;

my @item_types = (
    'Text',      'Journal', 'Reference', 'Bound volume',
    'Microform', 'CD-ROM',  'DVD',       'Electronic resource'
);

sub wizard_conv {
    my $self = shift;

    # Create CGI object
    my $q = $self->query();

    my $conv_form = '';

    $conv_form .= $q->start_html(
        -title   => 'Excel to MARC - Wizard',
        -Style   => { 'src' => '/marc/marc.css' },
        -BGCOLOR => '#94B6D8'
    );
    $conv_form .= $q->h2('Excel to MARC - Wizard');
    $conv_form .= "<div align=\"left\" class=\"form\">\n";

    my ( $file, $fh );
    if ( $file = $q->param('uploaded_file') ) {
        use Spreadsheet::ParseExcel::Simple;
        unless ( $fh = Spreadsheet::ParseExcel::Simple->read($file) ) {
            $conv_form .= "\'$file\' is not a valid Excel file.<br/><br/>\n";
            $conv_form .= $q->button(
                -class   => 'button',
                -value   => 'Back',
                -onClick => "history.back()"
            );
            $conv_form .= "</div>";

            $conv_form .= $self->footer();
            $conv_form .= $q->end_html;

            return $conv_form;
        }
    }
    else {
        $conv_form .= "No file selected.<br/><br/>\n";
        $conv_form .= $q->button(
            -class   => 'button',
            -value   => 'Back',
            -onClick => "history.back()"
        );
        $conv_form .= "</div>";

        $conv_form .= $self->footer();
        $conv_form .= $q->end_html;

        return $conv_form;
    }

    # Start the field mapping form/table ('fields')
    my $form = $q->start_form( -name => 'fields' );

    # Start the Main table
    $form .= $q->start_table( { -cellpadding => 1 } );
    $form .= "<tr valign=\"top\"><td>\n";

    $form .= $q->start_table( { -cellpadding => 2 } );

    $form .= '<tr><td>Record Status:</td><td>';
    $form .= $q->popup_menu(
        -name    => 'rec_status',
        -value   => \@rec_status_keys,
        -default => ['n'],
        -labels  => \%rec_status_labels
    );
    $form .= "</td></tr>\n";

    $form .= '<tr><td>Type of Record:</td><td>';
    $form .= $q->popup_menu(
        -name    => 'gmd',
        -value   => \@gmd_label_keys,
        -default => ['a'],
        -labels  => \%gmd_labels
    );
    $form .= "</td></tr>\n";

    $form .= '<tr><td>Bibliographic Level:</td><td>';
    $form .= $q->popup_menu(
        -name    => 'biblevel',
        -value   => \@bib_label_keys,
        -default => ['a'],
        -labels  => \%bib_level
    );
    $form .= "</td></tr>\n";

    $form .= '<tr><td>Physical Medium:</td><td>';
    $form .= $q->popup_menu(
        -name    => 'p_medium',
        -value   => \@p_medium_keys,
        -default => ['r'],
        -labels  => \%p_medium_labels
    );
    $form .= "</td></tr>\n";

    $form .= '<tr><td>Item Type:</td><td>';
    $form .= $q->popup_menu(
        -name    => 'item_type',
        -value   => \@item_types,
        -default => ['Text']
    );
    $form .= "</td></tr>\n";

    $form .= '<tr><td>Language Code:</td><td>';
    $form .= $q->textfield(
        -name      => 'lng',
        -default   => 'eng',
        -Size      => 8,
        -maxlength => 3
    );
    $form .= "</td></tr>\n";

    $form .= '<tr><td>Control Number Prefix:</td><td>';
    $form .= $q->textfield(
        -name      => 'prefix',
        -default   => 'xconv',
        -Size      => 8,
        -maxlength => 6
    );
    $form .= "</td></tr>\n";

    $form .= '<tr><td>Cataloguing Agency:</td><td>';
    $form .= $q->textfield(
        -name      => 'agency',
        -default   => 'xConv',
        -Size      => 8,
        -maxlength => 6
    );
    $form .= "</td></tr>\n";

    # End the MARC data settings table
    $form .= $q->end_table;
    $form .= "</td><br />\n";

    # Start the MARC fields mapping table
    $form .= "<td>\n";    # Make it part of main table
    $form .=
      $q->start_table( { -border => 1, -cellpadding => 2, -class => 'grid' } );

    $form .= '<tr>';
    $form .= '<td class="header"><b>Column</b></td>';
    $form .= '<td class="header"><b>Map MARC21 field</b></td>';
    $form .= '</tr>';

    foreach my $sheet ( $fh->sheets ) {
        my $header_flag = 0;
        while ( $sheet->has_data ) {
            chomp;
            my @first_row = $sheet->next_row if $header_flag != 1;
            my @row = $sheet->next_row;

            # Handle upto 52 columns
            my @range = ( 'A' .. 'Z' );
            push @range, ( 'AA' .. 'AZ' );

            # Show the header information for mapping each column
            for ( my $i = 0 ; $i <= $#first_row ; $i++ ) {
                next if $first_row[$i] =~ /^$/;
                $header_flag = 1;

                my $col_header = 'Column ' . $range[$i];
                $col_header = $first_row[$i] if $q->param('headers');

                # Map array position with column number
                my $col_id = $i + 1;

                $form .=
                  '<tr><td class="dashed"><em>' . $col_header . "</em></td>";
                $form .= $q->hidden( 'col_ids', $col_id );
                $form .= $q->hidden( 'col_header' . $col_id, $col_header );

                $form .= '<td class="dashed">';
                $form .= $q->textfield(
                    -name      => 'tag' . $col_id,
                    -id        => 'tag' . $col_id,
                    -default   => '',
                    -Size      => 3,
                    -maxlength => 3
                );

                $form .= "&nbsp;";
                $form .= $q->popup_menu(
                    -name    => 'tag_value',
                    -value   => \@tags,
                    -default => [''],
                    -labels  => \%marc_labels,
                    -onChange =>
"document.fields.tag$col_id.value = this[this.selectedIndex].value;"
                );
                $form .= '</td>';

                $form .= "</tr>\n";
            }    # End of 'for' loop for finding the mapping fields

            if ($header_flag) {

                # Populate data in array
                $form .= $q->hidden( 'data', join( '|', @row ) );
            }
        }    # End of 'while' for current sheet

    }    # End of 'foreach' for sheets

    # End MARC mapping table
    $form .= $q->end_table;
    $form .= "</td></tr>\n";

    $form .= $q->p;
    $form .=
      $q->hidden( -name => 'rm', -value => 'wizard_view', -override => 1 );

    $form .= "<tr align=\"center\"><td colspan=\"2\"><br>&nbsp;<br>";
    $form .= $q->radio_group(
        -name    => 'mode',
        -value   => [ 'Preview', 'MARC21', 'MARCXML', 'MARCMaker', 'ISIS' ],
        -default => ['Preview']
    );
    $form .= "<tr align=\"center\"><td colspan=\"2\"><br>";
    $form .= $q->submit( -class => 'button', -value => 'View MARC' );
    $form .= "&nbsp;&nbsp;";
    $form .= $q->button(
        -class   => 'button',
        -value   => 'Back',
        -onClick => "history.back()"
    );
    $form .= "</td></tr>\n";

    # End the Main table
    $form .= $q->end_table;

    $form .= $q->end_form;

    $conv_form .= $form;
    $conv_form .= "</div>";

    $conv_form .= $self->footer();

    $conv_form .= $q->end_html;

    return $conv_form;
}

sub custom_conv {
    my $self = shift;

    # Create CGI object
    my $q = $self->query();

    my $conv_form = '';

    my $jscript = <<END;
	window.onload = function() {
    	setupDependencies('fields'); //name of form
  	};
END

    $conv_form .= $q->start_html(
        -title  => 'Excel to MARC - Custom',
        -Style  => { 'src' => '/marc/marc.css' },
        -Script => [
            {
                -type => 'text/javascript',
                -src  => '/marc/FormManager.js'
            },
            $jscript
        ],
        -BGCOLOR => '#94B6D8'
    );
    $conv_form .= $q->h2('Excel to MARC - Custom');
    $conv_form .= "<div align=\"left\" class=\"form\">\n";

    my ( $file, $fh );
    if ( $file = $q->param('uploaded_file') ) {
        use Spreadsheet::ParseExcel::Simple;
        unless ( $fh = Spreadsheet::ParseExcel::Simple->read($file) ) {
            $conv_form .= "\'$file\' is not a valid Excel file.<br/><br/>\n";
            $conv_form .= $q->button(
                -class   => 'button',
                -value   => 'Back',
                -onClick => "history.back()"
            );
            $conv_form .= "</div>";

            $conv_form .= $self->footer();
            $conv_form .= $q->end_html;

            return $conv_form;
        }
    }
    else {
        $conv_form .= "No file selected.<br/><br/>\n";
        $conv_form .= $q->button(
            -class   => 'button',
            -value   => 'Back',
            -onClick => "history.back()"
        );
        $conv_form .= "</div>";

        $conv_form .= $self->footer();
        $conv_form .= $q->end_html;

        return $conv_form;
    }

    # Start the field mapping form/table ('fields')
    my $form = $q->start_form( -name => 'fields' );

    # Start the Main table
    $form .= $q->start_table( { -cellpadding => 2 } );
    $form .=
"<tr><td colspan=\"3\"><strong>Set the defaults</strong><br/><br/></td></tr>\n";
    $form .= "<tr valign=\"top\"><td>\n";

    # Start the default settings table
    $form .= $q->start_table( { -cellpadding => 2 } );

    $form .= '<tr><td>Record Status:</td><td>';
    $form .= $q->popup_menu(
        -name    => 'rec_status',
        -value   => \@rec_status_keys,
        -default => ['n'],
        -labels  => \%rec_status_labels
    );
    $form .= "</td></tr>\n";

    $form .= '<tr><td>Type of Record:</td><td>';
    $form .= $q->popup_menu(
        -name    => 'gmd',
        -value   => \@gmd_label_keys,
        -default => ['a'],
        -labels  => \%gmd_labels
    );
    $form .= "</td></tr>\n";

    $form .= '<tr><td>Bibliographic Level:</td><td>';
    $form .= $q->popup_menu(
        -name    => 'biblevel',
        -value   => \@bib_label_keys,
        -default => ['a'],
        -labels  => \%bib_level
    );
    $form .= "</td></tr>\n";

    $form .= '<tr><td>Physical Medium:</td><td>';
    $form .= $q->popup_menu(
        -name    => 'p_medium',
        -value   => \@p_medium_keys,
        -default => ['r'],
        -labels  => \%p_medium_labels
    );
    $form .= "</td></tr>\n";

    # End the default settings table
    $form .= $q->end_table;
    $form .= "</td><br />\n";

    # Separator column
    $form .= "<td>&nbsp;\n";
    $form .= "</td>\n";

    # Extended default settings table
    $form .= "<td>\n";
    $form .= $q->start_table( { -cellpadding => 2 } );

    $form .= '<tr><td>Item Type:</td><td>';
    $form .= $q->popup_menu(
        -name    => 'item_type',
        -value   => \@item_types,
        -default => ['Text']
    );
    $form .= "</td></tr>\n";

    $form .= '<tr><td>Language Code:</td><td>';
    $form .= $q->textfield(
        -name      => 'lng',
        -default   => 'eng',
        -Size      => 8,
        -maxlength => 3
    );
    $form .= "</td></tr>\n";

    $form .= '<tr><td>Control Number Prefix:</td><td>';
    $form .= $q->textfield(
        -name      => 'prefix',
        -default   => 'xconv',
        -Size      => 8,
        -maxlength => 6
    );
    $form .= "</td></tr>\n";

    $form .= '<tr><td>Cataloguing Agency:</td><td>';
    $form .= $q->textfield(
        -name      => 'agency',
        -default   => 'xConv',
        -Size      => 8,
        -maxlength => 6
    );
    $form .= "</td></tr>\n";

    # End extended default settings table
    $form .= $q->end_table;
    $form .= "</td></tr>\n";

    # Section 2
    $form .=
"<tr><td colspan=\"3\"><br/><br/><strong>Map each column</strong><br/><br/></td></tr>\n";

    # Start the MARC fields mapping table
    $form .= "<td colspan=\"3\">\n";    # Make it part of main table
    $form .=
      $q->start_table( { -border => 1, -cellpadding => 2, -class => 'grid' } );

    $form .= '<tr>';
    $form .= '<td class="header"><b>Column</b></td>';
    $form .= '<td class="header"><b>Subfields</b></td>';
    $form .= '<td class="header"><b>Map MARC21</b></td>';
    $form .= '</tr>';

    foreach my $sheet ( $fh->sheets ) {
        my $header_flag = 0;
        while ( $sheet->has_data ) {
            chomp;
            my @first_row = $sheet->next_row if $header_flag != 1;
            my @row = $sheet->next_row;

            # Handle upto 52 columns
            my @range = ( 'A' .. 'Z' );
            push @range, ( 'AA' .. 'AZ' );

            for ( my $i = 0 ; $i <= $#first_row ; $i++ ) {
                next if $first_row[$i] =~ /^$/;
                $header_flag = 1;

                my $col_header = 'Column ' . $range[$i];
                $col_header = $first_row[$i] if $q->param('headers');

                # Map array position with column number
                my $col_id = $i + 1;

                $form .=
                  '<tr><td class="dashed"><em>' . $col_header . '</em></td>';
                $form .= $q->hidden( 'col_ids', $col_id );
                $form .= $q->hidden( 'col_header' . $col_id, $col_header );

                $form .= '<td class="dashed">';

#				$form .= $q->checkbox(-name => 'comp' . $col_id,-checked=> 0,-Value => 0, -label => '');
                $form .= $q->radio_group(
                    -name    => 'comp' . $col_id,
                    -value   => [ 'One', 'More' ],
                    -default => ['One']
                );
                $form .= "&nbsp;";
                $form .= $q->textfield(
                    -name      => 'seprator' . $col_id,
                    -default   => '',
                    -Size      => 1,
                    -maxlength => 2,
                    -class     => "DEPENDS ON comp$col_id BEING More"
                );
                $form .= '</td>';

                $form .= '<td class="dashed">';
                $form .= $q->textfield(
                    -name      => 'tag' . $col_id,
                    -id        => 'tag' . $col_id,
                    -default   => '',
                    -Size      => 3,
                    -maxlength => 3
                );

                $form .= "&nbsp;";
                $form .= $q->textfield(
                    -name      => 'subf' . $col_id,
                    -default   => '',
                    -Size      => 1,
                    -maxlength => 1,
                    -class     => "DEPENDS ON comp$col_id BEING One"
                );
                $form .= $q->textfield(
                    -name      => 'subf' . $col_id,
                    -default   => '',
                    -Size      => 5,
                    -maxlength => 8,
                    -class     => "DEPENDS ON comp$col_id BEING More"
                );
                $form .= "&nbsp;";
                $form .= '</td>';

                $form .= "</tr>\n";
            }    # End of 'for' loop for finding the mapping fields

            if ($header_flag) {

                # Populate data in array
                $form .= $q->hidden( 'data', join( '|', @row ) );
            }
        }    # End of 'while' for current sheet

    }    # End of 'foreach' for sheets

    # End MARC mapping table
    $form .= $q->end_table;
    $form .= "</td></tr>\n";

    $form .= $q->p;
    $form .=
      $q->hidden( -name => 'rm', -value => 'custom_view', -override => 1 );

    $form .= "<tr align=\"center\"><td colspan=\"3\"><br>&nbsp;<br>";
    $form .= $q->radio_group(
        -name    => 'mode',
        -value   => [ 'Preview', 'MARC21', 'MARCXML', 'MARCMaker', 'ISIS' ],
        -default => ['Preview']
    );
    $form .= "<tr align=\"center\"><td colspan=\"3\"><br>";
    $form .= $q->submit( -class => 'button', -value => 'View MARC' );
    $form .= "&nbsp;&nbsp;";
    $form .= $q->button(
        -class   => 'button',
        -value   => 'Back',
        -onClick => "history.back()"
    );
    $form .= "</td></tr>\n";

    # End the Main table
    $form .= $q->end_table;

    $form .= $q->end_form;

    $conv_form .= $form;
    $conv_form .= "</div>";

    $conv_form .= $self->footer();

    $conv_form .= $q->end_html;

    return $conv_form;
}

sub custom_conv_o {
    my $self = shift;

    # Create CGI object
    my $q = $self->query();

    my $custom_view = '';
    $custom_view .= $q->start_html(
        -title   => 'Excel to MARC - Custom',
        -Style   => { 'src' => '/marc/marc.css' },
        -BGCOLOR => '#94B6D8'
    );
    $custom_view .= $q->h2('Excel to MARC - Custom');
    $custom_view .= "<div align=\"left\" class=\"upload\">\n";
    $custom_view .= $q->p(
        'Yet to be implemented.<br/><br/>Please check back later.<br/><br/>');
    $custom_view .= $q->button(
        -class   => 'button',
        -value   => 'Back',
        -onClick => "history.back()"
    );
    $custom_view .= "</div>";
    $custom_view .= $self->footer();
    $custom_view .= $q->end_html;

    return $custom_view;
}

1;
