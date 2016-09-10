# Name: XL2MARC.pm
# Version: 0.1 (Nov. 2008)
#
# Description: Converts MS-Excel data into MARC outputs. Handles file upload and uses two modules for conversion.

package XL2MARC;
use CGI::Pretty qw/:html3/;
use base 'CGI::Application';
use strict;

# Disable in production use
use warnings;

use XLView qw/wizard_conv custom_conv/;
use MARCView qw/marc_wizard marc_custom/;

# Set upload limit (bandwidth consideration)
$CGI::POST_MAX = 1024 * 500;    # 500kb

# Setup the application
sub setup {
    my $self = shift;
    $self->start_mode('upload');
    $self->run_modes(
        'upload'      => 'upload_file',
        'wizard_view' => 'marc_wizard',
        'custom_view' => 'marc_custom',
        'Wizard'      => 'wizard_conv',
        'Custom'      => 'custom_conv'
    );
}

# First screen for uploading Excel file
sub upload_file {
    my $self = shift;

    # Create CGI object
    my $q = $self->query();

    my $output = '';

    $output .= $q->start_html(
        -title   => 'Excel to MARC',
        -Style   => { 'src' => '/marc/marc.css' },
        -BGCOLOR => '#94B6D8'
    );
    $output .= $q->h2('Excel to MARC');
    $output .= "<div align=\"center\" class=\"upload\">\n";
    $output .= $q->start_multipart_form();
    $output .= $q->filefield( 'uploaded_file', 50 );
    $output .= $q->p;
    $output .= $q->checkbox( 'headers', 'checked', 'ON',
        'First row contains column headers' );
    $output .= $q->p;
    $output .= $q->radio_group(
        -name    => 'rm',
        -value   => [ 'Wizard', 'Custom' ],
        -default => ['Wizard']
    );
    $output .= $q->p;
    $output .= $q->submit( -class => 'button', -value => 'Submit' );
    $output .= $q->end_form;
    $output .= "</div>";

    #	$output .= $footer;
    $output .= $self->footer();
    $output .= $q->end_html;

    return $output;
}

# Page footer to be added in each screen
sub footer {
    my $footer = "<br/>&copy; 2008 Saiful Amin.\n";
    return $footer;
}

1;
