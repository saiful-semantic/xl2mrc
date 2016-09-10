#!/usr/bin/perl
#
# Name: convert.pl
# Version: 0.1 (Nov. 2008)
#
# Description: Converts MS-Excel data into MARC21

use strict;
use XL2MARC;

my $converter = XL2MARC->new();

$converter->run();