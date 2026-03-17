#!/usr/bin/perl
use strict;
use warnings;
use utf8;
use Encode;

# Fix UTF-8 mojibake (UTF-8 bytes read as Latin-1)
sub fix_mojibake {
  my $x = shift;
  return $x if !defined $x || $x eq '' || $x =~ /[\x{400}-\x{4ff}]/;
  my $d = eval { decode('UTF-8', encode('ISO-8859-1', $x), Encode::FB_CROAK) };
  defined $d ? $d : $x;
}

# Parse ERP xlsx - extract from sheet1.xml (inlineStr format)
my $sheet = 'xlsx_erp/xl/worksheets/sheet1.xml';
open my $fh, '<:encoding(UTF-8)', $sheet or die $!;
my $content = do { local $/; <$fh> };
close $fh;

sub excel_date {
    my ($n) = @_;
    return '' unless $n && $n =~ /^[\d.]+$/;
    my $days = int($n);
    my $epoch = ($days - 25569) * 86400;  # 25569 = 1970-01-01 in Excel
    my ($d,$m,$y) = (localtime($epoch))[3,4,5];
    return sprintf("%04d-%02d-%02d", $y+1900, $m+1, $d);
}

sub col_idx {
    my ($ref) = @_;
    $ref =~ /^([A-Z]+)/;
    my $col = $1;
    my $idx = 0;
    for my $c (split //, $col) {
        $idx = $idx * 26 + (ord($c) - ord('A') + 1);
    }
    return $idx - 1;
}

my %col_map = (A=>0, B=>1, C=>2, D=>3, E=>4, F=>5, G=>6, H=>7, I=>8, J=>9, K=>10, L=>11, M=>12);

my @tasks;
while ($content =~ /<row r="(\d+)">(.*?)<\/row>/gs) {
    my ($row_num, $row_content) = ($1, $2);
    next if $row_num == 1; # skip header
    
    my %cells;
    while ($row_content =~ /<c r="([A-Z]+)(\d+)"[^>]*>(.*?)<\/c>/gs) {
        my ($col, $val_content) = ($1, $3);
        my $idx = col_idx($col);
        my $val = '';
        if ($val_content =~ /<is><t>([^<]*)<\/t><\/is>/) {
            $val = $1;
        } elsif ($val_content =~ /<v>([^<]+)<\/v>/) {
            my $n = $1;
            if ($idx == 5 || $idx == 6 || $idx == 7) { # F,G,H - dates
                $val = excel_date($n);
            } else {
                $val = $n;
            }
        }
        $cells{$idx} = fix_mojibake($val) if $val ne '';
    }
    
    my $id = $cells{0} || '';
    next unless $id =~ /^(UCR|MCE|ERP)-?\d+/i;
    
    push @tasks, {
        id => $id,
        title => $cells{3} || '',
        state => $cells{10} || 'Unknown',
        type => $cells{8} || '',
        subsystem => $cells{12} || '-',
        created => $cells{5} || '',
        updated => $cells{6} || '',
    };
}

# Output JSON
use JSON;
use Encode;
binmode STDOUT, ':utf8';
my $project = decode('UTF-8', "\xD0\x95\xD0\xA0\xD0\x9F Cloud");  # ЕРП Cloud
my $json = encode_json({ project => $project, tasks => \@tasks });
# Write raw UTF-8 bytes (encode_json produces bytes; :utf8 would double-encode)
open my $out, '>:raw', 'erp-tasks-data.json' or die $!;
print $out $json;
close $out;
