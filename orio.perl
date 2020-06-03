#!/usr/bin/perl
use Spreadsheet::WriteExcel;
use Time::Piece;
use Unicode::Map();
($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst) = localtime();
my $csvFilename   = $ARGV[0];
my $FinalPath   = $ARGV[1];
my $date = localtime->strftime('%m/%d/%Y');
my $workbook = Spreadsheet::WriteExcel->new("$FinalPath/Hits-Report.xls");
my$worksheet = $workbook->add_worksheet("Hits-Count");
my $format0 = $workbook->add_format(
center_across => 1,
color => "black",
align => "vcenter",
bg_color => "green",
);
my $format = $workbook->add_format(
 center_across => 1,
 color => "black",
 align => "vcenter",
 border   => 2,
);
my $grand = $workbook->add_format(
 center_across => 1,
 bold => 1,
 color => "black",
 align => "vcenter",
 border   => 2,
);
my $format1 = $workbook->add_format(
 bold => 1,
 color => "black",
 border   => 2,
 bg_color => "green",
 font    =>'Lucida Calligraphy',
 size   => 10,
);
my $format3 = $workbook->add_format(
 bold => 1,
 color => "black",
 border   => 2,
 bg_color => "Yellow",
 align => "vcenter",
center_across => 1,
);
$worksheet->set_column(0,0,20);
#$worksheet->set_column('F1:H1',20);
$worksheet->merge_range('A1:B1', "         Hits-Count Report($date)                                                                                                                                                                                   ", $format1);
open(FH,"<$csvFilename") or die "Cannot open file: $!\n";
my ($x,$y) = (1,0);
while (<FH>){
#chomp -- get rid of the newline character
 chomp;
 # @list = split /\,/,$_;
 # read the fields in the current record into an array= "$_"
 @list = split ('\t',$_);

 foreach my $c (@list){
    if( $y == 6 )
    {
	$worksheet->set_column(0,0,85);
        $worksheet->write($x, $y++,"  ", $format0);
	$worksheet->set_column(0,0,20);
    }
   elsif ( $x == 1  )

    {
          $worksheet->set_column($x,$y,24);
          $worksheet->write($x, $y++, "$c",$format3);             
    }
else 
# 
{
        $worksheet->set_column($x,$y,20);
        $worksheet->write($x, $y++, "$c",$format);
	$worksheet->write( "A9", "   GrandTotal",$grand );
	$worksheet->write( "A11", "   Branch-Node",$format3 );
	$worksheet->write( "B11", "   Hits-Count",$format3 );
	$worksheet->write( "B9", "=SUM(B4+B5+B6+B7+B8)",$grand );
	$worksheet->write( "A17", "   GrandTotal",$grand );
	$worksheet->write( "B17", "=SUM(B12+B13+B14+B15+B16)",$grand );
}
 }
 $x++;$y=0;
}
close(FH);
$workbook->close();
