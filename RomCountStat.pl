use Cwd;
use Time::HiRes qw(gettimeofday);
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
$Win32::OLE::Warn = 3;

my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
  || Win32::OLE->new( 'Excel.Application', 'Quit' );

my %hash_model = (
	"PD1225"   => "Xplay",
	"PD1227F"  => "X3F",
	"PD1227L"  => "X3L",
	"PD1227T"  => "X3t",
	"PD1227V"  => "X3V",
	"PD1302"   => "XShot",
	"PD1302F"  => "X710F",
	"PD1303"   => "Xplay3S",
	"PD1303F"  => "X520F",
	"PD1304B"  => "Y613",
	"PD1304CF" => "Y613F",
	"PD1304CL" => "Y13L",
	"PD1304CV" => "Y913",
	"PD1304DL" => "Y13iL",
	"PD1309W"  => "Y622",
	"PD1309BL" => "Y22iL",
	"PD1309L"  => "Y22L",
	"PD1401BL" => "X5SL",
	"PD1401CL" => "vivoX5M",
	"PD1401F"  => "X5F",
	"PD1401L"  => "X5L",
	"PD1401V"  => "X5V",
	"PD1402"   => "Y18L",
	"PD1403F"  => "Y628",
	"PD1403L"  => "Y28L",
	"PD1403V"  => "Y928",
	"PD1408BL"  => "X5MAX+",
	"PD1408L"  => "X5MAXL",
	"PD1408V"  => "X5MAXV",
	"PD1410F"  => "Y627",
	"PD1410L"  => "Y27",
	"PD1410V"  => "Y927",
	"PD1419L"  => "Y23L",
	"PD1420L"  => "Y29L",
	"PD1421"  => "X5Pro",
	"PD1421D"  => "X5ProD",
	"PD1421L"  => "X5ProL",
	"PD1422L" => "Y33L"
);

my $dir = getcwd;
my $time = "从2010-01-01到2015-07-13详细升级信息表";

my $out_excel_file = "rom_count.xlsx";
main();

sub main {
	Process();
}

sub Process {

	my %hash_model_ver_max = ();
	
	while ( my ( $key_model, $value_model ) = each %hash_model ) {
		print $key_model. "=>" . $value_model . "\n";

		my %hash_model_ver = ();
#		$hash_model_ver_max{$value_model} = %hash_model_ver;

		#版本分布_PD1304CV_2015-04-23
		my $docName = $key_model . $time;
		my $filePath = $dir . "\/" . $docName . ".xls";    # 表格路径
		if ( !-e $filePath ) {
			print $filePath. "文件不存在，无法统计。\n";
			next;
		}

		my $workbook = $Excel->Workbooks->Open($filePath);
		my $sheet    = $workbook->Sheets("详细信息");
		my $rowcount = $sheet->usedrange->rows->count;
		my $numDRow  = "F" . $rowcount;
		my $dataArr  = $sheet->Range("A1:$numDRow")->{'Value'};

		for my $index ( 1 .. @$dataArr ) {
			my $data = $$dataArr[$index];
#			my @arr  = split( /_/, $$data[1] );
			my $ver  = $$data[1];
			if ( !exists( $hash_model_ver{$ver} ) ) {
				$hash_model_ver{$ver} = 0;
			}
			if ( $hash_model_ver{$ver} < $$data[2] ) {
				$hash_model_ver{$ver} = $$data[2];
			}
		}

		@keys_inorder = sort { $a <=> $b } keys %hash_model_ver;

		my $out_data = [];
		my $out_row_first = [ "Rom版本", "Rom最大用户数" ];
		push( @$out_data, $out_row_first );
		for ( my $index = 1 ; $index < @keys_inorder ; $index++ ) {
			my $out_row = [];
			push( @$out_row,  $keys_inorder[$index] );
			push( @$out_row,  $hash_model_ver{ $keys_inorder[$index] } );
			push( @$out_data, $out_row );
		}
		
		$hash_model_ver_max{$value_model} = $out_data;

		$workbook->Save();
		$workbook->Close();
	}
	
	my $out_book           = $Excel->Workbooks->Open($dir . "\/" . $out_excel_file );
#	my $sheetcount = $out_book->Worksheets->Count;
#	for ($index = $sheetcount; $index > 0; $index--){
#		print $out_book->Worksheets($index)->Delete;
#	}

	while ( my ( $key, $value ) = each %hash_model_ver_max ) {
		my $data = $value;
		my $data_count = @$data;
		my $out_sheet = $out_book->Worksheets->Add;
		$out_sheet->{name} = $key;
		$out_sheet->Range( "A1:" . "B" . $data_count )->{'Value'} = [@$data];
	}
	$out_book->Save;
	$out_book->Close;

}

