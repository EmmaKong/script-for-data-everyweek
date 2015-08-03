#!/usr/bin/perl 
#version : 1.0
#------------------------------------------------------------------------
#Target:                                                               
#   auto analysis tool of user feedback.    
#------------------------------------------------------------------------
use Cwd;
use Time::HiRes qw(gettimeofday);
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
$Win32::OLE::Warn = 3;  


#需要将待处理的文件名设为a.xlsx，同时将所有隐藏页删除
my @PHONEMODELS = qw(X520L Xplay3S X520F X3t X3L X3V X510t Xplay X710L Xshot X710F X5L Y22iL Y27 Y13L Y22L X5MaxL X5V Y28L Y23L X5S\sL X5Max\+ Y29L X5MaxV X5ProD X5M Y13iL Y33);  # Xplay3s: 0, X3t: 1, X510: 2, Xplay: 3, Xshot: 4
my @PAGENAME = qw(全部数据 筛选);
my $Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit');   
my $dir = getcwd;
my $X3t = "$dir/X3t.xlsx";
my $X3L = "$dir/X3L.xlsx";
my $X3V = "$dir/X3V.xlsx";
my $Xplay = "$dir/Xplay.xlsx";
my $Xplay3s = "$dir/Xplay3s.xlsx";
my $Xplay3sF = "$dir/Xplay3sF.xlsx";
my $Xshot = "$dir/Xshot.xlsx";
my $XshotF = "$dir/XshotF.xlsx";
my $X5L = "$dir/X5L.xlsx";
my $Y22iL = "$dir/Y22iL.xlsx";
my $Y27 = "$dir/Y27.xlsx";
my $Y13L = "$dir/Y13L.xlsx";
my $Y22L = "$dir/Y22L.xlsx";
my $X5MaxL = "$dir/X5MaxL.xlsx";
my $X5V = "$dir/X5V.xlsx";
my $Y28L = "$dir/Y28L.xlsx";
my $Y23L = "$dir/Y23L.xlsx";
my $X5S = "$dir/X5SL.xlsx";
my $X5MaxLL = "$dir/X5Max+.xlsx";
my $Y29L = "$dir/Y29L.xlsx";
my $X5MaxV = "$dir/X5MaxV.xlsx";
my $X5ProD = "$dir/X5ProD.xlsx";
my $X5M = "$dir/X5M.xlsx";
my $Y13iL = "$dir/Y13iL.xlsx";
my $Y33 = "$dir/Y33.xlsx";

main();

sub main{
	preProcess();	#预处理，将原始数据按机型对应	
}

sub preProcess{
	
	my ($start_sec, $start_usec) = gettimeofday();
	process($X3t, 1, $PHONEMODELS[3]);	
	my ($first_sec, $first_usec) = gettimeofday();
    my $timeDelta = ($first_usec - $start_usec) / 1000000 + ($first_sec - $start_sec);
    printf "X3t已耗时：%s秒\n", $timeDelta ;

	process($X3L, 1, $PHONEMODELS[4]);	
	my ($second_sec, $second_usec) = gettimeofday();
     $timeDelta = ($second_usec - $first_usec) / 1000000 + ($second_sec - $first_sec);
    printf "X3L已耗时：%s秒\n", $timeDelta ;
    
	process($X3V, 1, $PHONEMODELS[5]);	
	my ($third_sec, $third_usec) = gettimeofday();
     $timeDelta = ($third_usec - $second_usec) / 1000000 + ($third_sec - $second_sec);
    printf "X3V已耗时：%s秒\n", $timeDelta ;
      
	process($Xplay, 2, $PHONEMODELS[6], $PHONEMODELS[7]);	
	my ($fourth_sec, $fourth_usec) = gettimeofday();
    $timeDelta = ($fourth_usec - $third_usec) / 1000000 + ($fourth_sec - $third_sec);
	printf "Xplay已耗时：%s秒\n", $timeDelta ;
	
	process($Xshot, 2, $PHONEMODELS[8],$PHONEMODELS[9]);
	my ($fifth_sec, $fifth_usec) = gettimeofday();
    $timeDelta = ($fifth_usec - $fourth_usec) / 1000000 + ($fifth_sec - $fourth_sec);
	printf "Xshot已耗时：%s秒\n", $timeDelta ;
	
	process($XshotF, 1, $PHONEMODELS[10]);
	my ($sixth_sec, $sixth_usec) = gettimeofday();
    $timeDelta = ($sixth_usec - $fifth_usec) / 1000000 + ($sixth_sec - $fifth_sec);
	printf "XshotF已耗时：%s秒\n", $timeDelta ;

	process($X5L, 1, $PHONEMODELS[11]);
	my ($seven_sec, $seven_usec) = gettimeofday();
    $timeDelta = ($seven_usec - $sixth_usec) / 1000000 + ($seven_sec - $sixth_sec);
	printf "X5L已耗时：%s秒\n", $timeDelta ;

	process($Y22iL, 1, $PHONEMODELS[12]);
	my ($eighth_sec, $eighth_usec) = gettimeofday();
    $timeDelta = ($eighth_usec - $seven_usec) / 1000000 + ($eighth_sec - $seven_sec);
	printf "Y22iL已耗时：%s秒\n", $timeDelta ;
	
	process($Y27, 1, $PHONEMODELS[13]);
	my ($ninth_sec, $ninth_usec) = gettimeofday();
    $timeDelta = ($ninth_usec - $eighth_usec) / 1000000 + ($ninth_sec - $eighth_sec);
	printf "Y27已耗时：%s秒\n", $timeDelta ;

	process($Y13L, 1, $PHONEMODELS[14]);
	my ($tenth_sec, $tenth_usec) = gettimeofday();
    $timeDelta = ($tenth_usec - $ninth_usec) / 1000000 + ($tenth_sec - $ninth_sec);
	printf "Y13L已耗时：%s秒\n", $timeDelta ;
		
	process($Xplay3s, 2, $PHONEMODELS[0],$PHONEMODELS[1]);
	my ($eleven_sec, $eleven_usec) = gettimeofday();
    $timeDelta = ($eleven_usec - $tenth_usec) / 1000000 + ($eleven_sec - $tenth_sec);
	printf "Xplay3s已耗时：%s秒\n", $timeDelta ;
	
	process($Xplay3sF, 1, $PHONEMODELS[2]);
	my ($twelfth_sec, $twelfth_usec) = gettimeofday();
    $timeDelta = ($twelfth_usec - $eleven_usec) / 1000000 + ($twelfth_sec - $eleven_sec);
	printf "Xplay3sF已耗时：%s秒\n", $timeDelta ;

	process($Y22L, 1, $PHONEMODELS[15]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "Y22L已耗时：%s秒\n", $timeDelta ;

	process($X5MaxL, 1, $PHONEMODELS[16]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "X5MAXL已耗时：%s秒\n", $timeDelta ;

	process($X5V, 1, $PHONEMODELS[17]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "X5V已耗时：%s秒\n", $timeDelta ;	

	process($Y28L, 1, $PHONEMODELS[18]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "Y28L已耗时：%s秒\n", $timeDelta ;	
	
	process($Y23L, 1, $PHONEMODELS[19]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "Y23L已耗时：%s秒\n", $timeDelta ;	
	
	process($X5S, 1, $PHONEMODELS[20]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "X5S已耗时：%s秒\n", $timeDelta ;	
	
	process($X5MaxLL, 1, $PHONEMODELS[21]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "X5Max+已耗时：%s秒\n", $timeDelta ;	
	
	process($Y29L, 1, $PHONEMODELS[22]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "Y29L已耗时：%s秒\n", $timeDelta ;	
	
	process($X5MaxV, 1, $PHONEMODELS[23]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "Y29L已耗时：%s秒\n", $timeDelta ;	
	
	process($X5ProD, 1, $PHONEMODELS[24]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "X5ProD已耗时：%s秒\n", $timeDelta ;	
	
	process($X5M, 1, $PHONEMODELS[25]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "X5M已耗时：%s秒\n", $timeDelta ;	
	
	process($Y13iL, 1, $PHONEMODELS[26]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "Y13iL已耗时：%s秒\n", $timeDelta ;
	
	process($Y33, 1, $PHONEMODELS[27]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "Y33已耗时：%s秒\n", $timeDelta ;
	
}

sub process{
	my $filePath = shift;	
	my $number = shift;
	my $model = shift;
	my $model2;
	my $value = 1;
	
	#  modify kongqiao 20150725
	my $workbook = $Excel->Workbooks->Open($filePath); 
	my $allDataSheet = $workbook->Sheets(1);
  $allDataSheet->{name} = "全部数据";
	#$allDataSheet->Activate;	#删除前必须要激活当前窗口
	my $Rowcount=$allDataSheet->usedrange->rows->count;       #最大有效行数
	
	
	my $row = 2;	#从第二行开始遍历
	my $modelselect;
	
	#读出EXCEL数据到数组
	$totolRow=$Rowcount+1;
	$numDRow=X.$totolRow;
	$allDataArray = $allDataSheet->Range("A1:$numDRow")->{'Value'};
	$allDataLength = @$allDataArray;
	
	$firstSheet = $workbook->Worksheets->Add;
	$firstSheet->{name} = "所有数据";
	$firstSheet->Range("A1:$numDRow")->{'Value'} = $allDataArray;
	$firstSheet->Activate;	#删除前必须要激活当前窗口
	$firstSheet->Columns("A:A")->Delete;
	$firstSheet->Columns("D:F")->Delete;
	$firstSheet->Columns("F:G")->Delete;
	$DataArray = $firstSheet->Range("A1:$numDRow")->{'Value'};
	

	if ($number eq 1) {
		#printf "number = 1\n";
		for(2..$Rowcount){  	
		  $ref_array=$$DataArray[$row-1];
		  $value=$$ref_array[0];

			if($value =~ /$model$/i) {
				++$row;	#遇到了合适的，下一次就得加一行遍历
			} else {
				$position=$row-1;
				splice(@$DataArray,$position,1);#删除一行，整个表格会上移	
			}	
		}
	} elsif ($number eq 2) {
		$model2 = shift;		
		for(2..$Rowcount){  		
		  $ref_array=$$DataArray[$row-1];
		  $value=$$ref_array[0];
			if($value =~ /$model$/i || $value =~ /$model2$/i) {
				++$row;	#遇到了合适的，下一次就得加一行遍历
			} else {	
				$position=$row-1;
				splice(@$DataArray,$position,1);#删除一行，整个表格会上移	
			}	
		}
	}
	printf "%d find\n", ($row-2);
	
	#创建每页sheet，把数组写回EXCEL
	$pageNum = 2;
	$newSheet = $workbook->Worksheets->Add;
	$newSheet = $workbook->Worksheets->Add;
	$workbook->Sheets($pageNum)->{name} = "数据";
	$DataLenth = @$DataArray;
	$Dataend = X.$DataLenth;
	$workbook->Sheets($pageNum)->Range("A1:$Dataend")->{'value'} = $DataArray;	#传递的是数组引用
	
	#$workbook->Sheets($pageNum+1)->{name} = "所有数据";
		
	$workbook->Save();
	$workbook->Close();
}