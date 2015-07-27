#!/usr/bin/perl 
#version : 1.0
# xiaofangxu@vivo.com.cn, 2014.5.30
#------------------------------------------------------------------------
#Target:                                                               
#   auto analysis tool of user feedback.    
#------------------------------------------------------------------------
use Cwd;
use Time::HiRes qw(gettimeofday);
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
#use Excel::Writer::XLSX;
$Win32::OLE::Warn = 3;  
#���������
#X3L
#X3t
#X3V
#X510t Xplay
#X520F Xplay3SF
#X520L Xplay3S
#X5L 1401L
#X710F XshotF
#X710L Xshot
#Y22iL
#Y27

#��Ҫ����������ļ�����Ϊa.xlsx��ͬʱ����������ҳɾ��
my @PHONEMODELS = qw(X520L Xplay3S X520F X3t X3L X3V X510t Xplay X710L Xshot X710F X5L Y22iL Y27 Y13L Y22L X5MaxL X5V Y28L Y23L X5S\sL X5Max\+ Y29L X5MaxV X5ProD X5M Y13iL Y33);  # Xplay3s: 0, X3t: 1, X510: 2, Xplay: 3, Xshot: 4
my @PAGENAME = qw(ȫ������ ɸѡ);
my $Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit');   
my $dir = getcwd;


main();

sub main{
	preProcess();	#Ԥ������ԭʼ���ݰ����Ͷ�Ӧ	
}

sub preProcess{
	
	my ($start_sec, $start_usec) = gettimeofday();
	process($PHONEMODELS[3]);	
	my ($first_sec, $first_usec) = gettimeofday();
    my $timeDelta = ($first_usec - $start_usec) / 1000000 + ($first_sec - $start_sec);
    printf "X3t�Ѻ�ʱ��%s��\n", $timeDelta ;
=pod
	process($PHONEMODELS[4]);	
	my ($second_sec, $second_usec) = gettimeofday();
     $timeDelta = ($second_usec - $first_usec) / 1000000 + ($second_sec - $first_sec);
    printf "X3L�Ѻ�ʱ��%s��\n", $timeDelta ;
    
	process($PHONEMODELS[5]);	
	my ($third_sec, $third_usec) = gettimeofday();
     $timeDelta = ($third_usec - $second_usec) / 1000000 + ($third_sec - $second_sec);
    printf "X3V�Ѻ�ʱ��%s��\n", $timeDelta ;
      
	process($PHONEMODELS[6], $PHONEMODELS[7]);	
	my ($fourth_sec, $fourth_usec) = gettimeofday();
    $timeDelta = ($fourth_usec - $third_usec) / 1000000 + ($fourth_sec - $third_sec);
	printf "Xplay�Ѻ�ʱ��%s��\n", $timeDelta ;
	
	process($PHONEMODELS[8],$PHONEMODELS[9]);
	my ($fifth_sec, $fifth_usec) = gettimeofday();
    $timeDelta = ($fifth_usec - $fourth_usec) / 1000000 + ($fifth_sec - $fourth_sec);
	printf "Xshot�Ѻ�ʱ��%s��\n", $timeDelta ;
	
	process($PHONEMODELS[10]);
	my ($sixth_sec, $sixth_usec) = gettimeofday();
    $timeDelta = ($sixth_usec - $fifth_usec) / 1000000 + ($sixth_sec - $fifth_sec);
	printf "XshotF�Ѻ�ʱ��%s��\n", $timeDelta ;

	process($PHONEMODELS[11]);
	my ($seven_sec, $seven_usec) = gettimeofday();
    $timeDelta = ($seven_usec - $sixth_usec) / 1000000 + ($seven_sec - $sixth_sec);
	printf "X5L�Ѻ�ʱ��%s��\n", $timeDelta ;

	process($PHONEMODELS[12]);
	my ($eighth_sec, $eighth_usec) = gettimeofday();
    $timeDelta = ($eighth_usec - $seven_usec) / 1000000 + ($eighth_sec - $seven_sec);
	printf "Y22iL�Ѻ�ʱ��%s��\n", $timeDelta ;
	
	process($PHONEMODELS[13]);
	my ($ninth_sec, $ninth_usec) = gettimeofday();
    $timeDelta = ($ninth_usec - $eighth_usec) / 1000000 + ($ninth_sec - $eighth_sec);
	printf "Y27�Ѻ�ʱ��%s��\n", $timeDelta ;

	process($PHONEMODELS[14]);
	my ($tenth_sec, $tenth_usec) = gettimeofday();
    $timeDelta = ($tenth_usec - $ninth_usec) / 1000000 + ($tenth_sec - $ninth_sec);
	printf "Y13L�Ѻ�ʱ��%s��\n", $timeDelta ;
		
	process($PHONEMODELS[0],$PHONEMODELS[1]);
	my ($eleven_sec, $eleven_usec) = gettimeofday();
    $timeDelta = ($eleven_usec - $tenth_usec) / 1000000 + ($eleven_sec - $tenth_sec);
	printf "Xplay3s�Ѻ�ʱ��%s��\n", $timeDelta ;
	
	process($PHONEMODELS[2]);
	my ($twelfth_sec, $twelfth_usec) = gettimeofday();
    $timeDelta = ($twelfth_usec - $eleven_usec) / 1000000 + ($twelfth_sec - $eleven_sec);
	printf "Xplay3sF�Ѻ�ʱ��%s��\n", $timeDelta ;

	process($PHONEMODELS[15]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "Y22L�Ѻ�ʱ��%s��\n", $timeDelta ;

	process($PHONEMODELS[16]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "X5MAXL�Ѻ�ʱ��%s��\n", $timeDelta ;

	process($PHONEMODELS[17]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "X5V�Ѻ�ʱ��%s��\n", $timeDelta ;	

	process($PHONEMODELS[18]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "Y28L�Ѻ�ʱ��%s��\n", $timeDelta ;	
	
	process($PHONEMODELS[19]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "Y23L�Ѻ�ʱ��%s��\n", $timeDelta ;	
	
	process($PHONEMODELS[20]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "X5S�Ѻ�ʱ��%s��\n", $timeDelta ;	
	
	process($PHONEMODELS[21]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "X5Max+�Ѻ�ʱ��%s��\n", $timeDelta ;	
	
	process($PHONEMODELS[22]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "Y29L�Ѻ�ʱ��%s��\n", $timeDelta ;	
	
	process($PHONEMODELS[23]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "Y29L�Ѻ�ʱ��%s��\n", $timeDelta ;	
	
	process($PHONEMODELS[24]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "X5ProD�Ѻ�ʱ��%s��\n", $timeDelta ;	
	
	process($PHONEMODELS[25]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "X5M�Ѻ�ʱ��%s��\n", $timeDelta ;	
	
	process($PHONEMODELS[26]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "Y13iL�Ѻ�ʱ��%s��\n", $timeDelta ;
	
	process($PHONEMODELS[27]);
	my ($thirteenth_sec, $thirteenth_usec) = gettimeofday();
    $timeDelta = ($thirteenth_usec - $twelfth_usec) / 1000000 + ($thirteenth_sec - $twelfth_sec);
	printf "Y33�Ѻ�ʱ��%s��\n", $timeDelta ;
=cut	
}

sub process{
	
	my $model = shift;
	my $model2;
	my $value = 1;
	
	
	# ��������·��
	 my @allfilePathLast = <*.xlsx>;
	 foreach $path (@allfilePathLast){
	 		if($path =~ /ROM2.0_$model_.*�û�����/){
	  	 
	   		$filePathLast = $dir."\/".$path;
	   		print "��������·���ǣ� $filePathLast \n";
	 		}
	 }
 								  								
	 $workbookLast = $Excel->Workbooks->Open($filePathLast);
	
	 										
	 #�������� ȫ������ EXCEL���ݵ�����			��δ���κδ�������ݣ�						 
	 $SheetLast = $workbookLast->Sheets("ȫ������");
	 my $Rowcount = $SheetLast->usedrange->rows->count;       #�����Ч����
	 my $numDRow=MM.$Rowcount;
	 $DataArrayLast = $SheetLast->Range("A1:$numDRow")->{'Value'};
									
=pod
   # �½�һ��excel 
   my $filePath = $dir."\/".$model."Last".".xlsx";
   use File::Copy;
   copy($filePathLast, $filePath) or die "Copy failed: $!";
   
   
   $workbook = $Excel->Workbooks->Open($filePath);
   
   $DataArray = qw();
   # ������ ȫ������ sheet
   
   $workbook->Sheets(1)->Activate;
   $workbook->Sheets(1)->Delete;	  # ɾ��ǰ����Ҫ���ǰ����
   
   $sheet = $workbook->Sheets(2);
   $rowCount = $sheet->usedrange->rows->count;       #�����Ч����   
   $numRow = MM.$rowCount;
   $sheet->Range("A1:$numRow")->{'value'} = $DataArray;  # �����ÿ�
   $sheet->Activate;
   $sheet->Delete;
   
   $sheet = $workbook->Sheets(3);
   $rowCount = $sheet->usedrange->rows->count;       #�����Ч����   
   $numRow = MM.$rowCount;
   $sheet->Range("A1:$numRow")->{'value'} = $DataArray;  
   $sheet->Activate;
   $sheet->Delete;
   
   $sheet = $workbook->Sheets(4);
   $rowCount = $sheet->usedrange->rows->count;       #�����Ч����   
   $numRow = MM.$rowCount;
   $sheet->Range("A1:$numRow")->{'value'} = $DataArray;  
   $sheet->Activate;
   $sheet->Delete;
   
   
   $sheet = $workbook->Sheets(5);
   $rowCount = $sheet->usedrange->rows->count;       #�����Ч����   
   $numRow = MM.$rowCount;
   $sheet->Range("A1:$numRow")->{'value'} = $DataArray;  
   $sheet->Activate;
   $sheet->Delete;
   
   $sheet = $workbook->Sheets(6);
   $rowCount = $sheet->usedrange->rows->count;       #�����Ч����   
   $numRow = MM.$rowCount;
   $sheet->Range("A1:$numRow")->{'value'} = $DataArray;  
   $sheet->Activate;
   $sheet->Delete;
   
   $sheet = $workbook->Sheets(7);
   $rowCount = $sheet->usedrange->rows->count;       #�����Ч����   
   $numRow = MM.$rowCount;
   $sheet->Range("A1:$numRow")->{'value'} = $DataArray;  
   $sheet->Activate;
   $sheet->Delete;
   
   $sheet = $workbook->Sheets(8);
   $rowCount = $sheet->usedrange->rows->count;       #�����Ч����   
   $numRow = MM.$rowCount;
   $sheet->Range("A1:$numRow")->{'value'} = $DataArray;  
   $sheet->Activate;
   $sheet->Delete;
   
   $sheet = $workbook->Sheets(9);
   $rowCount = $sheet->usedrange->rows->count;       #�����Ч����   
   $numRow = MM.$rowCount;
   $sheet->Range("A1:$numRow")->{'value'} = $DataArray;  
   $sheet->Activate;
   $sheet->Delete;
   
   $sheet = $workbook->Sheets(10);
   $rowCount = $sheet->usedrange->rows->count;       #�����Ч����   
   $numRow = MM.$rowCount;
   $sheet->Range("A1:$numRow")->{'value'} = $DataArray;  
   $sheet->Activate;
   $sheet->Delete;
   
   $sheet = $workbook->Sheets(11);
   $rowCount = $sheet->usedrange->rows->count;       #�����Ч����   
   $numRow = MM.$rowCount;
   $sheet->Range("A1:$numRow")->{'value'} = $DataArray;  
   $sheet->Activate;
   $sheet->Delete;
   
   $sheet = $workbook->Sheets(12);
   $rowCount = $sheet->usedrange->rows->count;       #�����Ч����   
   $numRow = MM.$rowCount;
   $sheet->Range("A1:$numRow")->{'value'} = $DataArray;  
   $sheet->Activate;
   $sheet->Delete;
   
   $sheet = $workbook->Sheets(13);
   $rowCount = $sheet->usedrange->rows->count;       #�����Ч����   
   $numRow = MM.$rowCount;
   $sheet->Range("A1:$numRow")->{'value'} = $DataArray;  
   $sheet->Activate;
   $sheet->Delete;
   
   $sheet = $workbook->Sheets(14);
   $rowCount = $sheet->usedrange->rows->count;       #�����Ч����   
   $numRow = MM.$rowCount;
   $sheet->Range("A1:$numRow")->{'value'} = $DataArray;  
   $sheet->Activate;
   $sheet->Delete;
   
   $sheet = $workbook->Sheets(15);
   $rowCount = $sheet->usedrange->rows->count;       #�����Ч����   
   $numRow = MM.$rowCount;
   $sheet->Range("A1:$numRow")->{'value'} = $DataArray;  
   $sheet->Activate;
   $sheet->Delete;
   
   $sheet = $workbook->Sheets(16);
   $rowCount = $sheet->usedrange->rows->count;       #�����Ч����   
   $numRow = MM.$rowCount;
   $sheet->Range("A1:$numRow")->{'value'} = $DataArray;  
   $sheet->Activate;
   $sheet->Delete;
  
   $sheet = $workbook->Sheets(17);
   $rowCount = $sheet->usedrange->rows->count;       #�����Ч����   
   $numRow = MM.$rowCount;
   $sheet->Range("A1:$numRow")->{'value'} = $DataArray;  
   $sheet->Activate;
   $sheet->Delete;
=cut


		#�½�һ���յ�Excel�ļ���Ȼ�󱣴�
		my $book = $Excel->Workbooks->Add(); #�½�һ��������
		$book->SaveAs( $dir."\/".$model."Last".".xlsx") or die "Save failer."; #��������������ļ�


   
=pod   
   $firstSheet->Range("A1:$numDRow")->{'value'} = $DataArrayLast;
	 # my $firstSheet = $workbook->Sheets(1);
	 $firstSheet->Activate;	# ɾ��ǰ����Ҫ���ǰ����
	 my $Rowcount = $firstSheet->usedrange->rows->count;       #�����Ч����
	 $firstSheet->Columns("A:A")->Delete;
	 $firstSheet->Columns("D:F")->Delete;
	 $firstSheet->Columns("F:G")->Delete;
	 my $row = 2;	#�ӵڶ��п�ʼ����
	 my $modelselect;
=cut	
		
	$workbookLast->Save();
	$workbookLast->Close();
		
	$workbook->Save();
	$workbook->Close();
}