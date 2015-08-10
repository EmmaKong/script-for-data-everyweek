#!/usr/bin/perl 
#version : 1.0
# xiaofangxu@vivo.com.cn, 2014.5.30
#------------------------------------------------------------------------
#Target:                                                               
#   auto analysis tool of user feedback.    
#------------------------------------------------------------------------
#  �½�excel �ʼ������ڴ洢 ģʽϸ�� ����

# modify kongqiao

use Cwd;
use Time::HiRes qw(gettimeofday);
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
#use Win32::OLE::NLS qw(:LOCALE :TIME);   
$Win32::OLE::Warn = 3;  



my @PHONEMODELS = qw(X3t X3L X3V Xplay Xplay3s Xplay3sF Xshot XshotF X5L Y22iL Y27 Y13L Y22L X5MaxL X5V Y28L Y23L X5SL X5Max+ Y29L X5MaxV X5ProD X5M Y13iL Y33); 
#my @PHONEMODELS = qw(Xplay Xplay3s Xshot X5MaxL X5Max+ Y29L X5L X5MaxV X5ProD X5M);# X5M); 
my @PAGENAME = qw(ȫ������ ɸѡ);
my $Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit');   
my $dir = getcwd;

main();

sub main{
	preProcess();	#Ԥ������ԭʼ���ݰ����Ͷ�Ӧ	
}

sub preProcess{
	
	foreach (0...24){
	process($PHONEMODELS[$_]);	
	
	}
	
}

sub process{
	
	 my $model = shift;
	
	
	# ��������·��
	 my @allfilePathLast = <*.xlsx>;
	 foreach $path (@allfilePathLast){
	 		if($path =~ /ROM2.0_$model_.*�û�����/){
	  	 
	   		$filePathLast = $dir."\/".$path;
	 		}
	 }
 								  								
	 $workbookLast = $Excel->Workbooks->Open($filePathLast);
	
	 										
	 #�������� ȫ������ EXCEL���ݵ�����			��δ���κδ�������ݣ����д�֮����						 
	 $SheetLast = $workbookLast->Sheets("ȫ������");
	 my $Rowcount = $SheetLast->usedrange->rows->count;       #�����Ч����
	 my $numDRow = MM.$Rowcount;
	 $DataArrayLast = $SheetLast->Range("A1:$numDRow")->{'Value'};
									
	 #�½�һ���յ�Excel�ļ���Ȼ�󱣴�
	 my $workbook = $Excel->Workbooks->Add(); #�½�һ��������
	 my $filenew = $dir."\/".$model."DetailLast".".xlsx";  
	 print "�½��ļ��� $filenew \n";	 
	 my $newpath = $dir =~ s#/#\\#r;   # ��·���е� ��б�� �滻��б��	 
	 $file = $newpath.'\\'.$model.'DetailLast.xlsx';
	# $workbook->SaveAs($model."DetailLast".".xlsx") or die "Save failer."; #��������������ļ�  # Ĭ�Ͼ��Ǵ洢�ڵ�ǰ�ļ����£������˴�����һ������ֻ��д�ļ������ɣ����� 
   $workbook->SaveAs($file) or die "Save failer.";
   
   my $allDataSheet = $workbook->Sheets(1);
   $allDataSheet->{name} = "ȫ������";
   $allDataSheet->Range("A1:$numDRow")->{'value'} = $DataArrayLast; 
 
	# ɾ������Ҫ�� sheet
	$Sheet2 = $workbook->Sheets(2);
	$Sheet2->Activate;	# ɾ��ǰ����Ҫ���ǰ����
	$Sheet2->Delete;
	
	$workbookLast->Save();
	$workbookLast->Close();
		
	$workbook->Save();
	$workbook->Close();
	
}