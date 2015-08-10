#!/usr/bin/perl 
#version : 1.0
# xiaofangxu@vivo.com.cn, 2014.5.30
#------------------------------------------------------------------------
#Target:                                                               
#   auto analysis tool of user feedback.    
#------------------------------------------------------------------------
#  新建excel 问件，用于存储 模式细分 数据

# modify kongqiao

use Cwd;
use Time::HiRes qw(gettimeofday);
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
#use Win32::OLE::NLS qw(:LOCALE :TIME);   
$Win32::OLE::Warn = 3;  



my @PHONEMODELS = qw(X3t X3L X3V Xplay Xplay3s Xplay3sF Xshot XshotF X5L Y22iL Y27 Y13L Y22L X5MaxL X5V Y28L Y23L X5SL X5Max+ Y29L X5MaxV X5ProD X5M Y13iL Y33); 
#my @PHONEMODELS = qw(Xplay Xplay3s Xshot X5MaxL X5Max+ Y29L X5L X5MaxV X5ProD X5M);# X5M); 
my @PAGENAME = qw(全部数据 筛选);
my $Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit');   
my $dir = getcwd;

main();

sub main{
	preProcess();	#预处理，将原始数据按机型对应	
}

sub preProcess{
	
	foreach (0...24){
	process($PHONEMODELS[$_]);	
	
	}
	
}

sub process{
	
	 my $model = shift;
	
	
	# 上周数据路径
	 my @allfilePathLast = <*.xlsx>;
	 foreach $path (@allfilePathLast){
	 		if($path =~ /ROM2.0_$model_.*用户反馈/){
	  	 
	   		$filePathLast = $dir."\/".$path;
	 		}
	 }
 								  								
	 $workbookLast = $Excel->Workbooks->Open($filePathLast);
	
	 										
	 #读出上周 全部数据 EXCEL数据到数组			（未作任何处理的数据），有待之后处理						 
	 $SheetLast = $workbookLast->Sheets("全部数据");
	 my $Rowcount = $SheetLast->usedrange->rows->count;       #最大有效行数
	 my $numDRow = MM.$Rowcount;
	 $DataArrayLast = $SheetLast->Range("A1:$numDRow")->{'Value'};
									
	 #新建一个空的Excel文件，然后保存
	 my $workbook = $Excel->Workbooks->Add(); #新建一个工作簿
	 my $filenew = $dir."\/".$model."DetailLast".".xlsx";  
	 print "新建文件： $filenew \n";	 
	 my $newpath = $dir =~ s#/#\\#r;   # 将路径中的 反斜杠 替换成斜杠	 
	 $file = $newpath.'\\'.$model.'DetailLast.xlsx';
	# $workbook->SaveAs($model."DetailLast".".xlsx") or die "Save failer."; #保存这个工作部文件  # 默认就是存储在当前文件夹下！！！此处，第一个参数只用写文件名即可！！！ 
   $workbook->SaveAs($file) or die "Save failer.";
   
   my $allDataSheet = $workbook->Sheets(1);
   $allDataSheet->{name} = "全部数据";
   $allDataSheet->Range("A1:$numDRow")->{'value'} = $DataArrayLast; 
 
	# 删除不必要的 sheet
	$Sheet2 = $workbook->Sheets(2);
	$Sheet2->Activate;	# 删除前必须要激活当前窗口
	$Sheet2->Delete;
	
	$workbookLast->Save();
	$workbookLast->Close();
		
	$workbook->Save();
	$workbook->Close();
	
}