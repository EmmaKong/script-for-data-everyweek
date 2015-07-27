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
$Win32::OLE::Warn = 3;  

$Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit');   

#my @PHONEMODELS = qw(X3t X3V Xplay Xplay3s X5L Xshot Y22iL Y27 Y13L Y22L X5V Y28L Y23L X5SL);# X5MaxL);  # 机型
#my @PHONEMODELS = qw(Xplay Xplay3s Xshot X5MaxL);  # 机型
#my @PHONEMODELS = qw(Xplay Xplay3s Xshot X5MaxL X5Max+ Y29L X5L X5MaxV X5ProD);#  X5M);  # 机型
my @PHONEMODELS = qw(X5ProD);  # 机型
my $dir = getcwd;
my $workbook;
my $pageNum = 2;	

@TAGSERIES = qw(耗电 发热 系统性能 死机 重启 自动关机 触屏 桌面 黑屏 升级 vivo乐园 屏幕发黄 闪屏 场景桌面 小挂件 相机 通话 无响应 停止运行 闪退 兼容性 
i音乐 i视频 音频 i管家 浏览器 软件商店 状态栏 联系人 短信 电子邮件 文件管理 输入法 天气预报 便签 计算器 超清影院 电子书 
网络信号 WIFI GPS 蓝牙 指纹 OTG 智能体感 流量监控 时间(不|无法)准 帐号 小屏 游戏 第三方软件 云服务 充电 截屏 闹钟 photo+ i主题 HIFI 锁屏 相册 手电筒 收音机 字体 SD卡 屏幕亮度 震动 vivo语音助手 
日历 NFC ROOT 游戏中心 访客模式 USB vivo社区
眼球识别 
天籁K歌 录音 掉漆 建议 问题 抱怨 满意);


$totolRow = 600;#初始化该sheet只需要600行数据  # modify kongqiao
$numDRow = MM.$totolRow;#400行对应的excel长宽
$ModuleNum = 5;
$colMax = 85;  # 模式数 
$endRow = 27;
$startRow = 11;
$rom_count_name = "rom_count.xlsx";
main();

sub main{
	Process();	#过滤处理，处理完后进行版本筛选，筛选策略：先找到版本个数，按照个数拷贝相同的份数，进行删除操作。	
}


sub Process{
		foreach (@PHONEMODELS) {
										print $_."\n";
										$moduleTempNum=0;
										#打开x3t
										$filePath = $dir."\/".$_.".xlsx";	# 表格路径
										$workbook = $Excel->Workbooks->Open($filePath);
										$workSheet=$workbook->Sheets("统计");
										#$selectSheet = $workbook->Sheets($pageNum);
										$PhoneName=$_;
										#读出x3tEXCEL数据到数组
										$DataArray = $workSheet->Range("A1:$numDRow")->{'Value'};
										$DataLength=@$DataArray-1;

#=pod
									 									  
									  #$filePathLast = $dir."\/".$_."用户反馈".".xlsx";	# 表格路径
									  # modify kongqiao 20150714
									  my @allfilePathLast = <*.xlsx>;
	  								foreach $path (@allfilePathLast){
	  									if($path =~ /ROM2.0_$__.*用户反馈/){
	  										$filePathLast = $dir."\/".$path;
	  										print "上周数据路径是：$filePathLast \n";
	  									}
	 									}
										$workbookLast = $Excel->Workbooks->Open($filePathLast);
										$workSheetLast=$workbookLast->Sheets("统计");
										#读出EXCEL数据到数组
										$DataArrayLast = $workSheetLast->Range("A1:$numDRow")->{'Value'};
										$DataLength=@$DataArray-1;	
#=cut		

#处理数据

#=pod
										%rom_count = getRomcount($_);
										calculateRomRate($DataArray, \%rom_count);
#=cut				
				
										$firstCol=$$DataArray[0];
										$tenthCol=$$DataArray[9];
										$fourteenCol=$$DataArray[13];
										$seventeenCol=$$DataArray[16];
										$twentyfourthCol=$$DataArray[23];
										$sixthCol=$$DataArray[5];
										$seventhCol=$$DataArray[6];
										$eightthCol=$$DataArray[7];
										$eleventhCol=$$DataArray[10];
										$twelvethCol=$$DataArray[11];
										$thirteennCol=$$DataArray[12];
										
										$twentyFiveCol=$$DataArray[24];
										$twentySixCol=$$DataArray[25];
										$twentySevenCol=$$DataArray[26];
										
										$eighteenthCol=$$DataArray[17];
										$nineteenthCol=$$DataArray[18];
										$twentyCol=$$DataArray[19];
										
										$senToFour[0]=$$DataArray[1];
										$senToFour[1]=$$DataArray[2];
										$senToFour[2]=$$DataArray[3];
										
										foreach $count(1..$colMax){

													$$tenthCol[$count-1]=$$firstCol[$count-1];
													$$fourteenCol[$count-1]=$$firstCol[$count-1];
													$twentyeighthCol[$count-1]=$$firstCol[$count-1];
													if($count < 2){
															$$seventeenCol[$count-1]=$$firstCol[$count-1];
															$$seventeenCol[$count-1+9]=$$seventeenCol[$count-1];
															$$seventeenCol[1]="总投诉量";
															$$seventeenCol[2]="性能投诉量";
															$$seventeenCol[10]="耗电";
															$$seventeenCol[11]="发热";
															
															$$twentyfourthCol[$count-1]=$$firstCol[$count-1];
															$$twentyfourthCol[$count-1+9]=$$twentyfourthCol[$count-1];	
															$$twentyfourthCol[1]="性能投诉占比";
															$$twentyfourthCol[2]="耗电投诉占比";
															$$twentyfourthCol[3]="发热投诉占比";
															$$twentyfourthCol[10]="ROM人数";						
													}
													foreach $num(0..2){
														if(${$senToFour[$num]}[$colMax-1] eq ""){
																	${$senToFour[$num]}[$colMax-1]=0;
														}
														if($count < $colMax-1){
																	${$senToFour[$num]}[$colMax-1]+=${$senToFour[$num]}[$count];
														}
													}
													
										}
										$$DataArray[1]=$senToFour[0];
										$$DataArray[2]=$senToFour[1];
										$$DataArray[3]=$senToFour[2];
										$$DataArray[9]=$tenthCol;
										$$DataArray[13]=$fourteenCol;
										$$DataArray[16]=$seventeenCol;
										$$DataArray[23]=$twentyfourthCol;
										$tempCount=0;
#$endRow=27;
#$startRow=11;
										foreach $count(2..$colMax-1){
													
													$$sixthCol[$count-1]=${$senToFour[0]}[$colMax-1];
													$$seventhCol[$count-1]=${$senToFour[1]}[$colMax-1];
													$$eightthCol[$count-1]=${$senToFour[2]}[$colMax-1];
													if($count eq 2){
																$$eighteenthCol[$count-1]=${$senToFour[0]}[$colMax-1];
																$$nineteenthCol[$count-1]=${$senToFour[1]}[$colMax-1];
																$$twentyCol[$count-1]=${$senToFour[2]}[$colMax-1];	
																$$eighteenthCol[$count]=${$senToFour[0]}[4];
																$$nineteenthCol[$count]=${$senToFour[1]}[4];
																$$twentyCol[$count]=${$senToFour[2]}[4];	
																$$eighteenthCol[$count+8]=${$senToFour[0]}[1];
																$$nineteenthCol[$count+8]=${$senToFour[1]}[1];
																$$twentyCol[$count+8]=${$senToFour[2]}[1];	
																$$eighteenthCol[$count+9]=${$senToFour[0]}[3];
																$$nineteenthCol[$count+9]=${$senToFour[1]}[3];
																$$twentyCol[$count+9]=${$senToFour[2]}[3];	
																
																
																
																if($$eighteenthCol[$count-1] eq 0){
																		$$twentyFiveCol[$count-1]=0;
																}else{
																		$$twentyFiveCol[$count-1]=$$eighteenthCol[$count]/$$eighteenthCol[$count-1]*100;
																		$$twentyFiveCol[$count-1] = $$twentyFiveCol[$count-1]."\%";
																}

																if($$nineteenthCol[$count-1] eq 0){
																		$$twentySixCol[$count-1]=0;
																}else{
																		$$twentySixCol[$count-1]=$$nineteenthCol[$count]/$$nineteenthCol[$count-1]*100;
																		$$twentySixCol[$count-1] = $$twentySixCol[$count-1]."\%";
																}
																
																if($$twentyCol[$count-1] eq 0){
																		$$twentySevenCol[$count-1]=0;
																}else{
																		$$twentySevenCol[$count-1]=$$twentyCol[$count]/$$twentyCol[$count-1]*100;
																		$$twentySevenCol[$count-1] = $$twentySevenCol[$count-1]."\%";
																}
																
																if($$eighteenthCol[$count-1] eq 0){
																		$$twentyFiveCol[$count]=0;
																}else{
																		$$twentyFiveCol[$count]=$$eighteenthCol[$count+8]/$$eighteenthCol[$count-1]*100;
																	$$twentyFiveCol[$count] = $$twentyFiveCol[$count]."\%";
																}
																
																if($$nineteenthCol[$count-1] eq 0){
																		$$twentySixCol[$count]=0;
																}else{
																		$$twentySixCol[$count]=$$nineteenthCol[$count+8]/$$nineteenthCol[$count-1]*100;
																		$$twentySixCol[$count] = $$twentySixCol[$count]."\%";
																}
																
																if($$twentyCol[$count-1] eq 0){
																		$$twentySevenCol[$count]=0;
																}else{
																		$$twentySevenCol[$count]=$$twentyCol[$count+8]/$$twentyCol[$count-1]*100;
																		$$twentySevenCol[$count] = $$twentySevenCol[$count]."\%";
																}
																
																if($$eighteenthCol[$count-1] eq 0){
																		$$twentyFiveCol[$count+1]=0;
																}else{
																		$$twentyFiveCol[$count+1]=$$eighteenthCol[$count+9]/$$eighteenthCol[$count-1]*100;
																		$$twentyFiveCol[$count+1] = $$twentyFiveCol[$count+1]."\%";
																}
																
																if($$nineteenthCol[$count-1] eq 0){
																		$$twentySixCol[$count+1]=0;
																}else{
																		$$twentySixCol[$count+1]=$$nineteenthCol[$count+9]/$$nineteenthCol[$count-1]*100;
																		$$twentySixCol[$count+1] = $$twentySixCol[$count+1]."\%";
																}
																
																
																if($$twentyCol[$count-1] eq 0){
																		$$twentySevenCol[$count+1]=0;
																}else{
																		$$twentySevenCol[$count+1]=$$twentyCol[$count+9]/$$twentyCol[$count-1]*100;
																		$$twentySevenCol[$count+1] = $$twentySevenCol[$count+1]."\%";
																}
																
													}
													
													if($$sixthCol[$count-1] eq 0){
															$$eleventhCol[$count-1]=0;
													}else{
															$$eleventhCol[$count-1]=${$senToFour[0]}[$count-1]/$$sixthCol[$count-1]*100;
															$$eleventhCol[$count-1] = $$eleventhCol[$count-1]."\%";
													}
													
													if($$seventhCol[$count-1] eq 0){
															$$twelvethCol[$count-1]=0;
													}else{
															$$twelvethCol[$count-1]=${$senToFour[1]}[$count-1]/$$seventhCol[$count-1]*100;
															$$twelvethCol[$count-1] = $$twelvethCol[$count-1]."\%";
													}
													
													if($$eightthCol[$count-1] eq 0){
															$$thirteennCol[$count-1]=0;
													}else{
															$$thirteennCol[$count-1]=${$senToFour[2]}[$count-1]/$$eightthCol[$count-1]*100;
															$$thirteennCol[$count-1] = $$thirteennCol[$count-1]."\%";
													}
													
										}
										
											my $row1 = $$DataArray[0];
											my $row2 = $$DataArray[1];
											my $row3 = $$DataArray[2];
											my $row4 = $$DataArray[3];
	
											my $romcount_1 = $rom_count{$$row2[0]};
											my $romcount_2 = $rom_count{$$row3[0]};
											my $romcount_3 = $rom_count{$$row4[0]};
										
										$$twentyFiveCol[10]= $romcount_1;
										$$twentySixCol[10]= $romcount_2;
										$$twentySevenCol[10]= $romcount_3;
																
										
										$$DataArray[5]=$sixthCol;
										$$DataArray[6]=$seventhCol;
										$$DataArray[7]=$eightthCol;
										$$DataArray[10]=$eleventhCol;
										$$DataArray[11]=$twelvethCol;
										$$DataArray[12]=$thirteennCol;
										$$DataArray[24]=$twentyFiveCol;
										$$DataArray[25]=$twentySixCol;
										$$DataArray[26]=$twentySevenCol;
										
										foreach $count($startRow..$endRow){
													if($count < 14){
															$tempArray=$$DataArray[$count-1];
															$$tempArray[0]=${$senToFour[$tempCount++]}[0];
															#print "tempArray[0]是$$tempArray[0]\n";
															$$DataArray[$count-1]=$tempArray;
													}
													if($count eq 15){
													
															$tempCount=0;
													}
													if($count > 17 && $count < 21){
															$tempArray=$$DataArray[$count-1];
															$$tempArray[0]=${$senToFour[$tempCount]}[0];
															$$tempArray[9]=${$senToFour[$tempCount]}[0];
															$$DataArray[$count-1]=$tempArray;
															$tempCount++;
													}
													if($count eq 21){
													
															$tempCount=0;
													}
													if($count > 24 ){
															$tempArray=$$DataArray[$count-1];
															$$tempArray[0]=${$senToFour[$tempCount]}[0];
															$$tempArray[9]=${$senToFour[$tempCount]}[0];
															$$DataArray[$count-1]=$tempArray;
															$tempCount++;
													}
													
													
										}



										#$workSheet->Range("A1:$numDRow")->{'value'}=$DataArray;


										#按从大到小冒泡排序
										#$DataArray = $workSheet->Range("A1:$numDRow")->{'Value'};
										foreach $tempCount(2..$colMax-1-5){
													$ten[$tempCount-2]=$$tenthCol[$tempCount-1];
													$third[$tempCount-2]=$$thirteennCol[$tempCount-1];
													
										}
										#print "ten[$colMax-1-5-2]是$ten[$colMax-1-5-2]\n";
										$appArrayLength=$colMax-1-5-2+1;
										
										#print "appArrayLength长度为：$appArrayLength\n";
										for($i=0;$i<$appArrayLength-1;$i++){
											for($j=0;$j<$appArrayLength-$i-1;$j++){
													$jArray=$third[$j];
													$jPlusArray=$third[$j+1];
													
													$tenArray=$ten[$j];
													$tenPlusArray=$ten[$j+1];
													
													if($jArray eq ""){
														$jArray=0;
													}
													if($jPlusArray eq ""){
														$jPlusArray=0;
													}	
													if($jArray < $jPlusArray){
														$third[$j+1]=$jArray;
														$third[$j]=$jPlusArray;
														
														$ten[$j+1]=$tenArray;
														$ten[$j]=$tenPlusArray;	
													}
													
											}
										}

#=pod
										#开始分析本周数据是否有突变
										$thirdLast=$$DataArrayLast[12];	
										$fourLast=$$DataArrayLast[3];	
										$thirteennCol=$$DataArray[12];
										$twelvethCol=$$DataArray[11];
										$eleventhCol=$$DataArray[10];
										$fourCol=$$DataArray[3];

										$tenthCol=$$DataArray[9];
										$temp=0;
										#print "$$tenthCol[$colMax-1-5-1]是$$tenthCol[$colMax-1-5-1]\n";

#与相同版本比 要数量上大于3
#与不同版本比较，数量上不做要求
#
										foreach $tempCount(2..$colMax-1-5){
													if($$thirteennCol[$tempCount-1] > 0.01 && $$thirteennCol[$tempCount-1] ne 0 && $$twelvethCol[$tempCount-1] ne 0 && $$eleventhCol[$tempCount-1] ne 0  && $$thirteennCol[$tempCount-1] >= ($$twelvethCol[$tempCount-1]+$$twelvethCol[$tempCount-1]/2) && $$fourCol[$tempCount-1] > $$fourLast[$tempCount-1]+3){
													#if($$thirteennCol[$tempCount-1] > 0.01 && ($$thirteennCol[$tempCount-1] >= ($$thirdLast[$tempCount-1]+$$thirdLast[$tempCount-1]/2) || $$thirteennCol[$tempCount-1] >= ($$twelvethCol[$tempCount-1]+$$twelvethCol[$tempCount-1]/2)) && $$fourCol[$tempCount-1] > $$fourLast[$tempCount-1]+3){
													#if($$thirteennCol[$tempCount-1] > 0.01 && $$thirteennCol[$tempCount-1] ne 0  && $$thirteennCol[$tempCount-1] >= ($$twelvethCol[$tempCount-1]+$$twelvethCol[$tempCount-1]/2) ){#&& $$fourCol[$tempCount-1] > $$fourLast[$tempCount-1]+3){
																$suddenChange[$temp++]=$$tenthCol[$tempCount-1];
													}
										}	
										$tempCount=0;
										foreach $tempNum(134..134+$temp){
													$tempArray=$$DataArray[$tempNum-1];
													if($tempNum eq 134){
															
															$$tempArray[12]="本周突变模块是:";
													}else{
															
															$$tempArray[13]=$suddenChange[$tempCount];
															$tempCount++;
													}
													
													$$DataArray[$tempNum-1]=$tempArray;
										}
#=cut			
										#判断排名前八位
										$tempCount=0;
										foreach $temp(134..142){
													$tempArray=$$DataArray[$temp-1];
													if($temp eq 134){
															$$tempArray[7]="排名前八位是:";

													}else{
															$$tempArray[8]=$ten[$tempCount];
															
															$tempCount++;
													}
													
													$$DataArray[$temp-1]=$tempArray;
										}
										$workSheet->Range("A1:$numDRow")->{'value'}=$DataArray;
											

											
											
										#开始制图
										foreach $count(1..7){				
													if($count eq 1){	
															$rangeStart=A17;
															$rangeEnd=C20;
															$chartX=0;
															$chartY=431;#29个小方格
															$chartWidth=360;
															$chartHeight=225;
															$chartName=$PhoneName."总投诉量和性能投诉总量";
															createChart($workSheet,$rangeStart,$rangeEnd,$chartX,$chartY,$chartWidth,$chartHeight,$chartName);
														
															
													}elsif($count eq 2){
															$rangeStart=A24;
															$rangeEnd=B27;
															$chartX=430;
															$chartY=431;
															$chartWidth=360;
															$chartHeight=225;
															$chartName=$PhoneName."性能投诉占比";
															createChart($workSheet,$rangeStart,$rangeEnd,$chartX,$chartY,$chartWidth,$chartHeight,$chartName);
													}elsif($count eq 3){
															$rangeStart=A24;
															$rangeEnd=D27;
															$chartX=0;
															$chartY=674;
															$chartWidth=360;
															$chartHeight=225;
															$chartName=$PhoneName."各投诉占比";
															createChart($workSheet,$rangeStart,$rangeEnd,$chartX,$chartY,$chartWidth,$chartHeight,$chartName);
													}elsif($count eq 4){
															$rangeStart=A1;
															$rangeEnd=CA4;
															$chartX=0;
															$chartY=915;
															$chartWidth=1400;
															$chartHeight=281;
															$chartName=$PhoneName."的ROM版本总投诉量";
															createChart($workSheet,$rangeStart,$rangeEnd,$chartX,$chartY,$chartWidth,$chartHeight,$chartName);
													}elsif($count eq 5){
															$rangeStart=A10;
															$rangeEnd=CA13;
															$chartX=0;
															$chartY=1215;
															$chartWidth=1400;
															$chartHeight=281;
															$chartName=$PhoneName."的ROM版本总投诉占比";
															createChart($workSheet,$rangeStart,$rangeEnd,$chartX,$chartY,$chartWidth,$chartHeight,$chartName);
													}elsif($count eq 6){
															$value=$workSheet->Range("A15:A15")->{'Value'};
															print "value是$value\n";
															if($value ne ""){
															print "value是$value\n";
															$rangeStart=A14;
															$rangeEnd=CA15;
															$chartX=0;
															$chartY=1700;
															$chartWidth=1400;
															$chartHeight=281;
															$chartName=$PhoneName."的".$value."版本";
															createChart($workSheet,$rangeStart,$rangeEnd,$chartX,$chartY,$chartWidth,$chartHeight,$chartName);
															}
													}elsif($count eq 7){
															$rangeStart=A28;
															$rangeEnd=CA31;
															$chartX=0;
															$chartY=1500;
															$chartWidth=1400;
															$chartHeight=281;
															$chartName=$PhoneName."的ROM版本实销投诉占比";
															createChart($workSheet,$rangeStart,$rangeEnd,$chartX,$chartY,$chartWidth,$chartHeight,$chartName);
															
													}
													
										}
															

														
									
										$workbook->Save();
										$workbook->Close();
#=pod										
										$workbookLast->Save();
										$workbookLast->Close();
#=cut					
										#$workbookDetail->Save();
										#$workbookDetail->Close();					
					
			
					
					
					}

}


#获取数字转换成大写字母
sub getColumnName{  
    $first;  
    $lastvalue;  
    $columnNum=$_[0];
    #print "column的值为".$columnNum."\n";
    $result = "";  

    $first = int($columnNum / 27);  
    $lastvalue = $columnNum - ($first * 26);  
  	
  	#print "看一下".$first."            ".$lastvalue."\n"; 
  	$temp;
    if ($first gt 0){ #lt小于 gt大于 ne不等于 
    		$temp=$first+64;
    		#print $temp."\n";
        $result = chr($temp);  
      }
  
    if (lastvalue gt 0){
    		$temp=$lastvalue+64;
    		#print $temp."\n";
        $result = $result.chr($temp); 
      } 
 			#$result=$result;
      #print "暂时值为：".$result."\n"; 
  
} 



sub getRomcount{
	
	my $model = shift;
	#$mmodel=$model;#wuhongzhang 
	my $romfilepath = $dir."\/".$rom_count_name;
	my $workbook = $Excel->Workbooks->Open($romfilepath);
	my $sheet    = $workbook->Sheets($model);
	my $rowcount = $sheet->usedrange->rows->count;
	my $numDRow  = "F" . $rowcount;
	my $dataArr  = $sheet->Range("A1:$numDRow")->{'Value'};
	my %dict = ();
	
	for (my $index = 1; $index < @$dataArr; $index++){
		my $data = $$dataArr[$index];
		$dict{$$data[0]} = $$data[1];
	}
	$workbook->Save();
	$workbook->Close();
	return %dict;
}

sub calculateRomRate{
	my $dataArr = shift;
	my $dict_romcount_ref = shift;
	my %dict_romcount = %$dict_romcount_ref;
	
	my $row1 = $$dataArr[0];
	my $row2 = $$dataArr[1];
	my $row3 = $$dataArr[2];
	my $row4 = $$dataArr[3];
	my $row29 = $$dataArr[27];
	my $row30 = $$dataArr[28];
	my $row31 = $$dataArr[29];
	my $row32 = $$dataArr[30];
	
	@$row29 = @$row1;
	@$row30 = @$row2;
	@$row31 = @$row3;
	@$row32 = @$row4;
	
	my $romcount_1 = $dict_romcount{$$row2[0]};
	my $romcount_2 = $dict_romcount{$$row3[0]};
	my $romcount_3 = $dict_romcount{$$row4[0]};
	
	for (my $index = 1; $index <= $colMax; $index++){
		if($mmodel ne "X5ProD"){
		$$row30[$index] = $$row2[$index] / $romcount_1*100;
		$$row30[$index] = $$row30[$index]."\%";
	}
		$$row31[$index] = $$row3[$index] / $romcount_2*100;
		$$row31[$index] = $$row31[$index]."\%";
		$$row32[$index] = $$row4[$index] / $romcount_3*100;
		$$row32[$index] = $$row32[$index]."\%";
	}
	$$dataArr[28] = $row30;
	$$dataArr[29] = $row31;
	$$dataArr[30] = $row32;
	
}






#制图
sub createChart{  
	
		my ($tempcurSheet,$temprangeStart,$temprangeEnd,$tempchartX,$tempchartY,$tempchartWidth,$tempchartHeight,$tempChartName)=@_;
		#print "tempcurSheet是$tempcurSheet\n";
		#print "temprangeStart是$temprangeStart\n";
		#print "temprangeEnd是$temprangeEnd\n";
		#print "tempchartX是$tempchartX\n";
		#print "tempchartY是$tempchartY\n";
		#print "tempchartWidth是$tempchartWidth\n";
		#print "tempchartHeight是$tempchartHeight\n";
		#print "tempChartName是$tempChartName\n";
		
		
		$tempchartRange = $tempcurSheet->Range( "$temprangeStart:$temprangeEnd" );        # 数据源范围
		
		$Chart=$tempcurSheet->ChartObjects->Add($tempchartX, $tempchartY, $tempchartWidth, $tempchartHeight)->Chart;
		
		$Chart->SetSourceData(
		        {
		                Source =>$tempchartRange,                                #　数据源RANG
		                #PlotBy =>xlRows,                    # 以源数据库横向还是竖向为标准生成图表 xlRows eq 1
		        }
		);
	  $Chart->{ChartType} = xlColumnClustered;
	  $Chart->{'HasTitle'} = 1;
	  $Chart->ChartTitle->Characters->{Text} = $tempChartName;
		$Chart->ChartArea->Format->ThreeD->{'RotationX'}=0;
		$Chart->ChartArea->Format->ThreeD->{'RotationY'}=90;
		
  
} 