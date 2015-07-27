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

#my @PHONEMODELS = qw(X3t X3V Xplay Xplay3s X5L Xshot Y22iL Y27 Y13L Y22L X5V Y28L Y23L X5SL);# X5MaxL);  # ����
#my @PHONEMODELS = qw(Xplay Xplay3s Xshot X5MaxL);  # ����
#my @PHONEMODELS = qw(Xplay Xplay3s Xshot X5MaxL X5Max+ Y29L X5L X5MaxV X5ProD);#  X5M);  # ����
my @PHONEMODELS = qw(X5ProD);  # ����
my $dir = getcwd;
my $workbook;
my $pageNum = 2;	

@TAGSERIES = qw(�ĵ� ���� ϵͳ���� ���� ���� �Զ��ػ� ���� ���� ���� ���� vivo��԰ ��Ļ���� ���� �������� С�Ҽ� ��� ͨ�� ����Ӧ ֹͣ���� ���� ������ 
i���� i��Ƶ ��Ƶ i�ܼ� ����� ����̵� ״̬�� ��ϵ�� ���� �����ʼ� �ļ����� ���뷨 ����Ԥ�� ��ǩ ������ ����ӰԺ ������ 
�����ź� WIFI GPS ���� ָ�� OTG ������� ������� ʱ��(��|�޷�)׼ �ʺ� С�� ��Ϸ ��������� �Ʒ��� ��� ���� ���� photo+ i���� HIFI ���� ��� �ֵ�Ͳ ������ ���� SD�� ��Ļ���� �� vivo�������� 
���� NFC ROOT ��Ϸ���� �ÿ�ģʽ USB vivo����
����ʶ�� 
����K�� ¼�� ���� ���� ���� ��Թ ����);


$totolRow = 600;#��ʼ����sheetֻ��Ҫ600������  # modify kongqiao
$numDRow = MM.$totolRow;#400�ж�Ӧ��excel����
$ModuleNum = 5;
$colMax = 85;  # ģʽ�� 
$endRow = 27;
$startRow = 11;
$rom_count_name = "rom_count.xlsx";
main();

sub main{
	Process();	#���˴������������а汾ɸѡ��ɸѡ���ԣ����ҵ��汾���������ո���������ͬ�ķ���������ɾ��������	
}


sub Process{
		foreach (@PHONEMODELS) {
										print $_."\n";
										$moduleTempNum=0;
										#��x3t
										$filePath = $dir."\/".$_.".xlsx";	# ���·��
										$workbook = $Excel->Workbooks->Open($filePath);
										$workSheet=$workbook->Sheets("ͳ��");
										#$selectSheet = $workbook->Sheets($pageNum);
										$PhoneName=$_;
										#����x3tEXCEL���ݵ�����
										$DataArray = $workSheet->Range("A1:$numDRow")->{'Value'};
										$DataLength=@$DataArray-1;

#=pod
									 									  
									  #$filePathLast = $dir."\/".$_."�û�����".".xlsx";	# ���·��
									  # modify kongqiao 20150714
									  my @allfilePathLast = <*.xlsx>;
	  								foreach $path (@allfilePathLast){
	  									if($path =~ /ROM2.0_$__.*�û�����/){
	  										$filePathLast = $dir."\/".$path;
	  										print "��������·���ǣ�$filePathLast \n";
	  									}
	 									}
										$workbookLast = $Excel->Workbooks->Open($filePathLast);
										$workSheetLast=$workbookLast->Sheets("ͳ��");
										#����EXCEL���ݵ�����
										$DataArrayLast = $workSheetLast->Range("A1:$numDRow")->{'Value'};
										$DataLength=@$DataArray-1;	
#=cut		

#��������

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
															$$seventeenCol[1]="��Ͷ����";
															$$seventeenCol[2]="����Ͷ����";
															$$seventeenCol[10]="�ĵ�";
															$$seventeenCol[11]="����";
															
															$$twentyfourthCol[$count-1]=$$firstCol[$count-1];
															$$twentyfourthCol[$count-1+9]=$$twentyfourthCol[$count-1];	
															$$twentyfourthCol[1]="����Ͷ��ռ��";
															$$twentyfourthCol[2]="�ĵ�Ͷ��ռ��";
															$$twentyfourthCol[3]="����Ͷ��ռ��";
															$$twentyfourthCol[10]="ROM����";						
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
															#print "tempArray[0]��$$tempArray[0]\n";
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


										#���Ӵ�Сð������
										#$DataArray = $workSheet->Range("A1:$numDRow")->{'Value'};
										foreach $tempCount(2..$colMax-1-5){
													$ten[$tempCount-2]=$$tenthCol[$tempCount-1];
													$third[$tempCount-2]=$$thirteennCol[$tempCount-1];
													
										}
										#print "ten[$colMax-1-5-2]��$ten[$colMax-1-5-2]\n";
										$appArrayLength=$colMax-1-5-2+1;
										
										#print "appArrayLength����Ϊ��$appArrayLength\n";
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
										#��ʼ�������������Ƿ���ͻ��
										$thirdLast=$$DataArrayLast[12];	
										$fourLast=$$DataArrayLast[3];	
										$thirteennCol=$$DataArray[12];
										$twelvethCol=$$DataArray[11];
										$eleventhCol=$$DataArray[10];
										$fourCol=$$DataArray[3];

										$tenthCol=$$DataArray[9];
										$temp=0;
										#print "$$tenthCol[$colMax-1-5-1]��$$tenthCol[$colMax-1-5-1]\n";

#����ͬ�汾�� Ҫ�����ϴ���3
#�벻ͬ�汾�Ƚϣ������ϲ���Ҫ��
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
															
															$$tempArray[12]="����ͻ��ģ����:";
													}else{
															
															$$tempArray[13]=$suddenChange[$tempCount];
															$tempCount++;
													}
													
													$$DataArray[$tempNum-1]=$tempArray;
										}
#=cut			
										#�ж�����ǰ��λ
										$tempCount=0;
										foreach $temp(134..142){
													$tempArray=$$DataArray[$temp-1];
													if($temp eq 134){
															$$tempArray[7]="����ǰ��λ��:";

													}else{
															$$tempArray[8]=$ten[$tempCount];
															
															$tempCount++;
													}
													
													$$DataArray[$temp-1]=$tempArray;
										}
										$workSheet->Range("A1:$numDRow")->{'value'}=$DataArray;
											

											
											
										#��ʼ��ͼ
										foreach $count(1..7){				
													if($count eq 1){	
															$rangeStart=A17;
															$rangeEnd=C20;
															$chartX=0;
															$chartY=431;#29��С����
															$chartWidth=360;
															$chartHeight=225;
															$chartName=$PhoneName."��Ͷ����������Ͷ������";
															createChart($workSheet,$rangeStart,$rangeEnd,$chartX,$chartY,$chartWidth,$chartHeight,$chartName);
														
															
													}elsif($count eq 2){
															$rangeStart=A24;
															$rangeEnd=B27;
															$chartX=430;
															$chartY=431;
															$chartWidth=360;
															$chartHeight=225;
															$chartName=$PhoneName."����Ͷ��ռ��";
															createChart($workSheet,$rangeStart,$rangeEnd,$chartX,$chartY,$chartWidth,$chartHeight,$chartName);
													}elsif($count eq 3){
															$rangeStart=A24;
															$rangeEnd=D27;
															$chartX=0;
															$chartY=674;
															$chartWidth=360;
															$chartHeight=225;
															$chartName=$PhoneName."��Ͷ��ռ��";
															createChart($workSheet,$rangeStart,$rangeEnd,$chartX,$chartY,$chartWidth,$chartHeight,$chartName);
													}elsif($count eq 4){
															$rangeStart=A1;
															$rangeEnd=CA4;
															$chartX=0;
															$chartY=915;
															$chartWidth=1400;
															$chartHeight=281;
															$chartName=$PhoneName."��ROM�汾��Ͷ����";
															createChart($workSheet,$rangeStart,$rangeEnd,$chartX,$chartY,$chartWidth,$chartHeight,$chartName);
													}elsif($count eq 5){
															$rangeStart=A10;
															$rangeEnd=CA13;
															$chartX=0;
															$chartY=1215;
															$chartWidth=1400;
															$chartHeight=281;
															$chartName=$PhoneName."��ROM�汾��Ͷ��ռ��";
															createChart($workSheet,$rangeStart,$rangeEnd,$chartX,$chartY,$chartWidth,$chartHeight,$chartName);
													}elsif($count eq 6){
															$value=$workSheet->Range("A15:A15")->{'Value'};
															print "value��$value\n";
															if($value ne ""){
															print "value��$value\n";
															$rangeStart=A14;
															$rangeEnd=CA15;
															$chartX=0;
															$chartY=1700;
															$chartWidth=1400;
															$chartHeight=281;
															$chartName=$PhoneName."��".$value."�汾";
															createChart($workSheet,$rangeStart,$rangeEnd,$chartX,$chartY,$chartWidth,$chartHeight,$chartName);
															}
													}elsif($count eq 7){
															$rangeStart=A28;
															$rangeEnd=CA31;
															$chartX=0;
															$chartY=1500;
															$chartWidth=1400;
															$chartHeight=281;
															$chartName=$PhoneName."��ROM�汾ʵ��Ͷ��ռ��";
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


#��ȡ����ת���ɴ�д��ĸ
sub getColumnName{  
    $first;  
    $lastvalue;  
    $columnNum=$_[0];
    #print "column��ֵΪ".$columnNum."\n";
    $result = "";  

    $first = int($columnNum / 27);  
    $lastvalue = $columnNum - ($first * 26);  
  	
  	#print "��һ��".$first."            ".$lastvalue."\n"; 
  	$temp;
    if ($first gt 0){ #ltС�� gt���� ne������ 
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
      #print "��ʱֵΪ��".$result."\n"; 
  
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






#��ͼ
sub createChart{  
	
		my ($tempcurSheet,$temprangeStart,$temprangeEnd,$tempchartX,$tempchartY,$tempchartWidth,$tempchartHeight,$tempChartName)=@_;
		#print "tempcurSheet��$tempcurSheet\n";
		#print "temprangeStart��$temprangeStart\n";
		#print "temprangeEnd��$temprangeEnd\n";
		#print "tempchartX��$tempchartX\n";
		#print "tempchartY��$tempchartY\n";
		#print "tempchartWidth��$tempchartWidth\n";
		#print "tempchartHeight��$tempchartHeight\n";
		#print "tempChartName��$tempChartName\n";
		
		
		$tempchartRange = $tempcurSheet->Range( "$temprangeStart:$temprangeEnd" );        # ����Դ��Χ
		
		$Chart=$tempcurSheet->ChartObjects->Add($tempchartX, $tempchartY, $tempchartWidth, $tempchartHeight)->Chart;
		
		$Chart->SetSourceData(
		        {
		                Source =>$tempchartRange,                                #������ԴRANG
		                #PlotBy =>xlRows,                    # ��Դ���ݿ����������Ϊ��׼����ͼ�� xlRows eq 1
		        }
		);
	  $Chart->{ChartType} = xlColumnClustered;
	  $Chart->{'HasTitle'} = 1;
	  $Chart->ChartTitle->Characters->{Text} = $tempChartName;
		$Chart->ChartArea->Format->ThreeD->{'RotationX'}=0;
		$Chart->ChartArea->Format->ThreeD->{'RotationY'}=90;
		
  
} 