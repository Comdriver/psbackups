# ��������� ����������� ������ �� ������� � �������� �����
#
# �����: ����� �.�. (w31@zodcode.ru)
#
#=========== ���� ���������� ============
$sourcepath = 	"\\xxx.xxx.xxx.xxx\path"			#�������� 1
$localpath = 	"D:\localpath"			#�������� 2
$incsavepath = 	"D:\BackUp\Full backup"			#������ ������ �� ����� ������� (� �.�. ��������� � ���������)
$savepath = 	"D:\BackUp\daily backup"			#����� �� ��������� ���������� ����
$logfile = 	"D:\BackUp\"

$empty	=	"$logfile\EMPTY"
$maxcount = 3						#������� ����� �������
$datefolder = [System.DateTime]::Now.ToString("yyyy-MM-dd") #��� ����� �������� ������


#=========== �������� ������� �����
if(!(Test-Path -Path $savepath\$datefolder )){New-Item $savepath\$datefolder -type directory}		#���������� �����
if(!(Test-Path -Path $incsavepath )){New-Item $incsavepath -type directory}				#������ �����
if((Test-Path -Path $logfile\EMPTY )){Remove-Item -LiteralPath $empty -Force}				#������ �����
New-Item $empty -type directory

#=========== �������� ���������� �����
$folders = Get-ChildItem $savepath | ?{ $_.PSIsContainer } | Select-Object Name | Sort-Object Name -Descending		#������ �����
while ($folders.Count -gt $maxcount)											#�������� ���� �� �������� $maxcount �����
	{
	$remfolder = $($folders[$folders.Count-1]).name
	$remfolder = Join-Path $savepath $remfolder
	$remfolder
	robocopy "$empty" "$remfolder" /MIR
	Remove-Item -LiteralPath $remfolder -Force
	$folders = Get-ChildItem $savepath | ?{ $_.PSIsContainer } | Select-Object Name | Sort-Object Name -Descending
	}

#=========== ���������� ����� ������ � �����
$log=(join-path $logfile "\$datefolder copyLAN.log")
robocopy "$sourcepath" "$savepath\$datefolder" /E /FFT /R:5 /W:10 /Z /NP /NDL /XJD /MT:4 /unilog:"$log"		#�� �������� �����
$log=(join-path $logfile "\$datefolder copyLocal.log")
robocopy "$localpath" "$savepath\$datefolder\8. ������ �����" /E /FFT /R:5 /W:10 /Z /NP /NDL /XJD /MT:4 /unilog:"$log"		#�� ��������� �����

#=========== ������ ����� ������ � �����
$log=(join-path $logfile "\$datefolder copyFull.log")
robocopy "$savepath\$datefolder" "$incsavepath" /E /FFT /R:10 /W:10 /Z /NP /NDL /XJD /MT:6 /unilog:"$log"		#�� ���������� �����
