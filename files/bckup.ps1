# резервное копирование файлов по запросу в заданную папку
#
# Автор: Зотов В.И. (w31@zodcode.ru)
#
#=========== блок параметров ============
$sourcepath = 	"\\xxx.xxx.xxx.xxx\path"			#источник 1
$localpath = 	"D:\localpath"			#источник 2
$incsavepath = 	"D:\BackUp\Full backup"			#полная версия со всеми файлами (в т.ч. удалёнными в оригинале)
$savepath = 	"D:\BackUp\daily backup"			#копия за прошедшее количество дней
$logfile = 	"D:\BackUp\"

$empty	=	"$logfile\EMPTY"
$maxcount = 3						#сколько копий хранить
$datefolder = [System.DateTime]::Now.ToString("yyyy-MM-dd") #имя папки текущего бекапа


#=========== создание целевых папок
if(!(Test-Path -Path $savepath\$datefolder )){New-Item $savepath\$datefolder -type directory}		#ежедневные папки
if(!(Test-Path -Path $incsavepath )){New-Item $incsavepath -type directory}				#полная копия
if((Test-Path -Path $logfile\EMPTY )){Remove-Item -LiteralPath $empty -Force}				#пустая папка
New-Item $empty -type directory

#=========== проверка количества папок
$folders = Get-ChildItem $savepath | ?{ $_.PSIsContainer } | Select-Object Name | Sort-Object Name -Descending		#список папок
while ($folders.Count -gt $maxcount)											#удаление пока не остнется $maxcount папок
	{
	$remfolder = $($folders[$folders.Count-1]).name
	$remfolder = Join-Path $savepath $remfolder
	$remfolder
	robocopy "$empty" "$remfolder" /MIR
	Remove-Item -LiteralPath $remfolder -Force
	$folders = Get-ChildItem $savepath | ?{ $_.PSIsContainer } | Select-Object Name | Sort-Object Name -Descending
	}

#=========== ежедневная копия файлов и папок
$log=(join-path $logfile "\$datefolder copyLAN.log")
robocopy "$sourcepath" "$savepath\$datefolder" /E /FFT /R:5 /W:10 /Z /NP /NDL /XJD /MT:4 /unilog:"$log"		#из удалённой папки
$log=(join-path $logfile "\$datefolder copyLocal.log")
robocopy "$localpath" "$savepath\$datefolder\8. Личные папки" /E /FFT /R:5 /W:10 /Z /NP /NDL /XJD /MT:4 /unilog:"$log"		#из локальной папки

#=========== полная копия файлов и папок
$log=(join-path $logfile "\$datefolder copyFull.log")
robocopy "$savepath\$datefolder" "$incsavepath" /E /FFT /R:10 /W:10 /Z /NP /NDL /XJD /MT:6 /unilog:"$log"		#из ежедневной папки
