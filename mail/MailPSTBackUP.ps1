# выполняет резервное копирование всех почтовых ящиков в заданную папку

#=========== блок параметров ============
#$localpath = "\\EXCHNG\BackUp"					# локальна¤ копия
$savepath = "\\xxx.xxx.xxx.xxx\MailBackUp"				# сетева¤ копия
$maxcount = 2							# количество хранимых версий в сетевой копии
$datefolder = [System.DateTime]::Now.ToString("yyyy-MM-dd")	# формат имени папки


#=========== создание папки для сохранения
if(!(Test-Path -Path $savepath\$datefolder )){New-Item $savepath\$datefolder -type directory}
#if(!(Test-Path -Path $localpath )){New-Item $localpath -type directory} #не требуется для сетевых шар

#=========== очистка старых файлов в локальном бекапе
#Get-ChildItem -Path $localpath *.* -File -Recurse | foreach { $_.Delete()}

#=========== экспорт ящиков в папку

#создание сессии
$sess = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://Exch/PowerShell/ -Authentication Kerberos

#очистка старого списка заданий экспорта
Invoke-Command -Session $sess -ScriptBlock { Get-MailboxExportRequest | Remove-MailboxExportRequest -Confirm:$false }

#запись в лог
Add-Content $savepath\mail.log "$(Get-Date)`texport`t$datefolder"

#поиск ящиков (ѕ–»ћ≈„јЌ»≈: команда выполнит экспорт не более 1000 ящиков, при наличии 1000 и более добавить параметр -ResultSize unlimited дл¤ Get-Mailbox)
$AllMailboxes = Invoke-Command -Session $sess -ScriptBlock { Get-Mailbox }

#экспорт ящиков
$backup = join-path $savepath $datefolder
$logcount = 0 #количество обработанных ящиков
ForEach ( $MailBox in $AllMailboxes )
	{
	Invoke-Command -Session $sess -ScriptBlock {param ($MailBox, $path)  New-MailboxExportRequest -Mailbox $Mailbox -FilePath "$path\$($Mailbox).pst" } -ArgumentList $MailBox.Alias, $backup

	#ожидание экспорта выгрузка списка готовых ящиков
	while (((Invoke-Command -Session $sess -ScriptBlock {param ($MailBox) Get-MailboxExportRequest -mailbox $MailBox} -ArgumentList $MailBox.Alias) | ? {$_.Status -eq 'Queued' -or $_.Status -eq 'InProgress'}))
	     {Start-Sleep -s 10}
	$expmail = Invoke-Command -Session $sess -ScriptBlock {param ($MailBox) Get-MailboxExportRequest -mailbox $MailBox} -ArgumentList $MailBox.Alias
	Add-Content $savepath\mail.log "$(Get-Date)`t$($expmail.Status)`t$($MailBox.Alias)`t`t$($expmail.filepath)"
	$logcount ++
#	Invoke-Command -Session $sess -ScriptBlock {param ($ide) Remove-MailboxExportRequest $ide -Confirm:$false} -ArgumentList $expmail.Identity #подчищаем список экспорта
	}

#сюда добавить экспорт оставшейся части списка со статусом, отличным от
Add-Content $savepath\mail.log "====== Other results ======"
#Invoke-Command -Session $sess -ScriptBlock {Get-MailboxExportRequest -status failed | Get-MailboxExportRequestStatistics -IncludeReport} | Format-List > "$savepath\errors $datefolder.log"
#Add-Content $savepath\mail.log "$savepath\errors $datefolder.log"
Add-Content $savepath\mail.log "== END of other results ==="

#=========== удаление старых копий и лог размеров архива

#поиск версий в сетевой папке сохраненения, сортировка в убывающем алфавитном порядке
$folders = Get-ChildItem $savepath | ?{ $_.PSIsContainer } | Select-Object Name | Sort-Object Name -Descending
while ($folders.Count -gt $maxcount)
	{
	$remfolder = $($folders[$folders.Count-1]).name
	Remove-Item $savepath\$remfolder -recurse
	Add-Content $savepath\mail.log "$(Get-Date)`tremoved folder`t$remfolder backup"
	$folders = Get-ChildItem $savepath | ?{ $_.PSIsContainer } | Select-Object Name | Sort-Object Name -Descending
	}

# запись в лог размеров
$size =Get-ChildItem (join-path $savepath $datefolder) | Measure-Object -property length -sum
$boxes=(Get-ChildItem (join-path $savepath $datefolder) | Measure-Object).count
$size = "{0:N2}" -f ($size.sum / 1MB) + " MB"
Add-Content $savepath\mailsize.log "$($datefolder)`t$size`t$boxes ($logcount) mailboxes of $($AllMailboxes.count) processed"
Add-Content $savepath\mail.log "$(Get-Date)`t$($datefolder)`t$size saved`r`n`t`t`t$boxes ($logcount) mailboxes of $($AllMailboxes.count) processed`r`n======================================`r`n`r`n"


#Read-Host -Prompt "Press Enter to exit"
