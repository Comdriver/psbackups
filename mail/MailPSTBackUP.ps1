# ��������� ��������� ����������� ���� �������� ������ � �������� �����

#=========== ���� ���������� ============
#$localpath = "\\EXCHNG\BackUp"					# ��������� �����
$savepath = "\\xxx.xxx.xxx.xxx\MailBackUp"				# ������� �����
$maxcount = 2							# ���������� �������� ������ � ������� �����
$datefolder = [System.DateTime]::Now.ToString("yyyy-MM-dd")	# ������ ����� �����


#=========== �������� ����� ��� ����������
if(!(Test-Path -Path $savepath\$datefolder )){New-Item $savepath\$datefolder -type directory}
#if(!(Test-Path -Path $localpath )){New-Item $localpath -type directory} #�� ��������� ��� ������� ���

#=========== ������� ������ ������ � ��������� ������
#Get-ChildItem -Path $localpath *.* -File -Recurse | foreach { $_.Delete()}

#=========== ������� ������ � �����

#�������� ������
$sess = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://Exch/PowerShell/ -Authentication Kerberos

#������� ������� ������ ������� ��������
Invoke-Command -Session $sess -ScriptBlock { Get-MailboxExportRequest | Remove-MailboxExportRequest -Confirm:$false }

#������ � ���
Add-Content $savepath\mail.log "$(Get-Date)`texport`t$datefolder"

#����� ������ (����������: ������� �������� ������� �� ����� 1000 ������, ��� ������� 1000 � ����� �������� �������� -ResultSize unlimited ��� Get-Mailbox)
$AllMailboxes = Invoke-Command -Session $sess -ScriptBlock { Get-Mailbox }

#������� ������
$backup = join-path $savepath $datefolder
$logcount = 0 #���������� ������������ ������
ForEach ( $MailBox in $AllMailboxes )
	{
	Invoke-Command -Session $sess -ScriptBlock {param ($MailBox, $path)  New-MailboxExportRequest -Mailbox $Mailbox -FilePath "$path\$($Mailbox).pst" } -ArgumentList $MailBox.Alias, $backup

	#�������� �������� �������� ������ ������� ������
	while (((Invoke-Command -Session $sess -ScriptBlock {param ($MailBox) Get-MailboxExportRequest -mailbox $MailBox} -ArgumentList $MailBox.Alias) | ? {$_.Status -eq �Queued� -or $_.Status -eq �InProgress�}))
	     {Start-Sleep -s 10}
	$expmail = Invoke-Command -Session $sess -ScriptBlock {param ($MailBox) Get-MailboxExportRequest -mailbox $MailBox} -ArgumentList $MailBox.Alias
	Add-Content $savepath\mail.log "$(Get-Date)`t$($expmail.Status)`t$($MailBox.Alias)`t`t$($expmail.filepath)"
	$logcount ++
#	Invoke-Command -Session $sess -ScriptBlock {param ($ide) Remove-MailboxExportRequest $ide -Confirm:$false} -ArgumentList $expmail.Identity #��������� ������ ��������
	}

#���� �������� ������� ���������� ����� ������ �� ��������, �������� ��
Add-Content $savepath\mail.log "====== Other results ======"
#Invoke-Command -Session $sess -ScriptBlock {Get-MailboxExportRequest -status failed | Get-MailboxExportRequestStatistics -IncludeReport} | Format-List > "$savepath\errors $datefolder.log"
#Add-Content $savepath\mail.log "$savepath\errors $datefolder.log"
Add-Content $savepath\mail.log "== END of other results ==="

#=========== �������� ������ ����� � ��� �������� ������

#����� ������ � ������� ����� ������������, ���������� � ��������� ���������� �������
$folders = Get-ChildItem $savepath | ?{ $_.PSIsContainer } | Select-Object Name | Sort-Object Name -Descending
while ($folders.Count -gt $maxcount)
	{
	$remfolder = $($folders[$folders.Count-1]).name
	Remove-Item $savepath\$remfolder -recurse
	Add-Content $savepath\mail.log "$(Get-Date)`tremoved folder`t$remfolder backup"
	$folders = Get-ChildItem $savepath | ?{ $_.PSIsContainer } | Select-Object Name | Sort-Object Name -Descending
	}

# ������ � ��� ��������
$size =Get-ChildItem (join-path $savepath $datefolder) | Measure-Object -property length -sum
$boxes=(Get-ChildItem (join-path $savepath $datefolder) | Measure-Object).count
$size = "{0:N2}" -f ($size.sum / 1MB) + " MB"
Add-Content $savepath\mailsize.log "$($datefolder)`t$size`t$boxes ($logcount) mailboxes of $($AllMailboxes.count) processed"
Add-Content $savepath\mail.log "$(Get-Date)`t$($datefolder)`t$size saved`r`n`t`t`t$boxes ($logcount) mailboxes of $($AllMailboxes.count) processed`r`n======================================`r`n`r`n"


#Read-Host -Prompt "Press Enter to exit"