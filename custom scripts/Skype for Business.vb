
Dim FEPool As String = Connection.GetConfigParm("Custom\SkypeforBusiness\FEPool")
Dim SipAddressType As String = Connection.GetConfigParm("Custom\SkypeforBusiness\SipAddressType")
Dim UserPrincipalName As String = VI_GetValueofObject("Person","UID_Person",$UID_Person$,"DefaultEmailAddress")

Dim SMTPServer As String = Connection.GetConfigParm("Common\MailNotification\SMTPRelay")
Dim DefaultSender As String = Connection.GetConfigParm("Common\MailNotification\DefaultSender")

Dim Recipients As String = Connection.GetConfigParm("Custom\SkypeforBusiness\MailParameters\Recipients")
Dim Subject As String = Connection.GetConfigParm("Custom\SkypeforBusiness\MailParameters\Subject")
Dim MailBody As String = Connection.GetConfigParm("Custom\SkypeforBusiness\MailParameters\MailBody")

Dim script As New StringBuilder()

script.AppendLine("$DateTime = (Get-Date -Format 'dd-MM-yyyy;HH-mm-ss').toString()")
script.AppendLine("$FilesToSend = ""C:\IDM_scripts\Logs\Skype4Business\Skype4BusinessEnable($DateTime).log""")

script.AppendLine("Start-Transcript ""$FilesToSend""")

script.AppendLine("$FEPool = '" + FEPool + "'")
script.AppendLine("$SipAddressType = '" + SipAddressType + "'")
script.AppendLine("$UserPrincipalName = '" + UserPrincipalName + "'")

script.AppendLine("$SMTPServer = '" + SMTPServer + "'")
script.AppendLine("$DefaultSender = '" + DefaultSender + "'")

script.AppendLine("$Subject = '" + Subject + "'")
script.AppendLine("$MailBody = '" + MailBody + "'")

script.AppendLine("[string[]]$Recipients = " + Recipients)

script.AppendLine("Import-Module ""C:\Program Files\Common Files\Skype for Business Server 2015\Modules\SkypeForBusiness"" -verbose")

script.AppendLine("Enable-CsUser -Identity $UserPrincipalName -RegistrarPool $FEPool -SipAddressType $SipAddressType -Verbose -WhatIf")

'script.AppendLine("Disable-CsUser -Identity $UserPrincipalName -Verbose -WhatIf")

script.AppendLine("Stop-Transcript")

script.AppendLine("[string[]]$FilesToSend = $FilesToSend")

script.AppendLine("Send-MailMessage -From $DefaultSender -to $Recipients -Subject $Subject -Body $MailBody -SmtpServer $SMTPServer -Encoding UTF8 -BodyAsHtml -Attachments $FilesToSend -ErrorAction Continue -Verbose")

Value = script.ToString()

