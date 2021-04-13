<#
.SYNOPSIS
    This script will count the number of faxes received for each date.

.DESCRIPTION
    The data is contained in a file called ACTIVITYEX.log using Biscom Fax Suite. 
	This script will copy a portion of the log file that contains the same date 
    for a given column. Copying this data is accomplished by using
    Out-File.

    This script is setup as a scheduled task to run at 3 AM daily.

.NOTES
    File name - Count CS Faxes per day.ps1
    Version   - 1.0
    Date      - 7/21/2014
    Author    - Lou Garramone
    E-mail    - lou@lg4tech.com
#>

# Defining variables and loading classes.
$SmtpClient = New-Object System.Net.Mail.SmtpClient
$MailMessage = New-Object System.Net.Mail.MailMessage
# Get the date and subtract 1 day.
$Date = (Get-Date).AddDays(-1).ToString("MM/dd/yy")
[reflection.assembly]::LoadWithPartialName("System.IO") | Out-Null

# Mail server settings.
$SmtpClient.Host = "ip.address.here.x"
$MailMessage.From = ("FromEmail@email.com") #----------------------------------#
$MailMessage.To.Add("ToEmail@email.com") #----------------------------------#

# Read the data from file, copy each line that contains the current date and output to file.
Get-Content "C:\Program Files (x86)\Biscom\FaxcomQ_Customer_Service_Incoming\Biscom\FaxcomQ_Customer_Service_Incoming\bisfax\util\data\ACTIVITYEX.log" | ForEach-Object {if($_.contains("$Date")){$_ | Out-File "C:\Scripts\CS Fax Log.txt" -Append}}

# Count the number of lines in the new file and store to variable.
$LineCount = [System.IO.File]::ReadAllLines("C:\Scripts\CS Fax Log.txt").Count

# Attach the new log file that was filtered.
$MailMessage.Attachments.add("C:\Scripts\CS Fax Log.txt")
$MailMessage.Subject = "Total Customer Service Fax Count - $Date : $LineCount"

# E-mail the new log file.
$SmtpClient.Send($MailMessage) | Start-Sleep 2

# Clean up/remove file.
$MailMessage.Dispose()
$SmtpClient.Dispose()
if (Test-Path "C:\Scripts\CS Fax Log.txt"){
	Remove-Item "C:\Scripts\CS Fax Log.txt" -Force
	}