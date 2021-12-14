
# Get the OS build number so we know which update package to remove
$BuildNumber = (Get-ComputerInfo).OSBuildNumber

# Email parameters
$MailRecipient = "<EMAIL_ADDRESS_OF_DISTRIBUTION_GROUP>"
$MailSender = "<INSERT_PROPER_ACTOR>-$env:COMPUTERNAME@<EMAIL_DOMAIN>"
$Subject = "[Success] <INSERT_PROPER_ACTOR> - March Update Fix Installation"
$MailSMTPServer = "aspmx.l.google.com"

$LocalLogLocation = "C:\Windows\Temp\$env:COMPUTERNAME-march-update-removal-error.txt"



# If the build is 20H2 or 2004
if ($BuildNumber -eq '19042' -or $BuildNumber -eq '19041')
{
    try
    {
        # Copy the new Cumulative KB to a local source
        #Copy-Item -Path '\\FILESERVER\temp\windows10.0-kb5001567-x64_e3c7e1cb6fa3857b5b0c8cf487e7e16213b1ea83-2004-20H2.msu' -Destination 'C:\Windows\temp\'
        # Apply the new Cumulative KB
        wusa \\FILESERVER\temp\windows10.0-kb5001567-x64_e3c7e1cb6fa3857b5b0c8cf487e7e16213b1ea83-2004-20H2.msu /quiet /promptrestart

        $MailBody = "Computer $env:COMPUTERNAME installed the new March Cumulative Update on the system`n`n`n`I'm a bot and this message is automated. Replies will not be monitored. ;) glhf"

        Send-MailMessage -To $MailRecipient -From $MailSender -Subject $Subject -Body $MailBody -SmtpServer $MailSMTPServer
    }

    catch
    {
        Write-Output "$($_ | Select-Object -Property *)" > $LocalLogLocation

        $Subject = "[Failure] <INSERT_PROPER_ACTOR> - March Update Fix Installation"

        $MailBody = "Computer $env:COMPUTERNAME attempted to install an update but failed! The error message is attached.`n`n`n`I'm a bot and this message is automated. Replies will not be monitored. ;) glhf"

        # Send the email notification
        Send-MailMessage -To $MailRecipient -From $MailSender -Subject $Subject -Body $MailBody -Attachments $LocalLogLocation -SmtpServer $MailSMTPServer
    }
}

# If the build is 1909
elseif ($BuildNumber -eq '18363')
{
    try
    {
        # Copy the new Cumulative KB to a local source
        Copy-Item -Path '\\FILESERVER\temp\windows10.0-kb5001566-x64_b52b66b45562d5a620a6f1a5e903600693be1de0.msu' -Destination 'C:\Windows\temp\'
        # Apply the new Cumulative KB
        wusa 'C:\Windows\Temp\windows10.0-kb5001566-x64_b52b66b45562d5a620a6f1a5e903600693be1de0.msu' /promptrestart /quiet

        $MailBody = "Computer $env:COMPUTERNAME installed the new March Cumulative Update on the system`n`n`n`I'm a bot and this message is automated. Replies will not be monitored. ;) glhf"

        Send-MailMessage -To $MailRecipient -From $MailSender -Subject $Subject -Body $MailBody -SmtpServer $MailSMTPServer
        
        Remove-Item -Path 'C:\Windows\Temp\windows10.0-kb5001566-x64_b52b66b45562d5a620a6f1a5e903600693be1de0.msu'
    }

    catch
    {
        Write-Output "$($_ | Select-Object -Property *)" > $LocalLogLocation

        $Subject = "[Failure] <INSERT_PROPER_ACTOR> - March Update Fix Installation"

        $MailBody = "Computer $env:COMPUTERNAME attempted to install an update but failed! The error message is attached.`n`n`n`I'm a bot and this message is automated. Replies will not be monitored. ;) glhf"

        # Send the email notification
        Send-MailMessage -To $MailRecipient -From $MailSender -Subject $Subject -Body $MailBody -Attachments $LocalLogLocation -SmtpServer $MailSMTPServer
    }
}

else
{
    # The build number is not supported at this time    
}