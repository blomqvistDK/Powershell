$m = Get-Module -List ActiveDirectory
if(!$m) {
    Write-Host "AD Powershell Module is not installed. Install it and run the script again."
    Exit
}

else{
    Import-Module ActiveDirectory
}


#Mail
$smtpServer = "ex01.domain.com"
$from = "MAN01@domain.com"
$recipient = "user@domain.com"

#Variables
[string]$SavePath = "$PSScriptRoot\reports"
[int]$ComputerPasswordAgeDays = 90
[string]$ADSearchBase = "DC=domain,DC=com"


#--------------------------------------
#Name: sendMail 
#Arguments: 2 Arguments:
#           1. Computer Count
#           2. File to be attached
#Return: -
#Description: Sends an E-mail to recipient defined in $recipient containing the csv file and computer count.
#--------------------------------------

function sendMail{

     $mailinfo_computer_count = $args[0]
     $file = $args[1]
  

    $body = "Computers which have changed computer password within the last $ComputerPasswordAgeDays days.</br></br>"
    $body += $mailinfo_computer_count

    $subject = "Antal computere: " + $mailinfo_computer_count
    
                 
     #Sending email 
     Send-MailMessage -SmtpServer $smtpServer -From $from -To $recipient -Subject $subject -Body $body -BodyAsHtml -Attachments $file
  
}


$Date = (Get-Date -format "MM-dd-yyyy")
IF ((test-path $SavePath) -eq $False) { md $SavePath }
$ExportFile = "$SavePath\Workstations-$Date.csv"
$ComputerStaleDate = (Get-Date).AddDays(-$ComputerPasswordAgeDays)
$Computers = Get-ADComputer -SearchBase $ADSearchBase -filter { (passwordLastSet -ge $ComputerStaleDate) -and (OperatingSystem -notlike "*Server*") -and (OperatingSystem -like "*Windows*") } -properties Name, DistinguishedName, OperatingSystem,OperatingSystemServicePack, passwordLastSet,LastLogonDate,Description
$ComputerTrimmed = $Computers.Count
write-output "Number of computers: " ($ComputerTrimmed) 

$Computers | export-csv $ExportFile -NoTypeInformation
sendMail $ComputerTrimmed $ExportFile