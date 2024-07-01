# It's essential to perform a periodical check to see if your devices are up
# Windows Task Schedular enable admins perform that
# And PowerShell is a perfect tool to write automation scripts
# which enable to trigger alerts, send emails, etc.

Import-Module C:\Users\VM01\GitHubProjects\SendEmail\MailModule.psm1
$MailAccount=Import-Clixml -Path C:\Users\VM01\GitHubProjects\SendEmail\outlook.xml
$MailPort=587
$MailSMTPServer="smtp-mail.outlook.com"
$MailFrom=$MailAccount.Username
$Mailto="myemail@outlook.com"

$ServersFile="C:Users\VM01\GitHubProjects\ServersList.csv"

$Servers=Import-Csv -Path $ServersFile -Delimiter ","

$Export=[System.Collections.ArrayList]@()

foreach($Server in $Servers){
    $ServerName=$Server.ServerName
    $LastStatus=$Server.LastStatus
    $DownSince=$Server.DownSince
    $LastDownAlert=Server.LastDownAlertTime
    $Alert=$false
    $Connection=Test-Connection $Server.ServerName -Count 1
    $DateTime=Get-Date

    if($Connection.Status -eq "Success"){
        if($LastStatus -ne "Success"){
            $Server.$DownSince=$null
            $Server.$LastDownAlertTime=$null
            Write-Output "$ServerName is now up"
            $Alert=$true 
            $Subject="$ServerName is now up!"
            $Body="<h2>$ServerName is now up!</h2>"
            $Body="<p>$ServerName is now up at $DateTime</p>"           
        }
    }else{
        if($LastStatus -eq "Status"){
            Write-Output "$ServerName is now down"
            $Server.DownSince=$DateTime
            $server.LastDownAlertTime=$DateTime
            $Alert=$true
            $Subject="$ServerName is now down!"
            $Body="<h2>$ServerName is now down!</h2>"
            $Body="<p>$ServerName is now down at $DateTime</p>" 
        }else{
            $DownFor=$((Get-Date -Date $DateTime)-(Get-Date -Date $DownSince)).Days
            $SinceLastDownAlert=$((Get-Date -Date $DateTime)-(Get-Date -Date $LastDownAlert)).Days
            if(($DownFor -ge 1) -and ($SinceLastDownAlert -ge 1)){
                Write-Output "It has been $SinceLastDownAlert days since last alert"
                Write-Output "$ServerName is still down for $DownFor days"
                $Server.$LastDownAlertTime=$DateTime
                $Alert=$true
                $Subject="$ServerName is still down for $DownFor days!"
                $Body="<h2>$ServerName has been down for $DownFor days!</h2>"
                $Body="<p>$ServerName has been down since $DownSince</p>" 
            }
           
        }
    }

    if($Alert){
        Send-MailKitMessage -From $MailFrom -To $MailTo -SMTPServer $MailSMTPServer -Port $MailPort -Subject $Subject -Body $Body -BodyAsHtml -Credential $MailAccount
    }

    $Server.LastStatus=$Connection.Status
    $Server.LastCheckTime=$DateTime

    [void]$Export.Add($Server)
}

$Export | Export-Csv -Path $ServersFile -Delimeter "," -NoTypeInformation