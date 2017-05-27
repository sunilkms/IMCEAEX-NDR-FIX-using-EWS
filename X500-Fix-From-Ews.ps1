#
# Fix LegDn for Mail users
# useage
# to convert to x500 from csv : .\x500.ps1
# to make changes 	      : .\x500.ps1 -edit true
# 

param ($edit,$checkX500,$sendReport,$items=200)
. "D:\sunil\myscripts\X500fixews\find-Email-From-Folder-EWS.ps1"
. "D:\sunil\myscripts\X500fixews\MoveEmailtofolder-ews.ps1"

$skipped=0
$fixed=0
$multi=0
$notFound=0

$log = "D:\sunil\myscripts\X500fixews\log.txt"
Clear-Content $log

$m=(Get-Module | ? {$_.Name -like "Activedirectory"})
if (!$m) {"Connecting AD"; Import-Module ACtiveDirectory} else {"Already Connnected AD Skip.."}
$file = Get-X500Errorusers -items $items

Add-content -value ""  -path $log
Add-content -value "X500 AutoFix run Summary"  -path $log
Add-content -value "Number of Recipients found to be fixed:$($file.count)"  -path $log

Function CleanLegacyExchangeDN ([string]$imceaex)
{
    $imceaex = $imceaex.Replace("IMCEAEX-","")
    $imceaex = $imceaex.Replace("_","/")
    $imceaex = $imceaex.Replace("+20"," ")
    $imceaex = $imceaex.Replace("+5F","_")
    $imceaex = $imceaex.Replace("+28","(")
    $imceaex = $imceaex.Replace("+29",")")
    $imceaex = $imceaex.Replace("+2E",".")
    $imceaex = $imceaex.Replace("+2C",",")
    $imceaex = $imceaex.Replace("+21","!")
    $imceaex = $imceaex.Replace("+2B","+")
    $imceaex = $imceaex.Replace("+3D","=")
    $regex = New-Object System.Text.RegularExpressions.Regex('@.*')
    $imceaex = $regex.Replace($imceaex,"")
    $imceaex = $imceaex.Replace("+40","@")
    $imceaex # return object
}

function AddX500 {
        param (
                $user,
	              $X500
              )
Set-ADUser $user -Add @{proxyAddresses=$X500}
}

$data=@()
Add-content -value ""  -path $log
Add-content -value "USERS Validation - Below users Found with issues"  -path $log

foreach ($entry in $file) {

try {
	$us = Get-Recipient $Entry.Name.replace("'","" ) -ErrorAction Stop
	if ($us.count -gt 1) {
				write-host "Multiple Recipient found for user:" -n -f yellow
				write-host  $Entry.Name -f cyan -n
				write-host "Match SMTP from NDR" -f yellow
        Add-content -value "`nACTION REQUIRED ON THIS ITEM - Multiple Recipient found matching:$($Entry.name)" -path $log
				Add-content -value "`rCheck and add X500 ADDRESS: $(CleanLegacyExchangeDN $Entry.Address)" -path $log
				Add-content -value ""  -path $log
        $usr =$us[0].PrimarySmtpAddress.toString()
        $multi++
		} else {
			 $usr =$us.PrimarySmtpAddress.toString()
       Write-Host "User Found:$usr" -ForegroundColor Cyan		 				
		}		 
	$newEntry = New-Object -TypeName PSObject
	$newEntry | Add-Member -MemberType NoteProperty -Name user -Value $usr
	$newEntry | Add-Member -MemberType NoteProperty -Name X500 -Value $("X500:" + $(CleanLegacyExchangeDN $entry.Address))
	$data += $newEntry
  } 
catch { 
       if ($error[0].exception -match "null-valued")
                   {
                    $notFound++
                    write-host "Could'nt Catch Display Name From NDR." -f Red 
                    Write-host "Error Exception:" $error[0].Exception.Message
                    ""
                   }
              else {
                    $notFound++
                    write-host "Recipient Not Found or not in Correct Format" -f yellow -n 
		                ""     
                   }
		    Add-content -value $error[0].Exception  -path $log 
		}
}

function getcuX500 
       {
        param ($id)
        (Get-Recipient $id).EmailAddresses | ? {$_.Prefix -match "X500"} | % {$_.ProxyAddressString}
       }

if ($edit) {
    Add-content -value ""  -path $log
    Add-content -value "EDITED USERS"  -path $log
    Add-content -value ""  -path $log
    foreach ($u in $data)    
        {
		     write-host "checking if X500 Already Exist " -f cyan
		     $currentX500=getcuX500 -id $u.User
         if ($currentX500 -like $u.x500)
            {
             $skipped++
             Write-Host "X500 Already Exist for:"$u.User -f green
             Add-content -value "X500 address already exist user Modification was skipped: $($u.User)"  -path $log
            } 
        else {	
             $fixed++
             Write-host "Fixing:" $u.User
		         Write-host "X500Add:" $u.X500
		         Addx500 -user ((Get-Recipient $u.user).SamAccountName) -X500 $u.X500
                        
			       Add-content -value "`rUser Fixed:$($u.User)"  -path $log
             Add-content -value "X500 Added:$($u.X500)"  -path $log
             Add-content -value ""  -path $log	
		         }
	     }
    MoveFixedItems $items
}

if ($checkX500) 
{
  sleep 60
  Add-content -value ""  -path $log
  Add-content -value "POST X500 FIX CHECK - Addresses foud on Recipient"  -path $log
  Add-content -value ""  -path $log
  foreach ($u in $data) 
    {
     ""
	   write-host "Below are the currently Stampled X500 for:"$u.User -f Yellow
	   $x=(Get-Recipient $u.User).EmailAddresses | ? {$_.Prefix -match "X500"} | % {$_.ProxyAddressString}
	   if ($x -eq $null) 
              {
            	Write-Host "No X500 Address found in Recipient ProxyAddresses List"-f cyan
    		      } 
        else {
              $x
              Add-content -value "`rRecipient PrimarySMTP:$($u.user)"  -path $log
              Add-content -value "`r$x"  -path $log
              Add-content -value ""  -path $log
           	 }
    }
}

Add-content -value "****************************************************"  -path $log
Add-content -value "Number of NDR Email Received:$($fiItems.Items.count)"  -path $log
Add-content -value "`rNumber of Recipients found to be fixed:$($file.count)"  -path $log
Add-content -value "`rRecipient Fixed:$fixed"  -path $log
Add-content -value "`rRecipient skipped:$skipped"  -path $log
Add-content -value "`rRecipient matching multiple Entry:$multi "  -path $log
Add-content -value "`rUser Not Found:$notFound"  -path $log
Add-content -value "****************************************************"  -path $log

if ($sendReport) {
write-host "Sendng Report" -f Green
$to="sunil.chauhan@xyz.com"
$from="sunil.chauhan@xyz.com"
$subject="X500 Auto Fix Report: Manual Intervention needed"
$subject2="X500 Auto Fix Report"
$SMTPSRV="SMTPSRV"

#Sending Emails Notification
if ($multi) 
    {
    Send-MailMessage -to $to  -From $from -Subject $subject -Body $(cat $log | out-string) -SmtpServer $SMTPSrv
    }
else 
   {
    Send-MailMessage -to $to -From $from -Subject $subject2 -Body $(cat $log | out-string) -SmtpServer $smtpSrv
   }
"Done"
}
