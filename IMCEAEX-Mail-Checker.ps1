#Change Drive location for find-Email-From-Folder-EWS.ps1 script
. "D:\sunil\myscripts\X500fixEws\find-Email-From-Folder-EWS.ps1"

#-----Change Email Notification Details Below-------------------------------------
##################################################################################
$To="sunil@xyzdomain.com"
$From="sunil@xyzdomain.com"
$SmtpSrv="SmtpSrv.fqdn.domain.com"
$subject="X500 Auto Fix Report-ALL OK!! No New NDR Received"
##################################################################################

#Change Logs Drive Path
$log="D:\sunil\myscripts\X500fixews\log.txt"

$file = Get-X500Errorusers
if ($file.count -or $file.Name)
        	{
           Write-host "Recipient Found for Fix:"$file.Count
	         . "D:\sunil\myscripts\X500fixews\x500-fix-from-ews.ps1" -sendReport true -edit true -checkX500 true
	        }
 else     {
           "No Recipient found for processing"
           Clear-Content $log
           Add-content "X500 Auto Fix Report-ALL OK!! No New NDR Received:$(GEt-date -Format dd/MM/yy-HH:mm)" -path $log
           Send-MailMessage -to $to -From $from -Subject $subject -Body $(cat $log | out-string) -SmtpServer $SmtpSrv
	       }
