## Define UPN of the Account that has impersonation rights
function getcred {
$CredFile = "D:\sunil\myscripts\X500fixews\CredFile.xml"
$CredFromFile = Import-Clixml $CredFile
$CredFromFile.password = $CredFromFile.Password | ConvertTo-SecureString
$Cred=New-Object system.Management.Automation.PSCredential($CredFromFile.username, $CredFromFile.password)
$Global:Cred = $Cred
}

function Get-X500Errorusers {
param ($items=200)
getcred

$pass = $cred.GetNetworkCredential().Password
$AccountWithImpersonationRights = "AdminAccount@xyzDomain.com"
$Folder = "\LegecyDN"

#Define the SMTP Address of the mailbox to impersonate
$MailboxToImpersonate = "trapMailbox@xyzDomain.com"

$dllpath = "D:\sunil\Microsoft.Exchange.WebServices.dll"
Import-Module $dllpath

# Set Exchange Version
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2

#Create Exchange Service Object
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)

#Get valid Credentials using UPN for the ID that is used to impersonate mailbox
#$Service.UseDefaultCredentials = $true 

$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $AccountWithImpersonationRights, $pass

# Set the URL of the CAS (Client Access Server)
$uri=[system.URI] "https://mail.mydmain.com/ews/exchange.asmx" 
$service.url = $uri

#$service.AutodiscoverUrl($AccountWithImpersonationRights ,{$true})
Write-Host 'Using ' $AccountWithImpersonationRights ' to Impersonate ' $MailboxToImpersonate
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress,$MailboxToImpersonate );

#Connect to the Inbox and display basic statistics
$MailboxRootid= new-object Microsoft.Exchange.WebServices.Data.FolderId `
([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$ImpersonatedMailboxName)
$MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$MailboxRootid)
$fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(100)
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
$findFolderResults = $MailboxRoot.FindFolders($fvFolderView)
$legacyDNFolder=$findFolderResults | ? {$_.DisplayName -match "LegacyDN"}
$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet `
([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
#$psPropset.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::FirstClassProperties

#([Microsoft.Exchange.WebServices.Data.BasePropertySet]::HTML)
$psPropset.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::HTML

#Define ItemView to retrive just 10 Items  
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView($items)  
$Global:fiItems = $service.FindItems($legacyDNFolder.Id,$ivItemView) 

try { 
[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)
Write-host 'Email NDR found:' $fiItems.Items.count
} catch {Write-host "No Items Found in the Folder."}

$new= $fiitems | % {$_.ToRecipients | ? {$_.Address -match "IMCEAEX-"}}
$new | select -Unique Name, Address

#$trgFolder=$findFolderResults | ? {$_.DisplayName -match "fixed"}
#foreach ($i in $fiItems) {$i.move($trgFolder.id) | Out-null}
}
