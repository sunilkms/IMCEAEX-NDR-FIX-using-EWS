#Export Credentions
Function ExportCred 
                {
                 #Export Credentials to file
                 $cred = Get-Credential 
                 $CredFile = "D:\sunil\myscripts\X500fixews\CredFile.xml"
                 $CredToFile = $cred | Select-Object *
                 $CredToFile.password = $CredToFile.Password | ConvertFrom-SecureString
                 $CredToFile | Export-Clixml $CredFile
                }
#Get Cred
function getcred 
               {
                $CredFile = "D:\sunil\myscripts\X500fixews\CredFile.xml"
                $CredFromFile = Import-Clixml $CredFile
                $CredFromFile.password = $CredFromFile.Password | ConvertTo-SecureString
                $Cred=New-Object system.Management.Automation.PSCredential($CredFromFile.username, $CredFromFile.password)
                $Global:Cred = $Cred
               }
