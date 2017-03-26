$domain = 'domain\Username'
$pass = ConvertTo-SecureString -String 'somepassw@rd123' -AsPlainText -Force
$creds = New-Object System.Management.Automation.pscredential -ArgumentList $domain, $pass
$s = New-PSSession -ConfigurationName Microsoft.Exchange `
                   -ConnectionUri http://ExchangeServerName/PowerShell/ `
                   -Authentication Kerberos `
                   -Credential $creds
Import-PSSession $s

