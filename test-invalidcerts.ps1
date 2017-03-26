function Test-InvalidCert {
  <#
  .SYNOPSIS
    Checks root and intermediate certificate stores for certificates in the wrong store & moves them to the correct store.

  .DESCRIPTION
    Checks root and intermediate certificate stores for certificates in the wrong store & moves them to the correct store. Certificates in the wrong store can be very problematic for Lync Server and Skype for Business Server.

  .NOTES
    Version               : 1.2
    Wish list             : 
    Rights Required       : Local administrator on server
    Sched Task Required   : No
    Lync/Skype4B Version  : N/A
    Author/Copyright      : © Pat Richard, Office Servers and Services (Skype for Business) MVP - All Rights Reserved
    Email/Blog/Twitter    : pat@innervation.com  https://ucunleashed.com  @patrichard
    Donations             : https://www.paypal.me/PatRichard
    Dedicated Post        : https://ucunleashed.com/3903
    Disclaimer            : You running this script means you won't blame author(s) if this breaks your stuff. This script is
                            provided AS IS without warranty of any kind. Author(s) disclaim all implied warranties including,
                            without limitation, any implied warranties of merchantability or of fitness for a particular
                            purpose. The entire risk arising out of the use or performance of the sample scripts and
                            documentation remains with you. In no event shall author(s) be liable for any damages whatsoever
                            (including, without limitation, damages for loss of business profits, business interruption,
                            loss of business information, or other pecuniary loss) arising out of the use of or inability
                            to use the script or documentation. Neither this script, nor any part of it other than those
                            parts that are explicitly copied from others, may be republished without author(s) express written 
                            permission.
    Acknowledgements      :
    Assumptions           : ExecutionPolicy of AllSigned (recommended), RemoteSigned, or Unrestricted (not recommended)
    Limitations           :
    Known issues          : None yet, but I'm sure you'll find some!

  .LINK
    <blockquote data-secret="XEL5rIXalz" class="wp-embedded-content"><a href="https://www.ucunleashed.com/3903">Function: Test-InvalidCerts – Ensuring Certificates Are In The Correct Certificate Store</a></blockquote><iframe class="wp-embedded-content" sandbox="allow-scripts" security="restricted" style="position: absolute; clip: rect(1px, 1px, 1px, 1px);" src="https://www.ucunleashed.com/3903/embed#?secret=XEL5rIXalz" data-secret="XEL5rIXalz" width="500" height="282" title="“Function: Test-InvalidCerts – Ensuring Certificates Are In The Correct Certificate Store” — UC Unleashed" frameborder="0" marginwidth="0" marginheight="0" scrolling="no"></iframe>

  .EXAMPLE
    .\Test-InvalidCerts.ps1

    Description
    -----------
    Checks root and intermediate store for certs in the wrong store & moves them to the proper store.

  .INPUTS
    None. You cannot pipe objects to this script.
  #>
  [CmdletBinding(SupportsPaging)]
  param(
  ) # end of param block
  process{
    Write-Verbose -Message 'Checking for improper certs in root store'
    $InvalidCertsInRoot = Get-Childitem -Path cert:\LocalMachine\root -Recurse | Where-Object {$_.Issuer -ne $_.Subject}    

    if ($InvalidCertsInRoot){        
      Write-Verbose -Message ('{0} invalid certificates detected in Root Certificate Store' -f $InvalidCertsInRoot.count)
      ForEach ($cert in $InvalidCertsInRoot){
        Write-Verbose -Message ('Moving `"{0}`" to intermediate certificate store' -f $cert.subject)
        Move-Item -Path $cert.PSPath -Destination Cert:\LocalMachine\CA
      }
    }else{
      Write-Verbose -Message 'No invalid certs found in Root Certificate Store'
    }
    $InvalidCertsInInt = Get-Childitem -Path cert:\LocalMachine\Ca -Recurse | Where-Object {$_.Issuer -eq $_.Subject}  
    if ($InvalidCertsInInt){
      Write-Verbose -Message ('{0} invalid certificates detected in Intermediate Certificate Store' -f $InvalidCertsInInt.count)
      ForEach ($cert in $InvalidCertsInInt){
        Write-Verbose -Message ('Moving `"{0}`" to root certificate store' -f $cert.subject)
        Move-Item -Path $cert.PSPath -Destination Cert:\LocalMachine\Root
      }
    }else{
      Write-Verbose -Message 'No invalid certs found in Intermediate Certificate Store'
    }
  }
} # end function Test-InvalidCert
