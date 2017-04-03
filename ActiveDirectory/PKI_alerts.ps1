  #All settings changeable to suit your needs
  ##################################Please also see the IF STATEMENT below for filtering
  $NumberOfMonths = 8	#Number of month to look ahead
  #The recipient of the email alert. It supports the format of "Name <name@domain.com>"
  $EmailTarget = "<Enter the target email address or distribution list>" 
  #The sender of the email alert. It supports the format of "Name <name@domain.com>". Best practice to be sent from a real email address on your mail server
  $EmailFrom = "<Enter the from email address>" 
  #Mail server that allows smtp routing. This was used via an Exchange 2010 environment during the creation of this scsript
  $EmailSMTPServer = "<Enter the FQDN of your mail server>" 
  ##################################
  
  #Traffic light system for email colouring and makes the email easy to view
  $AmberColour = "FE8E00"
  $RedColour = "FE0000"
  $GreenColour = "00DD0D"
  
  #Used to colour the months number in the email body
  $BlueColour = "003EFE"
  
  #For customisation the email body was written in HTML code as a string format was proving difficult to adjust
  $Body = "<HTML><HEAD><META http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /><TITLE></TITLE></HEAD>"
  $Body += "<BODY bgcolor=""#FFFFFF"" style=""font-size: Small; font-family: TAHOMA; color: #000000""><P>"
  
  #Subject of the email sent
  $EmailSubject = "Certificate Expiration List"
  
  #The earliest date range in the check. I.E. any certificate about to expire after the date $After
  $After = Get-Date -Format d/MM/yyyy
  #The latest date range in the check. I.E. any certificate that will to expire before the date $Before which relies of the $NumberOfMonths set earlier
  $Before =(Get-Date).AddMonths($NumberOfMonths) | Get-Date -Format d/MM/yyyy
  #Restriction set for the date range for the cert collection
  $Restrict = "NotAfter>=$After,NotAfter<=$Before"
  #$CertText is a variable to work with in memory which collects the cert data using the inbuilt windows Certutil.exe
  #The values are then piped (|sls -Pattern "Requester Name:","Issued Common Name:","Certificate Expiration Date:") into the SLS powershell function to clean the data received up for the email body and check
  $CertText = certutil -view -restrict $Restrict -out "RequesterName,CommonName,Certificate Expiration Date" |Select-String -Pattern "Requester Name:","Issued Common Name:","Certificate Expiration Date:"
  
  #If the received data finds a certificate within the date range then the script will continue to work else it will end because no certificates are about to expire
  If ($CertText.Length -ge 3)
  {
    #Builing the first pieces of the body
    $Body += "<b>Dear Certificate Administrator</b><br>"
    $Body+="<br>" 
    $Body+="<b>Below is a list of Certificates due to expire in the next <b><font color=$BlueColour>$NumberOfMonths</b></font> months</b><br>"
    $Body+="<br>" 
    $Body+="<b><font color=$GreenColour>Green: Expiration date is further than 80 days away</font></b><br>"
    $Body+="<b><font color=$AmberColour>Amber: Expiration date is within the next 80 days</font></b><br>"
    $Body+="<b><font color=$RedColour>Red: Expiration date is within the next 40 days</font></b><br>"
    $Body+="<br>" 
    
    #This FOR LOOP works through each certificate found and colours them accordingly for the traffic light system
    For($intCounter=0; $intCounter -lt $CertText.length; $intCounter+=3)
    {
      #Filter out unwanted certificates
      #This is especially useful if autenrollment has been set up for devices on the network and there are potentially thousands of certificates
      #Filtering the certificates down will also improve the speed of the script
      #The filter is based on the line that contains "Requester Name:"
      #Example of a single filter to replace the top line of the IF STATEMENT
      #							If (!($CertText[$intCounter] -like "*domainname.com\Desktop*"))
      #Example of a multiple filter
      #							If (!($CertText[$intCounter] -like "*domainname.com\Desktop*" -or $CertText[$intCounter] -like "*domainname.com\tablet*"))
      ##################################
      If (!($CertText[$intCounter] -like "*domainname.com\Desktop*" -or $CertText[$intCounter] -like "*domainname.com\tablet*"))                                
      {
      ##################################
        #Get the expiration date from the certificate being worked on to determine the colour setting required
        $DateCheck = ($CertText[$intCounter+2]).tostring()
        $DateCheck = $DateCheck.Substring(31,$DateCheck.Length-31)
        $Datecheck = get-date($DateCheck)
         
        $StartDate=(GET-DATE)
        $TextColour = ""
        #The New-timespan powershell function normalises the date maths into an easily usable format instead and needing to do date maths
        $NumberofDays = (New-timespan -start $StartDate -End $DateCheck).Days
        
        If($NumberofDays -lt 40)
        {
          #Set colour to RED if the cert will expire within 40 days
          $TextColour = "$RedColour"
        }
        elseif ($NumberofDays -lt 80)
        {
          #Set colour to AMBER if the cert will expire within 40 to 80 day period
          $TextColour = "$AmberColour"
        }
        else
        {
          #Set colour to GREEN if the cert will expire more than 80 days away
          $TextColour = "$GreenColour"
        }
        #Add the certificate information to the body of the email
        $Text = ($CertText[$intCounter]).tostring()
        $Body += "$Text<br>"
        $Text = ($CertText[$intCounter+1]).tostring()
        $Body+="$Text<br>"
        $Text = ($CertText[$intCounter+2]).tostring()
        #Colour of the expiration date of the certificate is set accordingly from the IF STATEMENT above
        $Body+="<font color=$TextColour>$Text</font><br>"
        $Body+="<br>" 
      }
    }
    #Send your email alert
    Send-MailMessage -to $EmailTarget -from $EmailFrom -subject $EmailSubject -smtpserver $EmailSMTPServer -Body $Body -BodyAsHtml
  }
