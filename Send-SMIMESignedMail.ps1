###############################################################################
# Send an Email with an S/MIME Signature
# 21.03.2021 - Version 1.0 - Andres Bohren
###############################################################################

<#
.SYNOPSIS 
	This Script sends an S/MIME Signed Email. You need to have a PFX of your Certificate with a Password.
.LINK 
	https://blog.icewolf.ch
.EXAMPLE 
	.\Send-SMIMESignedMail.ps1 -MailFromPFXFile E:\a.bohren@icewolf.ch.pfx -MailFromPFXPassword "MyPFXPassword" -MailTo recipient@domain.tld -Subject "Test" -Body "Just a Test" -SMTPServer 172.21.175.61 -SMTPPort 25 
	.\Send-SMIMESignedMail.ps1 -MailFromPFXFile E:\a.bohren@icewolf.ch.pfx -MailFromPFXPassword "MyPFXPassword" -MailTo recipient@domain.tld -Subject "Test" -Body "Just a Test" -SMTPServer 172.21.175.61 -SMTPPort 25 -SMTPUsername "YourSMTPUsername" -SMTPPassword "YourSMTPPassword"

.DESCRIPTION
	This Script is based on smtp.smime.lib from rob kalmar
    https://www.powershellgallery.com/packages/smtp.smime.lib/1.0.4/Content/smtp.smime.lib.ps1
#> 


###############################################################################
#Script Input Parameters
###############################################################################
PARAM
(
	#[parameter(Mandatory=$true)][string]$MailFrom,
	[parameter(Mandatory=$true)][string]$MailFromPFXFile,
	[parameter(Mandatory=$true)][string]$MailFromPFXPassword,
	[parameter(Mandatory=$true)][string]$MailTo,
	[parameter(Mandatory=$true)][string]$Subject, 
	[parameter(Mandatory=$true)][string]$Body, 
	[parameter(Mandatory=$false)][string]$Attachment = $Null,
	[parameter(Mandatory=$true)][string]$SMTPServer,
	[parameter(Mandatory=$true)][string]$SMTPPort = "25",
	[parameter(Mandatory=$false)][string]$SMTPUsername = "",
	[parameter(Mandatory=$false)][string]$SMTPPassword = ""
)




###############################################################################
# Function Test-IsAscii
###############################################################################
Function Test-IsAscii
{
  <#
      .SYNOPSIS
      Tests if the input string only contains ASCII values.
      .DESCRIPTION
      Tests if the input string only contains ASCII values.
      Returns true if this is the case, otherwise it returns false.
      .PARAMETER $text
      A string to check.
      .EXAMPLE
      Test-IsAscii -text 'This string contains only ASCII'
  #>
  param
  (
    [parameter(Mandatory=$true)][string]$text
  )
  [byte[]]$utf8text = [Text.Encoding]::UTF8.GetBytes($text)
  [string]$asciitext = [Text.Encoding]::ASCII.GetString($utf8text)
  If ($text -eq $asciitext)
  {
    Return $True
  }
  Else
  {
    Return $False
  }
}

###############################################################################
# New-MailserverSettings
###############################################################################
function New-MailserverSettings
{
  <#
      .SYNOPSIS
      Create a new PSObject which sets the properties of a mail server
      .DESCRIPTION
      This function creates a new PSObject and sets the properties of a mail server object.
      The function returns a PSObject.
      .EXAMPLE
      Call this function with a string representing the fully qualified domain name of the mail server
      New-MailserverSettings "mail.example.com"
      .PARAMETER $FQDN
      A string containing fully qualified domain name of the mail server
      .PARAMETER $Output
      An integer indicating whether to send to mail to the mail server and/or to disk.
      Output = 0 -> save to eml file
      Output = 1 -> send to smtp server
      Output = 2 -> save to eml file and send to smtp server
      .PARAMETER $Port
      A string containing the port of the mail server to connect to, default to "25"
      .PARAMETER $STARTTLS
      A boolean indicating whether to use STARTTLS or not when connecting to the mail server, defaults to $False
      .PARAMETER $Username
      A string containing the username used for authentication on the mail server, defaults to ""
      .PARAMETER $Password
      A secure string containing the password used for authentication on the mail server, defaults to 1
  #>
  param
  (
    [parameter(Mandatory=$true)][string]$FQDN,
    [ValidateSet(0,1,2)][int]$Output = 1,
    [string]$Port = '25',
    [string]$STARTTLS = $False,
    [string]$VerifyServerCertificate = $True,
    [string]$Username = '',
    [string]$Password = ''
  )

  $MailServer = new-object -TypeName PSObject
  $MailServer | Add-Member -Name FQDN -Value $FQDN -MemberType NoteProperty
  $MailServer | Add-Member -Name Output -Value $Output -MemberType NoteProperty
  $MailServer | Add-Member -Name Port -Value $Port -MemberType NoteProperty
  $MailServer | Add-Member -Name STARTTLS -Value $STARTTLS -MemberType NoteProperty
  $MailServer | Add-Member -Name VerifyServerCertificate -Value $VerifyServerCertificate -MemberType NoteProperty
  $MailServer | Add-Member -Name Username -Value $Username -MemberType NoteProperty
  $MailServer | Add-Member -Name Password -Value $Password -MemberType NoteProperty
  Return $MailServer
}

###############################################################################
# Get-MimeType
###############################################################################
function Get-MimeType
{
  <#
      .SYNOPSIS
      Retrieves the Mime Type of a file extension from the Registry
      .DESCRIPTION
      This function retrieves the Mime Type of an file extension from the Windows Registry
      The function returns a string of the Mime Type.
      .EXAMPLE
      Call this function with a string representing the extension of a file
      Get-MimeType ".zip"
      .PARAMETER $Extension
      A string containing the extension of a file
  #>
  [CmdletBinding()]
  param
  (
    [string]$Extension = $null
  )

  $MimeType = 'application/octet-stream'

  $null = New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT

  Try
  {
    $Exists = Test-Path -Path "HKCR:$Extension" -ErrorAction Stop
    If ($Exists)
    {
      $Item = Get-ItemProperty -Path HKCR:$Extension -ErrorAction Stop
      $ItemContentType = $Item.'Content Type'
      If (($ItemContentType -eq $null) -or ($ItemContentType -eq ''))
      {
        $ItemContentType = 'application/octet-stream'
      }
      $MimeType = $ItemContentType
    }
  }
  Catch
  {
    $MimeType = 'application/octet-stream'
  }
  Return $MimeType
}

###############################################################################
# New-UTF8EncodedFileName
###############################################################################
function New-UTF8EncodedFileName
{
  <#
      .SYNOPSIS
      Given a filename string the function encodes it into UTF8.
      .DESCRIPTION
      Given a filename string the function encodes it into UTF8 for use in
      the Content-Type or Content-Disposition name/filename paramters
      .PARAMETER $FileName
      The filename to encode into UTF8
      .EXAMPLE
      New-UTF8EncodedFileName -FileName 'example.txt'
      Encodes the filename example.txt to UTF8.
  #>
  param
  (
    [parameter(Mandatory=$true)][string]$FileName
  )
    
  [String]$Header = ' =?UTF-8?B?'
  [String]$Footer = '?='
  [String]$EncodedB64FileName = $null
  [int]$ChunkSize = 60 #needs to be multiple of 4, otherwise Outlook does not like it.
  [String]$B64FileName = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($FileName))
  [int]$SizeB64FileName = $B64FileName.Length

  $UTF8EncodedFileName = New-Object -TypeName System.Text.StringBuilder 

  [int]$Position = 0
  While ($Position -le ($SizeB64FileName - $ChunkSize))
  {
    $null = $UTF8EncodedFileName.AppendLine()
    $EncodedB64FileName = $Header + $B64FileName.Substring($Position,$ChunkSize) + $Footer
    $null = $UTF8EncodedFileName.Append("$EncodedB64FileName")
    $Position = $Position + $ChunkSize   
  }
  If (($SizeB64FileName-$Position) -gt 0)
  {
    $null = $UTF8EncodedFileName.AppendLine()
    $EncodedB64FileName = $Header + $B64FileName.Substring($Position,$SizeB64FileName-$Position) + $Footer
    $null = $UTF8EncodedFileName.Append("$EncodedB64FileName") 
  }
    
  Return $UTF8EncodedFileName
}

###############################################################################
# New-HeaderMimePart
###############################################################################
function New-HeaderMimePart
{
  <#
      .SYNOPSIS
      Creates a new header for a Mime part
      .DESCRIPTION
      This function creates a new header for Mime part of a message
      The function returns a string
      .EXAMPLE
      Call this function with a string representing the Mime Type and file name of the Mime part to which the header will be added
      New-HeaderMimePart "application/octet-stream" "filename.zip"
      .PARAMETER $MimeType
      A string containing the Mime Type of a file
      .PARAMETER $MimeType
      A string containing the file name of a file
  #>
  param
  (
    [parameter(Mandatory=$true)][string]$MimeType,
    [parameter(Mandatory=$true)][string]$FileName
  )

  $MimePartHeader = New-Object -TypeName System.Text.StringBuilder 
  
  If ((Test-IsAscii -text $FileName) -and ($FileName.Length -le 64))
  {
    $null = $MimePartHeader.AppendLine("Content-Type: $MimeType; charset=`"utf-8`";")
    $null = $MimePartHeader.AppendLine(" name=`"$FileName`"")
    $null = $MimePartHeader.AppendLine('Content-Transfer-Encoding: base64')
    $null = $MimePartHeader.AppendLine('Content-Disposition: attachment;')
    $null = $MimePartHeader.AppendLine(" filename=`"$FileName`"")
    $null = $MimePartHeader.AppendLine() 
  }
  Else
  {
    [String]$UTF8EncodedFileName = New-UTF8EncodedFileName -FileName $FileName

    $null = $MimePartHeader.Append("Content-Type: $MimeType; charset=`"utf-8`"; name=`"")
    $null = $MimePartHeader.Append("$UTF8EncodedFileName")
    $null = $MimePartHeader.AppendLine('"')

    $null = $MimePartHeader.AppendLine('Content-Transfer-Encoding: base64')
    
    $null = $MimePartHeader.Append('Content-Disposition: attachment; filename="')
    $null = $MimePartHeader.Append("$UTF8EncodedFileName")
    $null = $MimePartHeader.AppendLine('"')

    $null = $MimePartHeader.AppendLine()
  }
    
  Return $MimePartHeader
}

###############################################################################
# New-MimePart
###############################################################################
function New-MimePart
{
  <#
      .SYNOPSIS
      Creates a new Mime part for a Mime Message
      .DESCRIPTION
      This function creates a new Mime part for a Mime message
      The function returns a string
      .EXAMPLE
      Call this function with an object representing a file
      New-MimePart file
      .PARAMETER $File
      An object representing a file
  #>
  param
  (
    [parameter(Mandatory=$true)][IO.FileInfo]$File
  )

  $Extension = $null
  $Extension = [IO.Path]::GetExtension($File)
  $MimeType = Get-MimeType -Extension $Extension

  $FileName = [IO.Path]::GetFileName($File)
  $Header = New-HeaderMimePart -MimeType $MimeType -FileName $FileName

  $MimePart = New-Object -TypeName System.Text.StringBuilder
  $null = $MimePart.Append($Header)

  [Byte[]] $BinaryData = [IO.File]::ReadAllBytes($File)
  [string] $Base64Value = [Convert]::ToBase64String($BinaryData, [Base64FormattingOptions]::InsertLineBreaks)

  $null = $MimePart.Append($Base64Value)
  $null = $MimePart.AppendLine()

  Return $MimePart
}

###############################################################################
# New-MimeMessage
###############################################################################
function New-MimeMessage
{
  <#
      .SYNOPSIS
      Create a Mime Message
      .DESCRIPTION
      This function creates multipart/mixed a Mime message
      The function returns a string
      .EXAMPLE
      Call this function with a string representing the body of the Mime Message and an object representing a file
      New-MimeMessage "body text" file
      .PARAMETER $File
      An string representing the body text of the message
      .PARAMETER $File
      An object representing a file
  #>
  [CmdletBinding()]
  param
  (
    [string]$Body = $null,
    [IO.FileInfo[]]$FileList = $null
  )

  $MIMEMessage = New-Object -TypeName System.Text.StringBuilder 

  If ($FileList -eq $null)
  {
    $null = $MIMEMessage.AppendLine("Content-Type: text/plain; charset=`"utf-8`"")
    $null = $MIMEMessage.AppendLine('Content-Transfer-Encoding: base64')

    $null = $MIMEMessage.AppendLine()
    
    [Byte[]] $BodyBytes = [Text.Encoding]::UTF8.GetBytes($Body)
    [string] $BodyB64 = [Convert]::ToBase64String($BodyBytes, [Base64FormattingOptions]::InsertLineBreaks)
    
    $null = $MIMEMessage.Append($BodyB64)
    $null = $MIMEMessage.AppendLine()
  }
  Else
  {
    [String]$Boundary = 'boundary-' + [guid]::NewGuid()
        
    $null = $MIMEMessage.AppendLine('Content-Type: multipart/mixed;') 
    $null = $MIMEMessage.AppendLine(" boundary=$Boundary") 
    $null = $MIMEMessage.AppendLine()
    $null = $MIMEMessage.AppendLine('This is a multi-part message in MIME format.')
    $null = $MIMEMessage.AppendLine("--$Boundary")
    $null = $MIMEMessage.AppendLine("Content-Type: text/plain; charset=`"utf-8`"")
    $null = $MIMEMessage.AppendLine('Content-Transfer-Encoding: base64')
    $null = $MIMEMessage.AppendLine()
    
    [Byte[]] $BodyBytes = [Text.Encoding]::UTF8.GetBytes($Body)
    [string] $BodyB64 = [Convert]::ToBase64String($BodyBytes, [Base64FormattingOptions]::InsertLineBreaks)
    
    $null = $MIMEMessage.Append($BodyB64)
    $null = $MIMEMessage.AppendLine()
    $null = $MIMEMessage.Append("--$Boundary")
   
    $File = $null

    Foreach ($File in $FileList)
    {
      $null = $MIMEMessage.AppendLine()  
      $Part = New-MimePart -File $File

      $null = $MIMEMessage.Append($Part)
      $null = $MIMEMessage.Append("--$Boundary")
    }
        
    $null = $MIMEMessage.AppendLine('--')
  }

  Return $MIMEMessage
}

###############################################################################
# New-SignedContent
###############################################################################
function New-SignedContent
{
  <#
      .SYNOPSIS
      Returns a signed byte array of a string using given certificate
      .DESCRIPTION
      This function signs a given string using the private key of a given certificate.
      The function returns an array of signed bytes.
      .EXAMPLE
      Call this function with a string to be signed and a certificate containing a private key
      New-SignedContent "Example string" certificate
      .PARAMETER $MIMEMessage
      A string containing the message to be signed
      .PARAMETER $Certificate
      A certificate which will be used to sign the Mime Message.
      .PARAMETER $EndCertOnly
      Set to $True to avoid issues with building up the certificate chain to the root certificate, defaults to $False
      .PARAMETER $Detached
      If set to $True the signatue will be detached so non smime capable clients can still read the message.
      .PARAMETER $Boundary
      Defines the boudary used in the multipart MIME message.
      .PARAMETER $DigestAlgorithm
      A string specifying which digest algorithm to use, defaults to "1.3.14.3.2.26" (SHA-1)
      "1.2.840.113549.2.5" #md5
      "1.3.14.3.2.26" #SHA-1
      "2.16.840.1.101.3.4.2.1" #SHA-256
      "2.16.840.1.101.3.4.2.2" #SHA-384
      "2.16.840.1.101.3.4.2.3" #SHA-512
      .PARAMETER $EnhancedProtection
      Turn ON or OFF enhanced protection. When ON the recipient and subject will be added as signed attributes to the signature.
      This enables you to check if they were changed in transit.
      Note: No mail client actually checks for this, so you need to program something yourself.
      .PARAMETER $Recipient
      The recipient the include as a signed attribute.
      .PARAMETER $Subject
      The subject to include as a signed attribute.
  #>
  param
  (
    [parameter(Mandatory=$true)][string]$MIMEMessage,
    [parameter(Mandatory=$true)][Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
    [bool]$EndCertOnly = $False,
    [bool]$Detached = $True,
    [string]$Boundary = 'boundary-1',
    [string]$DigestAlgorithm = '1.3.14.3.2.26',
    [bool]$EnhancedProtection = $True,
    [string]$Recipient = $null,
    [string]$Subject = $null        
  )
 
  Add-Type -AssemblyName System.Security
    
  [Byte[]] $BodyBytes = [Text.Encoding]::ASCII.GetBytes($MIMEMessage.ToString())
  $ContentInfo = New-Object -TypeName System.Security.Cryptography.Pkcs.ContentInfo -ArgumentList (,$BodyBytes) 
  $SignedCMS = New-Object -TypeName System.Security.Cryptography.Pkcs.SignedCms -ArgumentList $ContentInfo, $Detached

  $CMSSigner = New-Object -TypeName System.Security.Cryptography.Pkcs.CmsSigner -ArgumentList $Certificate 
  If ($EndCertOnly)
  {
    $CMSSigner.IncludeOption = [Security.Cryptography.X509Certificates.X509IncludeOption]::EndCertOnly
  }
  Else
  {
    $CMSSigner.IncludeOption = [Security.Cryptography.X509Certificates.X509IncludeOption]::WholeChain
  }
  $CMSSigner.DigestAlgorithm = $DigestAlgorithm

  $TimeStamp=Get-Date
  $Pkcs9SigningTime = New-Object -TypeName System.Security.Cryptography.Pkcs.Pkcs9SigningTime -ArgumentList ($TimeStamp)
  $null = $CMSSigner.SignedAttributes.Add($Pkcs9SigningTime)

  If ($EnhancedProtection)
  {
    If ($Recipient)
    {
      $documentNameAttribute = New-Object -TypeName System.Security.Cryptography.Pkcs.Pkcs9DocumentName -ArgumentList ("$Recipient")
      $null = $CMSSigner.SignedAttributes.Add($documentNameAttribute)
    }
    If ($Subject)
    {
      $documentDescriptionAttribute = New-Object -TypeName System.Security.Cryptography.Pkcs.Pkcs9DocumentDescription -ArgumentList ("$Subject")
      $null = $CMSSigner.SignedAttributes.Add($documentDescriptionAttribute)
    }
  }
    
  $SignedCMS.ComputeSignature($CMSSigner)
  [Byte[]] $SignedBytes = $SignedCMS.Encode()

  If ($Detached -eq $True)
  {
    $MultipartSigned = New-Object -TypeName System.Text.StringBuilder
        
    $null = $MultipartSigned.AppendLine("--$Boundary")
    $null = $MultipartSigned.AppendLine("$MIMEMessage")
    $null = $MultipartSigned.AppendLine("--$Boundary")
    $null = $MultipartSigned.AppendLine('Content-Type: application/pkcs7-signature; name="smime.p7s"')
    $null = $MultipartSigned.AppendLine('Content-Disposition: attachment; filename="smime.p7s"')
    $null = $MultipartSigned.AppendLine('Content-Transfer-Encoding: base64')
    $null = $MultipartSigned.AppendLine()

    [String]$Signature = [Convert]::ToBase64String($SignedBytes, [Base64FormattingOptions]::InsertLineBreaks)
    $null = $MultipartSigned.AppendLine("$Signature")

    $null = $MultipartSigned.AppendLine()
    $null = $MultipartSigned.AppendLine("--$Boundary--")
    
    [Byte[]] $SignedBytes = [Text.Encoding]::UTF8.GetBytes($MultipartSigned.ToString())
  }  

  Return $SignedBytes
}

###############################################################################
# Send-SMIMEMail
###############################################################################
function Send-SMIMEMail
{
  <#
      .SYNOPSIS
      Sends a SMIME message to a mail server
      .DESCRIPTION
      This function sends message to a mail server
      The function returns $True if the message is successfully sent otherwise $False
      .EXAMPLE
      Call this function with an object representing a mail server and an object of type System.Net.Mail.MailMessage
      Send-SMIMEMail mailserver message
      .PARAMETER $MailServer
      An object representing a mail server
      .PARAMETER $File
      An object representing an object of type System.Net.Mail.MailMessage
  #>
  param
  (
    [parameter(Mandatory=$true)][PSObject]$MailServer,
    [parameter(Mandatory=$true)][Net.Mail.MailMessage]$Message
  )

  $SendMail = $False
    
  $scriptPath = $(Split-Path -Path $script:MyInvocation.MyCommand.Path)

  $MailClient = New-Object -TypeName System.Net.Mail.SmtpClient -ArgumentList $MailServer.FQDN, $MailServer.Port
  $MailClient.Credentials = New-Object -TypeName System.Net.NetworkCredential -ArgumentList ($MailServer.Username, $MailServer.Password)

  $Delivery = $MailServer.Output
  switch ($Delivery)
  {
    1
    {
      $MailClient.DeliveryMethod = [Net.Mail.SmtpDeliveryMethod]::Network
      $MailClient.EnableSsl = [Convert]::ToBoolean($($MailServer.STARTTLS))
      
      If ([Convert]::ToBoolean($($MailServer.VerifyServerCertificate)))
      {
        [Net.ServicePointManager]::ServerCertificateValidationCallback = $null
      }
      Else
      {
        [Net.ServicePointManager]::ServerCertificateValidationCallback = { Return $True }
      }
      
      Try
      {
        $MailClient.Send($Message)
        $MailClient.Dispose()
        $SendMail = $True
      }
      Catch
      {
        $SendMail = $False
        $MailClient.Dispose()
      } 
    }

    2
    {
      $MailClient.DeliveryMethod = [Net.Mail.SmtpDeliveryMethod]::SpecifiedPickupDirectory
            
      $Exists = Test-Path -Path "$scriptPath\sent"
      if (!$Exists)
      {
        $null = New-Item -ItemType directory -Path "$scriptPath\sent"
      }
      $MailClient.PickupDirectoryLocation = "$scriptPath\sent"
      Try
      {
        $MailClient.Send($Message)
        $SendMail = $True
      }
      Catch
      {
        $SendMail = $False
      }
      $MailClient.DeliveryMethod = [Net.Mail.SmtpDeliveryMethod]::Network
      $MailClient.EnableSsl = [Convert]::ToBoolean($($MailServer.STARTTLS))
      
      If ([Convert]::ToBoolean($($MailServer.VerifyServerCertificate)))
      {
        [Net.ServicePointManager]::ServerCertificateValidationCallback = $null
      }
      Else
      {
        [Net.ServicePointManager]::ServerCertificateValidationCallback = { Return $True }
      }

      Try
      {
        $MailClient.Send($Message)
        $MailClient.Dispose()
        $SendMail = $True
      }
      Catch
      {
        $SendMail = $False
        $MailClient.Dispose()
      }
    }

    Default
    {
      $MailClient.DeliveryMethod = [Net.Mail.SmtpDeliveryMethod]::SpecifiedPickupDirectory
      $Exists = Test-Path -Path "$scriptPath\sent"
      if (!$Exists)
      {
        $null = New-Item -ItemType directory -Path "$scriptPath\sent"
      }
      $MailClient.PickupDirectoryLocation = "$scriptPath\sent"
      Try
      {
        $MailClient.Send($Message)
        $MailClient.Dispose()
        $SendMail = $True
      }
      Catch
      {
        $SendMail = $False
        $MailClient.Dispose()
      }
    }
  }
  Return $SendMail
}

###############################################################################
# Get-Oid
###############################################################################
function Get-Oid
{
  <#
      .SYNOPSIS
      Get the Oid of an Encryption or Digest algorithm.
      .DESCRIPTION
      This function return the Oid (as string) for a given encryption or digest algorithm (or $null when not known).
      .PARAMETER Algorithm
      A string of the name of the encryption or digest algorithm.
      .EXAMPLE
      Get-Oid -Algorithm 'aes128'
      This returns a string with the Oid of aes128 which is '2.16.840.1.101.3.4.1.2'
  #>

  param
  (
    [parameter(Mandatory=$true)][string]$Algorithm
  )
  
  Switch ($Algorithm)
  {
    # DigestAlgorithms
    'md5'     { $Oid = '1.2.840.113549.2.5' }
    'sha1'    { $Oid = '1.3.14.3.2.26' }
    'sha256'  { $Oid = '2.16.840.1.101.3.4.2.1' }
    'sha384'  { $Oid = '2.16.840.1.101.3.4.2.2' }
    'sha512'  { $Oid = '2.16.840.1.101.3.4.2.3' }
    
    # EncryptionAlgorithms
    'rc2'     { $Oid = '1.2.840.113549.3.2' }
    'des'     { $Oid = '1.3.14.3.2.7' }
    '3des'    { $Oid = '1.2.840.113549.3.7' }
    'aes128'  { $Oid = '2.16.840.1.101.3.4.1.2' }
    'aes192'  { $Oid = '2.16.840.1.101.3.4.1.22' }
    'aes256'  { $Oid = '2.16.840.1.101.3.4.1.42' }
    
    # Unknown
    Default   { $Oid = $null }
  }

  Return $Oid
}

###############################################################################
# New-MicAlg
###############################################################################
Function New-MicAlg
{
  <#
      .SYNOPSIS
      Given a digest algorithm returns a string for use in the micalg parameter in the Content-Type.
      .DESCRIPTION
      Given a digest algorithm returns a string for use in the micalg parameter in the Content-Type.
      This uses the specs of smime 3.1 (https://tools.ietf.org/html/rfc3851#section-3.4.3.2)
      If you want to use smime 3.2 see RFC 5751
      .PARAMETER $DigestAlgorithm
      The Digest Algorithm used for signing messages.
      .EXAMPLE
      New-MicAlg -DigestAlgorithm '1.3.14.3.2.26'
      Returns sha1
  #>


  param
  (
    [parameter(Mandatory=$true)][string]$DigestAlgorithm       
  )
    
  switch ($DigestAlgorithm)
  {
    '1.2.840.113549.2.5'     { $MicAlg = 'md5'}
    '1.3.14.3.2.26'          { $MicAlg = 'sha1'}
    '2.16.840.1.101.3.4.2.1' { $MicAlg = 'sha256'}
    '2.16.840.1.101.3.4.2.2' { $MicAlg = 'sha384'}
    '2.16.840.1.101.3.4.2.3' { $MicAlg = 'sha512'}
    Default                  { $MicAlg = 'sha1'}
  }
    
  Return $MicAlg
}


###############################################################################
# Send-SignedMail
###############################################################################
function Send-SignedMail
{
  <#
      .SYNOPSIS
      Send an encrypted mail message through a mail server
      .DESCRIPTION
      This function sends an encrypted message through a mail server using a specified certificate.
      The function returns $True when successfull otherwise $False
      .EXAMPLE
      Call this function with a mail server object, a certificate of the recipient, a string containing the sender mail address, a string with the subject text, a string containing the body text,
      an attachment representing a file.
      Send-EncryptedMail mailserver certificate "example@example.com" "subject text" "body text" file
      .PARAMETER $MailServer
      An object representing a mail server
      .PARAMETER $Recipient
      A string representing an e-mail address
      .PARAMETER $SenderCertificate
      A certificate which will be used to sign the body text and attachment.
      .PARAMETER $Subject
      A string representing the subject of a mail message
      .PARAMETER $Body
      A string representing the body text of a mail message
      .PARAMETER $Attachment
      An object representing an array of files.
      .PARAMETER $EndCertOnly
      Set to $True to avoid issues with building up the certificate chain to the root certificate, defaults to $False
      .PARAMETER $Detached
      If set to $True the signatue will be detached so non smime capable clients can still read the message.
      .PARAMETER $DigestAlgorithm
      A string specifying which digest algorithm to use, defaults to "1.3.14.3.2.26" (SHA-1)
      "1.2.840.113549.2.5" #md5
      "1.3.14.3.2.26" #SHA-1
      "2.16.840.1.101.3.4.2.1" #SHA-256
      "2.16.840.1.101.3.4.2.2" #SHA-384
      "2.16.840.1.101.3.4.2.3" #SHA-512
      .PARAMETER $EnhancedProtection
      Turn ON or OFF enhanced protection. When ON the recipient and subject will be added as signed attributes to the signature.
  #>
  param
  (
    [parameter(Mandatory=$true)][PSObject]$MailServer,
    [parameter(Mandatory=$true)][string]$Recipient,
    [parameter(Mandatory=$true)][Security.Cryptography.X509Certificates.X509Certificate2]$SenderCertificate,
    [string]$Subject = $null,
    [string]$Body = $null,
    [IO.FileInfo[]]$Attachments = $null,
    [bool]$EndCertOnly = $False,
    [bool]$Detached = $True,
    [string]$DigestAlgorithm = 'sha1',
    [bool]$EnhancedProtection = $True
  )

  $DigestAlgorithm = Get-Oid -Algorithm $DigestAlgorithm
  
  $SendSignedMail = $False
    
  [String]$Boundary = 'boundary-2-' + [guid]::NewGuid()
    
  $MIMEMessage = New-MimeMessage -Body $Body -FileList $Attachments
  $SignedBytes = New-SignedContent -MIMEMessage $MIMEMessage -Certificate $SenderCertificate -EndCertOnly $EndCertOnly -Detached $Detached -Boundary $Boundary -DigestAlgorithm $DigestAlgorithm -EnhancedProtection $EnhancedProtection -Recipient $Recipient -Subject $Subject

  $MemoryStream = New-Object -TypeName System.IO.MemoryStream -ArgumentList @(,$SignedBytes)
    
  If ($Detached -eq $True)
  {
    $MicAlg = New-MicAlg -DigestAlgorithm $DigestAlgorithm
    $ContentType = New-Object -TypeName System.Net.Mime.ContentType -ArgumentList "multipart/signed; boundary=`"$Boundary`"; protocol=`"application/pkcs7-signature`"; micalg=$MicAlg"
  }
  Else
  {
    $ContentType = New-Object -TypeName System.Net.Mime.ContentType -ArgumentList 'application/pkcs7-mime; smime-type=signed-data; name="smime.p7m"'
  }
    
  $AlternateView = New-Object -TypeName System.Net.Mail.AlternateView -ArgumentList ($MemoryStream, $ContentType) 
    
  If ($Detached -eq $True)
  {
    $AlternateView.TransferEncoding = '2'
  }
  Else
  {
    $AlternateView.TransferEncoding = '1'
  }
    
  $Sender = $SenderCertificate.GetNameInfo('EmailName', $False)

  $Message = New-Object -TypeName System.Net.Mail.MailMessage
  $Message.To.Add("<$Recipient>") 
  $Message.From = "<$Sender>"
  $Message.SubjectEncoding = [Text.Encoding]::UTF8
  $Message.Subject = $Subject
  $Message.Headers.Add('Return-Path', "<$Sender>")
  $Message.Headers.Add('X-Secured-By','Powershell S/MIME Toolkit v1.0') 
  $Message.AlternateViews.Add($AlternateView)

  $SendSignedMail = Send-SMIMEMail -MailServer $MailServer -Message $Message
  Return $SendSignedMail
}


###############################################################################
# Main Programm
###############################################################################


# enter your SMTP servername or IP address:
$FQDN = $SMTPServer
# enter the port to connect to (defaults: 25):
$Port = $SMTPPort
# use STARTTLS to switch to TLS/SSL connection:
$STARTTLS = $true
# should the server certificate be validated:
$VerifyServerCertificate = $false

# enter the username to authenticatie with:
$Username = $SMTPUsername
# enter the password used to authenticate:
$Password = $SMTPPassword

# use 0 for testing
# Output = 0 -> save to eml file
# Output = 1 -> send to smtp server
# Output = 2 -> save to eml file and send to smtp server
$Output = 1

# create a new MailServerSettings object:
$MailServerDetails = New-MailServerSettings -FQDN $FQDN -Output $Output -Port $Port -STARTTLS $STARTTLS -VerifyServerCertificate $VerifyServerCertificate -Username $Username -Password $Password 

# select the certificate used for the sender (used for signing messages):
# normally you would use:
$CertSenderFile = $MailFromPFXFile
$CertSenderPassword = $MailFromPFXPassword
$CertSender = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList ($CertSenderFile, $CertSenderPassword)

# get the From address from the certificate
$fromAddress = $CertSender.GetNameInfo('EmailName', $False)
$toAddress = $MailTo

# security settings:
# should the signed message be detached (readable for non s/mime clients):
$Detached = $True
# encryption algorithm to be used:
# valid values are: 'rc2','des','3des','aes128','aes192','aes256'
$EncryptionAlgorithm = 'aes256'
# when using encryption algorith 'rc2' you can specify a key length:
# valid values are: 40, 64, 128
$KeyLength = 128
# signing algorithm to be used:
# valid values are: 'md5','sha1','sha256','sha384','sha512'
$DigestAlgorithm = 'sha256'
# should the whole certificate chain be verified for the signing certificate:
# set this to $True if you do not have the root certificate (and/or intermediate certificates) of the signing certificate installed.
$EndCertOnly = $True
# if you want to use enhanced protection for signing the message:
# this adds signed attributes to the signed part
# it adds the recipient address to Pkcs9DocumentName
# and adds the subject to Pkcs9DocumentDescription
# this has no effect on most mail reader programs
$EnhancedProtection = $True

# the text of the body:
$Bodytext = $Body

# attchments you want to include:
# you can use wildcard to select multiple attachment
# set to $null if you do not want to include attachments
# normally you can use:
# $Attachments = get-item -Path "$(Split-Path -Path $script:MyInvocation.MyCommand.Path)\attachments\*"
#
# in this example we do not include attachments
$Attachments = $null

# set the subject and send using signing only:
#$Subject = 'SMIME Signed Emails'
$Result = Send-SignedMail -MailServer $mailserverdetails -Recipient $toAddress -SenderCertificate $certsender -Subject $subject -Body $bodytext -DigestAlgorithm $DigestAlgorithm -EndCertOnly $EndCertOnly -Detached $Detached -Attachments $attachments -EnhancedProtection $EnhancedProtection
$Result
