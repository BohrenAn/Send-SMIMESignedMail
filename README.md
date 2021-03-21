# Send-SMIMESignedMail
This Powershell Script sends an S/MIME Signed Email. 
You need to have a PFX of your Certificate with a Password.
From will be extracted from the PFX Certificate

## How to use it

Send Email without SMTP authentication
.\Send-SMIMESignedMail.ps1 -MailFromPFXFile E:\a.bohren@icewolf.ch.pfx -MailFromPFXPassword "MyPFXPassword" -MailTo recipient@domain.tld -Subject "Test" -Body "Just a Test" -SMTPServer 172.21.175.61 -SMTPPort 25 
```

Send Email with SMTP authentication
```
.\Send-SMIMESignedMail.ps1 -MailFromPFXFile E:\a.bohren@icewolf.ch.pfx -MailFromPFXPassword "MyPFXPassword" -MailTo recipient@domain.tld -Subject "Test" -Body "Just a Test" -SMTPServer 172.21.175.61 -SMTPPort 25 -SMTPUsername "YourSMTPUsername" -SMTPPassword "YourSMTPPassword"
```