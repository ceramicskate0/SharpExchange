# SharpExchange

This command line POC that shows how C# can be used to interact with Microsoft Exchange (EWS). Showing that it can be done in other tooling other than Powershell.
Yes this is a simple POC to show how it could be done. Its not 100%. You want to show off your l33t C# coder or red teamer skills open a pull request plz :)
This is for educational purposes only. Dont use for evil or illegal things.

## MS Documentation:

- https://learn.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.exchangeservice?view=exchange-ews-api
- https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications
- https://learn.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.mailbox?view=exchange-ews-api
- Google "Microsoft.Exchange.WebServices"  ;)

# Like the work dont forget to hit that Star Button

```
            
            Required Inputs (Must be in order shown):

            ReadEmailExchange.exe WEBDomain DomainName Password InternalDomainName DUMPItem
                Example WEBDomain: webmail.domain.com
                Example DomainName: User1
                Example Password: SecretPassword
                Example InternalDomainName: domain
                
                Options for DUMPItem:
                    Inbox
                    Sent
                    Drafts
                    Deleted
                    Skype
                    Attachments (Will Download Atatchments from the Inbox, DeletedItems, and Sent Items folders)
                    SendEmail ToEmailAddress~Subject~Body(Body can be file path)~AttachmentLocalFilePath(optional)
                    All (All == will try to dump all the items above)(I would default to this if unsure)

            Optional Inputs:

            ReadEmailExchange.exe WEBDomain DomainName Password InternalDomainName DUMPItem NumberOfSearchResultsToReturn
                            Example NumberOfSearchResultsToReturn (will return a maximum of the number,default 10): 10
                            Note: NumberOfSearchResultsToReturn must be a int/whole number

            Optional Inputs:

            ReadEmailExchange.exe WEBDomain DomainName Password InternalDomainName DUMPItem NumberOfSearchResultsToReturn OutputFileNameOrPath
                Example OutputFileNameOrPath: C:\file.csv
                Note: Program needs permission to write to location
                
```
## Dont use for evil or if not authorized to do so. This is for educational purposes only. Not an exploit. 

## #rd Party Depend:

- https://github.com/zzzprojects/html-agility-pack

- Microsoft.Exchange.WebServices

## If someone decides to use this (its already flagged by some A/V's on disk) here are some ideas for IOC:

- Its C#, so AMSI is likely in play on modern systems where it is enabled

- Can write text file to disk

- Uses default .NET user agent string (For example: ... .NET CLR ...)

- When Run the .NET exe could create the temp file in the user's account folder structure with its name.

- Many more opportunites exist if code is reviewed

## Credits:
- Stackover flow
- MS Docs
