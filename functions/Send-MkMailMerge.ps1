<#
.SYNOPSIS

Sends out a mail merge using MailKit.

.DESCRIPTION

Send-MkMailMerge takes a message template file and a .csv file to mail lots of people.

It accepts text files and HTML as templates.

If using plain text and it was created in Microsoft Word, make sure you save it as .txt
in MSDOS format. This will alert you to any characters like curved quotes and dashes that
are not ASCII. The script will also try to replace common non-ASCII characters with
equivilants if you fail to do this.

The .csv file must contain column headers and one column called "Email".

You should probalby test with differnet mail clients (Thunderbird) using the -TestAddress
parameter to make sure there isn't wonky encoding. Sending to GMail as well as M365 is
a useful test as well.

You can also specifiy an HTML template. You can embed images and tables in the HTML template.

.EXAMPLE

Send-MKMailMerge -From 'John Doe <jdoe@example.com>' '.\testmessagetemplate.html' .\test.csv

Sends messages from John Doe <jdoe@example.com> to the recipients in test.csv.

.EXAMPLE

Send-MailMerge.ps1 -From 'John Doe <jdoe@example.com>' '.\testmessagetemplate.txt' .\test.csv -testaddress blah@example.com

Sends test messages to blah@example.com for the first 3 rows of the CSV.
The test messages report who they would actually go to at the bottom.

.PARAMETER MessageTemplate

Path to text file containing the message template.

Values you want substituted from the .csv file can be included by using an '&' symbol followed
by the column name from the csv. So, you might have Email and FirstName as columns. You could use
"Dear &FirstName," in the body of the message and the value would be subsituted for each recipient.

A sample template might be:
####################
Dear &FirstName,

Your mailbox is going bye bye.

Peace out,
Me
####################

In can also be an HTML file. Pictures and tables can be embeded. In the example below,
&HeaderPath should correspond to a column (minus the &) in the .csv with the path to the image file.
&MyTable should correspond to a column with the HTML for a table of multiple items.
####################
<html>
<body>
    <img width=624
height=108 id="Picture 3" alt="my header image" src="&HeaderPath">
    <p>Dear &FirstName,</p>
    <p>These mailboxes is going bye bye.</p>
    &MyTable
    <p>Peace out,<br>
        Me
    </p>
</body>
</html>
####################

.PARAMETER Csv

Path to a PowerShell importable .csv file. This needs to have column headers. There must be
at least one column with the name "Email". Powershell will not import a csv if there are extraneous
commas. Using spaces or special characters in the column names hasn't been tested. Probably best
to avoid them. The columns with "CC" nad "BCC" headers will also be used to add additional recipients.
Addresses in these columns should be semi-colon or comma separated.

If you want to include a different attachment per recipient, use a column named "Attachment" to specify
the file path of the attachment.

If you want to use embeded images in the mail merge, there should be a column with the path to the image.
E.g. "HeaderPath" as the and and "C:\header.jpg" as the value (repeated each line if it's the same).
See the MessageTemplate help for how to embed HeaderPath in the HTML template.

If you want to embed an HTML table, you can generate it with "ConvertTo-Html -Fragment" and put the results
for each contact in a column. E.g. "MyTable" and values like:
"<table>
<colgroup><col/><col/></colgroup>
<tr><th>Name</th><th>SamAccountName</th></tr>
<tr><td>John Doe</td><td>jdoe</td></tr>
<tr><td>Jane Doe</td><td>janed</td></tr>
</table>"

.PARAMETER FromAddress

Address the message should be from.
If you would like a descriptive name as opposed to just an address, use the format:
"Name <address>"
For instance:
"John Doe <john.doe@example.com"

.PARAMETER Subject
Subject line of the message

.PARAMETER TestAddress

TestAddress indicates you want to test how your message will look. This will cause the script to
ignore the Email column of the .csv and send all messages to this addess. By default it will use
the first 3 lines of the .csv. The actual addresses will be appended to the body of the message.

.PARAMETER TestCount

Used with TestAddress to indicate how many rows of the .csv to read. Default is 3.

.PARAMETER CcAddress

Addresses to CC on all messages in the mail merge. They should be semi-colon or comma separated.

.PARAMETER BccAddress

Addresses to BCC on all messages in the mail merge. They should be semi-colon or comma separated.

.PARAMETER Attachment

Path to a file to attach to all messsages in the mail merge.

.PARAMETER SMTPServer

SMTP Server to use. Defaults to the value of $env:PSEmailServer if it is set.

.PARAMETER Credential

If the SMTP server requires authentication, you can temporarily save your credentials with:

$cred = Get-Credential

Then, pass the $cred variable for this parameter.

If -Credential is used, the script will default to port 587 and SSL to ensure the credentials are not passed in
the clear.

.PARAMETER Port
SMTP Port to use. Defaults to 25 if -Credential is not used. If -Credential is used, defaults to 587.

.PARAMETER SMIME
Whether to use SMIME to sign or sign and encrypt the message. If you use this, you must have the certificate
matching the FromAddress installed.

.PARAMETER CertStore
The certificate store to use. Defaults to CurrentUser. Can be LocalMachine if running as an administrator.

#>
function Send-MkMailMerge {
    [CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = 'High')]
    param (
        [Parameter(Mandatory = $true)]
        $MessageTemplate,
        [Parameter(Mandatory = $true)]
        $Csv,
        [Parameter(Mandatory = $true)]
        [string]
        $Subject,
        [Parameter(Mandatory = $true)]
        [string]
        $FromAddress,
        [string]
        $TestAddress,
        [int]
        $TestCount = 3,
        [string]
        $CcAddress,
        [string]
        $BccAddress,
        [string[]]
        $Attachment,
        [string]
        $SMTPServer = $env:PSEmailServer,
        [int]
        $Port,
        [PSCredential]
        $Credential,
        [ValidateSet('Sign', 'SignAndEncrypt')]
        [string]
        $SMIME,
        [System.Security.Cryptography.X509Certificates.StoreLocation]
        $CertStore = 'CurrentUser'
    )

    begin {
        # Sends and invividual message from the mail merge
        function Send-MailMergeMessage {
            [CmdletBinding()]
            param($ToAddress,
                $FromAddress,
                $Subject,
                $MessageBody,
                $ReplaceHash,
                [string]$CcAddress,
                [string]$BccAddress,
                [string[]]$Attachment,
                [string]$SMTPServer,
                [int]$Port,
                [PSCredential]$Credential,
                [switch]$BodyAsHtml,
                [ValidateSet('Sign', 'SignAndEncrypt')]
                [string]
                $SMIME,
                $CertStore
            )

            # Step common unicode stuff down to ASCII
            $ASCIIReplace = @(
                @{Original = "’"; Replacement = "'" }
                @{Original = '–'; Replacement = '-' }
                @{Original = "‘"; Replacement = '`' }
                @{Original = '”'; Replacement = '"' }
                @{Original = '“'; Replacement = '"' }
                @{Original = '…'; Replacement = '...' }
                @{Original = '£'; Replacement = 'GBP' }
                @{Original = '•'; Replacement = '*' }
                @{Original = ' '; Replacement = ' ' }
                @{Original = 'é'; Replacement = 'e' }
                @{Original = 'ï'; Replacement = 'i' }
                @{Original = '´'; Replacement = "'" }
                @{Original = '—'; Replacement = '-' }
                @{Original = "·"; Replacement = '*' }
                @{Original = '„'; Replacement = ',,' }
                @{Original = '€'; Replacement = 'EUR' }
                @{Original = '®'; Replacement = '(R)' }
                @{Original = '¹'; Replacement = '(1)' }
                @{Original = '«'; Replacement = '<<' }
                @{Original = 'è'; Replacement = 'e' }
                @{Original = 'á'; Replacement = 'a' }
                @{Original = '™'; Replacement = 'TM' }
                @{Original = '»'; Replacement = '>>' }
                @{Original = 'ç'; Replacement = 'c' }
                @{Original = '½'; Replacement = '1/2' }
                @{Original = '­'; Replacement = '-' }
                @{Original = '°'; Replacement = ' degrees ' }
                @{Original = 'ä'; Replacement = 'a' }
                @{Original = 'É'; Replacement = 'E' }
                @{Original = "‚"; Replacement = ',' }
                @{Original = 'ü'; Replacement = 'u' }
                @{Original = 'í'; Replacement = 'i' }
                @{Original = 'ë'; Replacement = 'e' }
                @{Original = 'ö'; Replacement = 'o' }
                @{Original = 'à'; Replacement = 'a' }
                @{Original = '¬'; Replacement = ' ' }
                @{Original = 'ó'; Replacement = 'o' }
                @{Original = 'â'; Replacement = 'a' }
                @{Original = 'ñ'; Replacement = 'n' }
                @{Original = 'ô'; Replacement = 'o' }
                @{Original = '¨'; Replacement = '**' }
                @{Original = 'å'; Replacement = 'a' }
                @{Original = 'ã'; Replacement = 'a' }
                @{Original = 'ˆ'; Replacement = '^' }
                @{Original = '©'; Replacement = '(c)' }
                @{Original = 'Ä'; Replacement = 'A' }
                @{Original = 'Ï'; Replacement = 'I' }
                @{Original = 'ò'; Replacement = 'o' }
                @{Original = 'ê'; Replacement = 'e' }
                @{Original = 'î'; Replacement = 'i' }
                @{Original = 'Ü'; Replacement = 'U' }
                @{Original = 'Á'; Replacement = 'A' }
                @{Original = 'ß'; Replacement = 'ss' }
                @{Original = '¾'; Replacement = '3/4' }
                @{Original = 'È'; Replacement = 'E' }
                @{Original = '¼'; Replacement = '1/4' }
                @{Original = '†'; Replacement = '+' }
                @{Original = '³'; Replacement = "'" }
                @{Original = '²'; Replacement = "'" }
                @{Original = 'Ø'; Replacement = 'O' }
                @{Original = '¸'; Replacement = ',' }
                @{Original = 'Ë'; Replacement = 'E' }
                @{Original = 'ú'; Replacement = 'u' }
                @{Original = 'Ö'; Replacement = 'O' }
                @{Original = 'û'; Replacement = 'u' }
                @{Original = 'Ú'; Replacement = 'U' }
                @{Original = 'Œ'; Replacement = 'Oe' }
                @{Original = 'º'; Replacement = '?' }
                @{Original = '‰'; Replacement = '0/00' }
                @{Original = 'Å'; Replacement = 'A' }
                @{Original = 'ø'; Replacement = 'o' }
                @{Original = "˜"; Replacement = '~' }
                @{Original = 'æ'; Replacement = 'ae' }
                @{Original = 'ù'; Replacement = 'u' }
                @{Original = '‹'; Replacement = '<' }
                @{Original = '±'; Replacement = '+/-' }
            )

            foreach ($replacePair in $ASCIIReplace) {
                $MessageBody = $MessageBody.Replace( $replacePair.Original, $replacePair.Replacement)
                $Subject = $Subject.Replace($replacePair.Original, $replacePair.Replacement)
            }

            if ($ReplaceHash) {
                # Uses an MS Office class but if it isn't defined in the CSS, it won't matter
                $singleSpaceP = "<p class=MsoNormal style='margin-top:0in;margin-bottom:0in;margin-bottom:.0001pt;line-height:normal'>"
                $replaceHash.GetEnumerator() | % {
                    if ($BodyAsHtml) {
                        # If we have more than one item and the first character isn't an angle bracket (indicates we're already
                        # gettting HTML) then wrap the items in single space paragraph tags.
                        if ($_.Value.count -gt 1 -and $_.Value[0] -notmatch "^[<]") {
                            $HtmlSubstitution = "$singleSpaceP$($_.Value -join "</p>`r`n$singleSpaceP")</p>"
                        } else {
                            $HtmlSubstitution = $_.Value
                        }
                        $MessageBody = $MessageBody.Replace('&amp;' + $_.Name, $htmlSubstitution)
                        $MessageBody = $MessageBody.Replace('&' + $_.Name, $htmlSubstitution)
                    } else {
                        $MessageBody = $MessageBody.Replace('&' + $_.Name, $_.Value -join "`r`n")
                    }

                    $Subject = $Subject.Replace('&' + $_.Name, $_.Value)
                }
            }

            $MKMailParams = @{
                From       = $FromAddress
                Subject    = $Subject
                Body       = $MessageBody
                To         = $ToAddress
                BodyAsHTML = $BodyAsHtml
                SmtpServer = $SMTPServer
                Port       = $Port
            }

            if ($SMIME) {
                $MKMailParams['SMIME'] = $SMIME
            }

            if ($CertStore) {
                $MKMailParams['CertStore'] = $CertStore
            }

            if ($Credential) {
                $MKMailParams['Credential'] = $Credential
            }

            if ($Attachment) {
                $MKMailParams['Attachments'] = $Attachment
            }
            if ($BccAddress) {
                $MKMailParams['Bcc'] = $BccAddress
            }
            if ($CcAddress) {
                $MKMailParams['Cc'] = $CcAddress
            }

            Send-MKMailMessage @MKMailParams
        }

        if (-not (Test-Path $MessageTemplate)) {
            throw "The message template path `"$MessageTemplate`"was invalid."
        }

        if (-not (Test-Path $Csv)) {
            throw "The csv path `"$Csv`" was invalid."
        }

        $Contacts = @(Import-Csv $csv -ErrorVariable ImportCSVError)
        if ($ImportCSVError) {
            throw "Could not import csv file properly. Make sure there is a header row and there are no extra commas.`r`n$($ImportCSVError[0].Exception.Message)"
        }

        if ($Contacts[0] | Get-Member -Name 'Email') {
            Write-Debug "Email property existed in .CSV."
        } else {
            throw "CSV did not contain an Email column."
        }

        $TemplateText = Get-Content $MessageTemplate

        $TemplateBody = $templateText -join "`n"

        if ($Attachment) {
            if (-not (Test-Path $Attachment) ) {
                throw "Attachment $Attachment could not be found."
            }
        }

        $GlobalCcAddress = $CcAddress
        $GlobalBccAddress = $BccAddress
        $GlobalAttachment = $Attachment

        $BodyAsHtml = $MessageTemplate -match "\.(htm|html)$"
    }

    Process {
        if ($TestAddress) {
            Write-Verbose "Sending up to $TestCount messages to $Testaddress for testing." -Verbose
            Write-Verbose "The email addresses have been appended to the body of the message." -Verbose
            Write-Verbose "The messages would go to the addresses below:" -Verbose
            $Contacts = $Contacts[0..($TestCount - 1 )]
        } else {
            $ShouldProcessResult = $PsCmdlet.ShouldProcess("$($Contacts.count) emails", "Send mail merge '$subject'.")
        }
        $CSVLine = 1
        foreach ($Contact in $Contacts) {
            $CSVLine++
            if ($Contact.Email -notmatch '@') {
                Write-Warning "Email address on line $CSVLine  of the CSV was invalid. Skipping."
                continue
            }
            Write-Verbose "$($Contact.Email)" -Verbose

            $ReplaceHash = @{}
            $Contact.psobject.properties | % {
                $ReplaceHash[$_.Name] = $_.Value -split "`r`n"
            }

            if ($Contact.Cc) {
                $CcAddress = "$GlobalCcAddress;$($Contact.Cc)"
            } else {
                $CcAddress = $globalCcAddress
            }

            if ($Contact.Bcc) {
                $BccAddress = "$GlobalBccAddress;$($Contact.Bcc)"
            } else {
                $BccAddress = $GlobalBccAddress
            }

            if ($Contact.Attachment) {
                $Attachment = @($Contact.Attachment.Split(';'))
                if ($GlobalAttachment) {
                    $Attachment += $GlobalAttachment
                }
            }

            # Parameters that are the same whether we're testing or not
            $SendMailMergeMessageParams = @{
                FromAddress = $FromAddress
                Subject     = $Subject
                ReplaceHash = $ReplaceHash
                Attachment  = $Attachment
                SMTPServer  = $SMTPServer
                Port        = $Port
                Credential  = $Credential
                BodyAsHtml  = $BodyAsHtml
            }
            if ($SMIME) {
                $SendMailMergeMessageParams['SMIME'] = $SMIME
            }
            if ($CertStore) {
                $SendMailMergeMessageParams['CertStore'] = $CertStore
            }
            # Parameters that are different if we're testing
            if ($TestAddress) {
                $SendMailMergeMessageParams['ToAddress'] = $TestAddress
                $SendMailMergeMessageParams['CcAddress'] = $TestAddress
                $SendMailMergeMessageParams['BccAddress'] = $TestAddress
                $SendMailMergeMessageParams['MessageBody'] = ($TemplateBody +
                    "`nWould go to: $($Contact.Email)" +
                    "`nWould CC: $CcAddress" +
                    "`nWould BCC: $BccAddress" +
                    "`nWould attach: $($Attachment -join "|")"
                )
            } else {
                $SendMailMergeMessageParams['ToAddress'] = $Contact.Email -split "[,;]\s*" | ? { $_ -notmatch '^\s*$' }
                $SendMailMergeMessageParams['CcAddress'] = $CcAddress -split "[,;]\s*" | ? { $_ -notmatch '^\s*$' }
                $SendMailMergeMessageParams['BccAddress'] = $BccAddress -split "[,;]\s*" | ? { $_ -notmatch '^\s*$' }
                $SendMailMergeMessageParams['MessageBody'] = $TemplateBody
            }

            if ($TestAddress -or $ShouldProcessResult) {
                Send-MailMergeMessage @SendMailMergeMessageParams
            }
        }
    }
}