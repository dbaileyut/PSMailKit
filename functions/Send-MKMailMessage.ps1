<#
.SYNOPSIS
    Send mail message using MailKit library
.DESCRIPTION
    The Send-MKMailMessage cmdlet sends an email message from within PowerShell.

    Uses https://www.nuget.org/packages/MailKit/ to faciliitate SMIME

.EXAMPLE
    Send-MKMailMessage -To "User01 <user01@example.com>" -From "User02 <user02@example.com>" -Subject "Test mail" -SMIME Sign

    This command sends an email message from User01 to User02.

    The mail message has a subject, which is required, but it does not have a body, which is optional. Also, because the SmtpServer parameter is not specified,
    Send-MKMailMessage uses the value of the $env:PSEmailServer preference variable for the SMTP server.

    Signs the email if a valid signing certificate and key is found in the current user's certificate store.
.EXAMPLE
    Send-MKMailMessage -From "User01 <user01@example.com>" -To "User02 <user02@example.com>", "User03 <user03@example.com>" -Subject "Sending the Attachment" -Body "Forgot to send the
    attachment. Sending now." -Attachments "data.csv" -Priority High -SmtpServer "smtp.fabrikam.com" -SMIME SignAndEncrypt -CertStore LocalMachine

    This command sends an email message with an attachment from User01 to two other users.

    It specifies a priority value of High.

    SMIME encrypts and signs if there's a valid signing certificate and key for the from address and the recipients' certificates can be found as well in
    the LocalMachine certificate store.

.EXAMPLE
    Send-MKMailMessage -To "User01 <user01@example.com>" -From "ITGroup <itdept@example.com>" -Cc "User02 <user02@example.com>" -bcc "ITMgr <itmgr@example.com>" -Subject "Don't forget today's
    meeting!" -Credential domain01\admin01 -UseSsl -SMIME Sign -CertStore LocalMachine

    This command sends an email message from User01 to the ITGroup mailing list with a copy (Cc) to User02 and a blind carbon copy (Bcc) to the IT manager (ITMgr).

    The command uses the credentials of a domain administrator and the UseSsl parameter.

    SMIME signs if there's a valid signing certificate and key for the from address in the LocalMachine certificate store.
.INPUTS
    String
        You can pipe the path and file names of attachments to Send-MKMailMessage
.OUTPUTS
    None
        This cmdlet does not generate any output.
.NOTES
    See also https://github.com/jstedfast/MailKit

#>
function Send-MKMailMessage {
    [CmdletBinding(
        SupportsShouldProcess = $true,
        ConfirmImpact = 'Medium')]
    Param (
        # Paths to attachments
        [Parameter(ValueFromPipeline = $True)]
        [ValidateNotNullOrEmpty()]
        [Alias("PsPath")]
        [String[]] $Attachments,

        # Blind CC addresses
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [String[]] $Bcc,

        <#
            Message body text. If the body is HTML, -BodyAsHTML needs to be specified.
            The HTML can include <img> tags with src attributes using local paths to
            the image files. These will be embeded in the MIME message.

            Example: <img alt="My alt" src="C:\myimage.png" >
        #>
        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [String] $Body,

        # Whether the message body is in HTML
        [Parameter()]
        [Alias("BAH")]
        [Switch] $BodyAsHtml,

        # Body encoding
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [Alias("BE")]
        [System.Text.Encoding] $Encoding,

        # Carbon copy addresses
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [String[]] $Cc,

        # Certificate store for signing/encrypting certificates
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [System.Security.Cryptography.X509Certificates.StoreLocation] $CertStore = 'CurrentUser',

        <#
        To Implement? MailKit doesn't seem to have an easy property for this
        #Delivery notification options
        [Parameter(Mandatory=$False,
                Position=-2147483648,
                ValueFromPipeline=$False,
                ValueFromPipelineByPropertyName=$False,
                ValueFromRemainingArguments=$False)]
            [ValidateNotNullOrEmpty()]
            [Alias("DNO")]
        [System.Net.Mail.DeliveryNotificationOptions] $DeliveryNotificationOption,
        #>

        # From address - use format "Display Name <emailaddr@blah.com>"
        # or just emailaddr@blah.com
        [Parameter(Mandatory = $True,
            Position = 4)]
        [ValidateNotNullOrEmpty()]
        [String] $From,

        # ReplyTo address - use format "Display Name <emailaddr@blah.com>"
        # or just emailaddr@blah.com
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [String[]] $ReplyTo,

        # DNS or IP address of the SMTP Server
        [Parameter(Position = 3)]
        [ValidateNotNullOrEmpty()]
        [Alias("ComputerName")]
        [String] $SmtpServer = $env:PSEmailServer,

        # Mail message priority
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [System.Net.Mail.MailPriority] $Priority,

        # Message subject
        [Parameter(Mandatory = $True,
            Position = 1)]
        [ValidateNotNullOrEmpty()]
        [Alias("sub")]
        [String] $Subject,

        # SMIME Sign or SignAndEncrypt
        [Parameter()]
        [ValidateSet('Sign', 'SignAndEncrypt')]
        [String] $SMIME,

        # To addresses
        [Parameter(Mandatory = $True,
            Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String[]] $To,

        # Credentials for the SMTP Server
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [PSCredential] $Credential,

        # Whether the SMTP server requires SSL
        [Parameter()]
        [Switch] $UseSsl,

        # SMTP Server port number
        [Parameter()]
        [ValidateRange(0, 2147483647)]
        [System.Int32] $Port = 25
    )

    begin {

        function Add-MKAddress ($AddressList, $AddressHeader, $Message) {
            foreach ($Addr in $AddressList) {
                try {
                    $Message.$AddressHeader.Add($Addr)
                } catch {
                    throw "Failed to add $AddressHeader address: `"$Addr`": $_"
                }
            }
            Write-Verbose ("$AddressHeader`: " + $Message.$AddressHeader)
        }

        function Get-MKPriority ([System.Net.Mail.MailPriority] $DotNetPriority) {
            switch ($DotNetPriority) {
                [System.Net.Mail.MailPriority]::High {
                    return [MimeKit.MessagePriority]::Urgent
                }
                [System.Net.Mail.MailPriority]::Low {
                    return [MimeKit.MessagePriority]::NonUrgent
                }
                [System.Net.Mail.MailPriority]::Normal {
                    return [MimeKit.MessagePriority]::Normal
                }
                default {
                    return $null
                }
            }

        }

        # Initialize attachement variables since we have to process the pipeline
        $AttachmentFails = @()

        # Build message
        $Message = [MimeKit.MimeMessage]::new()

        $AddressHeaders = @{
            'From' = $From
            'To'   = $To
            'Cc'   = $Cc
            'Bcc'  = $Bcc
            'ReplyTo'  = $ReplyTo
        }

        foreach ($Header in $AddressHeaders.Keys) {
            try {
                Add-MKAddress -AddressList $AddressHeaders[$Header] -AddressHeader $Header -Message $Message
            } catch {
                Write-Error $_
                return
            }
        }

        $Message.Subject = $Subject
        Write-Verbose ("Subject: " + $Message.Subject )

        $Builder = [MimeKit.BodyBuilder]::new()

        try {
            if ($BodyAsHtml) {
                $BodyProccessed = $Body
                # I know regex is not the best way to do this, but it works
                if ($Body -match "<img[^>]+src=`"(?<ImgPath>[^`"]+)`"") {
                    $ImgPaths = $Matches.ImgPath
                    foreach ($Path in $ImgPaths) {
                        $image = $Builder.LinkedResources.Add($Path)
                        $image.ContentId = [MimeKit.Utils.MimeUtils]::GenerateMessageId();
                        $BodyProccessed = $BodyProccessed.Replace($Path, "cid:" + $image.ContentId)
                    }
                }
                $Builder.HtmlBody = $BodyProccessed
            } else {
                $Builder.TextBody = $Body
            }
        } catch {
            Write-Error "Failed to add body.`r`n$_"
            return
        }

        # Confirm we have an SMTP server
        if ($SmtpServer -match "^\s*$") {
            Write-Error "SmtpServer was empty or null. Specify the parameter or set `$env:PSEmailServer."
            return
        }

        if ($SMIME) {
            # Create security context
            $Ctx = [MimeKit.Cryptography.WindowsSecureMimeContext]::new($CertStore)
        }

    }

    process {
        foreach ($Attachment in $Attachments) {
            if (Test-Path $Attachment -PathType Leaf) {
                $FullPath = (Get-Item $Attachment).FullName
                try {
                    $Builder.Attachments.Add($FullPath) | Out-Null
                } catch {
                    Write-Error "Failed to add attachment `"$Attachment`" to message body object. $_"
                    $AttachmentFails += $Attachment
                }
            } else {
                $AttachmentFails += $Attachment
            }
        }
    }

    end {
        if ($AttachmentFails.Count -eq 0) {

            $Message.Body = $Builder.ToMessageBody()

            if ($Encoding) {
                # Eh, not sure if this is the right thing... need to test
                $Message.Headers['Subject'] = [MimeKit.Header]::new($Encoding, 'Subject', $Subject)
                foreach ($BodyPart in $Message.BodyParts) {
                    if (($BodyPart.IsPlain -or $BodyPart.IsHtml) -and -not $BodyPart.IsAttachment) {
                        $BodyPart.ContentType.Charset = $Encoding
                    }
                }
            }

            if ($Priority) {
                $MKPriority = Get-MKPriority $Priority
                if ($null -ne $MKPriority) {
                    $Message.Priority = $MKPriority
                } else {
                    Write-Error "Failed to convert `"$Priority`" to MailKit enum."
                }
            }

            switch ($SMIME) {
                "Sign" {
                    try {
                        $Message.Sign($Ctx, [MimeKit.Cryptography.DigestAlgorithm]::Sha256)
                    } catch {
                        Write-Error "Failed to sign the message: $_"
                        return
                    }
                }
                "SignAndEncrypt" {
                    try {
                        $Message.SignAndEncrypt($Ctx, [MimeKit.Cryptography.DigestAlgorithm]::Sha256)
                    } catch {
                        Write-Error "Failed to sign and encrypt the message: $_"
                        return
                    }
                }
            }

            # Smtp Connection
            $SmtpClient = [MailKit.Net.Smtp.SmtpClient]::new()
            $SecureSocketOptions = $null
            $ShouldBeSecure = $UseSsl -or $Port -eq 587 -or $Port -eq 465 -or $Credential
            if ($UseSsl) {
                $SecureSocketOptions = [MailKit.Security.SecureSocketOptions]::Auto
            }
            if ($Port -eq 25 -and ($Credential -or $UseSsl)) {
                $SecureSocketOptions = [MailKit.Security.SecureSocketOptions]::StartTls
            }
            try {
                Write-Verbose "Connecting to `"$SmtpServer`" on port $Port. UseSSL: $UseSSL"
                $SmtpClient.Connect( $SmtpServer, $Port, $SecureSocketOptions )
            } catch {
                Write-Error "Failed to connenct to `"$SmtpServer`":`r`n$_"
            }

            if ($ShouldBeSecure -and -not $SmtpClient.IsSecure) {
                Write-Error "Failed to connect to secure SMTP server."
                return
            }

            if ($Credential) {
                if ($SmtpClient.IsSecure) {
                    Write-Verbose "Authenticating to `"$SmtpServer`" with `"$($Credential.UserName)`""
                    $SmtpClient.Authenticate($Credential)
                } else {
                    Write-Error "Cannot authenticate to unsecure SMTP server."
                }
            }

            if ($pscmdlet.ShouldProcess("Target", "Operation")) {
                try {
                    Write-Verbose "Sending message..."
                    $SmtpClient.Send($Message) | Out-Null
                } catch {
                    Write-Error "Failed to send message:`r`n$_"
                }
            }
        } else {
            Write-Error ("The following attachment paths are not valid files:`r`n" +
                "`t`"" + ($AttachmentFails -join "`"`r`n`t`"") + '"'
            )
        }

        try {
            if ($SmtpClient) {
                Write-Verbose "Disconnecting from `"$SmtpServer`""
                $SmtpClient.Disconnect($true)
            }
        } catch {
            Write-Error "Failed to disconnect cleanly from `"$SmtpServer`""
        }
    }
}