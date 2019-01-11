<#
.SYNOPSIS
    Send mail message using MailKit library
.DESCRIPTION
    The Send-MKMailMessage cmdlet sends an email message from within Windows PowerShell.

    Uses https://www.nuget.org/packages/MailKit/ to faciliitate SMIME

.EXAMPLE
    Send-MKMailMessage -To "User01 <user01@example.com>" -From "User02 <user02@example.com>" -Subject "Test mail"

    This command sends an email message from User01 to User02.

    The mail message has a subject, which is required, but it does not have a body, which is optional. Also, because the SmtpServer parameter is not specified, Send-MKMailMessage uses the value of
    the $PSEmailServer preference variable for the SMTP server.
.EXAMPLE
    Send-MKMailMessage -From "User01 <user01@example.com>" -To "User02 <user02@example.com>", "User03 <user03@example.com>" -Subject "Sending the Attachment" -Body "Forgot to send the
    attachment. Sending now." -Attachments "data.csv" -Priority High -dno onSuccess, onFailure -SmtpServer "smtp.fabrikam.com"

    This command sends an email message with an attachment from User01 to two other users.

    It specifies a priority value of High and requests a delivery notification by email when the email messages are delivered or when they fail.

.EXAMPLE
    Send-MKMailMessage -To "User01 <user01@example.com>" -From "ITGroup <itdept@example.com>" -Cc "User02 <user02@example.com>" -bcc "ITMgr <itmgr@example.com>" -Subject "Don't forget today's
    meeting!" -Credential domain01\admin01 -UseSsl

    This command sends an email message from User01 to the ITGroup mailing list with a copy (Cc) to User02 and a blind carbon copy (Bcc) to the IT manager (ITMgr).

    The command uses the credentials of a domain administrator and the UseSsl parameter.
.INPUTS
    System.String
        You can pipe the path and file names of attachments to Send-MKMailMessage
.OUTPUTS
    None
        This cmdlet does not generate any output.
.NOTES
    See also https://github.com/jstedfast/MailKit

#>
function Send-MKMailMessage {
    [CmdletBinding(
                   SupportsShouldProcess=$true,
                   PositionalBinding=$false,
                   ConfirmImpact='Medium')]
    [Alias()]
    Param (
        # Paths to attachments
        [Parameter(Mandatory=$False,
                Position=-2147483648,
                ValueFromPipeline=$True,
                ValueFromPipelineByPropertyName=$False,
                ValueFromRemainingArguments=$False)]
            [ValidateNotNullOrEmpty()]
            [Alias("PsPath")]
        [System.String[]] $Attachments,

        # Blind CC addresses
        [Parameter(Mandatory=$False,
                Position=-2147483648,
                ValueFromPipeline=$False,
                ValueFromPipelineByPropertyName=$False,
                ValueFromRemainingArguments=$False)]
            [ValidateNotNullOrEmpty()]
        [System.String[]] $Bcc,

        # Message body
        [Parameter(Mandatory=$False,
                Position=2,
                ValueFromPipeline=$False,
                ValueFromPipelineByPropertyName=$False,
                ValueFromRemainingArguments=$False)]
            [ValidateNotNullOrEmpty()]
        [System.String] $Body,

        # Whether the message body is in HTML
        [Parameter(Mandatory=$False,
                Position=-2147483648,
                ValueFromPipeline=$False,
                ValueFromPipelineByPropertyName=$False,
                ValueFromRemainingArguments=$False)]
            [Alias("BAH")]
        [Switch] $BodyAsHtml,

        # Body encoding
        [Parameter(Mandatory=$False,
                Position=-2147483648,
                ValueFromPipeline=$False,
                ValueFromPipelineByPropertyName=$False,
                ValueFromRemainingArguments=$False)]
            [ValidateNotNullOrEmpty()]
            [Alias("BE")]
        [System.Text.Encoding] $Encoding,

        # Carbon copy addresses
        [Parameter(Mandatory=$False,
                Position=-2147483648,
                ValueFromPipeline=$False,
                ValueFromPipelineByPropertyName=$False,
                ValueFromRemainingArguments=$False)]
            [ValidateNotNullOrEmpty()]
        [System.String[]] $Cc,

        # Delivery notification options
        [Parameter(Mandatory=$False,
                Position=-2147483648,
                ValueFromPipeline=$False,
                ValueFromPipelineByPropertyName=$False,
                ValueFromRemainingArguments=$False)]
            [ValidateNotNullOrEmpty()]
            [Alias("DNO")]
        [System.Net.Mail.DeliveryNotificationOptions] $DeliveryNotificationOption,

        # From address - use format "Display Name <emailaddr@blah.com>"
        # or just emailaddr@blah.com
        [Parameter(Mandatory=$True,
                Position=-2147483648,
                ValueFromPipeline=$False,
                ValueFromPipelineByPropertyName=$False,
                ValueFromRemainingArguments=$False)]
            [ValidateNotNullOrEmpty()]
        [System.String] $From,

        # DNS or IP of the SMTP Server
        [Parameter(Mandatory=$False,
                Position=3,
                ValueFromPipeline=$False,
                ValueFromPipelineByPropertyName=$False,
                ValueFromRemainingArguments=$False)]
            [ValidateNotNullOrEmpty()]
            [Alias("ComputerName")]
        [System.String] $SmtpServer = $PSEmailServer,

        # Mail message priority
        [Parameter(Mandatory=$False,
                Position=-2147483648,
                ValueFromPipeline=$False,
                ValueFromPipelineByPropertyName=$False,
                ValueFromRemainingArguments=$False)]
            [ValidateNotNullOrEmpty()]
        [System.Net.Mail.MailPriority] $Priority,

        # Message subject
        [Parameter(Mandatory=$True,
                Position=1,
                ValueFromPipeline=$False,
                ValueFromPipelineByPropertyName=$False,
                ValueFromRemainingArguments=$False)]
            [ValidateNotNullOrEmpty()]
            [Alias("sub")]
        [System.String] $Subject,

        # To addresses
        [Parameter(Mandatory=$True,
                Position=0,
                ValueFromPipeline=$False,
                ValueFromPipelineByPropertyName=$False,
                ValueFromRemainingArguments=$False)]
            [ValidateNotNullOrEmpty()]
        [System.String[]] $To,

        # Credentials for the SMTP Server
        [Parameter(Mandatory=$False,
                Position=-2147483648,
                ValueFromPipeline=$False,
                ValueFromPipelineByPropertyName=$False,
                ValueFromRemainingArguments=$False)]
            [ValidateNotNullOrEmpty()]
        [PSCredential] $Credential,

        # Whether the SMTP server requires SSL
        [Parameter(Mandatory=$False,
                Position=-2147483648,
                ValueFromPipeline=$False,
                ValueFromPipelineByPropertyName=$False,
                ValueFromRemainingArguments=$False)]
        [Switch] $UseSsl,

        # SMTP Server port number
        [Parameter(Mandatory=$False,
                Position=-2147483648,
                ValueFromPipeline=$False,
                ValueFromPipelineByPropertyName=$False,
                ValueFromRemainingArguments=$False)]
            [ValidateRange(0, 2147483647)]
        [System.Int32] $Port = 25
    )

    begin {
        # Smtp Connection
        if ($SmtpServer -match "^\s*$") {
            Write-Error "SmtpServer was empty or null. Specify the parameter or set `$PSEmailServer."
            return
        }
        $SmtpClient = [MailKit.Net.Smtp.SmtpClient]::new()
        $SecureSocketOptions = $null
        if ($UseSsl) {
            $SecureSocketOptions = 'Auto'
        }
        try {
            Write-Verbose "Connecting to `"$SmtpServer`" on port $Port. UseSSL: $UseSSL"
            $SmtpClient.Connect( $SmtpServer, $Port, $SecureSocketOptions )
        } catch {
            Write-Error "Failed to connenct to `"$SmtpServer`":`r`n$_"
        }

        if ($Credential) {
            $SmtpClient.Authenticate($Credential)
        }

        $AttachmentArr = @()
        $AttachmentErr = $false
    }

    process {
        foreach ($Attachment in $Attachments) {
            if (Test-Path $Attachment -PathType Leaf) {
                $AttachmentArr += $Attachment
            } else {
                Write-Error "Attachment `"$Attachment`" does not exist or is not a file."
                $AttachmentErr = $true
            }
        }
    }

    end {
        if (-not $AttachmentErr) {
            if ($pscmdlet.ShouldProcess("Target", "Operation")) {

            }
        }

        try {
            $SmtpClient.Disconnect($true)
        } catch {
            Write-Error "Failed to disconnect cleanly from `"$SmtpServer`""
        }
    }
}