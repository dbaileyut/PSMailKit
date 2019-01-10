<#
.SYNOPSIS
    Send mail message using MailKit library
.DESCRIPTION
    The Send-MailMessage cmdlet sends an email message from within Windows PowerShell.

    Uses https://www.nuget.org/packages/MailKit/ to faciliitate SMIME

.EXAMPLE
    Send-MailMessage -To "User01 <user01@example.com>" -From "User02 <user02@example.com>" -Subject "Test mail"

    This command sends an email message from User01 to User02.

    The mail message has a subject, which is required, but it does not have a body, which is optional. Also, because the SmtpServer parameter is not specified, Send-MailMessage uses the value of
    the $PSEmailServer preference variable for the SMTP server.
.EXAMPLE
    Send-MailMessage -From "User01 <user01@example.com>" -To "User02 <user02@example.com>", "User03 <user03@example.com>" -Subject "Sending the Attachment" -Body "Forgot to send the
    attachment. Sending now." -Attachments "data.csv" -Priority High -dno onSuccess, onFailure -SmtpServer "smtp.fabrikam.com"

    This command sends an email message with an attachment from User01 to two other users.

    It specifies a priority value of High and requests a delivery notification by email when the email messages are delivered or when they fail.

.EXAMPLE
    Send-MailMessage -To "User01 <user01@example.com>" -From "ITGroup <itdept@example.com>" -Cc "User02 <user02@example.com>" -bcc "ITMgr <itmgr@example.com>" -Subject "Don't forget today's
    meeting!" -Credential domain01\admin01 -UseSsl

    This command sends an email message from User01 to the ITGroup mailing list with a copy (Cc) to User02 and a blind carbon copy (Bcc) to the IT manager (ITMgr).

    The command uses the credentials of a domain administrator and the UseSsl parameter.
.INPUTS
    System.String
        You can pipe the path and file names of attachments to Send-MailMessage
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
        # Param1 help description
        [Parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ValidateCount(0,5)]
        [ValidateSet("sun", "moon", "earth")]
        [Alias("p1")]
        $Param1,

        # Param2 help description
        [Parameter(ParameterSetName='Parameter Set 1')]
        [AllowNull()]
        [AllowEmptyCollection()]
        [AllowEmptyString()]
        [ValidateScript({$true})]
        [ValidateRange(0,5)]
        [int]
        $Param2,

        # Param3 help description
        [Parameter(ParameterSetName='Another Parameter Set')]
        [ValidatePattern("[a-z]*")]
        [ValidateLength(0,15)]
        [String]
        $Param3
    )

    begin {
    }

    process {
        if ($pscmdlet.ShouldProcess("Target", "Operation")) {

        }
    }

    end {
    }
}