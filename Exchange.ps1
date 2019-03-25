# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'

function Send-RmWelcomeMail
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Employee', 'Faculty', 'Student', 'Shared')]
        [string]
        $Type
    )
    if ($Type -notin 'Employee', 'Faculty')
    {
        throw 'Wrong mailbox type. Welcome mail is only sent to employees.'
    }
    # Use SmtpClient instead of Send-MailMessage since the latter
    # always tries to authenticate with default credentials. A gMSA is
    # not allowed to authenticate to our Exchange SMTP receive connector.
    $smtpClient = New-Object -TypeName 'System.Net.Mail.SmtpClient'
    $smtpClient.UseDefaultCredentials = $false
    $smtpClient.Host = $Script:Config.Exchange.WelcomeMail.Server
    $msg = New-Object -TypeName 'System.Net.Mail.MailMessage'
    $msg.BodyEncoding = [System.Text.Encoding]::UTF8
    $msg.SubjectEncoding = [System.Text.Encoding]::UTF8
    $msg.IsBodyHtml = $true
    $msg.From = $Script:Config.Exchange.WelcomeMail.From
    $msg.Subject = $Script:Config.Exchange.WelcomeMail.Subject
    $msg.Body = $Script:Config.Exchange.WelcomeMail.Body
    $msg.To.Add($UserPrincipalName)
    $smtpClient.Send($msg)
}
