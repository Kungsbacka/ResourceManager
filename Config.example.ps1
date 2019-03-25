# Configuration
$Config = @{
    Logger = @{
        ConnectionString = 'Server=<Server>;Database=<Database>;Integrated Security=True'
        LogPath = '<log directory>'
    }
    SamlId = @{
        Domain = '<Domain>'
    }
    MicrosoftOnline = @{
        AccountName = '<License account name>'
        UsageLocation = 'SE'
    }
    # Office 365 credentials for an account with user administration privileges.
    # Password must be encrypted with the same credentials that is running the script.
    # Start a new PowerShell prompt as that user, run the following line
    # (Get-Credential).Password | ConvertFrom-SecureString | Set-Clipboard
    # and paste in the encrypted password.
    Office365 = @{
        MsolUser = '<Office 365 user with permissions to manage users>'
        MsolPassword = '<encrypted password>'
    }
    ExchangeOnline = @{
        Mailbox = @{
            EmailAddressPolicyEnabled = $false
            RetentionPolicy = '<Policy name>'
            AddressBookPolicy = '<Policy name>'
            Languages = @('sv-SE')
            Student = @{
                AddressBookPolicy = '<Policy name>'
            }
        }
        Owa = @{
            DateFormat = 'yyyy-MM-dd'
            TimeFormat = 'HH:mm'
            TimeZone = 'W. Europe Standard Time'
            Language = 'sv-SE'
        }
    }
    ExchangeOnprem = @{
        Servers = @(
            '<server1>'
            '<server2>'
            '<server3>'
            'etc'
        )
        ExchangeOnlineMailDomain = '<your tennant>.mail.onmicrosoft.com'
        Mailbox = @{
            EmailAddressPolicyEnabled = $false
            UseDatabaseQuotaDefaults = $false
            RetentionPolicy = '<Policy name>'
            AddressBookPolicy = '<Policy name>'
            IssueWarningQuota = 10000MB
            ProhibitSendQuota = 12000MB
            ProhibitSendReceiveQuota = 14000MB
            Student = @{
                RetentionPolicy = '<Policy name>'
                AddressBookPolicy = '<Policy name>'
                IssueWarningQuota = 5000MB
                ProhibitSendQuota = 5200MB
                ProhibitSendReceiveQuota = 5400MB
            }
        }
        Owa = @{
            DateFormat = 'yyyy-MM-dd'
            TimeFormat = 'HH:mm'
            TimeZone = 'W. Europe Standard Time'
            Language = 'sv-SE'
            LocalizeDefaultFolderName = $true
            OwaMailboxPolicy = '<Policy GUID>'
            Student = @{
                OwaMailboxPolicy = '<Policy GUID>'
            }
        }
        Calendar = @{
            DefaultCalendarPermission = 'Reviewer'
            WorkingHoursTimeZone = 'W. Europe Standard Time'
            ShowWeekNumbers = $true
            WeekStartDay = 'Monday'
            FirstWeekOfYear = 'FirstFourDayWeek'
            WorkingHoursStartTime = '08:00:00'
            WorkingHoursEndTime = '17:00:00'
            Student = @{
                WorkingHoursStartTime = '08:00:00'
                WorkingHoursEndTime = '16:00:00'
            }
        }
        AutoReply = @{
            DefaultMessage = ''
        }
    }
    Exchange = @{
        WelcomeMail = @{
            Server = '<SMTP server>'
            From = '<From address>'
            Subject = '<Subject line>'
            Body = @"
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
</head>
<body style="font-family:Calibri, Sans-serif; font-size:11pt; line-hight:120%;">
    <!-- content -->
</body>
</html>
"@
        }
    }
}
