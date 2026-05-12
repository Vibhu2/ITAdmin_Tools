# ============================================================
# SCRIPT  : Send-VBPasswordExpiryNotification.ps1
# VERSION : 1.1.0
# CHANGED : 12-05-2026 -- Added full help block; cleared TEST_USER for production
# AUTHOR  : Vibhu Bhatnagar
# PURPOSE : Notify AD users whose passwords expire in 14 or 7 days
# ENCODING: UTF-8 with BOM -- do not re-save without BOM
# ------------------------------------------------------------
# CHANGELOG (last 3-5 only -- full history in Git)
# v1.1.0 -- 12-05-2026 -- Finalized: help block added, TEST_USER cleared
# v1.0.0 -- 12-05-2026 -- Initial release
# ============================================================

<#
.SYNOPSIS
    Sends password expiry reminder emails to AD users via Office 365 SMTP.

.DESCRIPTION
    Queries Active Directory for enabled users and calculates how many days
    remain until each user's password expires based on the default domain
    password policy. Sends a reminder email at 14 days and again at 7 days.
    Logs all actions and exports a per-run CSV report.

    Requires a pre-exported PSCredential XML at the path defined in $CRED_PATH.
    To create: Get-Credential | Export-Clixml -Path 'C:\Scripts\smtpcred.xml'

    To test against a single user, set $TEST_USER to a valid SamAccountName.
    Set $TEST_USER to $null to run against all enabled AD users.

.PARAMETER None
    All configuration is in the CONFIGURATION block at the top of the script.

.EXAMPLE
    # Run in production against all enabled users
    .\Send-VBPasswordExpiryNotification.ps1

.EXAMPLE
    # Test against a single user (set $TEST_USER = 'jsmith' in config first)
    .\Send-VBPasswordExpiryNotification.ps1

.OUTPUTS
    Log  : C:\Scripts\PasswordExpiry.log
    CSV  : C:\Scripts\PasswordExpiryReport_yyyy-MM-dd.csv

.NOTES
    Version  : 1.1.0
    Author   : Vibhu
    Modified : 12-05-2026
    Requires : ActiveDirectory module, SMTP credential XML
    Schedule : Run daily via Task Scheduler as a domain service account
#>

$ErrorActionPreference = 'Stop'

# --- CONFIGURATION ---

$LOG_PATH    = 'C:\Scripts\PasswordExpiry.log'
$REPORT_PATH = Join-Path -Path 'C:\Scripts' -ChildPath "PasswordExpiryReport_$(Get-Date -Format 'yyyy-MM-dd').csv"
$CRED_PATH   = 'C:\Scripts\smtpcred.xml'

$SMTP_SERVER = 'smtp.office365.com'
$SMTP_PORT   = 587
$SMTP_FROM   = 'no-reply@garp.com'

$NOTIFY_DAYS = @(14, 7)

# Set to a SamAccountName to test against a single user, or $null for production
$TEST_USER   = $null

# --- HELPER FUNCTIONS ---

function Write-VBLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )
    $Line = "$(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')  $Message"
    Add-Content -Path $LOG_PATH -Value $Line -Encoding UTF8
}

# --- MAIN LOGIC ---

# Step 1 -- Initialise log and load dependencies
Import-Module -Name ActiveDirectory -ErrorAction Stop

Write-VBLog -Message '===== Script Started ====='

# Step 2 -- Load SMTP credential and domain password policy
$SmtpCred        = Import-Clixml -Path $CRED_PATH
$MaxPasswordAge  = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge

# Step 3 -- Get target users
if (-not [string]::IsNullOrEmpty($TEST_USER)) {
    $Users = Get-ADUser -Identity $TEST_USER `
        -Properties DisplayName, EmailAddress, PasswordLastSet, SamAccountName
}
else {
    $Users = Get-ADUser -Filter { Enabled -eq $true } `
        -Properties DisplayName, EmailAddress, PasswordLastSet, SamAccountName
}

# Step 4 -- Evaluate each user and send notification if due
$Report = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($user in $Users) {

    if ([string]::IsNullOrEmpty($user.EmailAddress)) {
        Write-VBLog -Message "SKIPPED  : $($user.SamAccountName) -- No email address"
        continue
    }

    if ($null -eq $user.PasswordLastSet) {
        Write-VBLog -Message "SKIPPED  : $($user.SamAccountName) -- No PasswordLastSet value"
        continue
    }

    $ExpiryDate    = $user.PasswordLastSet + $MaxPasswordAge
    $DaysRemaining = [math]::Round(($ExpiryDate - (Get-Date)).TotalDays)

    if ($NOTIFY_DAYS -notcontains $DaysRemaining) {
        Write-VBLog -Message "INFO     : $($user.SamAccountName) -- $DaysRemaining days remaining, no action"
        continue
    }

    $Subject = 'AWS WorkSpaces (EU) Password Expiring Soon'

    $Body = @"
Hello $($user.DisplayName),

This is an automated reminder that your AWS WorkSpaces (EU) password will expire on $($ExpiryDate.ToShortDateString()).
Please update your password before this date to prevent access issues.

When changing your password, remember it must be at least 16 characters long, include one uppercase letter,
one lowercase letter, one number, and one special character, and cannot reuse any of your last 24 passwords.
These requirements are enforced automatically.

This notification applies only to AWS WorkSpaces (EU) and does not apply to AWS WorkSpaces (US) or GARP systems.

How to Change Your Password:
1. Open the AWS WorkSpaces client installed on your laptop
2. On the sign-in screen, click "Forgot Password"
3. Follow the on-screen instructions to reset your password
4. Sign in again using your new password

Important Note:
Password resets must be completed using the AWS WorkSpaces client.

If you need assistance, contact helpdesk.alerts@garp.com

Regards,
GARP Helpdesk
"@

    try {
        Send-MailMessage `
            -To        $user.EmailAddress `
            -From      $SMTP_FROM `
            -Subject   $Subject `
            -Body      $Body `
            -SmtpServer $SMTP_SERVER `
            -Port      $SMTP_PORT `
            -UseSsl `
            -Credential $SmtpCred `
            -ErrorAction Stop

        Write-VBLog -Message "SUCCESS  : $($user.SamAccountName) -- Notified ($DaysRemaining days remaining)"

        $Report.Add([PSCustomObject]@{
            Date          = Get-Date -Format 'dd-MM-yyyy HH:mm:ss'
            Username      = $user.SamAccountName
            DisplayName   = $user.DisplayName
            Email         = $user.EmailAddress
            DaysRemaining = $DaysRemaining
            ExpiryDate    = $ExpiryDate.ToString('dd-MM-yyyy')
            Status        = 'Sent'
        })
    }
    catch {
        Write-VBLog -Message "ERROR    : $($user.SamAccountName) -- $($_.Exception.Message)"

        $Report.Add([PSCustomObject]@{
            Date          = Get-Date -Format 'dd-MM-yyyy HH:mm:ss'
            Username      = $user.SamAccountName
            DisplayName   = $user.DisplayName
            Email         = $user.EmailAddress
            DaysRemaining = $DaysRemaining
            ExpiryDate    = $ExpiryDate.ToString('dd-MM-yyyy')
            Status        = 'Failed'
        })
    }
}

# Step 5 -- Export report
if ($Report.Count -gt 0) {
    $Report | Export-Csv -Path $REPORT_PATH -NoTypeInformation -Encoding UTF8
    Write-VBLog -Message "REPORT   : Exported to $REPORT_PATH"
}
else {
    Write-VBLog -Message 'REPORT   : No users met notification criteria -- no file written'
}

Write-VBLog -Message '===== Script Completed ====='
