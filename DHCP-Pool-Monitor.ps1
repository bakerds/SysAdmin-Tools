#################################################################################################################
# 
# DHCP-Pool-Monitor.ps1
# Report Windows Server DHCP Pool utilization
#
# This script works with scopes configured for failover/load balancing across two Windows servers.
# It will not properly report for scopes that are available on multiple servers not configured for failover.
#
# Copyright (C) 2018 Dan Baker
# Copyright (C) 2018 Lancaster Mennonite School
# Released under the terms of the GNU General Public License
#
# Produces an email report that looks something like this:
#
# Scope ID    Free   In Use   Percent
# 10.0.0.0    55     0        ░░░░░░░░░░░░░░░░░░░░ 0%
# 10.0.1.0    88     12       ██░░░░░░░░░░░░░░░░░░ 12%
# 10.0.2.0    100    0        ░░░░░░░░░░░░░░░░░░░░ 0%
# 10.0.3.0    45     9        ███░░░░░░░░░░░░░░░░░ 17%
# 10.0.4.0    122    345      ███████████████░░░░░ 74%
# 10.0.6.0    309    194      ████████░░░░░░░░░░░░ 39%
# 10.0.8.0    190    40       ███░░░░░░░░░░░░░░░░░ 17%
#
##################################################################################################################

##################################################################################################################
#
# Domain controllers to query
$DomainControllers = "DC1", "DC2", "DC3", "DC4"
#
# Log folder path
$LogFolder = "C:\PowerShell\DHCP-Pool-Monitor\Logs"
#
# Email Recipient for report
$LogRecipient = "sysadmin@xyz.com"
#
##################################################################################################################

Set-ExecutionPolicy unrestricted

# Logging
$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$LogPath = $LogFolder + "\DHCP-Pool-Monitor-$(get-date -f yyyy-MM-ddTHH.mm.ss).log"
Start-Transcript -Path $LogPath

$scopes = @()

$DomainControllers | ForEach {
    (Get-DhcpServerv4ScopeStatistics -ComputerName $_) | ForEach {
        If (-not ($_.Free -eq 0 -and $_.InUse -eq 0)) {
            $scopes += $_
        }
    }
}

$uniqueScopes = $scopes | Sort -Property @{Expression = {$_."ScopeId"."Address"}} | Group "ScopeId" | %{ $_.Group | Select -First 1 }

$uniqueScopes | Format-Table

$emailHtml = "<table style='font-family:Monaco, monospace, Courier'><tr><th>Scope ID</th><th>Free</th><th>In Use</th><th>Percent</th></tr>"

$uniqueScopes | ForEach {
    $emailHtml += "<tr><td>" + $_.ScopeId + "</td><td>" + $_.Free + "</td><td>" + $_.InUse + "</td><td>"
    $bars = [math]::Round($_.PercentageInUse / 5)
    For ($i=0; $i -lt $bars; $i++) {
        $emailHtml += "&#9608;"
    }
    For ($i=20; $i -gt $bars; $i--) {
        $emailHtml += "&#9617;"
    }
    $emailHtml += " " + [math]::Round($_.PercentageInUse) + "%</td></tr>`n"
}

$emailHtml += "</table>"

$smtpServer="smtp.gmail.com"
$from = "My Server <server@xyz.com>"
$credentials = new-object Management.Automation.PSCredential "server@xyz.com", (Get-Content "C:\PowerShell\email-pwd.txt" | ConvertTo-SecureString)

Try {
    Sleep 2
    Send-Mailmessage -smtpServer $smtpServer -from $from -to $LogRecipient -subject "DHCP Scopes Report" -BodyAsHtml $emailHtml -priority High -UseSsl -Credential $credentials -Port 587 -ErrorAction Stop
}
Catch {
    Sleep 60
    Send-Mailmessage -smtpServer $smtpServer -from $from -to $LogRecipient -subject "DHCP Scopes Report" -BodyAsHtml $emailHtml -priority High -UseSsl -Credential $credentials -Port 587 -ErrorAction Stop
}

Stop-Transcript
