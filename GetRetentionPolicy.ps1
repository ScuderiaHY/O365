# Get all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited

# Array to hold mailbox information
$mailboxInfo = @()

foreach ($mailbox in $mailboxes) {
    $mailboxData = New-Object PSObject -Property @{
        DisplayName = $mailbox.DisplayName
        RetentionPolicy = ""
        RetentionTags = @()
    }

    try {
        # Get retention policy applied to the mailbox
        $retentionPolicy = Get-Mailbox $mailbox | Select-Object -ExpandProperty RetentionPolicy
        if ($retentionPolicy) {
            $mailboxData.RetentionPolicy = $retentionPolicy
        }

        # Get retention tags applied to the mailbox
        $retentionTags = Get-Mailbox $mailbox | Get-RetentionPolicyTag | Where-Object { $_.RetentionPolicyTagType -eq "All" }
        foreach ($tag in $retentionTags) {
            $mailboxData.RetentionTags += $tag.Name
        }

    } catch [Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException] {
        Write-Host "Object not found for mailbox $($mailbox.DisplayName). Skipping..."
        continue
    } catch {
        Write-Host "Error occurred while processing mailbox $($mailbox.DisplayName): $_"
    }

    $mailboxInfo += $mailboxData
}

# Export to CSV
$mailboxInfo | Export-Csv -Path "c:\temp\MailboxRetentionInfo.csv" -NoTypeInformation