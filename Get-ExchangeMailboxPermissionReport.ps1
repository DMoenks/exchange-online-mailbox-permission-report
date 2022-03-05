param([Parameter(Mandatory=$true)]
        [string]$credentialName,
        [Parameter(Mandatory=$true)]
        [string]$mailAddress)

$results = @{}

function Add-Result ([string]$Folder, [string]$User, [string]$Rights)
{
    if (-not $results.ContainsKey($Folder))
    {
        $results.Add($Folder, @{})
    }
    if (-not $results[$Folder].ContainsKey($User))
    {
        $results[$Folder].Add($User, [System.Collections.Generic.List[string]]::new())
    }
    if (-not $results[$Folder][$User].Contains($Rights))
    {
        $results[$Folder][$User].Add($Rights)
    }
}

Connect-ExchangeOnline -Credential (Get-AutomationPSCredential -Name $credentialName) -CommandName 'Get-Mailbox', 'Get-MailboxPermission', 'Get-MailboxFolderPermission', 'Get-MailboxFolderStatistics' | Out-Null

if ($null -ne ($mailbox = Get-Mailbox $mailAddress))
{
    foreach ($permission in $mailbox.GrantSendOnBehalfTo)
    {
        Add-Result -Folder '' -User $permission -Rights 'SendOnBehalf'
    }
    foreach ($permission in $mailbox | Get-MailboxPermission | Where-Object{$_.IsInherited -eq $false -and $_.Deny -eq $false})
    {
        Add-Result -Folder '' -User $permission.User -Rights $permission.AccessRights
    }
    foreach ($permission in Get-MailboxFolderPermission $mailAddress | Where-Object{-not ($_.User -in @('Anonymous', 'Default') -and $_.AccessRights -eq 'None')})
    {
        Add-Result -Folder '\' -User $permission.User -Rights $permission.AccessRights
    }
    foreach ($folder in $mailbox | Get-MailboxFolderStatistics | Where-Object{$_.FolderType -in @('Inbox', 'Outbox', 'SentItems', 'DeletedItems', 'Notes', 'Calendar', 'Contacts', 'Tasks', 'User Created')})
    {
        $folderPath = $folder.FolderPath.Replace('/', '\')
        foreach ($permission in Get-MailboxFolderPermission "${mailAddress}:$folderPath" | Where-Object{-not ($_.User -in @('Anonymous', 'Default') -and $_.AccessRights -eq 'None')})
        {
            if ($null -eq $permission.SharingPermissionFlags)
            {
                Add-Result -Folder $folderPath -User $permission.User -Rights $permission.AccessRights
            }
            else
            {
                Add-Result -Folder $folderPath -User $permission.User -Rights "$($permission.AccessRights) ($($permission.SharingPermissionFlags))"
            }
        }
    }
    '<html>' | Write-Output
        '<head>' | Write-Output
            '<style>' | Write-Output
                'table, th, td {' | Write-Output
                    'border: 1px solid black;' | Write-Output
                    'border-collapse: collapse;' | Write-Output
                '}' | Write-Output
                'th, td {' | Write-Output
                    'padding: 5px;' | Write-Output
                    'text-align: left;' | Write-Output
                    'vertical-align: top;' | Write-Output
                '}' | Write-Output
            '</style>' | Write-Output
        '</head>' | Write-Output
        '<body>' | Write-Output
            '<table>' | Write-Output
                '<tr>' | Write-Output
                    '<th>Folder</th>' | Write-Output
                    '<th>User</th>' | Write-Output
                    '<th>Rights</th>' | Write-Output
                '</tr>' | Write-Output
                foreach ($folder in $results.Keys | Sort-Object)
                {
                    '<tr>' | Write-Output
                    "<td rowspan=$($results[$folder].Keys.Count)>$folder</td>" | Write-Output
                    foreach ($user in $results[$folder].Keys | Sort-Object)
                    {
                        foreach ($right in $results[$folder][$user] | Sort-Object)
                        {
                            "<td>$user</td>" | Write-Output
                            "<td>$right</td>" | Write-Output
                            '</tr>' | Write-Output
                        }
                    }
                }
            '</table>' | Write-Output
        '</body>' | Write-Output
    '</html>' | Write-Output
}
