function Get-ShouldProcess {
    param (
        [MailItem]$MailItem
    )
    $MailItem.Message.Body | Out-File -FilePath 'C:\TransportAgentSamples\message.txt'
    return $false
}

function Get-ProcessedMessage {
        param (
        [MailItem]$MailItem
    )
    return $MailItems
}
