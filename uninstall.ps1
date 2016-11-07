$EXDIR="C:\TransportAgentSamples\BodyConversion"
Net Stop MSExchangeTransport

Write-Output "Disabling agent..."
Disable-TransportAgent -Identity "Body Conversion Sample" -Confirm:$false

Write-Output "Uninstalling agent..."
Uninstall-TransportAgent -Identity "Body Conversion Sample" -Confirm:$false

Write-Output "Deleting directories and files..."
Remove-Item $EXDIR\* -Force -Recurse -ErrorAction SilentlyContinue
Remove-Item $EXDIR -Force -Recurse -ErrorAction SilentlyContinue

Net Start MSExchangeTransport

Write-Output "Uninstall complete."
