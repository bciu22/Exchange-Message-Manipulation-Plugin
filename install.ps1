$EXDIR="C:\TransportAgentSamples\BodyConversion"
Net Stop MSExchangeTransport

Write-Output "Creating directories..."
New-Item -Type Directory -path $EXDIR -ErrorAction SilentlyContinue

Write-Output "Copying files..."
Copy-Item BodyConversion.dll $EXDIR -Force

Write-Output "Registering agent..."
Install-Transportagent -Name "Body Conversion Sample" -AssemblyPath $EXDIR\BodyConversion.dll -TransportAgentFactory Microsoft.Exchange.Samples.Agents.BodyConversion.BodyConversionFactory

Write-Output "Enabling agent..."
Enable-TransportAgent -Identity "Body Conversion Sample" 
Get-TransportAgent -Identity "Body Conversion Sample" 

Net Start MSExchangeTransport

Write-Output "Installation complete. Please exit the Exchange Management Shell."
