# Exchange Message Manipulation Plugin

This is an [Exchange SMTP Receive Agent] (https://technet.microsoft.com/en-us/library/bb125012(v=exchg.150).aspx) that allows Exchange Server Administrators to create PowerShell scripts to alter messages as the message is delivered to the Exchange Server.

Most of this code is derived from the example here: https://code.msdn.microsoft.com/Exchange-2013-Build-a-body-ed36ecb0 

## Architecture

 * Upon Enabling the TransportAgent, an event receiver is registered for [OnEndOfData](https://www.google.com/url?sa=t&rct=j&q=&esrc=s&source=web&cd=1&cad=rja&uact=8&ved=0ahUKEwjur6GByZnQAhWF44MKHeCXA1oQFggdMAA&url=https%3A%2F%2Fmsdn.microsoft.com%2Fen-us%2Flibrary%2Fmicrosoft.exchange.data.transport.smtp.smtpreceiveagent.onendofdata.aspx&usg=AFQjCNGw5ivmyyfOI3B1uGBjCgdRFEGM6Q&sig2=2BGhu-eQrs1IGSgFT_z99w&bvm=bv.137904068,d.amc)
* On each EndOfData event, the plugin enumerates all PSM1 files in the specified directory
* Foreach PSM1 file in the directory, the plugin passes the message body to the PSM1 file's ```Get-ShouldProcess``` function to determine if the plugin should process the message
* If ```Get-ShouldProcess``` returns true, the plugin passes the message body to the PSM1 file's ```Get-ProcessedMessage``` function, and writes the return of the function to the current event's messages output stream
* If ```Get-ShouldProcess``` returns false, the original message body is passed through in Exchange


## Installation

 * Complie the source, and copy the DLL to a protected folder
 * From Exchange Management Shell, run the command: 
    ```
    Net Stop MSExchangeTransport
    Install-Transportagent -Name "Body Conversion Sample" -AssemblyPath $EXDIR\BodyConversion.dll -TransportAgentFactory Microsoft.Exchange.Samples.Agents.BodyConversion.BodyConversionFactory```
    Enable-TransportAgent -Identity "Body Conversion Sample" 
    Get-TransportAgent -Identity "Body Conversion Sample" 
    Net Start MSExchangeTransport
    ```
  * Create PowerShell Modules in the directory 

