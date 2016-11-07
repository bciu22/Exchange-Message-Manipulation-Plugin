// ***************************************************************
// <copyright file="BodyConversion.cs" company="Microsoft">
//     Copyright (C) Microsoft Corporation.  All rights reserved.
// </copyright>
// <summary>
// Filters scripts out of email messages.
// </summary>
// ***************************************************************

namespace Microsoft.Exchange.Samples.Agents.BodyConversion
{
    using System;
    using System.IO;
    using System.Text;
    using System.Diagnostics;
    using Microsoft.Exchange.Data.Transport;
    using Microsoft.Exchange.Data.Transport.Smtp;
    using Microsoft.Exchange.Data.Transport.Email;
    using Microsoft.Exchange.Data.TextConverters;
    using System.Management.Automation;
    using System.Collections;
 

    /// <summary>
    /// Agent factory.
    /// </summary>
    public class BodyConversionFactory : SmtpReceiveAgentFactory
    {
        /// <summary>
        /// Creates a new BodyConversion.
        /// </summary>
        /// <param name="server">Exchange server.</param>
        /// <returns>A new BodyConversion.</returns>
        public override SmtpReceiveAgent CreateAgent(SmtpServer server)
        {
            return new BodyConversion();
        }
    }

    /// <summary>
    /// SmtpReceiveAgent for the BodyConversion sample.
    /// </summary>
    public class BodyConversion : SmtpReceiveAgent
    {
        /// <summary>
        /// An object to synchronize access to the log file.
        /// </summary>
        private object fileLock = new object();


        /// <summary>
        /// The constructor registers an end-of-data event handler.
        /// </summary>
        public BodyConversion()
        {
            Debug.WriteLine("[BodyConversion] Agent constructor");
            this.OnEndOfData += new EndOfDataEventHandler(this.OnEndOfDataHandler);
        }

        /// <summary>
        /// Invoked by Exchange when the entire message has been received.
        /// </summary>
        /// <param name="source">The source of this event.</param>
        /// <param name="eodArgs">The arguments for this event.</param>
        public void OnEndOfDataHandler(ReceiveMessageEventSource source, EndOfDataEventArgs eodArgs)
        {

            Debug.WriteLine("[BodyConversion] OnEndOfDataHandler is called");

            // The purpose of this sample is to show how TextConverters can be used
            // to update the body. This sample shows how to ensure that no active
            // content, such as scripts, can pass through in a message body.

            EmailMessage message = eodArgs.MailItem.Message;
            Stream originalBodyContent = null;
            Stream newBodyContent = null;
            StreamWriter writer = null;
            Encoding encoding;
            string charsetName;
            string[] ManipulationScripts = Directory.GetFiles("C:\\TransportAgentSamples","*.psm1");

            try
            {
                Body body = message.Body;
                BodyFormat bodyFormat = body.BodyFormat;
                newBodyContent = body.GetContentWriteStream();
                writer = new StreamWriter(newBodyContent);
                writer.Write("Found " + ManipulationScripts.Length.ToString() + " Scripts.  Listed below: " +Environment.NewLine);

                foreach (string Script in ManipulationScripts)
                {
                    PowerShell PSI =  PowerShell.Create();
                    string scripttext = System.IO.File.ReadAllText(Script);

                    writer.Write(scripttext);
                    /*PSI.AddScript(,false);
                    PSI.Invoke();
                    PSI.Commands.Clear();
                    PSI.AddCommand("Get-ShouldProcess").AddParameter("MailItem",eodArgs.MailItem);
                    var results = PSI.Invoke();
                    writer.Write(results.ToString()+Environment.NewLine+Environment.NewLine);*/
                   
                    writer.Write(Script + Environment.NewLine);
                }
                writer.Flush();
                writer.Close();
            }
            finally
            {
              
                if (originalBodyContent != null)
                {
                    originalBodyContent.Close();
                }

                if (newBodyContent != null)
                {
                    newBodyContent.Close();
                }
        

            }
            
        }
    }
}
