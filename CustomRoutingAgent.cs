/*  This program is meant to evaluate an email and determine if it is internal to the organization (domain name) or not.
    If it is not an internal message it will overide the message's routing and send it to a destination server.
 
    Incoming messages are not evaluated.
    This transport agent is useful for Exch 2010 hosting edition where each organization is seperate and a transport agent is needed.
   
    Last updated: 10/12/12 Gil
    
    Install Instructions
    1. Once the code has been compiled take the DLL and place it in a safe place on your hub server. ex: c:\program files\CustomRoutingAgent\CustomRoutingAgent.dll
    2. Open up EMS and submit the following cmd: Install-TransportAgent -Name "CustomRoutingAgent" -TransportAgentFactory "RoutingAgentOverride.mSRoutingAgentFactory" -AssemblyPath "c:\program files\CustomRoutingAgent\CustomRoutingAgent.dll"
    3. Enable agent: Enable-TransportAgent -Identity "CustomRoutingAgent" 
 
*/

//Needed NameSpaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Email;
using Microsoft.Exchange.Data.Transport.Smtp;
using Microsoft.Exchange.Data.Transport.Routing;
using Microsoft.Exchange.Data.Common;


namespace RoutingAgentOverride
{
    public class mSRoutingAgentFactory : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            RoutingAgent myAgent = new ownRoutingAgent();

            return myAgent;
        }
    }
}

public class ownRoutingAgent : RoutingAgent
{
    public ownRoutingAgent()
    {
        //subscribe to different events
        base.OnResolvedMessage += new ResolvedMessageEventHandler(ownRoutingAgent_OnResolvedMessage);
        
    }

    void ownRoutingAgent_OnResolvedMessage(ResolvedMessageEventSource source, QueuedMessageEventArgs e)
    {
        try
            {
            //Declare log string
                String NextLine = String.Empty;
                String delivery = e.MailItem.InboundDeliveryMethod.ToString();
           
            //Declare the Destination Server connector
                RoutingDomain routingDomain = new RoutingDomain("myDestionationServer.com");
                DeliveryQueueDomain deliveryQueueDomain = new DeliveryQueueDomain();
                RoutingOverride myRoutingOverride = new RoutingOverride(routingDomain, deliveryQueueDomain);

            //Yank out the domain of the sender
                var senderDomain = e.MailItem.FromAddress.DomainPart;

            //Go through the recipients and check which ones are NOT internal-> Overwrite their routing
                foreach (EnvelopeRecipient recp in e.MailItem.Recipients) 
                    {
                       //Apply only to messages that are coming from the mailbox server. We are discarding incoming messages here. Which have a type of Smtp.
                        if (delivery.Equals("Mailbox"))
                                {
                                    //Do internal check (same domain)       
                                    if (!recp.Address.DomainPart.Equals(senderDomain, StringComparison.OrdinalIgnoreCase))
                                    {
                                        //This isn't an internal email. Trying to override route.
                                        try
                                        {
                                            source.SetRoutingOverride(recp, myRoutingOverride);
                                        }
                                        catch (Exception ex)
                                        {
                                            NextLine = "Failed to override setting:" + ex;
                                            
                                            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"c:\temp\transportLogs.txt", true))
                                            {
                                                file.WriteLine(NextLine);
                                            }
                                        }
                                    }
                            } //end of main if statement
                    } //end of for loop 

            } //end of try
        catch (Exception ex)
            {
                //Debugging log
                String NextLine = "I'm inside the main catch statement"+ ex;
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"c:\temp\transportLogs.txt", true))
                {
                    file.WriteLine(NextLine);
                }
            }
     }

}

