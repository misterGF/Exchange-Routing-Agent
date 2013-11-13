    This program is meant to evaluate an email and determine if it is internal to the organization (domain name) or not.
    
    If it is not an internal message it will override the message's routing and send it to a destination server.
 
    Incoming messages are not evaluated.
    This transport agent is useful for Exch 2010 hosting edition where each organization is separate and a transport agent is needed.
   
    Last updated: 10/12/12 Gil
    
    Install Instructions
    1. Once the code has been compiled take the DLL and place it in a safe place on your hub server. ex: c:\program files\CustomRoutingAgent\CustomRoutingAgent.dll
    2. Open up EMS and submit the following cmd: Install-TransportAgent -Name "CustomRoutingAgent" -TransportAgentFactory "RoutingAgentOverride.mSRoutingAgentFactory" -AssemblyPath "c:\program files\CustomRoutingAgent\CustomRoutingAgent.dll"
    3. Enable agent: Enable-TransportAgent -Identity "CustomRoutingAgent" 