﻿/*
 * Created by SharpDevelop.
 * User: Alexander Petrovskiy
 * Date: 10/3/2014
 * Time: 8:47 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */

namespace Tmx.Client
{
    using System;
    using System.Net;
	using Spring.Rest.Client;
    using Tmx.Core.Types.Remoting;
    using Tmx.Interfaces.Exceptions;
    using Tmx.Interfaces.Remoting;
	using Tmx.Interfaces.Server;
    using Tmx.Core;
	using System.Collections.Generic;
    
    /// <summary>
    /// Description of StatusSender.
    /// </summary>
    public class StatusSender
    {
        readonly RestTemplate _restTemplate;
        
        public StatusSender(RestRequestCreator requestCreator)
        {
            _restTemplate = requestCreator.GetRestTemplate();
        }
        
        public virtual void Send(string status)
        {
            // TODO: add an error handler (??)
            try {
                _restTemplate.Put(UrnList.TestClients_Root + "/" + ClientSettings.Instance.ClientId + "/status", new DetailedStatus(status));
            }
            catch (Exception e) {
                throw new SendingDetailedStatusException("Failed to send detailed status. " + e.Message);
            }
        }
    }
}