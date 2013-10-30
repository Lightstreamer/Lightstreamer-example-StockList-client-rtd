#region License
/*
* Copyright 2013 Weswit Srl
*
* Licensed under the Apache License, Version 2.0 (the "License");
* you may not use this file except in compliance with the License.
* You may obtain a copy of the License at
*
* http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/
#endregion License

using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Net;
using Lightstreamer.DotNet.Client;

namespace Lightstreamer.DotNet.Client.Demo
{

    // This is the class handling the Lightstreamer Client,
    // the entry point for Lightstreamer update events.

    class LightstreamerClient
    {
        private string[] items;
        private string[] fields;

        private IRtdLightstreamerListener listener;
        private LSClient client;
        private SubscribedTableKey tableKey;
        private FlowForm flowForm = null;

        public LightstreamerClient(
            IRtdLightstreamerListener listener,
            string[] items, string[] fields,
            FlowForm flowForm)
        {
            if (listener == null)
            {
                throw new ArgumentNullException("listener is null");
            }
            this.items = items;
            this.fields = fields;
            this.listener = listener;
            this.flowForm = flowForm;
            client = new LSClient();
        }

        public void Stop()
        {
            if (tableKey != null)
                this.client.UnsubscribeTable(tableKey);
            tableKey = null;
            this.client.CloseConnection();
        }

        public void Start(string pushServerUrl, bool askCredentials)
        {
            ConnectionInfo connInfo = new ConnectionInfo();
            connInfo.PushServerUrl = pushServerUrl;
            connInfo.Adapter = "DEMO";
            StocklistConnectionListener ls = new StocklistConnectionListener(
                listener,this,pushServerUrl);

            bool connected = false;
            //this method will not exit until the openConnection returns without throwing an exception
            while (!connected)
            {
                try
                {
                    if (askCredentials)
                    {
                        if ( this.flowForm != null) 
                        {
                            String proxyUsr = flowForm.askProxyUsr();
                            String proxyPwd = flowForm.askProxyPwd();

                            IWebProxy proxy = WebRequest.DefaultWebProxy;
                            proxy.Credentials = new NetworkCredential(proxyUsr, proxyPwd);
                        }
                    }
                    //WebRequest.
                    this.client.OpenConnection(connInfo, ls);
                    connected = true;
                }
                catch (PushConnException e)
                {
                    listener.OnStatusChange(LightstreamerConnectionHandler.ERROR, e.Message);
                    if (e.Message.Contains("407"))
                    {
                        askCredentials = true;
                    }
                }
                catch (PushServerException e)
                {
                    listener.OnStatusChange(LightstreamerConnectionHandler.ERROR, e.Message);
                    if (e.Message.Contains("407"))
                    {
                        askCredentials = true;
                    }
                }
                catch (PushUserException e)
                {
                    listener.OnStatusChange(LightstreamerConnectionHandler.ERROR, e.Message);
                    if (e.Message.Contains("407"))
                    {
                        askCredentials = true;
                    }
                }

                if (!connected)
                {
                    Thread.Sleep(1000);
                }
            }

            SimpleTableInfo tableInfo = new ExtendedTableInfo(
            items, "MERGE", fields, true);
            tableInfo.DataAdapter = "QUOTE_ADAPTER";

            tableKey = client.SubscribeTable(tableInfo,
                new StocklistHandyTableListener(listener), false);
            
        }
    }

}
