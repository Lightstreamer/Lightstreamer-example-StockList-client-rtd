namespace TestRTD
{
    using com.lightstreamer.client;
    using NLog;
    using System;
    using System.Threading;

    // This is the class handling the Lightstreamer Client,
    // the entry point for Lightstreamer update events.

    class LSConnect
    {
        private string? adapter_set = null;
        private string? adapter_name = null;
        private string? pushServerUrl = null;
        private string? user = null;
        private string? password = null;
        private IRtdLightstreamerListener listener;
        private LightstreamerClient client;
        private Subscription tableInfo;
        private bool connected = false;
        private Logger log = null;
        private string max_freq = "";

        public LSConnect(
            IRtdLightstreamerListener listener)
        {
            if (listener == null)
            {
                throw new ArgumentNullException("listener is null");
            }

            log = NLog.LogManager.GetLogger("newrtdexceldemo");

            this.listener = listener;
            client = new LightstreamerClient(null, null);
            client.addListener(new StocklistConnectionListener(listener, this));
            log.Info("New LSConnect.");
        }

        public bool isConnected()
        {
            return this.connected;
        }

        public void setAdapterSet(string adapter_set)
        {
            this.adapter_set = adapter_set;
        }

        public void setUrl(string url)
        {
            this.pushServerUrl = url;
            log.Info(" ... push url: " + url);
        }

        public void setAdapterName(string adapter_name)
        {
            this.adapter_name = adapter_name;
        }

        public void setUser(string user)
        {
            this.user = user;
        }

        public void setPassword(string pwd)
        {
            this.password = pwd;
        }

        public void AddSubcribe(string[] items, string[] fields)
        {
            log.Info("Add Subscription: " + items[0]);

            tableInfo = new Subscription("MERGE", items, fields);
            tableInfo.DataAdapter = this.adapter_name;
            tableInfo.RequestedSnapshot = "yes";
                
            if ( !this.max_freq.Equals("") )
            {
                tableInfo.RequestedMaxFrequency = this.max_freq;
            }

            tableInfo.addListener(new StocklistSubListener(listener));

            client.subscribe(tableInfo);
        }

        public void Stop()
        {
            if (tableInfo != null)
                this.client.unsubscribe(tableInfo);
            tableInfo = null;
            this.client.disconnect();
        }

        public void Disconnected()
        {
            this.connected = false;
        }

        public void Start(bool askCredentials)
        {
            log.Info("Start connecting to push server ... ");
            //this method will not exit until the openConnection returns without throwing an exception
            while (!connected)
            {
                if (this.pushServerUrl != null)
                {
                    client.connectionDetails.ServerAddress = this.pushServerUrl;
                    client.connectionDetails.AdapterSet = this.adapter_set;

                    if (this.user != null)
                    {
                        if (this.user != "")
                        {
                            client.connectionDetails.User = this.user;
                        }
                    }

                    if (this.password != null)
                    {
                        if (this.password != "")
                        {
                            client.connectionDetails.Password = this.password;
                        }
                    }

                    try
                    {
                        if (askCredentials)
                        {
                            //if (this.flowForm != null)
                            //{
                            //    String proxyUsr = flowForm.askProxyUsr();
                            //    String proxyPwd = flowForm.askProxyPwd();

                            //    IWebProxy proxy = WebRequest.DefaultWebProxy;
                            //    proxy.Credentials = new NetworkCredential(proxyUsr, proxyPwd);
                            //}
                        }
                        //WebRequest.

                        log.Info("Connect.");

                        this.client.connect();
                        connected = true;
                    }
                    catch (Exception e)
                    {
                        listener.OnStatusChange(LightstreamerConnectionHandler.ERROR, e.Message);
                        if (e.Message.Contains("407"))
                        {
                            askCredentials = true;
                        }
                    }
                }

                if (!connected)
                {
                    Thread.Sleep(1000);
                }
            }

        }

        internal void setFrequency(string param_value)
        {
            this.max_freq = param_value;
        }

        internal void setForcedTransport(string param_value)
        {
            client.connectionOptions.ForcedTransport = param_value;
        }
    }
}