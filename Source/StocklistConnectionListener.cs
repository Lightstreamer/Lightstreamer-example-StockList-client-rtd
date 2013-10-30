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

using Lightstreamer.DotNet.Client;
using System.Threading;

namespace Lightstreamer.DotNet.Client.Demo
{

    class LightstreamerConnectionHandler
    {
        public const int DISCONNECTED = 0;
        public const int CONNECTING = 1;
        public const int CONNECTED = 2;
        public const int STREAMING = 3;
        public const int POLLING = 4;
        public const int STALLED = 5;
        public const int ERROR = 6;
    }

    class StocklistConnectionListener : IConnectionListener
    {

        private IRtdLightstreamerListener listener;
        private LightstreamerClient lsClient;
        private string pushServerUrl = null;
        private bool isPolling;

        public StocklistConnectionListener(IRtdLightstreamerListener listener, LightstreamerClient ls, String url)
        {
            if (listener == null)
            {
                throw new ArgumentNullException("listener");
            }
            this.listener = listener;

            this.lsClient = ls;
            this.pushServerUrl = url;
        }

        public void OnConnectionEstablished()
        {
            listener.OnStatusChange(LightstreamerConnectionHandler.CONNECTED,
                "Connected to Lightstreamer Server...");
        }

        public void OnSessionStarted(bool isPolling)
        {
            string message;
            int status;
            this.isPolling = isPolling;
            if (isPolling)
            {
                message = "Lightstreamer is pushing (smart polling mode)...";
                status = LightstreamerConnectionHandler.POLLING;
            }
            else
            {
                message = "Lightstreamer is pushing (streaming mode)...";
                status = LightstreamerConnectionHandler.STREAMING;
            }
            listener.OnStatusChange(status, message);
        }

        public void OnNewBytes(long b) { }

        public void OnDataError(PushServerException e)
        {
            listener.OnStatusChange(LightstreamerConnectionHandler.ERROR,
                "Data error");
        }

        public void OnActivityWarning(bool warningOn)
        {
            if (warningOn)
            {
                listener.OnStatusChange(LightstreamerConnectionHandler.STALLED,
                    "Connection stalled");
            }
            else
            {
                OnSessionStarted(this.isPolling);
            }
        }

        public void OnClose()
        {
            listener.OnStatusChange(LightstreamerConnectionHandler.DISCONNECTED,
                "Connection closed");
            if ( (this.lsClient != null) && (this.pushServerUrl != null) )
            {
                listener.OnStatusChange(LightstreamerConnectionHandler.DISCONNECTED,
                "Connection closed ... retrying ...");
                (new Thread(new ThreadStart(delegate() { this.lsClient.Start(this.pushServerUrl); }))).Start();
            }
          
        }

        public void OnEnd(int cause)
        {
            listener.OnStatusChange(LightstreamerConnectionHandler.DISCONNECTED,
                "Connection forcibly closed");


        }

        public void OnFailure(PushServerException e)
        {
            listener.OnStatusChange(LightstreamerConnectionHandler.ERROR,
                "Server failure");
        }

        public void OnFailure(PushConnException e)
        {
            listener.OnStatusChange(LightstreamerConnectionHandler.ERROR,
                "Connection failure");
        }
    }

}
