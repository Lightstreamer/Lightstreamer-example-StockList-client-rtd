using com.lightstreamer.client;
using NLog;
using System;

namespace TestRTD
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
    internal class StocklistConnectionListener : ClientListener
    {
        private IRtdLightstreamerListener listener;
        private LSConnect lsClient;
        private Logger log = null;

        public StocklistConnectionListener(IRtdLightstreamerListener listener, LSConnect ls)
        {
            if (listener == null)
            {
                throw new ArgumentNullException("listener");
            }
            this.listener = listener;

            this.lsClient = ls;

            log = NLog.LogManager.GetLogger("newrtdexceldemo");
        }

        public void onListenEnd()
        {
            // .
        }

        public void onListenStart()
        {
            // .
        }

        public void onPropertyChange(string property)
        {
            // .
        }

        public void onServerError(int errorCode, string errorMessage)
        {
            listener.OnStatusChange(LightstreamerConnectionHandler.ERROR, "Server failure: " + errorMessage);
        }

        public void onStatusChange(string status)
        {

            log.Info("onStatusChange: " + status);

            if (status.Equals("DISCONNECTED"))
            {
                listener.OnStatusChange(LightstreamerConnectionHandler.DISCONNECTED, "Connection forcibly closed");
            } else if (status.StartsWith("CONNECTING"))
            {
                listener.OnStatusChange(LightstreamerConnectionHandler.CONNECTING, "Connecting ... ");
            }
            else if (status.StartsWith("CONNECTED:") && status.EndsWith("-STREAMING") )
            {
                listener.OnStatusChange(LightstreamerConnectionHandler.STREAMING, "Lightstreamer is pushing (streaming mode) ... ");
            }
            else if (status.StartsWith("CONNECTED:") && status.EndsWith("-POLLING"))
            {
                listener.OnStatusChange(LightstreamerConnectionHandler.CONNECTED, "Lightstreamer is pushing (smart polling mode) ... ");
            }
            else if(status.StartsWith("STALLED"))
            {
                listener.OnStatusChange(LightstreamerConnectionHandler.STALLED, "Connection stalled");
            } else if (status.StartsWith("DISCONNECTED:WILL"))
            {
                listener.OnStatusChange(LightstreamerConnectionHandler.DISCONNECTED, "Connection closed ... retrying ... ");
            } else if (status.StartsWith("DISCONNECTED:TRYING-RECOVERY"))
            {
                listener.OnStatusChange(LightstreamerConnectionHandler.DISCONNECTED, " ... recovering ... ");
            }
        }
    }
}