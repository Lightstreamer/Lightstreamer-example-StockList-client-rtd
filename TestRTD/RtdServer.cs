using Microsoft.Office.Interop.Excel;
using com.lightstreamer.client;
using System.Runtime.InteropServices;

using NLog;

namespace TestRTD
{
    public interface IRtdLightstreamerListener
    {
        // Lightstreamer Client connection status changes arrive here
        void OnStatusChange(int status, string message);
        // Lightstreamer Client data updates arrive here
        void OnItemUpdate(int itemPos, string itemName, ItemUpdate update);
        // Lightstreamer Client lost updates info arrives here
        void OnLostUpdate(int itemPos, string itemName, int lostUpdates);
    }

    class RtdUpdateQueueItem
    {
        public int TopicID;
        public ItemUpdate update;
        public string field;
        public string value;

        public RtdUpdateQueueItem(int TopicID, string field, ItemUpdate update)
        {
            this.TopicID = TopicID;
            this.field = field;
            this.update = update;
        }

        public RtdUpdateQueueItem(int TopicID, string field, string value)
        {
            this.TopicID = TopicID;
            this.field = field;
            this.value = value;
        }

    }

    /// <summary>
    /// This is the main class of the library and handles the RTD Server and its
    /// communication with Excel, the Lightstreamer Client library and its
    /// communication with Lightstreamer Server.
    /// </summary>
    [ComVisible(true)]
    [Guid("DB1797F5-7198-4411-8563-D05F4E904956")]
    [ProgId("lightstreamer.rtdnew23")]
    public class RtdServer : IServer, IRtdLightstreamerListener
    {
        private string pushServerUrl = "";
        private bool feedingToggle = true;
        // This is set to 0 when FormClose is called, and makes
        // Hearthbeat() returning a problem.
        private int serverAlive = 1;

        private string serverStatus = "DISCONNECTED";

        private string[] items = { };

        private string[] fields = { };

        // some internal caches used in this class
        private Dictionary<string, ItemUpdate> itemCache = new Dictionary<string, ItemUpdate>();
        private Dictionary<int, string[]> topicIdMap = new Dictionary<int, string[]>();
        private Dictionary<string, int> reverseTopicIdMap = new Dictionary<string, int>();
        private Dictionary<string, List<string>> subsDone = new Dictionary<string, List<string>>();
        private Dictionary<string, List<string>> subsWait = new Dictionary<string, List<string>>();

        // this is a simple queue used to store updates that this class has to send to
        // Excel when RefreshData() is called. It's flushed out on RefreshData() and
        // filled in when OnItemUpdate() is called by Lightstreamer Client.
        private Queue<RtdUpdateQueueItem> updateQueue = new Queue<RtdUpdateQueueItem>();
        // avoid going to infinity, the update queue can, at most, have 15000 elements.
        private int updateQueueMaxLength = 15000;
        // Excel RTD variable
        private IRTDUpdateEvent rtdUpdateEvent = null;

        private LSConnect lsClient = null;
        private Logger log = null;

        // default RtdServer ProgID
        internal const string RTD_PROG_ID = "lightstreamer.newrtd2";
        // set this to the network computer name the RtdServer will run on
        // if the application will only serve local Excel requests, leave
        // this blank
        private const string RTD_SERVER = "";
        private Guid Guid = Guid.NewGuid();

        private void callWaitingSubs()
        {
            foreach (string key in subsWait.Keys)
            {
                List<string> l = null;

                if (subsWait.TryGetValue(key, out l))
                {
                    string[] listOfFields = new string[l.Count];
                    int i = 0;

                    foreach (string f in l)
                    {
                        listOfFields[i++] = f;
                    }

                    lsClient.AddSubcribe(new string[] { key }, listOfFields);
                    subsDone.Add(key, l);

                    log.Info("ConnectData - recover waiting Subscription, itemName: " + key + ", field: " + listOfFields);
                }
            }

            subsWait.Clear();
        }

        private void resubOnReconnect()
        {
            foreach (string key in subsDone.Keys)
            {
                List<string> l = null;

                if (subsDone.TryGetValue(key, out l))
                {
                    string[] listOfFields = new string[l.Count];
                    int i = 0;

                    foreach (string f in l)
                    {
                        listOfFields[i++] = f;
                    }

                    lsClient.AddSubcribe(new string[] { key }, listOfFields);

                    log.Info("ConnectData - recover waiting Subscription, itemName: " + key + ", field: " + listOfFields);
                }
            }
        }

        public RtdServer()
        {
            
            var config = new NLog.Config.LoggingConfiguration();

            var logfile = new NLog.Targets.FileTarget("logfile") { FileName = "TestRTD.log" };

            config.AddRule(LogLevel.Info, LogLevel.Fatal, logfile);

            NLog.LogManager.Configuration = config;

            log = NLog.LogManager.GetLogger("newrtdexceldemo");

            log.Info("RtdServer - " + RTD_PROG_ID + " - ");

            LightstreamerClient.setLoggerProvider(new NLgWrapper());

            // setup Lightstreamer Client instance
            lsClient = new LSConnect(this);


            // temporary assignement
            pushServerUrl = "http://localhost:8080";

        }

        public object ConnectData(int TopicID, ref Array Strings, ref bool GetNewValues)
        {
            List<string> item_fields = null;

            if (Strings.Length > 1)
            {

                if (((string)Strings.GetValue(0)).Equals("CONFIG"))
                {
                    string ls_server_url = (string)Strings.GetValue(1);
                    string ls_server_port = (string)Strings.GetValue(2);
                    string ls_adapter_set = (string)Strings.GetValue(3);
                    string ls_adapter_name = (string)Strings.GetValue(4);

                    string ls_user = (string)Strings.GetValue(5);
                    string ls_password = (string)Strings.GetValue(6);

                    pushServerUrl = ls_server_url;
                    if (!ls_server_port.Equals(""))
                    {
                        pushServerUrl += ":" + ls_server_port;
                    }
                    lsClient.setUrl(pushServerUrl);
                    lsClient.setAdapterSet(ls_adapter_set);
                    lsClient.setAdapterName(ls_adapter_name);
                    lsClient.setUser(ls_user);
                    lsClient.setPassword(ls_password);
                    

                    // register association between Topic ID and itemName and field.
                    topicIdMap[TopicID] = new string[] { "CONFIG" };
                    reverseTopicIdMap["CONFIG"] = TopicID;

                    log.Info("Setup Connection info (" + this.serverStatus + "), url: " + ls_server_url + ", port: " + ls_server_port + ", set: " + ls_adapter_set + ", data adapter: " + ls_adapter_name);

                    return this.serverStatus;
                }

                if (((string)Strings.GetValue(0)).Equals("OPTIONS"))
                {
                    string ls_parameter = (string)Strings.GetValue(1);
                    string param_value = (string)Strings.GetValue(2);

                    log.Info("Request for option setting; parameter: " + ls_parameter + ", value: " + param_value + ".");

                    if (ls_parameter.Equals("max_frequency")) {
                        lsClient.setFrequency(param_value);
                    }

                    if (ls_parameter.Equals("forced_transport"))
                    {
                        lsClient.setForcedTransport(param_value);
                    }

                    if (ls_parameter.Equals("stalled_timeout"))
                    {
                        lsClient.setStalledTimeout(param_value);
                    }

                    if (ls_parameter.Equals("proxy"))
                    {
                        string param_value2 = (string)Strings.GetValue(2);
                        string param_value3 = (string)Strings.GetValue(2);

                        lsClient.setProxy(param_value, param_value2, param_value3);
                    }

                    return "";
                }

                if (((string)Strings.GetValue(0)).Equals("LAST"))
                {
                    topicIdMap[TopicID] = new string[] { "LAST" };
                    reverseTopicIdMap["LAST"] = TopicID;

                    string wait = "...";
                    return wait;
                }

                if (((string)Strings.GetValue(0)).Equals("PX"))
                {
                    topicIdMap[TopicID] = new string[] { "PX" };
                    reverseTopicIdMap["PX"] = TopicID;

                    string pxw = "0";
                    return pxw;
                }


                string itemName = (string)Strings.GetValue(0);
                string field = (string)Strings.GetValue(1);

                log.Info("ConnectData, topic: " + TopicID + ", itemName: " + itemName + ", field: " + field);

                // register association between Topic ID and itemName and field.
                topicIdMap[TopicID] = new string[] { itemName, field };
                reverseTopicIdMap[itemName + "@" + field] = TopicID;
                ItemUpdate update;

                if (subsDone.TryGetValue(itemName, out item_fields))
                {
                    if (item_fields.Contains(field))
                    {
                        // if item is already in cache, then send it out, otherwise just send N/A string
                        // Excel will get a valid string as soon as it will be available.
                        if (itemCache.TryGetValue(itemName, out update))
                        {
                            string value = update.getValue(field);

                            log.Info("ConnectData, topic: " + TopicID + ", itemName: " + itemName + ", returning to Excel: " + value);

                            return value;
                        }
                    }
                    else
                    {
                        log.Info("ConnectData - add2 Subscription, itemName: " + itemName + ", field: " + field);
                        
                        item_fields.Add(field);
                        if (lsClient.isConnected())
                        {
                            lsClient.AddSubcribe(new string[] { itemName }, new string[] { field });
                        }

                    }

                }
                else
                {
                    log.Info("ConnectData - add Subscription, itemName: " + itemName + ", field: " + field);
                    
                    List<string> myFields = new List<string>();
                    myFields.Add(field);
                    if (lsClient.isConnected())
                    {
                        lsClient.AddSubcribe(new string[] { itemName }, new string[] { field });
                        subsDone.Add(itemName, myFields);
                    }
                    else
                    {
                        List<string> l = null;
                        log.Info("ConnectData - waiting Subscription, itemName: " + itemName + ", field: " + field);
                        
                        if (subsWait.TryGetValue(itemName, out l))
                        {
                            l.Add(field);
                        }
                        else
                        {
                            subsWait.Add(itemName, myFields);
                        }
                    }

                }
                log.Info("ConnectData, topic: " + TopicID + ", itemName: " + itemName + ", returning to Excel: Wait...");

                string res = "Wait...";
                return res;
            }
            log.Info("ConnectData, topic: " + TopicID + ", returning to Excel: ERROR");
            return "ERROR";
        }

        public void DisconnectData(int TopicID)
        {
            // Excel is no longer interested in this topic ID
            log.Info("RtdServer DisconnectData called for TopicID: " + TopicID);
            
            /* drop from our dictionary */
            string[] data;
            if (topicIdMap.TryGetValue(TopicID, out data))
            {
                reverseTopicIdMap.Remove(data[0] + "@" + data[1]);
            }
            topicIdMap.Remove(TopicID);
        }

        public int Heartbeat()
        {
            log.Info("RtdServer Heartbeat called");
            
            // is the server still alive?
            return serverAlive;
        }

        public Array RefreshData(ref int TopicCount)
        {
            log.Info("RtdServer RefreshData called");
            
            // updates we will push at most
            int enqueuedUpdates = updateQueue.Count();
            int updatesCount = 0;
            object[,] data = new object[2, enqueuedUpdates];
            while (enqueuedUpdates > 0)
            {
                RtdUpdateQueueItem item;
                try
                {
                    item = updateQueue.Dequeue();
                }
                catch (InvalidOperationException)
                {
                    break;
                }
                // build the update object
                data[0, updatesCount] = item.TopicID;
                if (item.field.Equals("price"))
                {
                    double result = Convert.ToDouble(item.update.getValue(item.field), System.Globalization.CultureInfo.CreateSpecificCulture("en-US"));

                    data[1, updatesCount] = result.ToString("N2", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"));
                }
                else if (item.field.Equals(""))
                {
                    if (topicIdMap[item.TopicID][0].Equals("PX"))
                    {
                        double px = Convert.ToDouble(item.value, System.Globalization.CultureInfo.CreateSpecificCulture("en-US"));
                        data[1, updatesCount] = px.ToString("N2", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"));
                    }
                    else
                    {
                        data[1, updatesCount] = item.value;
                    }
                }
                else
                {
                    data[1, updatesCount] = item.update.getValue(item.field);
                }
                ++updatesCount;
                --enqueuedUpdates;
            }
            TopicCount = updatesCount;
            return data;
        }

        /// <summary>
        /// Used by RtdServerTestForm project (which is used to test this class)
        /// </summary>
        public void SimulateStart()
        {
            updateQueue.Clear();
            topicIdMap.Clear();
            reverseTopicIdMap.Clear();
        }

        /// <summary>
        /// Used by RtdServerTestForm project (which is used to test this class)
        /// </summary>
        public void SimulateTerminate()
        {
            this.ServerTerminate();
            serverAlive = 0;
        }

        public int ServerStart(IRTDUpdateEvent CallbackObject)
        {
            log.Info("RtdServer started");
            
            reverseTopicIdMap.Clear();
            updateQueue.Clear();
            topicIdMap.Clear();
            (new Thread(new ThreadStart(delegate () { lsClient.Start(false); }))).Start();
            rtdUpdateEvent = CallbackObject;
            return 1;
        }

        public void ServerTerminate()
        {
            log.Info("RtdServer terminated");
            
            rtdUpdateEvent = null;
            if (lsClient != null)
            {
                lsClient.Stop();
            }
            lsClient = null;
            if (rtdUpdateEvent != null)
            {
                rtdUpdateEvent.Disconnect();
            }
            updateQueue.Clear();
            topicIdMap.Clear();
            reverseTopicIdMap.Clear();
        }

        public void TerminateLightstreamer()
        {
            if (lsClient != null)
            {
                lsClient.Stop();
            }
            lsClient = null;
        }

        public void ToggleFeeding()
        {
            feedingToggle = !feedingToggle;
            if (feedingToggle)
                if (rtdUpdateEvent != null)
                    rtdUpdateEvent.UpdateNotify();
        }

        public void OnStatusChange(int status, string message)
        {
            log.Info("OnStatusChange, status: " + status + ", " + message);
            
            if (status == 3 || status == 4)
            {
                callWaitingSubs();
                resubOnReconnect();
            }

            if (status == 0)
            {
                this.serverStatus = "DISCONNECTED";
            }
            else if (status == 1)
            {
                this.serverStatus = "CONNECTING";
            }
            else if (status == 2)
            {
                this.serverStatus = "CONNECTED";
            }
            else if (status == 3)
            {
                this.serverStatus = "STREAMING";
            }
            else if (status == 4)
            {
                this.serverStatus = "POLLING";
            }
            else if (status == 5)
            {
                this.serverStatus = "STALLED";
            }
            else
            {
                this.serverStatus = "ERROR";
            }

            int topicId;
            if (reverseTopicIdMap.TryGetValue("CONFIG", out topicId))
            {
                updateQueue.Enqueue(new RtdUpdateQueueItem(topicId, "", this.serverStatus));
                if (rtdUpdateEvent != null)
                {
                    // notify Excel that updates exist
                    // if this fails, it means that Excel is not ready
                    rtdUpdateEvent.UpdateNotify();

                    log.Info("OnItemUpdate, Excel notified. HB interval: " + rtdUpdateEvent.HeartbeatInterval);
                }
            }

        }

        public void OnItemUpdate(int itemPos, string itemName, ItemUpdate update)
        {
            List<string> item_fields = null;

            log.Info("OnItemUpdate, pos: " + itemPos + ", name: " + itemName + ", update:" + update);
            
            if (!feedingToggle)
            {
                log.Info("OnItemUpdate, not feeding Excel.");
                
                return;
            }

            itemCache[itemName] = update;


            if (subsDone.TryGetValue(itemName, out item_fields))
            {
                bool updatesExist = false;
                foreach (string field in item_fields)
                {
                    try
                    {
                        if (update.isValueChanged(field))
                        {
                            if (update.getValue(field) != null)
                            {
                                // push to update queue
                                if (updateQueue.Count() < updateQueueMaxLength)
                                {
                                    int topicId;
                                    if (reverseTopicIdMap.TryGetValue(itemName + "@" + field, out topicId))
                                    {
                                        updatesExist = true;
                                        updateQueue.Enqueue(new RtdUpdateQueueItem(topicId, field, update));
                                        if (field.Equals("time"))
                                        {
                                            int topicId_Last;
                                            if (reverseTopicIdMap.TryGetValue("LAST", out topicId_Last))
                                            {
                                                updateQueue.Enqueue(new RtdUpdateQueueItem(topicId_Last, "", update.getValue(field)));
                                            }
                                        }
                                        if (field.Equals("last_price"))
                                        {
                                            int topicId_Last_px;
                                            if (reverseTopicIdMap.TryGetValue("PX", out topicId_Last_px))
                                            {
                                                updateQueue.Enqueue(new RtdUpdateQueueItem(topicId_Last_px, "", update.getValue(field)));
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (ArgumentException)
                    {
                        // Skip ... 
                    }
                }
                if (updatesExist && (rtdUpdateEvent != null))
                {
                    // notify Excel that updates exist
                    // if this fails, it means that Excel is not ready
                    rtdUpdateEvent.UpdateNotify();

                    log.Info("OnItemUpdate, Excel notified. HB interval: " + rtdUpdateEvent.HeartbeatInterval);
                }
            }
        }

        public void OnLostUpdate(int itemPos, string itemName, int lostUpdates)
        {
            log.Info("OnLostUpdate, pos: " + itemPos + ", name: " + itemName + ", lost: " + lostUpdates);
        }

    }
}