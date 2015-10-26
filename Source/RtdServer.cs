#region License
/*
* Copyright (c) Lightstreamer Srl
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
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Lightstreamer.DotNet.Client;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;

namespace Lightstreamer.DotNet.Client.Demo
{

    public interface IRtdLightstreamerListener
    {
        // Lightstreamer Client connection status changes arrive here
        void OnStatusChange(int status, string message);
        // Lightstreamer Client data updates arrive here
        void OnItemUpdate(int itemPos, string itemName, IUpdateInfo update);
        // Lightstreamer Client lost updates info arrives here
        void OnLostUpdate(int itemPos, string itemName, int lostUpdates);
    }

    class RtdUpdateQueueItem
    {
        public int TopicID;
        public IUpdateInfo update;
        public string field;
        public string value;

        public RtdUpdateQueueItem(int TopicID, string field, IUpdateInfo update)
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
    [ProgId(RtdServer.RTD_PROG_ID), ComVisible(true)]
    public class RtdServer : IRtdServer, IRtdLightstreamerListener
    {

        private string pushServerUrl = null;
        private bool feedingToggle = true;
        // This is set to 0 when FormClose is called, and makes
        // Hearthbeat() returning a problem.
        private int serverAlive = 1;

        private string serverStatus = "DISCONNECTED";

        private string[] items = { };

        private string[] fields = { };

        // some internal caches used in this class
        private Dictionary<string, IUpdateInfo> itemCache = new Dictionary<string, IUpdateInfo>();
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

        private LightstreamerClient lsClient = null;
        private FlowForm flowForm = null;

        // default RtdServer ProgID
        internal const string RTD_PROG_ID = "lightstreamer.rtdexceldemo";
        // set this to the network computer name the RtdServer will run on
        // if the application will only serve local Excel requests, leave
        // this blank
        private const string RTD_SERVER = "";


        private void callWaitingSubs() 
        {
            foreach(string key in subsWait.Keys)
            {
                List<string> l = null;

                if (subsWait.TryGetValue(key, out l))
                {
                    string[] listOfFields = new string[l.Count];
                    int i = 0;

                    foreach(string f in l) 
                    {
                        listOfFields[i++] = f;
                    }

                    lsClient.AddSubcribe(new string[] { key }, listOfFields);
                    subsDone.Add(key, l);

                    flowForm.AppendExcelLog("ConnectData - recover waiting Subscription, itemName: " + key + ", field: " + listOfFields);
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

                    flowForm.AppendExcelLog("ConnectData - recover waiting Subscription, itemName: " + key + ", field: " + listOfFields);
                }
            }
        }

        public RtdServer()
        {
            flowForm = new FlowForm(this);
            flowForm.Activate();

            // setup Lightstreamer Client instance
            lsClient = new LightstreamerClient(this, null, null, this.flowForm);

            flowForm.Show();
            flowForm.BringToFront();

            pushServerUrl = "http://localhost:80";

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

                    pushServerUrl = ls_server_url;
                    if (!ls_server_port.Equals(""))
                    {
                        pushServerUrl += ":" + ls_server_port;
                    }
                    lsClient.setUrl(pushServerUrl);
                    lsClient.setAdapterSet(ls_adapter_set);
                    lsClient.setAdapterName(ls_adapter_name);

                    // register association between Topic ID and itemName and field.
                    topicIdMap[TopicID] = new string[] { "CONFIG" };
                    reverseTopicIdMap["CONFIG"] = TopicID;

                    flowForm.AppendExcelLog("Setup Connection info (" + this.serverStatus + "), url: " + ls_server_url + ", port: " + ls_server_port + ", set: " + ls_adapter_set + ", data adapter: " + ls_adapter_name);

                    return this.serverStatus;
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

                flowForm.AppendExcelLog("ConnectData, topic: " + TopicID + ", itemName: " +
                    itemName + ", field: " + field);

                // register association between Topic ID and itemName and field.
                topicIdMap[TopicID] = new string[] { itemName, field };
                reverseTopicIdMap[itemName + "@" + field] = TopicID;
                IUpdateInfo update;

                if (subsDone.TryGetValue(itemName, out item_fields))
                {
                    if (item_fields.Contains(field))
                    {
                        // if item is already in cache, then send it out, otherwise just send N/A string
                        // Excel will get a valid string as soon as it will be available.
                        if (itemCache.TryGetValue(itemName, out update))
                        {
                            string value = update.GetNewValue(field);
                            flowForm.AppendExcelLog("ConnectData, topic: " + TopicID + ", itemName: " +
                                itemName + ", returning to Excel: " + value);
                            return value;
                        }
                    }
                    else
                    {
                        flowForm.AppendExcelLog("ConnectData - add2 Subscription, itemName: " + itemName + ", field: " + field);
                        
                        item_fields.Add(field);
                        
                        if (lsClient.isConnected())
                        {
                            lsClient.AddSubcribe(new string[] { itemName }, new string[] { field });
                        }

                    }

                }
                else
                {
                    flowForm.AppendExcelLog("ConnectData - add Subscription, itemName: " + itemName + ", field: " + field);
                    List<string> myFields = new List<string>( );
                    myFields.Add(field);
                    if (lsClient.isConnected())
                    {
                        lsClient.AddSubcribe(new string[] { itemName }, new string[] { field });
                        subsDone.Add(itemName, myFields);
                    }
                    else 
                    {
                        List<string> l = null;
                        flowForm.AppendExcelLog("ConnectData - waiting Subscription, itemName: " + itemName + ", field: " + field);

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
                flowForm.AppendExcelLog("ConnectData, topic: " + TopicID + ", itemName: " +
                    itemName + ", returning to Excel: Wait...");
                string res = "Wait...";
                return res;
            }
            flowForm.AppendExcelLog("ConnectData, topic: " + TopicID +
                ", returning to Excel: ERROR");
            return "ERROR";
        }

        public void DisconnectData(int TopicID)
        {
            // Excel is no longer interested in this topic ID
            flowForm.AppendExcelLog("RtdServer DisconnectData called for TopicID: " + TopicID);
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
            flowForm.AppendExcelLog("RtdServer Heartbeat called");
            // is the server still alive?
            return serverAlive;
        }

        public Array RefreshData(ref int TopicCount)
        {
            flowForm.AppendExcelLog("RtdServer RefreshData called");

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
                    double result = Convert.ToDouble(item.update.GetNewValue(item.field), System.Globalization.CultureInfo.CreateSpecificCulture("en-US"));
                    data[1, updatesCount] = result.ToString("N2", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"));
                }
                else if (item.field.Equals("")) {
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
                    data[1, updatesCount] = item.update.GetNewValue(item.field);
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
            flowForm.AppendExcelLog("RtdServer started");
            reverseTopicIdMap.Clear();
            updateQueue.Clear();
            topicIdMap.Clear();
            (new Thread(new ThreadStart(delegate() { lsClient.Start(false); }))).Start();
            rtdUpdateEvent = CallbackObject;
            return 1;
        }

        public void ServerTerminate()
        {
            flowForm.AppendExcelLog("RtdServer terminated");
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
            flowForm.AppendLightstreamerLog("OnStatusChange, status: " + status + ", " +
                message);
            flowForm.UpdateConnectionStatusLabel(message);

            if (status == 3 || status == 4) 
            {
                callWaitingSubs();
                resubOnReconnect();
            }

            if ( status == 0 ) {
                this.serverStatus = "DISCONNECTED";
            } else if ( status == 1 ) {
                this.serverStatus = "CONNECTING";
            } else if ( status == 2 ) {
                this.serverStatus = "CONNECTED";
            } else if ( status == 3 ) {
                this.serverStatus = "STREAMING";
            } else if ( status == 4 ) {
                this.serverStatus = "POLLING";
            } else if ( status == 5 ) {
                this.serverStatus = "STALLED";
            } else {
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
                    flowForm.AppendExcelLog("OnItemUpdate, Excel notified. HB interval: " + rtdUpdateEvent.HeartbeatInterval);
                }
            }

        }

        public void OnItemUpdate(int itemPos, string itemName, IUpdateInfo update)
        {
            List<string> item_fields = null;

            flowForm.AppendLightstreamerLog("OnItemUpdate, pos: " + itemPos + ", name: " +
                itemName + ", update:" + update);
            flowForm.TickItemsCountLabel();

            if (!feedingToggle)
            {
                flowForm.AppendLightstreamerLog("OnItemUpdate, not feeding Excel.");
                return;
            }

            itemCache[itemName] = update;


            if (subsDone.TryGetValue(itemName, out item_fields)) 
            {
                bool updatesExist = false;
                foreach(string field in item_fields)
                {
                    try
                    {
                        if (update.IsValueChanged(field))
                        {
                            if (update.GetNewValue(field) != null)
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
                                                updateQueue.Enqueue(new RtdUpdateQueueItem(topicId_Last, "", update.GetNewValue(field)));
                                              }
                                        }
                                        if (field.Equals("last_price"))
                                        {
                                            int topicId_Last_px;
                                            if (reverseTopicIdMap.TryGetValue("PX", out topicId_Last_px))
                                            {
                                                updateQueue.Enqueue(new RtdUpdateQueueItem(topicId_Last_px, "", update.GetNewValue(field)));
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    } catch (ArgumentException ae) 
                    {
                        // Skip ... 
                    }
                }
                if (updatesExist && (rtdUpdateEvent != null))
                {
                    // notify Excel that updates exist
                    // if this fails, it means that Excel is not ready
                    rtdUpdateEvent.UpdateNotify();
                    flowForm.AppendExcelLog("OnItemUpdate, Excel notified. HB interval: " + rtdUpdateEvent.HeartbeatInterval);
                }
            }
        }

        public void OnLostUpdate(int itemPos, string itemName, int lostUpdates)
        {
            flowForm.AppendLightstreamerLog("OnLostUpdate, pos: " + itemPos + ", name: " +
                itemName + ", lost: " + lostUpdates);
        }

    }
}
