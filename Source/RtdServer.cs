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
using System.Linq;
using System.Text;
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
        public RtdUpdateQueueItem(int TopicID, string field, IUpdateInfo update)
        {
            this.TopicID = TopicID;
            this.field = field;
            this.update = update;
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

        // Lightstreamer Server information, make this point to a valid Lightstreamer
        // server URL.
        private const string pushServerHost = "http://push.lightstreamer.com";
        // port, if different from 80
        private const string pushServerPort = "";
        private string pushServerUrl = null;
        private bool feedingToggle = true;
        // This is set to 0 when FormClose is called, and makes
        // Hearthbeat() returning a problem.
        private int serverAlive = 1;

        // Lightstreamer items and their fields served by this library
        private string[] items = {"item1", "item2", "item3", "item4", "item5", "item6", "item7", "item8",
                                     "item9", "item10", "item11", "item12", "item13", "item14", "item15",
                                     "item16", "item17", "item18", "item19", "item20", "item21", "item22",
                                     "item23", "item24", "item25", "item26", "item27", "item28",
                                     "item29", "item30" };
        private string[] fields = {"stock_name", "last_price", "time", "pct_change", "bid_quantity", "bid",
                                      "ask", "ask_quantity", "min", "max", "ref_price", "open_price" };


        // some internal caches used in this class
        private Dictionary<string, IUpdateInfo> itemCache = new Dictionary<string, IUpdateInfo>();
        private Dictionary<int, string[]> topicIdMap = new Dictionary<int, string[]>();
        private Dictionary<string, int> reverseTopicIdMap = new Dictionary<string, int>();

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
        internal const string RTD_PROG_ID = "Lightstreamer.RtdExcelDemo";
        // set this to the network computer name the RtdServer will run on
        // if the application will only serve local Excel requests, leave
        // this blank
        private const string RTD_SERVER = "";

        public RtdServer()
        {
            // setup Lightstreamer Client instance
            lsClient = new LightstreamerClient(this, items, fields);

            flowForm = new FlowForm(this);
            flowForm.Activate();
            flowForm.Show();
            flowForm.BringToFront();

            pushServerUrl = pushServerHost;
            if (!pushServerPort.Equals(""))
            {
                pushServerUrl += ":" + pushServerPort;
            }
        }

        public object ConnectData(int TopicID, ref Array Strings, ref bool GetNewValues)
        {

            if (Strings.Length > 1)
            {
                string itemName = (string)Strings.GetValue(0);
                string field = (string)Strings.GetValue(1);

                flowForm.AppendExcelLog("ConnectData, topic: " + TopicID + ", itemName: " +
                    itemName + ", field: " + field);

                // register association between Topic ID and itemName and field.
                topicIdMap[TopicID] = new string[] { itemName, field };
                reverseTopicIdMap[itemName + "@" + field] = TopicID;
                IUpdateInfo update;
                // if item is already in cache, then send it out, otherwise just send N/A string
                // Excel will get a valid string as soon as it will be available.
                if (itemCache.TryGetValue(itemName, out update))
                {
                    string value = update.GetNewValue(field);
                    flowForm.AppendExcelLog("ConnectData, topic: " + TopicID + ", itemName: " +
                        itemName + ", returning to Excel: " + value);
                    return value;
                }
                flowForm.AppendExcelLog("ConnectData, topic: " + TopicID + ", itemName: " +
                    itemName + ", returning to Excel: Wait...");
                string wait = "Wait...";
                return wait;
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
                data[1, updatesCount] = item.update.GetNewValue(item.field);
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
            lsClient = new LightstreamerClient(this, items, fields);
            lsClient.Start(pushServerUrl);
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
            (new Thread(new ThreadStart(delegate() { lsClient.Start(pushServerUrl); }))).Start();
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
        }

        public void OnItemUpdate(int itemPos, string itemName, IUpdateInfo update)
        {

            flowForm.AppendLightstreamerLog("OnItemUpdate, pos: " + itemPos + ", name: " +
                itemName + ", update:" + update);
            flowForm.TickItemsCountLabel();

            if (!feedingToggle)
            {
                flowForm.AppendLightstreamerLog("OnItemUpdate, not feeding Excel.");
                return;
            }

            itemCache[itemName] = update;

            bool updatesExist = false;
            for (int i = 0; i < fields.Length; i++)
            {
                string field = fields[i];
                if (update.IsValueChanged(field))
                {
                    // push to update queue
                    if (updateQueue.Count() < updateQueueMaxLength)
                    {
                        int topicId;
                        if (reverseTopicIdMap.TryGetValue(itemName + "@" + field, out topicId))
                        {
                            updatesExist = true;
                            updateQueue.Enqueue(new RtdUpdateQueueItem(topicId, field, update));
                        }
                    }
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

        public void OnLostUpdate(int itemPos, string itemName, int lostUpdates)
        {
            flowForm.AppendLightstreamerLog("OnLostUpdate, pos: " + itemPos + ", name: " +
                itemName + ", lost: " + lostUpdates);
        }

    }
}
