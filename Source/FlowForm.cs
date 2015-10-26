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
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using Microsoft.VisualBasic;

namespace Lightstreamer.DotNet.Client.Demo
{
    public partial class FlowForm : Form
    {

        private int updatesReceived = 0;
        private RtdServer rtdServer = null;

        public FlowForm(RtdServer rtdServer)
        {
            this.rtdServer = rtdServer;
            InitializeComponent();
        }

        private void DebugForm_Load(object sender, EventArgs e)
        {

        }


        public void UpdateConnectionStatusLabel(string text)
        {
            string message = "Connection status: " + text;
            lblStatus.Invoke(new ThreadStart(
                delegate() { lblStatus.Text = message; }));
        }

        public void TickItemsCountLabel()
        {
            updatesReceived++; /* can overflow !*/
            string message = "Lightstreamer items: " + updatesReceived;
            lblItemCounter.Invoke(new ThreadStart(
                delegate() { lblItemCounter.Text = message; }));
        }

        public void AppendLightstreamerLog(string message)
        {
            try
            {
                tbLightstreamer.Invoke(new ThreadStart(
                    delegate() {
                        if (tbLightstreamer.TextLength > 5000)
                        {
                            tbLightstreamer.Clear();
                        }
                        tbLightstreamer.AppendText(message + "\r\n"); 
                    }));
            }
            catch (InvalidOperationException)
            {
                /* do nothing */
            }
        }

        public void AppendExcelLog(string message)
        {
            try
            {
                tbExcel.Invoke(new ThreadStart(
                    delegate() {
                        if (tbExcel.TextLength > 5000)
                        {
                            tbExcel.Clear();
                        }
                        tbExcel.AppendText(message + "\r\n"); 
                    }));
            }
            catch (InvalidOperationException)
            {
                /* do nothing */
            }
        }

        public System.String askProxyUsr() 
        {
            return Interaction.InputBox("Please provide username for Proxy Authentication:", "Proxy Credentials", "user");
        }

        public System.String askProxyPwd()
        {
            return Interaction.InputBox("Please provide password for Proxy Authentication:", "Proxy Credentials", "pwd");
        }

        private void cbxToggleStream_CheckedChanged(object sender, EventArgs e)
        {
            rtdServer.ToggleFeeding();
        }

        private void FlowForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            rtdServer.TerminateLightstreamer();
        }

        private void FlowForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // deny user to close the User interface. The process would be kept alive.
            // In this case, the process is Excel itself, so be really careful.
            // If you are going to allow User Interface close event, make sure to
            // not terminate Lightstreamer and the RtdServer by commenting
            // rtdServer.TerminateLightstreamer() in FlowForm_FormClosed()
            AppendExcelLog("User interface is closing, blocking");
            e.Cancel = true;
        }

    }
}
