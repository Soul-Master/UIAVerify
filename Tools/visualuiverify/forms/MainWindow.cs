//---------------------------------------------------------------------------
//
// <copyright file="MainWindow" company="Microsoft">
// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.
// </copyright>
// 
//
// Description: main window of the application
//
//---------------------------------------------------------------------------

using System;
using System.Windows.Forms;
using System.Windows.Automation;
using VisualUIAVerify.Controls;
using VisualUIAVerify.Features;
using VisualUIAVerify.Misc;
using VisualUIAVerify.Win32;
using System.Runtime.InteropServices;
using Microsoft.Test.UIAutomation;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using Microsoft.Test.UIAutomation.Logging;

namespace VisualUIAVerify.Forms
{
    public partial class MainWindow : Form
    {
        private ElementHighlighter _highlighter;
        private HotKey[] _hotKeys;
        string _configFile = @"uiverify.config";                        // Configuration file to persist settings between runs
        string _filterFileName = "Win7BugFilter.xml";                   // Win7BugFilter.xml is the default XML. This can be overwritten
        ApplicationState _applicationState = new ApplicationState();    // Object that loads/saves the persisted settings between runs
        static string BaseTitle = "Visual UI Automation Verify";

        /// <summary>
        /// initializes main window
        /// </summary>
        public MainWindow(params string[] args)
        {
            // Obtain the state of the application on last shut down
            ApplicationStateDeserialize();

            //set logging handlers
            Misc.ApplicationLogger.LoggingProgress += new ApplicationLogger.LoggingProgressEventDelegate(ApplicationLogger_LoggingProgressEvent);

            InitializeComponent();

            //initialize menu items tags
            rectangleHighlightingToolStripMenuItem.Tag = ElementHighlighterFactory.BoundingRectangle;
            fadingRectangleHighlightingToolStripMenuItem.Tag = ElementHighlighterFactory.FadingBoundingRectangle;
            raysAndRectangleHighlightingToolStripMenuItem.Tag = ElementHighlighterFactory.BoundingRectangleAndRays;
            noneHighlightingToolStripMenuItem.Tag = ElementHighlighterFactory.None;
            
            showCategoriesToolStripButton.Tag = PropertySort.Categorized;
            sortAlphabeticalToolStripButton.Tag = PropertySort.Alphabetical;

            //Initialize HighLighting
            switch (this._applicationState.HighLight)
            {
                case ElementHighlighterFactory.BoundingRectangle:
                    {
                        highlightingToolStripMenuItem_Click(this.rectangleHighlightingToolStripMenuItem, new EventArgs());
                        break;
                    }
                case ElementHighlighterFactory.FadingBoundingRectangle:
                    {
                        highlightingToolStripMenuItem_Click(this.fadingRectangleHighlightingToolStripMenuItem, new EventArgs());
                        break;
                    }
                case ElementHighlighterFactory.BoundingRectangleAndRays:
                    {
                        highlightingToolStripMenuItem_Click(this.raysAndRectangleHighlightingToolStripMenuItem, new EventArgs());
                        break;
                    }
                case ElementHighlighterFactory.None:
                    {
                        highlightingToolStripMenuItem_Click(this.noneHighlightingToolStripMenuItem, new EventArgs());
                        break;
                    }
            }

            //Initialize TopMost
            if (this._applicationState.ModeAlwaysOnTop)
            {
                alwaysOnTopToolStripMenuItem.Checked = true;
                alwaysOnTopToolStripMenuItem_Click(this.alwaysOnTopToolStripMenuItem, new EventArgs());
            }

            //Initialize Tracking HoverMode
            if (this._applicationState.ModeHoverMode)
            {
                this.hoverModeToolStripMenuItem.Checked = true;
                hoverModeToolStripMenuItem_Click(this.hoverModeToolStripMenuItem, new EventArgs());
            }

            //Initialize Tracking Focus tracking
            if (this._applicationState.ModeFocusTracking)
            {
                this.focusTrackingToolStripMenuItem1.Checked = true;
                focusTrackingToolStripMenuItem1_Click(this.focusTrackingToolStripMenuItem1, new EventArgs());
            }

            //Initilize Automation Tests control
            _automationElementTree.RootElement = AutomationElement.RootElement;

            RegisterHotKeys();

            if (args == null || (Array.IndexOf<string>(args, "NOCLIENTSIDEPROVIDER") != -1))
            {
                UnloadClientSideProviders();
            }
            else
            {
                this.Text = string.Format("{0} : {1}", BaseTitle, "Client Side Provider");
            }
            TestRuns.BugFilterFile = TestRuns.FilterOutBugs ? _filterFileName : String.Empty;
        }

        
        private void ApplicationStateDeserialize()
        {

            _configFile = Path.Combine(Directory.GetCurrentDirectory(), _configFile);

            if (File.Exists(_configFile))
            {
                Stream stream = File.Open(_configFile, FileMode.Open);
                BinaryFormatter formatter = new BinaryFormatter();
                _applicationState = (ApplicationState)formatter.Deserialize(stream);
                stream.Close();
            }
        }

        private void ApplicationStateSerialize()
        {
            Stream stream = File.Open(_configFile, FileMode.Create);
            BinaryFormatter formatter = new BinaryFormatter();
            formatter.Serialize(stream, _applicationState);
            stream.Close();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="e"></param>
        protected override void OnClosed(EventArgs e)
        {
            ApplicationStateSerialize();
            UnregisterHotKeys();
            base.OnClosed(e);
        }

        private void StopHighlighting()
        {
            if (this._highlighter != null)
            {
                this._highlighter.Dispose();
                this._highlighter = null;
            }
        }

        private void ShowLog()
        {

        }

        #region hot keys hook


        private void RegisterHotKeys()
        {
            this._hotKeys = new HotKey[]
            {
                new HotKey("Ctrl+Shift", "F5", new EventHandler(this.refreshElementToolStripButton_Click)),

                new HotKey("Ctrl+Shift", "F6", new EventHandler(this.goToParentToolStripButton_Click)),
                new HotKey("Ctrl+Shift", "F7", new EventHandler(this.goToFirstChildToolStripButton_Click)),
                new HotKey("Ctrl+Shift", "F8", new EventHandler(this.goToNextSiblingToolStripButton_Click)),
                new HotKey("Ctrl+Shift", "F9", new EventHandler(this.goToPrevSiblingToolStripButton_Click)),
                new HotKey("Ctrl+Shift", "F10", new EventHandler(this.goToLastChildToolStripButton_Click)),

                new HotKey("Ctrl+Shift", "T", new EventHandler(this.alwaysOnTopHotKey)),
                new HotKey("Ctrl+Shift", "H", new EventHandler(this.hoverModeHotKey)),
                new HotKey("Ctrl+Shift", "F", new EventHandler(this.focusTrackingHotKey))
            };

            int index = 0;
            foreach (HotKey hotKey in this._hotKeys)
            {
                if (!UnsafeNativeMethods.RegisterHotKey(base.Handle, index, hotKey.Mask, hotKey.VKCode))
                {
                    MessageBox.Show(string.Format("Cannot register HotKey {0}", hotKey.Description), "Visual UI Automation Verify Error");
                }
                index++;
            }
        }

        private void UnregisterHotKeys()
        {
            int index = 0;
            foreach (HotKey hotKey in this._hotKeys)
            {
                UnsafeNativeMethods.UnregisterHotKey(base.Handle, index++);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="msg"></param>
        protected override void WndProc(ref Message msg)
        {
            if (msg.Msg == 0x312)
            {
                this._hotKeys[msg.WParam.ToInt32()].Handler(this, EventArgs.Empty);
            }
            base.WndProc(ref msg);
        }
        
        #endregion

        #region logging

        void ApplicationLogger_LoggingProgressEvent(string message, int percentage)
        {
            ShowLoggingMessage(message, percentage);
        }

        private delegate void ShowLoggingMessageDelegate(string message, int percentage);

        private void ShowLoggingMessage(string message, int percentage)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new ShowLoggingMessageDelegate(ShowLoggingMessage), message, percentage);
            }
            else
            {
                this._messageToolStrip.Visible = true;
                this._progressToolStrip.Visible = true;

                this._messageToolStrip.Text = message;
                this._progressToolStrip.Value = percentage;


                if (percentage == 100)
                {
                    this._messageToolStrip.Visible = false;
                    this._progressToolStrip.Visible = false;
                }
            }
        }

        #endregion

        #region events

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //user wants to quit application
            Application.Exit();
        }

        private void goToParentToolStripButton_Click(object sender, EventArgs e)
        {
            //if some currentTestTypeRootNode is selected then go to its parent

            AutomationElementTreeNode selectedNode = this._automationElementTree.SelectedNode;
            if (selectedNode != null)
            {
                this._automationElementTree.GoToParentFromNode(selectedNode);
            }
        }

        private void goToFirstChildToolStripButton_Click(object sender, EventArgs e)
        {
            //if some currentTestTypeRootNode is selected then go to its first child

            AutomationElementTreeNode selectedNode = this._automationElementTree.SelectedNode;
            if (selectedNode != null)
            {
                this._automationElementTree.GoToFirstChildFromNode(selectedNode);
            }
        }

        private void goToNextSiblingToolStripButton_Click(object sender, EventArgs e)
        {
            //if some currentTestTypeRootNode is selected then go to its next sibling

            AutomationElementTreeNode selectedNode = this._automationElementTree.SelectedNode;
            if (selectedNode != null)
            {
                this._automationElementTree.GoToNextSiblingFromNode(selectedNode);
            }
        }

        private void goToPrevSiblingToolStripButton_Click(object sender, EventArgs e)
        {
            //if some currentTestTypeRootNode is selected then go to its previous sibling

            AutomationElementTreeNode selectedNode = this._automationElementTree.SelectedNode;
            if (selectedNode != null)
            {
                this._automationElementTree.GoToPreviousSiblingFromNode(selectedNode);
            }
        }

        private void goToLastChildToolStripButton_Click(object sender, EventArgs e)
        {
            //if some currentTestTypeRootNode is selected then go to its last child

            AutomationElementTreeNode selectedNode = this._automationElementTree.SelectedNode;
            if (selectedNode != null)
            {
                this._automationElementTree.GoToLastChildFromNode(selectedNode);
            }
        }

        private void FocusTrackingToolStripMenuItem_CheckedChanged(object sender, EventArgs e)
        {
            if (FocusTrackingToolStripMenuItem.Checked)
                _automationElementTree.StartFocusTracing();
            else
                _automationElementTree.StopFocusTracing();
        }

        private void alwaysOnTopHotKey(object sender, EventArgs e)
        {
            alwaysOnTopToolStripMenuItem.Checked = !alwaysOnTopToolStripMenuItem.Checked;
            this.TopMost = alwaysOnTopToolStripMenuItem.Checked;
        }

        private void alwaysOnTopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.TopMost = alwaysOnTopToolStripMenuItem.Checked;
            this._applicationState.ModeAlwaysOnTop = alwaysOnTopToolStripMenuItem.Checked;
        }

        private void highlightingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem senderMenuItem = (ToolStripMenuItem)sender;

            if (senderMenuItem.Checked)
                return;

            this._applicationState.HighLight = (string)senderMenuItem.Tag;

            rectangleHighlightingToolStripMenuItem.Checked = false;
            fadingRectangleHighlightingToolStripMenuItem.Checked = false;
            raysAndRectangleHighlightingToolStripMenuItem.Checked = false;
            noneHighlightingToolStripMenuItem.Checked = false;
            senderMenuItem.Checked = true;

            StopHighlighting();

            switch (this._applicationState.HighLight)
            {
                case ElementHighlighterFactory.None:
                    {
                        break;
                    }
                default:
                    {
                        _highlighter = ElementHighlighterFactory.CreateHighlighterById(this._applicationState.HighLight, this._automationElementTree);
                        _highlighter.StartHighlighting();
                        break;
                    }
            }
        }

        private void _automationElementTree_SelectedNodeChanged(object sender, EventArgs e)
        {
            //selected currentTestTypeRootNode has been changed so notify change to AutomationTests Control
            AutomationElementTreeNode selectedNode = _automationElementTree.SelectedNode;
            AutomationElement selectedElement = null;

            if (selectedNode != null)
                selectedElement = selectedNode.AutomationElement;
            
            _automationElementPropertyGrid.AutomationElement = selectedElement;
        }

        private void hoverModeHotKey(object sender, EventArgs e)
        {
            hoverModeToolStripMenuItem.Checked = !hoverModeToolStripMenuItem.Checked;

            hoverModeToolStripMenuItem_Click(sender, e);
        }
        
        private void hoverModeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this._applicationState.ModeHoverMode = hoverModeToolStripMenuItem.Checked;

            if (hoverModeToolStripMenuItem.Checked)
                this._automationElementTree.StartHoverMode();
            else
                this._automationElementTree.StopHoverMode();
        }
        
        private void propertyPaneToolStripButton_Click(object sender, EventArgs e)
        {
            ToolStripButton button = sender as ToolStripButton;

            if (button != null)
            {
                PropertySort propertySort = (PropertySort)button.Tag;

                PropertySort newValue = _automationElementPropertyGrid.PropertySort;

                if (button.Checked)
                    newValue |= propertySort;
                else
                    newValue &= ~propertySort;

                _automationElementPropertyGrid.PropertySort = newValue;
            }
        }

        private void refreshElementToolStripButton_Click(object sender, EventArgs e)
        {
            AutomationElementTreeNode node = _automationElementTree.SelectedNode;
            if (node != null)
            {
                _automationElementTree.RefreshNode(node);
            }
        }

        private void focusTrackingHotKey(object sender, EventArgs e)
        {
            focusTrackingToolStripMenuItem1.Checked = !focusTrackingToolStripMenuItem1.Checked;

            focusTrackingToolStripMenuItem1_Click(sender, e);
        }

        private void focusTrackingToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this._applicationState.ModeFocusTracking = focusTrackingToolStripMenuItem1.Checked;

            if (focusTrackingToolStripMenuItem1.Checked)
                _automationElementTree.StartFocusTracing();
            else
                _automationElementTree.StopFocusTracing();

        }
        
        private void refreshPropertyPaneToolStripButton_Click(object sender, EventArgs e)
        {
            _automationElementPropertyGrid.RefreshValues();
        }

        private void expandAllToolStripButton_Click(object sender, EventArgs e)
        {
            _automationElementPropertyGrid.ExpandAll = expandAllToolStripButton.Checked;
        }
        
        private void aboutVisualUIAVerifyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new AboutWindow().ShowDialog(this);
        }

        #endregion

        private void UnmanagedProxiesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UnloadClientSideProviders();
        }

        // Set up pInvoke for UiaRegisterProviderCallback
        [DllImport("UIAutomationCore.dll", CharSet = CharSet.Unicode)]
        private static extern void UiaRegisterProviderCallback(IntPtr callback);

        /// <summary>
        /// Unload client side provider.  This will default to the MSAA Proxy.
        /// </summary>
        public void UnloadClientSideProviders()
        {
            // First, do something to ensure the proxy loading call has been made
            AutomationElement root = AutomationElement.RootElement;

            // Register a Null callback, this tells UI Automation to use the new proxies in Core
            UiaRegisterProviderCallback(IntPtr.Zero);

            VerifyWeHaveUIAutomation();

            this.UnmanagedProxiesToolStripMenuItem.Enabled = false;
            Text = string.Format("{0} : {1}", BaseTitle, "No Client Side Provider");

            AutomationElementTreeNode node = _automationElementTree.RootNode;
            if (node != null)
            {
                _automationElementTree.RefreshNode(node);
            }
        }

        /// <summary>
        /// Quick test to determine if we have a valid UIAutomationCore dynamic library.
        /// </summary>
        public static void VerifyWeHaveUIAutomation()
        {
            try
            {
                AutomationElement root = AutomationElement.RootElement;
            }
            catch (ArgumentException)
            {
                MessageBox.Show("Exception has occured that indicates you are trying to turn off 'Client Side Provider'.  You may not have the most recent version of UIAutomationCore.dll installed on your system.  Please contact your accessibility contact to find out how to obtain a newer version of UIAutomationCoredll.  Visual UIVerify / UIAutomationCore are now unstable.  Please restart Visual UIVerify to use the default Client Side Provider (Windows Vista/.NET Framework 3.0).");
            }
            catch (Exception error)
            {
                while (null != error.InnerException)
                {
                    error = error.InnerException;
                }
                MessageBox.Show(error.Message + error.GetType().ToString());
                throw error;
            }
        }
        
        private void saveLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (saveLogFileDialog.ShowDialog() == DialogResult.OK)
            {
                UIVerifyLogger.GenerateXMLLog(saveLogFileDialog.FileName);
            }
        }
    }
}
