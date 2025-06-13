using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NAudio.CoreAudioApi;

namespace MicrophoneViewer
{
    public partial class MainForm : Form
    {
        private SplitContainer splitContainer;
        private TreeView treeView;
        private PropertyGrid propertyGrid;

        public MainForm()
        {
            InitializeComponents();
            Load += MainForm_Load;
        }

        private void InitializeComponents()
        {
            // Set up the split container
            splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 50
            };

            // Set up the tree view
            treeView = new TreeView
            {
                Dock = DockStyle.Fill
            };
            treeView.AfterSelect += TreeView_AfterSelect;

            // Set up the property grid
            propertyGrid = new PropertyGrid
            {
                Dock = DockStyle.Fill
            };

            // Add controls
            splitContainer.Panel1.Controls.Add(treeView);
            splitContainer.Panel2.Controls.Add(propertyGrid);
            Controls.Add(splitContainer);

            // Form properties
            Text = "Microphone Device Viewer";
            Width = 1200;
            Height = 600;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            PopulateTree();
        }

        private void PopulateTree()
        {
            var root = new TreeNode("Microphone Devices");
            var enumerator = new MMDeviceEnumerator();
            var devices = enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active);

            foreach (var device in devices)
            {
                var deviceNode = new TreeNode(device.FriendlyName) { Tag = device };

                // Enumerate APOs registered on the system (placeholder for device-specific mapping)
                var apos = GetAposFromRegistry(device);
                foreach (var apo in apos)
                {
                    var apoNode = new TreeNode(apo.FriendlyName) { Tag = apo };
                    deviceNode.Nodes.Add(apoNode);
                }

                root.Nodes.Add(deviceNode);
            }

            treeView.Nodes.Add(root);
            root.Expand();
        }

        private string ExtractDeviceGuid(string deviceId)
        {
            // The format is typically: {0.0.1.00000000}.{4d383790-1ef2-4496-9f11-de406c08f464}
            // We need to extract the second GUID part
            if (string.IsNullOrEmpty(deviceId))
                return string.Empty;

            int lastDotIndex = deviceId.LastIndexOf('.');
            if (lastDotIndex >= 0 && lastDotIndex < deviceId.Length - 1)
            {
                string guidPart = deviceId.Substring(lastDotIndex + 1);
                // Remove any remaining braces
                return guidPart;
            }

            return deviceId;
        }

        private ApoInfo[] GetAposFromRegistry(MMDevice device)
        {
            var result = new List<ApoInfo>();

            string extractedGuid = ExtractDeviceGuid(device.ID);
            System.Diagnostics.Debug.WriteLine($"Original device ID: {device.ID}");
            System.Diagnostics.Debug.WriteLine($"Extracted GUID: {extractedGuid}");

            // Try multiple known registry paths where APOs might be registered
            string[] possiblePaths = new string[] {
        @"SOFTWARE\Microsoft\Windows\CurrentVersion\Audio\Audio Processing Objects",
        @"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\" + extractedGuid + @"\FxProperties"
    };

            foreach (var path in possiblePaths)
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine($"Checking registry path: {path}");
                    using (var key = Registry.LocalMachine.OpenSubKey(path))
                    {
                        if (key == null)
                        {
                            System.Diagnostics.Debug.WriteLine($"Registry path not found: {path}");
                            continue;
                        }

                        foreach (var guidStr in key.GetSubKeyNames())
                        {
                            try
                            {
                                using (var sub = key.OpenSubKey(guidStr))
                                {
                                    if (sub == null) continue;

                                    var friendly = sub.GetValue("FriendlyName") as string ?? guidStr;
                                    var clsid = sub.GetValue("CLSID") as string ?? string.Empty;

                                    result.Add(new ApoInfo
                                    {
                                        Id = Guid.TryParse(guidStr, out Guid id) ? id : Guid.Empty,
                                        FriendlyName = friendly,
                                        Clsid = clsid,
                                        DeviceId = device.ID
                                    });
                                }
                            }
                            catch (Exception ex)
                            {
                                // Consider logging this exception
                                System.Diagnostics.Debug.WriteLine($"Error processing APO with ID {guidStr}: {ex.Message}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Consider logging this exception
                    System.Diagnostics.Debug.WriteLine($"Error accessing registry path {path}: {ex.Message}");
                }
            }

            return result.ToArray();
        }
        private void TreeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Tag is MMDevice device)
            {
                propertyGrid.SelectedObject = new MMDeviceInfo(device);
            }
            else
            {
                propertyGrid.SelectedObject = e.Node.Tag;
            }
        }
    }

    public class ApoInfo
    {
        public Guid Id { get; set; }
        public string FriendlyName { get; set; }
        public string Clsid { get; set; }
        public string DeviceId { get; set; }
        public override string ToString() => FriendlyName;
    }
}
