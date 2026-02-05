using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.ComApi;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;

using NavisApp = Autodesk.Navisworks.Api.Application;
using NavisColor = Autodesk.Navisworks.Api.Color;

namespace ClashToVP
{
    // ═══════════════════════════════════════════════════════════════════════════
    // CLASH TO VIEWPOINT PLUGIN
    // Filters: Only NEW, ACTIVE, or REVIEWED clashes
    // 1. Place section plane at middle of FIRST clashing item
    // 2. Zoom to it
    // 3. Highlight it (RED color)
    // ═══════════════════════════════════════════════════════════════════════════
    [PluginAttribute("ClashToVP.IsolateSection", "ADSK",
        DisplayName = "Clash VP - Section",
        ToolTip = "Section at Item1 + Zoom + Highlight",
        ExtendedToolTip = "Processes only New/Active/Reviewed clashes. Section plane at middle of first clash item.")]
    [AddInPluginAttribute(AddInLocation.AddIn,
        Icon = "PluginIcon16.png",
        LargeIcon = "PluginIcon32.png")]
    public class ClashIsolateSectionPlugin : AddInPlugin
    {
        private string logPath = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "ClashToVP_Log.txt");

        private System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
        private ProgressForm progressForm = null;

        private void DebugLog(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            string elapsed = stopwatch.IsRunning ? $" [{stopwatch.ElapsedMilliseconds}ms]" : "";
            string logMessage = timestamp + elapsed + " " + message;
            try { System.IO.File.AppendAllText(logPath, logMessage + Environment.NewLine); } catch { }
            System.Windows.Forms.Application.DoEvents();
        }

        private void UpdateProgress(int current, int total, string message)
        {
            if (progressForm != null && !progressForm.IsDisposed)
                progressForm.UpdateProgress(current, total, message);
            System.Windows.Forms.Application.DoEvents();
        }

        private bool IsCancelled()
        {
            System.Windows.Forms.Application.DoEvents();
            return progressForm != null && progressForm.CancelRequested;
        }

        public override int Execute(params string[] parameters)
        {
            try { System.IO.File.WriteAllText(logPath, ""); } catch { }
            stopwatch.Restart();

            DebugLog("=== Clash VP - Section Started ===");

            try
            {
                Document doc = NavisApp.ActiveDocument;
                if (doc == null) { MessageBox.Show("No active document."); return 0; }

                DocumentClash clashDoc = doc.GetClash();
                if (clashDoc == null) { MessageBox.Show("Could not access Clash Detective."); return 0; }

                if (clashDoc.TestsData.Tests.Count == 0) { MessageBox.Show("No Clash Tests found."); return 0; }

                ClashTest selectedTest = ClashHelper.SelectClashTest(clashDoc.TestsData.Tests);
                if (selectedTest == null) { MessageBox.Show("No Clash Test selected."); return 0; }

                DebugLog("Selected test: " + selectedTest.DisplayName);
                RunProcess(doc, selectedTest);
            }
            catch (Exception ex)
            {
                DebugLog("ERROR: " + ex.Message);
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                if (progressForm != null && !progressForm.IsDisposed) { progressForm.Close(); progressForm = null; }
            }

            stopwatch.Stop();
            DebugLog("=== Finished ===");
            return 0;
        }

        private void RunProcess(Document doc, ClashTest test)
        {
            // Get only NEW, ACTIVE, or REVIEWED clashes
            List<ClashResult> clashResults = ClashHelper.GetFilteredClashResults(test);
            int total = clashResults.Count;

            if (total == 0)
            {
                MessageBox.Show("No clash results found with status: New, Active, or Reviewed.");
                return;
            }

            DebugLog($"Found {total} clashes with status New/Active/Reviewed");

            progressForm = new ProgressForm(total);
            progressForm.Show();

            ComApi.InwOpState10 comState = null;
            try { comState = ComApiBridge.State; } catch { }

            // Folder: "Clash VP_[test]_[date]"
            string folderName = "Clash VP_" + test.DisplayName + "_" + DateTime.Now.ToString("yyyyMMdd_HHmm");
            GroupItem folderRef = ClashHelper.CreateFolder(doc, folderName);

            int count = 0, successCount = 0;

            foreach (ClashResult result in clashResults)
            {
                if (IsCancelled()) break;
                count++;
                UpdateProgress(count, total, result.DisplayName + " [" + result.Status + "]");
                DebugLog("Processing " + count + "/" + total + ": " + result.DisplayName + " [Status: " + result.Status + "]");

                try
                {
                    ProcessClash(doc, comState, folderRef, result, count);
                    successCount++;
                }
                catch (Exception ex) { DebugLog("Error: " + ex.Message); }
            }

            // Cleanup
            ClashHelper.DisableSectioning(comState);
            doc.CurrentSelection.Clear();
            doc.Models.ResetAllPermanentMaterials();

            if (progressForm != null) { progressForm.Close(); progressForm = null; }
            MessageBox.Show($"Complete!\n\nProcessed: {successCount} / {total}\n(New, Active, Reviewed only)\n\nFolder: {folderName}");
        }

        private void ProcessClash(Document doc, ComApi.InwOpState10 comState, GroupItem folder, ClashResult result, int index)
        {
            DebugLog("  -> Getting clash items...");
            ModelItem item1 = result.Item1;
            ModelItem item2 = result.Item2;

            if (item1 == null)
            {
                DebugLog("  -> Item1 is null, skipping...");
                return;
            }

            // Get bounding box of FIRST item (Item1)
            BoundingBox3D item1BBox = item1.BoundingBox();
            if (item1BBox == null || item1BBox.IsEmpty)
            {
                DebugLog("  -> Item1 bounding box is empty, skipping...");
                return;
            }

            // Get center point of Item1 - this is where we place the section plane
            Point3D item1Center = item1BBox.Center;
            DebugLog($"  -> Item1 center: X={item1Center.X:F3}, Y={item1Center.Y:F3}, Z={item1Center.Z:F3}");

            // STEP 1: Highlight Item1 with RED color
            DebugLog("  -> Highlighting items RED...");
            ModelItemCollection itemsToHighlight = new ModelItemCollection();
            itemsToHighlight.Add(item1);
            if (item2 != null) itemsToHighlight.Add(item2);
            doc.Models.OverridePermanentColor(itemsToHighlight, new NavisColor(1.0, 0.0, 0.0));

            // STEP 2: Select Item1 for zoom
            DebugLog("  -> Selecting Item1...");
            doc.CurrentSelection.Clear();
            doc.CurrentSelection.Add(item1);
            System.Windows.Forms.Application.DoEvents();

            // STEP 3: Place section plane at the middle of Item1
            DebugLog("  -> Placing section plane at Item1 center...");
            ClashHelper.EnableSectioningAtPoint(comState, item1Center, item1BBox);
            System.Windows.Forms.Application.DoEvents();

            // STEP 4: Zoom to Item1
            DebugLog("  -> Zooming to Item1...");
            ClashHelper.ZoomToSelection(comState);
            System.Windows.Forms.Application.DoEvents();

            // STEP 5: Save the viewpoint
            DebugLog("  -> Saving viewpoint...");
            ClashHelper.SaveViewpoint(doc, folder, result.DisplayName);

            // Reset for next clash
            DebugLog("  -> Resetting for next...");
            ClashHelper.DisableSectioning(comState);
            doc.Models.ResetPermanentMaterials(itemsToHighlight);
            doc.CurrentSelection.Clear();
            System.Windows.Forms.Application.DoEvents();
        }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // HELPER CLASS
    // ═══════════════════════════════════════════════════════════════════════════
    public static class ClashHelper
    {
        public static ClashTest SelectClashTest(SavedItemCollection tests)
        {
            List<ClashTest> clashTests = new List<ClashTest>();
            foreach (SavedItem item in tests)
            {
                if (item is ClashTest ct) clashTests.Add(ct);
            }

            if (clashTests.Count == 0) return null;
            if (clashTests.Count == 1) return clashTests[0];

            using (Form form = new Form())
            {
                form.Text = "Select Clash Test";
                form.Width = 350; form.Height = 150;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false; form.MinimizeBox = false;

                ComboBox combo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Left = 20, Top = 20, Width = 290 };
                foreach (ClashTest ct in clashTests) combo.Items.Add(ct.DisplayName);
                combo.SelectedIndex = 0;

                Button btnOk = new Button { Text = "OK", DialogResult = DialogResult.OK, Left = 150, Top = 60, Width = 75 };
                Button btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Left = 235, Top = 60, Width = 75 };

                form.Controls.Add(combo);
                form.Controls.Add(btnOk);
                form.Controls.Add(btnCancel);
                form.AcceptButton = btnOk;
                form.CancelButton = btnCancel;

                if (form.ShowDialog() == DialogResult.OK)
                    return clashTests[combo.SelectedIndex];
            }
            return null;
        }

        /// <summary>
        /// Get clash results filtered by status: New, Active, or Reviewed only
        /// </summary>
        public static List<ClashResult> GetFilteredClashResults(ClashTest test)
        {
            List<ClashResult> results = new List<ClashResult>();

            foreach (SavedItem item in test.Children)
            {
                if (item is ClashResult cr)
                {
                    // Filter: Only include New, Active, or Reviewed clashes
                    if (cr.Status == ClashResultStatus.New ||
                        cr.Status == ClashResultStatus.Active ||
                        cr.Status == ClashResultStatus.Reviewed)
                    {
                        results.Add(cr);
                    }
                }
            }

            return results;
        }

        public static GroupItem CreateFolder(Document doc, string name)
        {
            FolderItem folder = new FolderItem { DisplayName = name };
            doc.SavedViewpoints.AddCopy(folder);
            return doc.SavedViewpoints.Value[doc.SavedViewpoints.Value.Count - 1] as GroupItem;
        }

        /// <summary>
        /// Enable sectioning with a single horizontal plane at the center Z of the item
        /// This places the section plane right through the middle of Item1
        /// </summary>
        public static void EnableSectioningAtPoint(ComApi.InwOpState10 comState, Point3D centerPoint, BoundingBox3D bbox)
        {
            if (comState == null) return;

            try
            {
                dynamic view = comState.CurrentView;
                dynamic clipPlanes = view.ClippingPlanes();

                // Clear existing planes
                while (clipPlanes.Count > 0) clipPlanes.RemovePlane(1);

                // Create a single horizontal section plane at the center Z of Item1
                // Plane equation: ax + by + cz + d = 0
                // For horizontal plane at Z = centerZ: 0*x + 0*y + 1*z - centerZ = 0
                // Normal pointing up (0, 0, 1), d = -centerZ

                double sectionZ = centerPoint.Z;

                dynamic plane = comState.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaClipPlane, null, null);
                plane.Plane.SetValue(0, 0, 1, -sectionZ);  // Horizontal plane at centerZ
                plane.Enabled = true;
                clipPlanes.AddPlane(plane);

                // Enable sectioning
                clipPlanes.Enabled = true;
            }
            catch { }
        }

        /// <summary>
        /// Disable sectioning
        /// </summary>
        public static void DisableSectioning(ComApi.InwOpState10 comState)
        {
            if (comState == null) return;
            try
            {
                dynamic view = comState.CurrentView;
                dynamic clipPlanes = view.ClippingPlanes();
                clipPlanes.Enabled = false;
                while (clipPlanes.Count > 0) clipPlanes.RemovePlane(1);
            }
            catch { }
        }

        /// <summary>
        /// Zoom to the current selection
        /// </summary>
        public static void ZoomToSelection(ComApi.InwOpState10 comState)
        {
            if (comState == null) return;
            try
            {
                comState.ZoomInCurViewOnCurSel();
            }
            catch { }
        }

        public static void SaveViewpoint(Document doc, GroupItem folder, string name)
        {
            Viewpoint currentVp = doc.CurrentViewpoint.ToViewpoint();
            SavedViewpoint savedVp = new SavedViewpoint(currentVp);
            savedVp.DisplayName = name;
            doc.SavedViewpoints.InsertCopy(folder, folder.Children.Count, savedVp);
        }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // PROGRESS FORM
    // ═══════════════════════════════════════════════════════════════════════════
    public class ProgressForm : Form
    {
        private ProgressBar progressBar;
        private Label lblStatus;
        private Label lblProgress;
        private Button btnCancel;
        private int totalItems;

        public bool CancelRequested { get; private set; } = false;

        public ProgressForm(int total)
        {
            totalItems = total;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Processing Clashes (New/Active/Reviewed)...";
            this.Width = 450; this.Height = 180;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false; this.MinimizeBox = false;
            this.ControlBox = false; this.TopMost = true;

            lblStatus = new Label { Text = "Initializing...", Left = 20, Top = 20, Width = 400, AutoEllipsis = true };
            lblProgress = new Label { Text = "0 / " + totalItems, Left = 20, Top = 45, Width = 400 };
            progressBar = new ProgressBar { Left = 20, Top = 70, Width = 395, Height = 25, Minimum = 0, Maximum = totalItems, Value = 0 };
            btnCancel = new Button { Text = "Cancel", Left = 170, Top = 105, Width = 100, Height = 30 };
            btnCancel.Click += (s, e) => { CancelRequested = true; btnCancel.Enabled = false; btnCancel.Text = "Cancelling..."; };

            this.Controls.Add(lblStatus);
            this.Controls.Add(lblProgress);
            this.Controls.Add(progressBar);
            this.Controls.Add(btnCancel);
        }

        public void UpdateProgress(int current, int total, string message)
        {
            if (this.InvokeRequired) { this.Invoke(new Action(() => UpdateProgress(current, total, message))); return; }
            progressBar.Value = Math.Min(current, progressBar.Maximum);
            lblProgress.Text = current + " / " + total;
            lblStatus.Text = message;
            this.Refresh();
        }
    }
}