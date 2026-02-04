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
using NavisView = Autodesk.Navisworks.Api.View;

namespace ClashToVP
{
    // ═══════════════════════════════════════════════════════════════════════════
    // CLASH TO VIEWPOINT PLUGIN
    // Places section box directly at the clash point (shows plane axis)
    // Zooms tightly to clash items, marks with RED color
    // Folder: "Clash VP_[test]_[date]"
    // ═══════════════════════════════════════════════════════════════════════════
    [PluginAttribute("ClashToVP.IsolateSection", "ADSK",
        DisplayName = "Clash VP - Section",
        ToolTip = "Section Box at Clash Point",
        ExtendedToolTip = "Places section box directly at clash point, clash items colored RED")]
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
            List<ClashResult> clashResults = ClashHelper.GetClashResults(test);
            int total = clashResults.Count;
            if (total == 0) { MessageBox.Show("No clash results found."); return; }

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
                UpdateProgress(count, total, result.DisplayName);
                DebugLog("Processing " + count + "/" + total + ": " + result.DisplayName);

                try
                {
                    ProcessClash(doc, comState, folderRef, result, count);
                    successCount++;
                }
                catch (Exception ex) { DebugLog("Error: " + ex.Message); }
            }

            // Cleanup - Disable section planes
            ClashHelper.DisableSectionPlanes(doc, comState);
            doc.CurrentSelection.Clear();

            if (progressForm != null) { progressForm.Close(); progressForm = null; }
            MessageBox.Show($"Complete!\n\nSuccess: {successCount} / {total}\n\nFolder: {folderName}");
        }

        private void ProcessClash(Document doc, ComApi.InwOpState10 comState, GroupItem folder, ClashResult result, int index)
        {
            DebugLog("  -> Collecting items...");
            ModelItem item1 = result.Item1;
            ModelItem item2 = result.Item2;
            if (item1 == null && item2 == null) return;

            ModelItemCollection clashItems = new ModelItemCollection();
            if (item1 != null) clashItems.Add(item1);
            if (item2 != null) clashItems.Add(item2);

            BoundingBox3D bbox = ClashHelper.GetBoundingBox(clashItems);
            if (bbox == null || bbox.IsEmpty) return;

            // Get the exact clash point (center of the bounding box of clashing items)
            Point3D clashPoint = bbox.Center;
            DebugLog($"  -> Clash point: X={clashPoint.X:F3}, Y={clashPoint.Y:F3}, Z={clashPoint.Z:F3}");

            // Color clash items RED
            DebugLog("  -> Coloring items...");
            ClashHelper.ColorItems(doc, item1, item2);

            // Select clash items
            DebugLog("  -> Selecting clash items...");
            doc.CurrentSelection.Clear();
            doc.CurrentSelection.CopyFrom(clashItems);
            System.Windows.Forms.Application.DoEvents();

            // Create section BOX centered exactly at the clash point
            // This positions the section plane origin/axis right at the clash
            DebugLog("  -> Creating section box at clash point...");
            ClashHelper.CreateSectionBoxAtClashPoint(doc, comState, clashPoint, bbox);
            System.Windows.Forms.Application.DoEvents();

            // Zoom tightly to the clash point
            DebugLog("  -> Zooming to clash point...");
            ClashHelper.ZoomToClashPoint(doc, comState, clashPoint, bbox);
            System.Windows.Forms.Application.DoEvents();

            // Save viewpoint
            DebugLog("  -> Saving viewpoint...");
            ClashHelper.SaveViewpoint(doc, folder, "Clash " + index + " - " + result.DisplayName);

            // Reset for next clash
            DebugLog("  -> Resetting for next...");
            ClashHelper.DisableSectionPlanes(doc, comState);
            doc.Models.ResetPermanentMaterials(clashItems);
            doc.CurrentSelection.Clear();
            System.Windows.Forms.Application.DoEvents();
        }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // HELPER CLASS - Shared functionality
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

        public static List<ClashResult> GetClashResults(ClashTest test)
        {
            List<ClashResult> results = new List<ClashResult>();
            foreach (SavedItem item in test.Children)
            {
                if (item is ClashResult cr) results.Add(cr);
            }
            return results;
        }

        public static GroupItem CreateFolder(Document doc, string name)
        {
            FolderItem folder = new FolderItem { DisplayName = name };
            doc.SavedViewpoints.AddCopy(folder);
            return doc.SavedViewpoints.Value[doc.SavedViewpoints.Value.Count - 1] as GroupItem;
        }

        public static BoundingBox3D GetBoundingBox(ModelItemCollection items)
        {
            bool init = false;
            double minX = 0, minY = 0, minZ = 0, maxX = 0, maxY = 0, maxZ = 0;
            foreach (ModelItem item in items)
            {
                BoundingBox3D b = item.BoundingBox();
                if (b != null && !b.IsEmpty)
                {
                    if (!init)
                    {
                        minX = b.Min.X; minY = b.Min.Y; minZ = b.Min.Z;
                        maxX = b.Max.X; maxY = b.Max.Y; maxZ = b.Max.Z;
                        init = true;
                    }
                    else
                    {
                        minX = Math.Min(minX, b.Min.X); minY = Math.Min(minY, b.Min.Y); minZ = Math.Min(minZ, b.Min.Z);
                        maxX = Math.Max(maxX, b.Max.X); maxY = Math.Max(maxY, b.Max.Y); maxZ = Math.Max(maxZ, b.Max.Z);
                    }
                }
            }
            return init ? new BoundingBox3D(new Point3D(minX, minY, minZ), new Point3D(maxX, maxY, maxZ)) : null;
        }

        public static void ColorItems(Document doc, ModelItem item1, ModelItem item2)
        {
            ModelItemCollection clashItems = new ModelItemCollection();
            if (item1 != null) clashItems.Add(item1);
            if (item2 != null) clashItems.Add(item2);

            if (clashItems.Count > 0)
            {
                doc.Models.OverridePermanentColor(clashItems, new NavisColor(1.0, 0.0, 0.0)); // Red
            }
        }

        /// <summary>
        /// Creates a section BOX centered exactly at the clash point
        /// This places the section plane origin/gizmo right at the clash location
        /// Uses OrientedBox in JSON format which centers the box at the specified coordinates
        /// </summary>
        public static void CreateSectionBoxAtClashPoint(Document doc, ComApi.InwOpState10 comState, Point3D clashPoint, BoundingBox3D bbox)
        {
            // Calculate box size based on clash bounding box with some padding
            double sizeX = bbox.Max.X - bbox.Min.X;
            double sizeY = bbox.Max.Y - bbox.Min.Y;
            double sizeZ = bbox.Max.Z - bbox.Min.Z;

            // Add padding around the clash (make box slightly larger than clash)
            double padding = Math.Max(Math.Max(sizeX, sizeY), sizeZ) * 0.5;
            if (padding < 0.5) padding = 0.5; // Minimum padding

            // Create box centered on clash point
            double minX = clashPoint.X - (sizeX / 2) - padding;
            double minY = clashPoint.Y - (sizeY / 2) - padding;
            double minZ = clashPoint.Z - (sizeZ / 2) - padding;
            double maxX = clashPoint.X + (sizeX / 2) + padding;
            double maxY = clashPoint.Y + (sizeY / 2) + padding;
            double maxZ = clashPoint.Z + (sizeZ / 2) + padding;

            // Try .NET API first with JSON using OrientedBox (section box mode)
            try
            {
                NavisView activeView = doc.ActiveView;
                if (activeView != null)
                {
                    // Use OrientedBox format - this creates a section BOX centered at the clash
                    // The Box property contains [[minX, minY, minZ], [maxX, maxY, maxZ]]
                    // Rotation is [0,0,0] for axis-aligned box
                    string clippingPlanesJson = @"{
                        ""Type"": ""ClipPlaneSet"",
                        ""Version"": 1,
                        ""OrientedBox"": {
                            ""Type"": ""OrientedBox3D"",
                            ""Version"": 1,
                            ""Box"": [[" +
                                minX.ToString(System.Globalization.CultureInfo.InvariantCulture) + "," +
                                minY.ToString(System.Globalization.CultureInfo.InvariantCulture) + "," +
                                minZ.ToString(System.Globalization.CultureInfo.InvariantCulture) + "],[" +
                                maxX.ToString(System.Globalization.CultureInfo.InvariantCulture) + "," +
                                maxY.ToString(System.Globalization.CultureInfo.InvariantCulture) + "," +
                                maxZ.ToString(System.Globalization.CultureInfo.InvariantCulture) + @"]],
                            ""Rotation"": [0, 0, 0]
                        },
                        ""Enabled"": true
                    }";

                    activeView.SetClippingPlanes(clippingPlanesJson);
                    System.Windows.Forms.Application.DoEvents();
                    return;
                }
            }
            catch { }

            // Fallback to COM API - create 6 planes forming a box around the clash
            CreateSectionBoxUsingComAPI(comState, minX, minY, minZ, maxX, maxY, maxZ);
        }

        /// <summary>
        /// Fallback method using COM API for section box creation
        /// Creates 6 clipping planes forming a box around the clash point
        /// </summary>
        private static void CreateSectionBoxUsingComAPI(ComApi.InwOpState10 comState,
            double minX, double minY, double minZ, double maxX, double maxY, double maxZ)
        {
            if (comState == null) return;

            try
            {
                dynamic view = comState.CurrentView;
                dynamic clipPlanes = view.ClippingPlanes();

                // Clear existing planes
                while (clipPlanes.Count > 0) clipPlanes.RemovePlane(1);

                // Add 6 planes to form a box around the clash point
                // Plane equation: ax + by + cz + d = 0

                // Right plane (+X): normal (-1,0,0), d = maxX
                AddClipPlane(comState, clipPlanes, -1, 0, 0, maxX);

                // Left plane (-X): normal (1,0,0), d = -minX
                AddClipPlane(comState, clipPlanes, 1, 0, 0, -minX);

                // Back plane (+Y): normal (0,-1,0), d = maxY
                AddClipPlane(comState, clipPlanes, 0, -1, 0, maxY);

                // Front plane (-Y): normal (0,1,0), d = -minY
                AddClipPlane(comState, clipPlanes, 0, 1, 0, -minY);

                // Top plane (+Z): normal (0,0,-1), d = maxZ
                AddClipPlane(comState, clipPlanes, 0, 0, -1, maxZ);

                // Bottom plane (-Z): normal (0,0,1), d = -minZ
                AddClipPlane(comState, clipPlanes, 0, 0, 1, -minZ);

                clipPlanes.Enabled = true;
            }
            catch { }
        }

        private static void AddClipPlane(ComApi.InwOpState10 comState, dynamic clipPlanes,
            double a, double b, double c, double d)
        {
            try
            {
                dynamic plane = comState.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaClipPlane, null, null);
                plane.Plane.SetValue(a, b, c, d);
                plane.Enabled = true;
                clipPlanes.AddPlane(plane);
            }
            catch { }
        }

        /// <summary>
        /// Disables all section planes
        /// </summary>
        public static void DisableSectionPlanes(Document doc, ComApi.InwOpState10 comState)
        {
            // Try .NET API first
            try
            {
                NavisView activeView = doc.ActiveView;
                if (activeView != null)
                {
                    string disabledJson = @"{
                        ""Type"": ""ClipPlaneSet"",
                        ""Version"": 1,
                        ""Enabled"": false
                    }";

                    activeView.SetClippingPlanes(disabledJson);
                    return;
                }
            }
            catch { }

            // Fallback to COM API
            DisableSectionPlanesComAPI(comState);
        }

        private static void DisableSectionPlanesComAPI(ComApi.InwOpState10 comState)
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
        /// Zooms tightly to the clash point with a close-up view
        /// Camera positioned to look at the clash from an isometric angle
        /// </summary>
        public static void ZoomToClashPoint(Document doc, ComApi.InwOpState10 comState, Point3D clashPoint, BoundingBox3D bbox)
        {
            System.Windows.Forms.Application.DoEvents();

            // Calculate the size of the clash area for determining zoom distance
            double sizeX = bbox.Max.X - bbox.Min.X;
            double sizeY = bbox.Max.Y - bbox.Min.Y;
            double sizeZ = bbox.Max.Z - bbox.Min.Z;
            double clashSize = Math.Max(Math.Max(sizeX, sizeY), sizeZ);

            // Set a tight zoom distance based on clash size
            // Use a small multiplier for very close zoom
            double zoomDistance = Math.Max(clashSize * 2.0, 0.5); // Very close zoom

            try
            {
                // Create viewpoint looking at clash point from isometric angle
                Viewpoint vp = doc.CurrentViewpoint.ToViewpoint();

                // Camera angles for isometric view (looking down at clash)
                double angleH = Math.PI / 4;  // 45 degrees horizontal
                double angleV = Math.PI / 5;  // ~36 degrees vertical (looking down)

                // Position camera at calculated distance from clash point
                Point3D cameraPos = new Point3D(
                    clashPoint.X + zoomDistance * Math.Cos(angleH) * Math.Cos(angleV),
                    clashPoint.Y + zoomDistance * Math.Sin(angleH) * Math.Cos(angleV),
                    clashPoint.Z + zoomDistance * Math.Sin(angleV)
                );

                vp.Position = cameraPos;
                vp.Projection = ViewpointProjection.Perspective;

                // Apply the viewpoint
                doc.CurrentViewpoint.CopyFrom(vp);
                System.Windows.Forms.Application.DoEvents();

                // Use ZoomBox on a tight bounding box around the clash for precise framing
                double padding = clashSize * 0.3;
                BoundingBox3D tightBox = new BoundingBox3D(
                    new Point3D(clashPoint.X - padding, clashPoint.Y - padding, clashPoint.Z - padding),
                    new Point3D(clashPoint.X + padding, clashPoint.Y + padding, clashPoint.Z + padding)
                );

                vp = doc.CurrentViewpoint.ToViewpoint();
                vp.ZoomBox(tightBox);
                doc.CurrentViewpoint.CopyFrom(vp);
                System.Windows.Forms.Application.DoEvents();
            }
            catch
            {
                // Fallback: Try COM API zoom
                if (comState != null)
                {
                    try
                    {
                        comState.ZoomInCurViewOnCurSel();
                        System.Windows.Forms.Application.DoEvents();
                    }
                    catch { }
                }
            }
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
            this.Text = "Processing Clashes...";
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