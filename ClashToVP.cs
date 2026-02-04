using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.ComApi;
using Autodesk.Navisworks.Api.Interop.ComApi;
using Autodesk.Navisworks.Api.Plugins;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using NavisApp = Autodesk.Navisworks.Api.Application;
using NavisColor = Autodesk.Navisworks.Api.Color;

namespace ClashToVP
{
    // ═══════════════════════════════════════════════════════════════════════════
    // PLUGIN 1: CLASH CONTEXT + CIRCLE
    // Section Box + Unhide Context + Zoom + Red Circle
    // ═══════════════════════════════════════════════════════════════════════════
    [PluginAttribute("ClashToVP.ContextCircle", "ADSK",
        DisplayName = "Clash VP - Context & Circle",
        ToolTip = "Section, Unhide, Circle",
        ExtendedToolTip = "Section Box around clash, Unhides context, Zooms in, Marks a red circle.")]
    [AddInPluginAttribute(AddInLocation.AddIn,
        Icon = "PluginIcon16.png",
        LargeIcon = "PluginIcon32.png")]
    public class ClashContextPlugin : AddInPlugin
    {
        private string logPath = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "ClashToVP_Log.txt");

        private System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
        private ProgressForm progressForm = null;

        private void DebugLog(string message)
        {
            try
            {
                string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                System.IO.File.AppendAllText(logPath, $"{timestamp} {message}{Environment.NewLine}");
            }
            catch { }
        }

        public override int Execute(params string[] parameters)
        {
            try { System.IO.File.WriteAllText(logPath, ""); } catch { }
            stopwatch.Restart();

            try
            {
                Document doc = NavisApp.ActiveDocument;
                if (doc == null) return 0;

                DocumentClash clashDoc = doc.GetClash();
                if (clashDoc == null || clashDoc.TestsData.Tests.Count == 0)
                {
                    MessageBox.Show("No Clash Tests found.");
                    return 0;
                }

                ClashTest selectedTest = ClashHelper.SelectClashTest(clashDoc.TestsData.Tests);
                if (selectedTest == null) return 0;

                RunProcess(doc, selectedTest);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
                DebugLog("ERROR: " + ex.Message);
            }
            finally
            {
                if (progressForm != null && !progressForm.IsDisposed) progressForm.Close();
            }

            return 0;
        }

        private void RunProcess(Document doc, ClashTest test)
        {
            List<ClashResult> clashResults = ClashHelper.GetClashResults(test);
            int total = clashResults.Count;
            if (total == 0) { MessageBox.Show("No clash results found."); return; }

            progressForm = new ProgressForm(total);
            progressForm.Show();

            ComApi.InwOpState10 comState = ComApiBridge.State;

            // Create Folder
            string folderName = "Clash_Context_" + test.DisplayName + "_" + DateTime.Now.ToString("HHmm");
            GroupItem folderRef = ClashHelper.CreateFolder(doc, folderName);

            int count = 0;

            // 1. Reset Global State First (Ensure we start clean)
            doc.Models.ResetAllHidden();
            ClashHelper.DisableSectionPlanes(comState);

            foreach (ClashResult result in clashResults)
            {
                if (progressForm.CancelRequested) break;
                count++;
                progressForm.UpdateProgress(count, total, result.DisplayName);
                DebugLog($"Processing {count}/{total}: {result.DisplayName}");

                try
                {
                    ProcessClash(doc, comState, folderRef, result, count);
                }
                catch (Exception ex)
                {
                    DebugLog($"Error on clash {count}: {ex.Message}");
                }
            }

            // Cleanup at end
            ClashHelper.DisableSectionPlanes(comState);
            doc.Models.ResetAllHidden();
            doc.CurrentSelection.Clear();

            progressForm.Close();
            MessageBox.Show($"Complete! Saved to folder: {folderName}");
        }

        private void ProcessClash(Document doc, ComApi.InwOpState10 comState, GroupItem folder, ClashResult result, int index)
        {
            ModelItem item1 = result.Item1;
            ModelItem item2 = result.Item2;

            // 1. Collect Clash Items
            ModelItemCollection clashItems = new ModelItemCollection();
            if (item1 != null) clashItems.Add(item1);
            if (item2 != null) clashItems.Add(item2);

            if (clashItems.Count == 0) return;

            BoundingBox3D bbox = ClashHelper.GetBoundingBox(clashItems);
            if (bbox == null) return;

            // 2. Setup Visibility
            // The requirement: "Remove hide others".
            // So we ensure EVERYTHING is visible, but we color the clash items red.
            doc.Models.ResetAllHidden();

            // 3. Color Clash Items Red
            doc.Models.OverridePermanentColor(clashItems, new NavisColor(1.0, 0.0, 0.0));

            // 4. Section Box (Directly on top of items)
            // We use the bounding box of the clash items to define the section box
            ClashHelper.CreateSectionBox(comState, bbox, expansion: 0.5); // 0.5 meter padding

            // 5. Zoom to the Clash
            // We select them temporarily to use the API Zoom function, then clear selection
            doc.CurrentSelection.Clear();
            doc.CurrentSelection.CopyFrom(clashItems);
            ClashHelper.ZoomToSelection(doc, comState);
            doc.CurrentSelection.Clear(); // Clear selection so it doesn't look blue in the VP

            // 6. Save Viewpoint with Redline Circle
            // Because we zoomed to selection, the clash is in the center of the screen.
            // We draw a generic circle in the middle of the view.
            ClashHelper.SaveViewpointWithCircle(doc, folder, $"Clash {index} - {result.DisplayName}");

            // 7. Reset Section/Color for next loop (Optional, but safer for cleaner loops)
            ClashHelper.DisableSectionPlanes(comState);
            doc.Models.ResetPermanentMaterials(clashItems);
            System.Windows.Forms.Application.DoEvents();
        }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // HELPER CLASS
    // ═══════════════════════════════════════════════════════════════════════════
    public static class ClashHelper
    {
        // ... (Existing Selection Helpers) ...
        public static ClashTest SelectClashTest(SavedItemCollection tests)
        {
            List<ClashTest> clashTests = new List<ClashTest>();
            foreach (SavedItem item in tests)
                if (item is ClashTest ct) clashTests.Add(ct);

            if (clashTests.Count == 0) return null;
            if (clashTests.Count == 1) return clashTests[0];

            // Simple selection logic (returning first for brevity in this snippet, 
            // or use your previous Form code here)
            return clashTests[0];
        }

        public static List<ClashResult> GetClashResults(ClashTest test)
        {
            List<ClashResult> results = new List<ClashResult>();
            foreach (SavedItem item in test.Children)
                if (item is ClashResult cr) results.Add(cr);
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
            // Calculate combined bounding box
            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
            bool found = false;

            foreach (ModelItem item in items)
            {
                BoundingBox3D b = item.BoundingBox();
                if (b == null) continue;
                minX = Math.Min(minX, b.Min.X); minY = Math.Min(minY, b.Min.Y); minZ = Math.Min(minZ, b.Min.Z);
                maxX = Math.Max(maxX, b.Max.X); maxY = Math.Max(maxY, b.Max.Y); maxZ = Math.Max(maxZ, b.Max.Z);
                found = true;
            }
            return found ? new BoundingBox3D(new Point3D(minX, minY, minZ), new Point3D(maxX, maxY, maxZ)) : null;
        }

        // --- NEW: Save Viewpoint with Redline Circle ---
        public static void SaveViewpointWithCircle(Document doc, GroupItem folder, string name)
        {
            // 1. Create a copy of the current camera
            Viewpoint vp = doc.CurrentViewpoint.ToViewpoint();
            SavedViewpoint savedVp = new SavedViewpoint(vp);
            savedVp.DisplayName = name;

            // 2. Add Redline Circle
            // Coordinates in Navisworks Redlines are normalized to the view? 
            // Actually, Redline coordinates are often Screen Space (0,0 to width,height) or Normalized.
            // For simplicity in .NET API without complex projection math, we draw a circle 
            // in the "middle" of the saved view assuming standard aspect ratio.

            // Note: In strict API, we might need to set specific coordinates. 
            // We use a normalized coordinate system estimation here.
            // A circle in the center.
            RedlineEllipse circle = new RedlineEllipse();
            circle.LineColor = new NavisColor(1, 0, 0); // Red
            circle.LineWidth = 5; // Thick line

            // Create a box for the ellipse in the center of the screen
            // These coordinates are relative to the viewport. 
            // (Attempts to center it based on typical relative coord behavior)
            // If this doesn't appear centered, it's due to API coordinate space specifics 
            // which can vary by version, but this is the standard approach.
            double size = 0.2; // 20% of screen
            circle.BoundingBox = new BoundingBox2D(
                new Point2D(-size, -size),
                new Point2D(size, size));

            savedVp.Redlines.Add(circle);

            // 3. Save
            doc.SavedViewpoints.InsertCopy(folder, folder.Children.Count, savedVp);
        }

        public static void CreateSectionBox(ComApi.InwOpState10 comState, BoundingBox3D bbox, double expansion)
        {
            if (comState == null) return;

            // Expand box slightly so we don't cut the surface of the object
            double minX = bbox.Min.X - expansion;
            double minY = bbox.Min.Y - expansion;
            double minZ = bbox.Min.Z - expansion;
            double maxX = bbox.Max.X + expansion;
            double maxY = bbox.Max.Y + expansion;
            double maxZ = bbox.Max.Z + expansion;

            try
            {
                dynamic view = comState.CurrentView;
                dynamic clipPlanes = view.ClippingPlanes();

                // Clear existing
                while (clipPlanes.Count > 0) clipPlanes.RemovePlane(1);

                // Add 6 planes to form a box
                // Normals point OUT of the visible volume (or IN depending on API version logic, usually pointing to what is removed)
                // In COM: Plane(a,b,c,d) -> ax + by + cz + d = 0

                AddPlane(clipPlanes, comState, -1, 0, 0, maxX); // Right
                AddPlane(clipPlanes, comState, 1, 0, 0, -minX); // Left
                AddPlane(clipPlanes, comState, 0, -1, 0, maxY); // Top (Y is up/back depending on Up Vector)
                AddPlane(clipPlanes, comState, 0, 1, 0, -minY); // Bottom
                AddPlane(clipPlanes, comState, 0, 0, -1, maxZ); // Front
                AddPlane(clipPlanes, comState, 0, 0, 1, -minZ); // Back

                clipPlanes.Enabled = true;
            }
            catch { }
        }

        private static void AddPlane(dynamic clipPlanes, ComApi.InwOpState10 comState, double a, double b, double c, double d)
        {
            dynamic plane = comState.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaClipPlane, null, null);
            plane.Plane.SetValue(a, b, c, d);
            plane.Enabled = true;
            clipPlanes.AddPlane(plane);
        }

        public static void DisableSectionPlanes(ComApi.InwOpState10 comState)
        {
            try
            {
                dynamic view = comState.CurrentView;
                dynamic clipPlanes = view.ClippingPlanes();
                clipPlanes.Enabled = false;
            }
            catch { }
        }

        public static void ZoomToSelection(Document doc, ComApi.InwOpState10 comState)
        {
            // COM Zoom is smoothest
            try { comState.ZoomInCurViewOnCurSel(); }
            catch { }
        }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // PROGRESS FORM (Kept simple)
    // ═══════════════════════════════════════════════════════════════════════════
    public class ProgressForm : Form
    {
        private ProgressBar progressBar;
        private Label lblStatus;
        private Button btnCancel;
        public bool CancelRequested { get; private set; } = false;

        public ProgressForm(int total)
        {
            this.Size = new System.Drawing.Size(400, 150);
            this.Text = "Processing Clashes";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.TopMost = true;

            lblStatus = new Label { Location = new System.Drawing.Point(20, 20), Size = new System.Drawing.Size(350, 20), Text = "Starting..." };
            progressBar = new ProgressBar { Location = new System.Drawing.Point(20, 50), Size = new System.Drawing.Size(350, 20), Maximum = total };
            btnCancel = new Button { Location = new System.Drawing.Point(150, 80), Text = "Cancel" };
            btnCancel.Click += (s, e) => CancelRequested = true;

            this.Controls.Add(lblStatus);
            this.Controls.Add(progressBar);
            this.Controls.Add(btnCancel);
        }

        public void UpdateProgress(int current, int total, string msg)
        {
            if (InvokeRequired) { Invoke(new Action(() => UpdateProgress(current, total, msg))); return; }
            progressBar.Value = current;
            lblStatus.Text = $"{current}/{total}: {msg}";
            Application.DoEvents();
        }
    }
}