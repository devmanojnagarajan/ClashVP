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
    // PLUGIN 1: ISOLATE SECTION
    // Isolates clash items, creates section plane, zooms in
    // Folder: "Isolate Section_[test]_[date]"
    // ═══════════════════════════════════════════════════════════════════════════
    [PluginAttribute("ClashToVP.IsolateSection", "ADSK",
        DisplayName = "Clash VP - Isolate Section",
        ToolTip = "Isolate + Section Plane",
        ExtendedToolTip = "Hides other elements, creates section plane, clash items colored RED")]
    [AddInPluginAttribute(AddInLocation.AddIn,
        Icon = "PluginIcon16.png",
        LargeIcon = "PluginIcon32.png")]
    public class ClashIsolateSectionPlugin : AddInPlugin
    {
        private string logPath = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "ClashToVP_IsolateSection_Log.txt");

        private System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
        private ProgressForm progressForm = null;

        private void DebugLog(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            string elapsed = stopwatch.IsRunning ? $" [{stopwatch.ElapsedMilliseconds}ms]" : "";
            string logMessage = timestamp + elapsed + " [IsolateSection] " + message;
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

            DebugLog("=== Clash VP - Isolate Section Started ===");

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

            // Folder: "Isolate Section_[test]_[date]"
            string folderName = "Isolate Section_" + test.DisplayName + "_" + DateTime.Now.ToString("yyyyMMdd_HHmm");
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

            // Cleanup
            ClashHelper.DisableSectionPlanes(comState);
            doc.Models.ResetAllHidden();
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

            // Reset hidden state
            DebugLog("  -> Resetting hidden...");
            doc.Models.ResetAllHidden();
            System.Windows.Forms.Application.DoEvents();

            // Hide root items only (faster than hiding all descendants)
            DebugLog("  -> Hiding root items...");
            ModelItemCollection roots = new ModelItemCollection();
            foreach (Model model in doc.Models)
            {
                roots.Add(model.RootItem);
            }
            doc.Models.SetHidden(roots, true);
            System.Windows.Forms.Application.DoEvents();

            // Show clash items and their ancestors
            DebugLog("  -> Showing clash items...");
            ModelItemCollection itemsToShow = ClashHelper.GetItemsWithAncestorsAndDescendants(clashItems);
            doc.Models.SetHidden(itemsToShow, false);
            System.Windows.Forms.Application.DoEvents();

            // Create section plane
            DebugLog("  -> Creating section plane...");
            ClashHelper.CreateSectionBox(comState, bbox);

            // Color items
            DebugLog("  -> Coloring items...");
            ClashHelper.ColorItems(doc, item1, item2);

            // Select and zoom
            DebugLog("  -> Zooming...");
            doc.CurrentSelection.Clear();
            doc.CurrentSelection.CopyFrom(clashItems);
            ClashHelper.ZoomToSelection(doc, comState, bbox);

            // Save viewpoint
            DebugLog("  -> Saving viewpoint...");
            ClashHelper.SaveViewpoint(doc, folder, "Clash " + index + " - " + result.DisplayName);

            // Reset for next
            DebugLog("  -> Resetting for next...");
            ClashHelper.DisableSectionPlanes(comState);
            doc.Models.ResetAllHidden();
            doc.Models.ResetPermanentMaterials(clashItems);
            doc.CurrentSelection.Clear();
            System.Windows.Forms.Application.DoEvents();
        }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // PLUGIN 2: FADED (Template + Unfade)
    // Uses faded template, unfades clash items
    // Folder: "Faded_[test]_[date]"
    // ═══════════════════════════════════════════════════════════════════════════
    [PluginAttribute("ClashToVP.Faded", "ADSK",
        DisplayName = "Clash VP - Faded",
        ToolTip = "Template + Unfade",
        ExtendedToolTip = "Uses faded template viewpoint, unfades clash items, colored RED")]
    [AddInPluginAttribute(AddInLocation.AddIn,
        Icon = "PluginIcon16.png",
        LargeIcon = "PluginIcon32.png")]
    public class ClashFadedPlugin : AddInPlugin
    {
        private string logPath = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "ClashToVP_Faded_Log.txt");

        private System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
        private ProgressForm progressForm = null;
        private SavedViewpoint templateViewpoint = null;

        private void DebugLog(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            string elapsed = stopwatch.IsRunning ? $" [{stopwatch.ElapsedMilliseconds}ms]" : "";
            string logMessage = timestamp + elapsed + " [Faded] " + message;
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

        private void ApplyTemplate(Document doc)
        {
            if (templateViewpoint != null)
                doc.CurrentViewpoint.CopyFrom(templateViewpoint.Viewpoint);
            System.Windows.Forms.Application.DoEvents();
        }

        public override int Execute(params string[] parameters)
        {
            try { System.IO.File.WriteAllText(logPath, ""); } catch { }
            stopwatch.Restart();

            DebugLog("=== Clash VP - Faded Started ===");

            try
            {
                Document doc = NavisApp.ActiveDocument;
                if (doc == null) { MessageBox.Show("No active document."); return 0; }

                DocumentClash clashDoc = doc.GetClash();
                if (clashDoc == null) { MessageBox.Show("Could not access Clash Detective."); return 0; }

                if (clashDoc.TestsData.Tests.Count == 0) { MessageBox.Show("No Clash Tests found."); return 0; }

                ClashTest selectedTest = ClashHelper.SelectClashTest(clashDoc.TestsData.Tests);
                if (selectedTest == null) { MessageBox.Show("No Clash Test selected."); return 0; }

                // Select template
                templateViewpoint = ClashHelper.SelectTemplateViewpoint(doc);
                if (templateViewpoint == null)
                {
                    MessageBox.Show("No template viewpoint selected.\n\nCreate a saved viewpoint with faded elements first.");
                    return 0;
                }

                DebugLog("Selected test: " + selectedTest.DisplayName);
                DebugLog("Template: " + templateViewpoint.DisplayName);
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

            // Folder: "Faded_[test]_[date]"
            string folderName = "Faded_" + test.DisplayName + "_" + DateTime.Now.ToString("yyyyMMdd_HHmm");
            GroupItem folderRef = ClashHelper.CreateFolder(doc, folderName);

            // Apply template at start
            ApplyTemplate(doc);

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

            // Cleanup
            ClashHelper.DisableSectionPlanes(comState);
            ApplyTemplate(doc);

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

            // Unfade clash items (100% opacity)
            DebugLog("  -> Unfading clash items...");
            doc.Models.OverridePermanentTransparency(clashItems, 0.0);
            System.Windows.Forms.Application.DoEvents();

            // Color items
            DebugLog("  -> Coloring items...");
            ClashHelper.ColorItems(doc, item1, item2);

            // Create section plane
            DebugLog("  -> Creating section plane...");
            ClashHelper.CreateSectionBox(comState, bbox);

            // Select and zoom
            DebugLog("  -> Zooming...");
            doc.CurrentSelection.Clear();
            doc.CurrentSelection.CopyFrom(clashItems);
            ClashHelper.ZoomToSelection(doc, comState, bbox);

            // Save viewpoint
            DebugLog("  -> Saving viewpoint...");
            ClashHelper.SaveViewpoint(doc, folder, "Clash " + index + " - " + result.DisplayName);

            // Reset for next
            DebugLog("  -> Resetting for next...");
            ClashHelper.DisableSectionPlanes(comState);
            doc.Models.ResetPermanentMaterials(clashItems);
            doc.CurrentSelection.Clear();
            ApplyTemplate(doc);
            System.Windows.Forms.Application.DoEvents();
        }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // PLUGIN 3: ISOLATE (No Section Plane)
    // Isolates clash items only, no section plane
    // Folder: "Isolate_[test]_[date]"
    // ═══════════════════════════════════════════════════════════════════════════
    [PluginAttribute("ClashToVP.Isolate", "ADSK",
        DisplayName = "Clash VP - Isolate",
        ToolTip = "Isolate Only",
        ExtendedToolTip = "Hides other elements (no section plane), clash items colored RED")]
    [AddInPluginAttribute(AddInLocation.AddIn,
        Icon = "PluginIcon16.png",
        LargeIcon = "PluginIcon32.png")]
    public class ClashIsolatePlugin : AddInPlugin
    {
        private string logPath = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "ClashToVP_Isolate_Log.txt");

        private System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
        private ProgressForm progressForm = null;

        private void DebugLog(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            string elapsed = stopwatch.IsRunning ? $" [{stopwatch.ElapsedMilliseconds}ms]" : "";
            string logMessage = timestamp + elapsed + " [Isolate] " + message;
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

            DebugLog("=== Clash VP - Isolate Started ===");

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

            // Folder: "Isolate_[test]_[date]"
            string folderName = "Isolate_" + test.DisplayName + "_" + DateTime.Now.ToString("yyyyMMdd_HHmm");
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

            // Cleanup
            doc.Models.ResetAllHidden();
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

            // Reset hidden state
            DebugLog("  -> Resetting hidden...");
            doc.Models.ResetAllHidden();
            System.Windows.Forms.Application.DoEvents();

            // Hide root items only (faster)
            DebugLog("  -> Hiding root items...");
            ModelItemCollection roots = new ModelItemCollection();
            foreach (Model model in doc.Models)
            {
                roots.Add(model.RootItem);
            }
            doc.Models.SetHidden(roots, true);
            System.Windows.Forms.Application.DoEvents();

            // Show clash items and their ancestors
            DebugLog("  -> Showing clash items...");
            ModelItemCollection itemsToShow = ClashHelper.GetItemsWithAncestorsAndDescendants(clashItems);
            doc.Models.SetHidden(itemsToShow, false);
            System.Windows.Forms.Application.DoEvents();

            // NO section plane for this plugin

            // Color items
            DebugLog("  -> Coloring items...");
            ClashHelper.ColorItems(doc, item1, item2);

            // Select and zoom
            DebugLog("  -> Zooming...");
            doc.CurrentSelection.Clear();
            doc.CurrentSelection.CopyFrom(clashItems);
            ClashHelper.ZoomToSelection(doc, comState, bbox);

            // Save viewpoint
            DebugLog("  -> Saving viewpoint...");
            ClashHelper.SaveViewpoint(doc, folder, "Clash " + index + " - " + result.DisplayName);

            // Reset for next
            DebugLog("  -> Resetting for next...");
            doc.Models.ResetAllHidden();
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

        public static SavedViewpoint SelectTemplateViewpoint(Document doc)
        {
            List<SavedViewpoint> viewpoints = new List<SavedViewpoint>();
            CollectViewpoints(doc.SavedViewpoints.Value, viewpoints);

            if (viewpoints.Count == 0) return null;

            using (Form form = new Form())
            {
                form.Text = "Select Faded Template Viewpoint";
                form.Width = 400; form.Height = 180;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false; form.MinimizeBox = false;

                Label lbl = new Label { Text = "Select saved viewpoint with faded elements:", Left = 20, Top = 15, Width = 350 };
                ComboBox combo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Left = 20, Top = 45, Width = 340 };
                foreach (SavedViewpoint vp in viewpoints) combo.Items.Add(vp.DisplayName);
                combo.SelectedIndex = 0;

                Button btnOk = new Button { Text = "OK", DialogResult = DialogResult.OK, Left = 180, Top = 90, Width = 80 };
                Button btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Left = 280, Top = 90, Width = 80 };

                form.Controls.Add(lbl);
                form.Controls.Add(combo);
                form.Controls.Add(btnOk);
                form.Controls.Add(btnCancel);
                form.AcceptButton = btnOk;
                form.CancelButton = btnCancel;

                if (form.ShowDialog() == DialogResult.OK)
                    return viewpoints[combo.SelectedIndex];
            }
            return null;
        }

        private static void CollectViewpoints(SavedItemCollection items, List<SavedViewpoint> result)
        {
            foreach (SavedItem item in items)
            {
                if (item is SavedViewpoint vp) result.Add(vp);
                else if (item is GroupItem g) CollectViewpoints(g.Children, result);
            }
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

        public static ModelItemCollection GetAllItems(Document doc)
        {
            ModelItemCollection allItems = new ModelItemCollection();
            foreach (Model model in doc.Models)
            {
                allItems.Add(model.RootItem);
                foreach (ModelItem desc in model.RootItem.Descendants)
                    allItems.Add(desc);
            }
            return allItems;
        }

        public static ModelItemCollection GetItemsWithAncestorsAndDescendants(ModelItemCollection clashItems)
        {
            ModelItemCollection result = new ModelItemCollection();
            foreach (ModelItem item in clashItems)
            {
                result.Add(item);

                // Add ancestors
                ModelItem parent = item.Parent;
                while (parent != null)
                {
                    result.Add(parent);
                    parent = parent.Parent;
                }

                // Add descendants
                foreach (ModelItem desc in item.Descendants)
                    result.Add(desc);
            }
            return result;
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
            // Color both clashing items RED
            ModelItemCollection clashItems = new ModelItemCollection();
            if (item1 != null) clashItems.Add(item1);
            if (item2 != null) clashItems.Add(item2);

            if (clashItems.Count > 0)
            {
                doc.Models.OverridePermanentColor(clashItems, new NavisColor(1.0, 0.0, 0.0)); // Red
            }
        }

        public static void CreateSectionBox(ComApi.InwOpState10 comState, BoundingBox3D bbox)
        {
            if (comState == null) return;

            double pad = 0.5;
            double minX = bbox.Min.X - pad, minY = bbox.Min.Y - pad, minZ = bbox.Min.Z - pad;
            double maxX = bbox.Max.X + pad, maxY = bbox.Max.Y + pad, maxZ = bbox.Max.Z + pad;

            try
            {
                dynamic view = comState.CurrentView;
                dynamic clipPlanes = view.ClippingPlanes();
                while (clipPlanes.Count > 0) clipPlanes.RemovePlane(1);

                AddPlane(clipPlanes, comState, -1, 0, 0, maxX);
                AddPlane(clipPlanes, comState, 1, 0, 0, -minX);
                AddPlane(clipPlanes, comState, 0, -1, 0, maxY);
                AddPlane(clipPlanes, comState, 0, 1, 0, -minY);
                AddPlane(clipPlanes, comState, 0, 0, -1, maxZ);
                AddPlane(clipPlanes, comState, 0, 0, 1, -minZ);

                clipPlanes.Enabled = true;
            }
            catch { }
        }

        private static void AddPlane(dynamic clipPlanes, ComApi.InwOpState10 comState, double a, double b, double c, double d)
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

        public static void DisableSectionPlanes(ComApi.InwOpState10 comState)
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

        public static void ZoomToSelection(Document doc, ComApi.InwOpState10 comState, BoundingBox3D bbox)
        {
            System.Windows.Forms.Application.DoEvents();

            if (comState != null)
            {
                try { comState.ZoomInCurViewOnCurSel(); return; }
                catch { }
            }

            // Manual zoom fallback
            Point3D center = bbox.Center;
            double size = Math.Max(Math.Max(bbox.Max.X - bbox.Min.X, bbox.Max.Y - bbox.Min.Y), bbox.Max.Z - bbox.Min.Z);
            double dist = Math.Max(size * 2.0, 5.0);
            double angle = Math.PI / 4;

            Viewpoint vp = new Viewpoint();
            vp.Position = new Point3D(center.X + dist * Math.Cos(angle), center.Y + dist * Math.Sin(angle), center.Z + dist * 0.5);
            vp.Projection = ViewpointProjection.Perspective;

            try { doc.CurrentViewpoint.CopyFrom(vp); } catch { }
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
