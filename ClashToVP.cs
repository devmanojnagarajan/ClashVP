using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.ComApi;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;

// Resolves 'Application' ambiguity between Navisworks and WinForms
using NavisApp = Autodesk.Navisworks.Api.Application;

namespace ClashToVP
{
    [PluginAttribute("ClashToVP", "devmanojnagarajan",
        DisplayName = "Clash to Viewpoint",
        ToolTip = "Automates Viewpoints from selected Clash Test",
        ExtendedToolTip = "Select a clash test in the Clash Detective window, then run this tool.")]
    [AddInPluginAttribute(AddInLocation.AddIn,
        Icon = "PluginIcon16.png",
        LargeIcon = "PluginIcon32.png")]
    public class MainPlugin : AddInPlugin
    {
        // Log file path on desktop
        private string logPath = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "ClashToVP_Log.txt");

        // Debug log helper
        private void DebugLog(string message)
        {
            string logMessage = DateTime.Now.ToString("HH:mm:ss") + " [ClashToVP] " + message;
            System.Diagnostics.Debug.WriteLine(logMessage);

            try
            {
                System.IO.File.AppendAllText(logPath, logMessage + Environment.NewLine);
            }
            catch { }
        }

        public override int Execute(params string[] parameters)
        {
            // Clear previous log
            try { System.IO.File.WriteAllText(logPath, ""); } catch { }

            DebugLog("=== Plugin Execute Started ===");

            try
            {
                Document doc = NavisApp.ActiveDocument;
                DebugLog("Got active document");

                if (doc == null)
                {
                    MessageBox.Show("No active document found.");
                    DebugLog("ERROR: No active document");
                    return 0;
                }

                // Get the clash document part
                DocumentClash clashDoc = doc.GetClash();
                DebugLog("Got clash document");

                if (clashDoc == null)
                {
                    MessageBox.Show("Could not access Clash Detective.");
                    DebugLog("ERROR: clashDoc is null");
                    return 0;
                }

                // Get all clash tests
                DocumentClashTests testsData = clashDoc.TestsData;
                DebugLog("Got tests data. Test count: " + testsData.Tests.Count);

                if (testsData.Tests.Count == 0)
                {
                    MessageBox.Show("No Clash Tests found in this document.");
                    return 0;
                }

                // Let user select which test to process
                ClashTest selectedTest = SelectClashTest(testsData.Tests);

                if (selectedTest == null)
                {
                    MessageBox.Show("No Clash Test selected.");
                    DebugLog("No test selected");
                    return 0;
                }

                DebugLog("Selected test: " + selectedTest.DisplayName);
                DebugLog("Children count: " + selectedTest.Children.Count);

                RunClashAutomation(doc, selectedTest);
            }
            catch (Exception ex)
            {
                DebugLog("FATAL ERROR in Execute: " + ex.Message);
                DebugLog("Stack trace: " + ex.StackTrace);
                MessageBox.Show("Error: " + ex.Message + "\n\nCheck log file on Desktop.");
            }

            DebugLog("=== Plugin Execute Finished ===");
            return 0;
        }

        private ClashTest SelectClashTest(SavedItemCollection tests)
        {
            DebugLog("SelectClashTest called");

            List<ClashTest> clashTests = new List<ClashTest>();
            foreach (SavedItem item in tests)
            {
                DebugLog("Found item: " + item.DisplayName + " Type: " + item.GetType().Name);
                if (item is ClashTest ct)
                {
                    clashTests.Add(ct);
                }
            }

            DebugLog("Total ClashTests found: " + clashTests.Count);

            if (clashTests.Count == 0)
            {
                return null;
            }

            if (clashTests.Count == 1)
            {
                DebugLog("Only one test, auto-selecting: " + clashTests[0].DisplayName);
                return clashTests[0];
            }

            using (Form form = new Form())
            {
                form.Text = "Select Clash Test";
                form.Width = 350;
                form.Height = 150;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                ComboBox combo = new ComboBox();
                combo.DropDownStyle = ComboBoxStyle.DropDownList;
                combo.Left = 20;
                combo.Top = 20;
                combo.Width = 290;

                foreach (ClashTest ct in clashTests)
                {
                    combo.Items.Add(ct.DisplayName);
                }
                combo.SelectedIndex = 0;

                Button btnOk = new Button();
                btnOk.Text = "OK";
                btnOk.DialogResult = DialogResult.OK;
                btnOk.Left = 150;
                btnOk.Top = 60;
                btnOk.Width = 75;

                Button btnCancel = new Button();
                btnCancel.Text = "Cancel";
                btnCancel.DialogResult = DialogResult.Cancel;
                btnCancel.Left = 235;
                btnCancel.Top = 60;
                btnCancel.Width = 75;

                form.Controls.Add(combo);
                form.Controls.Add(btnOk);
                form.Controls.Add(btnCancel);
                form.AcceptButton = btnOk;
                form.CancelButton = btnCancel;

                if (form.ShowDialog() == DialogResult.OK)
                {
                    DebugLog("User selected: " + clashTests[combo.SelectedIndex].DisplayName);
                    return clashTests[combo.SelectedIndex];
                }
            }

            return null;
        }

        private void RunClashAutomation(Document doc, ClashTest test)
        {
            DebugLog("=== RunClashAutomation Started ===");

            // Count clash results
            int total = 0;
            foreach (SavedItem item in test.Children)
            {
                DebugLog("Child item: " + item.DisplayName + " Type: " + item.GetType().Name);
                if (item is ClashResult)
                {
                    total++;
                }
            }

            DebugLog("Total ClashResults: " + total);

            if (total == 0)
            {
                MessageBox.Show("No clash results found in the selected test.");
                return;
            }

            // Get COM state for zoom functionality
            ComApi.InwOpState10 comState = null;
            try
            {
                comState = ComApiBridge.State;
                DebugLog("Got COM state successfully");
            }
            catch (Exception ex)
            {
                DebugLog("WARNING: Could not get COM state: " + ex.Message);
                DebugLog("Will continue without COM zoom functionality");
            }

            int count = 0;
            int successCount = 0;
            int errorCount = 0;

            try
            {
                // Create folder for viewpoints
                DebugLog("Creating viewpoint folder...");
                string folderName = "Clashes_" + test.DisplayName + "_" + DateTime.Now.ToString("yyyyMMdd_HHmm");
                FolderItem clashFolder = new FolderItem();
                clashFolder.DisplayName = folderName;
                doc.SavedViewpoints.AddCopy(clashFolder);
                DebugLog("Folder created: " + folderName);

                int folderIndex = doc.SavedViewpoints.Value.Count - 1;
                GroupItem folderRef = doc.SavedViewpoints.Value[folderIndex] as GroupItem;

                if (folderRef == null)
                {
                    DebugLog("ERROR: Could not get folder reference");
                    MessageBox.Show("Failed to create viewpoint folder.");
                    return;
                }
                DebugLog("Got folder reference at index: " + folderIndex);

                // Process each clash result
                foreach (SavedItem issue in test.Children)
                {
                    if (issue is ClashResult result)
                    {
                        count++;
                        DebugLog("--- Processing clash " + count + "/" + total + ": " + result.DisplayName + " ---");

                        try
                        {
                            ProcessSingleClash(doc, comState, folderRef, result, count);
                            successCount++;
                            DebugLog("Clash " + count + " processed successfully");
                        }
                        catch (Exception ex)
                        {
                            errorCount++;
                            DebugLog("ERROR processing clash " + count + ": " + ex.Message);
                            DebugLog("Stack trace: " + ex.StackTrace);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLog("FATAL ERROR in RunClashAutomation: " + ex.Message);
                DebugLog("Stack trace: " + ex.StackTrace);
            }

            // Final cleanup
            try
            {
                DebugLog("Final cleanup...");
                doc.Models.ResetAllHidden();
                doc.CurrentSelection.Clear();
                DebugLog("Cleanup complete");
            }
            catch (Exception ex)
            {
                DebugLog("ERROR in final cleanup: " + ex.Message);
            }

            string summary = "Completed!\n\nSuccess: " + successCount + "\nErrors: " + errorCount + "\nTotal: " + total;
            summary += "\n\nLog file saved to Desktop.";
            DebugLog(summary);
            MessageBox.Show(summary);
        }

        private void ProcessSingleClash(Document doc, ComApi.InwOpState10 comState, GroupItem folder, ClashResult result, int index)
        {
            DebugLog("Step 1: Collecting clash items...");

            ModelItemCollection clashItems = new ModelItemCollection();

            if (result.Item1 != null)
            {
                clashItems.Add(result.Item1);
                DebugLog("  Item1: " + result.Item1.DisplayName);
            }
            else
            {
                DebugLog("  Item1 is NULL");
            }

            if (result.Item2 != null)
            {
                clashItems.Add(result.Item2);
                DebugLog("  Item2: " + result.Item2.DisplayName);
            }
            else
            {
                DebugLog("  Item2 is NULL");
            }

            if (clashItems.Count == 0)
            {
                DebugLog("  No items to process, skipping");
                return;
            }

            DebugLog("  Total clash items: " + clashItems.Count);

            Point3D clashCenter = result.Center;
            DebugLog("  Clash center: X=" + clashCenter.X + " Y=" + clashCenter.Y + " Z=" + clashCenter.Z);

            // Step 2: Reset view
            DebugLog("Step 2: Resetting view...");
            try
            {
                doc.Models.ResetAllHidden();
                DebugLog("  Reset done");
            }
            catch (Exception ex)
            {
                DebugLog("  ERROR in reset: " + ex.Message);
            }

            // Step 3: Apply isolation
            DebugLog("Step 3: Applying isolation...");
            try
            {
                ModelItemCollection allRoots = new ModelItemCollection();
                foreach (Model model in doc.Models)
                {
                    allRoots.Add(model.RootItem);
                }

                doc.Models.SetHidden(allRoots, true);
                DebugLog("  Hidden all roots");

                doc.Models.SetHidden(clashItems, false);
                DebugLog("  Unhidden clash items");

                doc.Models.OverridePermanentColor(clashItems, new Color(1.0, 0.0, 0.0));
                DebugLog("  Applied red color");
            }
            catch (Exception ex)
            {
                DebugLog("  ERROR in isolation: " + ex.Message);
                throw;
            }

            // Step 4: Select and zoom
            DebugLog("Step 4: Selecting and zooming...");
            try
            {
                doc.CurrentSelection.Clear();
                doc.CurrentSelection.CopyFrom(clashItems);
                DebugLog("  Selection set");

                // Use COM API to zoom if available
                if (comState != null)
                {
                    comState.ZoomInCurViewOnCurSel();
                    DebugLog("  Zoomed to selection (COM)");
                }
                else
                {
                    // Fallback: manual zoom
                    ZoomToItems(doc, clashItems);
                    DebugLog("  Zoomed to selection (manual)");
                }
            }
            catch (Exception ex)
            {
                DebugLog("  ERROR in zoom: " + ex.Message);
                // Try manual zoom as fallback
                try
                {
                    ZoomToItems(doc, clashItems);
                    DebugLog("  Fallback zoom completed");
                }
                catch { }
            }

            // Step 5: Save viewpoint
            DebugLog("Step 5: Saving viewpoint...");
            try
            {
                Viewpoint currentVp = doc.CurrentViewpoint.ToViewpoint();
                DebugLog("  Got current viewpoint");

                SavedViewpoint savedVp = new SavedViewpoint(currentVp);
                savedVp.DisplayName = "Clash " + index + " - " + result.DisplayName;
                DebugLog("  Created SavedViewpoint: " + savedVp.DisplayName);

                doc.SavedViewpoints.InsertCopy(folder, folder.Children.Count, savedVp);
                DebugLog("  Viewpoint saved");
            }
            catch (Exception ex)
            {
                DebugLog("  ERROR saving viewpoint: " + ex.Message);
                throw;
            }

            // Step 6: Reset for next clash
            DebugLog("Step 6: Resetting for next clash...");
            try
            {
                doc.Models.ResetAllHidden();
                doc.Models.ResetPermanentMaterials(clashItems);
                doc.CurrentSelection.Clear();
                DebugLog("  Reset complete");
            }
            catch (Exception ex)
            {
                DebugLog("  ERROR in reset: " + ex.Message);
            }

            DebugLog("ProcessSingleClash completed for clash " + index);
        }

        private void ZoomToItems(Document doc, ModelItemCollection items)
        {
            DebugLog("  ZoomToItems started");

            try
            {
                BoundingBox3D bbox = GetBoundingBox(items);

                if (bbox == null || bbox.IsEmpty)
                {
                    DebugLog("  Bounding box is null or empty");
                    return;
                }

                Point3D center = bbox.Center;
                double sizeX = bbox.Max.X - bbox.Min.X;
                double sizeY = bbox.Max.Y - bbox.Min.Y;
                double sizeZ = bbox.Max.Z - bbox.Min.Z;
                double maxSize = Math.Max(Math.Max(sizeX, sizeY), sizeZ);

                double distance = maxSize * 2.5;
                if (distance < 1.0) distance = 10.0;

                double angle = Math.PI / 4;
                double camX = center.X + distance * Math.Cos(angle);
                double camY = center.Y + distance * Math.Sin(angle);
                double camZ = center.Z + distance * 0.7;

                Viewpoint vp = CreateLookAtViewpoint(
                    new Point3D(camX, camY, camZ),
                    center);

                if (vp != null)
                {
                    doc.CurrentViewpoint.CopyFrom(vp);
                    DebugLog("  Viewpoint applied");
                }
            }
            catch (Exception ex)
            {
                DebugLog("  ERROR in ZoomToItems: " + ex.Message);
            }
        }

        private BoundingBox3D GetBoundingBox(ModelItemCollection items)
        {
            bool initialized = false;
            double minX = 0, minY = 0, minZ = 0;
            double maxX = 0, maxY = 0, maxZ = 0;

            foreach (ModelItem item in items)
            {
                BoundingBox3D itemBox = item.BoundingBox();
                if (itemBox != null && !itemBox.IsEmpty)
                {
                    if (!initialized)
                    {
                        minX = itemBox.Min.X;
                        minY = itemBox.Min.Y;
                        minZ = itemBox.Min.Z;
                        maxX = itemBox.Max.X;
                        maxY = itemBox.Max.Y;
                        maxZ = itemBox.Max.Z;
                        initialized = true;
                    }
                    else
                    {
                        minX = Math.Min(minX, itemBox.Min.X);
                        minY = Math.Min(minY, itemBox.Min.Y);
                        minZ = Math.Min(minZ, itemBox.Min.Z);
                        maxX = Math.Max(maxX, itemBox.Max.X);
                        maxY = Math.Max(maxY, itemBox.Max.Y);
                        maxZ = Math.Max(maxZ, itemBox.Max.Z);
                    }
                }
            }

            if (!initialized)
                return null;

            return new BoundingBox3D(new Point3D(minX, minY, minZ), new Point3D(maxX, maxY, maxZ));
        }

        private Viewpoint CreateLookAtViewpoint(Point3D eye, Point3D target)
        {
            try
            {
                Viewpoint vp = new Viewpoint();

                double forwardX = target.X - eye.X;
                double forwardY = target.Y - eye.Y;
                double forwardZ = target.Z - eye.Z;

                double forwardLength = Math.Sqrt(forwardX * forwardX + forwardY * forwardY + forwardZ * forwardZ);
                if (forwardLength < 0.0001) return null;

                forwardX /= forwardLength;
                forwardY /= forwardLength;
                forwardZ /= forwardLength;

                double upX = 0, upY = 0, upZ = 1;

                double rightX = forwardY * upZ - forwardZ * upY;
                double rightY = forwardZ * upX - forwardX * upZ;
                double rightZ = forwardX * upY - forwardY * upX;

                double rightLength = Math.Sqrt(rightX * rightX + rightY * rightY + rightZ * rightZ);
                if (rightLength < 0.0001)
                {
                    upX = 0; upY = 1; upZ = 0;
                    rightX = forwardY * upZ - forwardZ * upY;
                    rightY = forwardZ * upX - forwardX * upZ;
                    rightZ = forwardX * upY - forwardY * upX;
                    rightLength = Math.Sqrt(rightX * rightX + rightY * rightY + rightZ * rightZ);
                }

                rightX /= rightLength;
                rightY /= rightLength;
                rightZ /= rightLength;

                double actualUpX = rightY * forwardZ - rightZ * forwardY;
                double actualUpY = rightZ * forwardX - rightX * forwardZ;
                double actualUpZ = rightX * forwardY - rightY * forwardX;

                double actualUpLength = Math.Sqrt(actualUpX * actualUpX + actualUpY * actualUpY + actualUpZ * actualUpZ);
                actualUpX /= actualUpLength;
                actualUpY /= actualUpLength;
                actualUpZ /= actualUpLength;

                UnitVector3D rightUnit = new UnitVector3D(rightX, rightY, rightZ);
                UnitVector3D upUnit = new UnitVector3D(actualUpX, actualUpY, actualUpZ);
                UnitVector3D negForwardUnit = new UnitVector3D(-forwardX, -forwardY, -forwardZ);

                Rotation3D rotation = new Rotation3D(rightUnit, upUnit, negForwardUnit);

                vp.Position = eye;
                vp.Rotation = rotation;
                vp.Projection = ViewpointProjection.Perspective;

                return vp;
            }
            catch (Exception ex)
            {
                DebugLog("  ERROR in CreateLookAtViewpoint: " + ex.Message);
                return null;
            }
        }
    }
}