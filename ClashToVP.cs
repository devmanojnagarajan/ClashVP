using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.ComApi;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;

// Resolves 'Application' ambiguity between Navisworks and WinForms
using NavisApp = Autodesk.Navisworks.Api.Application;
using NavisColor = Autodesk.Navisworks.Api.Color;

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

        // Stopwatch for timing
        private System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();

        // Cancellation flag
        private bool cancelRequested = false;

        // Progress form
        private ProgressForm progressForm = null;

        // Debug log helper with millisecond timing
        private void DebugLog(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            string elapsed = stopwatch.IsRunning ? $" [{stopwatch.ElapsedMilliseconds}ms]" : "";
            string logMessage = timestamp + elapsed + " [ClashToVP] " + message;
            System.Diagnostics.Debug.WriteLine(logMessage);

            try
            {
                System.IO.File.AppendAllText(logPath, logMessage + Environment.NewLine);
            }
            catch { }

            // Allow UI to refresh
            System.Windows.Forms.Application.DoEvents();
        }

        // Update progress form
        private void UpdateProgress(int current, int total, string message)
        {
            if (progressForm != null && !progressForm.IsDisposed)
            {
                progressForm.UpdateProgress(current, total, message);
            }
            System.Windows.Forms.Application.DoEvents();
        }

        // Check if cancel was requested
        private bool IsCancelled()
        {
            System.Windows.Forms.Application.DoEvents();
            return cancelRequested || (progressForm != null && progressForm.CancelRequested);
        }

        public override int Execute(params string[] parameters)
        {
            // Clear previous log and start timer
            try { System.IO.File.WriteAllText(logPath, ""); } catch { }
            stopwatch.Restart();
            cancelRequested = false;

            DebugLog("========================================");
            DebugLog("=== Plugin Execute Started ===");
            DebugLog("========================================");

            try
            {
                DebugLog("STEP: Getting active document...");
                Document doc = NavisApp.ActiveDocument;
                DebugLog("SUCCESS: Got active document");

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
                stopwatch.Stop();
                DebugLog("========================================");
                DebugLog("FATAL ERROR in Execute: " + ex.Message);
                DebugLog("Stack trace: " + ex.StackTrace);
                DebugLog("========================================");
                MessageBox.Show("Error: " + ex.Message + "\n\nCheck log file on Desktop:\nClashToVP_Log.txt",
                    "Clash to Viewpoint Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Clean up progress form
                if (progressForm != null && !progressForm.IsDisposed)
                {
                    progressForm.Close();
                    progressForm.Dispose();
                    progressForm = null;
                }
            }

            stopwatch.Stop();
            DebugLog("=== Plugin Execute Finished ===");
            DebugLog("Total execution time: " + stopwatch.ElapsedMilliseconds + "ms");
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
            DebugLog("========================================");
            DebugLog("=== RunClashAutomation Started ===");
            DebugLog("========================================");

            // Count clash results
            DebugLog("STEP: Counting clash results...");
            List<ClashResult> clashResults = new List<ClashResult>();
            foreach (SavedItem item in test.Children)
            {
                if (item is ClashResult cr)
                {
                    clashResults.Add(cr);
                }
            }

            int total = clashResults.Count;
            DebugLog("RESULT: Total ClashResults = " + total);

            if (total == 0)
            {
                MessageBox.Show("No clash results found in the selected test.");
                return;
            }

            // Show progress form with cancel button
            progressForm = new ProgressForm(total);
            progressForm.Show();
            System.Windows.Forms.Application.DoEvents();

            // Get COM state for zoom functionality
            DebugLog("STEP: Getting COM state for zoom...");
            ComApi.InwOpState10 comState = null;
            try
            {
                comState = ComApiBridge.State;
                DebugLog("SUCCESS: Got COM state");
            }
            catch (Exception ex)
            {
                DebugLog("WARNING: Could not get COM state: " + ex.Message);
            }

            int count = 0;
            int successCount = 0;
            int errorCount = 0;
            int skippedCount = 0;

            try
            {
                // Create folder for viewpoints
                DebugLog("STEP: Creating viewpoint folder...");
                string folderName = "Clashes_" + test.DisplayName + "_" + DateTime.Now.ToString("yyyyMMdd_HHmm");
                FolderItem clashFolder = new FolderItem();
                clashFolder.DisplayName = folderName;
                doc.SavedViewpoints.AddCopy(clashFolder);
                DebugLog("SUCCESS: Folder created - " + folderName);

                int folderIndex = doc.SavedViewpoints.Value.Count - 1;
                GroupItem folderRef = doc.SavedViewpoints.Value[folderIndex] as GroupItem;

                if (folderRef == null)
                {
                    DebugLog("FATAL: Could not get folder reference");
                    MessageBox.Show("Failed to create viewpoint folder.");
                    return;
                }
                DebugLog("SUCCESS: Got folder reference at index " + folderIndex);

                // Process each clash result
                DebugLog("========================================");
                DebugLog("=== Starting Clash Processing Loop ===");
                DebugLog("========================================");

                foreach (ClashResult result in clashResults)
                {
                    // Check for cancellation
                    if (IsCancelled())
                    {
                        DebugLog("### CANCELLATION REQUESTED ###");
                        break;
                    }

                    count++;
                    long startMs = stopwatch.ElapsedMilliseconds;

                    UpdateProgress(count, total, "Processing: " + result.DisplayName);

                    DebugLog("");
                    DebugLog("########################################");
                    DebugLog("### CLASH " + count + " of " + total + " ###");
                    DebugLog("### Name: " + result.DisplayName);
                    DebugLog("########################################");

                    try
                    {
                        bool processed = ProcessSingleClash(doc, comState, folderRef, result, count, total);
                        if (processed)
                        {
                            successCount++;
                            long duration = stopwatch.ElapsedMilliseconds - startMs;
                            DebugLog("### CLASH " + count + " COMPLETE (took " + duration + "ms) ###");
                        }
                        else
                        {
                            skippedCount++;
                            DebugLog("### CLASH " + count + " SKIPPED ###");
                        }
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        DebugLog("### CLASH " + count + " FAILED ###");
                        DebugLog("ERROR: " + ex.Message);
                        DebugLog("Stack: " + ex.StackTrace);
                    }

                    System.Windows.Forms.Application.DoEvents();
                }
            }
            catch (Exception ex)
            {
                DebugLog("FATAL ERROR in RunClashAutomation: " + ex.Message);
                DebugLog("Stack trace: " + ex.StackTrace);
            }

            // Final cleanup
            DebugLog("========================================");
            DebugLog("=== Final Cleanup ===");
            DebugLog("========================================");
            try
            {
                DebugLog("Disabling section planes...");
                DisableSectionPlanes(doc, comState);

                DebugLog("Resetting all hidden...");
                doc.Models.ResetAllHidden();

                DebugLog("Clearing selection...");
                doc.CurrentSelection.Clear();

                DebugLog("Cleanup complete");
            }
            catch (Exception ex)
            {
                DebugLog("ERROR in final cleanup: " + ex.Message);
            }

            // Close progress form
            if (progressForm != null && !progressForm.IsDisposed)
            {
                progressForm.Close();
                progressForm.Dispose();
                progressForm = null;
            }

            stopwatch.Stop();
            long totalTime = stopwatch.ElapsedMilliseconds;
            double avgTime = successCount > 0 ? (double)totalTime / successCount : 0;

            DebugLog("========================================");
            DebugLog("=== PROCESSING COMPLETE ===");
            DebugLog("========================================");
            DebugLog("Success: " + successCount);
            DebugLog("Skipped: " + skippedCount);
            DebugLog("Errors:  " + errorCount);
            DebugLog("Total:   " + total);
            DebugLog("Time:    " + totalTime + "ms (" + (totalTime / 1000.0).ToString("F1") + " seconds)");
            DebugLog("Average: " + avgTime.ToString("F0") + "ms per clash");
            DebugLog("========================================");

            string summary = IsCancelled() ? "Processing Cancelled!\n\n" : "Processing Complete!\n\n";
            summary += "Success: " + successCount + "\n";
            summary += "Skipped: " + skippedCount + "\n";
            summary += "Errors: " + errorCount + "\n";
            summary += "Total: " + total + "\n\n";
            summary += "Time: " + (totalTime / 1000.0).ToString("F1") + " seconds\n\n";
            summary += "Log file: ClashToVP_Log.txt (Desktop)";

            MessageBox.Show(summary, "Clash to Viewpoint", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private bool ProcessSingleClash(Document doc, ComApi.InwOpState10 comState, GroupItem folder, ClashResult result, int index, int total)
        {
            // ═══════════════════════════════════════════════════════════════
            // STEP 1: COLLECT CLASH ITEMS
            // ═══════════════════════════════════════════════════════════════
            DebugLog("  [1/6] COLLECTING CLASH ITEMS...");

            ModelItemCollection clashItems = new ModelItemCollection();

            if (result.Item1 != null)
            {
                clashItems.Add(result.Item1);
                DebugLog("       -> Item1: " + result.Item1.DisplayName);
            }
            else
            {
                DebugLog("       -> Item1: NULL (skipped)");
            }

            if (result.Item2 != null)
            {
                clashItems.Add(result.Item2);
                DebugLog("       -> Item2: " + result.Item2.DisplayName);
            }
            else
            {
                DebugLog("       -> Item2: NULL (skipped)");
            }

            if (clashItems.Count == 0)
            {
                DebugLog("       -> NO ITEMS - SKIPPING THIS CLASH");
                return false;
            }

            DebugLog("       -> Total items: " + clashItems.Count);

            // Get bounding box for section plane
            BoundingBox3D clashBBox = GetBoundingBox(clashItems);
            if (clashBBox == null || clashBBox.IsEmpty)
            {
                DebugLog("       -> Could not get bounding box - SKIPPING");
                return false;
            }

            Point3D clashCenter = clashBBox.Center;
            DebugLog("       -> Midpoint: (" + clashCenter.X.ToString("F2") + ", " + clashCenter.Y.ToString("F2") + ", " + clashCenter.Z.ToString("F2") + ")");

            if (IsCancelled()) return false;

            // ═══════════════════════════════════════════════════════════════
            // STEP 2: CREATE SECTION PLANE AT MIDPOINT
            // ═══════════════════════════════════════════════════════════════
            DebugLog("  [2/6] CREATING SECTION PLANE AT MIDPOINT...");
            try
            {
                CreateSectionBox(doc, clashBBox, comState);
                DebugLog("       -> Section plane created at midpoint of clash items");
            }
            catch (Exception ex)
            {
                DebugLog("       -> ERROR creating section plane: " + ex.Message);
            }

            if (IsCancelled()) return false;

            // ═══════════════════════════════════════════════════════════════
            // STEP 3: COLOR CLASH ITEMS
            // ═══════════════════════════════════════════════════════════════
            DebugLog("  [3/6] COLORING CLASH ITEMS...");
            try
            {
                DebugLog("       -> Applying red color to clash items...");
                doc.Models.OverridePermanentColor(clashItems, new NavisColor(1.0, 0.0, 0.0));
                DebugLog("       -> Color applied");
            }
            catch (Exception ex)
            {
                DebugLog("       -> ERROR: " + ex.Message);
            }

            if (IsCancelled()) return false;

            // ═══════════════════════════════════════════════════════════════
            // STEP 4: ISOLATE AND ZOOM
            // ═══════════════════════════════════════════════════════════════
            DebugLog("  [4/6] ISOLATING AND ZOOMING...");
            try
            {
                // Hide all root items first
                DebugLog("       -> Hiding all model items...");
                ModelItemCollection allRoots = new ModelItemCollection();
                foreach (Model model in doc.Models)
                {
                    allRoots.Add(model.RootItem);
                }
                doc.Models.SetHidden(allRoots, true);

                // Show only clash items
                DebugLog("       -> Showing only clash items...");
                doc.Models.SetHidden(clashItems, false);

                // Select clash items
                DebugLog("       -> Selecting clash items...");
                doc.CurrentSelection.Clear();
                doc.CurrentSelection.CopyFrom(clashItems);

                // Zoom to selection
                if (comState != null)
                {
                    DebugLog("       -> Zooming to selection...");
                    comState.ZoomInCurViewOnCurSel();
                    DebugLog("       -> Zoom complete");
                }
                else
                {
                    ZoomToBox(doc, clashBBox);
                    DebugLog("       -> Manual zoom complete");
                }
            }
            catch (Exception ex)
            {
                DebugLog("       -> ERROR: " + ex.Message);
                try { ZoomToBox(doc, clashBBox); } catch { }
            }

            System.Windows.Forms.Application.DoEvents();

            if (IsCancelled()) return false;

            // ═══════════════════════════════════════════════════════════════
            // STEP 5: SAVE VIEWPOINT
            // ═══════════════════════════════════════════════════════════════
            DebugLog("  [5/6] SAVING VIEWPOINT...");
            try
            {
                DebugLog("       -> Capturing current view...");
                Viewpoint currentVp = doc.CurrentViewpoint.ToViewpoint();

                string vpName = "Clash " + index + " - " + result.DisplayName;
                DebugLog("       -> Saving as: " + vpName);
                SavedViewpoint savedVp = new SavedViewpoint(currentVp);
                savedVp.DisplayName = vpName;

                doc.SavedViewpoints.InsertCopy(folder, folder.Children.Count, savedVp);
                DebugLog("       -> Viewpoint saved");
            }
            catch (Exception ex)
            {
                DebugLog("       -> ERROR: " + ex.Message);
                throw;
            }

            if (IsCancelled()) return false;

            // ═══════════════════════════════════════════════════════════════
            // STEP 6: RESET AND MOVE TO NEXT
            // ═══════════════════════════════════════════════════════════════
            DebugLog("  [6/6] RESETTING AND MOVING TO NEXT...");
            try
            {
                DebugLog("       -> Disabling section planes...");
                DisableSectionPlanes(doc, comState);

                DebugLog("       -> Unhiding all items...");
                doc.Models.ResetAllHidden();

                DebugLog("       -> Removing color override...");
                doc.Models.ResetPermanentMaterials(clashItems);

                DebugLog("       -> Clearing selection...");
                doc.CurrentSelection.Clear();

                DebugLog("       -> Ready for next clash");
            }
            catch (Exception ex)
            {
                DebugLog("       -> ERROR: " + ex.Message);
            }

            DebugLog("  [DONE] Clash " + index + "/" + total + " processed");
            return true;
        }

        // Create a section box around the clash using COM API
        private void CreateSectionBox(Document doc, BoundingBox3D bbox, ComApi.InwOpState10 comState)
        {
            // Add padding around the clash (30% on each side)
            double padX = (bbox.Max.X - bbox.Min.X) * 0.3;
            double padY = (bbox.Max.Y - bbox.Min.Y) * 0.3;
            double padZ = (bbox.Max.Z - bbox.Min.Z) * 0.3;

            // Minimum padding of 0.5 units
            padX = Math.Max(padX, 0.5);
            padY = Math.Max(padY, 0.5);
            padZ = Math.Max(padZ, 0.5);

            double minX = bbox.Min.X - padX;
            double minY = bbox.Min.Y - padY;
            double minZ = bbox.Min.Z - padZ;
            double maxX = bbox.Max.X + padX;
            double maxY = bbox.Max.Y + padY;
            double maxZ = bbox.Max.Z + padZ;

            DebugLog("       -> Section box bounds:");
            DebugLog("          Min: (" + minX.ToString("F2") + ", " + minY.ToString("F2") + ", " + minZ.ToString("F2") + ")");
            DebugLog("          Max: (" + maxX.ToString("F2") + ", " + maxY.ToString("F2") + ", " + maxZ.ToString("F2") + ")");

            if (comState == null)
            {
                DebugLog("       -> COM state not available, skipping section planes");
                return;
            }

            try
            {
                // Get the current view's sectioning planes via COM
                dynamic view = comState.CurrentView;
                dynamic sectionPlanes = view.ClippingPlanes();

                // Clear existing section planes
                DebugLog("       -> Clearing existing section planes...");
                while (sectionPlanes.Count > 0)
                {
                    sectionPlanes.RemovePlane(1);
                }

                // Create 6 section planes to form a box
                // Plane equation: Ax + By + Cz + D = 0
                // Normal vector (A,B,C) points toward the visible (kept) side

                DebugLog("       -> Creating 6 section planes...");

                // Right plane (+X): normal (-1,0,0), keeps x < maxX
                AddSectionPlane(sectionPlanes, comState, -1, 0, 0, maxX);

                // Left plane (-X): normal (1,0,0), keeps x > minX
                AddSectionPlane(sectionPlanes, comState, 1, 0, 0, -minX);

                // Front plane (+Y): normal (0,-1,0), keeps y < maxY
                AddSectionPlane(sectionPlanes, comState, 0, -1, 0, maxY);

                // Back plane (-Y): normal (0,1,0), keeps y > minY
                AddSectionPlane(sectionPlanes, comState, 0, 1, 0, -minY);

                // Top plane (+Z): normal (0,0,-1), keeps z < maxZ
                AddSectionPlane(sectionPlanes, comState, 0, 0, -1, maxZ);

                // Bottom plane (-Z): normal (0,0,1), keeps z > minZ
                AddSectionPlane(sectionPlanes, comState, 0, 0, 1, -minZ);

                // Enable sectioning
                sectionPlanes.Enabled = true;

                DebugLog("       -> Section box created with " + sectionPlanes.Count + " planes");
            }
            catch (Exception ex)
            {
                DebugLog("       -> Section plane COM error: " + ex.Message);
                DebugLog("       -> Continuing without section planes...");
            }
        }

        private void AddSectionPlane(dynamic sectionPlanes, ComApi.InwOpState10 comState, double a, double b, double c, double d)
        {
            try
            {
                // Create a new plane using the COM object factory
                dynamic plane = comState.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaClipPlane, null, null);

                // Set plane equation coefficients
                plane.Plane.SetValue(a, b, c, d);
                plane.Enabled = true;

                // Add to collection
                sectionPlanes.AddPlane(plane);
            }
            catch (Exception ex)
            {
                DebugLog("       -> Error adding plane: " + ex.Message);
            }
        }

        private void DisableSectionPlanes(Document doc, ComApi.InwOpState10 comState)
        {
            if (comState == null) return;

            try
            {
                dynamic view = comState.CurrentView;
                dynamic sectionPlanes = view.ClippingPlanes();

                // Disable sectioning
                sectionPlanes.Enabled = false;

                // Remove all planes
                while (sectionPlanes.Count > 0)
                {
                    sectionPlanes.RemovePlane(1);
                }

                DebugLog("       -> Section planes disabled");
            }
            catch (Exception ex)
            {
                DebugLog("       -> Error disabling section planes: " + ex.Message);
            }
        }

        private void ZoomToBox(Document doc, BoundingBox3D bbox)
        {
            DebugLog("       -> ZoomToBox: Calculating view...");

            try
            {
                Point3D center = bbox.Center;
                double sizeX = bbox.Max.X - bbox.Min.X;
                double sizeY = bbox.Max.Y - bbox.Min.Y;
                double sizeZ = bbox.Max.Z - bbox.Min.Z;
                double maxSize = Math.Max(Math.Max(sizeX, sizeY), sizeZ);

                double distance = maxSize * 2.0;
                if (distance < 1.0) distance = 5.0;

                // Position camera at 45-degree angle
                double angle = Math.PI / 4;
                double camX = center.X + distance * Math.Cos(angle);
                double camY = center.Y + distance * Math.Sin(angle);
                double camZ = center.Z + distance * 0.5;

                Viewpoint vp = CreateLookAtViewpoint(
                    new Point3D(camX, camY, camZ),
                    center);

                if (vp != null)
                {
                    doc.CurrentViewpoint.CopyFrom(vp);
                    DebugLog("       -> Viewpoint applied");
                }
            }
            catch (Exception ex)
            {
                DebugLog("       -> ZoomToBox ERROR: " + ex.Message);
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

    // ═══════════════════════════════════════════════════════════════════════
    // PROGRESS FORM WITH CANCEL BUTTON
    // ═══════════════════════════════════════════════════════════════════════
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
            this.Width = 450;
            this.Height = 180;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ControlBox = false;
            this.TopMost = true;

            lblStatus = new Label();
            lblStatus.Text = "Initializing...";
            lblStatus.Left = 20;
            lblStatus.Top = 20;
            lblStatus.Width = 400;
            lblStatus.AutoEllipsis = true;

            lblProgress = new Label();
            lblProgress.Text = "0 / " + totalItems;
            lblProgress.Left = 20;
            lblProgress.Top = 45;
            lblProgress.Width = 400;

            progressBar = new ProgressBar();
            progressBar.Left = 20;
            progressBar.Top = 70;
            progressBar.Width = 395;
            progressBar.Height = 25;
            progressBar.Minimum = 0;
            progressBar.Maximum = totalItems;
            progressBar.Value = 0;

            btnCancel = new Button();
            btnCancel.Text = "Cancel";
            btnCancel.Left = 170;
            btnCancel.Top = 105;
            btnCancel.Width = 100;
            btnCancel.Height = 30;
            btnCancel.Click += BtnCancel_Click;

            this.Controls.Add(lblStatus);
            this.Controls.Add(lblProgress);
            this.Controls.Add(progressBar);
            this.Controls.Add(btnCancel);
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            CancelRequested = true;
            btnCancel.Enabled = false;
            btnCancel.Text = "Cancelling...";
            lblStatus.Text = "Cancellation requested, finishing current clash...";
        }

        public void UpdateProgress(int current, int total, string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => UpdateProgress(current, total, message)));
                return;
            }

            progressBar.Value = Math.Min(current, progressBar.Maximum);
            lblProgress.Text = current + " / " + total;
            lblStatus.Text = message;
            this.Refresh();
        }
    }
}
