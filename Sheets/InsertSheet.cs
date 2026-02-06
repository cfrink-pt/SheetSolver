using System;
using System.ComponentModel;
using System.Reflection;
using System.Security.Cryptography;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SheetSolver
{
    public static class DirectionPicker
    {
        public static string Show()
        {
            Form form = new Form
            {
                Text = "Select insert mating configuration.",
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                Width = 250,
                Height = 160
            };

            RadioButton rb2 = new RadioButton { Text = "2 Directions", Left = 20, Top = 15, Checked = true, AutoSize = true };
            RadioButton rb1 = new RadioButton { Text = "1 Direction", Left = 20, Top = 40, AutoSize = true };

            Button confirm = new Button { Text = "Confirm", Left = 20, Top = 75, DialogResult = DialogResult.OK };
            Button cancel = new Button { Text = "Cancel", Left = 110, Top = 75, DialogResult = DialogResult.Cancel };

            form.AcceptButton = confirm;
            form.CancelButton = cancel;
            form.Controls.AddRange(new Control[] { rb2, rb1, confirm, cancel });

            if (form.ShowDialog() == DialogResult.OK)
                return rb2.Checked ? "2 Directions" : "1 Direction";

            return null; // cancelled
        }
    }
    class InsertSheet
    {
        public void generate(ApplicationMgr mgr)
        {
            // lets poll the user to see if they have bi-directional inserts or not.
            bool doDoubleInsert = false;

            string drawingDocPath = mgr.drawingDocPath;

            string result = DirectionPicker.Show();
            int e = 0;
            string assyFile = mgr.assyFileDir + "\\" + mgr.assyFileName;

            if (result != null)
            {
                Console.WriteLine($"Selected: {result}");
                switch(result)
                {
                    case "2 Directions":
                        doDoubleInsert = true;

                        Console.WriteLine($"Attempting to open {assyFile}...");
                        ModelDoc2 assemblyDoc = (ModelDoc2)mgr.App.ActivateDoc3(assyFile, true, (int)swRebuildOnActivation_e.swRebuildActiveDoc, ref e);
                        mgr.PushRef(assemblyDoc);

                        Component2 sheetMetalPart = mgr.FetchSheetMetalInAssy();
                        mgr.PushRef(sheetMetalPart);

                        if (sheetMetalPart == null)
                        {
                            throw new InvalidOperationException("Failed to auto-fetch sheet metal part within assembly. Breaking...");
                        }

                        sheetMetalPart.ReferencedConfiguration = mgr.flatConfigurationName;
                        assemblyDoc.ForceRebuild3(true);
                        
                        if (MessageBox.Show("( 1 / 2 ) Orient your model with the insert scribe \"THIS SIDE\" facing you directly. When ready, select 'OK', otherwise, click 'Cancel' and run the macro again.", "Insert View Orientation 1", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                        {
                            throw new UserCancelledException("User cancelled operation: insert view creation 1/2");
                        }

                        assemblyDoc.NameView(mgr.insertView1);

                        if (MessageBox.Show("( 2 / 2 ) Orient your model with the OTHER insert scribe \"THIS SIDE\" facing you directly. When ready, select 'OK', otherwise, click 'Cancel' and run the macro again.", "Insert View Orientation 2", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                        {
                            throw new UserCancelledException("User cancelled operation: insert view creation 2/2");
                        }

                        assemblyDoc.NameView(mgr.insertView2);

                        int activateErr = 0;
                        mgr.App.ActivateDoc3(mgr.drawingDocPath, false, 0, ref activateErr);
                        break;

                    case "1 Direction":
                        Console.WriteLine($"Attempting to open {assyFile}...");
                        ModelDoc2 assyDoc = (ModelDoc2)mgr.App.ActivateDoc3(assyFile, true, (int)swRebuildOnActivation_e.swRebuildActiveDoc, ref e);
                        mgr.PushRef(assyDoc);

                        Component2 smPart = mgr.FetchSheetMetalInAssy();
                        mgr.PushRef(smPart);

                        if (smPart == null)
                        {
                            throw new InvalidOperationException("Failed to auto-fetch sheet metal part within assembly. Breaking...");
                        }

                        smPart.ReferencedConfiguration = mgr.flatConfigurationName;
                        assyDoc.ForceRebuild3(true);
                        
                        if (MessageBox.Show("Orient your model with the insert scribe \"THIS SIDE\" facing you directly. When ready, select 'OK', otherwise, click 'Cancel' and run the macro again.", "Insert View Orientation", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                        {
                            throw new UserCancelledException("User cancelled operation: insert view creation");
                        }

                        assyDoc.NameView(mgr.insertView1);

                        int activateErr1 = 0;
                        mgr.App.ActivateDoc3(mgr.drawingDocPath, false, 0, ref activateErr1);
                        break;
                };
            }
            else
                throw new UserCancelledException("User Cancelled at directionality of inserts.");

            // clean up for pre-ops.
            mgr.ClearSubStack();

            using (var popup = new LoadingPopup("Populating Title Block..."))
            {
                popup.Show();
                PopulateTitleBlock(mgr);
            }
            using (var popup = new LoadingPopup("Placing Inserts View..."))
            {
                popup.Show();
                PlaceInsertViews(mgr, doDoubleInsert);
            }
            using (var popup = new LoadingPopup("Creating Inserts Table..."))
            {
                popup.Show();
                CreateInsertTable(mgr);
            }
        }

        // ====================================================================

        public void PopulateTitleBlock(ApplicationMgr mgr)
        {
            try
            {
                DrawingDoc swDrawing = (DrawingDoc)mgr.App.ActiveDoc;
                mgr.PushRef(swDrawing);

                swDrawing.CreateDrawViewFromModelView3(mgr.Doc.GetTitle(), mgr.viewName, -mgr.drawingX*2, -mgr.drawingY*2, 0);
                
                bool ret = mgr.Doc.ForceRebuild3(false);
            }
            finally
            {
                Console.WriteLine("Tearing down substack... (PopulateTitleBlock)");
                mgr.ClearSubStack();
            }
        }

        public void PlaceInsertViews(ApplicationMgr mgr, bool doDoubleInsert)
        {
            try
            {
                switch (doDoubleInsert)
                {
                    case true:
                        break;

                    case false:
                        break;
                }
            }
            finally
            {
                Console.WriteLine("Tearing down substack... (PlaceInsertViews)");
                mgr.ClearSubStack();
            }
        }

        public void CreateInsertTable(ApplicationMgr mgr)
        {
            try
            {
                DrawingDoc swDrawing = (DrawingDoc)mgr.App.ActiveDoc;
                mgr.PushRef(swDrawing);

                TableAnnotation table = swDrawing.InsertTableAnnotation2(false, 0.01416083, .206895, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopLeft, @"\\storage\CAD\Solidworks\Phase Setting files\Templates\INSERT TABLE.sldtbt", 2, 5);
                mgr.PushRef(table);

                //FlatSheet.EditCell(mgr, "TestVal", "ITEM NO.", 1, 0);
            }
            finally
            {
                Console.WriteLine("Tearing down substack... (CreateInsertTable)");
                mgr.ClearSubStack();
            }
        }
    }
}