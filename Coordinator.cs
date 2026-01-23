using System;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Collections.Generic;

using View = SolidWorks.Interop.sldworks.View;
using FormsView = System.Windows.Forms.View;


namespace SheetSolver
{
    public class LoadingPopup : IDisposable
    {
        private Form _form;
        private Label _label;

        public LoadingPopup(string message = "Processing...", string title = "Please Wait")
        {
            _form = new Form()
            {
                Width = 200,
                Height = 100,
                Text = title,
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MinimizeBox = false,
                MaximizeBox = false,
                ControlBox = false
            };

            _label = new Label()
            {
                Text = message,
                AutoSize = false,
                Dock = DockStyle.Fill
            };

            _form.Controls.Add(_label);
        }

        public void Show()
        {
            _form.Show();
            _form.Refresh();
        }

        public void Close()
        {
            _form.Close();
        }

        public void UpdateMessage(string message)
        {
            _label.Text = message;
            _form.Refresh();
        }

        public void Dispose()
        {
            _form?.Dispose();
        }
    }
    class Coordinator
    {
        public void CreateDrawing()
        {
            ApplicationMgr mgr = new ApplicationMgr();

            // this first setting is simply to keep dimensions from prompting the user to input a value for a dimension on it's creation.
            mgr.App.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, true);

            try
            {
                // validate doc type before operation
                if (!mgr.VerifyDocType(swDocumentTypes_e.swDocPART))
                {
                    throw new InvalidOperationException($"Invalid Document type.\r\nUser opened \"{Enum.GetName(typeof(swDocumentTypes_e), mgr.Doc.GetType())}\", rather than a swDocPART.");
                }

                if (MessageBox.Show("Orient your model with the scribe facing the viewer orthogonally. When ready, select 'OK', otherwise, click 'Cancel' and run the macro again.", "Orientation Check", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                {
                    throw new UserCancelledException("User cancelled operation: Unprepared");
                }

                // we should perform a search of the configurations to validate one exists, then throw an error if not.
                // TODO: MAKE SURE THE FLAT PATTERN FEATURE IS UNSUPPRESSED.
                ValidateFlatPattern(mgr);
                // check that properties are filled in on the part.
                // right now, just validating for scribe value. TODO: EXPAND THIS LOGIC AS NECESSARY. 
                ValidatePropertiesExist(mgr);

                // initialize the drawing. this step ends with two views created, standard first page stuff.
                using (var popup = new LoadingPopup("Initializing drawing..."))
                {
                    popup.Show();
                    InitializeDrawingFromPartDoc(mgr);
                }

                // ========================================================================== //
                // START SHEET 1 OPERATIONS: FLAT. 

                FlatSheet flatSheet = new FlatSheet();
                flatSheet.generate(mgr);

                // ========================================================================== //

                // ========================================================================== //
                // START SHEET 2 OPERATIONS: FLAT. 


                // ========================================================================== //

            }
            finally
            {
                Console.WriteLine("Tearing Down Main...");
                mgr.TearDown();
            }
        }





        // Helper methods involved with initializing drawing process.
        private void ValidatePropertiesExist(ApplicationMgr mgr)
        {
            try
            {
                Dictionary<string, (int Type, string Value, int ResolvedStatus)> partPropMap = new Dictionary<string, (int Type, string Value, int ResolvedStatus)>();

                ModelDocExtension modelDocExtension = mgr.Doc.Extension;
                mgr.PushRef(modelDocExtension);

                CustomPropertyManager swPropMgr = modelDocExtension.get_CustomPropertyManager("");
                mgr.PushRef(swPropMgr);

                object propNamesObj = null;
                object propTypesObj = null;
                object propValuesObj = null;
                object resolvedValsObj = null;

                int result = swPropMgr.GetAll2(
                    ref propNamesObj,
                    ref propTypesObj,
                    ref propValuesObj,
                    ref resolvedValsObj
                );

                string[] propNames;
                int[] propTypes;
                string[] propValues;
                int[] resolvedVals;
                // if properties exist, cast them to arrays
                if (propNamesObj != null)
                {
                    propNames = (string[])propNamesObj;
                    propTypes = (int[])propTypesObj;
                    propValues = (string[])propValuesObj;
                    resolvedVals = (int[])resolvedValsObj;

                    for (int i = 0; i < propNames.Length; i++)
                    {
                        // for each property, lets STORE IT. WHOOP.
                        partPropMap.Add(propNames[i], (propTypes[i], propValues[i], resolvedVals[i]));
                    }
                }

                try
                {
                    string scribePropName = "";
                    string pNumberPropName = "";
                    foreach (var prop in partPropMap)
                    {
                        if(PropertyManager.getRegexValidation(prop.Key, @"(?i)^scribe$"))
                        {
                            scribePropName = prop.Key;
                            continue;
                        }

                        if(PropertyManager.getRegexValidation(prop.Key, @"(?i)^partnumber$"))
                        {
                            pNumberPropName = prop.Key;
                            continue;
                        }

                        if( scribePropName != "" & pNumberPropName != "")
                        {
                            break;
                        }
                    }

                    // Check if we actually found our properties
                    if (string.IsNullOrEmpty(scribePropName))
                    {
                        throw new InvalidOperationException("No valid property found for \"Scribe\". Please ensure your part document has a scribe property (N/A, REQUIRED, REFERENCE, ETC..)");
                    }

                    if (string.IsNullOrEmpty(pNumberPropName))
                    {
                        throw new InvalidOperationException("No valid \"partnumber\" property found. Please ensure a property with such title is created as a non-configuration specific property.");
                    }

                    Console.WriteLine($"Props found:\r\n Scribe Property: {scribePropName}\r\n PN Property: {pNumberPropName}");

                    var scribeProp = partPropMap[scribePropName];

                    if (PropertyManager.getRegexValidation(scribeProp.Value, " / "))
                    {
                        throw new InvalidOperationException("Please ensure \"Scribe\" property is populated in " + mgr.Doc.GetTitle() + " before running the macro.");
                    }
                }
                catch (InvalidOperationException)
                {
                    throw new InvalidOperationException("No valid property found for \"Scribe\". Please ensure your part document has a scribe property (N/A, REQUIRED, REFERENCE, ETC..)");
                }
            }
            finally
            {
                Console.WriteLine("Tearing down Substack... (ValidatePropertiesExist)");
                mgr.ClearSubStack();
            }
        }
        private void ValidateFlatPattern(ApplicationMgr mgr)
        {
            try
            {
                // first, lets fetch the configurations somehow. Check that configs exist first.
                int configurationCount = mgr.Doc.GetConfigurationCount();

                if (configurationCount == 0)
                {
                    throw new InvalidOperationException("Please ensure your part has a valid \"SM-FLAT-PATTERN\" configuration with the flat pattern feature unsuppressed.");
                }
                // so we know there actually are configurations stored, lets fetch the names.
                string[] configNames = (string[])mgr.Doc.GetConfigurationNames();

                // now lets search the array for a configuration matching "SM-FLAT-PATTERN"
                bool configMatchFound = false;
                string flatConfigName = "";
                foreach (string config in configNames)
                {
                    if (config.Contains("SM-FLAT-PATTERN"))
                    {
                        configMatchFound = true;
                        flatConfigName = config;
                    }
                }
                if (configMatchFound == false)
                {
                    throw new InvalidOperationException("No valid flat pattern configuration found. Please create a flat pattern configuration titled \"SM-FLAT-PATTERN\" with the flat pattern feature unsuppressed.");
                }

                mgr.Doc.ShowConfiguration2(flatConfigName);
            }
            finally
            {
                Console.WriteLine("Tearing Down... (ValidateFlatPattern())");
                mgr.ClearSubStack();
            }
        }
        private void InitializeDrawingFromPartDoc(ApplicationMgr mgr)
        {
            try
            {
                // Create the drawingdoc reference. Push it to the substack for cleanup later.
                DrawingDoc swDrawing = mgr.CreateAndMoveToDrawing();
                mgr.PushRef(swDrawing);

                // Create the view from model view
                View mainView = swDrawing.CreateDrawViewFromModelView3(mgr.Doc.GetTitle(), mgr.viewName, mgr.drawingX/2, mgr.drawingY/2, 0);
                mgr.PushRef(mainView);

                // Create our projected view from the main view.
                View mainViewProjected = swDrawing.CreateUnfoldedViewAt3(mgr.drawingX/2 + mgr.drawingX/4, mgr.drawingY/2, 0.0, false );
                mgr.PushRef(mainViewProjected);
            }
            finally
            {
                Console.WriteLine("Tearing Down Substack... (InitializeDrawingFromPartDoc())");
                mgr.ClearSubStack();
            }
        }

    }
}