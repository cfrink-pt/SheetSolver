using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Collections.Generic;

using View = SolidWorks.Interop.sldworks.View;
using FormsView = System.Windows.Forms.View;
using System.Linq;
using System.Xml.Schema;
using System.Windows.Forms.VisualStyles;

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

            try
            {
                //first, we should validate the part has a valid flat pattern. I dont want to handle creation here.
                // we should perform a search of the configurations to validate one exists, then throw an error if not.
                ValidateFlatPattern(mgr);
                
                // initialize the drawing. this step ends with two views created, standard first page stuff.
                using (var popup = new LoadingPopup("Initializing drawing..."))
                {
                    popup.Show();
                    InitializeDrawingFromPartDoc(mgr);
                }

                // next, we create the hole table
                using (var popup = new LoadingPopup("Generating hole table..."))
                {
                    popup.Show();
                    CreateHoleTable(mgr);
                }

                // next, we fetch properties.
                PropertyManager pMgr = new PropertyManager();
                using (var popup = new LoadingPopup("Populating properties..."))
                {
                    popup.Show();
                    PopulateProperties(mgr, pMgr);
                }
            }
            finally
            {
                Console.WriteLine("Tearing Down Main...");
                mgr.TearDown();
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
                if (!mgr.VerifyDocType(swDocumentTypes_e.swDocPART))
                {
                    throw new InvalidOperationException($"Invalid Document type.\r\nUser opened \"{Enum.GetName(typeof(swDocumentTypes_e), mgr.Doc.GetType())}\", rather than a swDocPART.");
                }

                if (MessageBox.Show("Orient your model with the scribe facing the viewer orthogonally. When ready, select 'OK', otherwise, click 'Cancel' and run the macro again.", "Orientation Check", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                {
                    throw new UserCancelledException("User cancelled operation: Unprepared");
                }


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

        private void CreateHoleTable(ApplicationMgr mgr)
        {
            try
            {
                DrawingDoc swDrawing = (DrawingDoc)mgr.App.ActiveDoc;
                mgr.PushRef(swDrawing);

                // get our selection manager
                SelectionMgr swSelMgr = (SelectionMgr)mgr.Doc.SelectionManager;
                mgr.PushRef(swSelMgr);

                // store our main view reference. Along with the sheet. forgot that it is the first view technically. 
                View sheet = (View)swDrawing.GetFirstView();
                mgr.PushRef(sheet);

                View mainView = (View)sheet.GetNextView();
                mgr.PushRef(mainView);

                if (mainView != null)
                {
                    // Clear current selection buffer
                    mgr.Doc.ClearSelection2(true);

                    Feature viewFeature = (Feature)swDrawing.FeatureByName(mainView.Name);
                    mgr.PushRef(viewFeature);

                    // select the feature directly
                    bool selected = viewFeature.Select2(false, 0);
                }
                else
                {
                    throw new InvalidOperationException("mainView incorrectly fetched. See CreateHoleTable() within Coordinator and ensure no null references.");
                }

                // now we have the view selected. Lets create the hole table
                // hole table itself handles cleaning up references.
                // TODO: Overhaul error handling to flag non-planar views to throw exceptions where appropriate.
                HoleTableEntrance htEntrance = new HoleTableEntrance();
                htEntrance.DoHoleTable(mgr);                
            }
            finally
            {
                Console.WriteLine("Tearing Down Substack... (CreateHoleTable())");
                mgr.ClearSubStack();
            }
        }
    
        private void PopulateProperties(ApplicationMgr mgr, PropertyManager pMgr)
        {
            try
            {   
                StorePropertyMap(mgr, pMgr);

                // now, begin fetching information. we clear substack in getsurface area so we will need to
                // initialize a new prop manager afterwards.
                pMgr.UserInitials = pMgr.GetUserInitials();
                pMgr.SurfaceArea = GetSurfaceArea(mgr);


                //get a new solidworks property manager
                ModelDoc2 swDrawing = (ModelDoc2)mgr.App.ActiveDoc;
                mgr.PushRef(swDrawing);

                ModelDocExtension modelDocExtension = swDrawing.Extension;
                mgr.PushRef(modelDocExtension);

                CustomPropertyManager swPropMgr = modelDocExtension.get_CustomPropertyManager("");
                mgr.PushRef(swPropMgr);
                

                // now, update properties in solidworks and pMgr propmap
                UpdateProperty(swPropMgr, pMgr, "Drawn By", value: pMgr.UserInitials);
                UpdateProperty(swPropMgr, pMgr, "Drawing title", value: pMgr.FormatFileNameForDrawingTitle(mgr.Doc.GetTitle()));

                // rebuild to populate terminal blocks.
                bool ret = swDrawing.ForceRebuild3(false);
                Console.WriteLine("Successfully rebuilt? " + ret);
            }
            finally
            {
                Console.WriteLine("Tearing down substack... (PopulateProperties)");
                mgr.ClearSubStack();
            }
        }

        public void UpdateProperty(
            CustomPropertyManager swPropMgr,
            PropertyManager pMgr,
            string propertyName,
            int? type = null,
            string value = null,
            int? resolvedStatus = null)
        {
            if (pMgr.propMap.ContainsKey(propertyName))
            {
                var current = pMgr.propMap[propertyName];

                // Update only the values that were provided (not null)
                int newType = type ?? current.Type;
                string newValue = value ?? current.Value;
                int newResolvedStatus = resolvedStatus ?? current.ResolvedStatus;

                // Update the dictionary
                pMgr.propMap[propertyName] = (newType, newValue, newResolvedStatus);

                // Push to SolidWorks (only if value changed)
                if (value != null)
                {
                    swPropMgr.Set(propertyName, newValue);
                }
            }
            else
            {
                throw new InvalidOperationException($"UpdateProperty() tried to update {propertyName}, but it did not exist within the property list.");
            }
        }

        private double GetSurfaceArea(ApplicationMgr mgr)
        {
            try
            {
                DrawingDoc swDrawing = (DrawingDoc)mgr.App.ActiveDoc;
                mgr.PushRef(swDrawing);

                View sheet = (View)swDrawing.GetFirstView();
                mgr.PushRef(sheet);

                View mainView = (View)sheet.GetNextView();
                mgr.PushRef(mainView);

                View sideView = (View)mainView.GetNextView();
                mgr.PushRef(sideView);

                // Clear current selection buffer
                mgr.Doc.ClearSelection2(true);

                Feature viewFeature = (Feature)swDrawing.FeatureByName(sideView.Name);
                mgr.PushRef(viewFeature);

                // select the feature directly
                bool selected = viewFeature.Select2(false, 0);

                View surfAreaView = swDrawing.CreateUnfoldedViewAt3(mgr.drawingX*1.5, mgr.drawingY/2, 0, false);
                mgr.PushRef(surfAreaView);


                // get a list of faces to evaluate
                List<Face2> swFaceList = new List<Face2>(); 

                object[] entityList = (object[])surfAreaView.GetVisibleEntities2(null, (int)swViewEntityType_e.swViewEntityType_Face);

                foreach (object obj in entityList)
                {
                    Entity swEnt = (Entity)obj;

                    int entType = (int)swEnt.GetType();
                    if (entType == (int)swSelectType_e.swSelFACES)
                    {
                        swFaceList.Add((Face2)swEnt);
                    }
                    else
                    {
                        Marshal.ReleaseComObject(swEnt);
                    }
                }

                double maxArea = 0;
                foreach (Face2 face in swFaceList)
                {
                    double area = face.GetArea();
                    if (area > maxArea) maxArea = area;
                    Marshal.ReleaseComObject(face);
                }
                swFaceList.Clear();

                // convert the area from square meters to square inches
                return maxArea*1550.0031;
            }
            finally
            {
                Console.WriteLine("Tearing down substack... (GetSurfaceArea)");
                mgr.ClearSubStack();
            }
        }    

        private void StorePropertyMap(ApplicationMgr mgr, PropertyManager pMgr)
        {
            try
            {
                ModelDoc2 swDrawing = (ModelDoc2)mgr.App.ActiveDoc;
                mgr.PushRef(swDrawing);

                ModelDocExtension modelDocExtension = swDrawing.Extension;
                mgr.PushRef(modelDocExtension);

                CustomPropertyManager swPropMgr = modelDocExtension.get_CustomPropertyManager("");
                mgr.PushRef(swPropMgr);

                pMgr.StorePropertyMap(swPropMgr);
            }
            finally
            {
                Console.WriteLine("Tearing Down Substack... (StorePropertyMap())");
                mgr.ClearSubStack();
            }
        }
    }
}