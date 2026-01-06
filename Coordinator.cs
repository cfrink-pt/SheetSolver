using System;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

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

            try
            {
                // first, we initialize the drawing.
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

            }
            finally
            {
                Console.WriteLine("Tearing Down Main...");
                mgr.TearDown();
            }

        }

        // the goal by the end of this method is to have the drawing opened and first two views placed.
        private void InitializeDrawingFromPartDoc(ApplicationMgr mgr)
        {
            try
            {
                // first validate we have a part opened.
                mgr.StoreOpenDoc();
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
    }
}