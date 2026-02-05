using System;
using SolidWorks.Interop.sldworks;

namespace SheetSolver
{
    class InsertSheet
    {
        public void generate(ApplicationMgr mgr)
        {
            using (var popup = new LoadingPopup("Populating Title Block..."))
            {
                popup.Show();
                PopulateTitleBlock(mgr);
            }
            //using (var popup = new LoadingPopup("Generating hole table..."))
            //{
            //    popup.Show();
            //    CreateHoleTable(mgr);
            //}

            //PropertyManager pMgr = new PropertyManager();
            //using (var popup = new LoadingPopup("Populating properties..."))
            //{
            //    popup.Show();
            //    PopulateProperties(mgr, pMgr);
            //}
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
    }
}