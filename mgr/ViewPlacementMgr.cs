using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Collections.Generic;
using System;
using SolidWorks.Interop.swconst;

namespace SheetSolver
{
    class ViewPlacementMgr
    {
        private Sheet _currentSheet;

        public ViewPlacementMgr(Sheet currentSheet)
        {
            _currentSheet = currentSheet;
        }

        // ############# METHODS (PUBLIC) ############

        public View PlaceView(ApplicationMgr mgr, double xLoc, double yLoc, View rootView = null, bool overrideBound = false)
        {
            View createdView = null;

            // Added extra debug here to determine what the state of mgr.doc is at runtime

            Console.WriteLine($"Current logged state of mgr.Doc = {mgr.Doc} / {mgr.Doc.GetType()}");

            if(mgr.Doc.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
            {
                throw new InvalidOperationException($"ViewPlacementMgr.PlaceView failed: mgr.Doc is not a SLDDRW pointer. Found {mgr.Doc.GetType()}. Expected {(int)swDocumentTypes_e.swDocDRAWING}");
            }

            if(!overrideBound)
            {
                if(!WithinBoundary(xLoc, yLoc, 0, mgr.drawingX, 0, mgr.drawingY))
                {
                    throw new InvalidOperationException($"ViewPlacementMgr.PlaceView failed: View placement landed outside of drawing bounds. Placement X, Y: {xLoc}, {yLoc} | Bound X, Y: {mgr.drawingX}, {mgr.drawingY}");
                }
            }

            DrawingDoc swDrawing = (DrawingDoc)mgr.Doc;
            
            if (rootView != null)
            {
                // option 1, place view normally. CreateDrawViewFromModelView3
                mgr.PushRef(swDrawing);

                createdView = swDrawing.CreateDrawViewFromModelView3(mgr.Doc.GetTitle(), mgr.viewName, xLoc, yLoc, 0); // mgr.viewName is fragile. Its just the generic stored front on view we generated at runtime
            }
            else
            {
                // option 2, place projected view. CreateUnfoldedViewAt3
                mgr.Doc.Extension.SelectByID2(rootView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

                createdView = swDrawing.CreateUnfoldedViewAt3(xLoc, yLoc, 0, true);
            }

            Console.WriteLine("Tearing down substack... (PlaceView)");
            mgr.ClearSubStack();

            return createdView;
        }

        private bool WithinBoundary(double x, double y, double xMin, double xMax, double yMin, double yMax)
        {
            if (x > xMax || x < xMin)
            {
                return false;
            }
            else if (y > yMax || y < yMin)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
    }
}
