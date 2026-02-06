using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Data.Common;

namespace SheetSolver
{
    class FlatSheet
    {
        public void generate(ApplicationMgr mgr)
        {
            using (var popup = new LoadingPopup("Generating hole table..."))
            {
                popup.Show();
                CreateHoleTable(mgr);
            }

            PropertyManager pMgr = new PropertyManager();
            using (var popup = new LoadingPopup("Populating properties..."))
            {
                popup.Show();
                PopulateProperties(mgr, pMgr);
            }

            using (var popup = new LoadingPopup("Dimensioning views..."))
            {
                popup.Show();
                DimViewsFlat(mgr);
            }
        }

        // ========================================================================== //
        // HELPER METHODS

        public void CreateHoleTable(ApplicationMgr mgr)
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
                // Getmatdata also fetches and stores whether or not the part is coated, and stores it within the pMgr.
                pMgr.Material = GetMatData(mgr, pMgr);


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
                if (pMgr.CoatStatus)
                {
                    UpdateProperty(swPropMgr, pMgr, "Finish", value: "BONE WHITE S/G TEXTURE");                  
                }
                else
                {
                    UpdateProperty(swPropMgr, pMgr, "Finish", value: "N/A");
                }

                // update surface area cell
                EditCell(mgr, pMgr.SurfaceArea + "IN\u00B2", "SURFACE AREA", 0, 1);
                
                // update coat cell
                string coatString = "NO";
                if (pMgr.CoatStatus == true)
                {
                    coatString = "YES";
                }
                EditCell(mgr, coatString, "FEATURES", 4, 1);

                // edit the cell in the features table if we picked an insert/bend/weld sheet.
                int row = 1;
                foreach (KeyValuePair<string, bool> kvp in mgr.sheetPreferences)
                {
                    if(kvp.Value)
                    {
                        EditCell(mgr, "YES", "FEATURES", row, 1);
                    }
                    row++;
                }

                // rebuild to populate terminal blocks.
                bool ret = swDrawing.ForceRebuild3(false);
            }
            finally
            {
                Console.WriteLine("Tearing down substack... (PopulateProperties)");
                mgr.ClearSubStack();
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

                // now we can delete the view.
                //turn the view into a feature.
                Feature surfAreaViewFeat = (Feature)swDrawing.FeatureByName(surfAreaView.Name);
                mgr.PushRef(surfAreaViewFeat);

                ModelDoc2 dwgModelDoc = (ModelDoc2)mgr.App.ActiveDoc;
                mgr.PushRef(dwgModelDoc);
                
                surfAreaViewFeat.Select2(false, 0);
                dwgModelDoc.DeleteSelection(false);    

                // convert the area from square meters to square inches
                return Math.Round(maxArea*1550.0031, 2);
            }
            finally
            {
                Console.WriteLine("Tearing down substack... (GetSurfaceArea)");
                mgr.ClearSubStack();
            }
        }    
        private string GetMatData(ApplicationMgr mgr, PropertyManager pMgr)
        {
            try
            {
                PartDoc partDoc = (PartDoc)mgr.Doc;
                mgr.PushRef(partDoc);

                string matDb;
                string matName;

                matName = partDoc.GetMaterialPropertyName2("", out matDb);
                
                if (PropertyManager.getRegexValidation(matName, "G90-P"))
                {
                    pMgr.CoatStatus = true;
                }

                return matName;
            }
            finally
            {
                Console.WriteLine("Tearing down substack... (GetMaterial)");
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
        public void UpdateProperty( CustomPropertyManager swPropMgr, PropertyManager pMgr, string propertyName, int? type = null, string value = null, int? resolvedStatus = null)
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
        public static void EditCell(ApplicationMgr mgr, string cellInput, string targetString, int rowId, int colID)
        {
            try
            {
                // grab the table we need to edit.
                DrawingDoc swDrawing = (DrawingDoc)mgr.App.ActiveDoc;
                mgr.PushRef(swDrawing);

                View viewSheet = (View)swDrawing.GetFirstView();
                mgr.PushRef(viewSheet);

                object[] tableAnnotations = (object[])viewSheet.GetTableAnnotations();
                
                foreach (TableAnnotation table in tableAnnotations)
                {
                    try
                    {
                        if (table.Text[0, 0] == targetString)
                        {
                            table.Text[rowId, colID] = cellInput;
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(table);
                    }
                }
            }
            finally
            {
                Console.WriteLine("Tearing down substack... (EditCell)");
                mgr.ClearSubStack();
            }
        }
        private void DimViewsFlat(ApplicationMgr mgr)
        {
           try
            {
                // fetch the view I want to evaluate. I think you can do Sheet.GetAllViews()
                // TODO: Evaluate the feasibility of this. Iterate through each view, dimension
                // the most we can reasonably dimension on the view, and then cull dims? TBD..
                
                DrawingDoc swDrawing = (DrawingDoc)mgr.App.ActiveDoc;
                mgr.PushRef(swDrawing);

                swDrawing.ActivateSheet("FLAT");

                Sheet sheet = (Sheet)swDrawing.GetCurrentSheet();
                mgr.PushRef(sheet);

                Object[] views = (Object[])sheet.GetViews();

                Dictionary<View, DimensionManager> dimViews = new Dictionary<View, DimensionManager>();

                DimensionManager initDmgr = new DimensionManager((View)views[0]);
                initDmgr.ExtractStraightEdgesFromView(mgr, (View)views[0]);

                // this is just storing the offset for the first view and using it universally.
                dimEdge[] bE = new dimEdge[4];
                bE = initDmgr.FindBoundEdges();

                double tW = initDmgr._xMax - initDmgr._xMin;
                double tH = initDmgr._yMax - initDmgr._yMin;
                double offset = Math.Min(tW, tH) / 5.0;

                initDmgr.ReleaseEdgeRefs();

                int viewIndex = 0;
                foreach (View view in views)
                {
                    try
                    {
                        // create a dim mgr for each view, perform dimension ops, release
                        DimensionManager dMgr = new DimensionManager(view);
                        dimViews.Add(view, dMgr);

                        //Console.WriteLine($"\r\n\r\nView Name: {view.GetName2()}\r\n");
                        dMgr.ExtractStraightEdgesFromView(mgr, view);

                        dimEdge[] boundEdges = new dimEdge[4];
                        boundEdges = dMgr.FindBoundEdges();

                        double dimX, dimY;
                        double totalWidth = dMgr._xMax - dMgr._xMin;
                        double totalHeight = dMgr._yMax - dMgr._yMin;

                        // This is where we will dimension views as per index. 
                        switch(viewIndex)
                        {
                            case 0:
                                Console.WriteLine("Dimensioning View 1");
                                
                                dimX = (boundEdges[0].X1 + boundEdges[1].X1) / 2.0;
                                dimY = dMgr._yMax + offset;

                                dMgr.DimensionEdges(
                                    mgr, 
                                    boundEdges[0], 
                                    boundEdges[1], 
                                    dimX, 
                                    dimY
                                    );

                                dimX = dMgr._xMin - offset;
                                dimY = (boundEdges[2].Y1 + boundEdges[3].Y1) / 2.0;

                                dMgr.DimensionEdges(
                                    mgr, 
                                    boundEdges[2], 
                                    boundEdges[3], 
                                    dimX, 
                                    dimY
                                    );

                                break;

                            case 1:
                                Console.WriteLine("Dimensioning View 2");

                                dimX = ((boundEdges[0].X1 + boundEdges[1].X1) / 2.0 ) + .02125;
                                dimY = dMgr._yMax + offset;

                                dMgr.DimensionEdges(
                                    mgr, 
                                    boundEdges[0], 
                                    boundEdges[1], 
                                    dimX, 
                                    dimY,
                                    swTolType_e.swTolSYMMETRIC
                                    );

                                    // TODO: Why is symmetric not applying properly? 

                                break;
                        }
                        
                        dMgr.ReleaseEdgeRefs();
                        viewIndex++;
                    }
                    finally
                    {
                        mgr.PushRef(view);
                    }
                }
            }
            finally
            {
                Console.WriteLine("Tearing down Substack... (DimViewsFlat)");
                mgr.ClearSubStack();
            }
        }
    }
}