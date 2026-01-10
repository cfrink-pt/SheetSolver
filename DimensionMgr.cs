using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace SheetSolver
{
    // ok so my idea is to store edge locations.
    // 

    // Rules for fetching edges to be dimensioned.
    //   The two edges being evaluated MUST be parallel.
    //    We must store the direction of the two edges being evaluated.
    //    This will give us an indicator of which axis to evaluate across for the next rule:
    //   The two edges must NOT be colinear. 

    // To add a dimension, we will need to get access to the ModelDoc extension.
    // it has a member called AddDimension, which takes: X, y and z as doubles, along with direction, which is within swSmartdimensionDirection_e.
    // returns a DisplayDimension.

    // we need to select the edges first using ModelDoc's SelectByID2.
    // once our edges are selected, we call .AddDimension(loc x, loc y, loc z, enum direction)
    // then we run ClearSelection2()

    

    // decided to go with a record struct so we could store unique xy pairs for 2 points, one for the start and one for the end point in the edge.
    // plus we get fast lookup which is lovely! we will iterate a lot. 
    public record struct dimEdge(
        double X1,
        double Y1,
        double Dx,
        double Dy,
        double Dz
    );

    class DimensionManager
    {
        public List<dimEdge> _edges;
        public HashSet<dimEdge> _evaluatedEdges;
        public View _view;
        public double _xMin { get; set; }
        public double _xMax { get; set; }
        public double _yMin { get; set; }
        public double _yMax { get; set; }

        public DimensionManager(View view)
        {
            _view = view;
            _edges = new List<dimEdge>();
            _evaluatedEdges = new HashSet<dimEdge>();
        }

        // extracting straight edges from view also will write to this object's internal "Space" representation
        public void ExtractStraightEdgesFromView(ApplicationMgr mgr, View view)
        {
            try
            {
                List<dimEdge> edges = new List<dimEdge>();

                int count = 0;
                object[] entityList = (object[])view.GetVisibleEntities2(null, (int)swViewEntityType_e.swViewEntityType_Edge);
                foreach (object edge in entityList)
                {
                    Entity edgeEnt = (Entity)edge;
                    int entType = (int)edgeEnt.GetType();
                    if (entType == (int)swSelectType_e.swSelEDGES)
                    {
                        // dont forget to cast from ent to edge
                        Edge swEdge = (Edge)edge;
                        mgr.PushRef(swEdge);

                        Curve edgeCurve = (Curve)swEdge.GetCurve();
                        mgr.PushRef(edgeCurve);

                        if (edgeCurve.Identity() == (int)swCurveTypes_e.LINE_TYPE)
                        {
                            count++;

                            // straight line. lets get a param.
                            double[] lineParams = new double[6];
                            lineParams = (double[])edgeCurve.LineParams;

                            double[] transformedLocation = new double[3];
                            double[] transformedVector = new double[3];
                            transformedLocation = TransformModelToSheetSpace(mgr, this._view, lineParams);
                            transformedVector = TransformDirectionModelToSheet(mgr, this._view, lineParams);

                            lineParams[0] = transformedLocation[0];
                            lineParams[1] = transformedLocation[1];
                            lineParams[2] = transformedLocation[2];

                            lineParams[3] = transformedVector[0];
                            lineParams[4] = transformedVector[1];
                            lineParams[5] = transformedVector[2];

                            // Clean up POSITION coordinates
                            for (int i = 0; i < 3; i++)
                            {
                                lineParams[i] = Math.Round(lineParams[i], 8);
                                if (Math.Abs(lineParams[i]) < 1e-10)
                                {
                                    lineParams[i] = 0;
                                }
                            }
                            
                            // Clean up DIRECTION components
                            for (int i = 3; i < 6; i++)
                            {
                                lineParams[i] = Math.Round(lineParams[i], 8);
                                if (Math.Abs(lineParams[i]) < 1e-15) 
                                {
                                    lineParams[i] = 0;
                                }
                            }

                            dimEdge dEdge = new dimEdge
                            {
                                X1 = lineParams[0],
                                Y1 = lineParams[1],
                                Dx = lineParams[3],
                                Dy = lineParams[4],
                                Dz = lineParams[5]
                            };
                            edges.Add(dEdge);
                            Console.WriteLine($"Edge {count}\r\n  X: {dEdge.X1} | Y: {dEdge.Y1}\r\n  Dx: {dEdge.Dx} | Dy: {dEdge.Dy} | Dz: {dEdge.Dz}");
                        }
                    }
                    else
                    {
                        Marshal.ReleaseComObject(edgeEnt);
                    }
                }

                if (edges.Count != 0)
                {
                    this._edges = edges;
                    StoreBounds(edges);
                }
                else
                {
                    throw new InvalidOperationException("ExtractStraightEdgesFromView failed to fetch valid edge entities from view reference. No valid Edges!");
                }
            }
            finally
            {
                Console.WriteLine("Tearing down substack... (ExtractStraightEdgesFromView)");
                mgr.ClearSubStack();
            }
        }

        public dimEdge[] FindBoundEdges()
        {
            // x edges = [0, 1]
            // y edges = [2, 3]
            dimEdge[] boundEdges = new dimEdge[4];
            
            bool xMinFound = false;
            bool xMaxFound = false;
            bool yMinFound = false;
            bool yMaxFound = false;
            foreach (dimEdge edge in this._edges)
            {
                if (edge.X1 == this._xMin && !this._evaluatedEdges.Contains(edge) && xMinFound == false)
                {
                    xMinFound = true;
                    this._evaluatedEdges.Add(edge);
                    boundEdges[0] = edge;
                    continue;
                }

                if (edge.X1 == this._xMax && !this._evaluatedEdges.Contains(edge) && xMaxFound == false)
                {
                    xMaxFound = true;
                    this._evaluatedEdges.Add(edge);
                    boundEdges[1] = edge;
                    continue;
                }

                if (edge.Y1 == this._yMin && !this._evaluatedEdges.Contains(edge) && yMinFound == false)
                {
                    yMinFound = true;
                    this._evaluatedEdges.Add(edge);
                    boundEdges[2] = edge;
                    continue;
                }

                if (edge.Y1 == this._yMax && !this._evaluatedEdges.Contains(edge) && yMaxFound == false)
                {
                    yMaxFound = true;
                    this._evaluatedEdges.Add(edge);
                    boundEdges[3] = edge;
                    continue;
                }
            }
            return boundEdges;
        }
        public void DimensionEdges(ApplicationMgr mgr, dimEdge edge1, dimEdge edge2)
        {
            try
            {
                ModelDoc2 dwDoc = (ModelDoc2)mgr.App.ActiveDoc;
                mgr.PushRef(dwDoc);

                dwDoc.Extension.SelectByID2("", "", edge1.X1, edge1.Y1, 0, true, 0, null, 0);
                dwDoc.Extension.SelectByID2("", "", edge2.X1, edge2.Y1, 0, true, 0, null, 0);

                dwDoc.AddDimension(0, 0, 0);
            }
            finally
            {
                Console.WriteLine("Tearing down substack...  (DimensionEdges)");
                mgr.ClearSubStack();
            }
        }

        private void StoreBounds(List<dimEdge> edgeStructs)
        {
            double xMin = 100;
            foreach (dimEdge edge in edgeStructs)
            {
                double x = edge.X1;
                if ( edge.X1 < xMin) xMin = x;
            }
            this._xMin = xMin;
            Console.WriteLine("Min x value stored = " + this._xMin);

            double xMax = 0;
            foreach (dimEdge edge in edgeStructs)
            {
                double x = edge.X1;
                if ( edge.X1 > xMax) xMax = x;
            }
            this._xMax = xMax;
            Console.WriteLine("Max x value stored = " + this._xMax);

            double yMin = 100;
            foreach (dimEdge edge in edgeStructs)
            {
                double y = edge.Y1;
                if ( edge.Y1 < yMin) yMin = y;
            }
            this._yMin = yMin;
            Console.WriteLine("Min y value stored = " + this._yMin);

            double yMax = 0;
            foreach (dimEdge edge in edgeStructs)
            {
                double y = edge.Y1;
                if ( edge.Y1 > yMax) yMax = y;
            }
            this._yMax = yMax;
            Console.WriteLine("Max y value stored = " + this._yMax);
        }

        public double[] TransformModelToSheetSpace(ApplicationMgr mgr, View view, double[] modelCoords)
        {
            try
            {
                // Math utility.
                MathUtility mathUtil = (MathUtility)mgr.App.GetMathUtility();
                mgr.PushRef(mathUtil);

                MathTransform viewTransform = view.ModelToViewTransform;
                mgr.PushRef(viewTransform);

                double[] point = new double[3];
                point[0] = modelCoords[0];
                point[1] = modelCoords[1];
                point[2] = modelCoords[2];

                MathPoint modelPt = (MathPoint)mathUtil.CreatePoint(point);
                mgr.PushRef(modelPt);

                //transform to sheet space
                MathPoint sheetPt = (MathPoint)modelPt.MultiplyTransform(viewTransform);
                mgr.PushRef(sheetPt);

                // pull the coords out. 
                double[] sheetCoords = (double[])sheetPt.ArrayData;

                return sheetCoords;
            }
            finally
            {
                mgr.ClearSubStack();
            }
        }

        public double[] TransformDirectionModelToSheet(ApplicationMgr mgr, View view, double[] directionCoords)
        {
            try
            {
                MathUtility mathUtil = (MathUtility)mgr.App.GetMathUtility();
                mgr.PushRef(mathUtil);

                MathTransform viewTransform = view.ModelToViewTransform;
                mgr.PushRef(viewTransform);

                //create vector in model space
                double[] dir = new double[3];
                dir[0] = directionCoords[3];
                dir[1] = directionCoords[4];
                dir[2] = directionCoords[5];

                MathVector modelVec = (MathVector)mathUtil.CreateVector(dir);
                mgr.PushRef(modelVec);

                // transformw hile ignoring translations
                MathVector sheetVec = (MathVector)modelVec.MultiplyTransform(viewTransform);
                mgr.PushRef(sheetVec);

                double[] sheetDirection = (double[])sheetVec.ArrayData;

                
                // NORMALIZE the vector to unit length
                double magnitude = Math.Sqrt(
                    sheetDirection[0] * sheetDirection[0] +
                    sheetDirection[1] * sheetDirection[1] +
                    sheetDirection[2] * sheetDirection[2]
                );
        
                if (magnitude > 1e-10)  // Avoid division by zero
                {
                    sheetDirection[0] /= magnitude;
                    sheetDirection[1] /= magnitude;
                    sheetDirection[2] /= magnitude;
                }
                else 
                {
                    Console.WriteLine($"  WARNING: Magnitude too small to normalize!");
                }
                

                return sheetDirection;
            }
            finally
            {
                mgr.ClearSubStack();
            }
        }
    }
}