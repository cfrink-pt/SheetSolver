using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security.Policy;

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

        public DimensionManager(View view)
        {
            _view = view;
            _edges = new List<dimEdge>();
            _evaluatedEdges = new HashSet<dimEdge>();
        }

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

                            // Clean up POSITION coordinates (indices 0-2)
                            for (int i = 0; i < 3; i++)
                            {
                                lineParams[i] = Math.Round(lineParams[i], 8);
                                if (Math.Abs(lineParams[i]) < 1e-10)
                                {
                                    lineParams[i] = 0;
                                }
                            }
                            
                            // Clean up DIRECTION components (indices 3-5) with MUCH smaller threshold
                            // because these are normalized unit vectors
                            for (int i = 3; i < 6; i++)
                            {
                                lineParams[i] = Math.Round(lineParams[i], 8);
                                if (Math.Abs(lineParams[i]) < 1e-15)  // Much smaller threshold!
                                {
                                    lineParams[i] = 0;
                                }
                            }

                            Console.WriteLine($"Edge {count}\r\n  X: {lineParams[0]} | Y: {lineParams[1]} | Z: {lineParams[2]}\r\n  Dx: {lineParams[3]} | Dy: {lineParams[4]} | Dz: {lineParams[5]}");
                            dimEdge dEdge = new dimEdge
                            {
                                X1 = lineParams[0],
                                Y1 = lineParams[1],
                                Dx = lineParams[3],
                                Dy = lineParams[4],
                                Dz = lineParams[5]
                            };
                            edges.Add(dEdge);
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
                    Console.WriteLine($"  After normalize: ({sheetDirection[0]}, {sheetDirection[1]}, {sheetDirection[2]})");
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