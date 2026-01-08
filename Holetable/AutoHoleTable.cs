using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Linq;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using SwView = SolidWorks.Interop.sldworks.View;
    //TODO: Eliminate all in-model evaluations for hole matching.
    // Instead, I implemented a read-only dict which defines features
    // from drawing space TableAnnotation objects. 
namespace SheetSolver
{
    public enum HoleType
    {
        [Description("10-32 E&T")]
        _1032E,
        [Description("8-32 E&T")]
        _832E,
        [Description("6-32 E&T")]
        _632E,
        [Description("1/4-20 E&T")]
        _2520E,
        [Description("0.875 Grommet")]
        _875G,
        [Description("Bridge Punch")]
        _060BP,
        [Description("Unknown, or Through-Hole")]
        Unknown
    }
    public class ToolingIdentifier
    {
        // inner surface areas of standard toolings. Using this dict to identify candidate toolings.
        private static readonly Dictionary<HoleType, double> KnownAreas = new()
        {
            { HoleType._1032E, 0.00000809 },
            { HoleType._832E, 0.00000323 },
            { HoleType._632E, 0.00000297 },
            { HoleType._875G, 0.00013301 },
            { HoleType._060BP, 0.00000584 }
        };

        public static HoleType Identify(double surfaceArea)
        {
            double formIDTol = 0.00000005;
            
            foreach (KeyValuePair<HoleType, double> kvp in KnownAreas)
            {
                if (Math.Abs(surfaceArea - kvp.Value) <= formIDTol )
                {
                    return kvp.Key;
                }
            }

            return HoleType.Unknown;
        }
    }
    public class HoleCandidate
    {
        // Loop identification
        public int LoopIndex { get; set; }

        // Edge Geometry
        public double CenterX { get; set; }
        public double CenterY { get; set; }
        public double CenterZ { get; set; }
        public double AxisX { get; set; }
        public double AxisY { get; set; }
        public double AxisZ { get; set; }
        public double Radius { get; set; }

        // Surface normals for the two adjacent faces
        public double[] ThisFaceNormal { get; set; }
        public double[] PartnerFaceNormal { get; set; }

        // Tangent at mid-coedge
        public double[] EdgeTangent { get; set; }

        // GetBox double array of 6 items
        // [ XCorner1, YCorner1, ZCorner1, XCorner2, YCorner2, ZCorner2 ]
        public double[] BoxData { get; set; }

        //Hole type enum
        public HoleType holeType { get; set; }

        // Reference back to edge for selection (keep this COM ref until final selection)
        public Edge EdgeRef { get; set; }
    }
    public class HoleDataExtractor
    {
        private VectorHelper vHelp = new VectorHelper();

        // lets extract all geometric data first so we minimize the time we hold on to com references.

        public List<HoleCandidate> ExtractHoleCandidates(List<Face2> faces, UserProgressBar pbar)
        {
            List<HoleCandidate> candidates = new List<HoleCandidate>();
            int loopIndex = 0;

            int faceIndex = 0;
            int totalFaces = faces.Count;

            foreach (Face2 swFace in faces)
            {
                pbar.UpdateTitle($"Scanning face {faceIndex+1} of {totalFaces}");
                pbar.UpdateProgress(faceIndex * 100 / totalFaces);
                System.Windows.Forms.Application.DoEvents();

                Loop2 swLoop = (Loop2)swFace.GetFirstLoop();

                while (swLoop != null)
                {
                    // Reject quickly. Anything that is an outer loop of a face, or a non 1-edged loop, is def not a hole.
                    if (swLoop.IsOuter() || swLoop.GetEdgeCount() != 1)
                    {
                        Loop2 nextLoop = (Loop2)swLoop.GetNext();
                        Marshal.ReleaseComObject(swLoop);
                        swLoop = nextLoop;
                        continue;
                    }

                    // This loop may be a hole. extract needed info.
                    HoleCandidate candidate = TryExtractCandidate(swLoop, loopIndex);

                    if (candidate != null)
                    {
                        candidates.Add(candidate);
                    }

                    loopIndex++;

                    // move to the next loop, release current
                    Loop2 tempLoop = swLoop;
                    swLoop = (Loop2)swLoop.GetNext();
                    Marshal.ReleaseComObject(tempLoop);
                }

                faceIndex++;
                
            }

            return candidates;

        }

        private double[] getFaceBoxDataFromCoEdge(CoEdge swCoedge)
        {
            double[] boxData;
            Face2 face = null;
            Loop2 loop = null;

            loop = (Loop2)swCoedge.GetLoop();
            face = (Face2)loop.GetFace();

            boxData = (double[])face.GetBox();

            if (loop != null) Marshal.ReleaseComObject(loop);
            if (face != null) Marshal.ReleaseComObject(face);

            return boxData;
        }

        private HoleCandidate TryExtractCandidate(Loop2 swLoop, int loopIndex)
        {
            CoEdge swThisCoedge = null;
            CoEdge swPartnerCoedge = null;
            Edge swEdge = null;
            Curve swCurve = null;
            Loop2 partnerLoop = null;
            Face2 partnerFace = null;
            double partnerFaceSurfArea = 0;

            try
            {
                // grab coedges
                swThisCoedge = (CoEdge)swLoop.GetFirstCoEdge();
                swPartnerCoedge = (CoEdge)swThisCoedge.GetPartner();

                //grab single edge and its curve
                object[] edgeArr = (object[])swLoop.GetEdges();
                swEdge = (Edge)edgeArr[0];
                swCurve = (Curve)swEdge.GetCurve();

                //Check circularity
                if (!swCurve.IsCircle())
                {
                    return null;
                }

                // Grab partner face surface area to idenfity hole type, if E&T, etc.
                partnerLoop = (Loop2)swPartnerCoedge.GetLoop();
                partnerFace = (Face2)partnerLoop.GetFace();

                partnerFaceSurfArea = partnerFace.GetArea();

                // now extract the needed data for C# ops
                double[] circleParams = (double[])swCurve.CircleParams;
                double[] thisNormal = vHelp.GetFaceNormalAtMidCoEdge(swThisCoedge);
                double[] partnerNormal = vHelp.GetFaceNormalAtMidCoEdge(swPartnerCoedge);
                double[] tangent = vHelp.GetTangentAtMidCoEdge(swThisCoedge);

                // Grab the box
                double[] partnerFaceBoxParams = getFaceBoxDataFromCoEdge(swPartnerCoedge);

                // build our data container
                return new HoleCandidate
                {
                    LoopIndex = loopIndex,
                    CenterX = circleParams[0],
                    CenterY = circleParams[1],
                    CenterZ = circleParams[2],
                    AxisX = circleParams[3],
                    AxisY = circleParams[4],
                    AxisZ = circleParams[5],
                    Radius = circleParams[6],
                    ThisFaceNormal = thisNormal,
                    PartnerFaceNormal = partnerNormal,
                    EdgeTangent = tangent,
                    BoxData = partnerFaceBoxParams,
                    holeType = ToolingIdentifier.Identify(partnerFaceSurfArea),
                    EdgeRef = swEdge // this ref is stored purely for later selection
                };
            }
            finally
            {
                //release everything except the stored edge for selection
                if (partnerFace != null) Marshal.ReleaseComObject(partnerFace);
                if (partnerLoop != null) Marshal.ReleaseComObject(partnerLoop);
                if (swCurve != null) Marshal.ReleaseComObject(swCurve);
                if (swPartnerCoedge != null) Marshal.ReleaseComObject(swPartnerCoedge);
                if (swThisCoedge != null) Marshal.ReleaseComObject(swThisCoedge);
            }
        }
    }
    public class HoleFilter
    {
        public List<HoleCandidate> FilterValidHoles(List<HoleCandidate> candidates, double exclusionRadius)
        {

            VectorHelper vHelp = new VectorHelper();

            List<HoleCandidate> validHoles = new List<HoleCandidate>();

            foreach (HoleCandidate candidate in candidates)
            {
                // Filter 1: Size check
                if (candidate.Radius >= exclusionRadius)
                {
                    // Too big - skip it, but release the edge reference
                    Marshal.ReleaseComObject(candidate.EdgeRef);
                    candidate.EdgeRef = null;
                    continue;
                }

                // Filter 2: Non-parallel normals (indicates a transition between surfaces)
                if (vHelp.VectorsAreParallel(candidate.ThisFaceNormal, candidate.PartnerFaceNormal))
                {
                    Marshal.ReleaseComObject(candidate.EdgeRef);
                    candidate.EdgeRef = null;
                    continue;
                }

                // Filter 3: Cross product / tangent check (confirms it's a hole, not a boss)
                double[] crossProduct = vHelp.GetCrossProduct(
                    candidate.ThisFaceNormal, 
                    candidate.PartnerFaceNormal
                );

                if (!vHelp.VectorsAreParallel(crossProduct, candidate.EdgeTangent))
                {
                    // Cross product unit vec did not point in the same direction as the unit vector of the edge tangency. Probably a boss extrude or not a perfect hole.
                    Marshal.ReleaseComObject(candidate.EdgeRef);
                    candidate.EdgeRef = null;
                    continue;
                }

                // Final Filter 4: Validate Partner face Box Data does not indicate this is a scribe.
                if (vHelp.CheckIfBoxDataIsScribe(candidate.BoxData))
                {
                    // Skip this, it's partner face has a depth less than that of a standard scribe.
                    Marshal.ReleaseComObject(candidate.EdgeRef);
                    candidate.EdgeRef = null;
                    continue;
                }

                // Passed all checks - this is a valid hole
                validHoles.Add(candidate);
            }

            return validHoles;
        }
    }
    public class HoleSelector
    {
        public int SelectHoles(List<HoleCandidate> validHoles, SelectData swSelData)
        {
            int selectedCount = 0;

            swSelData.Mark = 2;

            foreach (HoleCandidate hole in validHoles)
            {
                try
                {
                    // add this hole to selection buffer
                    if (hole.EdgeRef != null)
                    {
                        Entity swEntity = (Entity)hole.EdgeRef;
                        swEntity.Select4(true, swSelData);
                        selectedCount++;
                    }
                }
                finally
                {
                    // drop the reference
                    if (hole.EdgeRef != null)
                    {
                        Marshal.ReleaseComObject(hole.EdgeRef);
                        hole.EdgeRef = null;
                    }
                }
            }

            return selectedCount;
        }
    }
    public class TableProcessor
    {
        private static readonly Dictionary<string, string> knownFormingTools = new()
        {
            { "<MOD-DIAM>0.150<HOLE-DEPTH>0.019", "10-32 E&T" },
            { "<MOD-DIAM>0.150<HOLE-DEPTH>0.027", "10-32 E&T" },
            // Todo -> Better evaluation here. This is a not great solution to begin with,
            // but it will get the job done. May result in improper evaluations. 
            { "<MOD-DIAM>0.178", "10-32 TAP" },
            { "<MOD-DIAM>0.150<HOLE-DEPTH>0.086", "8-32 E&T" },
            { "<MOD-DIAM>0.150<HOLE-DEPTH>0.016", "8-32 E&T" },
            { "<MOD-DIAM>0.150<HOLE-DEPTH>0.011", "8-32 E&T" },
            // same here.
            { "<MOD-DIAM>0.189", "1/4-20 E&T" },
            { "<MOD-DIAM>0.138<HOLE-DEPTH>0.011", "6-32 E&T" },
            { "<MOD-DIAM>0.060 THRU", "BRIDGE PUNCH"},
            { "<MOD-DIAM>0.875<HOLE-DEPTH>0.075", "0.875in GROMMET" }
        };

        private Dictionary<int, string> GetFormingToolRows(ITableAnnotation table)
        {
            Dictionary<int, string> formingToolRows = new();

            int rowCount = table.RowCount;
            
            for (int i = 0; i < rowCount; i++)
            {
                foreach (var formingtool in knownFormingTools)
                {
                    if (table.Text[i, 1] == formingtool.Key)
                    {
                        formingToolRows.Add(i, formingtool.Value);
                        break;
                    }
                }
            }

            return formingToolRows;
        }

        private void EditHoleTable(ITableAnnotation table, Dictionary<int, string> formingToolRows)
        {
            int shift = 0;
            int loc = 0;
            foreach (var row in formingToolRows)
            {
                loc = row.Key + shift;
                table.InsertRow((int)swTableItemInsertPosition_e.swTableItemInsertPosition_After, loc);
                table.Text[loc+1, 1] = row.Value;

                table.MergeCells(loc, 0, loc+1, 0);
                table.MergeCells(loc, 2, loc+1, 2);

                shift++;
            }
        }

        public void generateHoletable(SwView view, string template)
        {
            HoleTableAnnotation swHoleTblAnnotation = view.InsertHoleTable2(false, 0.01416083, .206895, 1, "A", template);

            Annotation holeTableDatum = swHoleTblAnnotation.HoleTable.DatumOrigin.GetAnnotation();
            holeTableDatum.Visible = (int)swAnnotationVisibilityState_e.swAnnotationHidden;

            Marshal.ReleaseComObject(holeTableDatum);

            // cast hole table to standard table annotation.
            ITableAnnotation swTableAnnotation = (ITableAnnotation)swHoleTblAnnotation;

            Dictionary<int, string> formingToolRows = GetFormingToolRows(swTableAnnotation);

            EditHoleTable(swTableAnnotation, formingToolRows);


            Marshal.ReleaseComObject(swTableAnnotation);
        }
    }
    public class HoleProcessor
{
    private HoleDataExtractor extractor = new HoleDataExtractor();
    private HoleFilter filter = new HoleFilter();
    private HoleSelector selector = new HoleSelector();

    //private TableProcessor generator = new TableProcessor();

    private void DebugPrint(List<HoleCandidate> validHoles)
    {
        var counts = validHoles
            .GroupBy(h => h.holeType)
            .ToDictionary(g => g.Key, g => g.Count());

        foreach (var kvp in counts)
            Console.WriteLine($"{kvp.Key}: {kvp.Value}");

        Console.ReadLine();
    }
    
    public int ProcessFaces(List<Face2> faces, SelectData swSelData, double exclusionRadius, UserProgressBar pbar)
    {
        // return a list of holecandidate objects called candidates. 
        List<HoleCandidate> candidates = extractor.ExtractHoleCandidates(faces, pbar);
        
        List<HoleCandidate> validHoles = filter.FilterValidHoles(candidates, exclusionRadius);
        
        int count = selector.SelectHoles(validHoles, swSelData);

        //DebugPrint(validHoles);
        
        return count;
    }
}
    public class HoleTableEntrance
    {
        static UserProgressBar CreateProgressBar(SldWorks swApp, string title)
        {
            UserProgressBar pbar;
            swApp.GetUserProgressBar(out pbar);
            pbar.Start(0, 100, title);
            return pbar;
        }

        static SldWorks bindToApplication()
        {
            Console.Clear();
            Console.WriteLine("Connecting to Solidworks...");

            SldWorks swApp = null;

            try
            {
                swApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to connect to SolidWorks: " + ex.Message);
            }
            
            // Check if swApp is null. If so, we will need to throw an error. really isnt possible but, eh 
            if (swApp == null)
            {
                Console.Clear();
                Console.WriteLine("Failed to connect to solidworks. Make sure it is open, duh. Why even run this without solidworks?");
            }

            Console.WriteLine("Connected to Solidworks Successfully");

            swApp.Visible = true;

            return swApp;
        }

        static SwView getViewObject(SelectionMgr swSelMgr, ModelDoc2 swModel)
        {

            SwView swView = null;

            while (swView == null)
            {

                var swSelection = swSelMgr.GetSelectedObject6(1, -1);
                int selType = swSelMgr.GetSelectedObjectType3(1, -1);
                
                if ((swSelectType_e)selType == swSelectType_e.swSelDRAWINGVIEWS)
                {
                    swView = (SwView)swSelection;
                    swModel.ClearSelection2(true);
                    swSelMgr.DeSelect2(1, -1);
                    break;
                }
                else
                {
                    if (swSelection != null)
                    {
                        // if the user selected something that isnt a view object, but is still a com object.
                        Marshal.ReleaseComObject(swSelection);
                    }

                    //Console.Write("Please select a drawing view. \r\n");
                    //Console.WriteLine($"{(swSelectType_e)selType}");
                    //Console.ReadLine();                

                    MessageBox.Show("Please select a valid drawing view and run the macro again.");
                    break;
                }
            }
            
            swSelMgr.DeSelect2(1, -1);
            swModel.ClearSelection2(true);

            return swView;
        }

        public void DoHoleTable(ApplicationMgr mgr)
        {
            // Init model, view and selection manager obj
            ModelDoc2 swModel = (ModelDoc2)mgr.App.ActiveDoc;
            SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
            SwView swView = getViewObject(swSelMgr, swModel);

            // validate selection is a drawing view
            if (swView == null)
            {
                Console.WriteLine("Invalid Selection. Breaking...");
                Marshal.ReleaseComObject(swSelMgr);
                Marshal.ReleaseComObject(swModel);
                return;
            }

            // get a list of faces to evaluate
            List<Face2> swFaceList = new List<Face2>(); 

            object[] entityList = (object[])swView.GetVisibleEntities2(null, (int)swViewEntityType_e.swViewEntityType_Face);
            foreach (object obj in entityList)
            {
                Entity swEnt = (Entity)obj;
                int entType = (int)swEnt.GetType();
                if (entType == (int)swSelectType_e.swSelFACES)
                {
                    swFaceList.Add((Face2)swEnt);
                }
                else
                    Marshal.ReleaseComObject(swEnt);
            }

            UserProgressBar progressBar = CreateProgressBar(mgr.App, "Finding holes...");

            // Process holes
            SelectData swSelData = (SelectData)swSelMgr.CreateSelectData();
            double exclusionRadius = 0.01775;

            HoleProcessor holeProcessor = new HoleProcessor();
            holeProcessor.ProcessFaces(swFaceList, swSelData, exclusionRadius, progressBar);

            progressBar.UpdateTitle("Creating hole table...");

            // RELEASE ALL FACES.
            foreach (Face2 face in swFaceList)
            {
                Marshal.ReleaseComObject(face);
            }
            swFaceList.Clear();


            // Generate a hole table and store the table annotation object it returns for edits.
            TableProcessor holeTableProcessor = new TableProcessor();
            string holeTableTemplate = @"\\storage\CAD\Solidworks\Phase Setting files\Tables\HTS Hole Table.sldholtbt";
            holeTableProcessor.generateHoletable(swView, holeTableTemplate);

            progressBar.End();
            Marshal.ReleaseComObject(progressBar);
            Marshal.ReleaseComObject(swSelData);
            Marshal.ReleaseComObject(swView);
            Marshal.ReleaseComObject(swSelMgr);
            Marshal.ReleaseComObject(swModel);
        }
    }
}