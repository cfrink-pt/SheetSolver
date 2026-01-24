using SolidWorks.Interop.sldworks;
using System;
using System.Linq;

namespace SheetSolver
{
    class VectorHelper
    {
        // get the unit vector of any three dimensional vector. Returns array of 3 doubles.
        public double[] VecToUnitVec(double[] vec1)
        {
            double[] unitVec = null;
            double vMag = 0;

            vMag = Math.Pow((vec1[0] * vec1[0] + vec1[1] * vec1[1] + vec1[2] * vec1[2]), 0.5);
            unitVec = [(vec1[0] / vMag), (vec1[1] / vMag), (vec1[2] / vMag)];
            return unitVec;
        }

        public bool CheckIfBoxDataIsScribe(double[] boxData)
        {
            bool result = false;
            double[] absArray = new double[3];

            // box data:
            // [ XCorner1, YCorner1, ZCorner1, XCorner2, YCorner2, ZCorner2 ]

            absArray[0] = Math.Abs(boxData[0] - boxData[3]);
            absArray[1] = Math.Abs(boxData[1] - boxData[4]);
            absArray[2] = Math.Abs(boxData[2] - boxData[5]);

            double smallest = absArray.Min();

            if (smallest < 0.0001)
            {
                result = true;
            }

            return result;
        }

        public double[] GetFaceNormalAtMidCoEdge(CoEdge swCoEdge)
        {
            // declare the stuff we need to use. DUH. 
            Face2 swFace = default(Face2);
            // This surface will be used to give us our .EvaluateAtPoint normal values.
            Surface swSurface = default(Surface);
            Loop2 swLoop = default(Loop2);
            double[] varParams = null;
            double[] varPoint = null;
            double dblMidParam = 0;
            double[] dblNormal = new double[3];
            bool bFaceSenseReversed = false;

            // store a double of length 10 with curve params returned from the Coedge passed.
            varParams = (double[])swCoEdge.GetCurveParams();
            //calculating the midpoint of the coedge, regardless of direction.
            if (varParams[6] > varParams[7])
            {
                dblMidParam = (varParams[6] - varParams[7]) / 2 + varParams[7];
            }
            else
            {
                dblMidParam = (varParams[7] - varParams[6]) / 2 + varParams[6];
            }

            // Now we fetch the xyz double for the midpoint and store it.
            varPoint = (double[])swCoEdge.Evaluate(dblMidParam);

            // Now for the given coedge we need to find the face.
            // first though, we need to nab the loop that this passed coedge is referencing.
            swLoop = (Loop2)swCoEdge.GetLoop();
            // now we have the face for that loop (the coedge passed)
            swFace = (Face2)swLoop.GetFace();
            // then finally, the SURFACE which defines that face. (remember, it can have a different normal...)
            swSurface = (Surface)swFace.GetSurface();
            // This returns TRUE if the face and surface have opposite normals. false if they are the same. its stupid.
            bFaceSenseReversed = swFace.FaceInSurfaceSense();
            // The nice secret sauce. EvaluateAtPoint for ISurface interface calculates normal, principle directions, and principle curvatures of a surface @ that point. Muy bien.
            varParams = (double[])swSurface.EvaluateAtPoint(varPoint[0], varPoint[1], varPoint[2]);
            // True = surface normal is opposite to FACE normal.
            if (bFaceSenseReversed)
            {
                // REVERSE THAT SURFACE NORMAL. WE GOT TRUE AND APPARENTLY THAT MEANS ITS OPPOSITE TO THE FACE.
                dblNormal[0] = -varParams[0];
                dblNormal[1] = -varParams[1];
                dblNormal[2] = -varParams[2];
            }
            else
            {
                // surface normal and face normal are the same. so just store the calculated normal from the surf.
                dblNormal[0] = varParams[0];
                dblNormal[1] = varParams[1];
                dblNormal[2] = varParams[2];
            }
            return dblNormal;

        }

        public double[] GetTangentAtMidCoEdge (CoEdge swCoEdge)
        {
            double[] varParams = null;
            double dblMidParam = 0;
            double[] dblTangent = new double[3];
            // this is where we will do the SAMe thing we did in the other method to get the midpoint of the curve based on direction. Which really doesnt matter. I just pulled this snippet from the api docs.
            varParams = (double[])swCoEdge.GetCurveParams();
            if (varParams[6] > varParams[7])
            {
                dblMidParam = (varParams[6] - varParams[7]) / 2 + varParams[7];
            }
            else
            {
                dblMidParam = (varParams[7] - varParams[6]) / 2 + varParams[6];
            }
            // rather than the evaluate method on a surface, we will call it from the coedge interface to get tangency information rather than normal info. Returns array of 6 doubles.
            varParams = (double[])swCoEdge.Evaluate(dblMidParam);
            dblTangent[0] = varParams[3];
            dblTangent[1] = varParams[4];
            dblTangent[2] = varParams[5];
            return dblTangent;
        }
        public double[] GetCrossProduct (double[] varVec1, double[] varVec2)
        {
            double[] dblCross = new double[3];

            dblCross[0] = varVec1[1] * varVec2[2] - varVec1[2] * varVec2[1];
            dblCross[1] = varVec1[2] * varVec2[0] - varVec1[0] * varVec2[2];
            dblCross[2] = varVec1[0] * varVec2[1] - varVec1[1] * varVec2[0];

            return dblCross;
        }
        public bool VectorsAreParallel (double[] varVec1, double[] varVec2)
        {
            bool functionReturnValue = false;
            double dblMag = 0;
            double dblDot = 0;
            double[] dblUnit1 = new double[3];
            double[] dblUnit2 = new double[3];
            // get the vector's magnitude by squaring its components and taking the sqrt of the sum of all squares
            dblMag = Math.Pow((varVec1[0]*varVec1[0] + varVec1[1]*varVec1[1] + varVec1[2]*varVec1[2]), 0.5);
            // get the unit vector by dividing the vector by its magnitude for each leg
            dblUnit1[0] = varVec1[0] / dblMag;
            dblUnit1[1] = varVec1[1] / dblMag;
            dblUnit1[2] = varVec1[2] / dblMag;

            // do the same for the second vector 
            dblMag = Math.Pow((varVec2[0]*varVec2[0] + varVec2[1]*varVec2[1] + varVec2[2]*varVec2[2]), 0.5);
            dblUnit2[0] = varVec2[0] / dblMag;
            dblUnit2[1] = varVec2[1] / dblMag;
            dblUnit2[2] = varVec2[2] / dblMag;

            // take dot product.
            dblDot = dblUnit1[0] * dblUnit2[0] + dblUnit1[1] * dblUnit2[1] + dblUnit1[2] * dblUnit2[2];
            dblDot = Math.Abs(dblDot - 1.0);
            // dot product of one would indicate that the two unit vectors are parallel. we do some tolerancing here just in case, but its tiny.
            if (dblDot < 1E-10)
            {
                functionReturnValue = true;
            }
            else
            {
                functionReturnValue = false;
            }
            return functionReturnValue;
        }

        public double dotProduct (double[] vec1, double[] vec2)
        {
            double rValue = vec1[0] * vec2[0] + vec1[1] * vec2[1] + vec1[2] * vec2[2];
            return rValue;
        }
    }
}