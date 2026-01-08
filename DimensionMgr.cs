using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace SheetSolver
{
    class DimensionManager
    {
        public List<Curve> GetVisibleCurves(ApplicationMgr mgr, View view)
        {
            List<Curve> curves = new List<Curve>();

            object[] entityList = (object[])view.GetVisibleEntities2(null, (int)swViewEntityType_e.swViewEntityType_Face);
            foreach (object obj in entityList)
            {
                Entity swEnt = (Entity)obj;
                int entType = (int)swEnt.GetType();
                if (entType == (int)swSelectType_e.swSelEDGES)
                {
                    curves.Add((Curve)swEnt);
                }
                else
                {
                    Marshal.ReleaseComObject(swEnt);
                }
            }

            if (curves.Count != 0)
            {
                return curves;
            }
            else
            {
                throw new InvalidOperationException("GetVisibleCurves failed to fetch valid curve entities from view reference.");
            }
        }
    }
}