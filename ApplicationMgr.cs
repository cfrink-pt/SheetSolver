
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using System.Collections.Generic;
using System;
using SolidWorks.Interop.swconst;

namespace SheetSolver
{
    public class ApplicationMgr
    {
        public double drawingX = 0.2794;
        public double drawingY = 0.2159;
        public string drawingTemplate = @"\\storage\CAD\Solidworks\Phase Setting files\Templates\Phase Drawing, 4.0.drwdot";
        public string viewName = "TempView";
        public Stack<object> mainComStack = new Stack<object>();
        public Stack<object> subStack = new Stack<object>();
        public SldWorks App { get; set; }
        public ModelDoc2 Doc { get; set; }

        public ApplicationMgr()
        {
            try
            {
                App = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                mainComStack.Push(App);
            }
            catch (COMException)
            {
                MessageBox.Show("SolidWorks needs to be open before running this macro.");
                throw;
            }
        }

        public DrawingDoc CreateAndMoveToDrawing()
        {
            this.Doc.NameView(this.viewName);
            this.App.NewDocument(this.drawingTemplate, 0, 0, 0);

            DrawingDoc swDrawing = (DrawingDoc)this.App.ActiveDoc;
            return swDrawing;
        }
        public void StoreOpenDoc()
        {
            if (this.App.ActiveDoc != null)
            {
                ModelDoc2 swDoc = (ModelDoc2)this.App.ActiveDoc;
                this.Doc = swDoc;
                mainComStack.Push(swDoc);
            }
            else
            {
                throw new InvalidOperationException("No Active Document.");
            }
        }

        public bool VerifyDocType(swDocumentTypes_e docType)
        {
            bool result = false;

            if (this.Doc.GetType() == (int)docType)
            {
                result = true;
            }

            return result;
        }

        public void PushRef(object o)
        {
            if (o != null)
            {
                this.subStack.Push(o);
            }
            else
            {
                MessageBox.Show("PushRef(): No pushing null references to substack.");
                throw new NullReferenceException("Null references pushed to substack");
            }
        }

        public void ClearSubStack()
        {
            if (this.subStack.Peek() != null)
            {
                for (int i = 0; i < subStack.Count; i++)
                {
                    object currentRef = subStack.Peek();

                    if (currentRef != null) Marshal.ReleaseComObject(currentRef);

                    subStack.Pop();
                }
            }
            else
            {
                MessageBox.Show("ClearSubStack(): no references in substack to pop. Populate it with PushSubStack(object o) first.");
                throw new NullReferenceException("No references in substack to pop.");
            }
        }

        public void TearDown()
        {
            for (int i = 0; i < mainComStack.Count; i++)
            {
                object currentRef = mainComStack.Peek();

                if (currentRef != null) Marshal.ReleaseComObject(currentRef);

                mainComStack.Pop();
            }
        }
    }
}