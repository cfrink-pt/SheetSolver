using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using System.Collections.Generic;
using System;
using SolidWorks.Interop.swconst;
using System.Threading;

namespace SheetSolver
{
    public class ApplicationMgr
    {
        public Dictionary<string, bool> sheetPreferences { get; set; }
        public string partFileName { get; set; }
        public string partFileDir { get; set; }
        public string assyFileName { get; set; }
        public string assyFileDir { get; set; }
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

            StoreOpenDoc();

            sheetPreferences = new Dictionary<string, bool>
            {
                { "Insert", false },
                { "Bend", false },
                { "Weld", false}
            };
        }

        /// <summary>
        ///  Presents the user with options for which sheets this part requires. Stores resulting values within mgr.sheetPreferences.
        /// </summary>
        public void PollSheetPreference()
        {
            var form = new Form { Text = "Select Sheets", Width = 350, Height = 80 + sheetPreferences.Count * 30, StartPosition = FormStartPosition.CenterScreen };

            int y = 10;
            foreach (var key in new List<string>(sheetPreferences.Keys))
            {
                var cb = new CheckBox { Text = key, Checked = sheetPreferences[key], Left = 10, Top = y, Width = 200 };
                string k = key; 
                cb.CheckedChanged += (s, e) => sheetPreferences[k] = cb.Checked;
                form.Controls.Add(cb);
                y += 25;
            }

            var ok = new Button { Text = "OK", Left = 10, Top = y, DialogResult = DialogResult.OK };
            var cancel = new Button { Text = "Cancel", Left = 90, Top = y, DialogResult = DialogResult.Cancel };
            form.Controls.Add(ok);
            form.Controls.Add(cancel);
            form.AcceptButton = ok;
            form.CancelButton = cancel;

            if (form.ShowDialog() == DialogResult.Cancel)
            {
                throw new UserCancelledException("User cancelled sheet selection.");
            }
        }

        /// <summary>
        ///  Stores necessary metadata based on sheet selection within Application Manager.
        /// </summary>
        public void EvaluateSheetPreferences()
        {
            if (sheetPreferences["Insert"])
            {
                PollUserForInsertAssembly();
            }
            else
            {
                // do some sneaky checks to see if they can and should make an insert sheet but didnt.
                string thisFile = Doc.GetPathName();

                var thisInfo = new FileInfo(thisFile);
                this.partFileDir = thisInfo.DirectoryName;
                this.partFileName = thisInfo.Name;

                string insertAssyFile = partFileName.Replace(".SLDPRT", ".SLDASM");

                try
                {
                    // try for auto fetch BEHIND THE SCENES.
                    int e = 0, w = 0;
                    ModelDoc2 insertAssy = App.OpenDoc6(insertAssyFile, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref e, ref w);

                    var assyInfo = new FileInfo(insertAssy.GetPathName());
                    this.assyFileDir = assyInfo.DirectoryName;
                    this.assyFileName = assyInfo.Name;

                    // open our original doc. 
                    App.OpenDoc6(thisFile, (int)swDocumentTypes_e.swDocPART, 0, "", e, w);     
                    Marshal.ReleaseComObject(insertAssy);

                    // if we got to this point without throwing a nullref exception. the user may have forgotten about a potential file.
                    DialogResult result = MessageBox.Show(
                        $"Found \"{assyFileName}\" within \"{assyFileDir}\".\n\nWould you like to generate an insert sheet using this assembly?",
                        "Insert Assembly Found",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        sheetPreferences["Insert"] = true;
                    }
                }
                catch (NullReferenceException)
                {
                    // we sneakily tried to find an assembly and we could not find it.
                    // ok, the user was not wrong in their assessment of no valid inserts assembly. Proceed as normal.
                    this.assyFileDir = "";
                    this.assyFileName = "";
                }
            }

            if (sheetPreferences["Bend"])
            {
                
            }
            if (sheetPreferences["Weld"])
            {
                
            }
        }

        /// <summary>
        /// Ask the user the directory of the .asm file associated with this solidworks part.
        /// Also validates that the selected .asm file has a valid name.
        /// </summary>
        private void PollUserForInsertAssembly()
        {
            string thisFile = Doc.GetPathName();

            var thisInfo = new FileInfo(thisFile);
            this.partFileDir = thisInfo.DirectoryName;
            this.partFileName = thisInfo.Name;

            string insertAssyFile = partFileName.Replace(".SLDPRT", ".SLDASM");

            try
            {
                // try for auto fetch.
                int e = 0, w = 0;
                ModelDoc2 insertAssy = App.OpenDoc6(insertAssyFile, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref e, ref w);

                var assyInfo = new FileInfo(insertAssy.GetPathName());
                this.assyFileDir = assyInfo.DirectoryName;
                this.assyFileName = assyInfo.Name;

                // open our original doc. 
                App.OpenDoc6(thisFile, (int)swDocumentTypes_e.swDocPART, 0, "", e, w);     
                Marshal.ReleaseComObject(insertAssy);      
            }
            catch (NullReferenceException)
            {
                // auto fetch failed. begin find new file loop.
                bool validAssyFileFound = false;
                while(!validAssyFileFound)
                {
                    Console.WriteLine("Failed to auto-fetch insert assembly. Polling user...");
                    var form = new Form { Text = "Insert Assembly Location", Width = 675, Height = 165, StartPosition = FormStartPosition.CenterScreen };
                    var label = new Label { Text = "Assembly File:", Left = 10, Top = 18, Width = 80 };
                    var txtBox = new TextBox { Text = "Directory: ", Lines = [$"{thisFile}"], Left = 10, Top = 45, Width = 500 };
                    var browse = new Button { Text = "Browse...", Left = 515, Top = 45, Width = 80 };
                    browse.Click += (s, e) =>
                    {
                        string selectedFile = null;

                        var thread = new Thread(() =>
                        {
                            using var fileDialog = new OpenFileDialog
                            {
                                Title = "Select Assembly File",
                                Filter = "Assembly Files (*.SLDASM)|*.SLDASM|All Files (*.*)|*.*",
                                InitialDirectory = Path.GetDirectoryName(partFileDir) ?? ""
                            };
                            if (fileDialog.ShowDialog() == DialogResult.OK)
                            {
                                selectedFile = fileDialog.FileName;
                            }
                        });

                        thread.SetApartmentState(ApartmentState.STA);
                        thread.Start();
                        thread.Join();

                        if (selectedFile != null)
                        {
                            txtBox.Text = selectedFile;
                        }
                    };

                    var ok = new Button { Text = "OK", Left = 10, Top = 80, DialogResult = DialogResult.OK };
                    var cancel = new Button { Text = "Cancel", Left = 90, Top = 80, DialogResult = DialogResult.Cancel };

                    form.Controls.Add(label);
                    form.Controls.Add(txtBox);
                    form.Controls.Add(browse);
                    form.Controls.Add(ok);
                    form.Controls.Add(cancel);
                    form.AcceptButton = ok;
                    form.CancelButton = cancel;

                    if (form.ShowDialog() == DialogResult.Cancel)
                    {
                        throw new UserCancelledException("User cancelled sheet selection.");
                    }
                    
                    string userSelectedPath = txtBox.Text;
                    Console.WriteLine("DEBUG: " + userSelectedPath);

                    var userPath = new FileInfo(userSelectedPath);

                    // now check if it is valid.
                    try
                    {
                        // now lets check if that assembly has our part inside it.
                        // we could traverse the whole assembly, but lets just first check if the names line up.
                        if (partFileName.Substring(0, partFileName.Length - 7) != userPath.Name.Substring(0, userPath.Name.Length - 7)) // -7 here is what truncates the extension. messy but it kind of works. maybe.
                        {
                            throw new NullReferenceException();
                        }
                        
                        
                        int er = 0, wa = 0;
                        ModelDoc2 insertAssy = App.OpenDoc6(userSelectedPath, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref er, ref wa);

                        var assyInfo = new FileInfo(insertAssy.GetPathName());
                        this.assyFileDir = assyInfo.DirectoryName;
                        this.assyFileName = assyInfo.Name;

                        // open our original doc. 
                        App.OpenDoc6(thisFile, (int)swDocumentTypes_e.swDocPART, 0, "", er, wa);     
                        
                        validAssyFileFound = true;
                        Console.WriteLine("DEBUG: User inputted a valid assembly file.");




                        Marshal.ReleaseComObject(insertAssy);
                    }
                    catch (NullReferenceException)
                    {
                        MessageBox.Show($"{userSelectedPath} \r\n - is not a valid assembly file. Please select a valid .SLDASM file which contains your opened part file. Is your .SLDASM file you selected named identically to your part file?");
                    }
                }
            }
        }

        public DrawingDoc CreateAndMoveToDrawing()
        {
            this.Doc.NameView(this.viewName);
            this.App.NewDocument(this.drawingTemplate, 0, 0, 0);

            DrawingDoc swDrawing = (DrawingDoc)this.App.ActiveDoc;
            return swDrawing;
        }
        private void StoreOpenDoc()
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
            if (this.subStack.Count == 0)  // check count instead of Peek()
            {
                return; // nothing to clear, just exit silently
            }

            while (subStack.Count > 0)  // use while loop, safer than for loop here
            {
                object currentRef = subStack.Pop();
                if (currentRef != null) 
                {
                    Marshal.ReleaseComObject(currentRef);
                }
            }
        }

        public void TearDown()
        {
            while (mainComStack.Count > 0)
            {
                object currentRef = mainComStack.Pop();
                if (currentRef != null)
                {
                    Marshal.ReleaseComObject(currentRef);
                }
            }
        }
    }
}