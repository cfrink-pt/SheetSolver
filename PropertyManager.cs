using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Text.RegularExpressions;
using System.Windows.Forms;

/*
Properties we can auto-fetch
Drawn By
Surface Area

Properties we must poll for
Features? (insert, bend, weld, coat, etc)

*/
namespace SheetSolver
{
    class PropertyManager
    {
        public string UserInitials { get; set; }
        public double SurfaceArea { get; set; }

        public PropertyManager()
        {
            
        }

        public static bool getRegexValidation(string input, string expression)
        {
            Match match = Regex.Match(input, expression);
            if (match.Success)
            {
                return true;
            }
            return false;
        }

        public string GetUserInitials()
        {
            string currentUserName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string pattern = @"PHASE\\([a-zA-Z]{2})";

            Match match = Regex.Match(currentUserName, pattern);

            if (match.Success)
            {
                Console.WriteLine("Auto-Fetched user name: " + match.Groups[1].Value.ToUpper());
                return match.Groups[1].Value.ToUpper();
            }
            else
            {

                Console.WriteLine($"User system name: {currentUserName} did not match regex expression.");
                string userInput;
                do
                {

                    userInput = InputDialog.Show("Please enter your initials:");
                    if (userInput == null)
                    {
                        throw new UserCancelledException("User cancelled drawing creation at Initial entry position after improper system name fetching for custom properties.");
                    }

                // validate the user inputted a initial. 
                } while (!getRegexValidation(userInput, "^[A-Za-z]{2,3}$"));

                return userInput.ToUpper();

            }
        }
    }
    public static class InputDialog
    {
        public static string Show(string prompt, string title = "Input", string defaultValue = "")
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = prompt;
            textBox.Text = defaultValue;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(10, 10, 280, 20);
            textBox.SetBounds(10, 35, 280, 20);
            buttonOk.SetBounds(120, 70, 80, 25);
            buttonCancel.SetBounds(210, 70, 80, 25);

            form.ClientSize = new System.Drawing.Size(300, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterParent;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult result = form.ShowDialog();
            return result == DialogResult.OK ? textBox.Text : null;
        }
    }
}