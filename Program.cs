using System;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SheetSolver
{
    public class UserCancelledException : Exception
    {
        public UserCancelledException(string message) : base(message) { }
    }

    class Orchestrator
    {
        static void Main(string[] args)
        {
            Coordinator coordinator = new Coordinator();

            try
            {
                coordinator.CreateDrawing();
            }
            catch (UserCancelledException ex)
            {
                Console.WriteLine(ex.Message);
                return;
            }
            catch (InvalidOperationException ex)
            {
                Console.WriteLine(ex.Message);
                MessageBox.Show(ex.Message);
                return;
            }
        }
    }
}