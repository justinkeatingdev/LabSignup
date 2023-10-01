using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LabSignup
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            // decides which from runs
            Application.Run(new Agreement());
        }


    }

    public class LabInfo
    {
        public string LabName { get; set; }
        public string LabDay { get; set; }
        public string LabStart { get; set; }
        public string LabEnd { get; set; }
    }

    public class SigneeTitles
    {
        public string Title { get; set; }
    }

    public class SigneeInfo
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Title { get; set; }
        public string LabName { get; set; }
        public string LabDay { get; set; }
        public string LabStart { get; set; }
        public string LabEnd { get; set; }
        public string LabSignInTime { get; set; }
        public string LabHours { get; set; }
    }

    public class TitleInfo
    {
        public string Title { get; set; }
        public int Count { get; set; }
    }

    public class FacilitatorsInfo
    {
        public string FacilitatorsNames { get; set; }
        public string FacilitatorsTitles { get; set; }
        public string LabName { get; set; }
        public string LabDay { get; set; }
        public string LabStart { get; set; }
        public string LabEnd { get; set; }
        public string LabHours { get; set; }
        public string LearnerTitleswCount { get; set; }
        public int TotalFacilitators { get; set; }
        public double TotalFacilitatorsHours { get; set; }
    }


}
