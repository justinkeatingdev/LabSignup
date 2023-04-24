using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;

namespace LabSignup
{
    public partial class LabSignup : Form
    {
        public static List<string> labNames = new List<string>();
        public static List<LabInfo> allLabs = new List<LabInfo>();
        public static List<SigneeTitles> allTitles = new List<SigneeTitles>();
        public static List<string> allTitlesStrings = new List<string>();
        public static List<SigneeInfo> allSignee = new List<SigneeInfo>();
        public static List<SigneeInfo> currentSignee = new List<SigneeInfo>();
        public static string execPath = Path.GetDirectoryName(Application.ExecutablePath);
        public static string signeeList = execPath + $"/ExcelFiles/SignInSheet.xlsx";


        public LabSignup()
        {
            InitializeComponent();
            
        }

        private void LabSignup_Load(object sender, EventArgs e)
        {
            
            string file = execPath + "/ExcelFiles/LabDetails.xlsx";

            using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                var sheet = package.Workbook.Worksheets["Sheet1"];
                var labs = new LabSignup().GetList<LabInfo>(sheet);
                allLabs = labs;

                foreach (var lab in allLabs)
                {
                    string labDay = lab.LabDay.Replace("12:00:00 AM", "");
                    this.comboBox1.Items.Add($"{labDay}- {lab.LabName}");
                    labNames.Add($"{labDay}- {lab.LabName}");
                }
                this.comboBox1.SelectedIndex = -0;

            }

            string signeeTitles = execPath + "/ExcelFiles/Titles.xlsx";

            using (ExcelPackage package = new ExcelPackage(new FileInfo(signeeTitles)))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                var sheet = package.Workbook.Worksheets["Sheet1"];
                var titles = new LabSignup().GetList<SigneeTitles>(sheet);
                allTitles = titles;

                foreach (var title in allTitles.OrderBy(x=> x.Title))
                {
                    this.comboBox2.Items.Add($"{title.Title}");
                    allTitlesStrings.Add(title.Title);
                }
                this.comboBox2.SelectedIndex = -0;

            }

            using (ExcelPackage npackage = new ExcelPackage(new FileInfo(signeeList)))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                var nsheet = npackage.Workbook.Worksheets["Sheet1"];
                var signees = new LabSignup().GetList<SigneeInfo>(nsheet);
                allSignee = signees;

            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string labDay = "";
            string labStart = "";
            string labEnd = "";

            var labData = allLabs.AsQueryable().Where(l => l.LabName.ToLower().Contains(comboBox1.Text.Split('-').LastOrDefault().Trim().ToLower())).FirstOrDefault();
            if(labData != null)
            {
                labDay = labData.LabDay;
                labStart = labData.LabStart;
                labEnd = labData.LabEnd;
            }
            var signee = new SigneeInfo
            { FirstName = this.textBox1.Text, LastName=this.textBox2.Text, Title=this.comboBox2.Text, LabName=this.comboBox1.Text, LabDay= labDay.Replace("12:00:00 AM", ""), LabStart= labStart, LabEnd = labEnd, LabSignInTime = DateTime.Now.ToString(), LabHours = (DateTime.Parse(labEnd) - DateTime.Parse(labStart)).ToString() };
            currentSignee.Add(signee);

            InsertSigneeIntoSheet();

            this.textBox1.Clear();
            this.textBox2.Clear();
            this.comboBox1.SelectedIndex = -0;
            this.comboBox2.SelectedIndex = -0;

        }

        private void button2_Click(object sender, EventArgs e)
        {

            string newExcelFile = execPath + $"/ExcelFiles/LabsData-{DateTime.Now.Month + "-" +  DateTime.Now.Day + "-" + DateTime.Now.Year}.xlsx";
            newExcelFile = newExcelFile.Replace(" ", "-");
            new LabSignup().Export(newExcelFile);
        }

        private void Export(string file)
        {

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            var titlesString = string.Join(",", allTitlesStrings);

            using (ExcelPackage pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("LabsData");

                sheet.Cells[1, 1].LoadFromText($"Course Name,{titlesString},Total Learners");

                int rowStart = sheet.Dimension.Start.Row;
                int rowEnd = sheet.Dimension.End.Row;
                int tLearnerHoursColumn = -1;
                string tLearnerHoursColumnAddress = "";
                string firstTitleAddress = "";

                string cellRange = rowStart.ToString() + ":" + rowEnd.ToString();

                for (int i = 0; i < labNames.Count(); i++)
                {
                    sheet.Cells[i+2, 1].LoadFromText($"{labNames[i]}");
                }

                foreach(var s in allSignee)
                {
                    int cellRow = -1;
                    int cellColumn = -1;

                    foreach (var worksheetCell in sheet.Cells)
                    {
                        if (worksheetCell.Value.ToString() == s.LabName)
                        {
                            cellRow = worksheetCell.EntireRow.StartRow;
                        }
                        if (worksheetCell.Value.ToString() == s.Title)
                        {
                            cellColumn = worksheetCell.EntireColumn.StartColumn;
                        }
                    }

                    if(cellRow != -1 && cellColumn != -1)
                    {
                        var cellValue = string.IsNullOrEmpty(sheet.Cells[cellRow, cellColumn].Text) ? 0 : int.Parse(sheet.Cells[cellRow, cellColumn].Text.ToString());
                        cellValue += 1;

                        sheet.Cells[cellRow, cellColumn].LoadFromText($"{cellValue}");
                    }
                    
                }

                foreach (var worksheetCell in sheet.Cells)
                {
                    if (worksheetCell.Value.ToString() == "Total Learners")
                    {
                        tLearnerHoursColumn = worksheetCell.EntireColumn.StartColumn;
                        tLearnerHoursColumnAddress = ExcelCellAddress.GetColumnLetter(tLearnerHoursColumn - 1);
                    }

                    if (worksheetCell.Value.ToString() == allTitlesStrings.First().ToString())
                    {
                        var firstTitleColumn = worksheetCell.EntireColumn.StartColumn;
                        firstTitleAddress = ExcelCellAddress.GetColumnLetter(firstTitleColumn);
                    }
                }

                for (int i = 0; i < labNames.Count(); i++)
                {
                    sheet.Cells[i+2, tLearnerHoursColumn].Formula = $"=SUM({firstTitleAddress}{i+2}:{tLearnerHoursColumnAddress}{i+2})";
                }

                pck.SaveAs(new FileInfo(file));
            }
        }

        private void InsertSigneeIntoSheet()
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(signeeList)))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                var sheet = package.Workbook.Worksheets["Sheet1"];

                var lastRow = sheet.Dimension.End.Row;

                sheet.Cells[lastRow+1, 1].LoadFromCollection(currentSignee, false);

                package.Save();

                using (ExcelPackage npackage = new ExcelPackage(new FileInfo(signeeList)))
                {
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    var nsheet = npackage.Workbook.Worksheets["Sheet1"];
                    var signees = new LabSignup().GetList<SigneeInfo>(nsheet);
                    allSignee = signees;

                }
            }
        }

        private List<T> GetList<T>(ExcelWorksheet sheet)
        {
            List<T> list = new List<T>();
            var columnInfo = Enumerable.Range(1, sheet.Dimension.Columns).ToList().Select(n =>

                new { Index = n, ColumnName = sheet.Cells[1, n].Value.ToString() }
            );

            for(int row=2; row <= sheet.Dimension.Rows; row++)
            {
                T obj = (T)Activator.CreateInstance(typeof(T));
                foreach(var prop in typeof(T).GetProperties())
                {
                    int col = columnInfo.SingleOrDefault(c => c.ColumnName == prop.Name).Index;
                    var val = sheet.Cells[row, col].Value;
                    var propType = prop.PropertyType;
                    prop.SetValue(obj, Convert.ChangeType(val, propType));
                }
                list.Add(obj);
            }

            return list;
        }
    }
}
