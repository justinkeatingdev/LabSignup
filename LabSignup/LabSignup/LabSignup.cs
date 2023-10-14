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
        public static List<FacilitatorsInfo> allFacilitators = new List<FacilitatorsInfo>();
        public static List<SigneeInfo> currentSignee = new List<SigneeInfo>();
        public static List<FacilitatorsInfo> currentFacilitator = new List<FacilitatorsInfo>();
        public static string execPath = Path.GetDirectoryName(Application.ExecutablePath);
        public static string signeeList = execPath + $"/ExcelFiles/SignInSheet.xlsx";
        public static string FacilitatorsList = execPath + $"/ExcelFiles/FacilitatorsSignInSheet.xlsx";


        public LabSignup()
        {
            InitializeComponent();
            
        }

        private void LabSignup_Load(object sender, EventArgs e)
        {

            //int w = Screen.PrimaryScreen.Bounds.Width;
            //int h = Screen.PrimaryScreen.Bounds.Height;
            //this.Location = new Point(0, 0);
            //this.Size = new Size(w, h);

            this.WindowState = FormWindowState.Maximized;
            this.Hide();
            this.button2.Hide();
            this.panel1.Visible = false;
            this.panel2.Visible = false;

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
                    if (!labNames.Contains($"{labDay}- {lab.LabName}"))
                    {
                        labNames.Add($"{labDay}- {lab.LabName}");
                    }
                    this.comboBox4.Items.Add($"{labDay}- {lab.LabName}");
                }
                this.comboBox1.Text = "Select a Lab";
                this.comboBox4.Text = "Select a Lab";

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
                    this.Title.Items.Add($"{title.Title}");
                    this.dataGridView2.Columns.Add(title.Title, title.Title);
                }
                this.comboBox2.Text = "Select Your Title";


            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == string.Empty)
            {
                MessageBox.Show("Please enter your first name");
                return;
            }
            if (textBox2.Text == string.Empty)
            {
                MessageBox.Show("Please enter your last name");
                return;
            }
            if (comboBox1.Text == "Select a Lab")
            {
                MessageBox.Show("Please select a lab");
                return;
            }
            if (comboBox2.Text == "Select Your Title")
            {
                MessageBox.Show("Please select your title");
                return;
            }

            if (!string.IsNullOrEmpty(this.textBox1.Text))
            {

                string labDay = "";
                string labStart = "";
                string labEnd = "";

                var labData = allLabs.AsQueryable().Where(l => l.LabName.ToLower().Contains(comboBox1.Text.Split('-').LastOrDefault().Trim().ToLower())).FirstOrDefault();
                if (labData != null)
                {
                    labDay = labData.LabDay;
                    labStart = labData.LabStart;
                    labEnd = labData.LabEnd;
                }
                var signee = new SigneeInfo
                { FirstName = this.textBox1.Text, LastName = this.textBox2.Text, Title = this.comboBox2.Text, LabName = this.comboBox1.Text, LabDay = labDay.Replace("12:00:00 AM", ""), LabStart = labStart, LabEnd = labEnd, LabSignInTime = DateTime.Now.ToString(), LabHours = (DateTime.Parse(labEnd) - DateTime.Parse(labStart)).ToString() };
                currentSignee.Add(signee);

                InsertSigneeIntoSheet();
            }

            this.textBox1.Clear();
            this.textBox2.Clear();
            this.comboBox1.Text = "Select a Lab";
            this.comboBox2.Text = "Select Your Title";
            this.panel2.Visible = false;
            Agreement f2 = new Agreement();
            this.Hide();
            f2.Show();

        }

        private void button5_Click(object sender, EventArgs e)
        {

            if (comboBox4.Text == "Select a Lab")
            {
                MessageBox.Show("Please select a lab");
                return;
            }

            if (dataGridView1.Rows.Count - 1 > 0)
            {


                string labDay = "";
                string labStart = "";
                string labEnd = "";

                var labData = allLabs.AsQueryable().Where(l => l.LabName.ToLower().Contains(comboBox4.Text.Split('-').LastOrDefault().Trim().ToLower())).FirstOrDefault();
                if (labData != null)
                {
                    labDay = labData.LabDay;
                    labStart = labData.LabStart;
                    labEnd = labData.LabEnd;
                }

                //right here is needed logic to get datagridview data and form it to be a facilitator entry then need to send that to the facilitator signinsheet
                var facilNames = "";
                var facilTitles = "";

                var facilDataRows = dataGridView1.Rows.Count - 1;
                for (int i = 0; i <= facilDataRows - 1; i++)
                {

                    var ffirstname = dataGridView1.Rows[i].Cells[0].Value == null ? "" : dataGridView1.Rows[i].Cells[0].Value.ToString();
                    var flastname = dataGridView1.Rows[i].Cells[1].Value == null ? "" : " " + dataGridView1.Rows[i].Cells[1].Value.ToString();
                    facilNames += ffirstname + flastname;

                    if (i != facilDataRows - 1)
                    {
                        facilNames += ",";
                    }

                    facilTitles += dataGridView1.Rows[i].Cells[2].Value == null ? "" : dataGridView1.Rows[i].Cells[2].Value.ToString(); ;

                    if(dataGridView1.Rows[i].Cells[2].Value != null)
                    {
                        if (i != facilDataRows - 1)
                        {
                            facilTitles += ",";
                        }
                    }

                    if (ffirstname == String.Empty)
                    {
                        MessageBox.Show("Please enter first name for facilitator " + (i+1));
                        return;
                    }
                    if (flastname == String.Empty)
                    {
                        MessageBox.Show("Please enter last name for facilitator " + (i + 1));
                        return;
                    }
                    if (dataGridView1.Rows[i].Cells[2].Value == null)
                    {
                        MessageBox.Show("Please select a title for facilitator " + (i + 1));
                        return;
                    }

                }

                var facilLearnerColumns = dataGridView2.ColumnCount;
                var allTitleCount = allTitles.Count();
                var learnerTitles = "";

                foreach (DataGridViewColumn item in dataGridView2.Columns)
                {
                    var columnNumber = item.Index;

                    int titleTotalNumber = dataGridView2.Rows[0].Cells[columnNumber].Value == null ? 0 : int.Parse(dataGridView2.Rows[0].Cells[columnNumber].Value.ToString());

                    if (titleTotalNumber != 0)
                    {
                        learnerTitles += item.Name + "=" + titleTotalNumber + ",";

                    }  
                }

                List<string> facNames = facilNames.Split(',').ToList();
                int facCount = facNames.Count();

                if(learnerTitles.Count() > 1)
                {
                    learnerTitles = learnerTitles.Remove(learnerTitles.LastIndexOf(','));
                }

                double facTotalLabHoursCalculated = 0;
                if (labData != null)
                {
                    var facLabHours = DateTime.Parse(labEnd) - DateTime.Parse(labStart);
                    var totalLabMinutes = facLabHours.TotalMinutes;
                    facTotalLabHoursCalculated = facCount * (totalLabMinutes / 60);
                }
                

                var facilitator = new FacilitatorsInfo
                { FacilitatorsNames = facilNames, FacilitatorsTitles = facilTitles, LabName = this.comboBox4.Text, LabDay = labDay.Replace("12:00:00 AM", ""), LabStart = labStart, LabEnd = labEnd, LabHours = (DateTime.Parse(labEnd) - DateTime.Parse(labStart)).ToString(), LearnerTitleswCount = learnerTitles, TotalFacilitators = facCount, TotalFacilitatorsHours = facTotalLabHoursCalculated };
                currentFacilitator.Add(facilitator);

                ExportFacilitators();
            }
            else
            {
                MessageBox.Show("Please fill in first name, last name, and title");
                return;
            }

            this.dataGridView1.Rows.Clear();
            this.dataGridView2.Rows.Clear();
            this.comboBox4.Text = "Select a Lab";
            this.panel1.Visible = false;
            Agreement f2 = new Agreement();
            this.Hide();
            f2.Show();

        }

        private void button2_Click(object sender, EventArgs e)
        {

            //get learner signins from sheet
            using (ExcelPackage npackage = new ExcelPackage(new FileInfo(signeeList)))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                var nsheet = npackage.Workbook.Worksheets["Sheet1"];
                if (nsheet.Cells[1, 1].Count() != 0)
                {
                    var signees = new LabSignup().GetList<SigneeInfo>(nsheet);
                    allSignee = signees;
                }


            }

            //get facilitators from sheet
            using (ExcelPackage npackage = new ExcelPackage(new FileInfo(FacilitatorsList)))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                var nsheet = npackage.Workbook.Worksheets["Sheet1"];
                if (nsheet.Cells[1, 1].Count() !=0)
                {
                    var facilitators = new LabSignup().GetList<FacilitatorsInfo>(nsheet);
                    allFacilitators = facilitators;
                }

            }

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

                sheet.Cells[1, 1].LoadFromText($"Course Name,{titlesString},Total Learners,Total Learners Hours,Total Facilitators,Total Facilitators Hours");

                int rowStart = sheet.Dimension.Start.Row;
                int rowEnd = sheet.Dimension.End.Row;
                int tLearnersColumn = -1;
                string tLearnersColumnAddress = "";
                string tLearnersRowAddress = "";
                int tLearnerHoursColumn = -1;
                string tLearnerHoursColumnAddress = "";
                //int tFacilitatorsColumn = -1;
                //string tFacilitatorsColumnAddress = "";
                //int tFacilitatorsHoursColumn = -1;
                //string tFacilitatorsHoursColumnAddress = "";
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

                foreach (var f in allFacilitators)
                {
                    int cellRow = -1;
                    int cellColumn = -1;

                    foreach (var worksheetCell in sheet.Cells)
                    {
                        if (worksheetCell.Value.ToString() == f.LabName)
                        {
                            cellRow = worksheetCell.EntireRow.StartRow;
                        }
                    }

                    List<string> learnerTitleswcount = f.LearnerTitleswCount.Split(',').ToList();

                    foreach (var titlewcount in learnerTitleswcount)
                    {
                        if (!string.IsNullOrEmpty(titlewcount))
                        {
                            var title = titlewcount.Split('=').FirstOrDefault();
                            int count = int.Parse(titlewcount.Split('=').LastOrDefault());

                            foreach (var worksheetCell in sheet.Cells)
                            {
                                if (worksheetCell.Value.ToString() == title)
                                {
                                    cellColumn = worksheetCell.EntireColumn.StartColumn;
                                }

                            }

                            if (cellRow != -1 && cellColumn != -1)
                            {
                                var cellValue = string.IsNullOrEmpty(sheet.Cells[cellRow, cellColumn].Text) ? 0 : int.Parse(sheet.Cells[cellRow, cellColumn].Text.ToString());
                                cellValue += count;

                                sheet.Cells[cellRow, cellColumn].LoadFromText($"{cellValue}");
                            }
                        }

                    }

                    foreach (var worksheetCell in sheet.Cells)
                    {
                        if (worksheetCell.Value.ToString() == "Total Facilitators")
                        {
                            var tFacilitatorsColumn = worksheetCell.EntireColumn.StartColumn;
                            var tFacilitatorsColumnAddress = ExcelCellAddress.GetColumnLetter(tFacilitatorsColumn);

                            var cellValue = string.IsNullOrEmpty(sheet.Cells[cellRow, tFacilitatorsColumn].Text) ? 0 : int.Parse(sheet.Cells[cellRow, tFacilitatorsColumn].Text.ToString());
                            cellValue += f.TotalFacilitators;

                            sheet.Cells[cellRow, tFacilitatorsColumn].LoadFromText($"{cellValue}");

                        }

                        if (worksheetCell.Value.ToString() == "Total Facilitators Hours")
                        {
                            var tFacilitatorsHoursColumn = worksheetCell.EntireColumn.StartColumn;
                            var tFacilitatorsHoursColumnAddress = ExcelCellAddress.GetColumnLetter(tFacilitatorsHoursColumn);

                            double cellValue = string.IsNullOrEmpty(sheet.Cells[cellRow, tFacilitatorsHoursColumn].Text) ? 0 : double.Parse(sheet.Cells[cellRow, tFacilitatorsHoursColumn].Text.ToString());
                            cellValue += f.TotalFacilitatorsHours;

                            sheet.Cells[cellRow, tFacilitatorsHoursColumn].LoadFromText($"{cellValue}");
                        }
                    }

                }

                foreach (var worksheetCell in sheet.Cells)
                {
                    if (worksheetCell.Value.ToString() == "Total Learners")
                    {
                        tLearnersColumn = worksheetCell.EntireColumn.StartColumn;
                        tLearnersColumnAddress = ExcelCellAddress.GetColumnLetter(tLearnersColumn);
                        tLearnersRowAddress = ExcelCellAddress.GetColumnLetter(tLearnersColumn - 1);
                    }

                    if (worksheetCell.Value.ToString() == allTitlesStrings.First().ToString())
                    {
                        var firstTitleColumn = worksheetCell.EntireColumn.StartColumn;
                        firstTitleAddress = ExcelCellAddress.GetColumnLetter(firstTitleColumn);
                    }

                    if (worksheetCell.Value.ToString() == "Total Learners Hours")
                    {
                        tLearnerHoursColumn = worksheetCell.EntireColumn.StartColumn;
                        tLearnerHoursColumnAddress = ExcelCellAddress.GetColumnLetter(tLearnerHoursColumn);
                    }

                }

                for (int i = 0; i < labNames.Count(); i++)
                {
                    var labHours = DateTime.Parse(allLabs[i].LabEnd) - DateTime.Parse(allLabs[i].LabStart);
                    var totalLabMinutes = labHours.TotalMinutes;

                    sheet.Cells[i+2, tLearnersColumn].Formula = $"=SUM({firstTitleAddress}{i+2}:{tLearnersRowAddress}{i+2})";
                    sheet.Cells[i+2, tLearnerHoursColumn].Formula = $"={tLearnersColumnAddress}{i + 2}*{totalLabMinutes}/60";
                }

                pck.SaveAs(new FileInfo(file));
            }
        }

        private void ExportFacilitators()
        {

            using (ExcelPackage package = new ExcelPackage(new FileInfo(FacilitatorsList)))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                var sheet = package.Workbook.Worksheets["Sheet1"];

                sheet.Cells[1, 1].LoadFromText($"FacilitatorsNames,FacilitatorsTitles,LabName,LabDay,LabStart,LabEnd,LabHours,LearnerTitleswCount,TotalFacilitators,TotalFacilitatorsHours");

                var lastRow = sheet.Dimension.End.Row;

                sheet.Cells[lastRow + 1, 1].LoadFromCollection(currentFacilitator, false);

                package.Save();

                currentFacilitator.Clear();

            }
        }

        private void InsertSigneeIntoSheet()
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(signeeList)))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                var sheet = package.Workbook.Worksheets["Sheet1"];

                sheet.Cells[1, 1].LoadFromText($"FirstName,LastName,Title,LabName,LabDay,LabStart,LabEnd,LabSignInTime,LabHours");


                var lastRow = sheet.Dimension.End.Row;

                sheet.Cells[lastRow+1, 1].LoadFromCollection(currentSignee, false);

                package.Save();

                currentSignee.Clear();

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

        private void button3_Click(object sender, EventArgs e)
        {
            this.panel1.Visible = true;
            this.panel2.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.panel2.Visible = true;
            this.panel1.Visible = false;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (this.textBox3.Text == "CleVAsim1")
            {
                this.button2.Show();
            }
            else
            {
                this.button2.Hide();
            }
        }
    }
}
