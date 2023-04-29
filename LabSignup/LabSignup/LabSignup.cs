﻿using OfficeOpenXml;
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

            this.panel1.Visible = false;
            
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

        }

        private void button5_Click(object sender, EventArgs e)
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

            var facilitator = new FacilitatorsInfo
            { Names = this.textBox1.Text, Titles = this.comboBox2.Text, LabName = this.comboBox1.Text, LabDay = labDay.Replace("12:00:00 AM", ""), LabStart = labStart, LabEnd = labEnd, LabHours = (DateTime.Parse(labEnd) - DateTime.Parse(labStart)).ToString() };
            currentFacilitator.Add(facilitator);

            ExportFacilitators();

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

                sheet.Cells[1, 1].LoadFromText($"Course Name,{titlesString},Total Learners,Total Learners Hours,Total Facilitators,Total Facilitators Hours");

                int rowStart = sheet.Dimension.Start.Row;
                int rowEnd = sheet.Dimension.End.Row;
                int tLearnersColumn = -1;
                string tLearnersColumnAddress = "";
                string tLearnersRowAddress = "";
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
                var titlesString = string.Join(",", allTitlesStrings);

                sheet.Cells[1, 1].LoadFromText($"Facilitator Names,Facilitator Titles,Course Name,{titlesString},Total Facilitators,Total Facilitators Hours");


                var lastRow = sheet.Dimension.End.Row;

                sheet.Cells[lastRow + 1, 1].LoadFromCollection(currentFacilitator, false);

                package.Save();

                currentFacilitator.Clear();

                //using (ExcelPackage npackage = new ExcelPackage(new FileInfo(FacilitatorsList)))
                //{
                //    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                //    var nsheet = npackage.Workbook.Worksheets["Sheet1"];
                //    var facilitators = new LabSignup().GetList<SigneeInfo>(nsheet);
                //    allFacilitators = facilitators;

                //}
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

                currentSignee.Clear();

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

    }
}
