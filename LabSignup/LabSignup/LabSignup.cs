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

namespace LabSignup
{
    public partial class LabSignup : Form
    {
        public static List<string> labNames = new List<string>();
        public static List<LabInfo> allLabs = new List<LabInfo>();

        public LabSignup()
        {
            InitializeComponent();
            
        }

        private void LabSignup_Load(object sender, EventArgs e)
        {

            string file = @"D:\LabData.xlsx";

            using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                var sheet = package.Workbook.Worksheets["Sheet1"];
                var labs = new LabSignup().GetList<LabInfo>(sheet);
                allLabs = labs;

                foreach (var lab in allLabs)
                {
                    this.comboBox1.Items.Add(lab.LabName);
                }

            }

        }

        private void button1_Click(object sender, EventArgs e)
        {

           
        }

        private void button2_Click(object sender, EventArgs e)
        {


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
