using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LabSignup
{
    public partial class Agreement : Form
    {
        public Agreement()
        {
            InitializeComponent();
        }

        private void Agreement_Load(object sender, EventArgs e)
        {
            this.label1.Text = $"STOP";
        }
        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            LabSignup f2 = new LabSignup();
            f2.Show();
        }
    }
}
