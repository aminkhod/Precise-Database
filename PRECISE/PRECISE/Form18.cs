using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PRECISE
{
    public partial class Form18 : Form
    {
        public Form18()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "Product ID")
            {
                Form9 frm = new Form9();
                frm.Show();
            }

            else if (comboBox1.Text == "Product Name")
            {
                Form10 frm10 = new Form10();
                frm10.Show();
            }

            if (comboBox1.Text == "Category")
            {
                Form11 frm11 = new Form11();
                frm11.Show();
            }

            if (comboBox1.Text == "Product Cost Category")
            {
                Form12 frm12 = new Form12();
                frm12.Show();
            }
        }
    }
}
