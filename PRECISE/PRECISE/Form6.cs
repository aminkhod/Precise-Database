using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;

namespace PRECISE
{
    public partial class Form6 : Form
    {
        public Form6()
        {
            InitializeComponent();
        }
        // public string conString = "Data Source=DESKTOP-SM32JMN;Initial Catalog=ForTesting;Integrated Security=True";
        public string conString = System.Configuration.ConfigurationManager.ConnectionStrings["connection_string"].ConnectionString;

        private void button1_Click(object sender, EventArgs e)
        {



        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form16 frm16 = new Form16();
            frm16.Show();


            this.Hide();
        }

        
           

        private void button3_Click_1(object sender, EventArgs e)
        {
                Form18 frm18 = new Form18();
                frm18.Show();


            this.Hide();
        }

        }
    
}
