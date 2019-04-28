using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;

namespace PRECISE

{
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
        }
        // public string conString = "Data Source=DESKTOP-SM32JMN;Initial Catalog=ForTesting;Integrated Security=True";

        public string conString = System.Configuration.ConfigurationManager.ConnectionStrings["connection_string"].ConnectionString;
        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // MessageBox.Show(textBox1.Text);
            //public string conString = "Data Source=DESKTOP-SM32JMN;Initial Catalog=ForTesting;Integrated Security=True";
            SqlConnection con5 = new SqlConnection(conString);
            con5.Open();


            String st = "INSERT INTO dbo.Users(User_id,Name,Password,Permission) values (@User_id,@Name,@Password,@Permission)";
          //  String st = "INSERT INTO dbo.Users(Product_Cost_Category,Precise_Freight_and_Customs,Financing_per_month,TEBYAN_Freight_and_Customs,VIRGIN_percentage) values (@User_id,@Name,@Password,@Permission, @PPP)";


            SqlCommand cmd = new SqlCommand(st, con5);
           cmd.Parameters.AddWithValue("@User_id", textBox1.Text);
            cmd.Parameters.AddWithValue("@Name", textBox2.Text);
           cmd.Parameters.AddWithValue("@Password", textBox3.Text);
            cmd.Parameters.AddWithValue("@Permission", float.Parse(textBox4.Text));
            


            

            // cmd.Parameters.AddWithValue(" @QAR_RRP_QAR", float.Parse(textBox6.Text));
            cmd.ExecuteNonQuery();
            con5.Close();

            this.Hide();
        }
    }
}
