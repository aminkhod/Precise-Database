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
    public partial class Form7 : Form
    {
        public Form7()
        {
            InitializeComponent();
        }
        //   public string conString = "Data Source=DESKTOP-SM32JMN;Initial Catalog=ForTesting;Integrated Security=True";
        public string conString = System.Configuration.ConfigurationManager.ConnectionStrings["connection_string"].ConnectionString;

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection con7 = new SqlConnection(conString);
            SqlCommand cmd7 = new SqlCommand("Select * from dbo.Users;", con7);
            try
            {
                SqlDataAdapter sda = new SqlDataAdapter();
                sda.SelectCommand = cmd7;
                DataTable dbdataset = new DataTable();
                sda.Fill(dbdataset);
                BindingSource bSource = new BindingSource();
                bSource.DataSource = dbdataset;
                dataGridView1.DataSource = bSource;
                sda.Update(dbdataset);

               // change column name in dataGride
                dataGridView1.Columns["User_id"].HeaderText = "ID";


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
