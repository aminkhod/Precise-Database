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
    public partial class Form11 : Form
    {
        public Form11()
        {
            InitializeComponent();
            FillCombo();
        }
        // public string conString = "Data Source=DESKTOP-SM32JMN;Initial Catalog=ForTesting;Integrated Security=True";
        public string conString = System.Configuration.ConfigurationManager.ConnectionStrings["connection_string"].ConnectionString;
        

        void FillCombo()
        {
            SqlConnection con = new SqlConnection(conString);
            con.Open();
            String q = "Select *From dbo.Products_New ";

            SqlCommand cmd = new SqlCommand(q, con);
            SqlDataReader myreader = cmd.ExecuteReader();
            var lll = new List<string>();
            var lll_new = new List<string>();

            while (myreader.Read())
            {

                string u = myreader["Group_Name"].ToString();
                lll.Add(u);

           

            }

            lll_new.AddRange(lll.Distinct());

            foreach (string v in lll_new) {
                comboBox1.Items.Add(v);
            }
            




        }
        private void button1_Click(object sender, EventArgs e)
        {


            SqlConnection con11 = new SqlConnection(conString);
            con11.Open();
            string s1 = "Select *From dbo.Products_New where Group_Name = '" + comboBox1.Text.ToString() + "'";

            SqlCommand cmd9 = new SqlCommand(s1, con11);
            //SqlDataReader myreader = cmd9.ExecuteReader();
            string u = "";
            string[] aaa = new string[6];
            int i = 0;

            try
            {
                SqlDataAdapter sda = new SqlDataAdapter();
                sda.SelectCommand = cmd9;
                DataTable dbdataset = new DataTable();
                sda.Fill(dbdataset);
                BindingSource bSource = new BindingSource();
                bSource.DataSource = dbdataset;
                dataGridView1.DataSource = bSource;
                sda.Update(dbdataset);


                // change column name:
                dataGridView1.Columns["productsID"].HeaderText = "     ID";
                dataGridView1.Columns["SubGroup_Name"].HeaderText = "   Product \n    Cost \n  Category";
                dataGridView1.Columns["Group_Name"].HeaderText = "                   Category";
                dataGridView1.Columns["product"].HeaderText = "        Product";
                dataGridView1.Columns["Unit_cost_USD"].HeaderText = "Unit_cost \n  (USD)";
                dataGridView1.Columns["Unit_cost_AED"].HeaderText = "Unit_cost \n  (AED)";
                dataGridView1.Columns["Freight_customs"].HeaderText = "Freight& \n  Customs  (10%+5%)";
                dataGridView1.Columns["Financing"].HeaderText = " Financing \n 1%/Month";
                dataGridView1.Columns["PRECISE_Landed_Cost"].HeaderText = "PRECISE \n Landed \n  Cost";
                dataGridView1.Columns["PRECISE_Margin_AED"].HeaderText = "PRECISE \n  Margin \n  (AED)";
                dataGridView1.Columns["PRECISE_Margin_Cost_Ratio"].HeaderText = "PRECISE  Margin \n Cost Ratio (%)";
                dataGridView1.Columns["T_Cost_Price_USD"].HeaderText = "TEBYAN \n  Cost\n   Price \n  (USD)";
                dataGridView1.Columns["T_Cost_Price_QAR"].HeaderText = "TEBYAN \n  Cost \n  Price \n (QAR)";
                dataGridView1.Columns["T_Freight_customs_10per_add_5per"].HeaderText = "TEBYAN \n Freight& \n  Customs (10%+5%)";
                dataGridView1.Columns["T_Landed_Cost_QAR"].HeaderText = "TEBYAN \n Landed \n   Cost \n   (QAR)";
                dataGridView1.Columns["T_Margin_QAR"].HeaderText = "TEBYAN \n Margin \n (QAR) \n    ";
                dataGridView1.Columns["T_Margin_Per"].HeaderText = "TEBYAN \n Margin \n   (%) \n    ";

                dataGridView1.Columns["V_Retailer_Cost_Price_QAR"].HeaderText = "VIRGIN \n Retailer \n    Cost \n  Price*  (QAR)";
                dataGridView1.Columns["V_Retailer_Margin_QAR"].HeaderText = "VIRGIN \n Retailer \n Margin\n    (QAR) \n   ";
                dataGridView1.Columns["V_Retailer_Margin_per"].HeaderText = "VIRGIN \nRetailer \n Margin \n   \n   ";
                dataGridView1.Columns["UAE_RRP_AED"].HeaderText = "UAE RRP \n (AED)";
                dataGridView1.Columns["QAR_RRP_AED"].HeaderText = "QAR RRP \n (AED)";
                dataGridView1.Columns["QAR_RRP_QAR"].HeaderText = "QAR RRP \n (QAR)";


                DataGridViewColumn column = dataGridView1.Columns[0];
                column.Width = 60;

                DataGridViewColumn column1 = dataGridView1.Columns[1];
                column1.Width = 90;

                DataGridViewColumn column2 = dataGridView1.Columns[2];
                column2.Width = 200;

                // DataGridViewColumn column3 = dataGridView1.Columns[3];
                // column.Width = 60;

                DataGridViewColumn column4 = dataGridView1.Columns[4];
                column4.Width = 60;

                DataGridViewColumn column5 = dataGridView1.Columns[5];
                column5.Width = 60;

                DataGridViewColumn column6 = dataGridView1.Columns[6];
                column6.Width = 60;

                DataGridViewColumn column7 = dataGridView1.Columns[7];
                column7.Width = 60;
                DataGridViewColumn column8 = dataGridView1.Columns[8];
                column8.Width = 70;
                DataGridViewColumn column9 = dataGridView1.Columns[9];
                column9.Width = 60;
                DataGridViewColumn column10 = dataGridView1.Columns[10];
                column10.Width = 60;
                DataGridViewColumn column11 = dataGridView1.Columns[11];
                column11.Width = 60;
                DataGridViewColumn column12 = dataGridView1.Columns[12];
                column12.Width = 60;
                DataGridViewColumn column13 = dataGridView1.Columns[13];
                column13.Width = 60;
                DataGridViewColumn column14 = dataGridView1.Columns[14];
                column14.Width = 60;
                DataGridViewColumn column15 = dataGridView1.Columns[15];
                column15.Width = 60;
                DataGridViewColumn column16 = dataGridView1.Columns[16];
                column16.Width = 60;
                DataGridViewColumn column17 = dataGridView1.Columns[17];
                column17.Width = 60;
                DataGridViewColumn column18 = dataGridView1.Columns[18];
                column18.Width = 60;
                DataGridViewColumn column19 = dataGridView1.Columns[19];
                column19.Width = 60;
                DataGridViewColumn column20 = dataGridView1.Columns[20];
                column20.Width = 60;


                DataGridViewColumn column21 = dataGridView1.Columns[21];
                column21.Width = 60;
                DataGridViewColumn column22 = dataGridView1.Columns[22];
                column22.Width = 60;
                DataGridViewColumn column23 = dataGridView1.Columns[23];
                column23.Width = 60;






            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }
}
