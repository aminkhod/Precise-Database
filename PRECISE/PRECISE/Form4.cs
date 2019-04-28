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
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }
        // public string conString = "Data Source=DESKTOP-SM32JMN;Initial Catalog=ForTesting;Integrated Security=True";
        public string conString = System.Configuration.ConfigurationManager.ConnectionStrings["connection_string"].ConnectionString;

        private void button1_Click(object sender, EventArgs e)
        {
           // MessageBox.Show(textBox1.Text);
            //public string conString = "Data Source=DESKTOP-SM32JMN;Initial Catalog=ForTesting;Integrated Security=True";
            SqlConnection con1 = new SqlConnection(conString);
            con1.Open();
            String st = "INSERT INTO dbo.Products_New(productsID,SubGroup_Name,Group_Name,product,version,Unit_cost_USD,Unit_cost_AED,Freight_customs,Financing,PRECISE_Landed_Cost,PRECISE_Margin_AED,PRECISE_Margin_Cost_Ratio,T_Cost_Price_USD,T_Cost_Price_QAR,T_Freight_customs_10per_add_5per,T_Landed_Cost_QAR,T_Margin_QAR,T_Margin_Per,V_Retailer_Cost_Price_QAR,V_Retailer_Margin_QAR,V_Retailer_Margin_per,UAE_RRP_AED,QAR_RRP_AED,QAR_RRP_QAR) values (@productsID,@SubGroup_Name,@Group_Name,  @product,@version,@Unit_cost_USD,@Unit_cost_AED,@Freight_customs,@Financing,@PRECISE_Landed_Cost,@PRECISE_Margin_AED,@PRECISE_Margin_Cost_Ratio,@T_Cost_Price_USD,@T_Cost_Price_QAR,@T_Freight_customs_10per_add_5per,@T_Landed_Cost_QAR,@T_Margin_QAR,@T_Margin_Per,@V_Retailer_Cost_Price_QAR,@V_Retailer_Margin_QAR,@V_Retailer_Margin_per,@UAE_RRP_AED,@QAR_RRP_AED,@QAR_RRP_QAR)";
            SqlCommand cmd = new SqlCommand(st, con1);
            cmd.Parameters.AddWithValue("@productsID", textBox1.Text);
            cmd.Parameters.AddWithValue("@SubGroup_Name", textBox7.Text);

            cmd.Parameters.AddWithValue("@Group_Name", textBox2.Text);
            cmd.Parameters.AddWithValue("@product", textBox3.Text);
            cmd.Parameters.AddWithValue("@Unit_cost_USD", float.Parse(textBox4.Text));
            cmd.Parameters.AddWithValue("@UAE_RRP_AED", float.Parse(textBox5.Text));
            // cmd.Parameters.AddWithValue(" @QAR_RRP_QAR", float.Parse(textBox6.Text));
            cmd.Parameters.AddWithValue("@QAR_RRP_QAR", float.Parse(textBox6.Text));
            cmd.Parameters.AddWithValue("@version", textBox8.Text);
            //**************************C**************
            // double USD_Exchange_Rate_to_AED = 3.68;

            string s1 = "Select  * From dbo.Margin ";
            SqlCommand cmd9 = new SqlCommand(s1, con1);

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = cmd9;
            DataTable dbdataset = new DataTable();
            sda.Fill(dbdataset);


            string s10= (dbdataset.Rows[0][8]).ToString();
            double USD_Exchange_Rate_to_AED = float.Parse(s10);

            double Unit_cost_AED = (USD_Exchange_Rate_to_AED) * (float.Parse(textBox4.Text));

            cmd.Parameters.AddWithValue("@Unit_cost_AED", Unit_cost_AED);
            //***********************E**************

            


            string s2 = (dbdataset.Rows[0][1]).ToString();
            double Percent = float.Parse(s2);


          // MessageBox.Show(s2);

            // double Percent = 0;



            double Freight_customs = (float.Parse(textBox4.Text)) * (Percent/100);

            cmd.Parameters.AddWithValue("@Freight_customs", Freight_customs);
            //*************F***********************
           // double Percent_f = 0;
            string s3 = (dbdataset.Rows[0][2]).ToString();
            double Percent_f = float.Parse(s3);


            double Financing = (float.Parse(textBox4.Text)) *( Percent_f/100);

            cmd.Parameters.AddWithValue("@Financing", Financing);
            //*************G***********************


            double PRECISE_Landed_Cost = (Freight_customs) + (float.Parse(textBox4.Text)) + (Financing);

            cmd.Parameters.AddWithValue("@PRECISE_Landed_Cost", PRECISE_Landed_Cost);



            //**********************K***********************
            //double K11 = 0.33;
            string s4 = (dbdataset.Rows[0][3]).ToString();
            double K11 = float.Parse(s4);

            double T_Cost_Price_USD = ((PRECISE_Landed_Cost) * (K11/100)) + (PRECISE_Landed_Cost);


            double ttttt = Convert.ToDouble(String.Format("{0:0.00}", T_Cost_Price_USD));



            cmd.Parameters.AddWithValue("@T_Cost_Price_USD", ttttt);
            //****************H*******************

            double PRECISE_Margin_AED = (T_Cost_Price_USD) - (PRECISE_Landed_Cost);

            double ts = Convert.ToDouble(String.Format("{0:0.00}", PRECISE_Margin_AED));


            cmd.Parameters.AddWithValue("@PRECISE_Margin_AED", ts);
            //*******************I*****************
            double PRECISE_Margin_Cost_Ratio = (PRECISE_Margin_AED) / (float.Parse(textBox4.Text));

            double sss = Convert.ToDouble(String.Format("{0:0.00}", PRECISE_Margin_Cost_Ratio));

            cmd.Parameters.AddWithValue("@PRECISE_Margin_Cost_Ratio", sss);


            //********************L***************************
           // double W9 = 3.64;

            string s5 = (dbdataset.Rows[0][6]).ToString();
            double W9 = float.Parse(s5);

            double T_Cost_Price_QAR = (T_Cost_Price_USD) * (W9);

            double tttt = Convert.ToDouble(String.Format("{0:0.00}", T_Cost_Price_QAR));

            cmd.Parameters.AddWithValue("@T_Cost_Price_QAR", tttt);

            //******************M************************
           // double M11 = 0.15;
            string s6 = (dbdataset.Rows[0][4]).ToString();
            double M11 = float.Parse(s6);

            double T_Freight_customs_10per_add_5per = (T_Cost_Price_QAR) * (M11/100);


            double ttt = Convert.ToDouble(String.Format("{0:0.00}", T_Freight_customs_10per_add_5per));

            cmd.Parameters.AddWithValue("@T_Freight_customs_10per_add_5per", ttt);
            //********************N*************************

            double T_Landed_Cost_QAR = (T_Cost_Price_QAR) + (T_Freight_customs_10per_add_5per);
            double tt = Convert.ToDouble(String.Format("{0:0.00}", T_Landed_Cost_QAR));


            cmd.Parameters.AddWithValue("@T_Landed_Cost_QAR", tt);



            //***************R********************
          //  double R11 = 0.75;
            string s7 = (dbdataset.Rows[0][5]).ToString();
            double R11 = float.Parse(s7);

            double V_Retailer_Cost_Price_QAR = (float.Parse(textBox6.Text)) * (R11/100);

            cmd.Parameters.AddWithValue("@V_Retailer_Cost_Price_QAR", V_Retailer_Cost_Price_QAR);
            //*******************O******************
            double T_Margin_QAR = (V_Retailer_Cost_Price_QAR) - (T_Landed_Cost_QAR);

            double t = Convert.ToDouble(String.Format("{0:0.00}", T_Margin_QAR));


            cmd.Parameters.AddWithValue("@T_Margin_QAR", t);
            //*********************P*****************

            double T_Margin_Per = (T_Margin_QAR) / (V_Retailer_Cost_Price_QAR);


            double ppp = Convert.ToDouble(String.Format("{0:0.00}", T_Margin_Per));

            cmd.Parameters.AddWithValue("@T_Margin_Per", ppp);

            //******************S*********************

            double V_Retailer_Margin_QAR = (float.Parse(textBox6.Text)) - (V_Retailer_Cost_Price_QAR);

            cmd.Parameters.AddWithValue("@V_Retailer_Margin_QAR", V_Retailer_Margin_QAR);
            //***********************T*******************

            double V_Retailer_Margin_per = (V_Retailer_Margin_QAR) / (float.Parse(textBox6.Text));

            double pppp = Convert.ToDouble(String.Format("{0:0.00}", V_Retailer_Margin_per));

            cmd.Parameters.AddWithValue("@V_Retailer_Margin_per", pppp);
            //**************************W*************************
          //  double X9 = 1.01;

            string s8 = (dbdataset.Rows[0][7]).ToString();
            double X9 = float.Parse(s8);

            double QAR_RRP_AED = (float.Parse(textBox6.Text)) * (X9);

            cmd.Parameters.AddWithValue("@QAR_RRP_AED", QAR_RRP_AED);



            cmd.ExecuteNonQuery();
            con1.Close();

            this.Hide();
        }
    }
}
