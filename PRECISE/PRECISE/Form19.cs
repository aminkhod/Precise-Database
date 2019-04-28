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
using System.Data.SqlClient;
using System.IO;


using Excel = Microsoft.Office.Interop.Excel;
using ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat;


namespace PRECISE
{
    public partial class Form19 : Form
    {
        public Form19()
        {
            InitializeComponent();
        }
        public string conString = System.Configuration.ConfigurationManager.ConnectionStrings["connection_string"].ConnectionString;




        public int NumberOfClick = 1;




        private void button1_Click(object sender, EventArgs e)

        {

            NumberOfClick = NumberOfClick + 1;
            {
                SqlConnection con7 = new SqlConnection(conString);
                SqlCommand cmd7 = new SqlCommand("Select * from dbo.Products_New", con7);
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

                    // change column name in dataGride:
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


                SqlCommand cmd11 = new SqlCommand("Select * from dbo.Margin", con7);
                try
                {
                    SqlDataAdapter sda1 = new SqlDataAdapter();
                    sda1.SelectCommand = cmd11;
                    DataTable dbdataset1 = new DataTable();
                    sda1.Fill(dbdataset1);
                    BindingSource bSource = new BindingSource();
                    bSource.DataSource = dbdataset1;
                    dataGridView2.DataSource = bSource;
                    sda1.Update(dbdataset1);


                    dataGridView2.Columns["Product_Cost_Category"].HeaderText = "   Product Cost\n     Category";

                    dataGridView2.Columns["Precise_Freight_and_Customs"].HeaderText = "     Freight& \n      Customs";

                    dataGridView2.Columns["Financing_per_month"].HeaderText = "       Financing";


                    //dataGridView2.Columns["k11"].HeaderText = "      TEBYAN \n       Percentage ";
                    dataGridView2.Columns["k11"].HeaderText = "";

                    dataGridView2.Columns["TEBYAN_Freight_and_Customs"].HeaderText = "      TEBYAN \n       Freight& \n        Customs   ";

                    // dataGridView2.Columns["VIRGIN_percentage"].HeaderText = "        VIRGIN \n        percentage";
                    dataGridView2.Columns["VIRGIN_percentage"].HeaderText = "";
                    dataGridView2.Columns["w9"].HeaderText = "         USD \n Exchange Rate           to \n          QAR";
                    dataGridView2.Columns["x9"].HeaderText = "         QAR \n Exchange Rate           to \n         AED";
                    dataGridView2.Columns["v9"].HeaderText = "         USD \n Exchange Rate           to \n         AED";



                    DataGridViewColumn column0 = dataGridView2.Columns[0];
                    column0.Width = 130;


                    DataGridViewColumn column1 = dataGridView2.Columns[1];
                    column1.Width = 110;

                    DataGridViewColumn column2 = dataGridView2.Columns[2];
                    column2.Width = 110;

                    DataGridViewColumn column3 = dataGridView2.Columns[3];
                    column3.Width = 1;

                    DataGridViewColumn column4 = dataGridView2.Columns[4];
                    column4.Width = 110;

                    DataGridViewColumn column5 = dataGridView2.Columns[5];
                    column5.Width = 1;

                    DataGridViewColumn column6 = dataGridView2.Columns[6];
                    column6.Width = 110;
                    DataGridViewColumn column7 = dataGridView2.Columns[7];
                    column7.Width = 110;
                    DataGridViewColumn column8 = dataGridView2.Columns[8];
                    column8.Width = 110;







                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                // con7.Close();
            }


        }

        private void button2_Click(object sender, EventArgs e)
        {
            SqlConnection con1 = new SqlConnection(conString);
            con1.Open();
            int j = dataGridView1.Rows.Count;
            int[] arr = Enumerable.Range(0, j - 1).ToArray();
            foreach (var i in arr)
            {    //textBox4:
                dataGridView1.Rows[i].Cells["Unit_cost_USD"].Value = float.Parse(dataGridView1.Rows[i].Cells["Unit_cost_USD"].Value.ToString());
                // textBox5:
                dataGridView1.Rows[i].Cells["UAE_RRP_AED"].Value = float.Parse(dataGridView1.Rows[i].Cells["UAE_RRP_AED"].Value.ToString());
                /// textBox6:
                dataGridView1.Rows[i].Cells["QAR_RRP_QAR"].Value = float.Parse(dataGridView1.Rows[i].Cells["QAR_RRP_QAR"].Value.ToString());






                //**************************C**************
                //double USD_Exchange_Rate_to_AED = 3.68;

                string s1 = "Select  * From dbo.Margin ";
                SqlCommand cmd9 = new SqlCommand(s1, con1);

                SqlDataAdapter sda = new SqlDataAdapter();
                sda.SelectCommand = cmd9;
                DataTable dbdataset = new DataTable();
                sda.Fill(dbdataset);

                string s10 = (dbdataset.Rows[0][8]).ToString();
                // double USD_Exchange_Rate_to_AED = float.Parse(s10);

                double USD_Exchange_Rate_to_AED = (float.Parse(dataGridView2.Rows[0].Cells["v9"].Value.ToString()));

                double Unit_cost_AED = (USD_Exchange_Rate_to_AED) * (float.Parse(dataGridView1.Rows[i].Cells["Unit_cost_USD"].Value.ToString()));
                dataGridView1.Rows[i].Cells["Unit_cost_AED"].Value = Unit_cost_AED;

                //***********************E**************

                // double Percent = 0;



                string s2 = (dbdataset.Rows[0][1]).ToString();
                // double Percent = float.Parse(s2);

                double Percent = (int.Parse(dataGridView2.Rows[0].Cells["Precise_Freight_and_Customs"].Value.ToString()));


                double Freight_customs = (float.Parse(dataGridView1.Rows[i].Cells["Unit_cost_USD"].Value.ToString())) * (Percent / 100);
                dataGridView1.Rows[i].Cells["Freight_customs"].Value = Freight_customs;

                //*************F***********************
                // double Percent_f = 0;
                string s3 = (dbdataset.Rows[0][2]).ToString();
                // double Percent_f = float.Parse(s3);
                double Percent_f = (int.Parse(dataGridView2.Rows[0].Cells["Financing_per_month"].Value.ToString()));



                double Financing = (float.Parse(dataGridView1.Rows[i].Cells["Unit_cost_USD"].Value.ToString())) * (Percent_f / 100);
                dataGridView1.Rows[i].Cells["Financing"].Value = Financing;







                //*************G***********************


                double PRECISE_Landed_Cost = (Freight_customs) + (float.Parse(dataGridView1.Rows[i].Cells["Unit_cost_USD"].Value.ToString())) + (Financing);
                dataGridView1.Rows[i].Cells["PRECISE_Landed_Cost"].Value = PRECISE_Landed_Cost;



                //**********************K***********************
                // double K11 = 0.33;
                string s4 = (dbdataset.Rows[0][3]).ToString();
                //double K11 = float.Parse(s4);
                double K11 = (int.Parse(dataGridView2.Rows[0].Cells["K11"].Value.ToString()));



                double T_Cost_Price_USD = ((PRECISE_Landed_Cost) * (K11 / 100)) + (PRECISE_Landed_Cost);
                double ttttt = Convert.ToDouble(String.Format("{0:0.00}", T_Cost_Price_USD));



                dataGridView1.Rows[i].Cells["T_Cost_Price_USD"].Value = ttttt;


                //****************H*******************

                double PRECISE_Margin_AED = (T_Cost_Price_USD) - (PRECISE_Landed_Cost);


                double ts = Convert.ToDouble(String.Format("{0:0.00}", PRECISE_Margin_AED));

                dataGridView1.Rows[i].Cells["PRECISE_Margin_AED"].Value = ts;


                //*******************I*****************
                double PRECISE_Margin_Cost_Ratio = (PRECISE_Margin_AED) / (float.Parse(dataGridView1.Rows[i].Cells["Unit_cost_USD"].Value.ToString()));

                double sss = Convert.ToDouble(String.Format("{0:0.00}", PRECISE_Margin_Cost_Ratio));



                dataGridView1.Rows[i].Cells["PRECISE_Margin_Cost_Ratio"].Value = sss;



                //********************L***************************
                //double W9 = 3.64;
                string s5 = (dbdataset.Rows[0][6]).ToString();
                //  double W9 = float.Parse(s5);

                double W9 = (float.Parse(dataGridView2.Rows[0].Cells["W9"].Value.ToString()));


                double T_Cost_Price_QAR = (T_Cost_Price_USD) * (W9);

                double tttt = Convert.ToDouble(String.Format("{0:0.00}", T_Cost_Price_QAR));

                dataGridView1.Rows[i].Cells["T_Cost_Price_QAR"].Value = tttt;


                //******************M************************
                // double M11 = 0.15;
                string s6 = (dbdataset.Rows[0][4]).ToString();
                // double M11 = float.Parse(s6);
                double M11 = (int.Parse(dataGridView2.Rows[0].Cells["TEBYAN_Freight_and_Customs"].Value.ToString()));


                double T_Freight_customs_10per_add_5per = (T_Cost_Price_QAR) * (M11 / 100);

                double totalCost = Convert.ToDouble(String.Format("{0:0.00}", T_Freight_customs_10per_add_5per));

                dataGridView1.Rows[i].Cells["T_Freight_customs_10per_add_5per"].Value = totalCost;

                //********************N*************************

                double T_Landed_Cost_QAR = (T_Cost_Price_QAR) + (T_Freight_customs_10per_add_5per);

                double tt = Convert.ToDouble(String.Format("{0:0.00}", T_Landed_Cost_QAR));


                dataGridView1.Rows[i].Cells["T_Landed_Cost_QAR"].Value = tt;



                //***************R********************
                // double R11 = 0.75;
                string s7 = (dbdataset.Rows[0][5]).ToString();
                // double R11 = float.Parse(s7);


                double R11 = (int.Parse(dataGridView2.Rows[0].Cells["VIRGIN_percentage"].Value.ToString()));

                double V_Retailer_Cost_Price_QAR = (float.Parse(dataGridView1.Rows[i].Cells["QAR_RRP_QAR"].Value.ToString())) * (R11 / 100);
                dataGridView1.Rows[i].Cells["V_Retailer_Cost_Price_QAR"].Value = V_Retailer_Cost_Price_QAR;


                //*******************O******************
                double T_Margin_QAR = (V_Retailer_Cost_Price_QAR) - (T_Landed_Cost_QAR);


                double ttt = Convert.ToDouble(String.Format("{0:0.00}", T_Margin_QAR));

                dataGridView1.Rows[i].Cells["T_Margin_QAR"].Value = ttt;


                //*********************P*****************

                double T_Margin_Per = (T_Margin_QAR) / (V_Retailer_Cost_Price_QAR);

                double ppp = Convert.ToDouble(String.Format("{0:0.00}", T_Margin_Per));


                dataGridView1.Rows[i].Cells["T_Margin_Per"].Value = ppp;



                //******************S*********************

                double V_Retailer_Margin_QAR = (float.Parse(dataGridView1.Rows[i].Cells["QAR_RRP_QAR"].Value.ToString())) - (V_Retailer_Cost_Price_QAR);
                dataGridView1.Rows[i].Cells["V_Retailer_Margin_QAR"].Value = V_Retailer_Margin_QAR;

                //***********************T*******************


                double V_Retailer_Margin_per = (V_Retailer_Margin_QAR) / (float.Parse(dataGridView1.Rows[i].Cells["QAR_RRP_QAR"].Value.ToString()));
                double pppp = Convert.ToDouble(String.Format("{0:0.00}", V_Retailer_Margin_per));


                dataGridView1.Rows[i].Cells["V_Retailer_Margin_per"].Value = pppp;

                //**************************W*************************
                // double X9 = 1.01;
                string s8 = (dbdataset.Rows[0][7]).ToString();
                //  double X9 = float.Parse(s8);

                double X9 = (float.Parse(dataGridView2.Rows[0].Cells["X9"].Value.ToString()));


                double QAR_RRP_AED = (float.Parse(dataGridView1.Rows[i].Cells["QAR_RRP_QAR"].Value.ToString())) * (X9);
                dataGridView1.Rows[i].Cells["QAR_RRP_AED"].Value = QAR_RRP_AED;





            }
        }

        private void button3_Click(object sender, EventArgs e)
        {


            switch (NumberOfClick)
            {

                case 2:
                    {

                        SqlConnection con1 = new SqlConnection(conString);
                        con1.Open();

                        string sqlTrunc = "TRUNCATE TABLE  dbo.Products_New";
                        SqlCommand cmd3 = new SqlCommand(sqlTrunc, con1);
                        cmd3.ExecuteNonQuery();


                        int j = dataGridView1.Rows.Count;
                        int[] arr = Enumerable.Range(0, j - 1).ToArray();
                        foreach (var i in arr)
                        {
                            String st = "INSERT INTO dbo.Products_New(productsID,SubGroup_Name,Group_Name,product,version,Unit_cost_USD,Unit_cost_AED,Freight_customs,Financing,PRECISE_Landed_Cost,PRECISE_Margin_AED,PRECISE_Margin_Cost_Ratio,T_Cost_Price_USD,T_Cost_Price_QAR,T_Freight_customs_10per_add_5per,T_Landed_Cost_QAR,T_Margin_QAR,T_Margin_Per,V_Retailer_Cost_Price_QAR,V_Retailer_Margin_QAR,V_Retailer_Margin_per,UAE_RRP_AED,QAR_RRP_AED,QAR_RRP_QAR) values (@productsID,@SubGroup_Name,@Group_Name,  @product,@version,@Unit_cost_USD,@Unit_cost_AED,@Freight_customs,@Financing,@PRECISE_Landed_Cost,@PRECISE_Margin_AED,@PRECISE_Margin_Cost_Ratio,@T_Cost_Price_USD,@T_Cost_Price_QAR,@T_Freight_customs_10per_add_5per,@T_Landed_Cost_QAR,@T_Margin_QAR,@T_Margin_Per,@V_Retailer_Cost_Price_QAR,@V_Retailer_Margin_QAR,@V_Retailer_Margin_per,@UAE_RRP_AED,@QAR_RRP_AED,@QAR_RRP_QAR)";
                            SqlCommand cmd = new SqlCommand(st, con1);
                            cmd.Parameters.AddWithValue("@productsID", dataGridView1.Rows[i].Cells["productsID"].Value.ToString());
                            cmd.Parameters.AddWithValue("@SubGroup_Name", dataGridView1.Rows[i].Cells["SubGroup_Name"].Value.ToString());

                            cmd.Parameters.AddWithValue("@Group_Name", dataGridView1.Rows[i].Cells["Group_Name"].Value.ToString());
                            cmd.Parameters.AddWithValue("@product", dataGridView1.Rows[i].Cells["product"].Value.ToString());
                            cmd.Parameters.AddWithValue("@Unit_cost_USD", float.Parse(dataGridView1.Rows[i].Cells["Unit_cost_USD"].Value.ToString()));
                            cmd.Parameters.AddWithValue("@UAE_RRP_AED", float.Parse(dataGridView1.Rows[i].Cells["UAE_RRP_AED"].Value.ToString()));
                            // cmd.Parameters.AddWithValue(" @QAR_RRP_QAR", float.Parse(textBox6.Text));
                            cmd.Parameters.AddWithValue("@QAR_RRP_QAR", float.Parse(dataGridView1.Rows[i].Cells["QAR_RRP_QAR"].Value.ToString()));
                            cmd.Parameters.AddWithValue("@version", dataGridView1.Rows[i].Cells["version"].Value.ToString());
                            //**************************C**************
                            // double USD_Exchange_Rate_to_AED = 3.68;
                            // string s1 = "Select  * From dbo.Margin ";
                            // SqlCommand cmd9 = new SqlCommand(s1, con1);

                            // SqlDataAdapter sda = new SqlDataAdapter();
                            /// sda.SelectCommand = cmd9;
                            // DataTable dbdataset = new DataTable();
                            // sda.Fill(dbdataset);

                            //string s10 = (dbdataset.Rows[0][8]).ToString();
                            // double USD_Exchange_Rate_to_AED = float.Parse(s10);
                            // double USD_Exchange_Rate_to_AED = (float.Parse(dataGridView2.Rows[0].Cells["v9"].Value.ToString()));


                            // double Unit_cost_AED = (USD_Exchange_Rate_to_AED) * (float.Parse(dataGridView1.Rows[i].Cells["Unit_cost_USD"].Value.ToString()));

                            cmd.Parameters.AddWithValue("@Unit_cost_AED", float.Parse(dataGridView1.Rows[i].Cells["Unit_cost_AED"].Value.ToString()));
                            //***********************E**************
                            // double Percent = 0;
                            // string s2 = (dbdataset.Rows[0][1]).ToString();
                            // double Percent = float.Parse(s2);
                            //  double Percent = (int.Parse(dataGridView2.Rows[0].Cells["Precise_Freight_and_Customs"].Value.ToString()));



                            // double Freight_customs = (float.Parse(dataGridView1.Rows[i].Cells["Unit_cost_USD"].Value.ToString())) * (Percent/100);
                            cmd.Parameters.AddWithValue("@Freight_customs", float.Parse(dataGridView1.Rows[i].Cells["Freight_customs"].Value.ToString()));

                            // cmd.Parameters.AddWithValue("@Freight_customs", Freight_customs);
                            //*************F***********************
                            // double Percent_f = 0;
                            //string s3 = (dbdataset.Rows[0][2]).ToString();
                            // double Percent_f = float.Parse(s3);
                            // double Percent_f = (int.Parse(dataGridView2.Rows[0].Cells["Financing_per_month"].Value.ToString()));


                            // double Financing = (float.Parse(dataGridView1.Rows[i].Cells["Unit_cost_USD"].Value.ToString())) *( Percent_f/100);

                            // cmd.Parameters.AddWithValue("@Financing", Financing);

                            cmd.Parameters.AddWithValue("@Financing", float.Parse(dataGridView1.Rows[i].Cells["Financing"].Value.ToString()));

                            //*************G***********************


                            // double PRECISE_Landed_Cost = (Freight_customs) + (float.Parse(dataGridView1.Rows[i].Cells["Unit_cost_USD"].Value.ToString())) + (Financing);

                            // cmd.Parameters.AddWithValue("@PRECISE_Landed_Cost", PRECISE_Landed_Cost);
                            cmd.Parameters.AddWithValue("@PRECISE_Landed_Cost", float.Parse(dataGridView1.Rows[i].Cells["PRECISE_Landed_Cost"].Value.ToString()));



                            //**********************K***********************
                            // double K11 = 0.33;
                            // string s4 = (dbdataset.Rows[0][3]).ToString();
                            // double K11 = float.Parse(s4);
                            //  double K11 = (int.Parse(dataGridView2.Rows[0].Cells["K11"].Value.ToString()));


                            // double T_Cost_Price_USD = ((PRECISE_Landed_Cost) * (K11/100)) + (PRECISE_Landed_Cost);

                            //cmd.Parameters.AddWithValue("@T_Cost_Price_USD", T_Cost_Price_USD);
                            cmd.Parameters.AddWithValue("@T_Cost_Price_USD", float.Parse(dataGridView1.Rows[i].Cells["T_Cost_Price_USD"].Value.ToString()));

                            //****************H*******************

                            // double PRECISE_Margin_AED = (T_Cost_Price_USD) - (PRECISE_Landed_Cost);

                            // cmd.Parameters.AddWithValue("@PRECISE_Margin_AED", PRECISE_Margin_AED);

                            cmd.Parameters.AddWithValue("@PRECISE_Margin_AED", float.Parse(dataGridView1.Rows[i].Cells["PRECISE_Margin_AED"].Value.ToString()));

                            //*******************I*****************
                            // double PRECISE_Margin_Cost_Ratio = (PRECISE_Margin_AED) / (float.Parse(dataGridView1.Rows[i].Cells["Unit_cost_USD"].Value.ToString()));

                            // cmd.Parameters.AddWithValue("@PRECISE_Margin_Cost_Ratio", PRECISE_Margin_Cost_Ratio);
                            cmd.Parameters.AddWithValue("@PRECISE_Margin_Cost_Ratio", float.Parse(dataGridView1.Rows[i].Cells["PRECISE_Margin_Cost_Ratio"].Value.ToString()));



                            //********************L***************************
                            //double W9 = 3.64;
                            // string s5 = (dbdataset.Rows[0][6]).ToString();
                            // double W9 = float.Parse(s5);
                            // double W9 = (float.Parse(dataGridView2.Rows[0].Cells["W9"].Value.ToString()));

                            // double T_Cost_Price_QAR = (T_Cost_Price_USD) * (W9);

                            // cmd.Parameters.AddWithValue("@T_Cost_Price_QAR", T_Cost_Price_QAR);
                            cmd.Parameters.AddWithValue("@T_Cost_Price_QAR", float.Parse(dataGridView1.Rows[i].Cells["T_Cost_Price_QAR"].Value.ToString()));


                            //******************M************************
                            // double M11 = 0.15;

                            // string s6 = (dbdataset.Rows[0][4]).ToString();
                            // double M11 = float.Parse(s6);
                            // double M11 = (int.Parse(dataGridView2.Rows[0].Cells["TEBYAN_Freight_and_Customs"].Value.ToString()));

                            // double T_Freight_customs_10per_add_5per = (T_Cost_Price_QAR) * (M11/100);


                            //cmd.Parameters.AddWithValue("@T_Freight_customs_10per_add_5per", T_Freight_customs_10per_add_5per);
                            cmd.Parameters.AddWithValue("@T_Freight_customs_10per_add_5per", float.Parse(dataGridView1.Rows[i].Cells["T_Freight_customs_10per_add_5per"].Value.ToString()));


                            //********************N*************************


                            // double T_Landed_Cost_QAR = (T_Cost_Price_QAR) + (T_Freight_customs_10per_add_5per);

                            //  cmd.Parameters.AddWithValue("@T_Landed_Cost_QAR", T_Landed_Cost_QAR);
                            cmd.Parameters.AddWithValue("@T_Landed_Cost_QAR", float.Parse(dataGridView1.Rows[i].Cells["T_Landed_Cost_QAR"].Value.ToString()));



                            //***************R********************
                            // double R11 = 0.75;
                            // string s7 = (dbdataset.Rows[0][5]).ToString();
                            //double R11 = float.Parse(s7);
                            // double R11 = (int.Parse(dataGridView2.Rows[0].Cells["VIRGIN_percentage"].Value.ToString()));

                            //double V_Retailer_Cost_Price_QAR = (float.Parse(dataGridView1.Rows[i].Cells["QAR_RRP_QAR"].Value.ToString())) * (R11/100);

                            // cmd.Parameters.AddWithValue("@V_Retailer_Cost_Price_QAR", V_Retailer_Cost_Price_QAR);

                            cmd.Parameters.AddWithValue("@V_Retailer_Cost_Price_QAR", float.Parse(dataGridView1.Rows[i].Cells["V_Retailer_Cost_Price_QAR"].Value.ToString()));

                            //*******************O******************
                            // double T_Margin_QAR = (V_Retailer_Cost_Price_QAR) - (T_Landed_Cost_QAR);

                            // cmd.Parameters.AddWithValue("@T_Margin_QAR", T_Margin_QAR);

                            cmd.Parameters.AddWithValue("@T_Margin_QAR", float.Parse(dataGridView1.Rows[i].Cells["T_Margin_QAR"].Value.ToString()));

                            //*********************P*****************

                            //double T_Margin_Per = (T_Margin_QAR) / (V_Retailer_Cost_Price_QAR);

                            // cmd.Parameters.AddWithValue("@T_Margin_Per", T_Margin_Per);

                            cmd.Parameters.AddWithValue("@T_Margin_Per", float.Parse(dataGridView1.Rows[i].Cells["T_Margin_Per"].Value.ToString()));


                            //******************S*********************

                            // double V_Retailer_Margin_QAR = (float.Parse(dataGridView1.Rows[i].Cells["QAR_RRP_QAR"].Value.ToString())) - (V_Retailer_Cost_Price_QAR);

                            //cmd.Parameters.AddWithValue("@V_Retailer_Margin_QAR", V_Retailer_Margin_QAR);

                            cmd.Parameters.AddWithValue("@V_Retailer_Margin_QAR", float.Parse(dataGridView1.Rows[i].Cells["V_Retailer_Margin_QAR"].Value.ToString()));

                            //***********************T*******************

                            // double V_Retailer_Margin_per = (V_Retailer_Margin_QAR) / (float.Parse(dataGridView1.Rows[i].Cells["QAR_RRP_QAR"].Value.ToString()));

                            // cmd.Parameters.AddWithValue("@V_Retailer_Margin_per", V_Retailer_Margin_per);
                            cmd.Parameters.AddWithValue("@V_Retailer_Margin_per", float.Parse(dataGridView1.Rows[i].Cells["V_Retailer_Margin_per"].Value.ToString()));



                            //**************************W*************************
                            // double X9 = 1.01;

                            //string s8 = (dbdataset.Rows[0][7]).ToString();
                            //double X9 = float.Parse(s8);
                            //  double X9 = (float.Parse(dataGridView2.Rows[0].Cells["X9"].Value.ToString()));

                            //  double QAR_RRP_AED = (float.Parse(dataGridView1.Rows[i].Cells["QAR_RRP_QAR"].Value.ToString())) * (X9);

                            // cmd.Parameters.AddWithValue("@QAR_RRP_AED", QAR_RRP_AED);
                            cmd.Parameters.AddWithValue("@QAR_RRP_AED", float.Parse(dataGridView1.Rows[i].Cells["QAR_RRP_AED"].Value.ToString()));





                            cmd.ExecuteNonQuery();

                        }
                        string sqlTrunc2 = "TRUNCATE TABLE  dbo.Margin";
                        SqlCommand cmd4 = new SqlCommand(sqlTrunc2, con1);
                        cmd4.ExecuteNonQuery();


                        String st9 = "INSERT INTO dbo.Margin(Product_Cost_Category,Precise_Freight_and_Customs,Financing_per_month,k11,TEBYAN_Freight_and_Customs,VIRGIN_percentage,w9,x9,v9) values (@Product_Cost_Category,@Precise_Freight_and_Customs,@Financing_per_month,@k11,@TEBYAN_Freight_and_Customs,@VIRGIN_percentage,@w9,@x9,@v9)";
                        SqlCommand cmd99 = new SqlCommand(st9, con1);
                        cmd99.Parameters.AddWithValue("@Product_Cost_Category", dataGridView2.Rows[0].Cells["Product_Cost_Category"].Value.ToString());
                        cmd99.Parameters.AddWithValue("@Precise_Freight_and_Customs", int.Parse(dataGridView2.Rows[0].Cells["Precise_Freight_and_Customs"].Value.ToString()));
                        cmd99.Parameters.AddWithValue("@Financing_per_month", int.Parse(dataGridView2.Rows[0].Cells["Financing_per_month"].Value.ToString()));
                        cmd99.Parameters.AddWithValue("@k11", int.Parse(dataGridView2.Rows[0].Cells["k11"].Value.ToString()));
                        cmd99.Parameters.AddWithValue("@TEBYAN_Freight_and_Customs", int.Parse(dataGridView2.Rows[0].Cells["TEBYAN_Freight_and_Customs"].Value.ToString()));
                        cmd99.Parameters.AddWithValue("@VIRGIN_percentage", int.Parse(dataGridView2.Rows[0].Cells["VIRGIN_percentage"].Value.ToString()));
                        cmd99.Parameters.AddWithValue("@w9", float.Parse(dataGridView2.Rows[0].Cells["w9"].Value.ToString()));
                        cmd99.Parameters.AddWithValue("@x9", float.Parse(dataGridView2.Rows[0].Cells["x9"].Value.ToString()));
                        cmd99.Parameters.AddWithValue("@v9", float.Parse(dataGridView2.Rows[0].Cells["v9"].Value.ToString()));



                        cmd99.ExecuteNonQuery();

                        con1.Close();

                        this.Hide();





                    }
                    break;

                case 1:

                    break;
            }
        }


        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            switch (NumberOfClick)
            {

                case 2:
                    {

                        SqlConnection con1 = new SqlConnection(conString);
                        con1.Open();


                        string sqlTrunc = "TRUNCATE TABLE  dbo.For_print";
                        SqlCommand cmd3 = new SqlCommand(sqlTrunc, con1);
                        cmd3.ExecuteNonQuery();
                        con1.Close();


                        SqlConnection con111 = new SqlConnection(conString);
                        con111.Open();

                        int j = dataGridView1.Rows.Count;
                        int[] arr = Enumerable.Range(0, j - 1).ToArray();
                        foreach (var ii in arr)
                        {
                            String st = "INSERT INTO dbo.For_print(productsID,SubGroup_Name,Group_Name,product,version,Unit_cost_USD,Unit_cost_AED,Freight_customs,Financing,PRECISE_Landed_Cost,PRECISE_Margin_AED,PRECISE_Margin_Cost_Ratio,T_Cost_Price_USD,T_Cost_Price_QAR,T_Freight_customs_10per_add_5per,T_Landed_Cost_QAR,T_Margin_QAR,T_Margin_Per,V_Retailer_Cost_Price_QAR,V_Retailer_Margin_QAR,V_Retailer_Margin_per,UAE_RRP_AED,QAR_RRP_AED,QAR_RRP_QAR) values (@productsID,@SubGroup_Name,@Group_Name,  @product,@version,@Unit_cost_USD,@Unit_cost_AED,@Freight_customs,@Financing,@PRECISE_Landed_Cost,@PRECISE_Margin_AED,@PRECISE_Margin_Cost_Ratio,@T_Cost_Price_USD,@T_Cost_Price_QAR,@T_Freight_customs_10per_add_5per,@T_Landed_Cost_QAR,@T_Margin_QAR,@T_Margin_Per,@V_Retailer_Cost_Price_QAR,@V_Retailer_Margin_QAR,@V_Retailer_Margin_per,@UAE_RRP_AED,@QAR_RRP_AED,@QAR_RRP_QAR)";
                            SqlCommand cmd666 = new SqlCommand(st, con111);
                            cmd666.Parameters.AddWithValue("@productsID", dataGridView1.Rows[ii].Cells["productsID"].Value.ToString());
                            cmd666.Parameters.AddWithValue("@SubGroup_Name", dataGridView1.Rows[ii].Cells["SubGroup_Name"].Value.ToString());

                            cmd666.Parameters.AddWithValue("@Group_Name", dataGridView1.Rows[ii].Cells["Group_Name"].Value.ToString());
                            cmd666.Parameters.AddWithValue("@product", dataGridView1.Rows[ii].Cells["product"].Value.ToString());
                            cmd666.Parameters.AddWithValue("@Unit_cost_USD", float.Parse(dataGridView1.Rows[ii].Cells["Unit_cost_USD"].Value.ToString()));
                            cmd666.Parameters.AddWithValue("@UAE_RRP_AED", float.Parse(dataGridView1.Rows[ii].Cells["UAE_RRP_AED"].Value.ToString()));
                            // cmd.Parameters.AddWithValue(" @QAR_RRP_QAR", float.Parse(textBox6.Text));
                            cmd666.Parameters.AddWithValue("@QAR_RRP_QAR", float.Parse(dataGridView1.Rows[ii].Cells["QAR_RRP_QAR"].Value.ToString()));
                            cmd666.Parameters.AddWithValue("@version", dataGridView1.Rows[ii].Cells["version"].Value.ToString());
                            //**************************C**************
                            // double USD_Exchange_Rate_to_AED = 3.68;
                            // string s1 = "Select  * From dbo.Margin ";
                            // SqlCommand cmd9 = new SqlCommand(s1, con1);

                            // SqlDataAdapter sda = new SqlDataAdapter();
                            /// sda.SelectCommand = cmd9;
                            // DataTable dbdataset = new DataTable();
                            // sda.Fill(dbdataset);

                            //string s10 = (dbdataset.Rows[0][8]).ToString();
                            // double USD_Exchange_Rate_to_AED = float.Parse(s10);
                            // double USD_Exchange_Rate_to_AED = (float.Parse(dataGridView2.Rows[0].Cells["v9"].Value.ToString()));


                            // double Unit_cost_AED = (USD_Exchange_Rate_to_AED) * (float.Parse(dataGridView1.Rows[i].Cells["Unit_cost_USD"].Value.ToString()));

                            cmd666.Parameters.AddWithValue("@Unit_cost_AED", float.Parse(dataGridView1.Rows[ii].Cells["Unit_cost_AED"].Value.ToString()));
                            //***********************E**************
                            // double Percent = 0;
                            // string s2 = (dbdataset.Rows[0][1]).ToString();
                            // double Percent = float.Parse(s2);
                            //  double Percent = (int.Parse(dataGridView2.Rows[0].Cells["Precise_Freight_and_Customs"].Value.ToString()));



                            // double Freight_customs = (float.Parse(dataGridView1.Rows[i].Cells["Unit_cost_USD"].Value.ToString())) * (Percent/100);
                            cmd666.Parameters.AddWithValue("@Freight_customs", float.Parse(dataGridView1.Rows[ii].Cells["Freight_customs"].Value.ToString()));

                            // cmd.Parameters.AddWithValue("@Freight_customs", Freight_customs);
                            //*************F***********************
                            // double Percent_f = 0;
                            //string s3 = (dbdataset.Rows[0][2]).ToString();
                            // double Percent_f = float.Parse(s3);
                            // double Percent_f = (int.Parse(dataGridView2.Rows[0].Cells["Financing_per_month"].Value.ToString()));


                            // double Financing = (float.Parse(dataGridView1.Rows[i].Cells["Unit_cost_USD"].Value.ToString())) *( Percent_f/100);

                            // cmd.Parameters.AddWithValue("@Financing", Financing);

                            cmd666.Parameters.AddWithValue("@Financing", float.Parse(dataGridView1.Rows[ii].Cells["Financing"].Value.ToString()));

                            //*************G***********************


                            // double PRECISE_Landed_Cost = (Freight_customs) + (float.Parse(dataGridView1.Rows[i].Cells["Unit_cost_USD"].Value.ToString())) + (Financing);

                            // cmd.Parameters.AddWithValue("@PRECISE_Landed_Cost", PRECISE_Landed_Cost);
                            cmd666.Parameters.AddWithValue("@PRECISE_Landed_Cost", float.Parse(dataGridView1.Rows[ii].Cells["PRECISE_Landed_Cost"].Value.ToString()));



                            //**********************K***********************
                            // double K11 = 0.33;
                            // string s4 = (dbdataset.Rows[0][3]).ToString();
                            // double K11 = float.Parse(s4);
                            //  double K11 = (int.Parse(dataGridView2.Rows[0].Cells["K11"].Value.ToString()));


                            // double T_Cost_Price_USD = ((PRECISE_Landed_Cost) * (K11/100)) + (PRECISE_Landed_Cost);

                            //cmd.Parameters.AddWithValue("@T_Cost_Price_USD", T_Cost_Price_USD);
                            cmd666.Parameters.AddWithValue("@T_Cost_Price_USD", float.Parse(dataGridView1.Rows[ii].Cells["T_Cost_Price_USD"].Value.ToString()));

                            //****************H*******************

                            // double PRECISE_Margin_AED = (T_Cost_Price_USD) - (PRECISE_Landed_Cost);

                            // cmd.Parameters.AddWithValue("@PRECISE_Margin_AED", PRECISE_Margin_AED);

                            cmd666.Parameters.AddWithValue("@PRECISE_Margin_AED", float.Parse(dataGridView1.Rows[ii].Cells["PRECISE_Margin_AED"].Value.ToString()));

                            //*******************I*****************
                            // double PRECISE_Margin_Cost_Ratio = (PRECISE_Margin_AED) / (float.Parse(dataGridView1.Rows[i].Cells["Unit_cost_USD"].Value.ToString()));

                            // cmd.Parameters.AddWithValue("@PRECISE_Margin_Cost_Ratio", PRECISE_Margin_Cost_Ratio);
                            cmd666.Parameters.AddWithValue("@PRECISE_Margin_Cost_Ratio", float.Parse(dataGridView1.Rows[ii].Cells["PRECISE_Margin_Cost_Ratio"].Value.ToString()));



                            //********************L***************************
                            //double W9 = 3.64;
                            // string s5 = (dbdataset.Rows[0][6]).ToString();
                            // double W9 = float.Parse(s5);
                            // double W9 = (float.Parse(dataGridView2.Rows[0].Cells["W9"].Value.ToString()));

                            // double T_Cost_Price_QAR = (T_Cost_Price_USD) * (W9);

                            // cmd.Parameters.AddWithValue("@T_Cost_Price_QAR", T_Cost_Price_QAR);
                            cmd666.Parameters.AddWithValue("@T_Cost_Price_QAR", float.Parse(dataGridView1.Rows[ii].Cells["T_Cost_Price_QAR"].Value.ToString()));


                            //******************M************************
                            // double M11 = 0.15;

                            // string s6 = (dbdataset.Rows[0][4]).ToString();
                            // double M11 = float.Parse(s6);
                            // double M11 = (int.Parse(dataGridView2.Rows[0].Cells["TEBYAN_Freight_and_Customs"].Value.ToString()));

                            // double T_Freight_customs_10per_add_5per = (T_Cost_Price_QAR) * (M11/100);


                            //cmd.Parameters.AddWithValue("@T_Freight_customs_10per_add_5per", T_Freight_customs_10per_add_5per);
                            cmd666.Parameters.AddWithValue("@T_Freight_customs_10per_add_5per", float.Parse(dataGridView1.Rows[ii].Cells["T_Freight_customs_10per_add_5per"].Value.ToString()));


                            //********************N*************************


                            // double T_Landed_Cost_QAR = (T_Cost_Price_QAR) + (T_Freight_customs_10per_add_5per);

                            //  cmd.Parameters.AddWithValue("@T_Landed_Cost_QAR", T_Landed_Cost_QAR);
                            cmd666.Parameters.AddWithValue("@T_Landed_Cost_QAR", float.Parse(dataGridView1.Rows[ii].Cells["T_Landed_Cost_QAR"].Value.ToString()));



                            //***************R********************
                            // double R11 = 0.75;
                            // string s7 = (dbdataset.Rows[0][5]).ToString();
                            //double R11 = float.Parse(s7);
                            // double R11 = (int.Parse(dataGridView2.Rows[0].Cells["VIRGIN_percentage"].Value.ToString()));

                            //double V_Retailer_Cost_Price_QAR = (float.Parse(dataGridView1.Rows[i].Cells["QAR_RRP_QAR"].Value.ToString())) * (R11/100);

                            // cmd.Parameters.AddWithValue("@V_Retailer_Cost_Price_QAR", V_Retailer_Cost_Price_QAR);

                            cmd666.Parameters.AddWithValue("@V_Retailer_Cost_Price_QAR", float.Parse(dataGridView1.Rows[ii].Cells["V_Retailer_Cost_Price_QAR"].Value.ToString()));

                            //*******************O******************
                            // double T_Margin_QAR = (V_Retailer_Cost_Price_QAR) - (T_Landed_Cost_QAR);

                            // cmd.Parameters.AddWithValue("@T_Margin_QAR", T_Margin_QAR);

                            cmd666.Parameters.AddWithValue("@T_Margin_QAR", float.Parse(dataGridView1.Rows[ii].Cells["T_Margin_QAR"].Value.ToString()));

                            //*********************P*****************

                            //double T_Margin_Per = (T_Margin_QAR) / (V_Retailer_Cost_Price_QAR);

                            // cmd.Parameters.AddWithValue("@T_Margin_Per", T_Margin_Per);

                            cmd666.Parameters.AddWithValue("@T_Margin_Per", float.Parse(dataGridView1.Rows[ii].Cells["T_Margin_Per"].Value.ToString()));


                            //******************S*********************

                            // double V_Retailer_Margin_QAR = (float.Parse(dataGridView1.Rows[i].Cells["QAR_RRP_QAR"].Value.ToString())) - (V_Retailer_Cost_Price_QAR);

                            //cmd.Parameters.AddWithValue("@V_Retailer_Margin_QAR", V_Retailer_Margin_QAR);

                            cmd666.Parameters.AddWithValue("@V_Retailer_Margin_QAR", float.Parse(dataGridView1.Rows[ii].Cells["V_Retailer_Margin_QAR"].Value.ToString()));

                            //***********************T*******************

                            // double V_Retailer_Margin_per = (V_Retailer_Margin_QAR) / (float.Parse(dataGridView1.Rows[i].Cells["QAR_RRP_QAR"].Value.ToString()));

                            // cmd.Parameters.AddWithValue("@V_Retailer_Margin_per", V_Retailer_Margin_per);
                            cmd666.Parameters.AddWithValue("@V_Retailer_Margin_per", float.Parse(dataGridView1.Rows[ii].Cells["V_Retailer_Margin_per"].Value.ToString()));



                            //**************************W*************************
                            // double X9 = 1.01;

                            //string s8 = (dbdataset.Rows[0][7]).ToString();
                            //double X9 = float.Parse(s8);
                            //  double X9 = (float.Parse(dataGridView2.Rows[0].Cells["X9"].Value.ToString()));

                            //  double QAR_RRP_AED = (float.Parse(dataGridView1.Rows[i].Cells["QAR_RRP_QAR"].Value.ToString())) * (X9);

                            // cmd.Parameters.AddWithValue("@QAR_RRP_AED", QAR_RRP_AED);
                            cmd666.Parameters.AddWithValue("@QAR_RRP_AED", float.Parse(dataGridView1.Rows[ii].Cells["QAR_RRP_AED"].Value.ToString()));





                            cmd666.ExecuteNonQuery();


                        }


                        con111.Close();

                        SqlConnection con11 = new SqlConnection(conString);

                        con11.Open();
                        string sqlTrunc2 = "TRUNCATE TABLE  dbo.Margin_for_print";
                        SqlCommand cmd4 = new SqlCommand(sqlTrunc2, con11);
                        cmd4.ExecuteNonQuery();


                        String st9 = "INSERT INTO dbo.Margin_for_print(Product_Cost_Category,Precise_Freight_and_Customs,Financing_per_month,k11,TEBYAN_Freight_and_Customs,VIRGIN_percentage,w9,x9,v9) values (@Product_Cost_Category,@Precise_Freight_and_Customs,@Financing_per_month,@k11,@TEBYAN_Freight_and_Customs,@VIRGIN_percentage,@w9,@x9,@v9)";
                        SqlCommand cmd99 = new SqlCommand(st9, con11);
                        cmd99.Parameters.AddWithValue("@Product_Cost_Category", dataGridView2.Rows[0].Cells["Product_Cost_Category"].Value.ToString());
                        cmd99.Parameters.AddWithValue("@Precise_Freight_and_Customs", int.Parse(dataGridView2.Rows[0].Cells["Precise_Freight_and_Customs"].Value.ToString()));
                        cmd99.Parameters.AddWithValue("@Financing_per_month", int.Parse(dataGridView2.Rows[0].Cells["Financing_per_month"].Value.ToString()));
                        cmd99.Parameters.AddWithValue("@k11", int.Parse(dataGridView2.Rows[0].Cells["k11"].Value.ToString()));
                        cmd99.Parameters.AddWithValue("@TEBYAN_Freight_and_Customs", int.Parse(dataGridView2.Rows[0].Cells["TEBYAN_Freight_and_Customs"].Value.ToString()));
                        cmd99.Parameters.AddWithValue("@VIRGIN_percentage", int.Parse(dataGridView2.Rows[0].Cells["VIRGIN_percentage"].Value.ToString()));
                        cmd99.Parameters.AddWithValue("@w9", float.Parse(dataGridView2.Rows[0].Cells["w9"].Value.ToString()));
                        cmd99.Parameters.AddWithValue("@x9", float.Parse(dataGridView2.Rows[0].Cells["x9"].Value.ToString()));
                        cmd99.Parameters.AddWithValue("@v9", float.Parse(dataGridView2.Rows[0].Cells["v9"].Value.ToString()));



                        cmd99.ExecuteNonQuery();

                        con11.Close();
                        SqlConnection cnn;
                        // string connectionstring = null;
                        string sql = null;
                        string sql2 = null;
                        string data = null;
                        string data1 = null;

                        int i = 0;
                        // int j = 0;

                        Excel.Application xlApp2;
                        Excel.Workbook xlWorkBook2;
                        Excel.Worksheet xlWorkSheet2;
                        object misValue = System.Reflection.Missing.Value;


                        xlApp2 = new Microsoft.Office.Interop.Excel.Application();
                        xlWorkBook2 = xlApp2.Workbooks.Add(misValue);
                        xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(1);

                        //connectionstring = "Data Source=IN-WTS-SAM;Initial Catalog=MSNETDB;Integrated Security=True;Pooling=False";
                        cnn = new SqlConnection(conString);
                        cnn.Open();
                        sql = "SELECT * FROM dbo.For_print";
                        SqlDataAdapter dscmd = new SqlDataAdapter(sql, cnn);
                        DataSet ds = new DataSet();
                        dscmd.Fill(ds);

                        // foreach (DataTable dt in ds.Tables)
                        // {
                        // for (int i1 = 0; i1 < dt.Columns.Count; i1++)
                        // {
                        //    xlWorkSheet.Cells[1, i1 + 1] = dt.Columns[i1].ColumnName;
                        // }
                        //  }
                        sql2 = "SELECT * FROM dbo.Margin_for_print";
                        SqlDataAdapter dscmd2 = new SqlDataAdapter(sql2, cnn);
                        DataSet ds2 = new DataSet();
                        dscmd2.Fill(ds2);




                        foreach (DataTable dt in ds.Tables)
                        {





                            xlWorkSheet2.Cells[2, 1] = "3DOODLER CREATE";
                            xlWorkSheet2.Cells[2, 1].Font.Bold = true;

                            xlWorkSheet2.Cells[3, 1] = "PRECISE - AGENT";
                            xlWorkSheet2.Cells[3, 1].Font.Bold = true;
                            xlWorkSheet2.Cells[4, 1] = "19-Mar-18";
                            xlWorkSheet2.Cells[4, 1].Font.Bold = true;
                            xlWorkSheet2.Cells[6, 1] = "TEBYAN FROM WOBBLEWORKS HK";
                            xlWorkSheet2.Cells[6, 1].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 1] = "Product ID";
                            xlWorkSheet2.Cells[11, 1].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 2] = "         Category";
                            xlWorkSheet2.Cells[11, 2].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 3] = "                  Product";
                            xlWorkSheet2.Cells[11, 3].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 4] = "Version";
                            xlWorkSheet2.Cells[11, 4].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 5] = "Unit Cost \n   (USD)";
                            xlWorkSheet2.Cells[11, 5].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 6] = "Unit Cost\n   (AED)";
                            xlWorkSheet2.Cells[11, 6].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 7] = "  ";
                            xlWorkSheet2.Cells[11, 8] = "Freight& \n Customs\n(10%+5%)";
                            xlWorkSheet2.Cells[11, 8].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 9] = "Financing \n 1%/Month";
                            xlWorkSheet2.Cells[11, 9].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 10] = "PRECISE \n Landed Cost";
                            xlWorkSheet2.Cells[11, 10].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 11] = "PRECISE Margin \n (AED)";
                            xlWorkSheet2.Cells[11, 11].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 12] = "PRECISE Margin \n Cost Ratio (%)";
                            xlWorkSheet2.Cells[11, 12].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 13] = "  ";
                            xlWorkSheet2.Cells[11, 14] = "Cost Price* \n (USD)";
                            xlWorkSheet2.Cells[11, 14].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 15] = "Cost Price* \n (QAR)";
                            xlWorkSheet2.Cells[11, 15].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 16] = "Freight & \n Customs \n (10%+5%)";
                            xlWorkSheet2.Cells[11, 16].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 17] = "Landed   Cost \n (QAR)";
                            xlWorkSheet2.Cells[11, 17].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 18] = "Margin \n   (QAR)";
                            xlWorkSheet2.Cells[11, 18].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 19] = "Margin \n    (%)";
                            xlWorkSheet2.Cells[11, 19].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 20] = "  ";
                            xlWorkSheet2.Cells[11, 21] = "Retailer Cost \n Price* (QAR)";
                            xlWorkSheet2.Cells[11, 21].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 22] = "Retailer \n  Margin   (QAR)";
                            xlWorkSheet2.Cells[11, 22].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 23] = "Retailer \n Margin \n   (%)";
                            xlWorkSheet2.Cells[11, 23].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 24] = "  ";
                            xlWorkSheet2.Cells[11, 25] = "UAE RRP \n (AED)";
                            xlWorkSheet2.Cells[11, 25].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 26] = "QAR RRP \n (AED)";
                            xlWorkSheet2.Cells[11, 26].Font.Bold = true;
                            xlWorkSheet2.Cells[11, 27] = "QAR  RRP \n (QAR)";
                            xlWorkSheet2.Cells[11, 27].Font.Bold = true;

                            xlWorkSheet2.Cells[9, 25] = " USD  Exchange\n Rate to AED";
                            xlWorkSheet2.Cells[9, 25].Font.Bold = true;

                            xlWorkSheet2.Cells[9, 26] = " USD  Exchange \n Rate to QAR";
                            xlWorkSheet2.Cells[9, 26].Font.Bold = true;

                            xlWorkSheet2.Cells[9, 27] = " QAR  Exchange \n Rate to AED";
                            xlWorkSheet2.Cells[9, 27].Font.Bold = true;










                            // change size of coloumn :
                            xlWorkSheet2.Columns["A:A"].ColumnWidth = 11;
                            xlWorkSheet2.Columns["C:C"].ColumnWidth = 30;
                            xlWorkSheet2.Columns["P:P"].ColumnWidth = 12;
                            xlWorkSheet2.Columns["B:B"].ColumnWidth = 17;
                            xlWorkSheet2.Columns["H:H"].ColumnWidth = 11;
                            xlWorkSheet2.Columns["Y:Y"].ColumnWidth = 11;
                            xlWorkSheet2.Columns["Z:Z"].ColumnWidth = 11;
                            xlWorkSheet2.Columns["AA:AA"].ColumnWidth = 11;
                            xlWorkSheet2.Columns["I:I"].ColumnWidth = 11;



                            // coloumn name:

                            //  for (int i1 = 0; i1 < dt.Columns.Count-7; i1++)
                            ///   {

                            //     xlWorkSheet.Cells[12, i1 + 8] = dt.Columns[i1+7].ColumnName;
                            //  }

                        }

                        // to jadvalaie sql index az 0 ast.

                        for (i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                        {
                            int s = i + 12;
                            // first coloumn:
                            data1 = ds.Tables[0].Rows[i].ItemArray[0].ToString();
                            xlWorkSheet2.Cells[s + 1, 1] = data1;

                            for (j = 2; j <= 6; j++)
                            {
                                data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                                xlWorkSheet2.Cells[s + 1, j] = data;
                            }

                            for (j = 7; j <= 11; j++)
                            {
                                data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                                xlWorkSheet2.Cells[s + 1, j + 1] = data;
                            }
                            for (j = 12; j <= 17; j++)
                            {
                                data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                                xlWorkSheet2.Cells[s + 1, j + 2] = data;
                            }

                            for (j = 18; j <= 20; j++)
                            {
                                data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                                xlWorkSheet2.Cells[s + 1, j + 3] = data;
                            }



                            for (j = 21; j <= ds.Tables[0].Columns.Count - 1; j++)
                            {
                                data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                                xlWorkSheet2.Cells[s + 1, j + 4] = data;
                            }
                        }



                        //sorting based on a coloumn:
                        dynamic allDataRange = xlWorkSheet2.get_Range("13:100");
                        // dynamic allDataRange = xlWorkSheet.UsedRange;
                        allDataRange.Sort(allDataRange.Columns[2], Excel.XlSortOrder.xlAscending);

                        //sorting with range determined:
                        //  Excel.Range Fruits = xlWorkSheet.get_Range("B13", "B100");
                        // Fruits.Sort(Fruits.Columns[1], Excel.XlSortOrder.xlAscending
                        // Fruits.Columns[3], Excel.XlSortOrder.xlAscending,
                        // Fruits.Columns[4], Excel.XlSortOrder.xlAscending

                        //);
                        // missing, Excel.XlSortOrder.xlAscending,
                        // Excel.XlYesNoGuess.xlNo, missing, missing,
                        // Excel.XlSortOrientation.xlSortColumns,
                        // Excel.XlSortMethod.xlPinYin,
                        // Excel.XlSortDataOption.xlSortNormal,
                        //  Excel.XlSortDataOption.xlSortNormal,
                        // Excel.XlSortDataOption.xlSortNormal);



                        // color of a cell:
                        Excel.Range formatRange;
                        formatRange = xlWorkSheet2.get_Range("A6", "AA6");
                        formatRange.Interior.Color = System.Drawing.
                        ColorTranslator.ToOle(System.Drawing.Color.Beige);
                        //xlWorkSheet.Cells[1, 1] = "Red";


                        Excel.Range formatRange1;
                        formatRange1 = xlWorkSheet2.get_Range("Y9", "AA9");
                        formatRange1.Interior.Color = System.Drawing.
                        ColorTranslator.ToOle(System.Drawing.Color.Beige);

                        // Excel.Range formatRange2;
                        // formatRange2 = xlWorkSheet.get_Range("A11", "AA11");
                        // formatRange2.Interior.Color = System.Drawing.
                        //  ColorTranslator.ToOle(System.Drawing.Color.Beige);

                        Excel.Range formatRange3;
                        formatRange3 = xlWorkSheet2.get_Range("A10", "AA10");
                        formatRange3.Interior.Color = System.Drawing.
                        ColorTranslator.ToOle(System.Drawing.Color.Olive);


                        //fontsize:
                        formatRange.Font.Size = 15;


                        // Excel.Range formatRange3;
                        // formatRange3 = xlWorkSheet.get_Range("Y13","Y20");
                        // formatRange3.Interior.Color = System.Drawing.
                        // ColorTranslator.ToOle(System.Drawing.Color.LightGreen);






                        // merging rows:
                        xlWorkSheet2.get_Range("N10", "S10").Merge(false);
                        Excel.Range chartRange = xlWorkSheet2.get_Range("N10", "S10");
                        chartRange.FormulaR1C1 = "TEBYAN - FROM SUPPLIER (EX-WORKS)";
                        chartRange.HorizontalAlignment = 3;
                        chartRange.VerticalAlignment = 3;

                        chartRange.Font.Bold = true;





                        xlWorkSheet2.get_Range("V10", "W10").Merge(false);
                        Excel.Range chartRange2 = xlWorkSheet2.get_Range("V10", "W10");
                        chartRange2.FormulaR1C1 = "VIRGIN";
                        chartRange2.HorizontalAlignment = 3;
                        chartRange2.VerticalAlignment = 3;

                        chartRange2.Font.Bold = true;

                        //formatRange.Font.Size = 15;



                        // margin table:


                        string data22 = (ds2.Tables[0].Rows[0].ItemArray[1].ToString());
                        xlWorkSheet2.Cells[12, 8] = data22 + "%";


                        string data23 = (ds2.Tables[0].Rows[0].ItemArray[2].ToString());
                        xlWorkSheet2.Cells[12, 9] = data23 + "%";


                        //  string data24 = (ds2.Tables[0].Rows[0].ItemArray[3].ToString());
                        //  xlWorkSheet.Cells[12, 14] = data24 + "%";

                        string data25 = (ds2.Tables[0].Rows[0].ItemArray[4].ToString());
                        xlWorkSheet2.Cells[12, 16] = data25 + "%";


                        // string data26 = (ds2.Tables[0].Rows[0].ItemArray[5].ToString());
                        // xlWorkSheet.Cells[12, 21] = data26 + "%";

                        //v9:
                        string data27 = (ds2.Tables[0].Rows[0].ItemArray[8].ToString());
                        xlWorkSheet2.Cells[10, 25] = "AED    " + data27;
                        xlWorkSheet2.Cells[10, 25].Font.Bold = true;
                        //w9:
                        string data28 = (ds2.Tables[0].Rows[0].ItemArray[6].ToString());
                        xlWorkSheet2.Cells[10, 26] = "QAR    " + data28;
                        xlWorkSheet2.Cells[10, 26].Font.Bold = true;
                        //x9:
                        string data29 = (ds2.Tables[0].Rows[0].ItemArray[7].ToString());
                        xlWorkSheet2.Cells[10, 27] = "AED    " + data29;
                        xlWorkSheet2.Cells[10, 27].Font.Bold = true;









                        xlWorkBook2.SaveAs("informations2.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                        xlWorkBook2.Close(true, misValue, misValue);
                        xlApp2.Quit();

                        releaseObject(xlWorkSheet2);
                        releaseObject(xlWorkBook2);
                        releaseObject(xlApp2);



                        MessageBox.Show("Excel file created , you can find the file C:\\Users\\....\\Recent\\informations2.xls");

                        this.Hide();

                        cnn.Close();
                    }




                    break;

                case 1:

                    break;
            }

        }
    }
}
