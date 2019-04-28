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

using System.IO;

using Excel = Microsoft.Office.Interop.Excel;
using ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat;



namespace PRECISE
{
    public partial class Form17 : Form
    {
        public Form17()
        {
            InitializeComponent();
            // List<TreeNode> checkedNodes_2 = new List<TreeNode>();

        }

        public string conString = System.Configuration.ConfigurationManager.ConnectionStrings["connection_string"].ConnectionString;

        // sending an array from a fprm to another form:
        //  public List<TreeNode> checkedNodes2 = new List<TreeNode>();



        // public List<TreeNode> myArray

        // {
        //   get
        //   {
        //        return checkedNodes2;
        //   }
        //   set
        //    {
        //        checkedNodes2 = value;
        //    }
        //  }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {


            SqlConnection con7 = new SqlConnection(conString);
            SqlCommand cmd7 = new SqlCommand("Select * from dbo.Products_order", con7);
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

                DataGridViewColumn column5= dataGridView1.Columns[5];
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



            //  foreach (var mm in checkedNodes2) { 
            //  SqlConnection con9 = new SqlConnection(conString);
            // con9.Open();
            // string sqlTrunc2 = "DELETE FROM dbo.Products_order where ('TreeNode: '+ product) = '" + mm + "' ";
            // SqlCommand cmd2 = new SqlCommand(sqlTrunc2, con9);
            //  cmd2.ExecuteNonQuery();
            // con9.Close();

            //}
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

        private void button2_Click_1(object sender, EventArgs e)
        {

            // SqlConnection con8 = new SqlConnection(conString);
            // SqlCommand cmd8 = new SqlCommand("Select * from dbo.Products_order;", con8);
            // DataTable dt = new DataTable();
            //  con8.Fill(dt);


            SqlConnection cnn;
            // string connectionstring = null;
            string sql = null;
            string sql2 = null;
            string data = null;
            string data1 = null;

            int i = 0;
            int j = 0;

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;


            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //connectionstring = "Data Source=IN-WTS-SAM;Initial Catalog=MSNETDB;Integrated Security=True;Pooling=False";
            cnn = new SqlConnection(conString);
            cnn.Open();
            sql = "SELECT * FROM dbo.Products_order";
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
            sql2 = "SELECT * FROM dbo.Margin";
            SqlDataAdapter dscmd2 = new SqlDataAdapter(sql2, cnn);
            DataSet ds2 = new DataSet();
            dscmd2.Fill(ds2);




            foreach (DataTable dt in ds.Tables)
            {





                xlWorkSheet.Cells[2, 1] = "3DOODLER CREATE";
                xlWorkSheet.Cells[2, 1].Font.Bold = true;

                xlWorkSheet.Cells[3, 1] = "PRECISE - AGENT";
                xlWorkSheet.Cells[3, 1].Font.Bold = true;
                xlWorkSheet.Cells[4, 1] = "19-Mar-18";
                xlWorkSheet.Cells[4, 1].Font.Bold = true;
                xlWorkSheet.Cells[6, 1] = "TEBYAN FROM WOBBLEWORKS HK";
                xlWorkSheet.Cells[6, 1].Font.Bold = true;
                xlWorkSheet.Cells[11, 1] = "Product ID";
                xlWorkSheet.Cells[11, 1].Font.Bold = true;
                xlWorkSheet.Cells[11, 2] = "         Category";
                xlWorkSheet.Cells[11, 2].Font.Bold = true;
                xlWorkSheet.Cells[11, 3] = "                  Product";
                xlWorkSheet.Cells[11, 3].Font.Bold = true;
                xlWorkSheet.Cells[11, 4] = "Version";
                xlWorkSheet.Cells[11, 4].Font.Bold = true;
                xlWorkSheet.Cells[11, 5] = "Unit Cost \n   (USD)";
                xlWorkSheet.Cells[11, 5].Font.Bold = true;
                xlWorkSheet.Cells[11, 6] = "Unit Cost\n   (AED)";
                xlWorkSheet.Cells[11, 6].Font.Bold = true;
                xlWorkSheet.Cells[11, 7] = "  ";
                xlWorkSheet.Cells[11, 8] = "Freight& \n Customs\n(10%+5%)";
                xlWorkSheet.Cells[11, 8].Font.Bold = true;
                xlWorkSheet.Cells[11, 9] = "Financing \n 1%/Month";
                xlWorkSheet.Cells[11, 9].Font.Bold = true;
                xlWorkSheet.Cells[11, 10] = "PRECISE \n Landed Cost";
                xlWorkSheet.Cells[11, 10].Font.Bold = true;
                xlWorkSheet.Cells[11, 11] = "PRECISE Margin \n (AED)";
                xlWorkSheet.Cells[11, 11].Font.Bold = true;
                xlWorkSheet.Cells[11, 12] = "PRECISE Margin \n Cost Ratio (%)";
                xlWorkSheet.Cells[11, 12].Font.Bold = true;
                xlWorkSheet.Cells[11, 13] = "  ";
                xlWorkSheet.Cells[11, 14] = "Cost Price* \n (USD)";
                xlWorkSheet.Cells[11, 14].Font.Bold = true;
                xlWorkSheet.Cells[11, 15] = "Cost Price* \n (QAR)";
                xlWorkSheet.Cells[11, 15].Font.Bold = true;
                xlWorkSheet.Cells[11, 16] = "Freight & \n Customs \n (10%+5%)";
                xlWorkSheet.Cells[11, 16].Font.Bold = true;
                xlWorkSheet.Cells[11, 17] = "Landed   Cost \n (QAR)";
                xlWorkSheet.Cells[11, 17].Font.Bold = true;
                xlWorkSheet.Cells[11, 18] = "Margin \n   (QAR)";
                xlWorkSheet.Cells[11, 18].Font.Bold = true;
                xlWorkSheet.Cells[11, 19] = "Margin \n    (%)";
                xlWorkSheet.Cells[11, 19].Font.Bold = true;
                xlWorkSheet.Cells[11, 20] = "  ";
                xlWorkSheet.Cells[11, 21] = "Retailer Cost \n Price* (QAR)";
                xlWorkSheet.Cells[11, 21].Font.Bold = true;
                xlWorkSheet.Cells[11, 22] = "Retailer \n  Margin   (QAR)";
                xlWorkSheet.Cells[11, 22].Font.Bold = true;
                xlWorkSheet.Cells[11, 23] = "Retailer \n Margin \n   (%)";
                xlWorkSheet.Cells[11, 23].Font.Bold = true;
                xlWorkSheet.Cells[11, 24] = "  ";
                xlWorkSheet.Cells[11, 25] = "UAE RRP \n (AED)";
                xlWorkSheet.Cells[11, 25].Font.Bold = true;
                xlWorkSheet.Cells[11, 26] = "QAR RRP \n (AED)";
                xlWorkSheet.Cells[11, 26].Font.Bold = true;
                xlWorkSheet.Cells[11, 27] = "QAR  RRP \n (QAR)";
                xlWorkSheet.Cells[11, 27].Font.Bold = true;

                xlWorkSheet.Cells[9, 25] = " USD  Exchange\n Rate to AED";
                xlWorkSheet.Cells[9, 25].Font.Bold = true;

                xlWorkSheet.Cells[9, 26] = " USD  Exchange \n Rate to QAR";
                xlWorkSheet.Cells[9, 26].Font.Bold = true;

                xlWorkSheet.Cells[9, 27] = " QAR  Exchange \n Rate to AED";
                xlWorkSheet.Cells[9, 27].Font.Bold = true;



                






                // change size of coloumn :
                xlWorkSheet.Columns["A:A"].ColumnWidth = 11;
                xlWorkSheet.Columns["C:C"].ColumnWidth = 30;
                xlWorkSheet.Columns["P:P"].ColumnWidth = 12;
                xlWorkSheet.Columns["B:B"].ColumnWidth = 17;
                xlWorkSheet.Columns["H:H"].ColumnWidth = 11;
                xlWorkSheet.Columns["Y:Y"].ColumnWidth = 11;
                xlWorkSheet.Columns["Z:Z"].ColumnWidth = 11;
                xlWorkSheet.Columns["AA:AA"].ColumnWidth = 11;
                xlWorkSheet.Columns["I:I"].ColumnWidth = 11;



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
                xlWorkSheet.Cells[s + 1, 1] = data1;

                for (j = 2; j <= 6; j++)
                {
                    data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                    xlWorkSheet.Cells[s + 1, j] = data;
                }

                for (j = 7; j <= 11; j++)
                {
                    data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                    xlWorkSheet.Cells[s + 1, j + 1] = data;
                }
                for (j = 12; j <= 17; j++)
                {
                    data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                    xlWorkSheet.Cells[s + 1, j + 2] = data;
                }

                for (j = 18; j <= 20; j++)
                {
                    data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                    xlWorkSheet.Cells[s + 1, j + 3] = data;
                }



                for (j = 21; j <= ds.Tables[0].Columns.Count - 1; j++)
                {
                    data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                    xlWorkSheet.Cells[s + 1, j + 4] = data;
                }
            }



            //sorting based on a coloumn:
            dynamic allDataRange = xlWorkSheet.get_Range("13:100");
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
            formatRange = xlWorkSheet.get_Range("A6", "AA6");
            formatRange.Interior.Color = System.Drawing.
            ColorTranslator.ToOle(System.Drawing.Color.Beige);
            //xlWorkSheet.Cells[1, 1] = "Red";


            Excel.Range formatRange1;
            formatRange1 = xlWorkSheet.get_Range("Y9", "AA9");
            formatRange1.Interior.Color = System.Drawing.
            ColorTranslator.ToOle(System.Drawing.Color.Beige);

           // Excel.Range formatRange2;
           // formatRange2 = xlWorkSheet.get_Range("A11", "AA11");
           // formatRange2.Interior.Color = System.Drawing.
          //  ColorTranslator.ToOle(System.Drawing.Color.Beige);

            Excel.Range formatRange3;
            formatRange3 = xlWorkSheet.get_Range("A10", "AA10");
            formatRange3.Interior.Color = System.Drawing.
            ColorTranslator.ToOle(System.Drawing.Color.Olive);


            //fontsize:
            formatRange.Font.Size = 15;


            // Excel.Range formatRange3;
            // formatRange3 = xlWorkSheet.get_Range("Y13","Y20");
            // formatRange3.Interior.Color = System.Drawing.
            // ColorTranslator.ToOle(System.Drawing.Color.LightGreen);

            




            // merging rows:
            xlWorkSheet.get_Range("N10", "S10").Merge(false);
            Excel.Range chartRange = xlWorkSheet.get_Range("N10", "S10");
            chartRange.FormulaR1C1 = "TEBYAN - FROM SUPPLIER (EX-WORKS)";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;

            chartRange.Font.Bold = true;





            xlWorkSheet.get_Range("V10", "W10").Merge(false);
            Excel.Range chartRange2 = xlWorkSheet.get_Range("V10", "W10");
            chartRange2.FormulaR1C1 = "VIRGIN";
            chartRange2.HorizontalAlignment = 3;
            chartRange2.VerticalAlignment = 3;

            chartRange2.Font.Bold = true;

            //formatRange.Font.Size = 15;



            // margin table:
            

                string data22 = (ds2.Tables[0].Rows[0].ItemArray[1].ToString());
                xlWorkSheet.Cells[12, 8] = data22 + "%";


            string data23 = (ds2.Tables[0].Rows[0].ItemArray[2].ToString());
            xlWorkSheet.Cells[12, 9] = data23 + "%";


          //  string data24 = (ds2.Tables[0].Rows[0].ItemArray[3].ToString());
          //  xlWorkSheet.Cells[12, 14] = data24 + "%";

            string data25 = (ds2.Tables[0].Rows[0].ItemArray[4].ToString());
            xlWorkSheet.Cells[12, 16] = data25 + "%";


           // string data26 = (ds2.Tables[0].Rows[0].ItemArray[5].ToString());
           // xlWorkSheet.Cells[12, 21] = data26 + "%";

            //v9:
            string data27 = (ds2.Tables[0].Rows[0].ItemArray[8].ToString());
            xlWorkSheet.Cells[10, 25] = "AED    " + data27;
            xlWorkSheet.Cells[10, 25].Font.Bold = true;
            //w9:
            string data28 = (ds2.Tables[0].Rows[0].ItemArray[6].ToString());
            xlWorkSheet.Cells[10, 26] = "QAR    "+ data28;
            xlWorkSheet.Cells[10, 26].Font.Bold = true;
            //x9:
            string data29 = (ds2.Tables[0].Rows[0].ItemArray[7].ToString());
            xlWorkSheet.Cells[10, 27] = "AED    "+ data29;
            xlWorkSheet.Cells[10, 27].Font.Bold = true;









            xlWorkBook.SaveAs("informations.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);



            MessageBox.Show("Excel file created , you can find the file C:\\Users\\.....\\Recent\\informations.xls");

            this.Hide();


        }





    }









}


