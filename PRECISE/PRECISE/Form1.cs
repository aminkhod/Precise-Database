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
    public partial class Form1 : Form
    {
        public Form1()
        {

           // this.AutoSize = true;
           // this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            InitializeComponent();

        }



        //public string conString = "Data Source=DESKTOP-SM32JMN;Initial Catalog=ForTesting;Integrated Security=True";
        public string conString = System.Configuration.ConfigurationManager.ConnectionStrings["connection_string"].ConnectionString;


        private void button1_Click(object sender, EventArgs e)
        {

            // Excell

            // Excel.Application xlApp = new Excel.Application();
            // Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\E56626\Desktop\Teddy\VS2012\Sandbox\sandbox_test - Copy - Copy.xlsx");
            // Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            //Excel.Range xlRange = xlWorksheet.UsedRange;
            // Excell

            SqlConnection con = new SqlConnection(conString);
            con.Open();
            if (con.State == System.Data.ConnectionState.Open)
            {
                //String q="insert into dbo.User(id,name)values('"+username.Text.ToString()+"','"+password.Text.ToString() +"')";

                var dataSet = new DataSet();
                String q = "Select *From dbo.Users where User_id = '" + username.Text.ToString() + "'AND Password='" + password.Text.ToString() + "'";

                SqlCommand cmd = new SqlCommand(q, con);
                SqlDataReader myreader = cmd.ExecuteReader();
                //cmd.ExecuteNonQuery();

                //MessageBox.Show(myreader.ToString());
                // var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };

                //dataAdapter.Fill(dataSet);
                // MessageBox.Show(dataSet.ToString());


                // String st = "INSERT INTO dbo.Users(id,name) values (@no, @name)";
                // SqlCommand cmd2 = new SqlCommand(st, con);
                // cmd2.Parameters.AddWithValue("@no", textBox5.Text);
                //cmd2.Parameters.AddWithValue("@name", textBox6.Text);
                // cmd2.ExecuteNonQuery();


                int count = 0;
                // MessageBox.Show(myreader.Read());
                string u = "";
                while (myreader.Read())
                {
                    u = myreader["Permission"].ToString();

                    count = count + 1;

                    if (count == 1)
                    //if (count == 1) & (@Permission==1)
                    {
                        // MessageBox.Show("Username and Password is Correct", "Confirmation Message", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        // MessageBox.Show("You can change informations", "Confirmation Message", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        String q1 = "Select Permission From dbo.Users where User_id = '" + username.Text.ToString() + "'AND Password='" + password.Text.ToString() + "'";
                        SqlCommand cmd2 = new SqlCommand(q1, con);
                        if (u == 1.ToString())
                        {
                            //Form2 frm = new Form2();
                            //frm.Show();
                            Form14 frm = new Form14();
                            frm.Show();
                            this.Hide();

                        }
                        else
                        // users:
                        {
                            Form3 frm = new Form3();
                            frm.Show();
                            this.Hide();
                            //this.Close();

                        }

                    }


                    else
                    {
                        MessageBox.Show("Username and Password is Not Correct .  You can only read information", "Error Message!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }
                    //  if (dataSet == null)
                    // {
                    // MessageBox.Show("connection done sucssesfully");
                    // }



                }
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void idbox_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void pas_Click(object sender, EventArgs e)
        {

        }

        private void password_TextChanged(object sender, EventArgs e)
        {

        }


    }
}
