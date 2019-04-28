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
    public partial class Form16 : Form
    {
        public Form16()
        {
            InitializeComponent();
        }
        public string conString = System.Configuration.ConfigurationManager.ConnectionStrings["connection_string"].ConnectionString;
        public int index_sup, index_cat;
        private void Form16_Load(object sender, EventArgs e)
        {

        }

        public static List<T> RemoveDuplicates<T>(List<T> items)
        {
            return (from s in items select s).Distinct().ToList();
        }


        public List<string> Super_category_list = new List<string>();
        public List<string> Category_list = new List<string>();
        public List<string> Product_list = new List<string>();


        public List<string> Category_list_itssSuper = new List<string>();
        public List<string> Product_list_itsCategory = new List<string>();
        public string supp;
        public int index_sup_new;
        public int index_product;
        public int index_p;
        public int index_p_2;
        public int index_p_n;
        public string ssss;
        public int superG_index_in_tree;
        public int index_group_in_tree;

        public List<string> help1 = new List<string>();
        public List<string> help2 = new List<string>();

        public List<TreeNode> checkedNodes = new List<TreeNode>();




        int NumberOfClick = 0;
        private void button1_Click(object sender, EventArgs e)
        {

            NumberOfClick = NumberOfClick + 1;

            switch (NumberOfClick)
            {
                case 1:
                    {
                        // this is the first click
                        // . . .


                        //  treeView1.Nodes.Add("3DOODLER CREATE");

                        //  treeView1.Nodes[0].Nodes.Add("Create Pen Sets");
                        //  treeView1.Nodes[0].Nodes.Add("Create Accessories");
                        //  treeView1.Nodes[0].Nodes.Add("Create Project Kits");
                        //   treeView1.Nodes[0].Nodes.Add("Consumables");
                        SqlConnection con16 = new SqlConnection(conString);
                        con16.Open();
                        string Seqchildc = "Select *From dbo.Products_New";

                        SqlCommand cmd9 = new SqlCommand(Seqchildc, con16);

                        SqlDataAdapter dachildmnuc = new SqlDataAdapter(Seqchildc, con16);


                        SqlDataReader myreader = cmd9.ExecuteReader();
                        while (myreader.Read())
                        {
                            Product_list.Add(myreader["product"].ToString());
                            Category_list.Add(myreader["Group_Name"].ToString());
                            Super_category_list.Add(myreader["SubGroup_Name"].ToString());

                            Category_list_itssSuper.Add(myreader["SubGroup_Name"].ToString());
                            Product_list_itsCategory.Add(myreader["Group_Name"].ToString());

                        }


                        List<string> Super_category_list_new = RemoveDuplicates(Super_category_list);

                        //List<string> Product_list_new= RemoveDuplicates(Product_list);
                        List<string> Category_list_new = RemoveDuplicates(Category_list);



                        // childNode = treeView1.Nodes.Add(dr["FRM_NAME"].ToString());


                        // PopulateTreeView(Convert.ToInt32(dr["MNUSUBMENU"].ToString()), childNode);
                        foreach (var el in Super_category_list_new)
                        {

                            treeView1.Nodes.Add(el);
                        }



                        foreach (var el in Category_list_new)
                        {
                            foreach (var cate in Category_list)
                            {
                                if (el == cate)
                                {
                                    int index_cat = Category_list.FindIndex(a => a.Contains(cate));
                                    index_sup = index_cat;
                                    supp = Super_category_list[index_sup];

                                }


                            }
                            foreach (var j in Super_category_list_new)
                            {
                                if (j == supp)
                                {
                                    index_sup_new = Super_category_list_new.FindIndex(a => a.Contains(j));
                                }
                            }

                            treeView1.Nodes[index_sup_new].Nodes.Add(el);
                            //// if (index_sup_new==0)
                            //  {
                            //      help1.Add(el);
                            //   }
                            //   else
                            //   {
                            ////       help2.Add(el);
                            //    }


                        }



                        foreach (TreeNode tn in treeView1.Nodes)
                        {
                            int superG_index_in_tree = treeView1.Nodes.IndexOf(tn);
                            string super_in_tree = tn.ToString();
                            // MessageBox.Show(super_in_tree);
                            TreeNode gg = treeView1.Nodes[superG_index_in_tree];
                            foreach (TreeNode tnn in gg.Nodes)
                            {

                                int index_group_in_tree = gg.Nodes.IndexOf(tnn);
                                string group_in_tree = tnn.ToString();
                                help1.Add(index_group_in_tree.ToString());
                                help2.Add(group_in_tree);



                            }
                        }




                        foreach (var pp in Product_list)
                        {
                            index_product = Product_list.FindIndex(a => a.Contains(pp));
                            string cc = Category_list[index_product];
                            string ss = Super_category_list[index_product];
                            foreach (var ee in Category_list_new)
                            {
                                if (cc == ee)
                                {

                                    index_p = Category_list_new.FindIndex(a => a.Contains(ee));
                                }
                            }
                            foreach (var q in help2)
                            {

                                if (q == ("TreeNode: " + cc))
                                {
                                    //MessageBox.Show(q);
                                    // MessageBox.Show("TreeNode:" + cc);

                                    int iii = help2.FindIndex(a => a.Contains(q));
                                    string vv = help1[iii];
                                    index_p_n = Int32.Parse(vv);
                                }
                            }

                            foreach (var ssss in Super_category_list_new)
                            {
                                if (ssss == ss)
                                {
                                    index_p_2 = Super_category_list_new.FindIndex(a => a.Contains(ssss));
                                }


                            }

                            // if (index_p_2 == 0)
                            //// {
                            //      int index_p_n = help1.FindIndex(a => a.Contains(cc));

                            ////    }
                            //    else
                            //    {
                            //        int index_p_n = help2.FindIndex(a => a.Contains(cc));
                            //    }
                            //first:super    second: category            
                            treeView1.Nodes[index_p_2].Nodes[index_p_n].Nodes.Add(pp);
                        }




                    }

                    break;
                case 2:
                    // this is the second click
                    // . . .
                    break;
            }




        }
        
        // order button:
        private void button2_Click_1(object sender, EventArgs e)
        {
            {


                // the bellow line : indicates supergroups
                foreach (TreeNode node in treeView1.Nodes)
                {
                    if (node.Checked)
                    {
                        int superG_index_in_tree_forcheck = treeView1.Nodes.IndexOf(node);
                        string super_in_tree_forcheck = node.ToString();
                        foreach (TreeNode tnn in node.Nodes)
                        {
                            foreach (TreeNode tnnn in tnn.Nodes)
                            {
                                checkedNodes.Add(tnnn);
                            }
                        }

                    }



                }
                foreach (TreeNode node in treeView1.Nodes)
                {
                    foreach (TreeNode nn in node.Nodes)
                    {
                        if (nn.Checked)
                        {
                            foreach (TreeNode nnn in nn.Nodes)
                                checkedNodes.Add(nnn);
                        }
                    }
                }

                foreach (TreeNode node in treeView1.Nodes)
                {
                    foreach (TreeNode nn in node.Nodes)
                    {
                        foreach (TreeNode nnn in nn.Nodes)
                        {
                            if (nnn.Checked)
                            {
                                checkedNodes.Add(nnn);
                            }
                        }

                    }
                }








                SqlConnection con8 = new SqlConnection(conString);
                con8.Open();
                string sqlTrunc = "DELETE FROM dbo.Products_order";
                SqlCommand cmd = new SqlCommand(sqlTrunc, con8);
                cmd.ExecuteNonQuery();
                con8.Close();




                foreach (var mm in checkedNodes)
                {
                    SqlConnection con7 = new SqlConnection(conString);
                    con7.Open();
                    SqlCommand cmd7 = new SqlCommand("INSERT INTO   dbo.Products_order(productsID, SubGroup_Name, Group_Name, product, version, Unit_cost_USD, Unit_cost_AED, Freight_customs, Financing, PRECISE_Landed_Cost, PRECISE_Margin_AED, PRECISE_Margin_Cost_Ratio, T_Cost_Price_USD, T_Cost_Price_QAR, T_Freight_customs_10per_add_5per, T_Landed_Cost_QAR, T_Margin_QAR, T_Margin_Per, V_Retailer_Cost_Price_QAR, V_Retailer_Margin_QAR, V_Retailer_Margin_per, UAE_RRP_AED, QAR_RRP_AED, QAR_RRP_QAR) Select * from dbo.Products_New where ('TreeNode: '+ product) = '" + mm + "'  ", con7);



                    cmd7.ExecuteNonQuery();
                    con7.Close();


                    
                }

                Form17 frm17 = new Form17();
                // frm17.myArray = checkedNodes;
                frm17.Show();



            }

            this.Hide();

        }

        
        
        



    }
}
