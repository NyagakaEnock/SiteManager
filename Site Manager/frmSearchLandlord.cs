using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.Odbc;
namespace Site_Manager
{
    public partial class frmSearchLandlord : Form
    {
        public frmSearchLandlord()
        {
            InitializeComponent();
        }
        string sql;
        OdbcDataReader reader;


        private void loadLandLords()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();

                listView1.Columns.Clear();
                listView1.Items.Clear();
                listView1.Columns.Add("Landlord No", listView1.Width / 3);
                listView1.Columns.Add("Landlord Name", listView1.Width / 2);

                string sql = "SELECT * FROM ODASPAccount Where Status = 'A' AND AccountType = 'LLORD' oRDER BY AccountNo";

                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();

                listView1.Items.Clear();
                  while (reader.Read())
                    {

                        ListViewItem lv = new ListViewItem(reader["AccountNo"].ToString());
                       
                        lv.SubItems.Add(reader["CompanyName"].ToString());
                       
                        
                        listView1.Items.Add(lv);




                    
                }
                reader.Close();
                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void frmSearchLandlord_Load(object sender, EventArgs e)
        {
            loadLandLords();
        }

        private void listView1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
          // GeneralVariables GeneralVariables = new GeneralVariables();
           ListView.CheckedListViewItemCollection checkeditems = listView1.CheckedItems;
            foreach (ListViewItem items in checkeditems  ){
                GeneralVariables GeneralVariables = new GeneralVariables();
               textBox2.Text = items.Text;
             
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
               
             
                listView1.Columns.Clear();
                listView1.Items.Clear();
                listView1.Columns.Add("Landlord No", listView1.Width / 3);
                listView1.Columns.Add("Landlord Name", listView1.Width / 2);
                
                    sql = "SELECT * FROM ODASPAccount Where Status = 'A' AND AccountType = 'LLORD' AND AccountNo LIKE '%" + textBox1.Text + "%' or CompanyName LIKE '%" + textBox1.Text + "%' oRDER BY AccountNo";

                
                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
            

                listView1.Items.Clear();
                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv = new ListViewItem(reader["AccountNo"].ToString());
                        lv.SubItems.Add(reader["CompanyName"].ToString());
                        listView1.Items.Add(lv);

                    }

                }
                else {
                    ListViewItem lv = new ListViewItem("NO data Found");
                }
                reader.Close();
                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void listView1_MouseClick(object sender, MouseEventArgs e)
        {
           // loadLandLords();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();
             
            GeneralVariables.frmRStatement.strAccountNo = textBox2.Text;
            GeneralVariables.frmRStatement.ShowDialog();
        }

    }
}
