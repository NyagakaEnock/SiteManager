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
    public partial class frmODASMContractTermination : Form
    {
        public frmODASMContractTermination()
        {
            InitializeComponent();
        }
        OdbcCommand cmd;
        OdbcDataReader reader;
        public string CurrentUserName;
        private void getContractToTerminated()
        {
            try
            {

                listView1.Columns.Clear();
                listView1.Items.Clear();
                listView1.Columns.Add("COntract No", listView1.Width / 9);
                listView1.Columns.Add("Plot No", listView1.Width / 9);
                listView1.Columns.Add("COmpanyName", listView1.Width / 5);
                listView1.Columns.Add("COmmencementdate", listView1.Width / 9);
                listView1.Columns.Add("ExpiryDate", listView1.Width / 9);
                listView1.Columns.Add("Lease Duration", listView1.Width / 9);
                listView1.Columns.Add("Terminated", listView1.Width / 9);
                listView1.Columns.Add("TerminationDate", listView1.Width / 9);
                listView1.Columns.Add("Reason", listView1.Width / 9);
                ListView.ListViewItemCollection items = listView1.Items;
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
              
                cnn.Open();

                cmd = new OdbcCommand("SELECT LA.COntractNo,LA.PlotNo,A.COmpanyName,COmmencementdate,ExpiryDate,LeaseDuration,Terminated,TerminationDate,LA.Reason FROM ODASMLeaseAgreement LA INNER JOIN ODASPAccount A ON A.AccountNo=LA.AccountNo ", cnn);

                reader = cmd.ExecuteReader();

         

                if (reader.Read())
                {


                    while (reader.Read())
                    {
                        String dat = reader["COntractNo"].ToString();
                        ListViewItem lv3 = new ListViewItem(reader["COntractNo"].ToString());
                       
                            lv3.SubItems.Add(reader["PlotNo"].ToString());
                        

                       
                            lv3.SubItems.Add(reader["COmpanyName"].ToString());
                        
                        
                            lv3.SubItems.Add(Convert .ToDateTime ( reader["COmmencementdate"].ToString()).ToString ("MM/dd/yyyy"));
                        
                        
                            lv3.SubItems.Add(Convert .ToDateTime (reader["ExpiryDate"].ToString()).ToString ("MM/dd/yyyy"));
                        
                       
                            lv3.SubItems.Add(reader["LeaseDuration"].ToString());
                       

                        
                            lv3.SubItems.Add(reader["Terminated"].ToString());
                        

                        
                            lv3.SubItems.Add(reader["TerminationDate"].ToString());
                        
                            lv3.SubItems.Add(reader["Reason"].ToString());
                        
                        listView1.Items.Add(lv3);

                       

                    }
                   

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void frmODASMContractTermination_Load(object sender, EventArgs e)
        {
            getContractToTerminated();
        }

        private void listView1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            ListView.CheckedListViewItemCollection checkeditems = listView1.CheckedItems;
            foreach (ListViewItem items in checkeditems  ){
                txtContractNo.Text = items.Text;
                txtPlotNo.Text = items.SubItems [1].Text;
                search();
            }
        }
        private void search() {
            try
            {

                listView1.Columns.Clear();
                listView1.Items.Clear();
                listView1.Columns.Add("COntract No", listView1.Width / 9);
                listView1.Columns.Add("Plot No", listView1.Width / 9);
                listView1.Columns.Add("COmpanyName", listView1.Width / 5);
                listView1.Columns.Add("COmmencementdate", listView1.Width / 9);
                listView1.Columns.Add("ExpiryDate", listView1.Width / 9);
                listView1.Columns.Add("Lease Duration", listView1.Width / 9);
                listView1.Columns.Add("Terminated", listView1.Width / 9);
                listView1.Columns.Add("TerminationDate", listView1.Width / 9);
                listView1.Columns.Add("Reason", listView1.Width / 9);
                ListView.ListViewItemCollection items = listView1.Items;
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();

                cmd = new OdbcCommand("SELECT LA.COntractNo,LA.PlotNo,A.COmpanyName,COmmencementdate,ExpiryDate,LeaseDuration,Terminated,TerminationDate,LA.Reason FROM ODASMLeaseAgreement LA INNER JOIN ODASPAccount A ON A.AccountNo=LA.AccountNo WHERE PlotNo LIKE '%"+txtPlotNo .Text +"%'", cnn);

                reader = cmd.ExecuteReader();



                if (reader.Read())
                {


                    while (reader.Read())
                    {
                        String dat = reader["COntractNo"].ToString();
                        ListViewItem lv3 = new ListViewItem(reader["COntractNo"].ToString());

                        lv3.SubItems.Add(reader["PlotNo"].ToString());



                        lv3.SubItems.Add(reader["COmpanyName"].ToString());


                        lv3.SubItems.Add(Convert.ToDateTime(reader["COmmencementdate"].ToString()).ToString("MM/dd/yyyy"));


                        lv3.SubItems.Add(Convert.ToDateTime(reader["ExpiryDate"].ToString()).ToString("MM/dd/yyyy"));


                        lv3.SubItems.Add(reader["LeaseDuration"].ToString());



                        lv3.SubItems.Add(reader["Terminated"].ToString());



                        lv3.SubItems.Add(reader["TerminationDate"].ToString());
                        
                        lv3.SubItems.Add(reader["Reason"].ToString());

                        listView1.Items.Add(lv3);



                    }


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        
        }
        private void txtPlotNo_TextChanged(object sender, EventArgs e)
        {
            search();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try{
                if (txtContractNo.Text == "")
                {
                    MessageBox.Show("Select a contract to terminate", "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);

                } 
                else
                {
                    if(MessageBox.Show("Are you sure you want to Terminate this Contract? ", "Confirmation Required", MessageBoxButtons.YesNo , MessageBoxIcon.Question )==DialogResult .Yes )
                    {
                
                
                    GeneralVariables vars = new GeneralVariables();
                    OdbcConnection con = new OdbcConnection(vars.SQLstr);
                    con.Open();
                    cmd = new OdbcCommand("UPDATE ODASMLeaseAgreement SET Terminated='Y',TerminationDate='" + DateTime.Today.ToString("yyyy/MM/dd") + "',TerminatedBy='" + CurrentUserName + "',Reason='" + txtreason.Text + "' WHERE ContractNo LIKE '" + txtContractNo.Text + "'", con);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Contract Terminated Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    con.Close();
                    getContractToTerminated();
                    }
                }
            }catch (Exception ex){
            MessageBox .Show (ex.ToString ());
            
            }
        }
    }
}
