using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Site_Manager
{
    public partial class frmUVouchersPrepared : Form
    {
        public frmUVouchersPrepared()
        {
            InitializeComponent();
        }
        OdbcCommand cmd;
        OdbcDataReader reader;
        private void GetVouchersPrepared()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn2.Open();
                cnn.Open();
                OdbcDataReader RDR;
                string strSQL1 = "SELECT    COUNT(*) FROM ODASMVoucher AS V INNER JOIN ODASPAccount AS PA ON V.AccountNo = PA.AccountNo  INNER JOIN  ODASMInstallment AS I ON V.VoucherNo = I.VoucherNo WHERE (v.voucherdate>='" + startdate.Value.ToString("yyyy/MM/dd") + "' AND v.voucherdate<='" + startdate.Value.ToString("yyyy/MM/dd") + "')";
                  
                cmd = new OdbcCommand(strSQL1, cnn2);
                int c = Convert.ToInt32(cmd.ExecuteScalar().ToString());
                string strSQL = "SELECT     V.VoucherNo, V.VoucherDate, I.ContractNo, I.Installment, PA.CompanyName, V.Amount,V.Printed FROM ODASMVoucher AS V INNER JOIN ODASPAccount AS PA ON V.AccountNo = PA.AccountNo  INNER JOIN  ODASMInstallment AS I ON V.VoucherNo = I.VoucherNo WHERE (v.voucherdate>='"+ startdate.Value.ToString ("yyyy/MM/dd") + "' AND v.voucherdate<='" + lastdate .Value.ToString ( "yyyy/MM/dd") + "')";
                  
                cmd = new OdbcCommand(strSQL, cnn);
                RDR = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();

                listView1.Columns.Add("VoucherNo", listView1.Width / 7);
                listView1.Columns.Add("VoucherDate", listView1.Width / 7);
                listView1.Columns.Add("ContractNo", listView1.Width / 7);
                listView1.Columns.Add("Installment", listView1.Width / 7);
                listView1.Columns.Add("CompanyName", listView1.Width / 7);
                listView1.Columns.Add("Installment", listView1.Width / 7);
                listView1.Columns.Add("Printed", listView1.Width / 7);
                progressBar1.Visible = true;
                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = c+1;

                if (RDR.HasRows)
                {


                    while (RDR.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(RDR["VoucherNo"].ToString());
                        if (RDR["VoucherDate"].ToString() != "")
                        {
                            lv3.SubItems.Add(Convert.ToDateTime(RDR["VoucherDate"].ToString()).ToString("yyyy/MM/dd"));

                        }
                        if (RDR["ContractNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["ContractNo"].ToString());

                        }
                        if (RDR["Installment"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["Installment"].ToString());
                        }
                        if (RDR["CompanyName"].ToString() != "")
                        {

                            lv3.SubItems.Add(RDR["CompanyName"].ToString());
                        }
                        if (RDR["Installment"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["Installment"].ToString());
                        } if (RDR["Printed"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["Printed"].ToString());
                        }
                      

                        listView1.Items.Add(lv3);
                      //  progressBar1.Value = progressBar1.Value + 1;
                    }
                    


                  
                    progressBar1.Value = 0;
                    progressBar1.Visible = false;


                }

                RDR.Close();
                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void frmUVouchersPrepared_Load(object sender, EventArgs e)
        {

        }

        private void lastdate_CloseUp(object sender, EventArgs e)
        {
            GetVouchersPrepared();
        }

        private void startdate_CloseUp(object sender, EventArgs e)
        {
            GetVouchersPrepared();
        }

        private void listView1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            ListView.CheckedListViewItemCollection  chked = listView1.CheckedItems;

            foreach (ListViewItem items in chked ){
                
                textBox2.Text = items.Text;
              

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            GeneralVariables vars = new GeneralVariables();
              
            vars.vourcherReport.currentRecord = textBox2.Text;
            vars.vourcherReport.ShowDialog();
                
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try { 
                 if(textBox2 .Text ==""){
                     MessageBox.Show("Please select a record","Information Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                 }else {

            GeneralVariables GeneralVariables = new GeneralVariables();
            OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
            cnn2.Open();
            cmd = new OdbcCommand("UPDATE ODASMVoucherItem SET ItemName ='"+textBox1 .Text +"'where VoucherNo = '" + textBox2.Text + "'", cnn2);
          
              cmd.ExecuteNonQuery();
                cnn2.Close();
                MessageBox.Show("Vourcher Details No '" + textBox2.Text + " ' Update Success");
                 }
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();
            OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
            cnn2.Open();
            cmd = new OdbcCommand("SELECT *FROM ODASMVoucherItem where VoucherNo = '" + textBox2.Text + "'", cnn2);
            reader = cmd.ExecuteReader();
            reader.Read();
            textBox1.Text = reader["ItemName"].ToString();
            reader.Close();
            cnn2.Close();
           
        }
    }
}
