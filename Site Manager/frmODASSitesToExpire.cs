using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Site_Manager
{
    public partial class frmODASSitesToExpire : Form
    {
        public frmODASSitesToExpire()
        {
            InitializeComponent();
        }
        Excel.Application objApp;
        Excel._Workbook objBook;
        OdbcCommand cmd;
        OdbcDataReader reader;
        System.Data.DataTable dt;
        OdbcDataAdapter da;
        public String strReport;

        Microsoft.Office.Interop.Excel.Application excel; 
        private void showALLPlotsToExpire()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn2.Open();
                cnn.Open();
                OdbcDataReader RDR;
                string sql1 = "SELECT COUNT(*)  FROM ODASPPlot, ODASPAccount,ODASMLeaseAgreement WHERE  ODASMLeaseAgreement.PlotNo=ODASPPlot.PlotNo AND ODASPPLot.ExpiryDate <= '" + Convert.ToDateTime(txtLastDate.Text).ToString("yyyy/MM/dd") + "' and ODASPPlot.AccountNo = ODASPAccount.AccountNo  AND (ODASMLeaseAgreement.terminated='N' OR ODASMLeaseAgreement.Terminated IS NULL)";
                cmd = new OdbcCommand(sql1, cnn2);
                string sql;
                int c = Convert.ToInt32(cmd.ExecuteScalar().ToString());
                if(strReport ==""){
                    sql = "SELECT ODASPPlot.*,ODASPAccount.COmpanyName  FROM ODASPPlot, ODASPAccount,ODASMLeaseAgreement WHERE ODASMLeaseAgreement.PlotNo=ODASPPlot.PlotNo AND ODASPPLot.ExpiryDate >= '" + Convert.ToDateTime(txtStartDate.Value).ToString("yyyy/MM/dd") + "' and ODASPPLot.ExpiryDate <= '" + Convert.ToDateTime(txtLastDate.Value).ToString("yyyy/MM/dd") + "' and ODASPPlot.AccountNo = ODASPAccount.AccountNo and (ODASMLeaseAgreement.terminated='N' OR ODASMLeaseAgreement.Terminated IS NULL)";
                }
                else if (strReport == "ExpiredNotRenewed")
                {
                    sql = "SELECT ODASPPlot.*,ODASPAccount.COmpanyName  FROM ODASPPlot, ODASPAccount,ODASMLeaseAgreement WHERE  ODASMLeaseAgreement.PlotNo=ODASPPlot.PlotNo AND ODASPPLot.ExpiryDate <= '" + Convert.ToDateTime(txtLastDate.Text).ToString ("yyyy/MM/dd") + "' and ODASPPlot.AccountNo = ODASPAccount.AccountNo  AND (ODASMLeaseAgreement.terminated='N' OR ODASMLeaseAgreement.Terminated IS NULL) and (ODASMLeaseAgreement.Renewal='0' OR ODASMLeaseAgreement.Renewal IS NULL)";

                }
                else {
                    sql = "SELECT ODASPPlot.*,ODASPAccount.COmpanyName  FROM ODASPPlot, ODASPAccount,ODASMLeaseAgreement WHERE  ODASMLeaseAgreement.PlotNo=ODASPPlot.PlotNo AND ODASPPLot.ExpiryDate <= '" + Convert.ToDateTime(txtLastDate.Text).ToString ("yyyy/MM/dd") + "' and ODASPPlot.AccountNo = ODASPAccount.AccountNo  AND (ODASMLeaseAgreement.terminated='N' OR ODASMLeaseAgreement.Terminated IS NULL)";
               
                
                }
                
                cmd = new OdbcCommand(sql, cnn);
                RDR = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();

                listView1.Columns.Add("Plot No", listView1.Width / 5);
                listView1.Columns.Add("Company Name", listView1.Width / 5);
                listView1.Columns.Add("Physical Location", listView1.Width / 5);
                listView1.Columns.Add("DOC", listView1.Width / 5);
                listView1.Columns.Add("Expiry Date", listView1.Width / 5);
               
                progressBar1.Visible = true;
                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = c;

                if (RDR.HasRows)
                {


                    while (RDR.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(RDR["PlotNo"].ToString());
                        if (RDR["CompanyName"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["CompanyName"].ToString());

                        }
                        if (RDR["PhysicalLocation"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["PhysicalLocation"].ToString());

                        }
                        if (RDR["CommencementDate"].ToString() != "")
                        {
                            lv3.SubItems.Add(Convert.ToDateTime(RDR["CommencementDate"].ToString()).ToString("yyyy/MM/dd"));
                        }
                        if (RDR["expirydate"].ToString() != "")
                        {

                            lv3.SubItems.Add(Convert.ToDateTime(RDR["expirydate"].ToString()).ToString("yyyy/MM/dd"));
                        }
                   

                        listView1.Items.Add(lv3);
                        progressBar1.Value = progressBar1.Value + 1;
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
        private void frmODASSitesToExpire_Load(object sender, EventArgs e)
        {
            if (strReport == "")
            {

                this.Text = "Searching For Sites/Billboards To Expire - Within A date Range";
               
            }
            else {
                txtStartDate.Visible = false;
                this.Text = "Searching For Sites/Billboards To Expire - As At A Single Date";
            }
            showALLPlotsToExpire();
        }

        private void txtLastDate_CloseUp(object sender, EventArgs e)
        {
            showALLPlotsToExpire();
        }

        private void txtStartDate_CloseUp(object sender, EventArgs e)
        {
            showALLPlotsToExpire();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {



                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                int i = 2;

                int j = 1;
                int x = 0;
                int c = listView1.Items.Count - 1;
                progressBar1.Visible = true;

                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = c + 1;

                foreach (ListViewItem comp in listView1.Items)
                {

                    xlWorkSheet.Cells[i, j] = comp.Text.ToString();
                    for (int y = 1; y <= listView1.Columns.Count; y++)
                    {
                        xlWorkSheet.Cells[j, y] = listView1.Columns[y - 1].Text;
                    }
                    foreach (ListViewItem.ListViewSubItem drv in comp.SubItems)
                    {

                        xlWorkSheet.Cells[i, j] = drv.Text.ToString();
                        xlWorkSheet.Cells[listView1.Items.Count + 2, listView1.Columns.Count] = label4.Text;

                        j++;
                    }

                    j = 1;

                    i++;

                    xlWorkSheet.Cells[listView1.Items.Count + 2, listView1.Columns.Count - 1] = "Totals";

                    progressBar1.Value = progressBar1.Value + 1;
                }


                xlWorkBook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\ Sites Billboards To Expire .xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                MessageBox.Show("File stored in " + System.Windows.Forms.Application.StartupPath + "\\ Sites Billboards To Expire .xls", "Excel File Created Successfully", MessageBoxButtons.OK, MessageBoxIcon.Information);

                System.Diagnostics.Process.Start(System.Windows.Forms.Application.StartupPath + "\\ Sites Billboards To Expire .xls");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            GeneralVariables vars = new GeneralVariables();
            vars.Sites_to_Expire.strReport = strReport;
            vars.Sites_to_Expire.txtLastDate = txtLastDate.Text;
            vars.Sites_to_Expire.txtStartDate = txtStartDate.Text;
          //  vars.Sites_to_Expire 
            vars.Sites_to_Expire.ShowDialog();
        }
    }
}
