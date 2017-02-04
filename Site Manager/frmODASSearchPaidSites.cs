using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
namespace Site_Manager
{
    public partial class frmODASSearchPaidSites : Form
    {
        public frmODASSearchPaidSites()
        {
            InitializeComponent();
        }
        Excel.Application objApp;
        Excel._Workbook objBook;
        OdbcCommand cmd;
        OdbcDataReader reader;
        System.Data.DataTable dt;
        OdbcDataAdapter da;
        Microsoft.Office.Interop.Excel.Application excel; 
        private void showALLRentPAIDThisPeriod()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn2.Open();
                cnn.Open();
                OdbcDataReader RDR;
                string sql1 = "SELECT COUNT(*) FROM ODASMInstallment,ODASPPlot,ODASPAccount Where ODASMInstallment.PaymentDate >= '" + Convert.ToDateTime(txtStartDate.Value).ToString("yyyy/MM/dd") + "' and (ODASMInstallment.PaymentFlag = 'Y' or ODASMInstallment.PaymentFlag = 'P') AND ODASMInstallment.PaymentDate <= '" + Convert.ToDateTime(txtLastDate.Value).ToString("yyyy/MM/dd") + "' and ODASPPlot.PlotNo = ODASMInstallment.ContractNo and ODASPPlot.AccountNo = ODASPAccount.AccountNo ";

                cmd = new OdbcCommand(sql1, cnn2);
                int c = Convert.ToInt32(cmd.ExecuteScalar().ToString());
                string sql = "SELECT * FROM ODASMInstallment,ODASPPlot,ODASPAccount Where ODASMInstallment.PaymentDate >= '" + Convert.ToDateTime(txtStartDate.Value).ToString("yyyy/MM/dd") + "' and (ODASMInstallment.PaymentFlag = 'Y' or ODASMInstallment.PaymentFlag = 'P') AND ODASMInstallment.PaymentDate <= '" + Convert.ToDateTime(txtLastDate.Value).ToString("yyyy/MM/dd") + "' and ODASPPlot.PlotNo = ODASMInstallment.ContractNo and ODASPPlot.AccountNo = ODASPAccount.AccountNo ";

                cmd = new OdbcCommand(sql, cnn);
                RDR = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();

                listView1.Columns.Add("Mast No", listView1.Width / 8);
                listView1.Columns.Add("Location", listView1.Width / 8);
                listView1.Columns.Add("LandLord", listView1.Width / 8);
                listView1.Columns.Add("Starting", listView1.Width / 8);
                listView1.Columns.Add("Ending", listView1.Width / 8);
                listView1.Columns.Add("Installment", listView1.Width / 8);
                listView1.Columns.Add("Contract No", listView1.Width / 8);
                listView1.Columns.Add("Payment Date", listView1.Width / 8);
                listView1.Columns.Add("Amount Paid", listView1.Width / 8);
                progressBar1.Visible = true;
                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = c;

                if (RDR.HasRows)
                {


                    while (RDR.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(RDR["PlotNo"].ToString());
                        if (RDR["PhysicalLocation"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["PhysicalLocation"].ToString());

                        }
                        if (RDR["CompanyName"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["CompanyName"].ToString());

                        }
                        if (RDR["CommencementDate"].ToString() != "")
                        {
                            lv3.SubItems.Add(Convert.ToDateTime(RDR["CommencementDate"].ToString()).ToString("yyyy/MM/dd"));
                        }
                        if (RDR["expirydate"].ToString() != "")
                        {

                            lv3.SubItems.Add(Convert.ToDateTime(RDR["expirydate"].ToString()).ToString("yyyy/MM/dd"));
                        }
                        if (RDR["Installment"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["Installment"].ToString());
                        } if (RDR["ContractNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["ContractNo"].ToString());
                        }
                        if (RDR["PaymentDueDate"].ToString() != "")
                        {
                            lv3.SubItems.Add(Convert.ToDateTime(RDR["PaymentDueDate"].ToString()).ToString("yyyy/MM/dd"));
                        }
                        if (RDR["AmountPaid"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["AmountPaid"].ToString());
                        }

                        listView1.Items.Add(lv3);
                        progressBar1.Value = progressBar1.Value + 1;
                    }
                    double ColumnSum = 0.0;

                    for (int i = 0; i < this.listView1.Items.Count; i++)
                    {
                        ColumnSum += Convert.ToDouble(listView1.Items[i].SubItems[8].Text);
                    }


                    label4.Text = ColumnSum.ToString();
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
        private void frmODASSearchPaidSites_Load(object sender, EventArgs e)
        {

        }
      

        private void btnPrint_Click(object sender, EventArgs e)
        {

            GeneralVariables vars = new GeneralVariables();
            vars.rptRentDue.txtStartDate = txtStartDate.Text;
            vars.rptRentDue.txtLastDate = txtLastDate.Text;
            vars.rptRentDue.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
        }

        private void txtStartDate_CloseUp(object sender, EventArgs e)
        {
            showALLRentPAIDThisPeriod();
        }

        private void txtLastDate_ValueChanged(object sender, EventArgs e)
        {
            showALLRentPAIDThisPeriod();
        }

        private void button2_Click_1(object sender, EventArgs e)
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


                xlWorkBook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Sites Paid.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                MessageBox.Show("File stored in " + System.Windows.Forms.Application.StartupPath + "\\Sites Paid.xls", "Excel File Created Successfully", MessageBoxButtons.OK, MessageBoxIcon.Information);

                System.Diagnostics.Process.Start(System.Windows.Forms.Application.StartupPath + "\\Sites Paid.xls");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

      

    }
}
