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
    public partial class frmODASSearchSitesNotPaid : Form
    {

        public frmODASSearchSitesNotPaid()
        {
            InitializeComponent();
        }
        public string strReport;
        Excel.Application objApp;
        Excel._Workbook objBook;
        OdbcCommand cmd;
        OdbcDataReader reader;
        System.Data.DataTable dt;
        OdbcDataAdapter da;
        Microsoft.Office.Interop.Excel.Application excel; 

        private void showALLRentNOTPAIDThisPeriod()
        {
            try
            {
                
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn2.Open();
                cnn.Open();
                OdbcDataReader RDR;
                string sql1 = "Select COUNT(*) from ODASMInstallment, ODASPAccount,ODASPPlot Where ODASPPlot.PlotNo=ODASMInstallment.PlotNo AND (ODASMInstallment.Requisitioned = 'N' or ODASMInstallment.Requisitioned is null or ODASMInstallment.Requisitioned = 'Y' ) and ODASPAccount.AccountNo = ODASPPlot.AccountNo  AND (ODASMInstallment.PaymentDueDate>='" + Convert.ToDateTime(txtStartDate.Value).ToString("yyyy/MM/dd") + "' AND ODASMInstallment.PaymentDueDate<='" + Convert.ToDateTime(txtLastDate.Value).ToString("yyyy/MM/dd") + "') AND ODASMInstallment.Balance>0 AND ODASMInstallment.ContractNo IN (Select ContractNo From ODASMLeaseAgreement where (Terminated='N' OR Terminated IS NULL));";
                cmd = new OdbcCommand(sql1, cnn2);
                int c =Convert .ToInt32( cmd.ExecuteScalar().ToString());
                string sql = "Select * from ODASMInstallment, ODASPAccount,ODASPPlot Where ODASPPlot.PlotNo=ODASMInstallment.PlotNo AND (ODASMInstallment.Requisitioned = 'N' or ODASMInstallment.Requisitioned is null or ODASMInstallment.Requisitioned = 'Y' ) and ODASPAccount.AccountNo = ODASPPlot.AccountNo  AND (ODASMInstallment.PaymentDueDate>='" +Convert .ToDateTime ( txtStartDate.Value).ToString ( "yyyy/MM/dd") + "' AND ODASMInstallment.PaymentDueDate<='" +Convert .ToDateTime ( txtLastDate.Value).ToString ("yyyy/MM/dd") + "') AND ODASMInstallment.Balance>0 AND ODASMInstallment.ContractNo IN (Select ContractNo From ODASMLeaseAgreement where (Terminated='N' OR Terminated IS NULL));";
              
                cmd = new OdbcCommand(sql, cnn);
                RDR = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();
          
                listView1.Columns.Add("Mast No", listView1.Width / 8);
                listView1.Columns.Add("Location", listView1.Width / 8);
                listView1.Columns.Add("LandLord", listView1.Width / 8);
                listView1.Columns.Add("Starting", listView1.Width / 8);
                listView1.Columns.Add("Ending", listView1.Width / 8);
                listView1.Columns.Add("Sides", listView1.Width / 8);
                listView1.Columns.Add("Payment Due Date", listView1.Width / 8);
                listView1.Columns.Add("AmountDue", listView1.Width / 8);
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
                        if (RDR["NoofSites"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["NoofSites"].ToString());
                        }
                        if (RDR["PaymentDueDate"].ToString() != "")
                        {
                            lv3.SubItems.Add(Convert.ToDateTime(RDR["PaymentDueDate"].ToString()).ToString("yyyy/MM/dd"));
                        }
                        if (RDR["PaymentDue"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["PaymentDue"].ToString());
                        }

                        listView1.Items.Add(lv3);                        
                        progressBar1.Value = progressBar1.Value + 1;
                    }
                    double ColumnSum = 0.0;
                      
                    for (int i = 0; i < this.listView1.Items.Count; i++)
                    {
                        ColumnSum += Convert.ToDouble(listView1.Items[i].SubItems[7].Text);
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
        private void showALLRentPendingPaymentAsAtASingleDate()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn2.Open();
                cnn.Open();
                OdbcDataReader RDR;
                String sql1 = "SELECT COUNT(*) FROM ODASMInstallment,ODASPPlot,ODASPAccount,ODASMLeaseAgreement Where ODASMLeaseAgreement.PlotNo=ODASMInstallment.PlotNo AND ODASMLeaseAgreement.ContractNo=ODASMInstallment.ContractNo AND (ODASMLeaseAgreement.Terminated='N' OR ODASMLeaseAgreement.Terminated IS NULL) AND (ODASMInstallment.PaymentDueDate <= '" +Convert .ToDateTime ( txtLastDate.Value).ToString ("yyyy/MM/dd") + "') and ODASMInstallment.PaymentFlag = 'N' and ODASMInstallment.Balance > 0 and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo;";
                cmd = new OdbcCommand(sql1, cnn2);
               int c = Convert.ToInt32(cmd.ExecuteScalar().ToString());
               String sql = "SELECT ODASMInstallment. * FROM ODASMInstallment,ODASPPlot,ODASPAccount,ODASMLeaseAgreement Where ODASMLeaseAgreement.PlotNo=ODASMInstallment.PlotNo AND ODASMLeaseAgreement.ContractNo=ODASMInstallment.ContractNo AND (ODASMLeaseAgreement.Terminated='N' OR ODASMLeaseAgreement.Terminated IS NULL) AND (ODASMInstallment.PaymentDueDate <= '" +Convert .ToDateTime ( txtLastDate.Value).ToString ("yyyy/MM/dd") + "') and ODASMInstallment.PaymentFlag = 'N' and ODASMInstallment.Balance > 0 and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo ORDER BY CompanyName";
               
                cmd = new OdbcCommand(sql, cnn);
                RDR = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();

                listView1.Columns.Add("Installment No", listView1.Width / 6);
                listView1.Columns.Add("Plot No", listView1.Width / 6);
                listView1.Columns.Add("ContractNo", listView1.Width / 6);
                listView1.Columns.Add("Installment", listView1.Width / 6);
                listView1.Columns.Add("ContractYear", listView1.Width / 6);
                listView1.Columns.Add("ContractLength", listView1.Width / 6);
               
                progressBar1.Visible = true;
                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = c;

                if (RDR.HasRows)
                {


                    while (RDR.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(RDR["InstallmentNo"].ToString());
                        if (RDR["PlotNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["PlotNo"].ToString());

                        }
                        if (RDR["contractNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["contractNo"].ToString());

                        }
                        if (RDR["installment"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["installment"].ToString());
                        }
                        if (RDR["ContractYear"].ToString() != "")
                        {

                            lv3.SubItems.Add(RDR["ContractYear"].ToString());
                        }
                        if (RDR["ContractLength"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["ContractLength"].ToString());
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
        private void showALLRentVouchersPrepared()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn2.Open();
                cnn.Open();
                OdbcDataReader RDR;
                String sql1 = "SELECT COUNT(*)FROM ODASMInstallment,ODASPPlot,ODASPAccount Where ODASMInstallment.DateRequisitioned <= '" +Convert .ToDateTime ( txtLastDate.Value).ToString ("yyyy/MM/dd") + "' and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo AND (ODASMInstallment.Requisitioned ='Y')";
     cmd = new OdbcCommand(sql1, cnn2);
               int c = Convert.ToInt32(cmd.ExecuteScalar().ToString());
               String sql = "SELECT * FROM ODASMInstallment,ODASPPlot,ODASPAccount Where ODASMInstallment.DateRequisitioned <= '" +Convert .ToDateTime ( txtLastDate.Value).ToString ("yyyy/MM/dd") +"' and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo AND (ODASMInstallment.Requisitioned ='Y')  ORDER BY CompanyName";
     
                cmd = new OdbcCommand(sql, cnn);
                RDR = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();

                listView1.Columns.Add("Installment No", listView1.Width / 6);
                listView1.Columns.Add("Plot No", listView1.Width / 6);
                listView1.Columns.Add("ContractNo", listView1.Width / 6);
                listView1.Columns.Add("Installment", listView1.Width / 6);
                listView1.Columns.Add("ContractYear", listView1.Width / 6);
                listView1.Columns.Add("ContractLength", listView1.Width / 6);
               
                progressBar1.Visible = true;
                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = c;

                if (RDR.HasRows)
                {


                    while (RDR.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(RDR["InstallmentNo"].ToString());
                        if (RDR["PlotNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["PlotNo"].ToString());

                        }
                        if (RDR["contractNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["contractNo"].ToString());

                        }
                        if (RDR["installment"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["installment"].ToString());
                        }
                        if (RDR["ContractYear"].ToString() != "")
                        {

                            lv3.SubItems.Add(RDR["ContractYear"].ToString());
                        }
                        if (RDR["ContractLength"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["ContractLength"].ToString());
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
         private void  showALLRentPendingConfirmation()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn2.Open();
                cnn.Open();
                OdbcDataReader RDR;
                String sql1 =  "SELECT COUNT(*) FROM ODASMInstallment,ODASPPlot,ODASPAccount Where (ODASMInstallment.DateRequisitioned >= '" + Convert.ToDateTime(txtStartDate.Text ).ToString("yyyy/MM/dd") + "' AND ODASMInstallment.DateRequisitioned <= '" + Convert.ToDateTime(txtLastDate.Text ).ToString("yyyy/MM/dd") + "') and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo AND (ODASMInstallment.PaymentFlag ='Y') ";
                cmd = new OdbcCommand(sql1, cnn2);
               int c = Convert.ToInt32(cmd.ExecuteScalar().ToString());
               String sql = "SELECT * FROM ODASMInstallment,ODASPPlot,ODASPAccount Where (ODASMInstallment.DateRequisitioned >= '" + Convert.ToDateTime(txtStartDate.Text ).ToString("yyyy/MM/dd") + "' AND ODASMInstallment.DateRequisitioned <= '" + Convert.ToDateTime(txtLastDate.Text ).ToString("yyyy/MM/dd") + "') and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo AND (ODASMInstallment.PaymentFlag ='Y')  ORDER BY CompanyName";
    
                cmd = new OdbcCommand(sql, cnn);
                RDR = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();

                listView1.Columns.Add("Installment No", listView1.Width / 6);
                listView1.Columns.Add("Plot No", listView1.Width / 6);
                listView1.Columns.Add("ContractNo", listView1.Width / 6);
                listView1.Columns.Add("Installment", listView1.Width / 6);
                listView1.Columns.Add("ContractYear", listView1.Width / 6);
                listView1.Columns.Add("ContractLength", listView1.Width / 6);
               
                progressBar1.Visible = true;
                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = c;

                if (RDR.HasRows)
                {


                    while (RDR.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(RDR["InstallmentNo"].ToString());
                        if (RDR["PlotNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["PlotNo"].ToString());

                        }
                        if (RDR["contractNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["contractNo"].ToString());

                        }
                        if (RDR["installment"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["installment"].ToString());
                        }
                        if (RDR["ContractYear"].ToString() != "")
                        {

                            lv3.SubItems.Add(RDR["ContractYear"].ToString());
                        }
                        if (RDR["ContractLength"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["ContractLength"].ToString());
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
         private void showALLRentWithPaymentsConfirmed()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn2.Open();
                cnn.Open();
                OdbcDataReader RDR;
                String sql1 = "SELECT COUNT(*) FROM ODASMInstallment,ODASPPlot,ODASPAccount Where (ODASMInstallment.DateRequisitioned >= '" + Convert.ToDateTime(txtStartDate.Value).ToString("yyyy/MM/dd") + "' AND ODASMInstallment.DateRequisitioned <= '" + Convert.ToDateTime(txtLastDate.Value).ToString("yyyy/MM/dd") + "') and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo AND (ODASMInstallment.PaymentFlag ='Y')";
                cmd = new OdbcCommand(sql1, cnn2);
               int c = Convert.ToInt32(cmd.ExecuteScalar().ToString());
               String sql = "SELECT * FROM ODASMInstallment,ODASPPlot,ODASPAccount Where (ODASMInstallment.DateRequisitioned >= '" + Convert.ToDateTime(txtStartDate.Value).ToString("yyyy/MM/dd") + "' AND ODASMInstallment.DateRequisitioned <= '" + Convert.ToDateTime(txtLastDate.Value).ToString("yyyy/MM/dd") + "') and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo AND (ODASMInstallment.PaymentFlag ='Y')  ORDER BY CompanyName";
      
                cmd = new OdbcCommand(sql, cnn);
                RDR = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();

                listView1.Columns.Add("Installment No", listView1.Width / 6);
                listView1.Columns.Add("Plot No", listView1.Width / 6);
                listView1.Columns.Add("ContractNo", listView1.Width / 6);
                listView1.Columns.Add("Installment", listView1.Width / 6);
                listView1.Columns.Add("ContractYear", listView1.Width / 6);
                listView1.Columns.Add("ContractLength", listView1.Width / 6);
               
                progressBar1.Visible = true;
                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = c;

                if (RDR.HasRows)
                {


                    while (RDR.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(RDR["InstallmentNo"].ToString());
                        if (RDR["PlotNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["PlotNo"].ToString());

                        }
                        if (RDR["contractNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["contractNo"].ToString());

                        }
                        if (RDR["installment"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["installment"].ToString());
                        }
                        if (RDR["ContractYear"].ToString() != "")
                        {

                            lv3.SubItems.Add(RDR["ContractYear"].ToString());
                        }
                        if (RDR["ContractLength"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["ContractLength"].ToString());
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
    
        private void frmODASSearchSitesNotPaid_Load(object sender, EventArgs e)
        {
          
            if (strReport == "PendingPaymentAsAtASingleDate")
            {
                this.Text = "Rents Pending payment As At";
                txtLastDate.Visible = false;
            }
            else if (strReport == "PendingPayment")
            {
                this.Text = "Rents Pending payment As At";

            }
            else if (strReport == "VouchersPrepared")
            {
                this.Text = "Payment Vouchers Prepared Between";
            }
            else if (strReport == "PendingConfirmation")
            {
                this.Text = "Payments Pending Confirmation Between";
            }
            else if (strReport == "PaymentsConfirmed")
            {
                this.Text = "Payments Confirmed Between";
            }
            else
            {
                this.Text = "Outstanding Payments Between";
            }
        }

        private void txtLastDate_CloseUp(object sender, EventArgs e)
        {
            if (strReport == "PendingPaymentAsAtASingleDate")
            {
                showALLRentPendingPaymentAsAtASingleDate();
            }
            else if (strReport == "PendingPayment")
            {
                showALLRentPendingPaymentAsAtASingleDate();

            }
            else if (strReport == "VouchersPrepared")
            {
                showALLRentVouchersPrepared();
            }
            else if (strReport == "PendingConfirmation")
            {
                showALLRentPendingConfirmation();
            }
            else if (strReport == "PaymentsConfirmed")
            {
                showALLRentWithPaymentsConfirmed();
            }
            else
            {
                showALLRentNOTPAIDThisPeriod();
            }
           
        }

        private void txtStartDate_CloseUp(object sender, EventArgs e)
        {
            if (strReport == "PendingPaymentAsAtASingleDate")
            {
                showALLRentPendingPaymentAsAtASingleDate();
            }
            else if (strReport == "PendingPayment")
            {
                showALLRentPendingPaymentAsAtASingleDate();

            }
            else if (strReport == "VouchersPrepared")
            {
                showALLRentVouchersPrepared();
            }
            else if (strReport == "PendingConfirmation")
            {
                showALLRentPendingConfirmation();
            }
            else if (strReport == "PaymentsConfirmed")
            {
                showALLRentWithPaymentsConfirmed();
            }
            else
            {
                showALLRentNOTPAIDThisPeriod();
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {

            
            GeneralVariables vars = new GeneralVariables();
            if (strReport == "PaymentsConfirmed")
            {
                vars.rptRentDue.strReport = strReport;
            }
            else if (strReport == "PendingPaymentAsAtASingleDate")
            {
                vars.rptRentDue.strReport = strReport;
            }
            else if (strReport == "PendingPayment")
            {
                vars.rptRentDue.strReport = strReport;
            }
            else if (strReport == "VouchersPrepared")
            {
                vars.rptRentDue.strReport = strReport;
            }
            else if (strReport == "PendingConfirmation")
            {
                vars.rptRentDue.strReport = strReport;
            }
            else
            {


            }
            vars.rptRentDue.txtStartDate = txtStartDate.Text;
            vars.rptRentDue.txtLastDate = txtLastDate.Text;
            vars.rptRentDue.ShowDialog();
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
                int x=0;
                int c = listView1.Items.Count-1;
                progressBar1.Visible = true;
              
                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = c+1;

                    foreach (ListViewItem comp in listView1.Items)
                    {
                       
                        xlWorkSheet.Cells[i, j] = comp.Text.ToString();
                        for (int y = 1; y <= listView1.Columns.Count;y++ )
                        {
                            xlWorkSheet.Cells[j, y] = listView1 .Columns[y-1].Text ;
                        }
                        foreach (ListViewItem.ListViewSubItem drv in comp.SubItems)
                        {

                            xlWorkSheet.Cells[i, j] = drv.Text.ToString();
                            xlWorkSheet.Cells[listView1.Items.Count + 2, listView1.Columns.Count] =label4 .Text ;
                     
                            j++;
                              }
                       
                        j = 1;
                       
                        i++;

                        xlWorkSheet.Cells[listView1 .Items .Count +2, listView1.Columns.Count-1] = "Totals";

                        progressBar1.Value = progressBar1.Value + 1;
                    }
                   
                      
                xlWorkBook.SaveAs(System .Windows .Forms .Application .StartupPath +"\\Plots not Paid.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                MessageBox.Show("File stored in " + System.Windows.Forms.Application.StartupPath + "\\Plots not Paid.xls", "Excel File Created Successfully", MessageBoxButtons.OK, MessageBoxIcon.Information);

                System.Diagnostics.Process.Start(System.Windows.Forms.Application.StartupPath + "\\Plots not Paid.xls");
              
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
        }

        private void txtStartDate_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
