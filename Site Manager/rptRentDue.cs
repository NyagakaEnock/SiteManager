using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;

namespace Site_Manager
{
    public partial class rptRentDue : Form
    {
        public rptRentDue()
        {
            InitializeComponent();
        }
        string a;
        public string txtStartDate;
        public string txtLastDate;
        DataTable dt;
        DataSet ds;
        OdbcCommand cmd;
        public String sql;
        OdbcDataAdapter da;
        public String  strReport;
        public string title;
      //  public string sql;
        private void rptRentDue_Load(object sender, EventArgs e)
        {
            try
            {
                
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
              
                if(strReport =="PaymentsConfirmed"){
                    sql = "SELECT * FROM ODASMInstallment,ODASPPlot,ODASPAccount Where (ODASMInstallment.DateRequisitioned >= '" +Convert .ToDateTime ( txtStartDate).ToString ("yyyy/MM/dd") + "' AND ODASMInstallment.DateRequisitioned <= '" + Convert .ToDateTime ( txtLastDate).ToString ("yyyy/MM/dd") + "') and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo AND (ODASMInstallment.PaymentFlag ='Y')  ORDER BY CompanyName";
                    title = "Payments Confirmed Between "+txtStartDate +" to "+txtLastDate;
                }
                else if (strReport == "PendingPaymentAsAtASingleDate")
                {
                    sql = "SELECT  * FROM ODASMInstallment,ODASPPlot,ODASPAccount,ODASMLeaseAgreement Where ODASMLeaseAgreement.PlotNo=ODASMInstallment.PlotNo AND ODASMLeaseAgreement.ContractNo=ODASMInstallment.ContractNo AND (ODASMLeaseAgreement.Terminated='N' OR ODASMLeaseAgreement.Terminated IS NULL) AND (ODASMInstallment.PaymentDueDate <= '" + Convert.ToDateTime(txtLastDate).ToString("yyyy/MM/dd") + "') and ODASMInstallment.PaymentFlag = 'N' and ODASMInstallment.Balance > 0 and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo ORDER BY CompanyName";
                    title = "Rents Pending payment As At "+txtLastDate;
                }
                else if (strReport == "VouchersPrepared")
                {
                   
                    sql = "SELECT * FROM ODASMInstallment,ODASPPlot,ODASPAccount Where ODASMInstallment.DateRequisitioned <= '" +Convert .ToDateTime ( txtLastDate).ToString ("yyyy/MM/dd") + "' and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo AND (ODASMInstallment.Requisitioned ='Y')  ORDER BY CompanyName";

                    title = "Payment Vouchers Prepared Between " + txtStartDate + " to " + txtLastDate;
                }
                else if (strReport == "PendingPayment")
                {
                    sql = "SELECT  * FROM ODASMInstallment,ODASPPlot,ODASPAccount,ODASMLeaseAgreement Where ODASMLeaseAgreement.PlotNo=ODASMInstallment.PlotNo AND ODASMLeaseAgreement.ContractNo=ODASMInstallment.ContractNo AND (ODASMLeaseAgreement.Terminated='N' OR ODASMLeaseAgreement.Terminated IS NULL) AND (ODASMInstallment.PaymentDueDate <= '" + Convert.ToDateTime(txtLastDate).ToString("yyyy/MM/dd") + "') and ODASMInstallment.PaymentFlag = 'N' and ODASMInstallment.Balance > 0 and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo ORDER BY CompanyName";
                    title = "Rents Pending payment As At " + txtLastDate;
                }
                else if (strReport == "PendingConfirmation")
                {
                    sql = "SELECT * FROM ODASMInstallment,ODASPPlot,ODASPAccount Where (ODASMInstallment.DateRequisitioned >= '" +Convert .ToDateTime ( txtStartDate).ToString ("yyyy/MM/dd") + "' AND ODASMInstallment.DateRequisitioned <= '" +Convert .ToDateTime ( txtLastDate).ToString ("yyyy/MM/dd") + "') and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo AND (ODASMInstallment.PaymentFlag ='Y')  ORDER BY CompanyName";
                    title = "Payments Pending Confirmation Between " + txtStartDate + " to " + txtLastDate;
                }
                else
                {
                    sql = "Select * from ODASMInstallment, ODASPAccount,ODASPPlot Where ODASPPlot.PlotNo=ODASMInstallment.PlotNo AND (ODASMInstallment.Requisitioned = 'N' or ODASMInstallment.Requisitioned is null or ODASMInstallment.Requisitioned = 'Y' ) and ODASPAccount.AccountNo = ODASPPlot.AccountNo  AND (ODASMInstallment.PaymentDueDate>='" + Convert.ToDateTime(txtStartDate).ToString("yyyy/MM/dd") + "' AND ODASMInstallment.PaymentDueDate<='" + Convert.ToDateTime(txtLastDate).ToString("yyyy/MM/dd") + "') AND ODASMInstallment.Balance>0 AND ODASMInstallment.ContractNo IN (Select ContractNo From ODASMLeaseAgreement where (Terminated='N' OR Terminated IS NULL))";
                    title = "Outstanding Payments Between "+ txtStartDate + " to " + txtLastDate;
                }
             
                cmd = new OdbcCommand(sql, cnn);

                da = new OdbcDataAdapter(cmd);
               ds = new DataSet();
               da.Fill(ds, "ODASMInstallment");
           
                rptODASRentDue rpt = new rptODASRentDue();
               
              a = "{ODASMInstallment.PaymentDueDate} >= DateTime('" + Convert.ToDateTime(txtStartDate).ToString("MM/dd/yyyy") + "') and {ODASMInstallment.PaymentDueDate} <= DateTime('" + Convert.ToDateTime(txtLastDate).ToString("MM/dd/yyyy") + "') And  (ISNULL({ODASMLeaseAgreement.Terminated})  OR {ODASMLeaseAgreement.Terminated}='N' )";
              TextObject txttitle = (TextObject)rpt.ReportDefinition.Sections["Section2"].ReportObjects["txttitle"];
              txttitle.Text = title;
             
                if(strReport =="PaymentsConfirmed"){
                 rptODASRentpaid rptODASRentpaid= new rptODASRentpaid ();
                 rptODASRentpaid.RecordSelectionFormula = a;
                 rptODASRentpaid.SetDataSource(ds);
                crystalReportViewer1.ReportSource =rptODASRentpaid;
                TextObject txttitle2 = (TextObject)rptODASRentpaid.ReportDefinition.Sections["Section2"].ReportObjects["txttitle"];
                txttitle2.Text = title;
                }
                else if (strReport == "PendingPaymentAsAtASingleDate")
                {
                    ds = new DataSet();
                    da.Fill(ds, "ODASMInstallment");
                      a = "{ODASMInstallment.PaymentDueDate} <= DateTime('" + Convert.ToDateTime(txtLastDate).ToString("MM/dd/yyyy") + "') and (ISNULL({ODASMLeaseAgreement.Terminated})  OR {ODASMLeaseAgreement.Terminated}='N' )";
                    
                    rpt.RecordSelectionFormula = a;
                    rpt.SetDataSource(ds);
                    crystalReportViewer1.ReportSource = rpt;
                }
                else if (strReport == "VouchersPrepared")
                {
                    
                    rpt.SetDataSource(ds);
                    crystalReportViewer1.ReportSource = rpt;

                }
                else if (strReport == "PendingPayment")
                {
                  
                    rpt.SetDataSource(ds);
                    crystalReportViewer1.ReportSource = rpt;
                }
                else if (strReport =="PendingConfirmation")
                {
                    rpt.SetDataSource(ds);
                    crystalReportViewer1.ReportSource = rpt;
                }
                else
                {
                    rpt.RecordSelectionFormula = a;
                    rpt.SetDataSource(ds);
                    crystalReportViewer1.ReportSource = rpt;
                }
             

                cnn.Close();
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
        }
    }
}
