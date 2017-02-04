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
    public partial class frmRptODASAllRoadSites : Form
    {
        public frmRptODASAllRoadSites()
        {
            InitializeComponent();
        }
        OdbcCommand cmd;
        OdbcDataAdapter da;
        DataSet ds;
        private void frmRptODASAllRoadSites_Load(object sender, EventArgs e)
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                string sql = "Select * From ODASPPlot where OnRoadReserve='Y'";

                cmd = new OdbcCommand(sql, cnn);

                //  OdbcDataReader reader;
                // reader = cmd.ExecuteReader();
                // reader.Read();
                //  DataTable dt;
                // dt = new DataTable();
                // dt.Load(reader);
                //          reader.Close();
                da = new OdbcDataAdapter(cmd);
                ds = new DataSet();
               
                da.Fill(ds, "ODASPPlot");
                RptODASAllRoadSites rpt = new RptODASAllRoadSites();
                 
                //  a = "{ODASMInstallment.PaymentDueDate} >= DateTime('" + Convert.ToDateTime(txtStartDate).ToString("MM/dd/yyyy") + "') and {ODASMInstallment.PaymentDueDate} <= DateTime('" + Convert.ToDateTime(txtLastDate).ToString("MM/dd/yyyy") + "') And  (ISNULL({ODASMLeaseAgreement.Terminated})  OR {ODASMLeaseAgreement.Terminated}='N' )";

                //   rpt.RecordSelectionFormula = a;

                rpt.SetDataSource(ds);
                crystalReportViewer1.ReportSource = rpt;
                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
