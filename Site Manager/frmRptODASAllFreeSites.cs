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
    public partial class frmRptODASAllFreeSites : Form
    {
        public frmRptODASAllFreeSites()
        {
            InitializeComponent();
        }
        string a;
        public string txtStartDate;
        public string txtLastDate;
        DataTable dt;
        DataSet ds;
        OdbcCommand cmd;
        OdbcDataAdapter da;
        private void frmRptODASAllFreeSites_Load(object sender, EventArgs e)
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                string sql = "SELECT * FROM ODASPPlot,ODASPPlotSite where ODASPPlotSite.JobBriefNo is Null and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo ";

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
               
                da.Fill(ds, "ODASPPlotSite");
                da.Fill(ds, "ODASPPlot");
                RptODASAllFreeSites1 rpt = new RptODASAllFreeSites1();
                 
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
