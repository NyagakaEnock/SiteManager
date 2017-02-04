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
    public partial class Sites_to_Expire : Form
    {
        public Sites_to_Expire()
        {
            InitializeComponent();
        }
        public String strReport;
        OdbcCommand cmd;
        public String  txtLastDate;
        public String txtStartDate;
        private void Sites_to_Expire_Load(object sender, EventArgs e)
        {
            try {
               
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
             
                cnn.Open();
                string sql;
         
                if (strReport == "")
                {
                    sql = "SELECT ODASPPlot.*,ODASPAccount.COmpanyName  FROM ODASPPlot, ODASPAccount,ODASMLeaseAgreement WHERE ODASMLeaseAgreement.PlotNo=ODASPPlot.PlotNo AND ODASPPLot.ExpiryDate >= '" + Convert.ToDateTime(txtStartDate).ToString("yyyy/MM/dd") + "' and ODASPPLot.ExpiryDate <= '" + Convert.ToDateTime(txtLastDate).ToString("yyyy/MM/dd") + "' and ODASPPlot.AccountNo = ODASPAccount.AccountNo and (ODASMLeaseAgreement.terminated='N' OR ODASMLeaseAgreement.Terminated IS NULL)";
                }
                else if (strReport == "ExpiredNotRenewed")
                {
                    sql = "SELECT ODASPPlot.*,ODASPAccount.COmpanyName  FROM ODASPPlot, ODASPAccount,ODASMLeaseAgreement WHERE  ODASMLeaseAgreement.PlotNo=ODASPPlot.PlotNo AND ODASPPLot.ExpiryDate <= '" + Convert.ToDateTime(txtLastDate).ToString("yyyy/MM/dd") + "' and ODASPPlot.AccountNo = ODASPAccount.AccountNo  AND (ODASMLeaseAgreement.terminated='N' OR ODASMLeaseAgreement.Terminated IS NULL) and (ODASMLeaseAgreement.Renewal='0' OR ODASMLeaseAgreement.Renewal IS NULL)";

                }
                else
                {
                    sql = "SELECT ODASPPlot.*,ODASPAccount.COmpanyName  FROM ODASPPlot, ODASPAccount,ODASMLeaseAgreement WHERE  ODASMLeaseAgreement.PlotNo=ODASPPlot.PlotNo AND ODASPPLot.ExpiryDate <= '" + Convert.ToDateTime(txtLastDate).ToString("yyyy/MM/dd") + "' and ODASPPlot.AccountNo = ODASPAccount.AccountNo  AND (ODASMLeaseAgreement.terminated='N' OR ODASMLeaseAgreement.Terminated IS NULL)";
                }
                cmd = new OdbcCommand(sql, cnn);
                OdbcDataAdapter da;
                da = new OdbcDataAdapter(cmd);
                DataSet ds;
                ds = new DataSet();
                da.Fill(ds, "ODASPPlot");
                da.Fill(ds, "ODASPAccount");
                da.Fill(ds, "ODASMLeaseAgreement");
                RptODASSitesToExpire RPT = new RptODASSitesToExpire();

             //   RPT.SetParameterValue("NameTitle", "your parameter value");
                RPT.SetDataSource(ds);
                TextObject txttitle = (TextObject)RPT.ReportDefinition.Sections["Section1"].ReportObjects["txttitle"];
                if (strReport == "")
                {
                    txttitle.Text = "Sites expiring within the period Starting from " + Convert.ToDateTime(txtStartDate).ToString("dd/MM/yyyy") + " to " + Convert.ToDateTime(txtLastDate).ToString("dd/MM/yyyy") + "";
           
                }
                else {
                    txttitle.Text = "Sites to expire by " + Convert.ToDateTime(txtLastDate).ToString("dd/MM/yyyy");
                }
                  crystalReportViewer1.ReportSource = RPT;
                cnn.Close();
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
           

        }
    }
}
