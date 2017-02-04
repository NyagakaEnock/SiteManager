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
    public partial class frmRStatement : Form
    {
        public frmRStatement()
        {
            InitializeComponent();
        }
        public String  strAccountNo;
        private void frmRStatement_Load(object sender, EventArgs e)
        {
            OdbcDataReader reader;
            GeneralVariables GeneralVariables = new GeneralVariables();
            OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
            cnn.Open();
            OdbcDataAdapter da;
            DataSet ds;
            OdbcCommand cmd;
string sql="SELECT    * FROM         ODASMInstallment AS ODASMInstallment INNER JOIN " +
                      "ODASPPlot AS ODASPPlot ON ODASMInstallment.PlotNo = ODASPPlot.PlotNo INNER JOIN " +
                      "ODASPAccount AS ODASPAccount ON ODASPPlot.AccountNo = ODASPAccount.AccountNo WHERE ODASPAccount.AccountNo LIKE '" + strAccountNo + "'" +
"ORDER BY ODASPAccount.AccountNo";
            cmd = new OdbcCommand(sql ,cnn);
            ds = new DataSet();
            da = new OdbcDataAdapter(cmd);
            da.Fill(ds, "ODASMInstallment");
            //da.Fill(ds, "ODASPPlot");
            //da.Fill(ds, "ODASPAccount");
        // DataTable dt;
        //   dt = new DataTable();
         // reader = cmd.ExecuteReader();
         //  reader.Read();
         //  dt.Load(reader );
      
            rptStatement statement = new rptStatement();
            statement.RecordSelectionFormula = "{ODASPAccount.AccountNo} = '" + strAccountNo + "'";

            statement.SetDataSource(ds);
            crystalReportViewer1.ReportSource = statement;
        }
    }
}
