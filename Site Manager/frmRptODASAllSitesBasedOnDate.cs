using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Site_Manager
{
    public partial class frmRptODASAllSitesBasedOnDate : Form
    {
        public frmRptODASAllSitesBasedOnDate()
        {
            InitializeComponent();
        }
        public string strStartDate;
        public string strLastDate;
        private void frmRptODASAllSitesBasedOnDate_Load(object sender, EventArgs e)
        {
            RptODASAllSitesBasedOnDate rpt = new RptODASAllSitesBasedOnDate();
            rpt.RecordSelectionFormula ="{ODASPPlot.DateCreated}>=Date('" +Convert .ToDateTime  (strStartDate).ToString ("yyyy,MM,dd") + "') AND {ODASPPlot.DateCreated}<=Date('"+Convert .ToDateTime (strLastDate).ToString ("yyyy,MM,dd") + "')";
            crystalReportViewer1.ReportSource = rpt;
        }
    }
}
