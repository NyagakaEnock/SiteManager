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
    public partial class frmAllSites : Form
    {
        public frmAllSites()
        {
            InitializeComponent();
        }

        private void frmAllSites_Load(object sender, EventArgs e)
        {
            RptODASAllSites RptODASAllSites = new RptODASAllSites();
            RptODASAllSites.RecordSelectionFormula = "";
            crystalReportViewer1.ReportSource = RptODASAllSites;
        }
    }
}
