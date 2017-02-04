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
    public partial class rptODASPPlotSites : Form
    {
        public rptODASPPlotSites()
        {
            InitializeComponent();
        }
        public string currentRecord;
        private void rptODASPPlotSites_Load(object sender, EventArgs e)
        {
            rptPlotSites rptPlotSites = new rptPlotSites();
            rptPlotSites.RecordSelectionFormula = "{ODASPPlot.AccountNo} = '" + currentRecord + "'";
            crystalReportViewer1.ReportSource = rptPlotSites;
        }
    }
}
