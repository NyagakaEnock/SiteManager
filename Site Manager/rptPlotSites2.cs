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
    public partial class rptPlotSites2 : Form
    {
        public rptPlotSites2()
        {
            InitializeComponent();
        }
        public string currentRecord;
        private void rptPlotSites2_Load(object sender, EventArgs e)
        {
            rptPlotSites rptPlotSites = new rptPlotSites();
            rptPlotSites.RecordSelectionFormula = "{ODASPPlot.PlotNo} = '" + currentRecord + "'";
            crystalReportViewer1.ReportSource = rptPlotSites;
        }
    }
}
