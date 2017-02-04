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
    public partial class rptCouncilRates : Form
    {
        public rptCouncilRates()
        {
            InitializeComponent();
        }
        public String currentRecord;
        private void rptCouncilRates_Load(object sender, EventArgs e)
        {
             
            rptODASRRatesSchedule rptPlotSites = new rptODASRRatesSchedule();
          //  rptPlotSites.RecordSelectionFormula = "{ODASMCouncilRateDue.SiteNo} = '" + currentRecord + "' and {ODASMCouncilRateDue.CurrentYear}= '" & INPQRY2 & "' and {ODASMCouncilRateDue.JobBriefItemNo}= '" & CurrentRecord1 & "'";
            crystalReportViewer1.ReportSource = rptPlotSites;
        }
    }
}
