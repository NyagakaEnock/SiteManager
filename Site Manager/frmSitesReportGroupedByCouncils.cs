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
    public partial class frmSitesReportGroupedByCouncils : Form
    {
        public frmSitesReportGroupedByCouncils()
        {
            InitializeComponent();
        }
        public String strCouncilcode;
        private void frmSitesReportGroupedByCouncils_Load(object sender, EventArgs e)
        {
            rptSitesReportGroupedByCouncils rpt = new rptSitesReportGroupedByCouncils();
            if (strCouncilcode == "")
            {
                rpt.RecordSelectionFormula = "";
            }
            else {
                rpt.RecordSelectionFormula = "{ODASPCouncil.CouncilCode } = '" + strCouncilcode + "'";
            }
              
                crystalReportViewer1 .ReportSource =rpt;
        }
    }
}
