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
    public partial class frmPlotAllocation : Form
    {
        public frmPlotAllocation()
        {
            InitializeComponent();
        }

        private void frmPlotAllocation_Load(object sender, EventArgs e)
        {
            ODASRPlotAllocations ODASRPlotAllocations  = new ODASRPlotAllocations();
            ODASRPlotAllocations.RecordSelectionFormula = "";
            crystalReportViewer1.ReportSource = ODASRPlotAllocations;
        }
    }
}
