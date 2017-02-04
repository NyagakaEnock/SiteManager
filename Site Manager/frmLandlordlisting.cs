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
    public partial class frmLandlordlisting : Form
    {
        public frmLandlordlisting()
        {
            InitializeComponent();
        }

        private void frmLandlordlisting_Load(object sender, EventArgs e)
        {
            GeneralVariables vars = new GeneralVariables();
            ODASRLandLords ODASRLandLords = new ODASRLandLords();
            ODASRLandLords.RecordSelectionFormula = "{ODASPAccountType.LandLord} = 'Y'";
            crystalReportViewer1.ReportSource = ODASRLandLords;
        
        }
    }
}
