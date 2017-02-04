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
    public partial class rptContractAgreement : Form
    {
        public rptContractAgreement()
        {
            InitializeComponent();
        }
        public String CurrentRecord;
        private void rptContractAgreement_Load(object sender, EventArgs e)
        {
           
            rptODASLandlordContract1 rptODASLandlordContract1 = new rptODASLandlordContract1();
            rptODASLandlordContract1.RecordSelectionFormula = "{ODASMLeaseAgreement.ContractNo } = '" + CurrentRecord + "'";
            crystalReportViewer1.ReportSource = rptODASLandlordContract1;
        }
    }
}
