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
    public partial class rptODASRentPaymentInstallment : Form
    {
        public rptODASRentPaymentInstallment()
        {
            InitializeComponent();
        }
        public String currentRecord;
        private void rptODASRentPaymentInstallment_Load(object sender, EventArgs e)
        {
          
            ODASRentPaymentInstallments ODASRentInstallments = new ODASRentPaymentInstallments();
            
            ODASRentInstallments.RecordSelectionFormula = "{ODASMLeaseAgreement.ContractNo}='" + currentRecord + "'";
            crystalReportViewer1.ReportSource = ODASRentInstallments;
        }
    }
}
