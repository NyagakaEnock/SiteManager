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
    public partial class rptODASAgreementForm : Form
    {
        public rptODASAgreementForm()
        {
            InitializeComponent();
        }
        public String CurrentRecord;
        private void rptODASAgreementForm_Load(object sender, EventArgs e)
        {
            try
            {
                rptODASAgreementSchedule AgreementSchedule = new rptODASAgreementSchedule();
                AgreementSchedule.RecordSelectionFormula = "{ODASMLeaseAgreement.ContractNo } = '" + CurrentRecord + "'";
                AgreementSchedule.OpenSubreport("ODASRAgreementSchedule").RecordSelectionFormula = "{ODASMInstallment.ContractNo}= '" + CurrentRecord + "' and {ODASMInstallment.Installment} = '1'";
                crystalReportViewer1.ReportSource = AgreementSchedule;
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
        }
    }
}
