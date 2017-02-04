using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Odbc ;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows.Forms;

namespace Site_Manager
{
    public partial class frmPaymentConfirmation : Form
    {
        public frmPaymentConfirmation()
        {
            InitializeComponent();
        }
        OdbcCommand cmd;
        OdbcDataReader reader;
        public string CurrentUserName;
        Double PartialPaid;
        private void label6_Click(object sender, EventArgs e)
        {

        }
        private void DisableChilds(Control ctrl)
        {
            foreach (Control c in ctrl.Controls)
            {
                DisableChilds(c);
                if (c is TextBox)
                {
                    ((TextBox)(c)).Enabled = false;
                }
                if (c is CheckBox)
                {
                    ((CheckBox)(c)).Enabled = false;
                }
                if (c is RadioButton)
                {
                    ((RadioButton)(c)).Enabled = false;
                }
                if (c is ComboBox)
                {
                    ((ComboBox)(c)).Enabled = false;
                }
                if (c is DateTimePicker)
                {
                    ((DateTimePicker)(c)).Enabled = false;
                }
            }
        }
        private void ableChilds(Control ctrl)
        {
            foreach (Control c in ctrl.Controls)
            {
                ableChilds(c);
                if (c is TextBox)
                {
                    ((TextBox)(c)).Enabled = true ;
                }
                if (c is CheckBox)
                {
                    ((CheckBox)(c)).Enabled = true;
                }
                if (c is RadioButton)
                {
                    ((RadioButton)(c)).Enabled = true;
                }
                if (c is ComboBox)
                {
                    ((ComboBox)(c)).Enabled = true;
                }
                if (c is DateTimePicker)
                {
                    ((DateTimePicker)(c)).Enabled = true;
                }
            }
        }
        private void clearChilds(Control ctrl)
        {
            foreach (Control c in ctrl.Controls)
            {
                clearChilds(c);
                if (c is TextBox)
                {
                    ((TextBox)(c)).Text = "";
                }
                if (c is CheckBox)
                {
                    ((CheckBox)(c)).Checked = false;
                }
                if (c is RadioButton)
                {
                    ((RadioButton)(c)).Checked = false;
                }
                if (c is ComboBox)
                {
                    ((ComboBox)(c)).Text = "";
                }

            }
        }
        private void clearAllRecords()
        {
            clearChilds(this);
        }
        public void disableALLRECORD()
        {
            try
            {
                DisableChilds(this);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        public void ableALLRECORD()
        {
            try
            {
                ableChilds(this);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void GetRentRequisitioned()
        {
            try
            {

                listView1.Columns.Clear();
                listView1.Items.Clear();
                listView1.Columns.Add("Installment", listView1.Width / 6);
                listView1.Columns.Add("Contract No", listView1.Width / 6);
                listView1.Columns.Add("Due Date", listView1.Width / 6);
                listView1.Columns.Add("Amount", listView1.Width / 6);
                listView1.Columns.Add("Account No", listView1.Width / 6);
                listView1.Columns.Add("LandLord", listView1.Width / 6);
                listView1.Columns.Add("Requisition Date", listView1.Width / 6);
                listView1.Columns.Add("Serial", listView1.Width / 6);
                ListView.ListViewItemCollection items = listView1.Items;
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcCommand cmd2;
                cnn.Open();
                cnn2.Open();
                cmd2 = new OdbcCommand("Select COUNT(*) from ODASMInstallment I, ODASPAccount A,ODASMLeaseAgreement L Where (L.PlotNo=I.PlotNo AND L.ContractNo=I.ContractNo AND (L.Terminated='N' OR L.Terminated IS NULL) ) AND (I.Requisitioned = 'Y' OR I.Requisitioned ='P' ) and I.AccountNo = A.AccountNo AND (I.VoucherDate>='" + DTPStartDate.Text + "' AND I.VoucherDate<='" + DTPLastDate.Text + "') AND I.ChequeNo IS NULL", cnn);
                cmd = new OdbcCommand("Select * from ODASMInstallment I, ODASPAccount A,ODASMLeaseAgreement L Where (L.PlotNo=I.PlotNo AND L.ContractNo=I.ContractNo AND (L.Terminated='N' OR L.Terminated IS NULL) ) AND (I.Requisitioned = 'Y' OR I.Requisitioned ='P' ) and I.AccountNo = A.AccountNo AND (I.VoucherDate>='" + Convert.ToDateTime(DTPStartDate.Text).ToString("yyyy/MM/dd") + "' AND I.VoucherDate<='" + Convert.ToDateTime(DTPLastDate.Text).ToString("yyyy/MM/dd") + "') AND I.ChequeNo IS NULL", cnn2);



                String c = cmd2.ExecuteScalar().ToString();
                reader = cmd.ExecuteReader();
             
                if (Convert.ToInt32(c) != 0)
                {
                    progressBar1.Visible = true;
                    progressBar1.Value = 0;
                    progressBar1.Minimum = 0;
                    progressBar1.Maximum = Convert.ToInt32(c);

                }

                if (reader.Read())
                {


                    while (reader.Read())
                    {
                        String dat = reader["PaymentDueDate"].ToString();
                        ListViewItem lv3 = new ListViewItem(reader["InvoiceNo"].ToString());
                        if (reader["ContractNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["ContractNo"].ToString());
                        }

                        if (reader["PaymentDueDate"].ToString() != "")
                        {
                            lv3.SubItems.Add(Convert.ToDateTime(dat).ToString("MM/dd/yyyy"));
                        }
                        if (reader["AmountPaid"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["AmountPaid"].ToString());
                        }
                        if (reader["AccountNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["AccountNo"].ToString());
                        }
                        if (reader["CompanyName"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["CompanyName"].ToString());
                        }

                        if (reader["VoucherDate"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["VoucherDate"].ToString());
                        }

                        if (reader["InstallmentNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["InstallmentNo"].ToString());
                        }
                        listView1.Items.Add(lv3);

                        progressBar1.Value = progressBar1.Value + 1;


                    }
                    progressBar1.Visible = false;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void GetInvoicesREQUISITIONED()
        {
            try
            {
                listView2.Columns.Clear();
                listView2.Items.Clear();
                listView2.Columns.Add("Item No", listView1.Width / 4);
                listView2.Columns.Add("Invoice No", listView1.Width / 4);
                listView2.Columns.Add("LPO No", listView1.Width / 4);
                listView2.Columns.Add("Amount", listView1.Width / 4);


                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();
                cmd = new OdbcCommand("Select *  from ODASMVoucherItem Where VoucherNo = '" + txtVoucherNo.Text + "'", cnn);


                reader = cmd.ExecuteReader();
                listView2.Items.Clear();
                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["VoucherItemNo"].ToString());
                        if (reader["LPONo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["LPONo"].ToString());
                        }
                        if (reader["DocumentNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["DocumentNo"].ToString());
                        }

                        if (reader["AmountPaid"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["AmountPaid"].ToString());
                        }


                        listView2.Items.Add(lv3);




                    }


                }
                reader.Close();
                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        public String getPeriod()
        {

            String strMonth, strYr, result;
            strMonth = DateTime.Today.Month.ToString().Trim();
            if (strMonth.Length == 1)
            {

                strMonth = "0" + strMonth;

            }
            strYr = DateTime.Today.Year.ToString().Trim();
            result = strYr.Trim() + "/" + strMonth.Trim();
            return result;
        }
        private void loadPaymentDescription()
        {
            try
            {


                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();

                cmd = new OdbcCommand("SELECT * FROM ODASPCostCentre WHERE CostCentre = '" + txtCostCenter.Text + "'", cnn);


                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    txtPaymentDescription.Text = reader["COSTCENTREDescription"].ToString();
                }
                reader.Close();
                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void computeVOUCHERTOTAL()
        {
            try
            {


                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();

                cmd = new OdbcCommand("SELECT sum(AmountPaid) as totals from ODASMVoucherItem where VoucherNo = '" + txtVoucherNo.Text + "'", cnn);

                String c;
                c = cmd.ExecuteScalar().ToString();


                if (c == "")
                {
                    if (PartialPaid > 0)
                    {
                        txtVoucherAmount.Text = PartialPaid.ToString();
                    }
                    else
                    {
                        txtVoucherAmount.Text = Convert.ToDouble(GeneralVariables.VourcherPrepareForm.txtAmountPaid.Text).ToString();
                    }
                }
                else
                {
                    if (PartialPaid == 0)
                    {
                        txtVoucherAmount.Text = c;
                    }
                    else
                    {

                        txtVoucherAmount.Text = PartialPaid.ToString();
                    }
                }


                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void countVOUCHERITEMS()
        {
            try
            {


                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();

                cmd = new OdbcCommand("SELECT count(VoucherNo) as totals from ODASMVoucherItem where VoucherNo = '" + txtVoucherNo.Text + "'", cnn);

                String c;
                c = cmd.ExecuteScalar().ToString();

                if (c == "")
                {
                    txtItems.Text = "0";
                }

                else { txtItems.Text = c; }

                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void frmPaymentConfirmation_Load(object sender, EventArgs e)
        {
            GeneralVariables var = new GeneralVariables();
            var.NewRecord = false;
            disableALLRECORD();
          
            clearAllRecords();
            String dat = DTPStartDate.Text;
            DateTime dt = Convert.ToDateTime(dat);

            DTPStartDate.Value = dt.AddDays(-7);
            GetRentRequisitioned();
            GetInvoicesREQUISITIONED();
            txtCurrentPeriod.Text = getPeriod();
            loadPaymentDescription();
            //computeVOUCHERTOTAL();
            countVOUCHERITEMS();
        }
        private String getprevPendingContracts()
        {
            String str = "";
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASMInstallment WHERE  ContractNo='" + txtContractNo.Text + "' AND ContractYear<" + Convert.ToInt32(txtContractYear.Text) + " AND (Requisitioned='N' OR Requisitioned IS NULL) ", cnn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    str = str + reader["ContractYear"].ToString() + "- [" + reader["PaymentDueDate"].ToString() + "]";
                   
                }
                cnn.Close();

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return str;
        }
        private void listView1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            try
            {
               
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcDataReader reader2;
                cnn.Open();
                cnn2.Open();
                Double i, j;
                ListView.CheckedListViewItemCollection checkedItems = listView1.CheckedItems;
                foreach (ListViewItem item in checkedItems)
                {
                 
                    if (item.Checked == true)
                    {
                      // GetRentRequisitioned();
                        j = checkedItems.Count;
                        if (j == 0)
                        {
                            return;
                        }
                        txtAccountNo .Text =item .SubItems [4].Text ;
                        txtExpiryDate .Text =item .SubItems [2].Text ;
                        txtCouncilCode.Text = "";
                        txtPayeeDetails .Text =item .SubItems [5].Text ;
                        txtInvoiceAmount .Text =item .SubItems [3].Text ;
                        txtInvoiceBalance.Text = "0";
                        txtInstallmentNo .Text =item .SubItems [7].Text ;
                        cnn.Close();
                        cnn.Open();
                        cmd = new OdbcCommand("SELECT * FROM ODASMInstallment WHERE InstallmentNo LIKE '" +txtInstallmentNo.Text + "' ",cnn);
                        
                        reader = cmd.ExecuteReader();

                        if (reader.Read())
                        {
                            txtVoucherNo.Text = reader["vOUCHERnO"].ToString();
                            txtRequisitionDate.Text = reader["DateRequisitioned"].ToString();
                            txtInvoiceAmount.Text = reader["TotalRent"].ToString();
                            txtContractNo.Text = reader["ContractNo"].ToString();
                            txtContractYear.Text = reader["ContractYear"].ToString();

                            cmd = new OdbcCommand("SELECT * FROM ODASMVoucherItem WHERE vOUCHERnO LIKE '" +txtVoucherNo.Text + "'", cnn2);
                            cnn2.Close();
                            cnn2.Open();
                            reader2 = cmd.ExecuteReader();

                            if (reader2.Read())
                            {
                                txtTotalVoucherAmount.Text = reader2["AmountPaid"].ToString() + "";
                                txtInvoiceBalance.Text = reader2["Balance"].ToString() + "";
                            }
                            else {
                                txtTotalVoucherAmount.Text = reader["AmountPaid"].ToString() + "";
                                txtInvoiceBalance.Text = reader["Balance"].ToString() + "";
                            }
                            String strPrevPendingContracts;
                            strPrevPendingContracts = getprevPendingContracts();
                            if (strPrevPendingContracts.Length > 0)
                            {
                                /*  MessageBox.Show("There are pending installments (" + strPrevPendingContracts + ") for contract no [" + txtContractNo.Text + "] that need to be updated before you proceed", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                 item.Checked = false;
                                clearRECORD();
                                 txtVoucherItemNo.Text = "";
                                 txtAccountNo.Text = "";
                                 txtLPONo.Text = "";
                                 txtContractNo.Text = "";
                                 cboDocumentNo.Text = "";*/
                                return;
                            }
                            if (GeneralVariables.NewRecord == true)
                            {
                                GeneralVariables.NewRecord = false;
                            }
                            else
                            {
                               // DocumentNoLostFocus();
                            }
                            cnn2.Close();
                        }
                        reader.Close();
                        cnn.Close();
                    }
                    else item.Checked = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }
        private void updatePaymentDetails() {
            try {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();
                if (Convert.ToDouble(txtTotalVoucherAmount.Text) < Convert.ToDouble(txtInvoiceAmount.Text))
                {

                    cmd = new OdbcCommand("UPDATE ODASMInstallment SET PaymentFlag='P',PaymentDue='" + txtInvoiceBalance.Text +
                        "',Balance='" + txtInvoiceBalance.Text +
                        "',AmountPaid='" + txtTotalVoucherAmount.Text + "',PaymentDate='" + DateTime.Today.ToString("yyyy/MM/dd") + "',ChequeDate='" + Convert.ToDateTime(DTPChequeDate.Text).ToString("yyyy/MM/dd") + "',ChequeNo='" + txtChequeNo.Text + "'  WHERE InstallmentNo LIKE '" + txtInstallmentNo.Text + "'", cnn);
                    cmd.ExecuteNonQuery();
                }
                else {
                    cmd = new OdbcCommand("UPDATE ODASMInstallment SET PaymentFlag='Y',PaymentDate='" + DateTime .Today .ToString ("yyyy/MM/dd") +"',ChequeDate='"+ Convert .ToDateTime ( DTPChequeDate.Text ).ToString ("yyyy/MM/dd") + "',ChequeNo='" +txtChequeNo.Text +"'  WHERE InstallmentNo LIKE '" +txtInstallmentNo.Text + "'", cnn);
                    cmd.ExecuteNonQuery();
                
                }
                cnn.Close();
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
        }
        private void saveINSTALLMENTISSUED(){
            try {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn3 = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn4 = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcDataReader rsCONTROL, rsINSTALL, rsSAVE;
                String strINSTALL;
                Double numAmountDue, numBALANCE;
                numAmountDue = 0;
                cnn.Open();
                cnn2.Open();
                cnn3.Open();
                cnn4.Open();
                cmd = new OdbcCommand("SELECT * from ODASMInstallment I where I.VoucherNo = '" +txtVoucherNo.Text  + " '",cnn);
                rsCONTROL = cmd.ExecuteReader();

                cmd = new OdbcCommand("SELECT sum(TotalRent) as Totals from ODASMInstallment I where I.VoucherNo = '" +txtVoucherNo.Text + " ' ",cnn2);
                strINSTALL = cmd.ExecuteScalar().ToString();
                if(strINSTALL ==""){
                    numAmountDue = 0;
                }else {
                    numAmountDue =Convert .ToDouble ( strINSTALL);
                }
                numBALANCE = Convert.ToDouble(txtTotalVoucherAmount .Text);
                numAmountDue = Convert.ToDouble(numAmountDue) - Convert.ToDouble(txtTotalVoucherAmount.Text);
                while (rsCONTROL .Read ()){
                    if (cnn3.State == ConnectionState.Closed)
                    {
                    cnn3.Open ();
                    }
                    cmd = new OdbcCommand("SELECT * from ODASMInstallment where Installment = '" + rsCONTROL["Installment"].ToString () +
                        " ' and  ContractNo = '" + rsCONTROL["ContractNo"].ToString() + " ' and ContractYear = '" + rsCONTROL["ContractYear"].ToString() + " ' and  PaymentMode = '" + rsCONTROL["PaymentMode"].ToString() + " '", cnn3);
                    rsSAVE = cmd.ExecuteReader();
                    if(rsSAVE .Read ()){
                        cmd = new OdbcCommand("UPDATE ODASMInstallment SET Status='CHK-ISSUED',StatusDate='" + DateTime.Today.ToString("yyyy/MM/dd") + "',CurrentPeriod='" + txtCurrentPeriod.Text + "',PaymentDate='" + DateTime.Today.ToString("yyyy/MM/dd") + "' where Installment = '" + rsCONTROL["Installment"].ToString() +
                        " ' and  ContractNo = '" + rsCONTROL["ContractNo"].ToString() + " ' and ContractYear = '" + rsCONTROL["ContractYear"].ToString() + " ' and  PaymentMode = '" + rsCONTROL["PaymentMode"].ToString() + " '", cnn4);
                        cmd.ExecuteNonQuery();
                        cnn4.Close();
                        if (Convert.ToDouble(numAmountDue) >= Convert.ToDouble(txtInvoiceAmount.Text))
                        {
                            cmd = new OdbcCommand("UPDATE ODASMInstallment SET PaymentFlag='Y',Balance='" + 0 + "',AmountPaid='" + Convert.ToDouble(rsCONTROL["TotalRent"].ToString()) + "' where Installment = '" + rsCONTROL["Installment"].ToString() +
                             " ' and  ContractNo = '" + rsCONTROL["ContractNo"].ToString() + " ' and ContractYear = '" + rsCONTROL["ContractYear"].ToString() + " ' and  PaymentMode = '" + rsCONTROL["PaymentMode"].ToString() + " '", cnn4);
                            cnn4.Open();
                            cmd.ExecuteNonQuery();
                            cnn4.Close();
                        }
                        else {
                            if (Convert.ToDouble(rsCONTROL["TotalRent"].ToString()) <= numBALANCE)
                            {
                                cmd = new OdbcCommand("UPDATE ODASMInstallment SET PaymentFlag='N',Balance='" + 0 + "',AmountPaid='" + Convert.ToDouble(rsCONTROL["TotalRent"].ToString()) + "' where Installment = '" + rsCONTROL["Installment"].ToString() +
                                                             " ' and  ContractNo = '" + rsCONTROL["ContractNo"].ToString() + " ' and ContractYear = '" + rsCONTROL["ContractYear"].ToString() + " ' and  PaymentMode = '" + rsCONTROL["PaymentMode"].ToString() + " '", cnn4);
                                numBALANCE = numBALANCE - Convert.ToDouble(rsCONTROL["TotalRent"]);

                                cnn4.Open();
                                cmd.ExecuteNonQuery();
                                cnn4.Close();
                            }
                            else {
                                cmd = new OdbcCommand("UPDATE ODASMInstallment SET PaymentFlag='P',AmountPaid='" + numBALANCE  + "',Balance='" + Convert.ToDouble(rsCONTROL["TotalRent"].ToString()) + "' where Installment = '" + rsCONTROL["Installment"].ToString() +
                                                                " ' and  ContractNo = '" + rsCONTROL["ContractNo"].ToString() + " ' and ContractYear = '" + rsCONTROL["ContractYear"].ToString() + " ' and  PaymentMode = '" + rsCONTROL["PaymentMode"].ToString() + " '", cnn4);
                                numBALANCE = 0;

                                cnn4.Open();
                                cmd.ExecuteNonQuery();
                                cnn4.Close();
                            }
                        }
                    }
                }
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
          }
        private void btnSave_Click(object sender, EventArgs e)
        {
            try {
                if (MessageBox.Show("Are you sure you want to Perform this Action?", "Confirmation Required", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    GeneralVariables GeneralVariables = new GeneralVariables();
                    GeneralVariables.bLoadRecord = true;

                   
                    ListView.CheckedListViewItemCollection checkedItem = listView1.CheckedItems;
                    if (txtAccountNo.Text == "")
                    {
                        MessageBox.Show("Select at least one Record before you proceed", "Infromation Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    else if (txtChequeNo.Text == "")
                    {
                        MessageBox.Show("The Cheque No. is required", "Infromation Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        txtChequeNo.Focus();
                        return;
                    }
                    foreach (ListViewItem item in checkedItem)
                    {

                        if (item.Checked == true)
                        {
                            String strPrevPendingContracts;
                            strPrevPendingContracts = getprevPendingContracts();
                            if (strPrevPendingContracts.Length > 0)
                            {
                                /*  MessageBox.Show("There are pending installments (" + strPrevPendingContracts + ") for contract no [" + txtContractNo.Text + "] that need to be updated before you proceed", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                 item.Checked = false;
                                clearRECORD();
                                 txtVoucherItemNo.Text = "";
                                 txtAccountNo.Text = "";
                                 txtLPONo.Text = "";
                                 txtContractNo.Text = "";
                                 cboDocumentNo.Text = "";*/
                              //  return;
                            }
                            updatePaymentDetails();
                            saveINSTALLMENTISSUED();
                        }
                    }
                    GetRentRequisitioned();
                  //  MessageBox.Show("Successful");

                }
            }catch (Exception ex){
            MessageBox .Show (ex.ToString ());
            }
        }

        private void DTPStartDate_CloseUp(object sender, EventArgs e)
        {
            GetRentRequisitioned();
        }

        private void DTPLastDate_CloseUp(object sender, EventArgs e)
        {
            GetRentRequisitioned();
        }

        private void listView1_MouseClick(object sender, MouseEventArgs e)
        {
            GetRentRequisitioned();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GeneralVariables var = new GeneralVariables();
            var.NewRecord = true;
            ableALLRECORD();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            GeneralVariables VARS = new GeneralVariables();
            VARS.vourcherReport.currentRecord =txtVoucherNo.Text;
            VARS.vourcherReport.ShowDialog();
        }
    }
}
