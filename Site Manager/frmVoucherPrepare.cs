using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Odbc;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Site_Manager
{
    public partial class frmVoucherPrepare : Form
    {
        public frmVoucherPrepare()
        {
            InitializeComponent();
        }
        public Double PartialPaid=0;
        OdbcDataReader  reader;
        OdbcCommand cmd;
        public string CurrentUserName;
        Boolean bLoadRecord;
        String strPrevPendingContracts = "";
        
        private void DisableChilds(Control ctrl)
        {
            foreach (Control c in ctrl.Controls)
            {
                DisableChilds(c);
                if (c is TextBox)
                {
                    ((TextBox)(c)).Enabled = false;
                }
                if (c is CheckBox )
                {
                    ((CheckBox)(c)).Enabled = false;
                }
                if (c is RadioButton )
                {
                    ((RadioButton)(c)).Enabled = false;
                }
                if (c is ComboBox )
                {
                    ((ComboBox)(c)).Enabled = false;
                }
                if (c is DateTimePicker)
                {
                    ((DateTimePicker)(c)).Enabled = false;
                }
            }
        }
        private void AnableChilds(Control ctrl)
        {
            foreach (Control c in ctrl.Controls)
            {
                AnableChilds(c);
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
        private Boolean  ValidateRECORD() {

            if (txtInvoiceAmount.Text == "")
            {
            MessageBox.Show("The Requisition Amount cannot be left Blank","Information Required",MessageBoxButtons .OK );
            txtInvoiceAmount.Focus();
            return false;
            }
            else if (txtRequisitionDate.Text == "")
            {
                MessageBox.Show("The Payment Requisition Date cannot be left Blank", "Information Required", MessageBoxButtons.OK);
                txtRequisitionDate.Focus();
                return false;
            }
           /* else if (cboPaymentCode.Text == "" )
            {
                MessageBox.Show("The Claim Code is Required for all the Transaction", "Information Required", MessageBoxButtons.OK);
                cboPaymentCode.Focus();
                return false;
            }*/
            else if (cboDocumentNo.Text == "")
            {
                MessageBox.Show("The Document No is Used to Determine the Payees details", "Information Required", MessageBoxButtons.OK);
                cboDocumentNo.Focus();
                return false;
            }
            else if (txtCostCenter.Text == "OTP"&& txtRemark .Text =="")
            {
                MessageBox.Show("Please Give Detailed Remarks on This Payment", "Information Required", MessageBoxButtons.OK);
                txtRemark.Focus();
                return false;
            }
            else if (txtCostCenter.Text == "OTP" && txtPayeeDetails.Text == "")
            {
                MessageBox.Show("Enter the Payee Details", "Information Required", MessageBoxButtons.OK);
                txtPayeeDetails.Focus();
                return false;
            }
            else
            {
                return true;
            }
        }
        private void generateVorcherNo()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPLastNumbers WHERE AutoVoucherNo = 'Y'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {




                    int c = Convert.ToInt32(reader["VoucherNo"].ToString()) + 1;

                    switch (Convert.ToInt32(reader["VoucherNo"].ToString().Length))
                    {
                        case 1: txtVoucherNo.Text = reader["VoucherPrefix"].ToString() + "00000" + reader["VoucherNo"].ToString();
                            break;
                        case 2: txtVoucherNo.Text = reader["VoucherPrefix"].ToString() + "0000" + reader["VoucherNo"].ToString();
                            break;
                        case 3: txtVoucherNo.Text = reader["VoucherPrefix"].ToString() + "000" + reader["VoucherNo"].ToString();
                            break;
                        case 4: txtVoucherNo.Text = reader["VoucherPrefix"].ToString() + "00" + reader["VoucherNo"].ToString();
                            break;
                        case 5: txtVoucherNo.Text = reader["VoucherPrefix"].ToString() + "0" + reader["VoucherNo"].ToString();
                            break;
                        case 6: txtVoucherNo.Text = reader["VoucherPrefix"].ToString() + reader["VoucherNo"].ToString();
                            break;
                    }
                 
                    cnn.Close();
                    cnn.Open();
                    cmd = new OdbcCommand("UPDATE ODASPLastNumbers SET VoucherNo='" + c + "' ", cnn);
                    cmd.ExecuteNonQuery();

                }
                cnn.Close();

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }

        }
        private void GenerateVoucherItem()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * from ODASMVoucher where VoucherNo = '" +txtVoucherNo.Text + "'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    if (reader["Items"].ToString() == "")
                    {
                        txtItems.Text = "1";
                    }
                    else {
                        Double c;
                        c = Convert.ToDouble(reader["Items"].ToString())+1;
                        txtItems.Text = c.ToString ();
                    }

                }
                else {
                    txtItems.Text = "1";
                }
                txtVoucherItemNo.Text = txtVoucherNo.Text + "-" + txtItems.Text;
                cnn.Close();

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }

        }
        private void saveVOUCHERITEMS() {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
               
                cnn.Open();
                Double InvoiceAmount;
                cmd = new OdbcCommand("SELECT * from ODASMVOUCHERITEM where VoucherItemNo = '"+txtVoucherItemNo.Text + "'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    InvoiceAmount = Convert.ToDouble(txtInvoiceAmount.Text);
                    txtInvoiceBalance.Text = (InvoiceAmount - PartialPaid).ToString();
                    if (Convert.ToDouble(txtInvoiceBalance.Text) == 0)
                    {
                        cmd = new OdbcCommand("UPDATE ODASMVOUCHERITEM SET PaymentFlag='Y' where VoucherItemNo = '" + txtVoucherItemNo.Text + "'", cnn2);
                        cnn2.Open();
                        cmd.ExecuteNonQuery();
                        cnn2.Close();
                    }
                    else if (Convert.ToDouble(txtInvoiceBalance.Text) == Convert.ToDouble(txtInvoiceAmount.Text))
                    {
                        cmd = new OdbcCommand("UPDATE ODASMVOUCHERITEM SET PaymentFlag='N' where VoucherItemNo = '" + txtVoucherItemNo.Text + "'", cnn2);
                        cnn2.Open();
                        cmd.ExecuteNonQuery();
                        cnn2.Close();
                    }
                    else
                    {
                        cmd = new OdbcCommand("UPDATE ODASMVOUCHERITEM SET PaymentFlag='P' where VoucherItemNo = '" + txtVoucherItemNo.Text + "'", cnn2);
                        cnn2.Open();
                        cmd.ExecuteNonQuery();
                        cnn2.Close();
                    }
                    cmd = new OdbcCommand("UPDATE ODASMVOUCHERITEM SET InvoiceAmount='" + Convert.ToDouble(txtInvoiceAmount.Text) +
                        "',Balance='" + Convert.ToDouble(txtInvoiceBalance.Text) +
                        "',Remarks='" + txtRemark.Text + "',Status='" + txtStatus.Text + "' where VoucherItemNo = '" + txtVoucherItemNo.Text + "'", cnn2);
                    cnn2.Open();
                    cmd.ExecuteNonQuery();
                    cnn2.Close();
                    if (cboPaymentCode.Text == "OTP")
                    {
                        cmd = new OdbcCommand("UPDATE ODASMVOUCHERITEM SET ItemName='" + txtJobDetails.Text + "-" + "Allowances" + "',AmountPaid='" + Convert.ToDouble(txtTotalVoucherAmount.Text) + "' where VoucherItemNo = '" + txtVoucherItemNo.Text + "'", cnn2);
                        cnn2.Open();
                        cmd.ExecuteNonQuery();
                        cnn2.Close();
                    }
                    else
                    {
                        cmd = new OdbcCommand("UPDATE ODASMVOUCHERITEM SET ItemName='" + txtReference.Text + "' where VoucherItemNo = '" + txtVoucherItemNo.Text + "'", cnn2);
                        cnn2.Open();
                        cmd.ExecuteNonQuery();
                        cnn2.Close();
                        if (PartialPaid == 0)
                        {
                            cmd = new OdbcCommand("UPDATE ODASMVOUCHERITEM SET AmountPaid='" + Convert.ToDouble(txtAmountPaid.Text) + "' where VoucherItemNo = '" + txtVoucherItemNo.Text + "'", cnn2);
                            cnn2.Open();
                            cmd.ExecuteNonQuery();
                            cnn2.Close();
                        }
                        else
                        {
                            cmd = new OdbcCommand("UPDATE ODASMVOUCHERITEM SET AmountPaid='" + PartialPaid + "' where VoucherItemNo = '" + txtVoucherItemNo.Text + "'", cnn2);
                            cnn2.Open();
                            cmd.ExecuteNonQuery();
                            cnn2.Close();
                        }
                    }
                }
                else {
                    cmd = new OdbcCommand("INSERT INTO ODASMVOUCHERITEM(vOUCHERnO,VoucherItemNo,Preparedby,dateprepared,LPONo,DocumentNo) VALUES('" + txtVoucherNo.Text + 
                        "','" + txtVoucherItemNo.Text +
                        "','" + CurrentUserName +
                        "','" + DateTime.Today.ToString("MM/dd/yyyy") + "','" + txtLPONo.Text + "','" + cboDocumentNo.Text + "')", cnn2);
                    cnn2.Open();
                    cmd.ExecuteNonQuery();
                    cnn2.Close();
                    cmd = new OdbcCommand("SELECT * from ODASMVOUCHERITEM where VoucherItemNo = '" + txtVoucherItemNo.Text + "'", cnn);
                    reader = cmd.ExecuteReader();

                    if (reader.Read())
                    {
                        InvoiceAmount = Convert.ToDouble(txtInvoiceAmount.Text);
                        txtInvoiceBalance.Text = (InvoiceAmount - PartialPaid).ToString();
                        if (Convert.ToDouble(txtInvoiceBalance.Text) == 0)
                        {
                            cmd = new OdbcCommand("UPDATE ODASMVOUCHERITEM SET PaymentFlag='Y' where VoucherItemNo = '" + txtVoucherItemNo.Text + "'", cnn2);
                            cnn2.Open();
                            cmd.ExecuteNonQuery();
                            cnn2.Close();
                        }
                        else if (Convert.ToDouble(txtInvoiceBalance.Text) == Convert.ToDouble(txtInvoiceAmount.Text))
                        {
                            cmd = new OdbcCommand("UPDATE ODASMVOUCHERITEM SET PaymentFlag='N' where VoucherItemNo = '" + txtVoucherItemNo.Text + "'", cnn2);
                            cnn2.Open();
                            cmd.ExecuteNonQuery();
                            cnn2.Close();
                        }
                        else
                        {
                            cmd = new OdbcCommand("UPDATE ODASMVOUCHERITEM SET PaymentFlag='P' where VoucherItemNo = '" + txtVoucherItemNo.Text + "'", cnn2);
                            cnn2.Open();
                            cmd.ExecuteNonQuery();
                            cnn2.Close();
                        }
                        cmd = new OdbcCommand("UPDATE ODASMVOUCHERITEM SET InvoiceAmount='" + Convert.ToDouble(txtInvoiceAmount.Text) +
                            "',Balance='" + Convert.ToDouble(txtInvoiceBalance.Text) +
                            "',Remarks='" + txtRemark.Text + "',Status='" + txtStatus.Text + "' where VoucherItemNo = '" + txtVoucherItemNo.Text + "'", cnn2);
                        cnn2.Open();
                        cmd.ExecuteNonQuery();
                        cnn2.Close();
                        if (cboPaymentCode.Text == "OTP")
                        {
                            cmd = new OdbcCommand("UPDATE ODASMVOUCHERITEM SET ItemName='" + txtJobDetails.Text + "-" + "Allowances" + "',AmountPaid='" + Convert.ToDouble(txtTotalVoucherAmount.Text) + "' where VoucherItemNo = '" + txtVoucherItemNo.Text + "'", cnn2);
                            cnn2.Open();
                            cmd.ExecuteNonQuery();
                            cnn2.Close();
                        }
                        else
                        {
                            cmd = new OdbcCommand("UPDATE ODASMVOUCHERITEM SET ItemName='" + txtReference.Text + "' where VoucherItemNo = '" + txtVoucherItemNo.Text + "'", cnn2);
                            cnn2.Open();
                            cmd.ExecuteNonQuery();
                            cnn2.Close();
                            if (PartialPaid == 0)
                            {
                                cmd = new OdbcCommand("UPDATE ODASMVOUCHERITEM SET AmountPaid='" + Convert.ToDouble(txtAmountPaid.Text) + "' where VoucherItemNo = '" + txtVoucherItemNo.Text + "'", cnn2);
                                cnn2.Open();
                                cmd.ExecuteNonQuery();
                                cnn2.Close();
                            }
                            else
                            {
                                cmd = new OdbcCommand("UPDATE ODASMVOUCHERITEM SET AmountPaid='" + PartialPaid + "' where VoucherItemNo = '" + txtVoucherItemNo.Text + "'", cnn2);
                                cnn2.Open();
                                cmd.ExecuteNonQuery();
                                cnn2.Close();
                            }
                        }
                    }
                }
               
                cnn.Close();

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }

        
        }
        private void saveRecord() {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
              
                cnn.Open();
                cmd = new OdbcCommand("SELECT * from ODASMVoucher where VoucherNo = '" +txtVoucherNo.Text + "'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    cmd = new OdbcCommand("UPDATE ODASMVoucher SET Amount='" + Convert.ToDouble(txtVoucherAmount.Text) + "',Items='" + txtItems.Text + "',Reference='" + txtReference.Text + "',remark='"+txtPayeeDetails .Text +"' where VoucherNo = '" + txtVoucherNo.Text + "'", cnn2);
                    cnn2.Open();
                    cmd.ExecuteNonQuery();
                    cnn2.Close();
                }
                else
                {
                    cmd = new OdbcCommand("INSERT INTO ODASMVoucher(vOUCHERnO,PaymentCode,CostCenter,"+
                        "VoucherDate,AccountNo,dateprepared,Preparedby,"+
                        "Prepared,ChequePrepared,Authorized,Items,CurrentPeriod) VALUES('" + txtVoucherNo.Text +
                        "','"+cboPaymentCode .Text +
                        "','"+txtCostCenter .Text +
                        "','"+txtRequisitionDate .Text +
                        "','"+txtAccountNo .Text +
                        "','"+DateTime .Today.ToString ("MM/dd/yyyy") +
                        "','"+CurrentUserName  +
                        "','Y','N','N','" + 0 + "','" + getPeriod() + "')", cnn2);
                   
                    cnn2.Open();
                    cmd.ExecuteNonQuery();
                    cnn2.Close();
                    reader.Close();
                    //cnn.Close();
                    cmd = new OdbcCommand("select * from ODASPdefault", cnn);
                    reader = cmd.ExecuteReader();
                    if (reader .Read ()){
                        if (reader["VoucherApproval"].ToString() == "Y")
                        {
                            cmd = new OdbcCommand("UPDATE ODASMVoucher SET ApprovedBy='" + CurrentUserName +
                                "',DateApproved='" + DateTime.Today +
                                "',Status='VCH-APPROVED',Approved='Y'  where VoucherNo = '" + txtVoucherNo.Text + "'", cnn2);
                            cnn2.Open();
                            cmd.ExecuteNonQuery();
                            cnn2.Close();
                        }
                        else {
                            cmd = new OdbcCommand("UPDATE ODASMVoucher SET Status='" + txtStatus.Text + "',Approved='N'  where VoucherNo = '" + txtVoucherNo.Text + "'", cnn2);
                            cnn2.Open();
                            cmd.ExecuteNonQuery();
                            cnn2.Close();
                        }
                    
                    }

                }
                reader.Close();
                   cnn.Close();

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }

        }
        private void saveINVOICE() { 
            try {
         GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
              
                cnn.Open();
                cmd = new OdbcCommand("SELECT * from ODASMInvoice where InvoiceNo = '" +cboDocumentNo.Text +" ' and LPONo = '" +txtLPONo.Text + "'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    Double InvoiceBalance;
                    InvoiceBalance = Convert.ToDouble(reader["InvoiceBalance"].ToString()) - Convert.ToDouble(txtInvoiceAmount.Text); ;
                    cmd = new OdbcCommand("UPDATE ODASMInvoice SET Status='REQ-PREPARED',StatusDate='" + DateTime.Today + "',InvoiceBalance='" + InvoiceBalance + "' where InvoiceNo = '" + cboDocumentNo.Text + " ' and LPONo = '" + txtLPONo.Text + "'", cnn2);
                   
                    cnn2.Open();
                    cmd.ExecuteNonQuery();
                    cnn2.Close();
                    if (Convert.ToDouble(reader["InvoiceBalance"].ToString()) <= 0)
                    {
                        cmd = new OdbcCommand("UPDATE ODASMInvoice SET Requisitioned='Y',DateRequisitioned='" + DateTime.Today + "' where InvoiceNo = '" + cboDocumentNo.Text + " ' and LPONo = '" + txtLPONo.Text + "'", cnn2);

                        cnn2.Open();
                        cmd.ExecuteNonQuery();
                        cnn2.Close();
                    }
                    else {
                        cmd = new OdbcCommand("UPDATE ODASMInvoice SET Requisitioned='N',DateRequisitioned='" + DateTime.Today + "' where InvoiceNo = '" + cboDocumentNo.Text + " ' and LPONo = '" + txtLPONo.Text + "'", cnn2);

                        cnn2.Open();
                        cmd.ExecuteNonQuery();
                        cnn2.Close();
                    }
                }
                reader.Close();
                cnn.Close();
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
        }
        private void saveINSTALLMENT()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();
                cmd = new OdbcCommand("SELECT * from ODASMInstallment where InvoiceNo = '" +cboDocumentNo.Text +" ' and ContractNo = '" +txtLPONo.Text + "'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    cmd = new OdbcCommand("UPDATE ODASMInstallment SET Status='REQ-PREPARED',StatusDate='"+DateTime .Today +
                        "',vOUCHERnO='" + txtVoucherNo.Text + "',CurrentPeriod='" + txtCurrentPeriod.Text + "',VoucherDate='"+txtRequisitionDate .Text +"' where InvoiceNo = '" + cboDocumentNo.Text + " ' and ContractNo = '" + txtLPONo.Text + "'", cnn2);
                    cnn2.Open();
                    cmd.ExecuteNonQuery();
                    cnn2.Close();
                    if(Convert .ToDouble (txtInvoiceBalance .Text )==0){
                        cmd = new OdbcCommand("UPDATE ODASMInstallment SET Requisitioned='Y',DateRequisitioned='" + DateTime.Today + "' where InvoiceNo = '" + cboDocumentNo.Text + " ' and ContractNo = '" + txtLPONo.Text + "'", cnn2);
                        cnn2.Open();
                        cmd.ExecuteNonQuery();
                        cnn2.Close();
                    }else {
                        cmd = new OdbcCommand("UPDATE ODASMInstallment SET Requisitioned='P',DateRequisitioned='" + DateTime.Today + "' where InvoiceNo = '" + cboDocumentNo.Text + " ' and ContractNo = '" + txtLPONo.Text + "'", cnn2);
                        cnn2.Open();
                        cmd.ExecuteNonQuery();
                        cnn2.Close();
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
        private void saveRATESCHEDULE() {

            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();
                cmd = new OdbcCommand("SELECT * from ODASMCouncilRateDue where ReferenceNo = '" +cboDocumentNo.Text + " '", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    Double Balance;
                    Balance = Convert.ToDouble(reader["AmountDue"].ToString())-Convert .ToDouble (txtInvoiceAmount .Text );
                    cmd = new OdbcCommand("UPDATE ODASMCouncilRateDue SET Status='REQ-PREPARED',StatusDate='" + DateTime.Today + "',Balance='"+Balance +"' where ReferenceNo = '" + cboDocumentNo.Text + " '", cnn2);
                    cnn2.Open();
                    cmd.ExecuteNonQuery();
                    cnn2.Close();
                    if (Balance == 0)
                    {
                        cmd = new OdbcCommand("UPDATE ODASMCouncilRateDue SET Requisitioned='Y',DateRequisitioned='" + DateTime.Today + "' where ReferenceNo = '" + cboDocumentNo.Text + " '", cnn2);
                        cnn2.Open();
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                    }
                    else {
                        cmd = new OdbcCommand("UPDATE ODASMCouncilRateDue SET Requisitioned='N',DateRequisitioned='" + DateTime.Today + "' where ReferenceNo = '" + cboDocumentNo.Text + " '", cnn2);
                        cnn2.Open();
                        cmd.ExecuteNonQuery();
                        cnn.Close();
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
        private void updateRECORD() {
           
           if (ValidateRECORD()==true ) {
                
            if (txtVoucherNo .Text ==""){
            generateVorcherNo ();
            
           
            }
            if (cboPaymentCode.Text == "OTP")
            {
                GenerateVoucherItem();
                saveVOUCHERITEMS();
            }
            else
            {
                GenerateVoucherItem();
            
                saveVOUCHERITEMS();
                computeVOUCHERTOTAL();
               
            }
            countVOUCHERITEMS();
                saveRecord();
               // MessageBox.Show("xx");
            GetInvoicesREQUISITIONED();
          
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPCostCentre WHERE CostCentre = '" +txtCostCenter.Text + "'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    if (reader["Materials"].ToString()=="Y")
                    {
                        saveINVOICE();
                        GetInvoicesNotPaid();
                    }
                    else if (reader["Machinery"].ToString() == "Y")
                    {
                        saveINVOICE();
                        GetInvoicesNotPaid();
                    }
                    else if (reader["Rent"].ToString() == "Y")
                    {
                        saveINSTALLMENT();
                       
                    }
                    else if (reader["Rate"].ToString() == "Y")
                    {
                        saveRATESCHEDULE();
                        GetRateNotPaid();
                        
                    }
                    else if (reader["OtherCosts"].ToString() == "Y")
                    {
                        UpdateCostsAdded();
                        UpdateAdministrationCosts();
                        showALLOTHERREQUISITIONS();
                        
                    }

                   
                }
                disableALLRECORD();
            }
        }
        private void UpdateAdministrationCosts(){
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPAdminCosting WHERE CostItem = '" +txtReference.Text  +"'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    cmd = new OdbcCommand("UPDATE ODASPAdminCosting SET  Amount='" + Convert.ToDouble(txtAmountPaid.Text) + "' WHERE CostItem = '" + txtReference.Text + "'", cnn2);
                    cnn2.Open();
                    cmd.ExecuteNonQuery();
                    cnn2.Close();
                }
                else {
                    cmd = new OdbcCommand("INSERT INTO ODASPAdminCosting (CostItem,JobBriefNo,Approved,Authorized,valid,Preparedby,dateprepared,Amount) VALUES('" + txtCouncilCode .Text+ 
                        "','"+txtJobCardNo .Text +"','N','N','Y','"+CurrentUserName  +"','"+DateTime .Today +"','"+Convert .ToDouble (txtAmountPaid .Text )+"')", cnn2);
                    cnn2.Open();
                    cmd.ExecuteNonQuery();
                    cnn2.Close();
                }
                reader.Close();
                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
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
        private void enableALLRECORD() { 
             try
            {
                AnableChilds(this);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        
        }
        private void clearRECORD(){

        
            txtAmountPaid.Text = "0";
            txtInvoiceAmount.Text = "0";
            txtInvoiceBalance.Text = "0";
        }
        private void LoadRequisition()
        {
            try {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("Select * from ODASMVoucher  Where ODASMVoucher.VoucherNo = '" +txtVoucherNo.Text + "'", cnn);
                reader = cmd.ExecuteReader();
                if (reader .Read ()){
                    cboPaymentCode.Text = reader["PaymentCode"].ToString();
                    txtCostCenter.Text = reader["CostCenter"].ToString();
                    txtInvoiceAmount.Text = reader["Amount"].ToString();
                    txtReference.Text = reader["Reference"].ToString()+"";
                    txtRequisitionDate.Text = reader["VoucherDate"].ToString();
                    txtStatus.Text = reader["Status"].ToString()+"";
                }
                reader.Close();
                cnn.Close();
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
               
               
    }
        private void loadClaimDescription()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPPaymentCode WHERE PaymentCode = '" +cboPaymentCode.Text + "' ", cnn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    cboPaymentCode.Text = reader["PaymentCode"].ToString();
                    txtCostCenter.Text = reader["CostCenter"].ToString();
                    txtPaymentCodeDescription.Text = reader["PaymentCodeDescription"].ToString();
                }
                reader.Close();
                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


        }
        private void UpdateCostsAdded() {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();
                cmd = new OdbcCommand("SELECT * from ODASMJobCard where JobCardNo = '" +txtJobCardNo.Text + " '", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                   
                    if (Convert .ToDouble ( reader ["TotalCost"].ToString ()) == 0)
                    {
                        cmd = new OdbcCommand("UPDATE ODASMJobCard SET TotalCost='" + 0 + "' where JobCardNo = '" + txtJobCardNo.Text + " '", cnn2);
                        cnn2.Open();
                        cmd.ExecuteNonQuery();
                        cnn2.Close();
                    }
                    Double TotalCost;
                    TotalCost = Convert.ToDouble(reader["TotalCost"].ToString()) + Convert.ToDouble(txtTotalVoucherAmount.Text );
                    cmd = new OdbcCommand("UPDATE ODASMJobCard SET TotalCost='" + TotalCost + "' where JobCardNo = '" + txtJobCardNo.Text + " '", cnn2);
                        cnn2.Open();
                        cmd.ExecuteNonQuery();
                        cnn2.Close();
                    
                }
                reader.Close();
                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void loadPAYMENTRECORD()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPPaymentCode WHERE PaymentCode = '" +cboPaymentCode.Text + "' ", cnn);
                reader = cmd.ExecuteReader();
                
                if (reader.Read())
                {
                   
                    txtReference.Text = reader["PaymentCodeDescription"].ToString().Trim ();
                    txtCostCenter.Text = reader["CostCenter"].ToString();
                    txtPaymentCodeDescription.Text = reader["PaymentCodeDescription"].ToString();
                }
                reader.Close();
                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


        }
        private String  getprevPendingContracts()
        {
            String str = "";
            try
            {
                
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASMInstallment WHERE  ContractNo='" +txtContractNo.Text + "' AND ContractYear<" +Convert .ToInt32 (txtContractYear.Text) + " AND (Requisitioned='N' OR Requisitioned IS NULL) ", cnn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    str = str + reader["ContractYear"].ToString() + "- [" + reader["PaymentDueDate"].ToString() + "]";
                    MessageBox.Show(str + " " + str.Length.ToString() + " " + reader["Requisitioned"].ToString());
                }
                cnn.Close();
                
            }
                
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString ());
            }
          
            return str;
        }
        
              private void loadLANDLORDDETAILS()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("select * from ODASMInstallment I, ODASPAccount A where I.InvoiceNo = '" +cboDocumentNo.Text  +"' and I.AccountNo =  A.AccountNo", cnn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    txtInvoiceAmount.Text = reader["PaymentDue"].ToString();
                    txtAmountPaid.Text = reader["PaymentDue"].ToString();
                    txtStatus.Text = "REQ-PREP";
                    txtInvoiceBalance.Text = "0";
                    txtAccountNo.Text = reader["AccountNo"].ToString();
                    txtPayeeDetails.Text = reader["CompanyName"].ToString();
                    txtProductCode.Text = "N/A";
                    if (txtCostCenter.Text == "RENT")
                    {

                    }
                    else {
                        txtReference.Text = txtPaymentCodeDescription.Text.Trim () + "-" + cboDocumentNo.Text.Trim ();
                    }
                
                }
                reader.Close();
                cnn.Close();
            }
            catch (Exception ex) {
                MessageBox.Show(ex.ToString ());
            }
         }

        private void loadINVOICEDETAILS()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("select * from ODASMInvoice I, ODASMRequisition R, ODASPAccount A where I.InvoiceNo = '" + cboDocumentNo.Text  + "' and I.LPONo = R.LPONo and R.AccountNo =  A.AccountNo", cnn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    txtInvoiceAmount.Text = reader["PriceInclusive"].ToString();
                    txtVoucherAmount.Text = reader["PriceInclusive"].ToString();
                    txtStatus.Text = "REQ-PREP";
                    txtReference.Text = txtPaymentCodeDescription.Text +"- "+ cboDocumentNo.Text.Trim();
                    txtAccountNo.Text = reader["AccountNo"].ToString();
                   
                }
                reader.Close();
                cnn.Close();
            }
            catch (Exception ex) {
                MessageBox.Show(ex.ToString ());
            }
         }
        private void loadCOUNCILDETAILS()
        {
            try
            {
                String str;
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcDataReader invoiceReader=null ;
                cnn.Open();
                cnn2.Open();
                cmd = new OdbcCommand("Select * From ODASMCouncilRateDue Where ReferenceNo like '" +cboDocumentNo.Text  + "'", cnn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    if (reader["Face"].ToString() == "Y")
                    {
                        cmd = new OdbcCommand("select * from ODASPPlotSite S, ODASPPlot P,ODASPCouncil C  where P.PlotNo = S.PlotNo and P.CouncilCode = C.CouncilCode and S.SiteNo  = '" + reader["SiteNo"].ToString() + "'", cnn2);
                        invoiceReader = cmd.ExecuteReader();
                    }
                    else if (reader["BillBoard"].ToString() == "Y")
                    {
                        cmd = new OdbcCommand("select * from ODASPPlotMast S, ODASPPlot P,ODASPCouncil C  where P.PlotNo = S.PlotNo and P.CouncilCode = C.CouncilCode and S.MastNo  = '" + reader["SiteNo"].ToString() + "'", cnn2);
                        invoiceReader = cmd.ExecuteReader();
                    }
                    else if (reader["BillBoard"].ToString() == "" && reader["Face"].ToString() == "")
                       
                    {
                        return;
                    }
                    if(invoiceReader .Read ()){
                        txtInvoiceAmount.Text = invoiceReader["AmountDue"].ToString();
                        txtVoucherAmount.Text = invoiceReader["AmountDue"].ToString();
                        txtInvoiceBalance.Text = "0";
                        txtStatus.Text = "REQ-PREP";
                        txtReference.Text = txtPaymentCodeDescription.Text + "- " + cboDocumentNo.Text;
                        txtCouncilCode.Text = invoiceReader["CouncilCode"].ToString();
                        reader.Close();
                        cnn.Close();
                        cnn.Open();
                        cmd = new OdbcCommand("select * from ODASMjobBrief JB, ODASMjobBriefItems JBI where JB.jobBriefNo = JBI.JobBriefNo and JBI.JobBriefitemNo = '" + invoiceReader["JobBriefItemNo"].ToString () + "'",cnn);
                        reader = cmd.ExecuteReader();
                        if (reader.Read())
                        {
                            txtProductCode.Text = reader["ProductCode"].ToString();

                        }
                        else {
                            txtProductCode.Text = "N/A";
                        }
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
        private void DocumentNoLostFocusLOAN() {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("select * from ALISMLoanManagement, ODASMJobBrief, ODASPAccount where ALISMLoanManagement.LoanNo = '" +cboDocumentNo.Text +"' and ALISMLoanManagement.JobBriefNo = ODASMJobBrief.JobBriefNo and ODASMJobBrief.AccountNo LIKE ODASPAccount.AccountNo", cnn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    txtInvoiceAmount.Text = reader["LoanAmount"].ToString();
                    txtVoucherAmount.Text = reader["PriceInclusive"].ToString();
                    txtStatus.Text = "REQ-PREP";
                    txtReference.Text = txtReference.Text + "- " + cboDocumentNo.Text.Trim();
                    txtPayeeDetails.Text = reader["othernames"].ToString()+" "+ reader["CompanyName"].ToString();
                     
                }
                reader.Close();
                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void DocumentNoLostFocus()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPCostCentre WHERE CostCentre = '" + txtCostCenter.Text + "' ", cnn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {

                    if (reader["AdminCost"].ToString().Trim() == "Y")
                    {
                        DocumentNoLostFocusLOAN();
                    }
                    else if (reader["Machinery"].ToString().Trim() == "Y")
                    {
                       
                    }
                    else if (reader["Materials"].ToString().Trim() == "Y")
                    {

                       loadINVOICEDETAILS();
                    }
                    else if (reader["Rent"].ToString().Trim() == "Y")
                    {

                        loadLANDLORDDETAILS();
                    }
                    else if (reader["rate"].ToString().Trim() == "Y")
                    {
                        loadCOUNCILDETAILS();
                    }
                    else if (reader["OtherCosts"].ToString().Trim() == "Y")
                    {
                    
                    }
                    else if (reader["ManPower"].ToString()=="Y")
                    {
                    
                    }

                }
                if (GeneralVariables.bMedicalRequisition == true)
                {
                    //loadSpecificMEDICALGRID();
                }
                reader.Close();
                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


        }
        private void SelectCostCenter()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPCostCentre WHERE CostCentre = '" +txtCostCenter.Text + "' ", cnn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                   
                    if (reader["AdminCost"].ToString().Trim() == "Y") {
                        
                    }
                    else if (reader["Machinery"].ToString().Trim() == "Y")
                    {
                        
                    }
                    else if (reader["Materials"].ToString().Trim() == "Y")
                    {
                       
                        GetInvoicesNotPaid();
                    }
                    else if (reader["Rent"].ToString().Trim() == "Y")
                    {
                        
                        GetRentNotRequisitioned();
                    }
                    else if (reader["rate"].ToString().Trim() == "Y")
                    {
                        
                        GetRateNotPaid();
                    }
                    else if (reader["OtherCosts"].ToString().Trim() == "Y")
                    {
                        
                        showALLOTHERREQUISITIONS();
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
        private void GetInvoicesNotPaid() {
            try
            {
                listView1.Columns.Clear();
                listView1.Items.Clear();
                listView1.Columns.Add("Invoice No",listView1 .Width /6);
                listView1.Columns.Add("LPO No", listView1.Width / 6);
                listView1.Columns.Add("Invoice Date", listView1.Width / 6);
                listView1.Columns.Add("Job Brief No", listView1.Width / 6);
                listView1.Columns.Add("Account No", listView1.Width / 6);
                listView1.Columns.Add("Sublier", listView1.Width / 6);
              
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("Select *  from ODASMInvoice I, ODASPAccount A Where (I.Requisitioned = 'N' or I.Requisitioned is null) and I.AccountNo = A.AccountNo", cnn);
                reader = cmd.ExecuteReader();
                listView1.Items.Clear();
                if (reader.Read())
                {
                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["InvoiceNo"].ToString());
                        if (reader["LPONo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["LPONo"].ToString());
                        }
                        if (reader["InvoiceDate"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["InvoiceDate"].ToString());
                        }
                        if (reader["JobCardNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["JobCardNo"].ToString());
                        }
                        if (reader["AccountNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["AccountNo"].ToString());
                        }
                        if (reader["CompanyName"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["CompanyName"].ToString());
                        }
                        listView1.Items.Add(lv3);




                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void GetRateNotPaid()
        {
            try
            {
                listView1.Columns.Clear();
                listView1.Items.Clear();
                listView1.Columns.Add("Reference No", listView1.Width / 6);
                listView1.Columns.Add("Site No", listView1.Width / 6);
                listView1.Columns.Add("Start Date", listView1.Width / 6);
                listView1.Columns.Add("Amount Due", listView1.Width / 6);
                listView1.Columns.Add("End Date", listView1.Width / 6);
                listView1.Columns.Add("Date Due", listView1.Width / 6);
                listView1.Columns.Add("Job Brief", listView1.Width / 6);

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("Select *  from ODASMCouncilRatedue C Where C.AmountDue > 0 and C.DueDate <= '" +DateTime .Today .ToString ("yyy/MM/dd")+ "'  and C.Requisitioned = 'N' and C.Paid = 'N'", cnn);
                reader = cmd.ExecuteReader();
                listView1.Items.Clear();
                if (reader.Read())
                {
                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["ReferenceNo"].ToString());
                        if (reader["SiteNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["SiteNo"].ToString());
                        }
                        if (reader["StartDate"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["StartDate"].ToString());
                        }
                        if (reader["EndDate"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["EndDate"].ToString());
                        }
                        if (reader["DueDate"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["DueDate"].ToString());
                        }
                        if (reader["AmountDue"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["AmountDue"].ToString());
                        }
                        if (reader["JobBriefItemNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["JobBriefItemNo"].ToString());
                        }
                        listView1.Items.Add(lv3);




                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void loadAllowance()
        {
            try
            {
                
              
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               
                cnn.Open();

                cmd = new OdbcCommand("SELECT * FROM ODASPAdminCosting WHERE CostItem = '" +txtReference.Text + "' ", cnn);

               
                reader = cmd.ExecuteReader();
          
                if (reader.Read ())
                {
                    txtPaymentDescription.Text = reader["CostingItemName"].ToString();
                }
                reader.Close();
                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        private void loadJobBriefAccountNo()
        {
            try
            {


                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();

                cmd = new OdbcCommand("SELECT * FROM ODASMJobBrief WHERE JobBriefNo = '" +txtJobCardNo.Text +"'  ", cnn);


                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    txtAccountNo.Text = reader["AccountNo"].ToString();
                }
                reader.Close();
                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void loadAccountName()
        {
            try
            {


                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();

                cmd = new OdbcCommand("SELECT * FROM ODASPAccount WHERE AccountNo = '" + txtAccountNo.Text + "'   ", cnn);


                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    txtJobDetails.Text = reader["CompanyName"].ToString();
                } 
                reader.Close();
                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void GetRentNotRequisitioned()
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
                listView1.Columns.Add("Plot Details", listView1.Width / 6);
                listView1.Columns.Add("Installment No", listView1.Width / 6);
                ListView.ListViewItemCollection items = listView1.Items;
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcCommand cmd2;
                cnn.Open();
                cnn2.Open ();
                cmd2 = new OdbcCommand("Select COUNT(*) from ODASMInstallment I, ODASPAccount A,ODASPPlot P Where P.PlotNo=I.PlotNo AND (I.Requisitioned = 'N' or I.Requisitioned is null or I.Requisitioned = 'Y' ) and A.AccountNo = P.AccountNo  AND (I.PaymentDueDate>='" + DTPStartDate.Text  + "' AND I.PaymentDueDate<='" +DTPLastDate.Text + "') AND I.Balance>0", cnn);
                cmd = new OdbcCommand("Select * from ODASMInstallment I, ODASPAccount A,ODASPPlot P Where P.PlotNo=I.PlotNo AND (I.Requisitioned = 'N' or I.Requisitioned is null or I.Requisitioned = 'Y' ) and A.AccountNo = P.AccountNo  AND (I.PaymentDueDate>='" + DTPStartDate.Text + "' AND I.PaymentDueDate<='" + DTPLastDate.Text + "') AND I.Balance>0", cnn2);
              
             
             
               String  c = cmd2.ExecuteScalar().ToString ();
                 reader = cmd.ExecuteReader();
                
               
              
                if (reader.Read()  )
                {
                   
                    progressBar1.Visible = true;
                    progressBar1.Value = 0;
                    progressBar1.Minimum = 0;
                    progressBar1.Maximum = Convert.ToInt32(c); 
                     
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
                            lv3.SubItems.Add(Convert.ToDateTime(dat).ToString ("MM/dd/yyyy"));
                        }
                        if (reader["PaymentDue"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PaymentDue"].ToString());
                        }
                        if (reader["AccountNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["AccountNo"].ToString());
                        }
                        if (reader["CompanyName"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["CompanyName"].ToString());
                        }
                        
                        
                        
                        foreach (ListViewItem i in items)
                        {
                          //  i.SubItems[6].Text = "[ LR - " + reader["LRNo"].ToString() + " ] " + reader["PhysicalLocation"].ToString() + ". REF. No " + reader["ContractNo"].ToString();
                           /// i.SubItems[6].Text = reader["InstallmentNo"].ToString();
                        }
                        
                        
                            lv3.SubItems.Add("[ LR - " + reader["LRNo"].ToString() + " ] " + reader["PhysicalLocation"].ToString() + ". REF. No " + reader["ContractNo"].ToString());
                          
                        
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
                MessageBox.Show(ex.ToString ());
            }
        }
        private void showALLOTHERREQUISITIONS()
        {
            try
            {
                listView1.Columns.Clear();
                listView1.Items.Clear();
                listView1.Columns.Add("Request No", listView1.Width / 5);
                listView1.Columns.Add("Request", listView1.Width / 5);
                listView1.Columns.Add("Job Card", listView1.Width / 5);
                listView1.Columns.Add("Amount", listView1.Width / 5);
                listView1.Columns.Add("Requisition No", listView1.Width / 5);
                listView1.Columns.Add("Remarks", listView1.Width / 5);
               
              
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               
                cnn.Open();
                 cmd = new OdbcCommand("SELECT * FROM ODASMRequisitionItems Where Request='Y' and Authorized = 'Y' and approved = 'Y' and (issued = 'N' or issued is null)", cnn);

               
                reader = cmd.ExecuteReader();
                listView1.Items.Clear();
                if (reader.HasRows)
                {
                    

                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["ItemNo"].ToString());
                        if (reader["ProductCode"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["ProductCode"].ToString());
                        }
                        if (reader["JobCardNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["JobCardNo"].ToString());
                        }
                        if (reader["AllowanceTotals"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["AllowanceTotals"].ToString());
                        }
                        if (reader["RequisitionNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["RequisitionNo"].ToString());
                        }
                        if (reader["RequestPurpose"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["RequestPurpose"].ToString());
                        }
                      
                        listView1.Items.Add(lv3);

                       


                    }
                    reader.Close();
                    cnn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void  GetInvoicesREQUISITIONED()
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
                  cmd = new OdbcCommand("Select *  from ODASMVoucherItem Where VoucherNo = '" +txtVoucherNo.Text  + "'", cnn);

               
                reader = cmd.ExecuteReader();
                listView2.Items.Clear();
                if (reader.HasRows )
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
                MessageBox.Show(ex.ToString ());
            }
        }
        private void loadPaymentDescription()
        {
            try
            {
                
              
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               
                cnn.Open();

                cmd = new OdbcCommand("SELECT * FROM ODASPCostCentre WHERE CostCentre = '" +txtCostCenter.Text + "'", cnn);

               
                reader = cmd.ExecuteReader();
          
                if (reader.Read ())
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
       
        private void  computeVOUCHERTOTAL()
        {
            try
            {
                
              
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               
                cnn.Open();

                cmd = new OdbcCommand("SELECT sum(AmountPaid) as totals from ODASMVoucherItem where VoucherNo = '" +txtVoucherNo.Text + "'", cnn);

                String c;
                c = cmd.ExecuteScalar ().ToString ();

               
                    if (c == "")
                    {
                        if (PartialPaid > 0)
                        {
                            txtVoucherAmount.Text = PartialPaid.ToString();
                        }
                        else
                        {
                            txtVoucherAmount.Text = Convert.ToDouble(txtAmountPaid.Text).ToString();
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

                cmd = new OdbcCommand("SELECT count(VoucherNo) as totals from ODASMVoucherItem where VoucherNo = '" +txtVoucherNo.Text + "'", cnn);

                String  c;
                c = cmd.ExecuteScalar().ToString();

                    if( c == "")
                    {
                        txtItems.Text = "0";
                    }
                
                else { txtItems.Text =  c; }
               
                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        public String getPeriod()
        {

            String strMonth, strYr, result;
            strMonth = DateTime.Today .Month.ToString().Trim();
            if (strMonth.Length == 1)
            {

                strMonth = "0" + strMonth;

            }
            strYr = DateTime.Today.Year.ToString().Trim();
            result = strYr.Trim() + "/" + strMonth.Trim();
            return result;
        }
        
        private void frmVoucherPrepare_Load(object sender, EventArgs e)
        {
            try {
               
                disableALLRECORD();
                GeneralVariables GeneralVariables = new GeneralVariables();
                cboPaymentCode = new TextBox ();
              
               if(cboPaymentCode .Text =="OTP"){
                   txtPayeeDetails.Enabled = true;
               }

                
                clearRECORD();
                txtPayeeDetails.Enabled = false;
                getPeriod();
                txtCurrentPeriod.Text = getPeriod();
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void cboPaymentCode_TextChanged(object sender, EventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();
            if (GeneralVariables.bapproveREQUISITION == true)
            {
                LoadRequisition();
                loadClaimDescription();

            }
            else
            {
                loadPAYMENTRECORD();
                //SelectCostCenter();
                GetInvoicesREQUISITIONED();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();

            GeneralVariables.NewRecord = true;
            enableALLRECORD();
            txtVoucherNo.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (listView1 .CheckedItems .Count ==0)
                { 
                MessageBox.Show("Please select a record?", "Information Required", MessageBoxButtons.OK , MessageBoxIcon.Error );
                }
                else if (MessageBox.Show("Are you sure you want to Perform this Action?", "Confirmation Required", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    GeneralVariables GeneralVariables = new GeneralVariables();
                    GeneralVariables.bLoadRecord = true;
                     if (Convert.ToDouble(txtInvoiceAmount.Text) > Convert.ToDouble(txtAmountPaid.Text))
                    {
                        PartialPaid = Convert.ToDouble(txtAmountPaid.Text);
                    }
                    ListView.CheckedListViewItemCollection checkedItem = listView1.CheckedItems;
                    if (cboDocumentNo.Text == "")
                        {
                            MessageBox.Show("Select at least one Record before you proceed", "Infromation Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return;
                        }
                    foreach (ListViewItem item in checkedItem)
                    {
                       
                          if (item.Checked == true)
                        {
                            txtVoucherNo.Text = "";
                            //  bLoadRecord = false;
                            if (strPrevPendingContracts.Length > 0)
                            {
                              //  MessageBox.Show("There are pending installments (" + strPrevPendingContracts + ") for contract no [" + txtContractNo.Text + "] that need to be updated before you proceed", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                              
                               // return;
                            }
                            updateRECORD();

                        }
                        
                        }
                        SelectCostCenter();


                        Cursor.Current = Cursors.Default ;
                    
                }
                   }
            catch (Exception ex) {
                MessageBox.Show(ex.ToString ());
            }
        }

        private void DTPStartDate_ValueChanged(object sender, EventArgs e)
        {
       //     SelectCostCenter();
        }

        private void DTPLastDate_ValueChanged(object sender, EventArgs e)
        {
          //  SelectCostCenter();
        }

        private void listView1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            try
            {
               
                Double i, j;
                ListView.CheckedListViewItemCollection checkedItems = listView1.CheckedItems;
                foreach (ListViewItem item in checkedItems  ){

                    if (item.Checked == true)
                    {
                        j = checkedItems.Count;
                        if (j == 0)
                        {
                            return;
                        }
                        if (txtCostCenter.Text == "OTP")
                        {
                            cboDocumentNo.Text = item.Text;
                            txtAmountPaid.Text = item.SubItems[3].Text;
                            txtInvoiceAmount.Text = item.SubItems[3].Text;
                            txtJobCardNo.Text = item.SubItems[2].Text;
                            txtRemark.Text = item.SubItems[5].Text;
                            txtCouncilCode.Text = item.SubItems[1].Text;
                            txtReference.Text = item.SubItems[1].Text;
                            txtReqNo.Text = item.SubItems[4].Text;
                            loadAllowance();
                            loadJobBriefAccountNo();
                            loadAccountName();
                        }
                        else
                        {
                            txtAccountNo.Text = item.SubItems[4].Text;
                            cboDocumentNo.Text = item.Text;
                            txtLPONo.Text = item.SubItems[1].Text;
                            txtExpiryDate.Text = item.SubItems[2].Text;
                            txtCouncilCode.Text = "";
                            txtAmountPaid.Text = item.SubItems[3].Text;
                            txtInvoiceAmount.Text = item.SubItems[3].Text;
                            txtVoucherAmount.Text = "0";
                            txtInvoiceBalance.Text = "0";
                            if (bLoadRecord == true)
                            {
                                txtReference.Text = item.SubItems[6].Text;
                            }
                            txtInstallmentNo.Text = item.SubItems[7].Text;
                            GeneralVariables GeneralVariables = new GeneralVariables();
                            OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);

                            cnn.Open();

                            cmd = new OdbcCommand("SELECT * FROM ODASMInstallment WHERE InstallmentNo LIKE '%" + txtInstallmentNo.Text + "%'    ", cnn);


                            reader = cmd.ExecuteReader();

                            if (reader.Read())
                            {
                                txtContractNo.Text = reader["ContractNo"].ToString();
                                txtContractYear.Text = reader["ContractYear"].ToString();
                            }
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
                                GeneralVariables.NewRecord = false ;
                            }
                            else {
                                DocumentNoLostFocus();
                            }
                            reader.Close();
                            cnn.Close();
                        }
                    }
                    else item.Checked = false;
                }
            }
            catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
        }

        private void chkSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
        ListView.CheckedListViewItemCollection checkedItems = listView1.CheckedItems;
       ListView .ListViewItemCollection listItems= listView1.Items ;
       foreach (ListViewItem item in listItems)
        {
            if (chkSelectAll.Checked == true)
            {
                item.Checked = true;
            }
            else item.Checked = false;
        }
            }
            catch (Exception ex) {
                MessageBox.Show (ex.ToString ());
            }
        }

        private void txtAmountPaid_TextChanged(object sender, EventArgs e)
        {
            try {
                GeneralVariables var = new GeneralVariables();
                if (var.NewRecord == true)
                {
                    return;
                }
                computeVOUCHERTOTAL();
                if(txtInvoiceAmount .Text ==""){
                return ;
                }
                if (txtAmountPaid.Text == "")
                { 
                    return;
                }
                if (txtCostCenter.Text == "OTP")
                {
                    txtInvoiceBalance.Text = "0";
                    txtInvoiceBalance.Enabled = false;

                }
                else { 
                if(txtVoucherAmount .Text ==""){
                    txtVoucherAmount.Text = "0";

                }
                txtInvoiceBalance.Text =( Convert.ToDouble(txtInvoiceAmount.Text) - Convert.ToDouble(txtAmountPaid.Text )).ToString ();
                txtVoucherAmount.Text = (Convert.ToDouble(txtAmountPaid .Text )).ToString ();
                }
            }catch (Exception ex){
                MessageBox.Show(ex.ToString());
            }
        }

        private void DTPLastDate_CloseUp(object sender, EventArgs e)
        {
           SelectCostCenter();
        }

        private void DTPStartDate_CloseUp(object sender, EventArgs e)
        {
            SelectCostCenter();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            GeneralVariables VARS = new GeneralVariables();
            VARS.vourcherReport.currentRecord  = txtVoucherNo.Text;
            VARS.vourcherReport.ShowDialog();
        }
        private void anableAll(Control ctrl)
        {
            foreach (Control c in ctrl.Controls)
            {
                anableAll(c);
                if (c is TextBox)
                {
                    ((TextBox)(c)).Enabled = true;
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
                if (c is NumericUpDown)
                {
                    ((NumericUpDown)(c)).Enabled = true;
                }
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            anableAll(this );
        }
    }
}
