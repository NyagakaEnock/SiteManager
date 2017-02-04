using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.Odbc;
namespace Site_Manager
{
    public partial class frmLandLord : Form
    {
        public frmLandLord()
        {
            InitializeComponent();
        }
      
        OdbcCommand cmd;
        OdbcDataReader reader;
        DataTable dTable;
        OdbcDataReader odbcrdr;
        DataSet ds;
        Boolean search=false ;
        OdbcDataAdapter da;
        private void loadTowns()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();

                ds = new DataSet();
                cmd = new OdbcCommand("SELECT Town AS SelectField FROM ODASPTown ORDER BY Town", cnn);
                da = new OdbcDataAdapter(cmd);

                da.Fill(ds, "ODASPTown");
                dTable = ds.Tables[0];
                cboTownCode.Items.Clear();

                foreach (DataRow drow in dTable.Rows)
                {
                    cboTownCode.Items.Add(drow["SelectField"].ToString());
                }
                cnn.Close();
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }
        private void LoadAccountType()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPAccountType WHERE AccountType = '" +cboAccountType.Text + "'", cnn);
                reader = cmd.ExecuteReader();
                if(reader .Read ()){
                    txtAccountTypeDescription.Text = reader["AccountTypeDescription"].ToString();
                }
                cnn.Close();
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }

        private void cboTownCode_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT CouncilCode FROM ODASPCouncil WHERE Council='" + cboTownCode.SelectedItem.ToString() + "'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {


                    txtTownDescription.Text = reader["CouncilCode"].ToString();

                }
                cnn.Close();

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.Message);

            }
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
        private void loadDefault() { 
        txtDateCreated.Text = DateTime .Today .ToString ("MM/dd/yyyy");
        txtPostalAddress.Text = "P. O. Box ";
        txtPhysicalAddress.Text = "XX";
        cboTownCode.Text = "NAIROBI";
        txtTelephoneNo.Text = "XX";
        txtTownDescription.Text = "NBI";
        txtStatus.Text = "ACTIVE";
        txtDateCreated.Text = DateTime.Today.ToString("MM/dd/yyyy");
        txtContactDepartment.Text = "XX";
        txtContactDesignation.Text = "XX";
        txtContactName.Text = "XX";
        txtMobileNo.Text = "XX";
        txtTelephoneExtention.Text = "XX";
        txtemailAddress.Text = "XX";
        }
        private void getLANDLORDTYPE() { 
         
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT *  FROM ODASPAccountType WHERE AccountType = 'LLORD'", cnn);
                reader = cmd.ExecuteReader();
                if(reader .Read ()){
                    cboAccountType.Text = reader["AccountTypeDescription"].ToString();
                }
                cnn.Close();
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }
        private void getLANDLORDS()
        {
            try
            {
                string sql;
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn2.Open();
                cnn.Open();
               if (search == true)
                {
                    sql = "SELECT * FROM ODASPAccount Where Status = 'A' AND AccountType = 'LLORD' AND CompanyName LIKE '%"+txtSearch .Text +"%' oRDER BY AccountNo";

                }
                else
                {
                    sql = "SELECT * FROM ODASPAccount Where Status = 'A' AND AccountType = 'LLORD' oRDER BY AccountNo";
                }
               //sql = "SELECT * FROM ODASPAccount Where Status = 'A' AND AccountType = 'LLORD' oRDER BY AccountNo";
             
                cmd = new OdbcCommand(sql, cnn);
                odbcrdr = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();
                listView1.Columns.Add("Land Lord No", listView1.Width / 3);
                listView1.Columns.Add("Names", listView1.Width / 3);
                listView1.Columns.Add("Status", listView1.Width / 3);


                if (odbcrdr.HasRows)
                {
                    while (odbcrdr.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(odbcrdr["AccountNo"].ToString());
                        if (odbcrdr["CompanyName"].ToString() != "")
                        {
                            lv3.SubItems.Add(odbcrdr["CompanyName"].ToString());
                        }
                        if (odbcrdr["Status"].ToString() != "")
                        {
                            lv3.SubItems.Add(odbcrdr["Status"].ToString());
                        }


                        listView1.Items.Add(lv3);
                    }

                }
                odbcrdr.Close();
                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString ());
            }
        }
      private void showALLSITESByLandlord()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn2.Open();
                cnn.Open();
                OdbcDataReader RDR;
                string sql = "SELECT * FROM ODASPAccount, ODASpPlot Where ODASPAccount.AccountNo = ODASPPlot.AccountNo and ODASPPlot.AccountNo = '" +txtLandLordNo.Text + "' Order by ODASPPlot.PlotNo";

                cmd = new OdbcCommand(sql, cnn);
                RDR = cmd.ExecuteReader();
                listView2.Items.Clear();
                listView2.Columns.Clear();
                listView2.Columns.Add("Plot No", listView2.Width / 3);
                listView2.Columns.Add("Location", listView2.Width / 3);
                listView2.Columns.Add("DOC", listView2.Width / 3);
                listView2.Columns.Add("Expiry", listView2.Width / 3);
                listView2.Columns.Add("Rent Due", listView2.Width / 3);


                if (RDR.HasRows)
                {


                    while (RDR.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(RDR["PlotNo"].ToString());
                        if (RDR["PhysicalLocation"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["PhysicalLocation"].ToString());
                        }
                        if (RDR["CommencementDate"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["CommencementDate"].ToString());
                        }
                        if (RDR["expirydate"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["expirydate"].ToString());
                        }
                        if (RDR["AnnualRent"].ToString() != "")
                        {
                            lv3.SubItems.Add(RDR["AnnualRent"].ToString());
                        }


                        listView2.Items.Add(lv3);





                    }

                }
                RDR.Close();
                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString ());
            }
        }
        private void AnableChilds(Control ctrl)
        {
            foreach (Control c in ctrl.Controls)
            {
                AnableChilds(c);
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
            }
        }
        private void frmLandLord_Load(object sender, EventArgs e)
        {
           getLANDLORDS();
            loadTowns();
            disableALLRECORD();
            loadDefault();
           LoadAccountType();
            getLANDLORDTYPE();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            GeneralVariables vars = new GeneralVariables();
            vars.NewRecord = true;
            AnableChilds(this);
        }
        private void loadRECORD() { 
            try{
        GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPAccount  WHERE AccountNo = '" +txtLandLordNo.Text + "'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    txtLandLordName.Text = reader["CompanyName"].ToString();
                    txtLandLordNo.Text = reader["AccountNo"].ToString();
                    txtemailAddress.Text = reader["EmailAddress"].ToString();
                    txtMobileNo.Text = reader["MobileNo"].ToString();
                    txtPhysicalAddress.Text = reader["PhysicalAddress"].ToString();
                    txtPostalAddress.Text = reader["PostalAddress"].ToString();
                    txtTelephoneNo.Text = reader["TelephoneNo"].ToString();
                    cboAccountType.Text = reader["AccountType"].ToString();
                    cboTownCode.Text = reader["Towncity"].ToString();
                    txtContactDesignation.Text = reader["ContactTitle"].ToString();
                    txtContactName.Text = reader["ContactPerson"].ToString();
               
                
                
                
                }
                reader.Close();
                cnn.Close();
        }catch (Exception ex){
        MessageBox .Show (ex.ToString ());
        }
        }
    
        private void generateLandLordNo()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("Select * from ODASPLAstNumbers Where AutoLandLordNo = 'Y'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {


                    // txtLandlord1 .Text  = ;

                    switch (Convert.ToInt32(reader["LandLordNo"].ToString().Length))
                    {
                        case 1: txtLandLordNo.Text = reader["LandLordPrefix"].ToString() + "0000" + Convert.ToInt32(reader["LandLordNo"].ToString().Trim()) + 1;
                            break;
                        case 2: txtLandLordNo.Text = reader["LandLordPrefix"].ToString() + "000" + Convert.ToInt32(reader["LandLordNo"].ToString().Trim()) + 1;
                            break;
                        case 3: txtLandLordNo.Text = reader["LandLordPrefix"].ToString() + "00" + Convert.ToInt32(reader["LandLordNo"].ToString().Trim()) + 1;
                            break;
                        case 4: txtLandLordNo.Text = reader["LandLordPrefix"].ToString() + "0" + Convert.ToInt32(reader["LandLordNo"].ToString().Trim()) + 1;
                            break;
                        case 5: txtLandLordNo.Text = reader["LandLordPrefix"].ToString() + Convert.ToInt32(reader["LandLordNo"].ToString().Trim()) + 1;
                            break;
                    }
                    int c = Convert.ToInt32(reader["LandLordNo"].ToString()) + 1;
                    cnn.Close();
                    cnn.Open();


                    cmd = new OdbcCommand("UPDATE ODASPLAstNumbers SET LandLordNo='" + c + "' ", cnn);
                    cmd.ExecuteNonQuery();
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
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPAccount  WHERE AccountNo = '" +txtLandLordNo.Text + "'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    cmd = new OdbcCommand(" UPDATE ODASPAccount SET CompanyName ='" + txtLandLordName.Text  +
                        "',CreatedBy='" + GeneralVariables.CurrentUserName +
                        "',DateCreated='" + DateTime.Today +
                        "',Status='A',EmailAddress='" + txtemailAddress.Text +
                        "',MobileNo='" + txtMobileNo.Text +
                        "',PostalAddress='" + txtPostalAddress.Text +
                        "' ,Towncity='" + txtTownDescription.Text +
                        "',ContactTitle='" + txtContactDesignation.Text +
                        "',ContactPerson='" + txtContactName.Text +
                        "',PhysicalAddress='" + txtPhysicalAddress.Text +
                        "',TelephoneNo='" + txtTelephoneNo.Text +
                        "',AccountType='" + cboAccountType.Text +
                        "' WHERE AccountNo = '" + txtLandLordNo.Text + "'", cnn);
                    cnn.Close();
                    cnn.Open();
                    cmd.ExecuteNonQuery();
                    cnn.Close();
                }
                else {
                    generateLandLordNo();
                    cmd = new OdbcCommand("INSERT INTO ODASPAccount(AccountNo,CompanyName,CreatedBy,DateCreated,Status)"+
                        " VALUES('" + txtLandLordNo.Text + "','" + txtLandLordName .Text +
                        "','" + GeneralVariables.CurrentUserName + "','" + DateTime.Today + "','A')", cnn);
                    cnn.Close();
                    cnn.Open();
                    cmd.ExecuteNonQuery();
                    cnn.Close();
                    cmd = new OdbcCommand(" UPDATE ODASPAccount SET CompanyName ='" + txtLandLordName +
                       "',CreatedBy='" + GeneralVariables.CurrentUserName +
                       "',DateCreated='" + DateTime.Today +
                       "',Status='A',EmailAddress='" + txtemailAddress.Text +
                       "',MobileNo='" + txtMobileNo.Text +
                       "',PostalAddress='" + txtPostalAddress.Text +
                       "',Towncity='" + txtTownDescription.Text +
                       "',ContactTitle='" + txtContactDesignation.Text +
                       "',ContactPerson='" + txtContactName.Text +
                       "',PhysicalAddress='" + txtPhysicalAddress.Text +
                       "',TelephoneNo='" + txtTelephoneNo.Text +
                       "',AccountType='" + cboAccountType.Text +
                       "' WHERE AccountNo = '" + txtLandLordNo.Text + "'", cnn);
                    cnn.Close();
                    cnn.Open();
                    cmd.ExecuteNonQuery();
                    cnn.Close();
                }
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
        }
        private Boolean ValidateData() {

         if (txtLandLordName.Text == "")
            {
                MessageBox.Show("The Account Name Cannot be Left Blank", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtLandLordName.Focus();
                return false;
            }
         else if (txtAccountTypeDescription.Text == "")
         {
             MessageBox.Show("The Account Type Entered is invalid", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
             txtAccountTypeDescription.Focus();
             return false;
         }
         else if (txtPhysicalAddress.Text == "")
         {
             MessageBox.Show("The Physical Address cannot be Left Blank", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
             txtPhysicalAddress.Focus();
             return false;
         }
         else if (txtPostalAddress.Text == "")
         {
             MessageBox.Show("The Postal Address of the LandLord is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
             txtPostalAddress.Focus();
             return false;
         }
         else if (txtTelephoneNo.Text == "")
         {
             MessageBox.Show("The Telephone Contact is needed for ease of Access", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
             txtTelephoneNo.Focus();
             return false;
         }
         else if (txtTownDescription.Text == "")
         {
             MessageBox.Show("The Town Code is neccessary", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
             txtTownDescription.Focus();
             return false;
         }
         else if (cboAccountType.Text == "")
         {
             MessageBox.Show("The Account Type is Mandatory", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
             cboAccountType.Focus();
             return false;
         }
            else {

                return true;
            }
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();
            if (ValidateData())
            {
            if (GeneralVariables.NewRecord ==true )
            {
                generateLandLordNo();
            }
            saveRecord();
            
            GeneralVariables.NewRecord = false; 
           // getLANDLORDS();
            disableALLRECORD();
                
            }

        }

        private void listView1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            ListView.CheckedListViewItemCollection check = listView1.CheckedItems;
              foreach (ListViewItem checkeditems in check)
            {
                txtLandLordNo.Text = checkeditems.Text;
           loadRECORD();
            showALLSITESByLandlord();
            }
        }

        private void txtLandLordName_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtLandLordName_Leave(object sender, EventArgs e)
        {
            txtContactName.Text = txtLandLordName.Text;
            txtLandLordName.Text = txtLandLordName.Text.ToUpper();
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
           search = true;
          getLANDLORDS();
        }

        private void txtLandLordNo_TextChanged(object sender, EventArgs e)
        {
            
          
        }

        private void listView1_MouseClick(object sender, MouseEventArgs e)
        {
            getLANDLORDS();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            GeneralVariables vars = new GeneralVariables();
            vars.rptODASPPlotSites.currentRecord = txtLandLordNo.Text;
            vars.rptODASPPlotSites.ShowDialog();
        }
    }
}
