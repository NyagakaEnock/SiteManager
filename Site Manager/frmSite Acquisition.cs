using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Odbc;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace Site_Manager
{
    public partial class frmSite_Acquisition : Form
    {
        OdbcCommand cmd;
        OdbcDataReader reader;
        DataTable dTable;
        DataSet ds;
        OdbcDataAdapter da;
        public String CurrentUserName;
        int j, k, i;
        public frmSite_Acquisition()
        {
            InitializeComponent();
        }

        private void textBox30_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox46_TextChanged(object sender, EventArgs e)
        {

        }
        public string getTownCode()
        {
           string townCode = " ";
            System .Windows .Forms .Form ThisForm=System .Windows .Forms .Application .OpenForms ["frmMain"];
            townCode = ((frmMain)ThisForm).textBox2.Text.ToString();
    

           return townCode;
        }
        public string getCurrentUser()
        {
            string user = " ";
            System.Windows.Forms.Form ThisForm = System.Windows.Forms.Application.OpenForms["frmMain"];
            user = ((frmMain)ThisForm).txtUser.Text.ToString();


            return user;
        }
        private void frmSite_Acquisition_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'dataSet1.ODASPPlot' table. You can move, or remove it, as needed.
            this.oDASPPlotTableAdapter.Fill(this.dataSet1.ODASPPlot);
            txtIncrementStartYear.Text = "0";
            txtTownCode.Text = getTownCode();
            ttxtTown.Text = txtTownCode.Text;
            txtUser.Text = getCurrentUser();
          
         DisAnable();
         txtSearch.Enabled = true;
         loadPaymentModes();
         loadMedia();
         loadLandLords();
         showALLPROPERTIES1();
         loadPlots();
         loadMasts();
         loadCouncil();
           
        }
        private void Anable()
        {
            numericUpDown1.Enabled = true;
            rdoYes.Checked = true;
            rdoFirm.Checked = true;
            chkLease.Checked = true;
            txtAcquiredBy.Enabled = true;
            txtAddress.Enabled = true;
            txtAmountDue.Enabled = true;
            txtAnnualRent.Enabled = true;
            txtComments.Enabled = true;
            txtContractNo .Enabled =true ;
            txtContractYear.Enabled = true;
            txtCouncil.Enabled = true;
            txtCurrenPeriod.Enabled = true;
            txtDetails.Enabled = true;
            txtEmail.Enabled = true;
            txtExpiryDate.Enabled = true;
            txtFaceNo.Enabled = true;
            txtFaceRate.Enabled = true;
            txtHighwayZone.Enabled = true;
            txtIncreamentInterval.Enabled = true;
            txtIncreamentRecent.Enabled = true;
            txtIncreamentStarts.Enabled = true;
            txtIncreamentType.Enabled = true;
            txtInstallments.Enabled = true;
            txtInstallPaymentDue.Enabled = true;
            txtInvoiceNo.Enabled = true;
            txtLandlord1.Enabled = true;
            txtLandlord2.Enabled = true;
            txtLocation.Enabled = true;
            txtLRNo.Enabled = true;
            txtMastDetails.Enabled = true;
            txtMastNo.Enabled = true;
            txtMastRate.Enabled = true;
            txtMeterNo.Enabled = true;
            txtNoMast.Enabled = true;
            txtOtherDeatils.Enabled = true;
            txtPaimentInterval.Enabled = true;
            txtPayDueDate.Enabled = true;
            txtPaymentMode.Enabled = true;
            txtPercentageAmount.Enabled = true;
            txtPlotNo.Enabled = true;
            txtPropertyCode .Enabled =true ;
            txtRate .Enabled =true ;
            txtRent .Enabled =true ;
            txtSearch .Enabled =true ;
            txtSerialNo .Enabled =true ;
            txtSiteNo .Enabled =true ;
            txtStatus .Enabled =true ;
            txtTransactionNo .Enabled =true ;
            cmbCombo.Enabled = true;
            cmbMedia.Enabled = true;
            cmbPaymenMode.Enabled = true;
            cmbSize.Enabled = true;


        }

        private void DisAnable()
        {
            btnHelp.Enabled = false;
            txtTownCity.Enabled = false;
            txtMobileNo.Enabled = false;
            acquisitionDate.Enabled = false;
            CommencementDate.Enabled = false;
            ttxtTown.Enabled = false;
            numericUpDown1.Enabled = false;
            txtAcquiredBy.Enabled = false ;
            txtAddress.Enabled = false;
            txtAmountDue.Enabled = false;
            txtAnnualRent.Enabled = false;
            txtComments.Enabled = false;
            txtContractNo.Enabled = false;
            txtContractYear.Enabled = false;
            txtCouncil.Enabled = false;
            txtCurrenPeriod.Enabled = false;
            txtDetails.Enabled = false;
            txtEmail.Enabled = false;
            txtExpiryDate.Enabled = false;
            txtFaceNo.Enabled = false;
            txtFaceRate.Enabled = false;
            txtHighwayZone.Enabled = false;
            txtIncreamentInterval.Enabled = false;
            txtIncreamentRecent.Enabled = false;
            txtIncreamentStarts.Enabled = false;
            txtIncreamentType.Enabled = false;
            txtInstallments.Enabled = false;
            txtInstallPaymentDue.Enabled = false;
            txtInvoiceNo.Enabled = false;
            txtLandlord1.Enabled = false;
            txtLandlord2.Enabled = false;
            txtLocation.Enabled = false;
            txtLRNo.Enabled = false;
            txtMastDetails.Enabled = false;
            txtMastNo.Enabled = false;
            txtMastRate.Enabled = false;
            txtMeterNo.Enabled = false;
            txtNoMast.Enabled = false;
            txtOtherDeatils.Enabled = false;
            txtPaimentInterval.Enabled = false;
            txtPayDueDate.Enabled = false;
            txtPaymentMode.Enabled = false;
            txtPercentageAmount.Enabled = false;
            txtPlotNo.Enabled = false;
            txtPropertyCode.Enabled = false;
            txtRate.Enabled = false;
            txtRent.Enabled = false;
            txtSearch.Enabled = false;
            txtSerialNo.Enabled = false;
            txtSiteNo.Enabled = false;
            txtStatus.Enabled = false;
            txtTransactionNo.Enabled = false;
            cmbCombo.Enabled = false;
            cmbMedia.Enabled = false;
            cmbPaymenMode.Enabled = false;
            cmbSize.Enabled = false;


        }
        private void cmbCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT CouncilCode FROM ODASPCouncil WHERE Council='" + cmbCombo.SelectedItem .ToString ()   + "'", cnn);
                reader = cmd.ExecuteReader();
          
                if (reader.Read())
                {


                    txtCouncil.Text = reader["CouncilCode"].ToString();
                    txtTownCode.Text = reader["CouncilCode"].ToString();
                    ttxtTown.Text = reader["CouncilCode"].ToString();
                }
                cnn.Close();

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.Message);

            }
        
        }

        private void cmbCombo_Click(object sender, EventArgs e)
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
               
                ds = new DataSet();
                cmd = new OdbcCommand("SELECT Council FROM ODASPCouncil", cnn);
                da = new OdbcDataAdapter(cmd);

                da.Fill(ds, "ODASPCouncil");
                dTable = ds.Tables[0];
                cmbCombo.Items.Clear();

                foreach (DataRow drow in dTable.Rows)
                {
                    cmbCombo.Items.Add(drow["Council"].ToString());

                }
                cnn.Close();
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }

        private void loadPaymentModes(){
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();

                ds = new DataSet();
                cmd = new OdbcCommand("SELECT PaymentModeDescription FROM ODASPPaymentMode", cnn);
                da = new OdbcDataAdapter(cmd);

                da.Fill(ds, "ODASPPaymentMode");
                dTable = ds.Tables[0];
                cmbPaymenMode.Items.Clear();

                foreach (DataRow drow in dTable.Rows)
                {
                    cmbPaymenMode.Items.Add(drow["PaymentModeDescription"].ToString());
                     

                }
                cnn.Close();
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }
        private void loadContracts() {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();



                string sql = "SELECT ContractNo,PlotNo,CommencementDate,ExpiryDate,LeaseDuration,AnnualRent FROM ODASMLeaseAgreement WHERE PLotNo LIKE '"+ txtPlotNo.Text  + "'  ORDER BY CommencementDate DESC";

                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();

                listView6.Items.Clear();
                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv = new ListViewItem(reader["ContractNo"].ToString());

                        lv.SubItems.Add(reader["PlotNo"].ToString());
                        lv.SubItems.Add(reader["CommencementDate"].ToString());
                        lv.SubItems.Add(reader["ExpiryDate"].ToString());
                        lv.SubItems.Add(reader["LeaseDuration"].ToString());
                        lv.SubItems.Add(reader["AnnualRent"].ToString());
                      
                        listView6.Items.Add(lv);




                    }
                }
                reader.Close();
                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void loadMedia()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();

                ds = new DataSet();
                cmd = new OdbcCommand("SELECT MediaDescription FROM ODASPMedia", cnn);
                da = new OdbcDataAdapter(cmd);

                da.Fill(ds, "ODASPMedia");
                dTable = ds.Tables[0];
                cmbMedia.Items.Clear();
               
                foreach (DataRow drow in dTable.Rows)
                {
                    cmbMedia.Items.Add(drow["MediaDescription"].ToString());
                     


                }
                cnn.Close();
               
               
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }

        private void loadMediaSize()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();

                ds = new DataSet();
                cmd = new OdbcCommand("SELECT MediaSize FROM ODASPMediaSize WHERE MediaCode='"+ txtMedia .Text + "'", cnn);
                da = new OdbcDataAdapter(cmd);

                da.Fill(ds, "ODASPMediaSize");
                dTable = ds.Tables[0];
                cmbSize.Items.Clear();

                foreach (DataRow drow in dTable.Rows)
                {
                    cmbSize.Items.Add(drow["MediaSize"].ToString());
                }
                cnn.Close();
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }
        
       private void loadLandLords(){
           try
           {
               GeneralVariables GeneralVariables = new GeneralVariables();
               OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               cnn.Open();
            


                string sql = "SELECT * FROM ODASPAccount Where Status = 'A' AND AccountType = 'LLORD' oRDER BY AccountNo";
              
               OdbcCommand cmd = new OdbcCommand(sql, cnn);
               reader = cmd.ExecuteReader();

               listView1.Items.Clear();
               if (reader.HasRows)
               {


                   while (reader.Read())
                   {

                       ListViewItem lv = new ListViewItem(reader["AccountNo"].ToString());
                       // lv.SubItems[0].Text = reader.GetString(0).ToString();
                       lv.SubItems.Add(reader["CompanyName"].ToString() );
                       lv.SubItems.Add(reader["Status"].ToString());
                       // lv.SubItems.Add(reader.GetString(1));
                       listView1.Items.Add(lv);




                   }
               }
               reader.Close();
               cnn.Close();

           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.Message);
           }
       }
       private void filterLandLords()
       {
           try
           {
               GeneralVariables GeneralVariables = new GeneralVariables();
               OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               cnn.Open();



               string sql = "SELECT * FROM ODASPAccount Where Status = 'A' AND AccountType = 'LLORD' AND AccountNo='" + txtLandlord1.Text + "' oRDER BY AccountNo ";
                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
               reader = cmd.ExecuteReader();

               listView1.Items.Clear();
               if (reader.HasRows)
               {


                   while (reader.Read())
                   {

                       ListViewItem lv = new ListViewItem(reader["AccountNo"].ToString());
                       // lv.SubItems[0].Text = reader.GetString(0).ToString();
                       lv.SubItems.Add(reader["CompanyName"].ToString());
                       lv.SubItems.Add(reader["Status"].ToString());
                       // lv.SubItems.Add(reader.GetString(1));
                       listView1.Items.Add(lv);




                   }
               }
               reader.Close();
               cnn.Close();

           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.Message);
           }
       }

       private void searchLandLords()
       {
           try
           {
               GeneralVariables GeneralVariables = new GeneralVariables();
               OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               cnn.Open();



               string sql = "SELECT * FROM ODASPAccount Where (CompanyName like '%" + txtSearch.Text + "%' OR PhysicalAddress like '%" + txtSearch.Text + "%' OR PostalAddress like '%" + txtSearch.Text + "%' and MobileNo like '%" + txtSearch.Text + "%' ) AND Status = 'A' AND AccountType = 'LLORD' Order by AccountNo";

               OdbcCommand cmd = new OdbcCommand(sql, cnn);
               reader = cmd.ExecuteReader();

               listView1.Items.Clear();
               if (reader.HasRows)
               {


                   while (reader.Read())
                   {

                       ListViewItem lv = new ListViewItem(reader["AccountNo"].ToString());
                       // lv.SubItems[0].Text = reader.GetString(0).ToString();
                       lv.SubItems.Add(reader["CompanyName"].ToString());
                       lv.SubItems.Add(reader["Status"].ToString());
                       // lv.SubItems.Add(reader.GetString(1));
                       listView1.Items.Add(lv);




                   }
               }
               reader.Close();
               cnn.Close();

           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.Message);
           }
       }
       private void showALLPROPERTIES1() {
           try
           {
               GeneralVariables GeneralVariables = new GeneralVariables();
               OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               cnn.Open();
               string sql = "SELECT * FROM ODASPProperties";
              
               OdbcCommand cmd = new OdbcCommand(sql, cnn);
               reader = cmd.ExecuteReader();

               listView2.Items.Clear();
               if (reader.HasRows)
               {


                   while (reader.Read())
                   {

                       ListViewItem lv2 = new ListViewItem(reader["PropertyCode"].ToString());
                       // lv.SubItems[0].Text = reader.GetString(0).ToString();
                       lv2.SubItems.Add(reader["PropertyDescription"].ToString());

                       listView2.Items.Add(lv2);




                   }
               }
               reader.Close();
               cnn.Close();

           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.Message);
           }
       }

       private void loadPlots()
       {
           try
           {
               GeneralVariables GeneralVariables = new GeneralVariables();
               OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               cnn.Open();
               string sql = "SELECT * FROM ODASPPlot P, ODASPAccount A where A.AccountNo = P.AccountNo and P.TownCode = '" +txtTownCode.Text + "' ";

               OdbcCommand cmd = new OdbcCommand(sql, cnn);
               reader = cmd.ExecuteReader();

               listView3.Items.Clear();
               if (reader.HasRows)
               {


                   while (reader.Read())
                   {

                       ListViewItem lv3 = new ListViewItem(reader["PlotNo"].ToString());
                       // lv.SubItems[0].Text = reader.GetString(0).ToString();
                       lv3.SubItems.Add(reader["PlotName"].ToString());
                       lv3.SubItems.Add(reader["LRNo"].ToString());
                       lv3.SubItems.Add(reader["PhysicalLocation"].ToString());
                       lv3.SubItems.Add(reader["TownCode"].ToString());

                       listView3.Items.Add(lv3);




                   }
               }
               reader.Close();
               cnn.Close();

           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.Message);
           }
       }

       private void loadMasts()
       {
           try
           {
               GeneralVariables GeneralVariables = new GeneralVariables();
               OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               cnn.Open();
               string sql = "SELECT * FROM ODASPPlotMast  WHERE ODASPPlotMast.PlotNo = '" + txtPlotNo.Text + "' ";

               OdbcCommand cmd = new OdbcCommand(sql, cnn);
               reader = cmd.ExecuteReader();

               listView4.Items.Clear();
               if (reader.HasRows)
               {


                   while (reader.Read())
                   {

                       ListViewItem lv3 = new ListViewItem(reader["MastNo"].ToString());
                       // lv.SubItems[0].Text = reader.GetString(0).ToString();
                       lv3.SubItems.Add(reader["PlotNo"].ToString());
                      
                       lv3.SubItems.Add(reader["MastDetails"].ToString());
                       lv3.SubItems.Add(reader["AnnualRent"].ToString());

                       listView4.Items.Add(lv3);




                   }
               }
               reader.Close();
               cnn.Close();

           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.Message);
           }
       }
       private void loadFaces()
       {
           try
           {
               GeneralVariables GeneralVariables = new GeneralVariables();
               OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               cnn.Open();
               string sql =" SELECT * FROM ODASPPlotSite, ODASpPlot Where ODASPPlotSite.PlotNo = ODASPPlot.PlotNo AND (ODASPPlotSite.MastNo = '" + txtMastNo.Text + "' or ODASPPlotSite.PlotNo = '" + txtPlotNo.Text + "')";

               OdbcCommand cmd = new OdbcCommand(sql, cnn);
               reader = cmd.ExecuteReader();

               listView5.Items.Clear();
               if (reader.HasRows)
               {


                   while (reader.Read())
                   {

                       ListViewItem lv3 = new ListViewItem(reader["SiteNo"].ToString());
                       if (reader["MastNo"].ToString() != "") {
                           lv3.SubItems.Add(reader["MastNo"].ToString());
                       }
                       if (reader["PlotNo"].ToString() != "")
                       {
                           lv3.SubItems.Add(reader["PlotNo"].ToString());
                       }
                       if (reader["PlotName"].ToString() != "")
                       {
                           lv3.SubItems.Add(reader["PlotName"].ToString());
                       }
                      
                       

                       listView5.Items.Add(lv3);




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
       private void saveProperty()
       {
           try
           {
               GeneralVariables GeneralVariables = new GeneralVariables();
               OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
              
               cnn.Open();
               string sql = "SELECT * FROM ODASMSiteProperties  WHERE SiteNo = '" +txtSiteNo.Text.Trim () + "' and PropertyCode = '" +txtPropertyCode.Text + "'";
               ListView.ListViewItemCollection items = listView2.Items;
               OdbcCommand cmd = new OdbcCommand(sql, cnn);
               reader = cmd.ExecuteReader();
               j = listView2.Items.Count;
               for (i = 1; i <= j; i++)
               {
                   foreach (ListViewItem item in items  )
                       if (item.Checked == true)
                   {
                       txtPropertyCode.Text = item.Text;
                       if (reader.Read())
                       {
                           cmd = new OdbcCommand("UPDATE ODASMSiteProperties SET OtherDetails='" + txtOtherDeatils.Text +
                               "',AmountDue='" + Convert.ToString(txtAmountDue.Text) + "',DateAssigned='" + dateAssigned.Text + "',CommencementDate'" + startDate.Text + "',PlotNo='" + txtPlotNo.Text + "' WHERE SiteNo = '" + txtSiteNo.Text.Trim() + "' and PropertyCode = '" + txtPropertyCode.Text + "'", cnn2);
                           if (cnn2.State == ConnectionState.Closed)
                           {
                               cnn2.Open();
                           }
                           cnn2.Open();
                           cmd.ExecuteNonQuery();
                           cnn2.Close();
                       }
                       else {
                           cmd = new OdbcCommand("INSERT INTO ODASMSiteProperties(SiteNo,PropertyCode,Status,PreparedBY,DatePrepared,OtherDetails,AmountDue,DateAssigned,CommencementDate,PlotNo)"+
                                                                           " VALUES('" + txtSiteNo.Text + "','" + txtPropertyCode.Text + "','ACTIVE','" + CurrentUserName + "','" + DateTime.Today + "','" + txtOtherDeatils.Text + "','" + txtAmountDue.Text + "','" + dateAssigned.Text + "','" + startDate.Text + "','" + txtPlotNo.Text + "')", cnn2);
                           if (cnn2.State == ConnectionState.Open)
                           {
                               cnn2.Close();
                           }
                           cnn2.Open();
                           if (cnn2.State == ConnectionState.Closed)
                           {
                               cnn2.Open();
                           }
                          // cnn2.Open();
                           cmd.ExecuteNonQuery();
                           cnn2.Close();
                       }

                   }
               }
               
               
           }
           catch (Exception ex) {
               MessageBox.Show(ex.ToString ());
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
               if (c is NumericUpDown )
               {
                   ((NumericUpDown)(c)).Enabled = false;
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
       private void loadInstallments()
       {
           try
           {
               GeneralVariables GeneralVariables = new GeneralVariables();
               OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               cnn.Open();
               string sql = "SELECT * FROM ODASMInstallment I Where I.PlotNo='" +txtPlotNo.Text + "' AND I.ContractNo = '" + txtContractNo.Text  +"' Order by InstallmentNo";

               OdbcCommand cmd = new OdbcCommand(sql, cnn);
               reader = cmd.ExecuteReader();

               listView7.Items.Clear();
               if (reader.HasRows)
               {


                   while (reader.Read())
                   {

                       ListViewItem lv3 = new ListViewItem(reader["InstallmentNo"].ToString());
                       if (reader["TotalRent"].ToString() != "")
                       {
                           lv3.SubItems.Add(reader["TotalRent"].ToString());
                       }
                       if (reader["PaymentDueDate"].ToString() != "")
                       {
                           lv3.SubItems.Add(reader["PaymentDueDate"].ToString());
                       }
                       if (reader["PaymentFlag"].ToString() != "")
                       {
                           lv3.SubItems.Add(reader["PaymentFlag"].ToString());
                       }
                       if (reader["ContractYear"].ToString() != "")
                       {
                           lv3.SubItems.Add(reader["ContractYear"].ToString());
                       }
                       if (reader["CurrentPeriod"].ToString() != "")
                       {
                           lv3.SubItems.Add(reader["CurrentPeriod"].ToString());
                       }
                       if (reader["PaymentMode"].ToString() != "")
                       {
                           lv3.SubItems.Add(reader["PaymentMode"].ToString());
                       }
                       if (reader["InvoiceNo"].ToString() != "")
                       {
                           lv3.SubItems.Add(reader["InvoiceNo"].ToString());
                       }
                       if (reader["PaymentDue"].ToString() != "")
                       {
                           lv3.SubItems.Add(reader["PaymentDue"].ToString());
                       }
                       

                       listView7.Items.Add(lv3);




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
       private void filterPlots()
       {
           try
           {
               GeneralVariables GeneralVariables = new GeneralVariables();
               OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               cnn.Open();
               string sql = "SELECT * FROM ODASPPlot P, ODASPAccount A where A.AccountNo = P.AccountNo and P.TownCode = '" + txtTownCode.Text + "' AND PlotNo LIKE '%" + txtPlotNo.Text + "%'";

               OdbcCommand cmd = new OdbcCommand(sql, cnn);
               reader = cmd.ExecuteReader();

               listView3.Items.Clear();
               if (reader.HasRows)
               {


                   while (reader.Read())
                   {

                       ListViewItem lv3 = new ListViewItem(reader["PlotNo"].ToString());
                       // lv.SubItems[0].Text = reader.GetString(0).ToString();
                       lv3.SubItems.Add(reader["PlotName"].ToString());
                       lv3.SubItems.Add(reader["LRNo"].ToString());
                       lv3.SubItems.Add(reader["PhysicalLocation"].ToString());
                       lv3.SubItems.Add(reader["TownCode"].ToString());

                       listView3.Items.Add(lv3);




                   }
               }
               reader.Close();
               cnn.Close();

           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.Message);
           }
       }
       private void searchPlots() {
           try
           {
               GeneralVariables GeneralVariables = new GeneralVariables();
               OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               cnn.Open();
               string sql = "SELECT * FROM ODASPPlot P, ODASPAccount A where A.AccountNo = P.AccountNo  AND (PlotNo LIKE '%" + txtSearch.Text + "%' OR PlotName LIKE '%" + txtSearch .Text + "%')"; 

               OdbcCommand cmd = new OdbcCommand(sql, cnn);
               reader = cmd.ExecuteReader();

               listView3.Items.Clear();
               if (reader.HasRows)
               {


                   while (reader.Read())
                   {

                       ListViewItem lv3 = new ListViewItem(reader["PlotNo"].ToString());
                       // lv.SubItems[0].Text = reader.GetString(0).ToString();
                       lv3.SubItems.Add(reader["PlotName"].ToString());
                       lv3.SubItems.Add(reader["LRNo"].ToString());
                       lv3.SubItems.Add(reader["PhysicalLocation"].ToString());
                       lv3.SubItems.Add(reader["TownCode"].ToString());

                       listView3.Items.Add(lv3);




                   }
               }
               reader.Close();
               cnn.Close();

           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.Message);
           }
       }
       private void filterCouncil()
       {
           try
           {
               GeneralVariables GeneralVariables = new GeneralVariables();
               OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               cnn.Open();
               string sql = "SELECT * FROM ODASPCouncil WHERE CouncilCode = '" + txtCouncil.Text + "'";

               OdbcCommand cmd = new OdbcCommand(sql, cnn);
               reader = cmd.ExecuteReader();

               listView3.Items.Clear();
               if (reader.HasRows)
               {


                   while (reader.Read())
                   {

                       ListViewItem lv3 = new ListViewItem(reader["PlotNo"].ToString());
                       // lv.SubItems[0].Text = reader.GetString(0).ToString();
                       lv3.SubItems.Add(reader["PlotName"].ToString());
                       lv3.SubItems.Add(reader["LRNo"].ToString());
                       lv3.SubItems.Add(reader["PhysicalLocation"].ToString());
                       lv3.SubItems.Add(reader["TownCode"].ToString());

                       listView3.Items.Add(lv3);




                   }
               }
               reader.Close();
               cnn.Close();

           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.Message);
           }
       }
       private void loadRecord()
       { 
       String sql ="select * from ODASPPlot Where PlotNo = '" + txtPlotNo.Text + "'";
       try
       {

           GeneralVariables GeneralVariables = new GeneralVariables();
           OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
           cnn.Open();
           cmd = new OdbcCommand(sql, cnn);
           reader = cmd.ExecuteReader();
          
           if (reader.Read())
           {
               if (reader["OnRoadReserve"].ToString() == "N")
               {
                   rdoNo.Checked = true;
               }
               else { rdoYes.Checked = true; }

              
               txtPlotNo.Text = reader["PlotNo"].ToString();
               txtStatus.Text = reader["Status"].ToString();
               txtPaymentMode.Text = reader["PaymentMode"].ToString();
               txtIncreamentInterval.Text = reader["RentVariationType"].ToString();
               txtIncreamentType.Text = reader["AnnualRentIncrementType"].ToString();
               if (reader["IncrementStartYear"].ToString() == "")
               {
                   txtIncrementStartYear.Text = "2";
               }
               else { txtIncrementStartYear.Text = reader["IncrementStartYear"].ToString(); }

               if(reader ["IncrementFrequency"].ToString ()==""){
                   txtIncreamentStarts.Text = "1";
               }
               else { txtIncreamentStarts.Text = reader["IncrementFrequency"].ToString(); }

               if (reader["PaymentInterval"].ToString () == "")
               {
                   txtPaimentInterval.Text = "1";
               }
               else { txtPaimentInterval.Text = reader["PaymentInterval"].ToString(); }

               if (reader["NoOfMasts"].ToString() == "")
               {
                txtNoMast .Text ="0";
                }
               else 
               {
                txtNoMast .Text =reader ["NoOfMasts"].ToString ();
               }
                if(reader ["NoofSites"].ToString ()=="")
                {
                    txtFaceNo.Text = "0";

                }
                else { txtFaceNo.Text = reader["NoofSites"].ToString(); }
                if (reader["AnnualRentIncrement"].ToString () == "")
                {
                    txtPercentageAmount.Text = "0";
                }
                else { txtPercentageAmount.Text = reader["AnnualRentIncrement"].ToString();
                }
                if (reader["AnnualRent"].ToString() == "")
                {
                    txtRent.Text = "0";
                }
                else { txtRent.Text = reader["AnnualRent"].ToString(); }
                if (reader["AnnualRate"].ToString() == "")
                {
                    txtRate.Text = "0";
                }
                else { txtRate.Text = reader["AnnualRate"].ToString(); }

                txtAcquiredBy.Text = reader["AcquiredBy"].ToString();
                txtTownCode.Text = reader["TownCode"].ToString();
                acquisitionDate.Text = reader["AcquisitionDate"].ToString();
               // int test;
               
              /*  if (int.TryParse(reader["LeaseDuration"].ToString(), out test))
                {
                    numericUpDown1.Value = 1;
                   
                }
                else {*/
                    //numericUpDown1.Value= Convert.ToInt32(reader["LeaseDuration"].ToString());

                    if (reader["LeaseDuration"].ToString() == "0")
                    {
                        numericUpDown1.Value = 1;

                    }
                    else {
                        numericUpDown1.Value = Convert.ToInt32(reader["LeaseDuration"].ToString());
                    }
               // }
                if (reader["WithLease"].ToString() == "Y")
                {
                    chkLease.Checked = true;

                }
                else { chkLease.Checked = false; }
               txtLRNo .Text = reader ["LRNo"].ToString ();

               txtCouncil.Text = reader["CouncilCode"].ToString();
               txtLocation.Text = reader["PhysicalLocation"].ToString();
               CommencementDate.Text = reader["CommencementDate"].ToString();
               txtExpiryDate.Text = reader["expirydate"].ToString();
               txtComments.Text = reader["Comments"].ToString();
               txtLandlord1.Text = reader["AccountNo"].ToString();

           }
           reader .Close ();
           cnn.Close();
          
       }
       catch (Exception EX)
       {
           MessageBox.Show(EX.ToString());

       }
       }
       private void loadLandLordRecord(){

           try
           {

               GeneralVariables GeneralVariables = new GeneralVariables();
               OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               cnn.Open();
               cmd = new OdbcCommand("SELECT * FROM ODASPAccount  WHERE AccountNo = '" + txtLandlord1.Text + "'", cnn); 
               reader = cmd.ExecuteReader();
               //dTable = new DataTable();
               // dTable.Load(reader);
               //cmbCombo.Items.Clear();
               if (reader.Read())
               {


                   txtLandlord2.Text = reader["CompanyName"].ToString();
                   txtEmail.Text = reader["EmailAddress"].ToString();
                   txtMobileNo.Text = reader["MobileNo"].ToString();
                   txtAddress.Text = reader["PostalAddress"].ToString();
                   txtTownCity.Text = reader["Towncity"].ToString();
               }
               cnn.Close();

           }
           catch (Exception EX)
           {
               MessageBox.Show(EX.ToString());

           }
    }

       private void loadMastRecord() {
           try
           {

               GeneralVariables GeneralVariables = new GeneralVariables();
               OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               cnn.Open();
               cmd = new OdbcCommand("select * from ODASPPlotMast Where PlotNo = '" + txtPlotNo.Text  + "'", cnn);
               reader = cmd.ExecuteReader();
               if (reader.Read())
               {


                   txtMastDetails.Text = reader["MastDetails"].ToString();
                   txtMastNo.Text = reader["MastNo"].ToString();
                   txtAnnualRent.Text = reader["AnnualRent"].ToString();
                  if (reader["AnnualRent"].ToString() == "" )
                  {
                       txtAnnualRent.Text = "0";
                      
                   }
                   else { txtAnnualRent.Text = reader["AnnualRent"].ToString(); }


                   if (reader["AnnualRate"].ToString() == "")
                   {
                       txtMastRate.Text = "0";
                   }
                    else { 
                       txtMastRate.Text = reader["AnnualRate"].ToString(); 
                   }

                    txtMeterNo.Text = reader["MeterNo"].ToString();
                    cmbMedia.Text = reader["TypeOfMast"].ToString();
                 
                   
                    
                  //  if (reader["OwenedByClient"].ToString() == "Y")
                  //  {
                   //     rdoClient.Checked = true;
                   /// }
                   // else { rdoFirm.Checked = true; }
               }
               cnn.Close();

           }
           catch (Exception EX)
           {
               MessageBox.Show(EX.ToString());

           }
       
       }
       private void loadFaceRecord()
       {
               try
               {

                   GeneralVariables GeneralVariables = new GeneralVariables();
                   OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                   cnn.Open();
                   cmd = new OdbcCommand("select * from ODASPPlotSite Where MastNo = '" + txtMastNo.Text + "'", cnn);
                   reader = cmd.ExecuteReader();
                   if (reader.Read())
                   {


                       txtSiteNo.Text = reader["SiteNo"].ToString();
                       txtDetails.Text = reader["SiteDetails"].ToString();
                       txtFaceRate.Text = reader["Rates"].ToString();
                       txtHighwayZone.Text = reader["HighwayZone"].ToString();
                       cmbMedia.Text = reader["MediaCode"].ToString();
                       txtMedia.Text = reader["MediaCode"].ToString();
                       txtMediaSize.Text = reader["MediaSize"].ToString();
                       cmbSize.Text = reader["MediaSize"].ToString();


                   }
                   cnn.Close();

               }
               catch (Exception EX)
               {
                   MessageBox.Show(EX.ToString());

               }
       }
       private void saveMast()
       {
           try
           {

               GeneralVariables GeneralVariables = new GeneralVariables();
               OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
              
               cnn.Open();
               cmd = new OdbcCommand("select * from ODASPPlotMast Where MastNo = '" + txtMastNo.Text + "'", cnn);
               reader = cmd.ExecuteReader();
               if (reader.Read())
               {
                   cmd = new OdbcCommand("UPDATE ODASPPlotMast SET Status='" + txtStatus.Text + "',MeterNo='" + txtMeterNo.Text +
                                            "',PlotNo='" + txtPlotNo.Text.Trim() +
                                            "',MastDetails='" + txtMastDetails.Text +
                                            "',AnnualRent='" + txtAnnualRent.Text +
                                            "',AnnualRate='" + txtMastRate.Text +
                                            "',LeaseDuration='" + numericUpDown1.Text +
                                           "',CommencementDate='" + CommencementDate.Text +
                                            "',expirydate='" + txtExpiryDate.Text +
                                            "',TypeOfMast='" + txtMedia.Text +
                                            "',MediaSize='" + txtMediaSize.Text +
                                            "',OwenedByClient='" + txtClient.Text + "' WHERE MastNo='" +txtMastNo .Text +"' ", cnn);
                   cnn.Close();
                   cnn.Open();
                   cmd.ExecuteNonQuery();
                   cnn.Close();
               }
               else {
                   
                   generateMastNo();
               
                   cnn.Close();
                   cmd = new OdbcCommand("INSERT INTO ODASPPlotMast(MastNo,CreatedBy,DateCreated,Created,"+
                   "Approved,Authorized,LeasePrepared,RentPaid,RatePaid,AllocationDate,NoofSites,PlotNo,"+
                   "MastDetails,AnnualRent,AnnualRate,LeaseDuration,CommencementDate,expirydate,TypeOfMast,MediaSize,OwenedByClient,Status,MeterNo) VALUES('" + txtMastNo.Text +
                       "','"+CurrentUserName  +
                       "','"+DateTime .Today +
                       "','Y','N','N','N','" + 0 + "','" + 0 +
                       "','" + DateTime.Today + "','" + 0 +
                       "','" + txtPlotNo.Text + 
                       "','" + txtMastDetails.Text  +
                       "','" + txtAnnualRent.Text + 
                       "','" + txtMastRate.Text +
                       "','" + numericUpDown1.Text +
                       "','" + CommencementDate.Text +
                       "','" + txtExpiryDate.Text +
                       "','" + txtMedia.Text +
                       "','" + txtMediaSize.Text + 
                       "','" + txtClient .Text + 
                       "','"+txtStatus .Text +
                       "','"+txtMeterNo .Text +"')", cnn);
                   cnn.Open();
                   cmd.ExecuteNonQuery();
               }
               cnn.Close();

           }
           catch (Exception EX)
           {
               MessageBox.Show(EX.ToString());

           }
       }
       private void savaSite() {
           try
           {

               GeneralVariables GeneralVariables = new GeneralVariables();
               OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               cnn.Open();
               cmd = new OdbcCommand("select * from ODASPPlotSite Where SiteNo = '" +txtSiteNo.Text +"'", cnn);
               reader = cmd.ExecuteReader();

               if (reader.Read())
               {
                   cmd = new OdbcCommand("UPDATE ODASPPlotSite SET Status = '" + txtStatus.Text +
                   "',MeterNo='" + txtMeterNo.Text + "',HighwayZone='" + txtHighwayZone.Text +
                   "',MastNo='" + txtMastNo.Text + "',PlotNo='" + txtPlotNo.Text.Trim() +
                   "',SiteDetails='" + txtDetails.Text + "',MediaSize='" + txtMediaSize.Text + "',MediaCode='" + txtMedia.Text + "',Active='" + txtActive.Text + "' WHERE SiteNo='"+txtSiteNo .Text +"'", cnn);
                   cnn.Close();
                   cnn.Open();
                   cmd.ExecuteNonQuery();
                   cnn.Close();
               }
               else {
                   cmd = new OdbcCommand("INSERT INTO ODASPPlotSite(SiteNo,CreatedBy,DateCreated,Created,PropertiesAssigned,Approved,Authorized,"+
                       "RatePaid,AllocationDate,JobBriefNo,Active,Status,MeterNo,HighwayZone,MastNo,PlotNo,SiteDetails,MediaSize,MediaCode)VALUES('"+txtSiteNo .Text .Trim ()+
                       "','" + CurrentUserName + "','" + DateTime.Today + "','Y','N','N','N','" + 0 +
                       "','"+DateTime .Today +"','','Y','"+txtStatus .Text +
                       "','" + txtMeterNo.Text + "','" + txtHighwayZone.Text + "','" + txtMastNo.Text + "','" + txtPlotNo.Text + "','" + txtDetails.Text + "','" + txtMediaSize.Text + "','" + txtMedia.Text + "')", cnn);
           cnn.Close (); 
                   cnn .Open ();
                   cmd.ExecuteNonQuery();
               }
               cnn.Close();

           }
           catch (Exception EX)
           {
               MessageBox.Show(EX.ToString());

           }
       }
        private void btnNew_Click(object sender, EventArgs e)
        {
            GeneralVariables vars = new GeneralVariables();
            vars.entryinProgress = true;
            anableAll(this );
            clear();
            txtPlotNo.Enabled = false;
            txtLandlord1.Enabled = false;
            txtPercentageAmount.Enabled = true;
            txtIncreamentType.Enabled = true;
            txtTownCode.Enabled = false;
            txtExpiryDate.Enabled = false;
           
        }
        private void clear() {
            clearAll(this );
            txtTownCode.Enabled = false;
            txtExpiryDate.Enabled = false;
            txtContractNo.Text = "";
            txtIncreamentInterval.Text = "N";
            txtIncreamentType.Text = "N";
            txtPaimentInterval.Text = "1";
            txtIncreamentStarts.Text = "1";
            txtPlotNo.Text = "";
            txtLandlord1.Text = "";
            txtIncrementStartYear.Text = "0";
            txtAnnualRent.Text = "0";
            rdoFirm.Checked = true;
            rdoYes.Checked = true;
            dateAssigned.Text = DateTime.Today.ToString("MM/dd/yyyy");
            startDate.Text = DateTime.Today.ToString("MM/dd/yyyy");
            numericUpDown1.Text = "0";
            txtMastNo.Text = "0";
            txtPercentageAmount.Text = "0";
            txtFaceNo.Text = "0";
            CommencementDate.Text = DateTime.Today.ToString("MM/dd/yyyy");
            acquisitionDate .Text = DateTime.Today.ToString("MM/dd/yyyy");
            txtExpiryDate.Text = DateTime.Today.ToString("MM/dd/yyyy");

        }
        private void generateLandLordNo() {
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
                        case 1: txtLandlord1 .Text =reader["LandLordPrefix"].ToString() +"0000"+Convert .ToInt32 ( reader["LandLordNo"].ToString().Trim ())+1;
                            break;
                        case 2: txtLandlord1.Text = reader["LandLordPrefix"].ToString() + "000" + Convert.ToInt32(reader["LandLordNo"].ToString().Trim()) + 1;
                            break;
                        case 3: txtLandlord1.Text = reader["LandLordPrefix"].ToString() + "00" + Convert.ToInt32(reader["LandLordNo"].ToString().Trim()) + 1;
                            break;
                        case 4: txtLandlord1.Text = reader["LandLordPrefix"].ToString() + "0" + Convert.ToInt32(reader["LandLordNo"].ToString().Trim()) + 1;
                            break;
                        case 5: txtLandlord1.Text = reader["LandLordPrefix"].ToString() + Convert.ToInt32(reader["LandLordNo"].ToString().Trim()) + 1;
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
        private void generatePlotNo()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPLastNumbers WHERE AutoPlotNo = 'Y'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {



                   
                    int c = Convert.ToInt32(reader["PlotNo"].ToString()) + 1;

                    switch (Convert.ToInt32(reader["LandLordNo"].ToString().Length))
                    {
                        case 1: txtPlotNo.Text = reader["PlotNoPrefix"].ToString() + "00000" + reader["PlotNo"].ToString();
                            break;
                        case 2: txtPlotNo.Text = reader["PlotNoPrefix"].ToString() + "0000" + reader["PlotNo"].ToString();
                            break;
                        case 3: txtPlotNo.Text = reader["PlotNoPrefix"].ToString() + "000" + reader["PlotNo"].ToString();
                            break;
                        case 4: txtPlotNo.Text = reader["PlotNoPrefix"].ToString() + "00" + reader["PlotNo"].ToString();
                            break;
                        case 5: txtPlotNo.Text = reader["PlotNoPrefix"].ToString() + "0" + reader["PlotNo"].ToString();
                            break;
                        case 6: txtPlotNo.Text = reader["PlotNoPrefix"].ToString() + reader["PlotNo"].ToString();
                            break;
                    }
                  
                    cnn.Close();
                    cnn.Open();
                    cmd = new OdbcCommand("UPDATE ODASPLastNumbers SET PlotNo='" + c + "' ", cnn);
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
                cmd = new OdbcCommand("SELECT * FROM ODASPAccount WHERE AccountNo = '" + txtLandlord1.Text + "'", cnn);
                reader = cmd.ExecuteReader();
               


                    if (reader.Read())
                    {
                        cnn.Close();
                        cnn.Open();

                        cmd = new OdbcCommand("UPDATE ODASPAccount SET EmailAddress='" + 
                            txtEmail.Text + "',MobileNo='" + 
                            txtMobileNo.Text + "',PostalAddress='" + 
                            txtAddress.Text + "',AccountType='LLORD',Towncity='" + 
                            txtTownCity.Text + "',CompanyName='" + txtLandlord2.Text + 
                            "',PhysicalAddress='" + txtLocation.Text + 
                            "'WHERE AccountNo='" + txtLandlord1.Text + "'", cnn);
                        cmd.ExecuteNonQuery();  
                    }
                    else
                    {
                        if (txtLandlord1.Text == "")
                        {
                            generateLandLordNo();
                        }
                       
                        cmd = new OdbcCommand("INSERT INTO ODASPAccount(AccountNo,CompanyName,CreatedBy,"+
                            "DateCreated,Status,EmailAddress,MobileNo,PostalAddress,AccountType,Towncity,"+
                            "PhysicalAddress)VALUES('" + txtLandlord1.Text + 
                            "','" + txtLandlord2.Text +
                            "','" + CurrentUserName  + "','" + DateTime.Today.ToString("MM/dd/yyyy") +
                            "','A','" + txtEmail.Text + "','" + txtMobileNo.Text + "','" + txtAddress .Text +
                            "','LLORD','"+txtTownCity .Text +"','"+txtLocation .Text +"')", cnn);

                        cmd.ExecuteNonQuery();
                        
                       
                    }
                    cnn.Close();

                
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        
        }
        private void updateRecord() {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               // OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                
                cmd = new OdbcCommand("select * from ODASPPlot Where PlotNo = '" + txtPlotNo.Text  + "'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    cmd = new OdbcCommand("UPDATE ODASPPlot SET Status='" + txtStatus.Text +
                        "',AnnualRentIncrement='" + txtPercentageAmount.Text +
                        "',PaymentMode='" + txtPaymentMode.Text +
                        "',RentVariationType='"+txtIncreamentInterval.Text +
                        "',AnnualRentIncrementType='" + txtIncreamentType.Text +
                        "',IncrementStartYear='" + txtIncrementStartYear.Text +
                        "',IncrementFrequency='" + txtIncreamentStarts.Text +
                        "',PaymentInterval='" + txtPaimentInterval.Text +
                        "',WithLease='" +txtLease .Text  +
                        "',AcquisitionDate='" + acquisitionDate.Text +
                        "',MeterNo='" + txtMeterNo.Text +
                        "',PlotName='" + txtLandlord2.Text +
                        "',CouncilCode='" + txtCouncil.Text +
                        "',AccountNo='" + txtLandlord1.Text +
                        "',OnRoadReserve='" + txtYes.Text +
                        "',TownCode='" + txtTownCode.Text +
                        "',LeaseDuration='" + numericUpDown1.Text +
                        "',LRNo='" + txtLRNo.Text +
                        "',PhysicalLocation='" + txtLocation.Text +
                        "',CommencementDate='" + CommencementDate.Text +
                        "',expirydate='" + txtExpiryDate.Text +
                        "',AcquiredBy='" + txtAcquiredBy.Text +
                        "',Comments='" + txtComments.Text +
                        "' Where PlotNo = '" + txtPlotNo.Text + "'",cnn );
                    cnn.Close();
                    cnn.Open();
                    cmd.ExecuteNonQuery();
                    
                    cnn.Close();
                    cnn.Open();
                    //Count the Number of Masts
                    cmd = new OdbcCommand("select  Count(PlotNo) as NoofMasts from ODASPPlotMast Where PlotNo = '" + txtPlotNo.Text + "'", cnn);
                    txtNoMast.Text = cmd.ExecuteScalar().ToString();

                    cnn.Close();
                    cnn.Open();
                    if (txtNoMast.Text == "")
                    {
                        cmd = new OdbcCommand("UPDATE ODASPPlot SET NoOfMasts='" + 0 + "' Where PlotNo = '" + txtPlotNo.Text + "'",cnn);
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                    }
                    else if (txtNoMast.Text != "")
                    {
                       
                        cmd = new OdbcCommand("UPDATE ODASPPlot SET NoOfMasts='" + txtNoMast.Text + "' Where PlotNo = '" + txtPlotNo.Text + "'",cnn);
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                    }
                    // Procedure to Count the Number of Sites
                    cnn.Open();
                    cmd = new OdbcCommand("select  Count(SiteNo) as NoOfSites from ODASPPlotSite Where MastNo = '" + txtMastNo.Text + "'", cnn);
                    txtFaceNo.Text = cmd.ExecuteScalar().ToString();

                    cnn.Close();
                    cnn.Open();
                    if (txtFaceNo.Text == "")
                    {
                        cmd = new OdbcCommand("UPDATE ODASPPlot SET NoofSites='" + 0 + "' Where PlotNo = '" + txtPlotNo.Text + "'",cnn);
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                    }
                    else if (txtFaceNo.Text != "")
                    {
                       // cnn.Open();
                        cmd = new OdbcCommand("UPDATE ODASPPlot SET NoOfMasts='" + txtFaceNo.Text + "' Where PlotNo = '" + txtPlotNo.Text + "'", cnn);
                        cmd.ExecuteNonQuery();

                    }



                    
                }
                else {
                    //GeneralVariables GeneralVariables = new GeneralVariables();
                    OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                    String StrCouncilAccountNo;
                    String str;
                    cnn2.Open();
                    OdbcDataReader odbcreader;
                    cmd = new OdbcCommand("SELECT COUNT(*) FROM ODASPPlot WHERE CouncilCode = '" + txtCouncil.Text + "'", cnn2);
                    str = cmd.ExecuteScalar().ToString ();

                    if ((Convert .ToInt32 (str))!=0)
                    {
                        if (Convert.ToInt32(str) <= 8)
                        {
                            StrCouncilAccountNo = txtCouncil.Text + "-" + "00" + Convert.ToInt32(str) + 1;
                        }
                        else if ((Convert.ToInt32(str) <= 9) && Convert.ToInt32(str) < 100)
                        {
                            StrCouncilAccountNo = txtCouncil.Text + "-" + "0" + Convert.ToInt32(str) + 1;
                        }
                        else { StrCouncilAccountNo = txtCouncil.Text + "-" + Convert.ToInt32(str) + 1; }

                    }
                    else
                    {
                        StrCouncilAccountNo = txtCouncil.Text + "-" + "001";
                    }

                    cnn2.Close();
                    generatePlotNo();
                    
                    cnn.Close();
                    cnn.Open();
                    cmd = new OdbcCommand("INSERT INTO ODASPPlot(CouncilAccountNo,PlotNo,CreatedBy,DateCreated,Created,Approved,"+
                        "Authorized,RentPaid,RatePaid,NoOfMasts,ContractStatus,ContractStatusDate,RateStatus,"+
                        "RateStatusDate,RentStatus,RentStatusDate,"+
                        "Status,AnnualRentIncrement,PaymentMode,AnnualRentIncrementType,IncrementStartYear,IncrementFrequency,"+
                        "PaymentInterval,WithLease,AcquisitionDate,MeterNo,PlotName,CouncilCode,AccountNo,OnRoadReserve,"+
                        "TownCode,LeaseDuration,LRNo,PhysicalLocation,CommencementDate,expirydate,AcquiredBy,Comments) VALUES('" + StrCouncilAccountNo + "','" + txtPlotNo.Text +
                        "','" + CurrentUserName  + "','" + DateTime.Today.ToString("MM/dd/yyyy") + "','Y','N','N','" + 0 +
                        "','"+0+"','"+0+
                        "','NOT-STARTED','" + DateTime.Today.ToString("MM/dd/yyyy") + 
                        "','NOT-PAID','" + DateTime.Today.ToString("MM/dd/yyyy") +
                        "','NOT-PAID','" + DateTime.Today.ToString("MM/dd/yyyy") +
                        "','" + txtStatus .Text +
                        "','" + txtPercentageAmount.Text + "','" + txtPaymentMode.Text +
                        "','" + txtIncreamentType.Text + "','" + txtIncrementStartYear.Text + "','" + txtIncreamentStarts.Text +
                        "','" + txtPaimentInterval.Text + "','" + txtLease.Text +
                        "','" + acquisitionDate.Text + "','" + txtMeterNo .Text +
                        "','" + txtLandlord2.Text + "','" + txtCouncil.Text + 
                        "','" + txtLandlord1.Text + "','" + txtYes .Text +
                        "','" + txtTownCode.Text + "','" + numericUpDown1.Text +
                        "','" + txtLRNo.Text + "','" + txtLocation .Text +
                        "','" + CommencementDate.Text + "','" + txtExpiryDate .Text +
                        "','" + txtAcquiredBy.Text + "','" + txtComments .Text + "')", cnn);
                    cmd.ExecuteNonQuery();
                
                
                }
                cnn.Close();

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        
        }
        private void generateMastNo()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT count(*) as NoOfMasts  FROM ODASPPlotSite WHERE PlotNo = '" +txtPlotNo.Text +"'", cnn);
                String c=cmd.ExecuteScalar ().ToString ();

                if (c != "")
                {
                    txtMastNo.Text = txtPlotNo.Text.Trim() + "-" + (Convert.ToDouble(c.Trim ())) + 1;
                }
                else txtMastNo.Text = "0";
                cnn.Close();

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }

        }
        private void updateANNUALRate() {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("select * from ODASPPlotMast Where MastNo = '" +txtMastNo.Text + "'", cnn);
               reader = cmd.ExecuteReader();

                if (reader .Read() )
                {
                    cnn.Close();
                    cnn.Open();
                    cmd = new OdbcCommand("select  sum(ODASPPlotSite.Rates) as Totals from ODASPPlotSite Where PlotNo = '" +txtPlotNo.Text + "'",cnn);
                    String  c;
                    c = cmd.ExecuteScalar().ToString();
                    cnn.Close();
                    if (c == "")
                    {
                        cnn.Open();
                        cmd = new OdbcCommand("UPDATE ODASPPlotMast SET AnnualRate='" + 0 + "' Where MastNo = '" + txtMastNo.Text + "' ", cnn);
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                    }
                    else {
                        cnn.Open();
                        cmd = new OdbcCommand("UPDATE ODASPPlotMast SET AnnualRate='" + c + "' Where MastNo = '" + txtMastNo.Text + "' ", cnn);
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                    }
                    cnn.Open();
                    cmd = new OdbcCommand("select * from ODASPPlotMast Where MastNo = '" + txtMastNo.Text + "'", cnn);
                    reader = cmd.ExecuteReader();
                    if (reader .Read ()){
                        txtMastRate.Text = reader["AnnualRate"].ToString();
                        cnn.Close();
                    }
                }
               
                cnn.Close();

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }
        private void updateANNUALRent()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("select * from ODASPPlot Where PlotNo = '" +txtPlotNo.Text  + "'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    cnn.Close();
                    cnn.Open();
                    cmd = new OdbcCommand("select  sum(ODASPPlotMast.AnnualRent) as Totals from ODASPPlotMast Where PlotNo = '" +txtPlotNo.Text + "'", cnn);
                    String c;
                    c = cmd.ExecuteScalar().ToString();
                    cnn.Close();
                    if (c == "")
                    {
                        cnn.Open();
                        cmd = new OdbcCommand("UPDATE ODASPPlot SET AnnualRent='" + 0 + "' Where PlotNo = '" + txtPlotNo.Text + "' ", cnn);
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                    }
                    else
                    {
                        cnn.Open();
                        cmd = new OdbcCommand("UPDATE ODASPPlot SET AnnualRent='" + c + "' Where PlotNo = '" + txtPlotNo.Text + "' ", cnn);
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                    }
                    cnn.Open();
                    cmd = new OdbcCommand("select * from ODASPPlot Where PlotNo = '" + txtPlotNo.Text + "'", cnn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        txtAnnualRent.Text = reader["AnnualRent"].ToString();
                        
                        cnn.Close();
                    }
                }

                cnn.Close();

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
    }
        private void updateALLSites()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("select * from ODASPPlot Where PlotNo = '" +txtPlotNo.Text + "'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    cnn.Close();
                    cnn.Open();
                    cmd = new OdbcCommand("select  COUNT(plotNo) AS SiteCount from ODASPPlotSite Where PlotNo = '" +txtPlotNo.Text  +"'", cnn);
                    String c;
                    c = cmd.ExecuteScalar().ToString();
                    
                    cnn.Close();
                    if (c == "")
                    {
                        cnn.Open();
                        cmd = new OdbcCommand("UPDATE ODASPPlot SET NoofSites='" + 0 + "' Where PlotNo = '" + txtPlotNo.Text + "' ", cnn);
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                    }
                    else
                    {
                        cnn.Open();
                        cmd = new OdbcCommand("UPDATE ODASPPlot SET NoofSites='" + c + "' Where PlotNo = '" + txtPlotNo.Text + "' ", cnn);
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                    }
                    cnn.Open();
                    cmd = new OdbcCommand("select * from ODASPPlot Where PlotNo = '" + txtPlotNo.Text + "'", cnn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        txtFaceNo.Text = reader["NoofSites"].ToString();
                         
                        cnn.Close();
                    }
                    cnn.Open();
                    cmd = new OdbcCommand("select  COUNT(plotNo) AS MastCount from ODASPPlotMast Where PlotNo = '" +txtPlotNo.Text  + "'", cnn);
                    String X;
                    X = cmd.ExecuteScalar().ToString();
                    
                    cnn.Close();
                    if (X == "")
                    {
                        cnn.Open();
                        cmd = new OdbcCommand("UPDATE ODASPPlot SET NoOfMasts='" + 0 + "' Where PlotNo = '" + txtPlotNo.Text + "' ", cnn);
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                    }
                    else
                    {
                        cnn.Open();
                        cmd = new OdbcCommand("UPDATE ODASPPlot SET NoOfMasts='" + X + "' Where PlotNo = '" + txtPlotNo.Text + "' ", cnn);
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                    }
                    cnn.Open();
                    cmd = new OdbcCommand("select * from ODASPPlot Where PlotNo = '" + txtPlotNo.Text + "'", cnn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        txtNoMast.Text = reader["NoOfMasts"].ToString();
                         
                        cnn.Close();
                    }
                }

                cnn.Close();

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }

          }
        private void clearRECORD() {
            txtAmountDue.Text = "0";
            txtPropertyCode.Text = "";
            txtOtherDeatils.Text = "";
       
        }

        protected bool CheckDate(String date)
        {

            try
            {

                DateTime dt = DateTime.Parse(date);

                return true;
            }

            catch
            {

                return false;

            }

        }
        private void enableInstallments()
        {
            txtCurrenPeriod.Enabled = true;
            txtTransactionNo.Enabled = true;
            txtInstallments.Enabled = true;
            txtPayDueDate.Enabled = true;
            txtContractYear.Enabled = true;
            txtCurrenPeriod.Enabled = true;

        }
        public Boolean  validINSTALLMENT() {
           
                if (txtTransactionNo.Text == "")
                {
                    MessageBox.Show("Transaction No is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtTransactionNo.Focus();
                    return false;
                }
                else if (txtInstallments.Text == "" || Convert.ToDouble (txtInstallments.Text) <= 0)
                {
                    MessageBox.Show("Installment is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtInstallments.Focus();
                    return false;
                }
                else if (txtInstallPaymentDue.Text == "" || Convert.ToDouble (txtInstallPaymentDue.Text) <= 0)
                {
                    MessageBox.Show("The Payment Due Must be > Zero", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtInstallPaymentDue.Focus();
                    return false;
                }
                else if (CheckDate(txtPayDueDate.Text) == false)
                {
                    MessageBox.Show("The Payment Due Date Captured is Invalid ", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtPayDueDate.Focus();
                    return false;
                }
                else if (txtContractYear.Text == "")
                {
                    MessageBox.Show("Contract Year is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtContractYear.Focus();
                    return false;
                }
                else if (txtCurrenPeriod.Text == "")
                {
                    MessageBox.Show("Current Period is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtCurrenPeriod.Focus();
                    return false;
                }
                else return true;
            

            
        }
        private void saveINSTALLMENT()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASMInstallment WHERE InstallmentNo = '" +txtTransactionNo.Text + "'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {

                }
                else {
                    cmd = new OdbcCommand("INSERT INTO ODASMInstallment(ContractYear,TotalRent,PaymentDueDate,InstallmentPercent,PaymentDue,Balance,PaymentFlag,CurrentPeriod)VALUES('" + txtContractYear.Text +
                        "','"+txtInstallPaymentDue .Text +
                        "','"+txtIncreamentRecent .Text +"','"+txtInstallPaymentDue .Text +"','"+txtInstallPaymentDue .Text +"','"+txtPaymentFlag .Text +"','"+getPeriod (DateTime .Today )+"')", cnn);
                    cnn.Close();
                    cnn.Open();
                    cmd.ExecuteNonQuery();
                    cnn.Close();
                    MessageBox.Show("Changes Made Successfully");
                }
            }
            catch (Exception ex) {
                MessageBox.Show(ex.ToString () );
            }
        }
        private void GenerateContractNo() {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPLastNumbers WHERE AutoContractNo = 'Y'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {


                    int c = Convert.ToInt32(reader["ContractNo"].ToString()) + 1;

                    switch (Convert.ToInt32(reader["ContractNo"].ToString().Length))
                    {
                        case 1: txtContractNo.Text = reader["ContractNoPrefix"].ToString() + "00000" +Convert .ToInt32 ( reader["ContractNo"].ToString().Trim());
                            break;
                        case 2: txtContractNo.Text = reader["ContractNoPrefix"].ToString() + "0000" + Convert.ToInt32(reader["ContractNo"].ToString().Trim());
                            break;
                        case 3: txtContractNo.Text = reader["ContractNoPrefix"].ToString() + "000" + Convert.ToInt32(reader["ContractNo"].ToString().Trim());
                            break;
                        case 4: txtContractNo.Text = reader["ContractNoPrefix"].ToString() + "00" + Convert.ToInt32(reader["ContractNo"].ToString().Trim());
                            break;
                        case 5: txtContractNo.Text = reader["ContractNoPrefix"].ToString() + "0" + Convert.ToInt32(reader["ContractNo"].ToString().Trim());
                            break;
                        case 6: txtContractNo.Text = reader["ContractNoPrefix"].ToString() + Convert.ToInt32(reader["ContractNo"].ToString().Trim());
                            break;
                    }
                   
                    cnn.Close();
                    cnn.Open();


                    cmd = new OdbcCommand("UPDATE ODASPLAstNumbers SET ContractNo='" + c + "' ", cnn);
                    cmd.ExecuteNonQuery();
                }
                cnn.Close();

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        
        }
        private void SaveContractDetails(){

            if (txtContractNo.Text == "") {
               
            }
          // GenerateContractNo();
            GeneralVariables GeneralVariables = new GeneralVariables();
            OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
            cnn.Open();
            cmd = new OdbcCommand("SELECT * FROM ODASMLeaseAgreement WHERE ContractNo='" +txtContractNo.Text + "'",cnn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                cmd = new OdbcCommand("UPDATE ODASMLeaseAgreement SET PlotNo='" + txtPlotNo.Text +
                    "',AgreementDate='" + CommencementDate.Text +
                    "',AccountNo='" + txtLandlord1.Text +
                    "',CommencementDate='" + CommencementDate.Text +
                    "',expirydate='" + txtExpiryDate.Text +
                    "',AnnualRent='" + txtRent.Text +
                    "',PaymentMode='" + txtPaymentMode.Text +
                    "',LeaseDuration='" + numericUpDown1.Value +
                    "',AcquisitionDate='" + acquisitionDate.Text +
                    "',AnnualRentIncrement='" + txtPercentageAmount.Text +
                    "',RentVariationType='" + txtIncreamentInterval.Text +
                    "',AnnualRentIncrementType='" + txtIncreamentType.Text +
                    "',IncrementStartYear='" + txtIncrementStartYear.Text +
                    "',IncrementFrequency='" + txtIncreamentStarts.Text +
                    "',WithLease='" + txtLease.Text +
                    "',PaymentInterval='" + txtPaimentInterval.Text +
                    "',Comments='" + txtComments.Text + "' WHERE ContractNo='" + txtContractNo.Text + "' ", cnn);
                cnn.Close();
                cnn.Open();
                cmd.ExecuteNonQuery();
                cnn.Close();
            }
            else {
                GenerateContractNo();
                cmd = new OdbcCommand("INSERT INTO ODASMLeaseAgreement(ContractNo,DatePrepared,PreparedBY,"+
                    "PlotNo,AgreementDate,AccountNo,CommencementDate,expirydate,"+
                    "AnnualRent,PaymentMode,LeaseDuration,AcquisitionDate,AnnualRentIncrement,"+
                    "RentVariationType,AnnualRentIncrementType,IncrementStartYear,IncrementFrequency,WithLease,"+
                    "PaymentInterval,Comments) VALUES('"+txtContractNo .Text +"','"+DateTime .Today .ToString ("MM/dd/yyyy")+
                    "','"+CurrentUserName +"','"+txtPlotNo .Text +"','"+CommencementDate .Text +
                    "','" + txtLandlord1.Text + "','" + CommencementDate .Text + "','"+txtExpiryDate .Text +"','"+txtRent .Text +
                    "','" + txtPaymentMode.Text + "','" + numericUpDown1.Value + "','" + acquisitionDate .Text +
                    "','" + txtPercentageAmount.Text + "','" + txtIncreamentInterval.Text + "','" + txtIncreamentType.Text +
                    "','" + txtIncrementStartYear.Text + "','" + txtIncreamentStarts.Text + "','" + txtLease .Text +
                    "','" + txtPaimentInterval.Text + "','" + txtComments.Text + "')", cnn);
                cnn.Close();
                cnn.Open();
                cmd.ExecuteNonQuery();
                cnn.Close();
            }
           
        
        }
        public String   getPeriod(DateTime  x){

            String strMonth, strYr,result;
            strMonth = x.Month.ToString().Trim ();
            if(strMonth .Length ==1){

                strMonth = "0" + strMonth;

            }
            strYr = x.Year.ToString().Trim();
            result = strYr.Trim() + "/" + strMonth.Trim();
            return result;
        }
        private void GenerateInstallmentPayable() {
            try{
                Double ContractYear;
                Double ContractYear2;
                Double InstallmentPercent=0;
                Double PaymentDue;
                Double  ContractLength=1;
                Double AmountPaid=0;
                Double Balance=0;
                
                Double TotalRent=0;


            Double MaxInstallments, InstallmentAmount;
            GeneralVariables GeneralVariables = new GeneralVariables();
            OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
            OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
            OdbcConnection cnn3 = new OdbcConnection(GeneralVariables.SQLstr);
            OdbcConnection cnn4 = new OdbcConnection(GeneralVariables.SQLstr);
            cnn3.Open();
               OdbcDataReader  reader2;
               OdbcDataReader ODBCR;
            cnn.Open();
            cmd = new OdbcCommand("SELECT * FROM ODASPPaymentMode WHERE PaymentMode = '" +txtPaymentMode .Text + "'",cnn);
            reader2 = cmd.ExecuteReader();
            if (reader2.Read())
            {

                MaxInstallments = Convert.ToDouble(reader2["PaymentsInAYear"].ToString()) * Convert.ToDouble(numericUpDown1.Text);
                InstallmentAmount = Convert.ToDouble(txtAnnualRent.Text) / Convert.ToDouble(reader2["PaymentsInAYear"].ToString());
                PaymentDue = InstallmentAmount;
               
            }else {
            MaxInstallments = 0;
                InstallmentAmount =0;
            }
           // reader2.Close();
            cnn.Close();
            cnn.Open();
            for (int i=1; i <= MaxInstallments;i++ )
            {
                //==========================================================================================================================
                if (cnn.State == ConnectionState.Closed) {
                    cnn.Open();
                }
                cmd = new OdbcCommand("SELECT * FROM ODASMInstallment WHERE PlotNo='" + txtPlotNo.Text+ "' AND ContractNo = '" +txtContractNo.Text  + "' and Installment = '" + i + "'",cnn);
                reader=cmd.ExecuteReader ();
                if (reader.Read())
                {


                    String date= CommencementDate.Text .ToString();
                    DateTime dt = Convert.ToDateTime(date); 
                      
                    //cnn.Close();
                    //cnn.Open();
            cmd = new OdbcCommand("SELECT * FROM ODASPPaymentMode WHERE PaymentMode = '" +txtPaymentMode .Text + "'",cnn2);
            if (cnn2.State == ConnectionState.Open )
            {
                cnn2.Close ();
            }
                    cnn2.Open();
                    ODBCR = cmd.ExecuteReader();
         // cnn.Close();
          // cnn2.Open();
           if (ODBCR.Read())
             {
                 
                // cnn2.Close();
                // cnn2.Open();
                 //cnn.Open();
                 DateTime newDate = dt.AddMonths((i - 1) * (Convert.ToInt32((ODBCR["CoverPeriod"].ToString()))));
                 String strdat = reader["PaymentDueDate"].ToString();
                DateTime date2 = Convert.ToDateTime(strdat);
                cmd = new OdbcCommand("UPDATE ODASMInstallment SET PaymentDueDate='" + newDate + "',CurrentPeriod='" + getPeriod(date2) + "',TotalRent='"+ 0+"' WHERE PlotNo='" + txtPlotNo.Text + "' AND ContractNo = '" + txtContractNo.Text + "' and Installment = '" + i + "'", cnn3);
                cnn3.Close();
              /// cnn.Open();
                 cnn3.Open();
                 cmd.ExecuteNonQuery  ();
                 //cnn.Close();
                 if (1 * (Convert.ToInt32((ODBCR["CoverPeriod"].ToString()))) % 12 != 0)
                 {
                     
                     ContractYear = ((i * (Convert.ToInt32((reader2["CoverPeriod"].ToString())))) / 12)+1;
                      ContractYear2 = Math.Truncate(ContractYear);
                     cmd = new OdbcCommand("UPDATE ODASMInstallment SET ContractYear='" + ContractYear2 + "' WHERE PlotNo='" + txtPlotNo.Text + "' AND ContractNo = '" + txtContractNo.Text + "' and Installment = '" + i + "'", cnn3);
                    cnn3.Open();
                     cmd.ExecuteNonQuery();
                     cnn3.Close();
                 }
                 else {
                    
                     ContractYear = (i * (Convert.ToInt32((ODBCR["CoverPeriod"].ToString())))) / 12;
                      ContractYear2 = Math.Truncate(ContractYear);
                     cmd = new OdbcCommand("UPDATE ODASMInstallment SET ContractYear='" + ContractYear2 + "' WHERE PlotNo='" + txtPlotNo.Text + "' AND ContractNo = '" + txtContractNo.Text + "' and Installment = '" + i + "'", cnn3);
                     //cnn.Close();
                    // cnn3.Open();
                     cmd.ExecuteNonQuery();
                    // cnn.Close();
                     // Increase the Rent Where appropriate
                     if ((Convert.ToInt32(reader["ContractYear"].ToString())) >= (Convert.ToInt32(txtIncrementStartYear.Text.ToString ())))
                     {
                         if (Convert.ToInt32(reader["ContractYear"].ToString()) % Convert.ToInt32(txtIncreamentStarts.Text) == 0)
                         {
                             if (txtIncreamentType.Text == "P")
                             {
                                
                                 InstallmentPercent = InstallmentAmount * (Convert.ToDouble(txtPercentageAmount.Text) / 100);
                                 cmd = new OdbcCommand("UPDATE ODASMInstallment SET InstallmentPercent='" + InstallmentPercent + "' WHERE PlotNo='" + txtPlotNo.Text + "' AND ContractNo = '" + txtContractNo.Text + "' and Installment = '" + i + "'", cnn3);

                                 cnn3.Close();
                                 cnn3.Open();
                                 cmd.ExecuteNonQuery();
                                 cnn3.Close();

                             }
                             else if (txtIncreamentType.Text == "A")
                             {
                                 cmd = new OdbcCommand("UPDATE ODASMInstallment SET InstallmentPercent='" + txtPercentageAmount.Text + "' WHERE PlotNo='" + txtPlotNo.Text + "' AND ContractNo = '" + txtContractNo.Text + "' and Installment = '" + i + "'", cnn3);

                                 cnn3.Close();
                                 cnn3.Open();
                                 cmd.ExecuteNonQuery();
                                 cnn3.Close();
                             }
                             else
                             {
                                 cmd = new OdbcCommand("UPDATE ODASMInstallment SET InstallmentPercent='" + 0 + "' WHERE PlotNo='" + txtPlotNo.Text + "' AND ContractNo = '" + txtContractNo.Text + "' and Installment = '" + i + "'", cnn3);

                                 cnn3.Close();
                                 cnn3.Open();
                                 cmd.ExecuteNonQuery();
                                 cnn3.Close();

                             }
                             if (reader["InstallmentPercent"].ToString() == "")
                             {
                                 cmd = new OdbcCommand("UPDATE ODASMInstallment SET InstallmentPercent='" + 0 + "' WHERE PlotNo='" + txtPlotNo.Text + "' AND ContractNo = '" + txtContractNo.Text + "' and Installment = '" + i + "'", cnn3);

                                 cnn3.Close();
                                 cnn3.Open();
                                 cmd.ExecuteNonQuery();
                                 cnn3.Close();
                                 InstallmentAmount = Convert.ToDouble(InstallmentAmount);
                             }
                             else
                             {
                                 InstallmentAmount = Convert.ToDouble(InstallmentAmount) + (Convert.ToDouble(reader["InstallmentPercent"]));
                             }
                         }
                     }
                     else  
                     {
                         cmd = new OdbcCommand("UPDATE ODASMInstallment SET InstallmentPercent='" + 0 + "' WHERE PlotNo='" + txtPlotNo.Text + "' AND ContractNo = '" + txtContractNo.Text + "' and Installment = '" + i + "'", cnn3);
                         cnn3.Close();
                         cnn3.Open();
                         cmd.ExecuteNonQuery();
                         ///cnn.Close();
                     }
                 }
                 if (((Convert.ToDouble(reader["ContractYear"].ToString())) - 1) % (Convert.ToInt32(txtPaimentInterval.Text)) == 0)
                 {
                    
                     PaymentDue = Convert.ToDouble(InstallmentAmount) * Convert.ToDouble(txtPaimentInterval.Text);
                     cmd = new OdbcCommand("UPDATE ODASMInstallment SET PaymentDue='" + PaymentDue + "',ContractLength='" + txtPaimentInterval.Text + "' WHERE PlotNo='" + txtPlotNo.Text + "' AND ContractNo = '" + txtContractNo.Text + "' and Installment = '" + i + "'", cnn3);

                     cnn3.Close();
                     cnn3.Open();
                     cmd.ExecuteNonQuery();
                    // cnn.Close();
                 }
                 else
                 {
                     cmd = new OdbcCommand("UPDATE ODASMInstallment SET PaymentDue='" + 0 + "' WHERE PlotNo='" + txtPlotNo.Text + "' AND ContractNo = '" + txtContractNo.Text + "' and Installment = '" + i + "'", cnn3);

                     cnn3.Close();
                     cnn3.Open();
                     cmd.ExecuteNonQuery();
                     //cnn.Close();
                 }
                 String amt = reader["AmountPaid"].ToString();
                 String paydue1 = reader["PaymentDue"].ToString();
                 if (amt == ""||paydue1 =="")
                 {
                     amt = "0";
                     paydue1 = "0";
                 }
                
                 if ((Convert.ToDouble(amt)) > (Convert.ToDouble(paydue1)))
                 {
                     cmd = new OdbcCommand("UPDATE ODASMInstallment SET AmountPaid='" + Convert.ToInt32(reader["PaymentDue"].ToString()) + "',PaymentDue='" + 0 + "',Balance='"+0+"' WHERE PlotNo='" + txtPlotNo.Text + "' AND ContractNo = '" + txtContractNo.Text + "' and Installment = '" + i + "'", cnn3);

                     cnn3.Close();
                     cnn3.Open();
                     cmd.ExecuteNonQuery();
                    // cnn.Close();
                 
                 }
                  
                 Double paydue = Convert.ToDouble(reader["PaymentDue"].ToString()) - Convert.ToDouble(reader["AmountPaid"].ToString());
                 
                 cmd = new OdbcCommand("UPDATE ODASMInstallment SET PaymentDue='" + paydue + "' WHERE PlotNo='" + txtPlotNo.Text + "' AND ContractNo = '" + txtContractNo.Text + "' and Installment = '" + i + "'", cnn3);

                 cnn3.Close();
                 cnn3.Open();
                 cmd.ExecuteNonQuery();
                // cnn.Close();
                 if (Convert.ToDouble(reader["PaymentDue"].ToString()) == 0)
                 {
                     cmd = new OdbcCommand("UPDATE ODASMInstallment SET Requisitioned='Y' WHERE PlotNo='" + txtPlotNo.Text + "' AND ContractNo = '" + txtContractNo.Text + "' and Installment = '" + i + "'", cnn3);

                     cnn3.Close();
                     cnn3.Open();
                     cmd.ExecuteNonQuery();
                    // cnn.Close();
                     if (Convert.ToDouble(reader["AmountPaid"].ToString()) > 0)
                     {
                         cmd = new OdbcCommand("UPDATE ODASMInstallment SET PaymentFlag='Y' WHERE PlotNo='" + txtPlotNo.Text + "' AND ContractNo = '" + txtContractNo.Text + "' and Installment = '" + i + "'", cnn3);

                         cnn3.Close();
                         cnn3.Open();
                         cmd.ExecuteNonQuery();
                         //cnn.Close();
                     }
                     txtIncreamentInterval.Text = "";
                 }
                 Double rent;
                 rent = Convert.ToDouble(InstallmentAmount) - Convert.ToDouble(reader["AmountPaid"].ToString());
                 cmd = new OdbcCommand("UPDATE ODASMInstallment SET TotalRent='" + rent + "',PaymentDue='" + Convert.ToDouble(InstallmentAmount) + "',Balance='" + Convert.ToDouble(reader["PaymentDue"].ToString()) + "' WHERE PlotNo='" + txtPlotNo.Text + "' AND ContractNo = '" + txtContractNo.Text + "' and Installment = '" + i + "'", cnn3);

                 cnn3.Close();
                 cnn3.Open();
                 cmd.ExecuteNonQuery();
                // cnn.Close();
                 
                 cnn.Close();
            }
                    reader .Close ();
                }

                    //======================
                else {
                    OdbcConnection con = new OdbcConnection(GeneralVariables.SQLstr);
                    String invoiceNo;
                    invoiceNo = txtContractNo.Text.Trim() + "-" + i.ToString().Trim();
                    cnn.Close();
                    cnn.Open();
                    cmd = new OdbcCommand("SELECT * FROM ODASPPaymentMode WHERE PaymentMode = '" + txtPaymentMode.Text + "'", cnn);
                    reader2 = cmd.ExecuteReader();
                    if (reader2.Read())
                    {
                       
                        MaxInstallments = Convert.ToDouble(reader2["PaymentsInAYear"].ToString()) * Convert.ToDouble(numericUpDown1.Text);
                        InstallmentAmount = Convert.ToDouble(txtAnnualRent.Text) / Convert.ToDouble(reader2["PaymentsInAYear"].ToString());
                        PaymentDue = InstallmentAmount;
                         for (int x = 1; x <= MaxInstallments; x++)
                        {
                            if (cnn3.State == ConnectionState.Closed) {
                                cnn3.Open();
                            }
                            String date = CommencementDate.Text.ToString();
                            DateTime dt = Convert.ToDateTime(date);
                            DateTime newDate = dt.AddMonths((x - 1) * (Convert.ToInt32((reader2["CoverPeriod"].ToString()))));
                            
                                
                                

                            if (1 * (Convert.ToInt32((reader2["CoverPeriod"].ToString()))) % 12 != 0)
                            {

                                ContractYear = ((x * (Convert.ToInt32((reader2["CoverPeriod"].ToString())))) / 12) + 1;
                                ContractYear2 = Math.Truncate(ContractYear);
                            }
                            else
                            {

                                ContractYear = (x * (Convert.ToInt32((reader2["CoverPeriod"].ToString())))) / 12;
                                ContractYear2 = Math.Truncate(ContractYear);
                            }
                           
                            cmd = new OdbcCommand("INSERT INTO ODASMInstallment(PlotNo,ContractNo,Installment,PaymentMode,InvoiceNo," +
                                                   "AccountNo,PaymentFlag,Requisitioned,Status,StatusDate," +
                                                   "ContractYear,InstallmentPercent,ContractLength,AmountPaid,Balance,TotalRent,PaymentDueDate,CurrentPeriod) VALUES('" + txtPlotNo.Text + "','" + txtContractNo.Text +
                                                   "','" + x + "','" + txtPaymentMode.Text + "','" + invoiceNo +
                                                   "','" + txtLandlord1.Text + "','N','N','INSTAL-CREATED','" + DateTime.Today.ToString("MM/dd/yyyy") +
                                                   "','" + ContractYear2 + "','" + InstallmentPercent + "','" + ContractLength + "','" + AmountPaid + "','" + Balance + "','" + TotalRent + "','" + newDate + "','" + getPeriod(dt) + "')", cnn3);
                            cnn3.Close();
                            cnn3.Open();
                            
                            cmd.ExecuteScalar();
                            cnn3.Close();
                          
                           
                        }
                        
                    }
                    //Insert CODE GOES HERE==========================
                    return;

                    //==========================================================================================================
                    
                }

             
            }
            }catch (Exception ex){
            MessageBox .Show (ex.ToString ());
            }
        }
      
        private void showACTUALPROPERTIES1() {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASMSiteProperties,ODASPProperties where SiteNo = '" +txtSiteNo.Text + "' and ODASMSiteProperties.PropertyCode=ODASPProperties.PropertyCode;", cnn);
                reader = cmd.ExecuteReader();
                listView8.Items.Clear();
                if (reader.Read())
                {
                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["PropertyCode"].ToString());
                        if (reader["PropertyDescription"].ToString()!="")
                        {
                            lv3.SubItems.Add(reader["PropertyDescription"].ToString());
                        }
                        
                       

                        listView8.Items.Add(lv3);




                    }

                }
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message );
            }

        }
        private void enableSITE() {
            txtIncreamentInterval.Enabled = true;
            txtSiteNo.Enabled = false;
            txtPaymentMode.Enabled = false;
            txtDetails.Enabled = true;
            txtMastRate.Enabled = false;
            chkActiveSize.Checked =true ;
            txtIncreamentType.Enabled = true;
            txtContractNo.Enabled = true;
            txtFaceRate.Enabled = true;
            numericUpDown1.Enabled = true;
            txtMeterNo.Enabled = true;
            txtHighwayZone.Enabled = true;
            txtPaimentInterval.Enabled = true;
            cmbSize.Enabled = true;
            txtMastNo.Enabled = false;
            txtTownCode.Enabled = false;
            txtMastRate.Enabled = true;
            txtMastDetails.Enabled = true;
        }
        private void enableMAST() {
            txtMastNo.Enabled = false;
            cmbMedia.Enabled = true;
            txtMastRate.Enabled = false;
            txtMastDetails.Enabled = true;
            txtMeterNo.Enabled = true;
            txtTownCode.Enabled = false;
            txtAnnualRent.Enabled = true;
        }
        private void enablePLOT() {
            acquisitionDate.Enabled = false;
            txtRate.Enabled = false;
            txtRent.Enabled = false;
            txtMeterNo.Enabled = true;
            cmbCombo.Enabled = true;
            txtTownCity.Enabled = true;
            txtExpiryDate.Enabled = false;
            numericUpDown1.Enabled = true;
            txtLandlord2.Enabled = true;
            txtMobileNo.Enabled = true;
            txtAddress.Enabled = true;
            txtEmail.Enabled = true;
            txtLRNo.Enabled = true;
            txtLocation.Enabled = true;
            txtPlotNo.Enabled = true;
            CommencementDate.Enabled = true;
            txtLRNo.Enabled = true;
            txtHighwayZone.Enabled = true;
            txtCouncil.Enabled = false;
            txtLandlord1.Enabled = false;
            txtAcquiredBy.Enabled = true;
            cmbPaymenMode.Enabled = true;
            acquisitionDate.Enabled = true;
            numericUpDown1.Enabled = true;
            txtIncreamentType.Enabled = true;
            chkLease.Enabled = true;
            chkLease.Checked = true;
            txtComments.Enabled = true;
            txtPaimentInterval.Enabled = true;
            txtIncrementStartYear.Enabled = true;





        }

        private void enableOptions() {
            rdoClient.Enabled = true;
            rdoFirm.Enabled = true;
            rdoYes.Enabled = true;
            rdoNo.Enabled = true;
        }
        private void cmbPaymenMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
               
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPPaymentMode WHERE PaymentModeDescription='" + cmbPaymenMode.SelectedItem.ToString() + "'", cnn);
                reader = cmd.ExecuteReader();
             
                if (reader.Read())
                {
                   
                   
                    txtPaymentMode.Text = reader["PaymentMode"].ToString();

                }
                cnn.Close();

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }

        private void cmbMedia_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPMedia WHERE MediaDescription='" + cmbMedia.SelectedItem.ToString() + "'", cnn);
                reader = cmd.ExecuteReader();
                //dTable = new DataTable();
                // dTable.Load(reader);
                //cmbCombo.Items.Clear();
                if (reader.Read())
                {


                    txtMedia.Text = reader["MediaCode"].ToString();
                   
                }
                cnn.Close();
                txtMediaSize.Text = "";
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }

        private void cmbMedia_Click(object sender, EventArgs e)
        {
            loadMediaSize();
        }

        private void cmbSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPMediaSize WHERE MediaSize='" + cmbSize.SelectedItem.ToString() + "'", cnn);
                reader = cmd.ExecuteReader();
              
                if (reader.Read())
                {

                    txtMediaSize.Text = reader["MediaSize"].ToString();
                }
                cnn.Close();
               
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            int year = DateTime.Today.Year ;
            decimal newyear = year + numericUpDown1.Value;

            txtExpiryDate.Text = DateTime.Today.ToString("MM/dd/") + newyear;
        }

        private void txtLandlord2_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            
                searchPlots();
     // searchLandLords();
            
        }
        private void anableAll(Control ctrl)
        {
            foreach (Control c in ctrl.Controls)
            {
                anableAll(c);
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
                if (c is NumericUpDown )
                {
                    ((NumericUpDown)(c)).Enabled = true;
                }
            }
        }
        private void clearAll(Control ctrl)
        {
            foreach (Control c in ctrl.Controls)
            {
                clearAll(c);
                if (c is TextBox)
                {
                    ((TextBox)(c)).Text  = "";
                }
                
                
                if (c is ComboBox)
                {
                    ((ComboBox)(c)).Text = "";
                }
               
            }
        }
        private void listView3_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            ListView.CheckedListViewItemCollection checkedItems = listView3.CheckedItems;
           
        //    progressBar1.Maximum = listView3.Items.Count;
            foreach (ListViewItem item in checkedItems)
            {
                progressBar1.Visible = true;
                progressBar1.Minimum = 0;
                progressBar1.Value = 0;
               
             

                txtPlotNo.Text = item.SubItems[0].Text;
                loadFaces();
               
                progressBar1.Value = progressBar1.Value + 1;
                loadPaymentMode();
                progressBar1.Value = progressBar1.Value+1;
                
                progressBar1.Value = progressBar1.Value + 1;
                
                progressBar1.Value = progressBar1.Value + 1;
                getContractNo();
                progressBar1.Value = progressBar1.Value + 1;
                loadInstallments();
                progressBar1.Value = progressBar1.Value + 1;
                loadMastRecord();
                progressBar1.Value = progressBar1.Value + 1;
                loadFaceRecord();
                progressBar1.Value = progressBar1.Value + 1;
                loadContracts();
                progressBar1.Value = progressBar1.Value + 1;
                loadRecord();
                progressBar1.Value = progressBar1.Value + 1;
                loadLandLordRecord();


                loadRecord();
                progressBar1.Value = progressBar1.Value + 1;
                loadLandLordRecord();
                progressBar1.Value ++;
                filterLandLords();

               
                progressBar1.Value = progressBar1.Value + 1;
                progressBar1.Value = 0;
                progressBar1.Visible = false;
            }
            
        }

        private void txtPlotNo_TextChanged(object sender, EventArgs e)
        {
            filterPlots();
           
        }
        private void loadCouncil() {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPCouncil WHERE CouncilCode = '" + txtCouncil .Text +  "'", cnn);
                reader = cmd.ExecuteReader();
              
                if (reader.Read())
                {


                    cmbCombo.Items.Add(reader["Council"].ToString());
                }
                cnn.Close();

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }
        private void loadPaymentMode() {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPPaymentMode WHERE PaymentMode = '"+ txtPaymentMode .Text  + "'", cnn);
                reader = cmd.ExecuteReader();
                //dTable = new DataTable();
                // dTable.Load(reader);
                //cmbCombo.Items.Clear();
                if (reader.Read())
                {


                   // cmbCombo.Items.Add(reader["PaymentModeDescription"].ToString());
                    cmbPaymenMode.Text = reader["PaymentModeDescription"].ToString();
                }
                cnn.Close();

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }
        private void getContractNo() {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASMLeaseAgreement WHERE PlotNo = '" + txtPlotNo.Text + "'", cnn);
                reader = cmd.ExecuteReader();
          
                if (reader.Read())
                {


                    txtContractNo.Text = reader["ContractNo"].ToString();
                }
                cnn.Close();

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }

        private void txtMastNo_TextChanged(object sender, EventArgs e)
        {
            //
        }

        private void listView1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            ListView.CheckedListViewItemCollection checkedItems = listView1.CheckedItems;

            foreach (ListViewItem item in checkedItems)
            {
                txtLandlord1.Text = "";

                txtLandlord1.Text = item.SubItems[0].Text;
                loadLandLordRecord();
                filterLandLords();
                loadContracts();
            }
        }

        private void listView7_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            ListView.CheckedListViewItemCollection checkedItems = listView7.CheckedItems;

             foreach (ListViewItem item in checkedItems)
             { 
             
             txtSerialNo .Text =item .SubItems [0].Text ;
           
             }
        }
        private void loadInstallmentRecord(){

            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASMInstallment  WHERE InstallmentNo = '"+ txtSerialNo .Text + "' ", cnn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    if (reader["PaymentDue"].ToString() == "")
                    {
                        txtInstallPaymentDue.Text = "0";
                    }
                    else { txtInstallPaymentDue.Text = reader["PaymentDue"].ToString(); }

                    if(reader["InstallmentPercent"].ToString()==""){
                        txtIncreamentRecent.Text = "0";
                    
                    }
                    
                        else {txtIncreamentRecent.Text = reader["InstallmentPercent"].ToString();}

                    txtInstallments.Text = reader["Installment"].ToString();

                    txtInvoiceNo.Text = reader["InvoiceNo"].ToString();

                    txtPayDueDate.Text =Convert .ToDateTime ( reader["PaymentDueDate"].ToString()).ToString ("yyyy/MM/dd");


                    txtCurrenPeriod.Text = reader["CurrentPeriod"].ToString();
                    txtContractYear.Text = reader["ContractYear"].ToString();
                    

                }
                cnn.Close();

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }

        private void txtSerialNo_TextChanged(object sender, EventArgs e)
        {
            loadInstallmentRecord();
        }

        private Boolean validFace() {
          
         String expDate;
            expDate =txtExpiryDate .Text ;
            DateTime expdt= Convert .ToDateTime (expDate);
            if (txtLandlord2.Text == "")
            {
                MessageBox.Show("Landlord Name is required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtLandlord2.Focus();
                return false;
            }
            else if (txtIncreamentType.Text != "P" && txtIncreamentType.Text != "N" && txtIncreamentType.Text != "A")
            {
                MessageBox.Show("The Annual Increment Can be (P)ercent, (A)mount or (N)one", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtIncreamentType.Focus();
                return false;
            }
            else if ((txtLRNo.Text == "") && (rdoNo.Checked = true))
            {
                MessageBox.Show("LR No is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtLRNo.Focus();
                return false;
            }
            else if (Convert.ToDouble(txtPercentageAmount.Text) <= 0 && txtIncreamentType.Text != "N")
            {
                MessageBox.Show("You Must Indicate the Basis of the Annual Rent Increment as either P or Amount", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtPercentageAmount.Focus();
                return false;
            }
            else if ((txtPaymentMode.Text == "") && (rdoNo.Checked = true))
            {
                MessageBox.Show("The Payment Mode cannot be Left Blank, there is need to Compute the Installment Due", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cmbPaymenMode.Focus();
                return false;
            }
            else if (txtAddress.Text == "")
            {
                MessageBox.Show("The Postal Address is Required ", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtAddress.Focus();
                return false;

            }
            else if (txtMobileNo.Text == "")
            {

                MessageBox.Show("The Mobile Phone Number for the Contact is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMobileNo.Focus();
                return false;
            }
            else if (txtEmail.Text == "")
            {

                MessageBox.Show("The Email address of the Contact is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtEmail.Focus();
                return false;
            }
            else if (txtCouncil.Text == "")
            {
                MessageBox.Show("The Council is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cmbCombo.Focus();

                return false;
            }
            else if (txtMediaSize.Text == "")
            {

                MessageBox.Show("The Media Size is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cmbSize.Focus();

                return false;
            }
            else if (txtMedia.Text == "")
            {
                MessageBox.Show("The Media type is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cmbMedia.Focus();

                return false;

            }else if(acquisitionDate .Text ==""){

                MessageBox.Show("The Acquisition Date is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                acquisitionDate.Focus();

                return false;
            }
            
            else if(Convert .ToDateTime (acquisitionDate.Value)>DateTime .Today ){
            MessageBox.Show("The Acquistion Date cannot be in the Future", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                acquisitionDate.Focus();

                return false;
            }
            else if (txtExpiryDate.Text == "") {
                MessageBox.Show("The Expiry Date is Necessary ", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtExpiryDate.Focus();

                return false;
            
            }else if(CommencementDate .Text ==""){
                MessageBox.Show("The Commencement date of the Site is Required ", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                CommencementDate.Focus();

                return false;

            }else if(txtExpiryDate .Text ==""){
                MessageBox.Show("The  expiry date is required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtExpiryDate.Focus();

                return false;
            }
            else if (Convert.ToDateTime(CommencementDate.Text.ToString()) > expdt)
            {
                 MessageBox.Show("The Commencement Date cannot be greater than the expiry date", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                CommencementDate.Focus();

                return false;
            }else if(numericUpDown1 .Value <0){
                MessageBox.Show("The Duration of the Lease is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                numericUpDown1.Focus();
                return false;

            }else if(txtLRNo .Text ==""){
                MessageBox.Show("The LR Number is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtLRNo.Focus();
                return false;
           
            }else if(txtLocation .Text ==""){
                MessageBox.Show("The Physical Address of the Site is Necessary", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtLocation.Focus();
                return false;
            
            }else if(txtClient .Text ==""){
                MessageBox.Show("The specification of the Structure / Mast Ownership is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                rdoClient.Focus();
                 
                return false;
            }else if(rdoFirm .Checked ==true && txtAnnualRent .Text =="0" && rdoNo .Checked ==true ){
                MessageBox.Show("The Annual Rent for the Mast / Structure cannot be ZERO!", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtAnnualRent.Focus();
                 
                return false;

            }
            else if (rdoClient.Checked == true && Convert.ToDouble(txtAnnualRent.Text) > 0)
            {
                MessageBox.Show(" The Annual Rent is not Payable when the Mast / Structure is Owned by the Client!.Set it to ZERO.", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtAnnualRent.Focus();
                 
                return false;
            }else if(txtYes .Text ==""){
             MessageBox.Show("The specification of the plot (whether on Road Reserve or not) is Required.", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
             rdoYes.Focus();
                  
                return false;
            }else if(rdoYes .Checked ==false && Convert.ToDouble (txtAnnualRent .Text )==0 && rdoFirm .Checked ==true ){
                MessageBox.Show("The Annual rent For  the Mast / Structure cannot be ZERO!", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtAnnualRent.Focus();
                  
                return false;
            
            }else if(txtTownCode .Text ==""){
                MessageBox.Show("The Town Code is Required.", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtTownCode.Focus();
                  
                return false;
            }else if(txtMastRate .Text ==""){
                MessageBox.Show(" The Annual Rate for the BillBoard is Required.", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMastRate.Focus();
                  
                return false;

            }else if (txtDetails .Text == "")
            {
                MessageBox.Show(" The Exact Location of the Mast is Required.", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtDetails.Focus();

                return false;
            }
            else if (txtFaceRate.Text == "")
            { 
                MessageBox.Show("The Annual Rate for the Site is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtFaceRate.Focus();

                return false;
            }
            else if (txtMeterNo.Text == "")
            { 
                MessageBox.Show("The Meter No is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMeterNo.Focus();

                return false;
            }else if(txtHighwayZone .Text ==""){
                MessageBox.Show(" The Highway Zone is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtHighwayZone.Focus();

                return false;

            }else if (Convert .ToInt32 ( txtIncrementStartYear.Text) <0)
            {
                MessageBox.Show("Invalid Increment Start Installment", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtIncrementStartYear.Focus();

                return false;

            }
            else if (Convert.ToInt32(txtIncreamentStarts.Text) < 0)
            {
                MessageBox.Show("Invalid Increment Frequency", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtIncreamentStarts.Focus();

                return false;

            }
            else
            {
                return true;
            }
        }  
        private void btnSave_Click(object sender, EventArgs e)
        {
            try {
               
                if (validFace())
                {
                  
                    DialogResult DialogResult;
                    DialogResult = new DialogResult();
                    if (MessageBox.Show("Are you sure you want to Perform this Action?", "Confirm Save", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                       
                        progressBar1.Visible = true;
                        progressBar1.Value = progressBar1.Value + 1;
                       
                        saveRecord();
                        progressBar1.Value = progressBar1.Value + 1;
                      
                        loadLandLords();
                        progressBar1.Value = progressBar1.Value + 1;
                      
                        updateRecord();
                        progressBar1.Value = progressBar1.Value + 1;
                       
                        saveMast();
                        progressBar1.Value = progressBar1.Value + 1;
                      
                        savaSite();
                        progressBar1.Value = progressBar1.Value + 1;
                       
                        updateANNUALRate();

                        progressBar1.Value = progressBar1.Value + 1;
                     
                        updateANNUALRent();
                        progressBar1.Value = progressBar1.Value + 1;
                       
                        updateALLSites();
                        progressBar1.Value = progressBar1.Value + 1;
                      
                        SaveContractDetails();
                        progressBar1.Value = progressBar1.Value + 1;
                       
                        GenerateInstallmentPayable();
                        progressBar1.Value = progressBar1.Value + 1;
                        
                        loadMasts();
                        progressBar1.Value = progressBar1.Value + 1;
                        loadPlots();
                        
                        loadFaces();

                        loadInstallments();
                       
                        loadLandLordRecord();
                        progressBar1.Value = progressBar1.Value + 1;
                        showALLPROPERTIES1();
                        showACTUALPROPERTIES1();
                        progressBar1.Value = progressBar1.Value + 1;
                     
                      
                        MessageBox.Show("Record Saved Successfully");
                        progressBar1.Value = 0;
                        progressBar1.Visible = false;
                    }
                }
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
        }

        private void rdoYes_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoYes.Checked == true)
            {
                txtYes.Text = "Y";
            }
        }

        private void rdoNo_CheckedChanged(object sender, EventArgs e)
        {
            if(rdoNo .Checked ==true ){
                txtYes.Text = "N";
            }
            
        }

        private void chkLease_CheckedChanged(object sender, EventArgs e)
        {
            if (chkLease.Checked == true)
            {
                txtLease.Text = "Y";
            }
            else txtLease.Text = "N";
        }

        private void rdoClient_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoClient.Checked == true) {
                txtClient.Text = "Y";
            }
        }

        private void rdoFirm_CheckedChanged(object sender, EventArgs e)
        {
            if(rdoFirm.Checked ==true ){
                txtClient.Text = "N";
            }
        }

        private void chkActiveSize_CheckedChanged(object sender, EventArgs e)
        {
            if(chkActiveSize .Checked ==true ){
                txtActive.Text = "Y";
            }else txtActive .Text ="N";
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            try {

                if (txtPlotNo.Text == "")
                {
                    MessageBox.Show("No Record selected to edit", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else {
                    GeneralVariables vars = new GeneralVariables();
                    vars.entryinProgress = true;
                    btnSave.Text = "       Save Changes";
                    enableMAST();
                    enableOptions();
                    enablePLOT();
                    enableSITE();
                    btnHelp.Enabled = true;
                    txtPlotNo.Enabled = false;
                    txtPercentageAmount.Enabled = true;
                    enableInstallments();
                }
            }catch (Exception ex){
                MessageBox.Show(ex.Message );
            }
        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
           
        }

        private void tabPage3_Click(object sender, EventArgs e)
        {
            label54.Text = "Search Plot";
        }

        private void btnHelp_Click(object sender, EventArgs e)
        {
            try {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               
                if (MessageBox.Show("Are you sureyou want to delete this record?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                    cmd = new OdbcCommand("DELETE FROM ODASPPlot WHERE PlotNo LIKE '" +txtPlotNo.Text + "'", cnn);
                    cnn.Open();
                    cmd.ExecuteNonQuery();
                    cnn.Close();
                    cmd = new OdbcCommand("DELETE FROM ODASPPlotMast WHERE PlotNo LIKE '" + txtPlotNo.Text + "'", cnn);
                    cnn.Open();
                    cmd.ExecuteNonQuery();
                    cnn.Close();
                    cmd = new OdbcCommand("DELETE FROM ODASPPlotSite WHERE PlotNo LIKE '" +txtPlotNo.Text + "'", cnn);
                    cnn.Open();
                    cmd.ExecuteNonQuery();
                    cnn.Close();
                    cmd = new OdbcCommand("DELETE FROM ODASMInstallment WHERE ContractNo LIKE '" +txtPlotNo.Text + "'", cnn);
                    cnn.Open();
                    cmd.ExecuteNonQuery();
                    cnn.Close();
                    MessageBox.Show("Record deleted Successfully","Successful",MessageBoxButtons .OK , MessageBoxIcon .Information );
                }
                
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message );
            }
                
        }

        private void txtLocation_TextChanged(object sender, EventArgs e)
        {
            txtMastDetails.Text = txtLocation.Text;
        }

        private void txtContractNo_TextChanged(object sender, EventArgs e)
        {
            loadInstallments();
        }
       
        private void btnChange_Click(object sender, EventArgs e)
        {
            try { 
            if (btnChange .Text =="Change"){
                btnChange.Text = "Save";
                txtIncreamentRecent .Enabled =false ;
                txtInstallPaymentDue.Enabled =false ;
            }
            else if (btnChange.Text == "Save")
            {
                btnChange.Text = "Change";
                if (validINSTALLMENT()) {
                    saveINSTALLMENT();
                    loadInstallments();

                }
                saveINSTALLMENT();
                loadInstallments();

            }
            }catch (Exception ex){
                MessageBox.Show(ex.ToString () );
            }
        }

        private void chkPaymentFlag_CheckedChanged(object sender, EventArgs e)
        {
            if(chkPaymentFlag.Checked ==true ){
                txtPaymentFlag.Text = "Y";
            }
            else txtPaymentFlag.Text = "N";
        }
        private void enableRECORD() {
            txtAmountDue.Enabled = true;
            txtPropertyCode.Enabled = true;
            txtOtherDeatils.Enabled = true;
            startDate.Enabled = true;
            dateAssigned.Enabled = true;
        }
       
        private void btnRight_Click(object sender, EventArgs e)
        {
            try {
               
            if(txtPlotNo .Text ==""){
                MessageBox.Show("Plot No Missing", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            else if (txtMastNo.Text == "")
            { 
                MessageBox.Show("Mast No Missing", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            else if (txtSiteNo.Text == "")
            {
                MessageBox.Show("Site No Missing", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            else {
                if (btnRight.Text.Trim () == ">>") {

                    btnRight.Text = "Save";
                    btnLeft.Enabled = false;
                    clearRECORD();
                    enableRECORD();
                    showALLPROPERTIES1();
                    showACTUALPROPERTIES1();
                }else if(btnRight.Text == "Save"){
                     ListView.ListViewItemCollection  checkedItems = listView2.Items;

            
                j=listView2 .Items .Count ;
                k = 0;
                if (j != 0) {
                    for (i = 1; i <= j; i++) {
                        foreach (ListViewItem item in checkedItems)
                        {
                            if(item.Checked  ==true ){
                                k = k + 1;
                             }
                    }
                    }
                    if (ValidRecord())
                    {
                        if (MessageBox.Show("Are you sure you want to perform this action?", "Confirm Save", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                            saveProperty();
                            disableALLRECORD();
                            showACTUALPROPERTIES1();
                            showALLPROPERTIES1();
                            btnRight.Text = ">>";
                            btnLeft.Text = "<<";
                            btnLeft.Enabled = true;
                            MessageBox.Show("Record Saved Successfully","Successful");
                        }
                        
                        
                    }

                }

                }
            }
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
        }
        private Boolean  ValidRecord() {
            if (txtSiteNo.Text == "") {
                MessageBox.Show("Site No is required","Information Required",MessageBoxButtons .OK ,MessageBoxIcon.Exclamation );
                txtSiteNo.Focus();
                return false;

            }else if(k==0){
                MessageBox.Show("Please choose one or more properties Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtSiteNo.Focus();
                return false;
            }else 
            return true;
        }
        private void btnLeft_Click(object sender, EventArgs e)
        {

        }

        private void txtYes_TextChanged(object sender, EventArgs e)
        {

        }

        private void CommencementDate_ValueChanged(object sender, EventArgs e)
        {
           
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            GeneralVariables vars = new GeneralVariables();
            vars.rptPlotSites2.currentRecord = txtPlotNo.Text;
            vars.rptPlotSites2.ShowDialog();
        }

        private void txtLandlord2_MouseLeave(object sender, EventArgs e)
        {
           
        }

        private void txtLandlord2_Leave(object sender, EventArgs e)
        {
            txtLandlord2.Text = txtLandlord2.Text.ToUpper();
        }

        private void bindingNavigatorPositionItem_Click(object sender, EventArgs e)
        {

        }

        private void bindingNavigatorMoveLastItem_Click(object sender, EventArgs e)
        {

        }

        private void txtLocation_Leave(object sender, EventArgs e)
        {
            txtLocation.Text = txtLocation.Text.ToUpper();
        }

        private void frmSite_Acquisition_FormClosing(object sender, FormClosingEventArgs e)
        {
            MessageBox.Show("Data entry ");
            return;
            GeneralVariables vars = new GeneralVariables();
            if(vars.entryinProgress == true){
                MessageBox.Show("Data entry in Progress. If you really want to cancel current operation click refresh. NOTE all unsaved data will be lost.","Data Entry in progress",MessageBoxButtons.OK ,MessageBoxIcon .Warning );
                return;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GeneralVariables vars = new GeneralVariables();
            if ( MessageBox.Show("NOTE all unsaved data will be lost.", "Data Entry in progress", MessageBoxButtons.YesNo , MessageBoxIcon.Warning)==DialogResult .Yes )
            {
               
                vars.entryinProgress = false;
                anableAll(this);
                clear();
                txtPlotNo.Enabled = false;
                txtLandlord1.Enabled = false;
                txtPercentageAmount.Enabled = true;
                txtIncreamentType.Enabled = true;
                txtTownCode.Enabled = false;
                txtExpiryDate.Enabled = false;
            }
        }

        private void frmSite_Acquisition_FormClosed(object sender, FormClosedEventArgs e)
        {
            MessageBox.Show("Data entry2 ");
            return;
        }

       
    }
}
