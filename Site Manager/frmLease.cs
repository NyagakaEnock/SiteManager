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
    public partial class frmLease : Form
    {
        public frmLease()
        {
            InitializeComponent();
        }
        Double PaymentsInAYear;
        OdbcCommand cmd;
        OdbcDataReader reader;
        DataTable dTable;
        DataSet ds;
        public string CurrentUserName;
        OdbcDataAdapter da;
        String ContarctNo;
        private void loadPaymentModes()
        {
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
                cboPaymentMode.Items.Clear();
                 
                foreach (DataRow drow in dTable.Rows)
                {
                    cboPaymentMode.Items.Add(drow["PaymentModeDescription"].ToString());


                }
                cnn.Close();
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }
        private void cboPaymentMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPPaymentMode WHERE PaymentModeDescription='" + cboPaymentMode.SelectedItem.ToString() + "'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {


                    txtpaymentMode.Text = reader["PaymentMode"].ToString();
                     
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
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn2.Open();
                cnn.Open();

                string sql = "SELECT * FROM ODASPAccount Where Status = 'A' AND AccountType = 'LLORD' oRDER BY AccountNo";

                cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView3.Items.Clear();
                listView3.Columns.Clear();
                listView3.Columns.Add("Land Lord No", listView3.Width / 3);
                listView3.Columns.Add("Names", listView3.Width / 3);
                listView3.Columns.Add("Status", listView3.Width / 3);


                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["AccountNo"].ToString());
                        if (reader["CompanyName"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["CompanyName"].ToString());
                        }
                        if (reader["Status"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["Status"].ToString());
                        }


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
        private void getMastsToLease()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn2.Open();
                cnn.Open();

                string sql = "SELECT * FROM ODASPPlotMast Where PlotNo = '" +txtPlotNo.Text + "' AND OwenedByClient = 'N'";

                cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView2.Items.Clear();
                listView2.Columns.Clear();
                listView2.Columns.Add("Structure No", listView3.Width / 3);
                listView2.Columns.Add("Media", listView3.Width / 3);
                listView2.Columns.Add("Size", listView3.Width / 3);


                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["MastNo"].ToString());
                        if (reader["TypeOfMast"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["TypeOfMast"].ToString());
                        }
                        if (reader["MediaSize"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["MediaSize"].ToString());
                        }


                        listView2.Items.Add(lv3);





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
        private void getLeasableMasts() {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn2.Open();
                cnn.Open();

                string sql = "SELECT * FROM ODASPPlotMast Where PlotNo = '" +txtPlotNo.Text +"' AND OwenedByClient = 'N' and (LeasePrepared ='N' or LeasePrepared is null)";

                cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView2.Items.Clear();
                listView2.Columns.Clear();
                listView2.Columns.Add("Structure No", listView3.Width / 3);
                listView2.Columns.Add("Media", listView3.Width / 3);
                listView2.Columns.Add("Size", listView3.Width / 3);


                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["MastNo"].ToString());
                        if (reader["TypeOfMast"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["TypeOfMast"].ToString());
                        }
                        if (reader["MediaSize"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["MediaSize"].ToString());
                        }


                        listView2.Items.Add(lv3);





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
        private void showALLLandLORDSites() {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn2.Open();
                cnn.Open();

                string sql = "SELECT * FROM ODASMLeaseAgreement where AccountNo = '" +txtLandLordNo.Text + "'";

                cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView4.Items.Clear();
                listView4.Columns.Clear();
                listView4.Columns.Add("Contract", listView4.Width / 6);
                listView4.Columns.Add("Plot", listView4.Width / 6);
                listView4.Columns.Add("LandLord", listView4.Width / 6);
                listView4.Columns.Add("Agreement Date", listView4.Width / 6);
                listView4.Columns.Add("Signed", listView4.Width / 6);
                listView4.Columns.Add("Signed By", listView4.Width / 6);


                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["ContractNo"].ToString());
                        if (reader["ContractNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PlotNo"].ToString());
                        }
                        if (reader["AccountNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["AccountNo"].ToString());
                        }
                        if (reader["AsSigned"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["AsSigned"].ToString());
                        } if (reader["AgreementDate"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["AgreementDate"].ToString());
                        } if (reader["SignedBy"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["SignedBy"].ToString());
                        }

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
        private void setALLAcquiredSites()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn2.Open();
                cnn.Open();
              string sql = "SELECT *  FROM ODASPPlot  where (OnRoadReserve = 'N' or OnRoadReserve is null)";

                cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                ListProperties.Items.Clear();
                ListProperties.Columns.Clear();
                ListProperties.Columns.Add("Plot No", ListProperties.Width / 3);
                ListProperties.Columns.Add("Plot Name", ListProperties.Width / 3);
                ListProperties.Columns.Add("Physical Location", ListProperties.Width / 3);


                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["PlotNo"].ToString());
                        if (reader["PlotName"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PlotName"].ToString());
                        }
                        if (reader["PhysicalLocation"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PhysicalLocation"].ToString());
                        }


                        ListProperties.Items.Add(lv3);

                        



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
        private void GenerateContractNo()
        {
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
                        case 1: txtContractNo.Text = reader["ContractNoPrefix"].ToString() + "00000" + Convert.ToInt32(reader["ContractNo"].ToString().Trim());
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
        private void saveRecord() {
            try {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                int K,J;
                K = 0;
                J = listView2.Items.Count;
                ListView.ListViewItemCollection items = listView2.Items;
                ListView.CheckedListViewItemCollection checkedItems = ListProperties.CheckedItems;
                for (int i = 0; i <= J;i++ )
                {
                    foreach (ListViewItem item in checkedItems)
                    {
                    if(item .Checked ==true ){
                        textBox1.Text = checkedItems.Count.ToString();
                          K = checkedItems.Count;
                    }
                }
                }
               
                
                cnn2.Open();
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASMLeaseAgreement  WHERE ContractNo = '" +txtContractNo.Text + "'",cnn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    cmd = new OdbcCommand("UPDATE ODASMLeaseAgreement SET PaymentMode='" + txtpaymentMode.Text +
                        "',PlotNo='" + txtPlotNo.Text +
                        "',AgreementDate='" + DTPickerAgreementDate.Text +
                        "',AccountNo='" + txtLandLordNo.Text + "',NoOfBillBoards='" + K + "',WitnessLandLord='" + txtWitnessLandLord.Text + "' WHERE  ContractNo = '" + txtContractNo.Text + "'", cnn2);
                    cmd.ExecuteNonQuery();
                    cnn2.Close();
                    cnn2.Open();
                    if (chkYes.Checked == true)
                    {
                        cmd = new OdbcCommand("UPDATE ODASMLeaseAgreement SET IncreamentAnnualRent='Y',PercentageIncreament='" + txtPercentage.Text + "'  WHERE  ContractNo = '" + txtContractNo.Text + "'", cnn2);
                        cmd.ExecuteNonQuery();
                        cnn2.Close();
                    }
                    else
                    {
                        cmd = new OdbcCommand("UPDATE ODASMLeaseAgreement SET IncreamentAnnualRent='N',PercentageIncreament='" + 0 + "'  WHERE  ContractNo = '" + txtContractNo.Text + "'", cnn2);
                        cmd.ExecuteNonQuery();
                        cnn2.Close();
                    }
                    if (OptStandard.Checked == true)
                    {
                        cnn2.Open();
                        cmd = new OdbcCommand("UPDATE ODASMLeaseAgreement SET Standard='Y'  WHERE  ContractNo = '" + txtContractNo.Text + "'", cnn2);
                        cmd.ExecuteNonQuery();
                        cnn2.Close();
                    }
                    else
                    {
                        cnn2.Open();
                        cmd = new OdbcCommand("UPDATE ODASMLeaseAgreement SET Standard='N'  WHERE  ContractNo = '" + txtContractNo.Text + "'", cnn2);
                        cmd.ExecuteNonQuery();
                        cnn2.Close();
                    }
                    cnn.Close();
                }
                   
                else {
                    cnn.Close();
                    cnn.Open();
                    GenerateContractNo();
                    cmd = new OdbcCommand("INSERT INTO ODASMLeaseAgreement (ContractNo,CompanyCode,Preparedby,dateprepared,Renewal,Status,Authorized,PlotNo) " +
                                               "VALUES('" + txtContractNo.Text + "','MAG','" + GeneralVariables.CurrentUserName + "','" + DateTime.Today + "','" + 1 + "','DRAFT AGREEMENT','N','" + txtPlotNo.Text + "')", cnn);
                    cmd.ExecuteNonQuery();
                    cnn.Close();
                    if (chkDeallocate.Checked == true)
                    {
                        cnn.Open();
                        cmd = new OdbcCommand("UPDATE ODASMLeaseAgreement SET AsSigned='N'  WHERE  ContractNo = '" + txtContractNo.Text + "'", cnn2);
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                    }
                    else 
                    {
                        cnn.Open();
                        cmd = new OdbcCommand("UPDATE ODASMLeaseAgreement SET AsSigned='Y'  WHERE  ContractNo = '" + txtContractNo.Text + "'", cnn2);
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                    }
                    cnn2.Close();
                    cnn2.Open();
                    cmd = new OdbcCommand("select * from ODASPdefault ", cnn2);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        if (reader["AutoApproval"].ToString() == "Y")
                        {
                            cnn.Open();
                            cmd = new OdbcCommand("UPDATE ODASMLeaseAgreement SET Approved='Y'  WHERE  ContractNo = '" + txtContractNo.Text + "'", cnn);
                            cmd.ExecuteNonQuery();
                            cnn.Close();
                        }
                        else {
                            cnn.Open();
                            cmd = new OdbcCommand("UPDATE ODASMLeaseAgreement SET Approved='N'  WHERE  ContractNo = '" + txtContractNo.Text + "'", cnn);
                            cmd.ExecuteNonQuery();
                            cnn.Close();
                        }
                        reader.Close();
                        //=============================================
                        cmd = new OdbcCommand("UPDATE ODASMLeaseAgreement SET PaymentMode='" + txtpaymentMode.Text +
         "',PlotNo='" + txtPlotNo.Text +
         "',AgreementDate='" + DTPickerAgreementDate.Text +
         "',AccountNo='" + txtLandLordNo.Text + "',NoOfBillBoards='" + K + "',WitnessLandLord='" + txtWitnessLandLord.Text + "' WHERE  ContractNo = '" + txtContractNo.Text + "'", cnn2);
                        cmd.ExecuteNonQuery();
                        cnn2.Close();
                        cnn2.Open();
                        if (chkYes.Checked == true)
                        {
                            cmd = new OdbcCommand("UPDATE ODASMLeaseAgreement SET IncreamentAnnualRent='Y',PercentageIncreament='" + txtPercentage.Text + "'  WHERE  ContractNo = '" + txtContractNo.Text + "'", cnn2);
                            cmd.ExecuteNonQuery();
                            cnn2.Close();
                        }
                        else
                        {
                            cmd = new OdbcCommand("UPDATE ODASMLeaseAgreement SET IncreamentAnnualRent='N',PercentageIncreament='" + 0 + "'  WHERE  ContractNo = '" + txtContractNo.Text + "'", cnn2);
                            cmd.ExecuteNonQuery();
                            cnn2.Close();
                        }
                        if (OptStandard.Checked == true)
                        {
                            cnn2.Open();
                            cmd = new OdbcCommand("UPDATE ODASMLeaseAgreement SET Standard='Y'  WHERE  ContractNo = '" + txtContractNo.Text + "'", cnn2);
                            cmd.ExecuteNonQuery();
                            cnn2.Close();
                        }
                        else
                        {
                            cnn2.Open();
                            cmd = new OdbcCommand("UPDATE ODASMLeaseAgreement SET Standard='N'  WHERE  ContractNo = '" + txtContractNo.Text + "'", cnn2);
                            cmd.ExecuteNonQuery();
                            cnn2.Close();
                        }
                        cnn.Close();
                    }
                } //=============================================
            }catch (Exception ex){
               MessageBox .Show (ex.ToString());
            }
        }
        private void updateSITE() {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("select * from ODASPPlot Where PlotNo = '" +txtPlotNo.Text  +"'",cnn);
                reader = cmd.ExecuteReader();

                if(reader .Read ()){
                    cnn.Close();
                    cnn.Open();
                    cmd = new OdbcCommand("UPDATE ODASPPlot SET AccountNo='"+txtLandLordNo .Text +"' Where PlotNo = '" + txtPlotNo.Text + "' ", cnn);
                    cmd.ExecuteNonQuery();
                    cnn.Close();
                    if (chkDeallocate.Checked == true)
                    {
                        cnn.Close();
                        cnn.Open();
                        cmd = new OdbcCommand("UPDATE ODASPPlot SET Status='UN-ALLOCATED' Where PlotNo = '" + txtPlotNo.Text + "' ", cnn);
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                    }
                    else {
                        
                            cnn.Close();
                            cnn.Open();
                            cmd = new OdbcCommand("UPDATE ODASPPlot SET Status='SITE-ACQUIRED' Where PlotNo = '" + txtPlotNo.Text + "' ", cnn);
                            cmd.ExecuteNonQuery();
                        cnn.Close();
                        
                    }
                }
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
        }
        private void updateLeasedPlotMasts(){
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASPPlotMast WHERE MastNo = '" + mastNo.Text   + "'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read()) {
                    cnn.Close();
                    cnn.Open();
                    cmd = new OdbcCommand("UPDATE ODASPPlotMast SET LeasePrepared='Y',ContractNo='"+txtContractNo .Text +"' WHERE MastNo = '" + mastNo.Text + "'", cnn);
                    cmd.ExecuteNonQuery();
                    cnn.Close();
                }
            }catch (Exception ex){
            MessageBox .Show (ex.ToString ());
            }
       }
        private void upDateLeaseAnnualRent()
        {
            
                try
                {
                    GeneralVariables GeneralVariables = new GeneralVariables();
                    OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                    OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                    OdbcDataReader rdr;
                  
                    String c;
                    cnn.Open();
                    cmd = new OdbcCommand("SELECT sum(AnnualRent) FROM ODASPPlotMast WHERE ContractNo = '" +txtContractNo.Text +"'", cnn);
                    c = cmd.ExecuteScalar().ToString ();
                    cnn.Close();
                    cnn.Open();
                    cmd = new OdbcCommand("SELECT  ODASPPlotMast.ExpiryDate,ODASPPlotMast.CommencementDate,ODASPPlotMast.ContractNo, ODASPPlotMast.LeaseDuration  FROM ODASPPlotMast WHERE ContractNo = '" + txtContractNo.Text + "' Group By ContractNo, LeaseDuration,CommencementDate,ExpiryDate", cnn);
                  
                    reader =cmd.ExecuteReader ();
                
                    if (reader.Read())
                    {
                        cnn2.Open();
                        cmd = new OdbcCommand("SELECT * FROM ODASMLeaseAgreement WHERE ContractNo = '" +reader["ContractNo"].ToString ()+ "'",cnn2);
                         
                        rdr=  cmd.ExecuteReader();
                        if(rdr.Read ()){
                            cnn2.Close();
                            cnn2.Open();
                            cmd = new OdbcCommand("UPDATE ODASMLeaseAgreement SET AnnualRent='" + c + "',LeaseDuration='" + reader["LeaseDuration"].ToString() + "',CommencementDate='" + reader["CommencementDate"].ToString() + "',expirydate='" + reader["expirydate"].ToString() + "' WHERE ContractNo = '" + reader["ContractNo"].ToString() + "'", cnn2);
                            cmd.ExecuteNonQuery();
                            cnn2.Close();
                           
                        }

                      
                    }
                    cnn.Close();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            
        }
        private void updateInstallments() {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn3 = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn4 = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn5 = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection CON = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcDataReader rdr;
                OdbcDataReader rdr1;
                String c;
                cnn.Open();

                cmd = new OdbcCommand("SELECT * FROM ODASMLeaseAgreement WHERE ContractNo = '" +txtContractNo.Text + "'", cnn);
                
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                   
                    if (GeneralVariables.bAllowProcess)
                    {
                        cnn2.Open();
                        cmd = new OdbcCommand("SELECT * FROM ODASPPaymentMode WHERE PaymentMode = '" + txtpaymentMode.Text + "'", cnn2);
                        rdr = cmd.ExecuteReader();
                        if (rdr.Read())
                        {
                            PaymentsInAYear = Convert.ToDouble(rdr["PaymentsInAYear"].ToString());
                        }
                    }
                    else {
                        cnn2.Open();
                        cmd = new OdbcCommand("SELECT * FROM ODASPPaymentMode WHERE PaymentMode = '" + reader["PaymentMode"].ToString ()+ "'", cnn2);
                        rdr = cmd.ExecuteReader();
                        if (rdr.Read())
                        {
                            PaymentsInAYear = Convert.ToDouble(rdr["PaymentsInAYear"].ToString());
                        }
                    }
                  
                  Verify:
                    cnn3.Open();
                    cmd= new OdbcCommand ("SELECT COUNT(*) FROM ODASPInstallment WHERE ODASPInstallment.PaymentMode = '" + reader["PaymentMode"].ToString ()+ "' and ODASPInstallment.LeasePeriod = '1'",cnn3);
                    String count = cmd.ExecuteScalar().ToString();
                   
                    cnn3.Close();
                    cnn3.Open();
                    cmd = new OdbcCommand("SELECT * FROM ODASPInstallment WHERE ODASPInstallment.PaymentMode = '" + reader["PaymentMode"].ToString() + "' and ODASPInstallment.LeasePeriod = '1'",cnn3);
                    rdr1 = cmd.ExecuteReader();
                    if (Convert .ToInt32 ( count ) == 0) {
                        cnn2.Close();
                        cnn2.Open();
                        cmd = new OdbcCommand("SELECT * FROM ODASPPaymentMode WHERE PaymentMode = '" + reader["PaymentMode"].ToString() + "'",cnn2);
                        rdr = cmd.ExecuteReader();
                        if(rdr.Read ()){
                       int PInts=1;
                       int Dur=0;
                       while (PInts != (Convert.ToInt32(rdr["CoverPeriod"].ToString()) + 1))
                       {
                           cnn4.Open();
                            cmd = new OdbcCommand("INSERT INTO ODASPInstallment(LeasePeriod,PaymentMode,Installment,Duration,InstallmentDescription,dateprepared,Preparedby) VALUES('"+1+
                                "','" + reader["PaymentMode"].ToString() + "','" + PInts + "','" + Dur + "','Installment'" + PInts + ",'" + DateTime.Today + "','" + CurrentUserName + "')", cnn4);
                            cmd.ExecuteNonQuery();
                           PInts = PInts + 1;
                            Dur = Dur + Convert.ToInt32(rdr["CoverPeriod"].ToString());
                            cnn4.Close();
                       }
                       goto Verify;
                        }
                    }
                        Double installmentsAmount, LeaseP, Rent;
                        LeaseP = 1;
                      
                            Rent = Convert.ToDouble(reader["AnnualRent"].ToString());
                         
                        DateTime dt = Convert.ToDateTime(reader["CommencementDate"]);
                       
                        OdbcDataReader rsSave;
                        while (LeaseP < (Convert .ToInt32 ( reader["LeaseDuration"].ToString())+1))
                        {
                            installmentsAmount = Rent / PaymentsInAYear;
                            while (rdr1.Read())
                            {
                                cnn5.Open();
                                cmd = new OdbcCommand("SELECT * FROM ODASMInstallment WHERE ContractYear = '" + LeaseP + "' and ContractNo = '" + reader["ContractNo"].ToString() + "' and Installment = '" + rdr1["Installment"].ToString() + "' and PaymentMode = '" + reader["PaymentMode"].ToString() + "'", cnn5);
                                rsSave = cmd.ExecuteReader();
                                if (rsSave.Read())
                                {
                                  if(CON.State ==ConnectionState .Closed){
                                      CON.Open();
                                  }
                                   // CON.Open();
                                  cmd = new OdbcCommand("UPDATE ODASMInstallment SET ContractNo='" + ContracttNo() +
                                      "',AccountNo'" + txtLandLordNo.Text +
                                      "',Installment='" + rdr["Installment"].ToString() + "',PaymentMode='" + txtpaymentMode.Text +
                                      "',InvoiceNo='" + generateInstallmentNo ()+ "',ContractYear='"+LeaseP +"',TotalRent='"+Rent +
                                      "',CurrentPeriod='" + getPeriod() + "',PaymentDueDate'" + dt.AddMonths(Convert.ToInt32(rdr1["Duration"].ToString())) +
                                      "',InstallmentPercent='" + rdr["InstallmentPercent"].ToString() + "',PaymentDue='" + installmentsAmount +
                                      "',Balance'" + installmentsAmount + "',PaymentFlag='N' WHERE ContractYear = '" + LeaseP + "' and ContractNo = '" + reader["ContractNo"].ToString() + "' and Installment = '" + rdr["Installment"].ToString() + "' and PaymentMode = '" + reader["PaymentMode"].ToString() + "'", CON);
                                    cmd.ExecuteNonQuery();
                                    CON.Close();
                                }
                                else
                                {
                                    if (CON.State == ConnectionState.Closed)
                                    {
                                        CON.Open();
                                    }
                                   
                                    cmd = new OdbcCommand("INSERT INTO ODASMInstallment(PlotNo,ContractNo,AccountNo,Installment,PaymentMode,InvoiceNo,ContractYear,TotalRent,CurrentPeriod,PaymentDueDate,InstallmentPercent,PaymentDue,Balance,PaymentFlag)"+
                                        " VALUES('"+txtPlotNo .Text +"','" + ContracttNo() + "','" + txtLandLordNo.Text + "','" + rdr1["Installment"].ToString() + "','"+txtpaymentMode .Text +"','" + generateInstallmentNo() + 
                                        "','" + LeaseP + "','" + Rent + "','" + getPeriod() + 
                                        "','" + dt.AddMonths(Convert.ToInt32(rdr1["Duration"].ToString())) +
                                        "','" + rdr1["InstallmentPercent"].ToString() + "','" + installmentsAmount + "','" + installmentsAmount + "','N')", CON);
                                    cmd.ExecuteNonQuery();
                                    CON.Close();
                                }
                                cnn5.Close();
                            }
                                LeaseP = LeaseP + 1;
                                dt = dt.AddYears(1);
                            if(optPercentage .Checked ==true ){
                                Rent = Rent * (100 + Convert.ToInt32(reader["PercentageIncreament"])) / 100;
                            }
                            else if (optAmount.Checked == true)
                            {
                                Rent = Rent + Convert.ToInt32(reader["PercentageIncreament"].ToString());
                            }
                            else {
                                Rent = Rent * (100 + Convert.ToInt32(reader["PercentageIncreament"])) / 100;
                            }
                            
                        
                    }

                }
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
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
        public String ContracttNo()
        {

            String Contratc = "";
            try
            {
                GeneralVariables vars = new GeneralVariables();
                OdbcConnection con = new OdbcConnection(vars.SQLstr);
                con.Open();

                cmd = new OdbcCommand("SELECT * FROM ODASMLeaseagreement WHERE ContractNo = '" + txtContractNo.Text + "' and AccountNo = '" + txtLandLordNo.Text + "'", con);
                reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        Contratc = reader["ContractNo"].ToString(); 
                      }
                    con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return Contratc;
        }
        public String generateInstallmentNo() {
           
             String InstallmentNo="";
            try {
                GeneralVariables vars = new GeneralVariables();
                OdbcConnection con = new OdbcConnection(vars.SQLstr );
                OdbcConnection con2 = new OdbcConnection(vars.SQLstr);
                OdbcConnection con3 = new OdbcConnection(vars.SQLstr);
               
                con2.Open();
                con3.Open();
                OdbcDataReader ODBR;
                con.Open();
                int count;
                int count2;
                cmd = new OdbcCommand("SELECT COUNT(*) FROM ODASMLeaseagreement WHERE ContractNo = '" +txtContractNo.Text + "' and AccountNo = '" +txtLandLordNo.Text + "'",con);
                count =Convert .ToInt32 ( cmd.ExecuteScalar().ToString ());
                con.Close();
                con.Open();
               
                cmd = new OdbcCommand("SELECT * FROM ODASMLeaseagreement WHERE ContractNo = '" + txtContractNo.Text + "' and AccountNo = '" + txtLandLordNo.Text + "'", con);
                reader = cmd.ExecuteReader();
              if(count !=0){
                  if (reader.Read())
                  {
                     

                      cmd = new OdbcCommand("SELECT COUNT(*)FROM ODASMInstallment WHERE ContractNo = '" + txtContractNo.Text  + "'", con3);
                      count2 = Convert.ToInt32(cmd.ExecuteScalar().ToString());
                      con3.Close();

                    
                      cmd = new OdbcCommand("SELECT * FROM ODASMInstallment WHERE ContractNo = '" + txtContractNo.Text  + "'", con2);
                      ODBR = cmd.ExecuteReader();
                      if (ODBR.Read())
                      {
                          InstallmentNo = ODBR["ContractNo"].ToString() + "-" + count2 + 1;
                      }
                      else
                      {
                          InstallmentNo = txtContractNo.Text  + "-" + 1;

                      }

                  }
                
              }
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }

            return InstallmentNo;
        }
        private void showALLINSTALLMENTSDUE() {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn2.Open();
                cnn.Open();
                string sql = "SELECT * FROM ODASMInstallment I Where I.PlotNo='" +txtPlotNo.Text + "' AND I.ContractNo = '" +txtContractNo.Text + "' Order by InstallmentNo ";

                cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                ListALLInstallments.Items.Clear();
                ListALLInstallments.Columns.Clear();
                ListALLInstallments.Columns.Add("Installment No", ListALLInstallments.Width / 8);
                ListALLInstallments.Columns.Add("Rent", ListALLInstallments.Width / 8);
                ListALLInstallments.Columns.Add("Payment Date", ListALLInstallments.Width / 8);
                ListALLInstallments.Columns.Add("Flag", ListALLInstallments.Width / 8);
                ListALLInstallments.Columns.Add("Current Year", ListALLInstallments.Width / 8);
                ListALLInstallments.Columns.Add("Payment Period", ListALLInstallments.Width / 8);
                ListALLInstallments.Columns.Add("Payment Mode", ListALLInstallments.Width / 8);
                ListALLInstallments.Columns.Add("Inv No", ListALLInstallments.Width / 8);
                ListALLInstallments.Columns.Add("Payment Due", ListALLInstallments.Width / 8);


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
                        } if (reader["PaymentFlag"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PaymentFlag"].ToString());
                        } if (reader["ContractYear"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["ContractYear"].ToString());
                        } if (reader["CurrentPeriod"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["CurrentPeriod"].ToString());
                        } if (reader["PaymentMode"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PaymentMode"].ToString());
                        } if (reader["InvoiceNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["InvoiceNo"].ToString());
                        } if (reader["PaymentDue"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PaymentDue"].ToString());
                        }


                        ListALLInstallments.Items.Add(lv3);





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
        private void frmLease_Load(object sender, EventArgs e)
        {
            GeneralVariables vrs = new GeneralVariables();
           
            loadPaymentModes();
            getLANDLORDS();
            setALLAcquiredSites();
            if (vrs.bPlotRenewal == true)
            {
                getMastsToLease();
            }
            else {
                getLeasableMasts();
            }
            showALLLandLORDSites();
        }

        private void listView3_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            ListView.ListViewItemCollection items = listView3.Items;
            foreach (ListViewItem item in items  ){
            if(item .Checked ==true ){
               txtLandLordNo .Text = item.Text;
                txtNames .Text =item .SubItems [1].Text ;
               showALLLandLORDSites();
               
            }
            }
        }

        private void ListProperties_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            ListView.ListViewItemCollection items = ListProperties.Items;
            foreach (ListViewItem item in items)
            {
                
                if (item.Checked == true)
                {
                   
                    item.Checked = true;
                    txtPlotNo.Text = item.Text;
                   
                    showALLLandLORDSites();
                    
                }
            }
        }

        private void ListProperties_MouseClick(object sender, MouseEventArgs e)
        {
          setALLAcquiredSites();
        }

        private void listView3_MouseClick(object sender, MouseEventArgs e)
        {
            getLANDLORDS();
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
        private void AnableeChilds(Control ctrl)
        {
            foreach (Control c in ctrl.Controls)
            {
                AnableeChilds(c);
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
        private void btnSave_Click(object sender, EventArgs e)
        {
            
            try {
                GeneralVariables vars = new GeneralVariables();
             

                if(txtLandLordNo .Text ==""){
                    MessageBox.Show("Land Lord is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            
                }else if(txtpaymentMode .Text ==""){
                    MessageBox.Show("Payment Mode is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            
                }
                else if (textBox1.Text == "")
                {
                    MessageBox.Show("Please select one or more BillBoards to Lease!", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    saveRecord();
                    if (textBox1.Text == "0")
                    {
                        return;
                    }
                    else
                        updateSITE();
                    updateLeasedPlotMasts();
                    upDateLeaseAnnualRent();
                   
                    updateInstallments();
                    disableALLRECORD();
                    getLANDLORDS();
                    showALLLandLORDSites();
                    showALLINSTALLMENTSDUE();
                   
                }
             
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
        }

        private void listView2_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            int K, J;
            K = 0;
            J = listView2.Items.Count;
            GeneralVariables vrs = new GeneralVariables();
            ListView.ListViewItemCollection items = listView2.Items;
            ListView.CheckedListViewItemCollection checkedItems = listView2 . CheckedItems;
            for (int i = 0; i <= J; i++)
            {
                foreach (ListViewItem item in checkedItems)
                {
                    if (item.Checked == true)
                    {
                        textBox1.Text = checkedItems.Count.ToString ();
                        mastNo.Text = item.Text;
                        if (vrs.bPlotRenewal == true)
                        {
                            getMastsToLease();
                        }
                        else
                        {
                            getLeasableMasts();
                        }
                    }
                }
            }
        }

        private void cmdPrint_Click(object sender, EventArgs e)
        {
            GeneralVariables vars = new GeneralVariables();
            vars.rptODASRentPaymentInstallment.currentRecord =txtContractNo.Text;
            vars.rptODASRentPaymentInstallment.ShowDialog();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            GeneralVariables vars = new GeneralVariables();
           
            vars.rptContractAgreement.CurrentRecord = txtContractNo.Text;
          
            vars.rptContractAgreement.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AnableeChilds(this );
        }
    }
}
