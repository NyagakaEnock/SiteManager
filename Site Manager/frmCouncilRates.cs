using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Site_Manager
{
    public partial class frmCouncilRates : Form
    {
        public frmCouncilRates()
        {
            InitializeComponent();
        }
        OdbcCommand cmd;
        OdbcDataReader reader;
        DateTime dtLastDateOfYear;
        string sql;
        string BillBoard;
        string Face;
        string SiteNo;
        String CurrentUserName;
        private void textBox18_TextChanged(object sender, EventArgs e)
        {

        }
        private void ShowSITESWITHRATES()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
               
                cnn.Open();

                string sql = "SELECT *  FROM ODASMCouncilRateDue R, ODASPPlotSite S where S.SiteNo = R.SiteNo and (R.AmountDue > 0) ";

                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView4.Items.Clear();
                listView4.Columns.Clear();
                listView4.Columns.Add("Site No", listView4.Width / 4);
                listView4.Columns.Add("Media code", listView4.Width / 4);
                listView4.Columns.Add("Media Size", listView4.Width / 4);
                listView4.Columns.Add("Payment Modes", listView4.Width / 4);
                listView4.Columns.Add("Start Date", listView4.Width / 4);
                listView4.Columns.Add("End Date", listView4.Width / 4);
                listView4.Columns.Add("Amount", listView4.Width / 4);
              
                if (reader.HasRows)
                {

                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["SiteNo"].ToString());

                        if (reader["MediaCode"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["MediaCode"].ToString());
                        } if (reader["MediaSize"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["MediaSize"].ToString());
                        } if (reader["PaymentMode"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PaymentMode"].ToString());
                        } if (reader["StartDate"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["StartDate"].ToString());
                        } if (reader["EndDate"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["EndDate"].ToString());
                        } if (reader["AmountDue"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["AmountDue"].ToString());
                        }




                        listView4.Items.Add(lv3);

                       

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
        private void ListALLSITES()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();

                string sql = "SELECT * FROM ODASPPlotSite, ODASPPlot Where  ODASPPlotSite.PlotNo = ODASPPlot.PlotNo and ODASPPlot.CouncilCode = '" +txtTownCode.Text + "' and ODASPPlot.OnRoadReserve='Y' ";

                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView3.Items.Clear();
                listView3.Columns.Clear();
                listView3.Columns.Add("Site No", listView4.Width / 4);
                listView3.Columns.Add("B.Board No", listView4.Width / 4);
                listView3.Columns.Add("Plot", listView4.Width / 4);
                listView3.Columns.Add("Site", listView4.Width / 4);
              
                if (reader.HasRows)
                {

                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["SiteNo"].ToString());

                        if (reader["MastNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["MastNo"].ToString());
                        } if (reader["PlotName"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PlotName"].ToString());
                        } if (reader["SiteDetails"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["SiteDetails"].ToString());
                        } 



                        listView3.Items.Add(lv3);



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
        private void AnableAll(Control ctrl)
        {
            foreach (Control c in ctrl.Controls)
            {
                AnableAll(c);
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
        private void ListALLCOUNCILACCOUNTS()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();

                string sql = "SELECT * FROM ODASPAccount A, ODASPAccountType T Where A.AccountType = T.AccountType and A.status = 'A' and T.Council = 'Y'";

                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView2.Items.Clear();
                listView2.Columns.Clear();
                listView2.Columns.Add("Account No", listView4.Width / 4);
                listView2.Columns.Add("Company Name", listView4.Width / 4);
                listView2.Columns.Add("Town", listView4.Width / 4);
                listView2.Columns.Add("Type", listView4.Width / 4);

                if (reader.HasRows)
                {

                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["AccountNo"].ToString());

                        if (reader["CompanyName"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["CompanyName"].ToString());
                        } if (reader["Towncity"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["Towncity"].ToString());
                        } if (reader["AccountType"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["AccountType"].ToString());
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
        private void loadRECORD()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();

                string sql = "SELECT P.*,(PS.MediaSize) as Media,PS.*, PM.*,(PM.CommencementDate)as ComDate,(PM.ExpiryDate) as ExDate FROM ODASPPlot P,ODASPPlotmast PM, ODASPPlotSite PS WHERE PS.SiteNo = '" +txtSiteNo.Text + "' and PS.MastNo = PM.MastNo and PS.PlotNo = P.PlotNo";

                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();


                if (reader.Read())
                {
                    txtPlotNo.Text = reader["PlotNo"].ToString();
                    txtSiteDetails.Text = reader["SiteDetails"].ToString();
                    txtMediaCode.Text = reader["MediaCode"].ToString();
                    txtMediaSize.Text = reader["Media"].ToString();
                    txtCommencementDate.Text = reader["ComDate"].ToString();
                    txtExpiryDate.Text = reader["ExDate"].ToString();
                    txtAccountNo.Text = reader["CouncilAccountNo"].ToString();
         
                }
                reader.Close();

                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void updateSITE()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcDataReader odbr;
                cnn2.Open();
                cnn.Open();
                if (txtMediaCode.Text == "BIL")
                {
                    string sql = "Select * From ODASPPlotSite where MastNo = '" + txtMast.Text + "'";

                    OdbcCommand cmd = new OdbcCommand(sql, cnn);
                    reader = cmd.ExecuteReader();


                    while (reader.Read())
                    {
                        cmd = new OdbcCommand("Select * from ODASPPlotSite Where siteNo = '" + reader["SiteNo"].ToString() + "'", cnn2);
                        odbr = cmd.ExecuteReader();
                        if (odbr.Read())
                        {
                            cnn2.Close();
                            cnn2.Open();
                            cmd = new OdbcCommand(" UPDATE ODASPPlotSite SET RateStatus='" + txtStatus.Text + "' ,RateDue='" + Convert.ToDouble(txtAmount.Text) + "',RateDueDate='" + Convert.ToDouble(txtRateDueDate.Text) + "'Where siteNo = '" + reader["SiteNo"].ToString() + "'");
                            cmd.ExecuteNonQuery();
                            cnn2.Close();
                        }
                    }
                    reader.Close();

                    cnn.Close();

                }
                else {


                    string sql = "Select * from ODASPPlotSite Where siteNo = '" +txtSiteNo.Text + "' ";

                    OdbcCommand cmd = new OdbcCommand(sql, cnn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        cnn.Close();
                        cnn.Open();
                        cmd = new OdbcCommand("UPDATE ODASPPlotSite SET RateStatus='" + txtStatus.Text + "',RateDue='" + Convert.ToDouble(txtAmount.Text) + "',RateDueDate='"+Convert .ToDouble (txtRateDueDate .Text )+"' Where siteNo = '" + txtSiteNo.Text + "'", cnn);
                        cnn.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }
        private void UpdateJobBriefItems()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();

                string sql = "Select * From ODASMJObBriefItems where JobBriefItemNo = '" +txtJobBriefItemNo.Text + "'";

                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();


                if (reader.Read())
                {
                    cnn.Close();
                    cnn.Open();
                    cmd = new OdbcCommand("UPDATE ODASMJObBriefItems SET RatesComputed='Y' where JobBriefItemNo = '" + txtJobBriefItemNo.Text + "'", cnn);
                    cmd.ExecuteNonQuery();
                }
                reader.Close();

                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }
        private void updateRATESCHEDULE()
        {
          
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();
                if (GeneralVariables.bBillBoard == true || GeneralVariables.bStreetSign ==true )
                {
                    SiteNo = txtMast.Text;

                 sql = "Select * from ODASMCouncilRateDue Where JobBriefItemNo= '" +txtJobBriefItemNo.Text + "' and SiteNo = '" +txtMast.Text + "' and StartDate = '" +Convert .ToDateTime ( txtRateStartDate.Text).ToString ("MMMM dd,YYYY") + "'";
                }else {
                    SiteNo = txtSiteNo.Text;
                sql ="Select * from ODASMCouncilRateDue Where JobBriefItemNo= '" +txtJobBriefItemNo.Text +"' and SiteNo = '" +txtSiteNo.Text + "' and StartDate = '"  +Convert .ToDateTime ( txtRateStartDate.Text).ToString ("MMMM dd,YYYY") + "'";
                }

                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();


                if (reader.Read())
                {
                    if (GeneralVariables.bBillBoard == true || GeneralVariables.bStreetSign == true)
                    {
                        BillBoard = "Y";
                        Face = "N";
                    }
                    else
                    {
                        BillBoard = "N";
                        Face = "Y";
                    }
                    txtReferenceNo.Text = reader["ReferenceNo"].ToString();
                    cnn.Close();
                    cnn.Open();
                    cmd = new OdbcCommand("UPDATE ODASMCouncilRateDue SET BillBoard='" + BillBoard + "',Face='" + Face + "',EndDate='" + Convert.ToDateTime(txtRateExpiryDate.Text).ToString("MMMM dd,yyyy") +
                        "',AmountDue='" + Convert.ToDouble(txtAmount.Text) + "',DueDate='" + txtRateDueDate.Text +
                        "',PaymentMode='" + cboPaymentMode.Text + "',Duration='" + txtDuration.Text +
                        "',Status='RATES-PREPARED',Balance='" + Convert.ToDouble(txtAmount.Text) +
                        "' Where JobBriefItemNo= '" + txtJobBriefItemNo.Text + "' and SiteNo = '" + SiteNo + "' and StartDate = '" + Convert.ToDateTime(txtRateStartDate.Text).ToString("MMMM dd,YYYY") + "'", cnn);
                    cmd.ExecuteNonQuery();
                }
                else {

                    if (GeneralVariables.bBillBoard == true || GeneralVariables.bStreetSign == true)
                    {
                        SiteNo = txtMast.Text;
                        BillBoard = "Y";
                        Face = "N";
                    }
                    else
                    {
                        SiteNo = txtSiteNo.Text;
                        BillBoard = "N";
                        Face = "Y";
                    }
                    cnn.Close();
                    cnn.Open();
                    cmd = new OdbcCommand("INSERT INTO ODASMCouncilRateDue(SiteNo,BillBoard,Face,JobBriefItemNo,StartDate,dateprepared,Preparedby,CurrentYear,paid,Requisitioned,EndDate,AmountDue,DueDate,PaymentMode,Duration,Status,Balance)VALUES('"+SiteNo +
                        "','" + BillBoard + "','" + Face + "','" + txtJobBriefItemNo.Text + 
                        "','" + Convert.ToDateTime(txtRateStartDate .Text ).ToString("MMMM dd,YYYY") +
                        "','" + DateTime.Today + "','" + CurrentUserName + "','" + txtCurrentYear.Text + "','N','N','" + Convert.ToDateTime(txtRateExpiryDate.Text).ToString("MMMM dd,YYYY") + 
                        "','"+Convert .ToDouble (txtAmount .Text )+"','"+cboPaymentMode .Text +"','"+Convert .ToInt32 ( txtDuration .Text) +
                        "','RATES-PREPARED','"+Convert .ToDouble (txtAmount .Text )+"')", cnn);
                    cmd.ExecuteNonQuery();
                }
                reader.Close();

                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void calcDUEDATE()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();

                string sql = "Select * from ODASPDefault";

                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();


                if (reader.Read())
                {

                    txtRateDueDate.Text = Convert.ToDateTime(txtRateStartDate.Text).AddDays(-1 * Convert.ToInt32(reader["DefaultRateDays"].ToString())).ToString();
                }
                reader.Close();

                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }
        private void saveRecord()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();

                string sql = "Select * From ODASPYear Where CurrentYear = '" + DateTime .Today .Year + "'";

                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();


                if (reader.Read())
                {
                    dtLastDateOfYear = Convert.ToDateTime(reader["EndDate"].ToString());
                }

                while (Convert .ToDateTime (txtRateExpiryDate .Text )<=dtLastDateOfYear && Convert .ToDateTime (txtRateStartDate .Text )<=dtLastDateOfYear){
                updateRATESCHEDULE();
                UpdateJobBriefItems();
                       txtRateStartDate.Text = Convert.ToDateTime(txtRateExpiryDate .Text ).AddDays (1).ToString ();
                    txtRateExpiryDate.Text = Convert.ToDateTime(txtRateStartDate.Text).AddMonths(Convert .ToInt32  (txtDuration .Text )).ToString ();
                if(Convert .ToDateTime (txtRateExpiryDate )>dtLastDateOfYear ){
                    txtRateExpiryDate.Text = dtLastDateOfYear.ToString();
                
                }
                txtReferenceNo.Text = "";
                calcDUEDATE();
                }
               
                reader.Close();

                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }
        private void disableAllRecord(Control ctrl)
        {
            foreach (Control c in ctrl.Controls)
            {
                disableAllRecord(c);
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
        private void frmCouncilRates_Load(object sender, EventArgs e)
        {
            disableAllRecord(this);
            ShowSITESWITHRATES();
            ListALLSITES();
            ListALLCOUNCILACCOUNTS();
        }

        private void listView3_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            ListView.CheckedListViewItemCollection checkeditems = listView3.CheckedItems;
           foreach (ListViewItem items in checkeditems ){
               txtSiteNo.Text = items.Text;
               txtPlotName.Text =items .SubItems[2].Text ;
               txtMast.Text =items .SubItems [1].Text ;
               loadRECORD();
           }
        }

        private void listView3_MouseClick(object sender, MouseEventArgs e)
        {
            ListALLSITES();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AnableAll(this);
        }
        private Boolean ValidateRECORD()
        {
            try
            {
                if (txtAmount.Text == "")
                {
                    MessageBox.Show("The Council Rate Amount is required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtAmount.Focus();
                    return false;
                }
                else if (Convert.ToInt32(txtAmount.Text) <= 0)
                {
                    MessageBox.Show("The Council Rate Amount MUST be Greater Than Zero", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtAmount.Focus();
                    return false;
                }
                else if (txtCurrentYear.Text == "")
                {
                    MessageBox.Show("The Current Year MUST NOT be Blank", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtCurrentYear.Focus();
                    return false;
                }
                else if (txtAccountNo.Text == "")
                {
                    MessageBox.Show("The Account no entered is Invalid", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtAccountNo.Focus();
                    return false;
                }
                else if (txtJobBriefItemNo.Text == "")
                {
                    MessageBox.Show("A Job Brief Item Number is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtJobBriefItemNo.Focus();
                    return false;
                }
                else if (txtDuration.Text == "")
                {
                    MessageBox.Show("The Duration is required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtDuration.Focus();
                    return false;
                }
                else if (Convert.ToInt32(txtDuration.Text) <= 0)
                {
                    MessageBox.Show("The Duration MUST be Greater Than Zero", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtDuration.Focus();
                    return false;
                }
                else if (txtMediaCode.Text == "")
                {
                    MessageBox.Show("The Media Is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtMediaCode.Focus();
                    return false;
                }
                else if (txtMediaSize.Text == "")
                {
                    MessageBox.Show("The Size of the Media is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtMediaSize.Focus();
                    return false;
                }
                else if (txtPlotNo.Text == "")
                {
                    MessageBox.Show("The Plot No is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtPlotNo.Focus();
                    return false;
                }
                else if (txtSiteNo.Text == "")
                {
                    MessageBox.Show("The Site No is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtSiteNo.Focus();
                    return false;
                }
                else if (txtRateDueDate.Text == "")
                {
                    MessageBox.Show("The Date the Council Rate is Due is Manadatory", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtRateDueDate.Focus();
                    return false;
                }
                else if (txtRateExpiryDate.Text == "")
                {
                    MessageBox.Show("The Exppiry Date for the Council Rate is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtRateExpiryDate.Focus();
                    return false;
                }
                else if (txtRateStartDate.Text == "")
                {
                    MessageBox.Show("The Start Date for the Council Rate is Required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtRateStartDate.Focus();
                    return false;
                }
                else if (txtRateDueDate.Text == "" || txtRateStartDate.Text == "")
                {
                    MessageBox.Show("The Due date and Start Date is required", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtRateStartDate.Focus();
                    return false;
                }
                else if (Convert.ToDouble(txtRateDueDate.Text) < Convert.ToDouble(txtRateStartDate.Text))
                {
                    MessageBox.Show("The Due date MUST come before the Start Date for the Rate", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtRateStartDate.Focus();
                    return false;
                }
                else
                {
                    return true;
                }
            }catch ( Exception ex){
                
                MessageBox.Show(ex.ToString ());
                return false;
            }
           
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (ValidateRECORD())
            {
            //    saveRecord();
            //    updateSITE();
            //    ShowSITESWITHRATES();
            //    ListALLSITES();
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            GeneralVariables vars = new GeneralVariables();
            vars.rptCouncilRates.ShowDialog();
        }
    }
}
