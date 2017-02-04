using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Odbc ;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Security.Principal;
namespace Site_Manager
{

    public partial class frmMain : Form 
    {
        OdbcCommand  cmd;
        OdbcDataReader reader;
        public OdbcConnection conn;
        String conSTR;
        public String CurrentUserName;
        public TextBox myTextbox;
     
       public String  currentRecord;
        public frmMain()
        {
            InitializeComponent();
          
        }
        private void showALLSitesWithoutProperties() {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                cnn2.Open();
                cnn.Open();
                cmd = new OdbcCommand ("SELECT COUNT(*)  FROM ODASPPlot,ODASPPlotMast where ODASPPlotMast.PlotNo = ODASPPlot.PlotNo and (ODASPPlotMast.PropertiesAssigned='N' or ODASPPlotMast.PropertiesAssigned is null)",cnn2 );
                String c;
                c = cmd.ExecuteScalar().ToString ();
                progressBar1.Minimum = 0;
              
                progressBar1.Maximum =Convert .ToInt32 ( c);
                progressBar1.Visible=true;
                string sql = "SELECT *  FROM ODASPPlot,ODASPPlotMast where ODASPPlotMast.PlotNo = ODASPPlot.PlotNo and (ODASPPlotMast.PropertiesAssigned='N' or ODASPPlotMast.PropertiesAssigned is null)";

                 cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();
                listView1.Columns.Add("BBoard No", listView1.Width / 9);
                listView1.Columns.Add("Faces", listView1.Width / 15);
                listView1.Columns.Add("Plot No", listView1.Width / 9);
                listView1.Columns.Add("Plot Name", listView1.Width / 4);
                listView1.Columns.Add("Physcical Location", listView1.Width / 3);
                listView1.Columns.Add("Structure", listView1.Width / 7);

                if (reader.HasRows)
                {
                  
                      
                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["MastNo"].ToString());
                        if (reader["NoofSites"].ToString()!="")
                        {
                        lv3.SubItems.Add(reader["NoofSites"].ToString());
                        }
                        if (reader["PlotNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PlotNo"].ToString());
                        }
                        if (reader["PlotName"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PlotName"].ToString());
                        }
                        if (reader["PhysicalLocation"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PhysicalLocation"].ToString());
                        }
                        if (reader["TypeOfMast"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["TypeOfMast"].ToString());
                        }

                        listView1.Items.Add(lv3);

                        progressBar1.Value = progressBar1.Value + 1;
                       
                      

                    }
                    progressBar1.Visible = false;
                    progressBar1.Value = 0;
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
                cmd = new OdbcCommand("SELECT COUNT(*)  FROM ODASPPlot  where (OnRoadReserve = 'N' or OnRoadReserve is null) ", cnn2);
                String c;
                c = cmd.ExecuteScalar().ToString();
                progressBar1.Minimum = 0;

                progressBar1.Maximum = Convert.ToInt32(c);
                progressBar1.Visible = true;
                string sql = "SELECT *  FROM ODASPPlot  where (OnRoadReserve = 'N' or OnRoadReserve is null) ";

                cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();
                listView1.Columns.Add("Plot No", listView1.Width / 3);
                listView1.Columns.Add("Plot Name", listView1.Width / 3);
                listView1.Columns.Add("Physical Location", listView1.Width / 3);
              

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
                       

                        listView1.Items.Add(lv3);

                        progressBar1.Value = progressBar1.Value + 1;



                    }
                    progressBar1.Visible = false;
                    progressBar1.Value = 0;
                }
                reader.Close();
                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void SearchshowALLSitesWithoutProperties()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
          
                string sql = "SELECT *  FROM ODASPPlot,ODASPPlotMast where ODASPPlotMast.PlotNo = ODASPPlot.PlotNo and (ODASPPlotMast.PropertiesAssigned='N' or ODASPPlotMast.PropertiesAssigned is null AND (ODASPPlot.PlotNo LIKE '%" + textBox1.Text + "%' OR ODASPPlot.PlotName LIKE '%" + textBox1.Text + "%'))";

                 cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();
                listView1.Columns.Add("BBoard No", listView1.Width / 9);
                listView1.Columns.Add("Faces", listView1.Width / 15);
                listView1.Columns.Add("Plot No", listView1.Width / 9);
                listView1.Columns.Add("Plot Name", listView1.Width / 4);
                listView1.Columns.Add("Physcical Location", listView1.Width / 3);
                listView1.Columns.Add("Structure", listView1.Width / 7);

                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["MastNo"].ToString());
                        if (reader["NoofSites"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["NoofSites"].ToString());
                        }
                        if (reader["PlotNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PlotNo"].ToString());
                        }
                        if (reader["PlotName"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PlotName"].ToString());
                        }
                        if (reader["PhysicalLocation"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PhysicalLocation"].ToString());
                        }
                        if (reader["TypeOfMast"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["TypeOfMast"].ToString());
                        }

                        listView1.Items.Add(lv3);




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
        public  void showALLAvailableFaces()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcConnection cnn2 = new OdbcConnection(GeneralVariables.SQLstr);
                
                OdbcCommand cmd1;
                cnn.Open();
                cmd1 = new OdbcCommand();
                cmd1 = new OdbcCommand("Select COUNT(*) From ODASPPlot,ODASPTown,ODASPPlotSite,ODASPPlotMast Where ODASPTown.Town like '" + currentRecord + "%' and ODASPPlot.PlotNo = ODASPPlotMast.PlotNo and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo and ODASPPlotMast.ExpiryDate>'" + DateTime.Today.ToString("MMMM dd,yyyy") + "' and ODASPPlot.TownCode = ODASPTown.TownCode ", cnn);
                String c;
                c = cmd1.ExecuteScalar().ToString();
                progressBar1.Minimum = 0;

                progressBar1.Maximum = Convert.ToInt32(c);
                progressBar1.Visible = true;
                cnn.Close();
                cnn.Open();
                string sql = "Select * From ODASPPlot,ODASPTown,ODASPPlotSite,ODASPPlotMast Where ODASPTown.Town like '" + currentRecord  + "%' and ODASPPlot.PlotNo = ODASPPlotMast.PlotNo and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo and ODASPPlotMast.ExpiryDate>'" +DateTime .Today .ToString ("MMMM dd,yyyy") + "' and ODASPPlot.TownCode = ODASPTown.TownCode";
                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();

                listView1.Columns.Add("Site No", listView1.Width / 8);
                listView1.Columns.Add("BillBoard No", listView1.Width / 8);
                listView1.Columns.Add("Free From", listView1.Width / 8);
                listView1.Columns.Add("Till", listView1.Width / 8);
                listView1.Columns.Add("Free Days", listView1.Width / 8);
                listView1.Columns.Add("Town", listView1.Width / 8);
                listView1.Columns.Add("Site Details", listView1.Width / 8);
                listView1.Columns.Add("Plot Location", listView1.Width / 8);

                cnn2.Open();
                OdbcDataReader rdr;
                if (reader.HasRows)
                {


                    while (reader.Read())
                    {
                        if (cnn2.State == ConnectionState.Closed)
                        {
                            cnn2.Open();
                        }
                        cmd = new OdbcCommand("Select min(scheduleDate) as StartDate, max(scheduleDate)as EndDate from ODASMSiteSchedule Where SiteNo  = '" + reader["SiteNo"].ToString() + "' and (Reserved = 'N' or JobBriefItemNo is null) and ScheduleDate >'" + DateTime.Today.ToString("yyyy-MM-dd") + "'", cnn2);
                        rdr = cmd.ExecuteReader();
                        rdr.Read();
                        MessageBox.Show(reader["SiteNo"].ToString());
                        DateTime StartDate = Convert.ToDateTime(rdr["StartDate"].ToString());
                        DateTime EndDate = Convert.ToDateTime(rdr ["EndDate"].ToString());
                        String df = EndDate.Subtract(StartDate).ToString();

                      
                          ListViewItem lv3 = new ListViewItem(reader["SiteNo"].ToString());

                        if (reader["MastNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["MastNo"].ToString());
                        }
                        if (rdr["StartDate"].ToString() != "")
                        {
                            lv3.SubItems.Add(rdr["StartDate"].ToString());
                        } if (rdr["EndDate"].ToString() != "")
                        {
                            lv3.SubItems.Add(rdr["EndDate"].ToString());
                        }
                        lv3.SubItems.Add(df);
                        if (reader["Town"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["Town"].ToString());
                        } if (reader["SiteDetails"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["SiteDetails"].ToString());
                        } if (reader["PhysicalLocation"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PhysicalLocation"].ToString());
                        }
                        
                        
                        listView1.Items.Add(lv3);

                        progressBar1.Value = progressBar1.Value + 1;


                    }
                }
                reader.Close();
                progressBar1.Value = 0;
                progressBar1.Visible = false;
                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString ());
            }

        }
        public  void frmMain_Load(object sender, EventArgs e)
        {
            try
            {
              
                layoutSettings();
                timer1.Enabled = true;
                treeView1.ImageList = imageList1;
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message );
            }

        }

        private void loadPlots()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);

                cnn.Open();
                string sql = "SELECT * FROM ODASPTown WHERE towncode LIKE '%" + textBox1.Text + "%' OR  town LIKE '%" + textBox1.Text + "%'";
              
                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();
                listView1.Columns.Add("Town Code", listView1.Width /2);
                listView1.Columns.Add("Town", listView1.Width / 2);
                
                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["TownCode"].ToString());
                        // lv.SubItems[0].Text = reader.GetString(0).ToString();
                        lv3.SubItems.Add(reader["Town"].ToString());


                        listView1.Items.Add(lv3);




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
        private void layoutSettings() {
            panel1.Width = this.Width;
            panel2.Height = this.Height - 150;
            treeView1.Height = panel2.Height;
            panel3.Width = this.Width - panel2.Width;
            panel3.Height = this.Height -150;

            listView1.Height = panel3.Height - 5 - progressBar1.Height;
            timer1.Enabled = true;
            listView1.Width = panel3.Width - 30;
            progressBar1.Width = panel3.Width - 30;
            panel4.Width = listView1.Width / 2;
         //   progressBar1.Margin.Top = listView1.Height;
        }

 
       
        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();
            listViews displayTowns = new listViews();
            if (treeView1.SelectedNode.Name == "SiteAcquisition")
            {
                
                loadPlots();

            }
            else if (treeView1.SelectedNode.Name == "VourcherPreparation")
            {
                Cursor.Current = Cursors.WaitCursor;

               
                GeneralVariables.VourcherPrepareForm.cboPaymentCode.Text = "RENT";
                GeneralVariables.VourcherPrepareForm.CurrentUserName = CurrentUserName;
                GeneralVariables.VourcherPrepareForm.ShowDialog();
                Cursor.Current = Cursors.Default;
            }
            else if (treeView1.SelectedNode.Name == "landlord")
            {
                GeneralVariables.LandLord.ShowDialog();
            }
            else if (treeView1.SelectedNode.Name == "PrepareLease")
            {
                setALLAcquiredSites();
                
            }
            else if (treeView1.SelectedNode.Name == "AssignProperties")
            {
                showALLSitesWithoutProperties();
                
            }
            else if (treeView1.SelectedNode.Name == "PaymentConfirmation")
            {
                Cursor.Current = Cursors.WaitCursor;
                GeneralVariables.PaymentConfirmation.cboPaymentCode.Text = "RENT";
               
                GeneralVariables.PaymentConfirmation.ShowDialog();
                Cursor.Current = Cursors.Default;
            }
            else if (treeView1.SelectedNode.Name == "SetCouncilRates")
            {
                
                showALLCOUNCILS();
            }
            else if (treeView1.SelectedNode.Name == "Contaracts")
            {
                
                getALLAllocatedSites();
            }
            else if (treeView1.SelectedNode.Name == "ContractsAuthorization")
            {
                 getALLApprovedSites();
                    
            }
            else if (treeView1.SelectedNode.Name == "EditLease")
            {
                
               getCURRENTLEASES();
            }
            else if (treeView1.SelectedNode.Name == "PrintRatesSchedule")
            {
                 getALLsitesRatesPrepared();
                
            }
            else if (treeView1.SelectedNode.Name == "PrintRentInstallmentsheet")
            {
               //getALLNACADAContracts()();
               getALLContracts();
            }
            else if (treeView1.SelectedNode.Name == "EditMainContractClauses")
            {
                getALLNACADAContracts();
                // getALLContracts();
            }
            else if (treeView1.SelectedNode.Name == "PrintEditedContracts")
            {
                getALLNACADAContracts();
                // getALLContracts();
            }
            else if (treeView1.SelectedNode.Name == "PrintMainContract")
            {
                //getALLNACADAContracts()();
                getALLContracts();
            }
            else if (treeView1.SelectedNode.Name == "NACADA")
            {
                getALLNACADAContracts();
               // getALLContracts();
               // getCURRENTLEASES();
            }
            else if (treeView1.SelectedNode.Name == "PrintSchedule")
            {

                // ListALLSitesToFree();
                 getCURRENTLEASES();
            }
            else if (treeView1.SelectedNode.Name == "FreeAssignedSites")
            {

                ListALLSitesToFree();
                // AllSitesOnRoadReserve();
            }
            else if (treeView1.SelectedNode.Name == "SitesonRoadReserve")
            {

                // AllNonEagleStructures();
                AllSitesOnRoadReserve();
            }
            else if (treeView1.SelectedNode.Name == "NonCompanyStructures")
            {

                AllNonEagleStructures();
                //  AllPlotRents()();
            }
            else if (treeView1.SelectedNode.Name == "AnnualRentforAllPlots")
            {

                // RateSchedules();
                 AllPlotRents();
            }
            else if (treeView1.SelectedNode.Name == "AllExpiredSites")
            {

                RateSchedules();
                //  LeaseduetoExpire();
            }
            else if (treeView1.SelectedNode.Name == "LeaseduetoExpire")
            {

                //showALLFreeSites();
                getLEASESDUEToEXPIRE();
            }
            else if (treeView1.SelectedNode.Name == "AllFreeSites")
            {

                showALLFreeSites();
                //getALLApprovedMasts()();
            }
            else if (treeView1.SelectedNode.Name == "Load")
            {

               // showALLFreeSites();
                getALLApprovedMasts();
            }
            else if (treeView1.SelectedNode.Name == "PrepareNotice")
            {
              //  getNOTICESAPPROVED();
                getALLContracts();
            }
            else if (treeView1.SelectedNode.Name == "NoticeAuthorization")
            {
                  getNOTICESAPPROVED();
                //getALLContracts();
            }
            else if (treeView1.SelectedNode.Name == "SendNotice")
            {
                getNOTICESAUTHORIZED();
                //getALLContracts();
            }
            else if (treeView1.SelectedNode.Name == "Termination")
            {
                getCONTRACTSToTerminate();
                //getALLContracts();
            }
            else if (treeView1.SelectedNode.Name == "Renewal")
            {
                getCONTRACTSToRenew();
                //getALLContracts();
            }
            else if (treeView1.SelectedNode.Name == "RenewalContacts")
            {
                getCONTRACTSRenewed();
                //getALLContracts();
            }
            else if (treeView1.SelectedNode.Name == "NoticesPreapared")
            {
                getNoticesPrepared();
                //getALLContracts();
            }
            else if (treeView1.SelectedNode.Name == "NoticesApproved")
            {
               
                getNOTICESAPPROVED();
                //getALLContracts();
            }
            else if (treeView1.SelectedNode.Name == "NoticesAuthorised")
            {
                getALLNoticesAuthorized();

            }
            else if (treeView1.SelectedNode.Name == "NoticesSent")
            {
                getAllNoticesSent();

            }
            else if (treeView1.SelectedNode.Name == "NoticesAcknowledged")
            {
                getNoticesReceived();

            }
            else if (treeView1.SelectedNode.Name == "RenewJobs")
            {
                showNoticesAuthorized();

            }
            else if (treeView1.SelectedNode.Name == "Pinned")
            {
                getALLContracts();

            }
            else if (treeView1.SelectedNode.Name == "ShowEmptyBillBoards")
            {
                ShowAllValidEmptyBillBoards();

            }
            else if (treeView1.SelectedNode.Name == "ShowAllLandLords")
            {
                showALLLandlords();

            }
            else if (treeView1.SelectedNode.Name == "ShowSitesUn-Allocated")
            {
                showALLSitesUnAllocated();

            }
            else if (treeView1.SelectedNode.Name == "ShowSiteswithAdverts")
            {
                showALLSitesAllocated();

            }
            else if (treeView1.SelectedNode.Name == "ShowSitesReserved")
            {
                showALLSitesReserved();

            }
            else if (treeView1.SelectedNode.Name == "ShowSitestoFree")
            {
                showALLSitesToFree();

            }
            else if (treeView1.SelectedNode.Name == "SiteMaintananceSchedule")
            {
                showAllJobsCompleted();

            }
            else if (treeView1.SelectedNode.Name == "SiteDueForMaintainance")
            {
                ShowAllWorksDueForMaintenance();

            }
            else if (treeView1.SelectedNode.Name == "1Monthto")
            {
                ShowAllWorksDueForMaintenanceONEMonth();

            }
            else if (treeView1.SelectedNode.Name == "SitesScheduled")
            {
                AllSiteSchedule();

            }
            else if (treeView1.SelectedNode.Name == "NewSiteSchedule")
            {
                getALLsites();

            }
            else if (treeView1.SelectedNode.Name == "SitesUnpaid")
            {
                GeneralVariables.frmODASSearchSitesNotPaid.ShowDialog();

            }
            else if (treeView1.SelectedNode.Name == "SitesPaid")
            {
                GeneralVariables.frmODASSearchPaidSites.ShowDialog();

            } 
              
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();
            GeneralVariables.SiteAcquisition.ShowDialog();
        }

        private void clientContractAgreementFormToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();
            GeneralVariables.AgreementForm.ShowDialog();
        }

        private void siteRegistrationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            loadPlots();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            GeneralVariables GeneralVariables = new GeneralVariables();
            GeneralVariables.VourcherPrepareForm.cboPaymentCode.Text = "RENT";
            GeneralVariables.VourcherPrepareForm.CurrentUserName = CurrentUserName;
            GeneralVariables.VourcherPrepareForm.ShowDialog();
            Cursor.Current = Cursors.Default ;
         
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            //lblDate.Text = DateTime.Today.ToString("MM/dd/yyyy");
        }

        private void getALLAllocatedSites() {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcCommand cmd1;
                cnn.Open();
                cmd1 = new OdbcCommand();
                cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPPlot P, ODASMLeaseAgreement LA where P.PlotNo = LA.PlotNo  and (LA.Approved = 'N' or LA.Approved is null)", cnn);
                String c;
                c = cmd1.ExecuteScalar().ToString();
                progressBar1.Minimum = 0;

                progressBar1.Maximum = Convert.ToInt32(c);
                progressBar1.Visible = true;
                cnn.Close();
                cnn.Open();
                string sql = "SELECT * FROM ODASPPlot P, ODASMLeaseAgreement LA where P.PlotNo = LA.PlotNo  and (LA.Approved = 'N' or LA.Approved is null)";

                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();
                listView1.Columns.Add("ContractNo", listView1.Width / 5);
                listView1.Columns.Add("PlotNo", listView1.Width / 5);
                listView1.Columns.Add("PlotName", listView1.Width / 5);
                listView1.Columns.Add("Physical", listView1.Width / 5);
                listView1.Columns.Add("LandLord", listView1.Width / 5);
                

                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["ContractNo"].ToString());
                        if (reader["PlotNo"].ToString()!="")
                        {
                        lv3.SubItems.Add(reader["PlotNo"].ToString());
                        }
                        if (reader["PlotName"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PlotName"].ToString());
                        }
                        if (reader["PhysicalLocation"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PhysicalLocation"].ToString());
                        }
                        if (reader["AccountNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["AccountNo"].ToString());
                        }

                        listView1.Items.Add(lv3);

                        progressBar1.Value = progressBar1.Value + 1;


                    }
                }
                reader.Close();
                progressBar1.Value = 0;
                progressBar1.Visible = false;
                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        
        }
        private void getALLApprovedSites()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcCommand cmd1;
                cnn.Open();
                cmd1 = new OdbcCommand();
                cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPPlot P, ODASMLeaseAgreement LA where  P.PlotNo = LA.PlotNo and (LA.Authorized is null or LA.Authorized = 'N') and LA.Approved = 'Y'", cnn);
                String c;
                c = cmd1.ExecuteScalar().ToString();
                progressBar1.Minimum = 0;

                progressBar1.Maximum = Convert.ToInt32(c);
                progressBar1.Visible = true;
                cnn.Close();
                cnn.Open();
                string sql = "SELECT * FROM ODASPPlot P, ODASMLeaseAgreement LA where  P.PlotNo = LA.PlotNo and (LA.Authorized is null or LA.Authorized = 'N') and LA.Approved = 'Y'";

                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();
                listView1.Columns.Add("ContractNo", listView1.Width / 5);
                listView1.Columns.Add("PlotNo", listView1.Width / 5);
                listView1.Columns.Add("PlotName", listView1.Width / 5);
                listView1.Columns.Add("Physical", listView1.Width / 5);
                listView1.Columns.Add("LandLord", listView1.Width / 5);


                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["ContractNo"].ToString());
                        if (reader["PlotNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PlotNo"].ToString());
                        }
                        if (reader["PlotName"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PlotName"].ToString());
                        }
                        if (reader["PhysicalLocation"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PhysicalLocation"].ToString());
                        }
                        if (reader["AccountNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["AccountNo"].ToString());
                        }

                        listView1.Items.Add(lv3);

                        progressBar1.Value = progressBar1.Value + 1;


                    }
                }
                reader.Close();
                progressBar1.Value = 0;
                progressBar1.Visible = false;
                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void getCURRENTLEASES()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcCommand cmd1;
                cnn.Open();
                cmd1 = new OdbcCommand();
                cmd1 = new OdbcCommand("SELECT COUNT(*)  FROM ODASPPlot P,ODASPAccount A, ODASMLeaseAgreement LA where  LA.PlotNo = P.PlotNo AND P.AccountNo = A.AccountNo ", cnn);
                String c;
                c = cmd1.ExecuteScalar().ToString();
                progressBar1.Minimum = 0;

                progressBar1.Maximum = Convert.ToInt32(c);
                progressBar1.Visible = true;
                cnn.Close();
                cnn.Open();
                string sql = "SELECT *  FROM ODASPPlot P,ODASPAccount A, ODASMLeaseAgreement LA where  LA.PlotNo = P.PlotNo AND P.AccountNo = A.AccountNo ";

                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();
                listView1.Columns.Add("Contract No", listView1.Width / 6);
                listView1.Columns.Add("PlotNo", listView1.Width / 6);
                listView1.Columns.Add("Start Date", listView1.Width / 6);
                listView1.Columns.Add("Expiry Date", listView1.Width / 6);
                listView1.Columns.Add("Physical Location", listView1.Width / 6);
                listView1.Columns.Add("LandLord", listView1.Width / 6);


                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["ContractNo"].ToString());
                        if (reader["PlotNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PlotNo"].ToString());
                        }
                        if (reader["CommencementDate"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["CommencementDate"].ToString());
                        }
                        if (reader["expirydate"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["expirydate"].ToString());
                        }
                        if (reader["PhysicalLocation"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PhysicalLocation"].ToString());
                        }
                        if (reader["CompanyName"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["CompanyName"].ToString());
                        }

                        listView1.Items.Add(lv3);

                        progressBar1.Value = progressBar1.Value + 1;


                    }
                }
                reader.Close();
                progressBar1.Value = 0;
                progressBar1.Visible = false;
                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void getALLContracts()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcCommand cmd1;
                cnn.Open();
                cmd1 = new OdbcCommand();
                cmd1 = new OdbcCommand("SELECT COUNT(*)  FROM  ODASPPlot ", cnn);
                String c;
                c = cmd1.ExecuteScalar().ToString();
                progressBar1.Minimum = 0;

                progressBar1.Maximum = Convert.ToInt32(c);
                progressBar1.Visible = true;
                cnn.Close();
                cnn.Open();
                string sql = "SELECT (L.ExpiryDate) as EDates, L.*, P.*  FROM ODASMLeaseAgreement L, ODASPPlot P where L.Assigned = 'Y' and (L.Terminated = 'N' or L.Terminated is null) AND P.PlotNo = L.PlotNo";

                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();
                listView1.Columns.Add("Contract No", listView1.Width / 5);
                listView1.Columns.Add("PlotNo", listView1.Width / 5);
                listView1.Columns.Add("Plot Name", listView1.Width / 5);
                listView1.Columns.Add("LandLord", listView1.Width / 5);
                listView1.Columns.Add("Expiry Date", listView1.Width / 5);
            

                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["ContractNo"].ToString());
                        if (reader["PlotNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PlotNo"].ToString());
                        }
                        if (reader["PlotName"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PlotName"].ToString());
                        }
                        if (reader["AccountNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["AccountNo"].ToString());
                        }
                        if (reader["EDates"].ToString() != "")
                        {
                            lv3.SubItems.Add(Convert .ToDateTime ( reader["EDates"].ToString()).ToString ("MM/dd/yyyy"));
                        }
                       

                        listView1.Items.Add(lv3);

                        progressBar1.Value = progressBar1.Value + 1;


                    }
                }
                reader.Close();
                progressBar1.Value = 0;
                progressBar1.Visible = false;
                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void showALLCOUNCILS()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcCommand cmd1;
                cnn.Open();
                cmd1 = new OdbcCommand();
                cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPCouncil Where Status = 'A'", cnn);
                String c;
                c = cmd1.ExecuteScalar().ToString();
                progressBar1.Minimum = 0;

                progressBar1.Maximum = Convert.ToInt32(c);
                progressBar1.Visible = true;
                cnn.Close();
                cnn.Open();
                string sql = "SELECT CouncilCode, Council, UseCalendarYear, Status FROM ODASPCouncil Where Status = 'A' ";

                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();
                listView1.Columns.Add("Council Code", listView1.Width / 5);
                listView1.Columns.Add("Council", listView1.Width / 5);
                listView1.Columns.Add("Use Calendar Year?", listView1.Width / 5);
                listView1.Columns.Add("Status", listView1.Width / 5);
               

                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["CouncilCode"].ToString());
                       
                        if (reader["Council"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["Council"].ToString());
                        }
                        if (reader["UseCalendarYear"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["UseCalendarYear"].ToString());
                        }
                        if (reader["Status"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["Status"].ToString());
                        }
                       

                        listView1.Items.Add(lv3);

                        progressBar1.Value = progressBar1.Value + 1;


                    }
                }
                reader.Close();
                progressBar1.Value = 0;
                progressBar1.Visible = false;
                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void getALLNACADAContracts()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcCommand cmd1;
                cnn.Open();
                cmd1 = new OdbcCommand();
                cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPPlot", cnn);
                String c;
                c = cmd1.ExecuteScalar().ToString();
                progressBar1.Minimum = 0;

                progressBar1.Maximum = Convert.ToInt32(c);
                progressBar1.Visible = true;
                cnn.Close();
                cnn.Open();
                string sql = "SELECT (L.ExpiryDate) as EDates, L.*, P.*  FROM ODASMLeaseAgreement L, ODASPPlot P where L.Assigned = 'Y' and (L.Terminated = 'N' or L.Terminated is null) AND P.PlotNo = L.PlotNo ";

                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();
                listView1.Columns.Add("ContractNo", listView1.Width / 5);
                listView1.Columns.Add("Plot", listView1.Width / 5);
                listView1.Columns.Add("LandLord", listView1.Width / 5);
                listView1.Columns.Add("Expiry Date", listView1.Width / 5);


                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["ContractNo"].ToString());

                        if (reader["PlotNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PlotNo"].ToString());
                        }
                        if (reader["AccountNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["AccountNo"].ToString());
                        }
                        if (reader["EDates"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["EDates"].ToString());
                        }


                        listView1.Items.Add(lv3);

                        progressBar1.Value = progressBar1.Value + 1;


                    }
                }
                reader.Close();
                progressBar1.Value = 0;
                progressBar1.Visible = false;
                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void getALLsitesRatesPrepared()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcCommand cmd1;
                cnn.Open();
                cmd1 = new OdbcCommand();
                cmd1 = new OdbcCommand("SELECT COUNT(distinct CR.CurrentYear) FROM  ODASMCouncilRateDue CR", cnn);
                String c;
                c = cmd1.ExecuteScalar().ToString();
                progressBar1.Minimum = 0;

                progressBar1.Maximum = Convert.ToInt32(c);
                progressBar1.Visible = true;
                cnn.Close();
                cnn.Open();
                string sql = "SELECT  distinct CR.CurrentYear, CR.SiteNo,CR.JobBriefItemNo  FROM  ODASMCouncilRateDue CR ";

                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();
                listView1.Columns.Add("Site No", listView1.Width / 3);
                listView1.Columns.Add("Job Item No on Site", listView1.Width / 3);
                listView1.Columns.Add("Current Year", listView1.Width / 3);
     


                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["SiteNo"].ToString());

                        if (reader["CurrentYear"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["CurrentYear"].ToString());
                        }
                        if (reader["JobBriefItemNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["JobBriefItemNo"].ToString());
                        }
                    


                        listView1.Items.Add(lv3);

                        progressBar1.Value = progressBar1.Value + 1;


                    }
                }
                reader.Close();
                progressBar1.Value = 0;
                progressBar1.Visible = false;
                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
         private void ListALLSitesToFree()
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                OdbcCommand cmd1;
                cnn.Open();
                cmd1 = new OdbcCommand();
                cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPPlot,ODASPPlotSite where ODASPPlotSite.JobBriefNo is Not Null and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo", cnn);
                String c;
                c = cmd1.ExecuteScalar().ToString();
                progressBar1.Minimum = 0;

                progressBar1.Maximum = Convert.ToInt32(c);
                progressBar1.Visible = true;
                cnn.Close();
                cnn.Open();
                string sql = "SELECT * FROM ODASPPlot,ODASPPlotSite where ODASPPlotSite.JobBriefNo is Not Null and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo ";

                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();
                listView1.Items.Clear();
                listView1.Columns.Clear();
                listView1.Columns.Add("Site No", listView1.Width / 5);
                listView1.Columns.Add("Site Details", listView1.Width / 5);
                listView1.Columns.Add("Plot No", listView1.Width / 5);
                listView1.Columns.Add("Plot Name", listView1.Width / 5);
                listView1.Columns.Add("Status", listView1.Width / 5);
     


                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv3 = new ListViewItem(reader["SiteNo"].ToString());

                       // if (reader["SiteDetails"].ToString() != "")
                      //  {
                            lv3.SubItems.Add(reader["SiteDetails"].ToString());
                        //}
                        if (reader["PlotNo"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PlotNo"].ToString());
                        } if (reader["PlotName"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["PlotName"].ToString());
                        } if (reader["Status"].ToString() != "")
                        {
                            lv3.SubItems.Add(reader["Status"].ToString());
                        }
                    


                        listView1.Items.Add(lv3);

                        progressBar1.Value = progressBar1.Value + 1;


                    }
                }
                reader.Close();
                progressBar1.Value = 0;
                progressBar1.Visible = false;
                cnn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
         private void AllSitesOnRoadReserve()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPPlot,ODASPPlotMast where ODASPPlot.OnRoadReserve ='Y' and ODASPPlot.PlotNo = ODASPPlotMast.PlotNo ", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT * FROM ODASPPlot,ODASPPlotMast where ODASPPlot.OnRoadReserve ='Y' and ODASPPlot.PlotNo = ODASPPlotMast.PlotNo ";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("BillBoard No", listView1.Width / 4);
                 listView1.Columns.Add("Details", listView1.Width / 4);
                 listView1.Columns.Add("Plot No", listView1.Width / 4);
                 listView1.Columns.Add("Plot Name", listView1.Width / 4);
               


                 if (reader.HasRows)
                 {


                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["MastNo"].ToString());

                         if (reader["MastDetails"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["MastDetails"].ToString());
                         }
                         if (reader["PlotNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotNo"].ToString());
                         } if (reader["PlotName"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotName"].ToString());
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }

         }
         private void AllPlotRents()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPPlot, ODASPTown WHERE ODASPPlot.AnnualRent is not null and ODASPPlot.Towncode = ODASPTown.towncode", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT * FROM ODASPPlot, ODASPTown WHERE ODASPPlot.AnnualRent is not null and ODASPPlot.Towncode = ODASPTown.towncode";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("PlotNo", listView1.Width / 4);
                 listView1.Columns.Add("Location", listView1.Width / 4);
                 listView1.Columns.Add("Town", listView1.Width / 4);
                 listView1.Columns.Add("Rent", listView1.Width / 4);



                 if (reader.HasRows)
                 {


                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["PlotNo"].ToString());

                         if (reader["PhysicalLocation"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PhysicalLocation"].ToString());
                         }
                         if (reader["Town"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["Town"].ToString());
                         } if (reader["AnnualRent"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["AnnualRent"].ToString());
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }

         }
         private void AllNonEagleStructures()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPPlot,ODASPPlotMast where ODASPPlotMast.OwenedByClient ='Y' and ODASPPlot.PlotNo = ODASPPlotMast.PlotNo ", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT * FROM ODASPPlot,ODASPPlotMast where ODASPPlotMast.OwenedByClient ='Y' and ODASPPlot.PlotNo = ODASPPlotMast.PlotNo ";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("BillBoard No", listView1.Width / 4);
                 listView1.Columns.Add("Details", listView1.Width / 4);
                 listView1.Columns.Add("Plot No", listView1.Width / 4);
                 listView1.Columns.Add("Plot Name", listView1.Width / 4);



                 if (reader.HasRows)
                 {


                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["MastNo"].ToString());

                         if (reader["MastDetails"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["MastDetails"].ToString());
                         }
                         if (reader["PlotNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotNo"].ToString());
                         } if (reader["PlotName"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotName"].ToString());
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }

         }
         private void RateSchedules()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPPlot, ODASMJobBriefItems, ODASPPlotSite, ODASPTown, ODASMJobBrief,ODASMCouncilRatesPayable WHERE ODASPPlot.PlotNo = ODASPPlotSite.PLotNo and ODASMJobBriefItems.JobBriefItemNo = ODASPPlotSite.JobBriefItemNo and ODASMJobBriefItems.JobBriefNo = ODASMJobBrief.JobBriefNo and ODASPPlot.Towncode = ODASPTown.towncode and ODASPPlotSite.SiteNo = ODASMCouncilRatesPayable.SiteNo", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT * FROM ODASPPlot, ODASMJobBriefItems, ODASPPlotSite, ODASPTown, ODASMJobBrief,ODASMCouncilRatesPayable WHERE ODASPPlot.PlotNo = ODASPPlotSite.PLotNo and ODASMJobBriefItems.JobBriefItemNo = ODASPPlotSite.JobBriefItemNo and ODASMJobBriefItems.JobBriefNo = ODASMJobBrief.JobBriefNo and ODASPPlot.Towncode = ODASPTown.towncode and ODASPPlotSite.SiteNo = ODASMCouncilRatesPayable.SiteNo";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Site No", listView1.Width / 6);
                 listView1.Columns.Add("Location", listView1.Width / 6);
                 listView1.Columns.Add("Town", listView1.Width / 6);
                 listView1.Columns.Add("Advert", listView1.Width / 6);
                 listView1.Columns.Add("Rates Payable", listView1.Width / 6);
                 listView1.Columns.Add("Rates DueDate", listView1.Width / 6);



                 if (reader.HasRows)
                 {

                     MessageBox.Show(reader.Read().ToString());
                     while (reader.Read())
                     {
                        
                         ListViewItem lv3 = new ListViewItem(reader["SiteNo"].ToString());

                         if (reader["PhysicalLocation"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PhysicalLocation"].ToString());
                         }
                         if (reader["Town"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["Town"].ToString());
                         } if (reader["ProductCode"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["ProductCode"].ToString());
                         } if (reader["RatePayable"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["RatePayable"].ToString());
                         } if (reader["RateDueDate"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["RateDueDate"].ToString());
                         }
                         


                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }

         }
         private void getLEASESDUEToEXPIRE()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT count (*) FROM ODASMLeaseAgreement L,ODASPPlot P WHERE (L.Terminated is null or L.Terminated ='N') AND P.PlotNo = L.PlotNo and P.ExpiryDate > '" + DateTime.Today.ToString("MMMM dd,yyyy") + "'", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT * FROM ODASMLeaseAgreement L,ODASPPlot P WHERE (L.Terminated is null or L.Terminated ='N') AND P.PlotNo = L.PlotNo and P.ExpiryDate > '" +DateTime.Today .ToString ("MMMM dd,yyyy") +"'";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("PlotNo", listView1.Width / 6);
                 listView1.Columns.Add("PlotName", listView1.Width / 6);
                 listView1.Columns.Add("Physical Location", listView1.Width / 6);
                 listView1.Columns.Add("LandLord", listView1.Width / 6);
                 listView1.Columns.Add("Expiry Date", listView1.Width / 6);
                 listView1.Columns.Add("Commencent Date", listView1.Width / 6);



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
                         } if (reader["AccountNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["AccountNo"].ToString());
                         } if (reader["ContractNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["ContractNo"].ToString());
                         } if (reader["expirydate"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["expirydate"].ToString());
                         } if (reader["CommencementDate"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["CommencementDate"].ToString());
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }

         }
         private void showALLFreeSites()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPPlot,ODASPPlotSite where ODASPPlotSite.JobBriefNo is Null and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo ", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT * FROM ODASPPlot,ODASPPlotSite where ODASPPlotSite.JobBriefNo is Null and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo ";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Site No", listView1.Width / 5);
                 listView1.Columns.Add("Site Details", listView1.Width / 5);
                 listView1.Columns.Add("Plot No", listView1.Width / 5);
                 listView1.Columns.Add("Plot Name", listView1.Width / 5);
                 listView1.Columns.Add("Status", listView1.Width / 5);
               


                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["SiteNo"].ToString());

                         if (reader["SiteDetails"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["SiteDetails"].ToString());
                         }
                         if (reader["PlotNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotNo"].ToString());
                         } if (reader["PlotName"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotName"].ToString());
                         } if (reader["Status"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["Status"].ToString());
                         } 



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }

         }
         private void getALLApprovedMasts()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPPlotSite", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT * FROM ODASPPlotSite";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Site No", listView1.Width / 4);
                 listView1.Columns.Add("Plot No", listView1.Width / 4);
                 listView1.Columns.Add("Mast No", listView1.Width / 4);
                 listView1.Columns.Add("Physical Location", listView1.Width / 4);
               



                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["SiteNo"].ToString());

                         if (reader["PlotNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotNo"].ToString());
                         }
                         if (reader["MastNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["MastNo"].ToString());
                         } if (reader["SiteDetails"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["SiteDetails"].ToString());
                         } 



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }

         }
         private void getNOTICESAPPROVED()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*)  FROM ODASMLeaseAgreement where NoticeApproved = 'Y' AND (NoticeAUTHORIZED = 'N' OR NoticeAUTHORIZED is null)", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT *  FROM ODASMLeaseAgreement where NoticeApproved = 'Y' AND (NoticeAUTHORIZED = 'N' OR NoticeAUTHORIZED is null)";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Contract No", listView1.Width / 5);
                 listView1.Columns.Add("Plot", listView1.Width / 5);
                 listView1.Columns.Add("LandLord", listView1.Width / 5);
                 listView1.Columns.Add("Approved By", listView1.Width / 5);
                 listView1.Columns.Add("Approval date", listView1.Width / 5);




                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["ContractNo"].ToString());

                         if (reader["PlotNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotNo"].ToString());
                         }
                         if (reader["AccountNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["AccountNo"].ToString());
                         } if (reader["NoticeApprovedBy"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["NoticeApprovedBy"].ToString());
                         } if (reader["NoticeApprovalDate"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["NoticeApprovalDate"].ToString());
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }

         }
         private void getNOTICESAUTHORIZED()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*)  FROM ODASMLeaseAgreement where NoticeAuthorized = 'Y' and (NoticeDispatched is null or NoticeDispatched = 'N')", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT *  FROM ODASMLeaseAgreement where NoticeAuthorized = 'Y' and (NoticeDispatched is null or NoticeDispatched = 'N')";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Contract No", listView1.Width / 5);
                 listView1.Columns.Add("Plot", listView1.Width / 5);
                 listView1.Columns.Add("LandLord", listView1.Width / 5);
                 listView1.Columns.Add("Notice Date", listView1.Width / 5);
                 listView1.Columns.Add("Prepared By", listView1.Width / 5);




                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["ContractNo"].ToString());

                         if (reader["PlotNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotNo"].ToString());
                         }
                         if (reader["AccountNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["AccountNo"].ToString());
                         } if (reader["NoticeDate"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["NoticeDate"].ToString());
                         } if (reader["NoticePreparedBy"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["NoticePreparedBy"].ToString());
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }

         }
         private void getCONTRACTSToTerminate()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*)  FROM ODASMLeaseAgreement where NoticeDispatched = 'Y' and (Terminated is null or Terminated = 'N') and ReasonsForNotice = 'Termination of Contract'", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT *  FROM ODASMLeaseAgreement where NoticeDispatched = 'Y' and (Terminated is null or Terminated = 'N') and ReasonsForNotice = 'Termination of Contract'";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Contract No", listView1.Width / 5);
                 listView1.Columns.Add("Plot", listView1.Width / 5);
                 listView1.Columns.Add("LandLord", listView1.Width / 5);
                 listView1.Columns.Add("Notice Date", listView1.Width / 5);
                 listView1.Columns.Add("Prepared By", listView1.Width / 5);




                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["ContractNo"].ToString());

                         if (reader["PlotNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotNo"].ToString());
                         }
                         if (reader["AccountNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["AccountNo"].ToString());
                         } if (reader["NoticeDate"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["NoticeDate"].ToString());
                         } if (reader["NoticePreparedBy"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["NoticePreparedBy"].ToString());
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }

         }
         private void getCONTRACTSToRenew()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT( *)  FROM ODASMLeaseAgreement A,ODASPPlot P where A.NoticeDispatched = 'Y' and (A.Renewed is null or A.Renewed = 'N') and A.ReasonsForNotice = 'Renewal of Contract' and A.AccountNo = P.AccountNo", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT *  FROM ODASMLeaseAgreement A,ODASPPlot P where A.NoticeDispatched = 'Y' and (A.Renewed is null or A.Renewed = 'N') and A.ReasonsForNotice = 'Renewal of Contract' and A.AccountNo = P.AccountNo";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Contract No", listView1.Width / 5);
                 listView1.Columns.Add("Plot", listView1.Width / 5);
                 listView1.Columns.Add("LandLord", listView1.Width / 5);
                 listView1.Columns.Add("Expiry Date", listView1.Width / 5);
                 listView1.Columns.Add("Notice Date", listView1.Width / 5);
                



                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["ContractNo"].ToString());

                         if (reader["PlotNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotNo"].ToString());
                         }
                         if (reader["AccountNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["AccountNo"].ToString());
                         } if (Convert .ToDateTime ( reader["expirydate"].ToString()).ToString ("yyyy/MM/dd") != "")
                         {
                             lv3.SubItems.Add(Convert.ToDateTime(reader["expirydate"].ToString()).ToString("yyyy/MM/dd"));
                         } if (reader["NoticeDate"].ToString() != "")
                         {
                             lv3.SubItems.Add(Convert .ToDateTime ( reader["NoticeDate"].ToString()).ToString ("yyyy/MM/dd"));
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }

         }
         private void getCONTRACTSRenewed()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*)  FROM ODASMLeaseRenewals LR, ODASMLeaseAgreement A,ODASPPlot P where  LR.ContractNo = A.ContractNo and A.AccountNo = P.AccountNo", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT *  FROM ODASMLeaseRenewals LR, ODASMLeaseAgreement A,ODASPPlot P where  LR.ContractNo = A.ContractNo and A.AccountNo = P.AccountNo";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Contract No", listView1.Width / 6);
                 listView1.Columns.Add("Plot", listView1.Width / 6);
                 listView1.Columns.Add("LandLord", listView1.Width / 6);
                 listView1.Columns.Add("Expiry Date", listView1.Width / 6);
                 listView1.Columns.Add("Date Renewed", listView1.Width / 6);
                 listView1.Columns.Add("New Expiry Date", listView1.Width / 6);

                 


                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["ContractNo"].ToString());

                         if (reader["PlotNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotNo"].ToString());
                         }
                         if (reader["AccountNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["AccountNo"].ToString());
                         } if (Convert.ToDateTime(reader["expirydate"].ToString()).ToString("yyyy/MM/dd") != "")
                         {
                             lv3.SubItems.Add(Convert.ToDateTime(reader["expirydate"].ToString()).ToString("yyyy/MM/dd"));
                         } if (reader["RenewalDate"].ToString() != "")
                         {
                             lv3.SubItems.Add(Convert.ToDateTime(reader["RenewalDate"].ToString()).ToString("yyyy/MM/dd"));
                         }
                         if (reader["expirydate"].ToString() != "")
                         {
                             lv3.SubItems.Add(Convert.ToDateTime(reader["expirydate"].ToString()).ToString("yyyy/MM/dd"));
                         }


                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }

         }
         private void getNoticesPrepared()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*)  FROM ODASMLeaseAgreement where NoticePrepared = 'Y'", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT *  FROM ODASMLeaseAgreement where NoticePrepared = 'Y'";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Contract No", listView1.Width / 5);
                 listView1.Columns.Add("Plot", listView1.Width / 5);
                 listView1.Columns.Add("LandLord", listView1.Width / 5);
                 listView1.Columns.Add("Notice Date", listView1.Width / 5);
                 listView1.Columns.Add("Prepared By", listView1.Width / 5);
                 



                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["ContractNo"].ToString());

                         if (reader["PlotNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotNo"].ToString());
                         }
                         if (reader["AccountNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["AccountNo"].ToString());
                         } if (Convert.ToDateTime(reader["NoticeDate"].ToString()).ToString("yyyy/MM/dd") != "")
                         {
                             lv3.SubItems.Add(Convert.ToDateTime(reader["NoticeDate"].ToString()).ToString("yyyy/MM/dd"));
                         } if (reader["NoticePreparedBy"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["NoticePreparedBy"].ToString());
                         }
                        


                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }

         }
         private void getALLNoticesAuthorized()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*)  FROM ODASMLeaseAgreement where NoticeAuthorized = 'Y'", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT *  FROM ODASMLeaseAgreement where NoticeAuthorized = 'Y'";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Contract No", listView1.Width / 5);
                 listView1.Columns.Add("Plot", listView1.Width / 5);
                 listView1.Columns.Add("LandLord", listView1.Width / 5);
                 listView1.Columns.Add("Authorization Date", listView1.Width / 5);
                 listView1.Columns.Add("Authorized By", listView1.Width / 5);




                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["ContractNo"].ToString());

                         if (reader["PlotNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotNo"].ToString());
                         }
                         if (reader["AccountNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["AccountNo"].ToString());
                         } if (Convert.ToDateTime(reader["NoticeAuthorizationDate"].ToString()).ToString("yyyy/MM/dd") != "")
                         {
                             lv3.SubItems.Add(Convert.ToDateTime(reader["NoticeAuthorizationDate"].ToString()).ToString("yyyy/MM/dd"));
                         } if (reader["AuthorizedBy"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["AuthorizedBy"].ToString());
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }

         }
         private void getAllNoticesSent()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*)  FROM ODASMLeaseAgreement where NoticeDispatched = 'Y'", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT *  FROM ODASMLeaseAgreement where NoticeDispatched = 'Y'";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Contract No", listView1.Width / 5);
                 listView1.Columns.Add("Plot", listView1.Width / 5);
                 listView1.Columns.Add("LandLord", listView1.Width / 5);
                 listView1.Columns.Add("Notice Date", listView1.Width / 5);
                 listView1.Columns.Add("Prepared By", listView1.Width / 5);




                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["ContractNo"].ToString());

                         if (reader["PlotNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotNo"].ToString());
                         }
                         if (reader["AccountNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["AccountNo"].ToString());
                         } if (Convert.ToDateTime(reader["NoticeDate"].ToString()).ToString("yyyy/MM/dd") != "")
                         {
                             lv3.SubItems.Add(Convert.ToDateTime(reader["NoticeDate"].ToString()).ToString("yyyy/MM/dd"));
                         } if (reader["NoticePreparedBy"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["NoticePreparedBy"].ToString());
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }

         }
         private void getNoticesReceived()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*)  FROM ODASMLeaseAgreement where NoticeDispatched = 'Y' and NoticeReceived = 'Y'", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT *  FROM ODASMLeaseAgreement where NoticeDispatched = 'Y' and NoticeReceived = 'Y'";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Contract No", listView1.Width / 5);
                 listView1.Columns.Add("Plot", listView1.Width / 5);
                 listView1.Columns.Add("LandLord", listView1.Width / 5);
                 listView1.Columns.Add("Notice Date", listView1.Width / 5);
                 listView1.Columns.Add("Prepared By", listView1.Width / 5);




                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["ContractNo"].ToString());

                         if (reader["PlotNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotNo"].ToString());
                         }
                         if (reader["AccountNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["AccountNo"].ToString());
                         } if (Convert.ToDateTime(reader["NoticeDate"].ToString()).ToString("yyyy/MM/dd") != "")
                         {
                             lv3.SubItems.Add(Convert.ToDateTime(reader["NoticeDate"].ToString()).ToString("yyyy/MM/dd"));
                         } if (reader["NoticePreparedBy"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["NoticePreparedBy"].ToString());
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }

         }
         private void showNoticesAuthorized()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*)  FROM ODASMJobBriefItems,ODASPAccount,ODASMJobBrief where ODASMJobBriefItems.NoticeAuthorized = 'Y' and ODASMJobBriefItems.Status = 'NOTICE-AUTHORIZED' and ODASMJobBriefItems.JobBriefNo = ODASMJobBrief.JobBriefNo and ODASMJobBrief.AccountNo = ODASPAccount.AccountNo", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT (JBI.ExpiryDate)as EDate,JBI.*,JB.*,A.*  FROM ODASMJobBriefItems JBI,ODASPAccount A,ODASMJobBrief JB where JBI.NoticeAuthorized = 'Y' and JBI.Status = 'NOTICE-AUTHORIZED' and JBI.JobBriefNo = JB.JobBriefNo and JB.AccountNo = A.AccountNo";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Job Item No", listView1.Width / 5);
                 listView1.Columns.Add("Site", listView1.Width / 5);
                 listView1.Columns.Add("Client", listView1.Width / 5);
                 listView1.Columns.Add("Expiry Date", listView1.Width / 5);
                 listView1.Columns.Add("Notice Date", listView1.Width / 5);
                 listView1.Columns.Add("Renewal period", listView1.Width / 5);
                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["JobBriefItemNo"].ToString());

                         if (reader["SiteNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["SiteNo"].ToString());
                         }
                         if (reader["CompanyName"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["CompanyName"].ToString());
                         } if (Convert.ToDateTime(reader["EDate"].ToString()).ToString("yyyy/MM/dd") != "")
                         {
                             lv3.SubItems.Add(Convert.ToDateTime(reader["EDate"].ToString()).ToString("yyyy/MM/dd"));
                         } if (reader["NoticeReceivedDate"].ToString() != "")
                         {
                             lv3.SubItems.Add(Convert.ToDateTime(reader["NoticeReceivedDate"].ToString()).ToString("yyyy/MM/dd"));
                         } if (reader["RenewalPeriod"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["RenewalPeriod"].ToString());
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.ToString ());
             }

         }
         private void ShowAllValidEmptyBillBoards()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPPlot P,ODASPTown T, ODASPAccount AC WHERE P.BillBoard = 'Y' and T.TownCode = P.TownCode and P.AccountNo = AC.AccountNo ", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT * FROM ODASPPlot P,ODASPTown T, ODASPAccount AC WHERE P.BillBoard = 'Y' and T.TownCode = P.TownCode and P.AccountNo = AC.AccountNo ";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("BillBord No", listView1.Width / 7);
                 listView1.Columns.Add("BillBord Details", listView1.Width / 7);
                 listView1.Columns.Add("Location", listView1.Width / 7);
                 listView1.Columns.Add("Faces Free", listView1.Width / 7);
                 listView1.Columns.Add("Land Lord", listView1.Width / 7);
                 listView1.Columns.Add("Start Date", listView1.Width / 7);
                 listView1.Columns.Add("Expiry Date", listView1.Width / 7);
                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["MastNo"].ToString());

                         if (reader["MastDetails"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["MastDetails"].ToString());
                         }
                         if (reader["PhysicalLocation"].ToString() != "" && reader["Town"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PhysicalLocation"].ToString() + "IN" + reader["Town"].ToString());
                         }
                             lv3.SubItems.Add(c);
                             if (reader["CompanyName"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["CompanyName"].ToString());
                         } 
                         if (reader["CommencementDate"].ToString() != "")
                             {
                                 lv3.SubItems.Add(Convert.ToDateTime(reader["CommencementDate"].ToString()).ToString("yyyy/MM/dd"));
                             }
                         if (reader["expirydate"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["expirydate"].ToString());
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.ToString());
             }

         }
         private void showALLLandlords()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*)  FROM ODASPAccount WHERE AccountType = 'LLORD'", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT * FROM ODASPAccount WHERE AccountType = 'LLORD'";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("LandLord No", listView1.Width / 9);
                 listView1.Columns.Add("Name", listView1.Width / 9);
                 listView1.Columns.Add("Physical Address", listView1.Width / 9);
                 listView1.Columns.Add("City", listView1.Width / 9);
                 listView1.Columns.Add("Postal Address", listView1.Width / 9);
                 listView1.Columns.Add("Telephone", listView1.Width / 9);
                 listView1.Columns.Add("Mobile No", listView1.Width / 9);
                 listView1.Columns.Add("E-Mail", listView1.Width /9);
                 listView1.Columns.Add("Fax", listView1.Width / 9);



                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["AccountNo"].ToString());

                         if (reader["CompanyName"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["CompanyName"].ToString());
                         }
                         if (reader["PhysicalAddress"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PhysicalAddress"].ToString());
                         } if (reader["Towncity"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["Towncity"].ToString());
                         } if (reader["PostalAddress"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PostalAddress"].ToString());
                         } if (reader["TelephoneNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["TelephoneNo"].ToString());
                         } if (reader["MobileNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["MobileNo"].ToString());
                         } if (reader["EmailAddress"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["EmailAddress"].ToString());
                         } if (reader["FAxNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["FAxNo"].ToString());
                         }

                         

                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }

         }
         private void showALLSitesUnAllocated()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPPlot,ODASPPlotSite where ODASPPlotSite.Status ='SITE-AVAILABLE' and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo ", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT * FROM ODASPPlot,ODASPPlotSite where ODASPPlotSite.Status ='SITE-AVAILABLE' and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo ";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Site No", listView1.Width / 5);
                 listView1.Columns.Add("Site Details", listView1.Width / 5);
                 listView1.Columns.Add("Plot No", listView1.Width / 5);
                 listView1.Columns.Add("Plot Name", listView1.Width / 5);
                 
                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["SiteNo"].ToString());

                         if (reader["SiteDetails"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["SiteDetails"].ToString());
                         }
                         if (reader["PlotNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotNo"].ToString());
                         }  if (reader["PlotName"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotName"].ToString());
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.ToString());
             }

         }
         private void showALLSitesAllocated()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPPlot,ODASPPlotSite,ODASMJobBrief,ODASPAccount where ODASPPlot.PlotNo = ODASPPlotSite.PlotNo and ODASPPlotSite.JobBriefNo = ODASMJobBrief.JobBriefNo and ODASMJobBrief.AccountNo = ODASPAccount.AccountNo", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT * FROM ODASPPlot,ODASPPlotSite,ODASMJobBrief,ODASPAccount where ODASPPlot.PlotNo = ODASPPlotSite.PlotNo and ODASPPlotSite.JobBriefNo = ODASMJobBrief.JobBriefNo and ODASMJobBrief.AccountNo = ODASPAccount.AccountNo";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Site No", listView1.Width / 5);
                 listView1.Columns.Add("Site Details", listView1.Width / 5);
                 listView1.Columns.Add("Plot No", listView1.Width / 5);
                 listView1.Columns.Add("Plot Name", listView1.Width / 5);
                 listView1.Columns.Add("Client", listView1.Width / 5);
                 listView1.Columns.Add("Date Started", listView1.Width / 5);
                 listView1.Columns.Add("Expiry Datee", listView1.Width / 5);
                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["SiteNo"].ToString());

                         if (reader["SiteDetails"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["SiteDetails"].ToString());
                         }
                         if (reader["PlotNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotNo"].ToString());
                         } if (reader["PlotName"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotName"].ToString());
                         } if (reader["CompanyName"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["CompanyName"].ToString());
                         }
                         if (reader["JCStartDate"].ToString() != "")
                         {
                             lv3.SubItems.Add(Convert .ToDateTime ( reader["JCStartDate"].ToString()).ToString ("yyyy/MM/dd"));
                         } if (reader["JCExpiryDate"].ToString() != "")
                         {
                             lv3.SubItems.Add(Convert.ToDateTime(reader["JCExpiryDate"].ToString()).ToString("yyyy/MM/dd"));
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.ToString());
             }

         }
         private void showALLSitesReserved()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPPlot,ODASPPlotSite,ODASMJobBrief,ODASPAccount where ODASPPlotSite.Status ='SITE-RESERVED' and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo and ODASPPlotSite.JobBriefNo = ODASMJobBrief.JobBriefNo and ODASMJobBrief.AccountNo = ODASPAccount.AccountNo", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT * FROM ODASPPlot,ODASPPlotSite,ODASMJobBrief,ODASPAccount where ODASPPlotSite.Status ='SITE-RESERVED' and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo and ODASPPlotSite.JobBriefNo = ODASMJobBrief.JobBriefNo and ODASMJobBrief.AccountNo = ODASPAccount.AccountNo";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Site No", listView1.Width / 5);
                 listView1.Columns.Add("Site Details", listView1.Width / 5);
                 listView1.Columns.Add("Plot No", listView1.Width / 5);
                 listView1.Columns.Add("Plot Name", listView1.Width / 5);
                 listView1.Columns.Add("Client", listView1.Width / 5);
               
                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["SiteNo"].ToString());

                         if (reader["SiteDetails"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["SiteDetails"].ToString());
                         }
                         if (reader["PlotNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotNo"].ToString());
                         } if (reader["PlotName"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotName"].ToString());
                         } if (reader["CompanyName"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["CompanyName"].ToString());
                         }
                        



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.ToString());
             }

         }
         private void showAllJobsCompleted()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASMJobBriefItems JBI,ODASMJobBrief JB,ODASPAccount A,ODASPmedia ME WHERE JB.Closed = 'Y' and JBI.JobBriefNo = JB.JobBriefNo and JB.AccountNo = A.AccountNo and JBI.MediaCode = ME.MediaCode and JBI.ExpiryDate > '" + DateTime.Today.ToString("MMMM dd,yyyy") + "'", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT * FROM ODASMJobBriefItems JBI,ODASMJobBrief JB,ODASPAccount A,ODASPmedia ME WHERE JB.Closed = 'Y' and JBI.JobBriefNo = JB.JobBriefNo and JB.AccountNo = A.AccountNo and JBI.MediaCode = ME.MediaCode and JBI.ExpiryDate > '" + DateTime .Today .ToString( "MMMM dd,yyyy") +"'";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("JobBrief ItemNo", listView1.Width / 5);
                 listView1.Columns.Add("SiteNo", listView1.Width / 5);
                 listView1.Columns.Add("Customer", listView1.Width / 5);
                 listView1.Columns.Add("Media", listView1.Width / 5);
                 listView1.Columns.Add("Maitanance Date", listView1.Width / 5);
                 listView1.Columns.Add("Town", listView1.Width / 5);
                 listView1.Columns.Add("BB", listView1.Width / 5);

                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["JobBriefItemNo"].ToString());

                         if (reader["SiteNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["SiteNo"].ToString());
                         }
                         if (reader["CompanyName"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["CompanyName"].ToString());
                         } if (reader["MediaDescription"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["MediaDescription"].ToString());
                         } if (reader["MaintananceDueDate"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["MaintananceDueDate"].ToString());
                         } if (reader["Town"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["Town"].ToString());
                         } if (reader["BillBoard"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["BillBoard"].ToString());
                         }
                         



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.ToString());
             }

         }
         private void showALLSitesToFree()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPPlot,ODASPPlotSite,ODASMJobBrief where ODASPPlotSite.Status ='SITE-ALLOCATED' and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo and ODASPPlotSite.JobBriefNo = ODASMJobBrief.JobBriefNo and ODASPPlotSite.JobBriefNo is not null and ODASPPlotSite.JCExpiryDate <'" + DateTime.Today.ToString("MMMM dd,yyyy") + "'", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT * FROM ODASPPlot,ODASPPlotSite,ODASMJobBrief where ODASPPlotSite.Status ='SITE-ALLOCATED' and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo and ODASPPlotSite.JobBriefNo = ODASMJobBrief.JobBriefNo and ODASPPlotSite.JobBriefNo is not null and ODASPPlotSite.JCExpiryDate <'"+DateTime .Today .ToString ("MMMM dd,yyyy") + "'";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Site No", listView1.Width / 5);
                 listView1.Columns.Add("Site Details", listView1.Width / 5);
                 listView1.Columns.Add("Plot No", listView1.Width / 5);
                 listView1.Columns.Add("Plot Name", listView1.Width / 5);
                 listView1.Columns.Add("Product", listView1.Width / 5);
                 listView1.Columns.Add("Date Started", listView1.Width / 5);
                 listView1.Columns.Add("Expiry Datee", listView1.Width / 5);
                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["SiteNo"].ToString());

                         if (reader["SiteDetails"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["SiteDetails"].ToString());
                         }
                         if (reader["PlotNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotNo"].ToString());
                         } if (reader["PlotName"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotName"].ToString());
                         } if (reader["ProductCode"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["ProductCode"].ToString());
                         }
                         if (reader["JCStartDate"].ToString() != "")
                         {
                             lv3.SubItems.Add(Convert.ToDateTime(reader["JCStartDate"].ToString()).ToString("yyyy/MM/dd"));
                         } if (reader["JCExpiryDate"].ToString() != "")
                         {
                             lv3.SubItems.Add(Convert.ToDateTime(reader["JCExpiryDate"].ToString()).ToString("yyyy/MM/dd"));
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.ToString());
             }

         }
         private void ShowAllWorksDueForMaintenance()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPTown T, ODASPPlot P,ODASMJobBriefItems JBI,ODASMJobBrief JB,ODASPPlotSite PS,ODASPAccount A,ODASPmedia ME WHERE JBI.JobBriefItemNo = PS.JobBriefItemNo and JBI.JobBriefNo = JB.JobBriefNo and JB.AccountNo = A.AccountNo and JBI.MediaCode = ME.MediaCode and JBI.MaintananceDueDate > '"+ DateTime .Today .ToString ("MMMM dd,yyyy") + "' and PS.PlotNo = P.PlotNo and P.TownCode = T.TownCode", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT * FROM ODASPTown T, ODASPPlot P,ODASMJobBriefItems JBI,ODASMJobBrief JB,ODASPPlotSite PS,ODASPAccount A,ODASPmedia ME WHERE JBI.JobBriefItemNo = PS.JobBriefItemNo and JBI.JobBriefNo = JB.JobBriefNo and JB.AccountNo = A.AccountNo and JBI.MediaCode = ME.MediaCode and JBI.MaintananceDueDate > '" + DateTime .Today .ToString ("MMMM dd,yyyy") + "' and PS.PlotNo = P.PlotNo and P.TownCode = T.TownCode";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Site No", listView1.Width / 7);
                 listView1.Columns.Add("Site Details", listView1.Width / 7);
                 listView1.Columns.Add("Customer", listView1.Width / 7);
                 listView1.Columns.Add("Media", listView1.Width / 7);
                 listView1.Columns.Add("Product", listView1.Width / 7);
                 listView1.Columns.Add("Maitanance Date", listView1.Width / 7);
                 listView1.Columns.Add("Town", listView1.Width / 7);
                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["SiteNo"].ToString());

                         if (reader["SiteDetails"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["SiteDetails"].ToString());
                         }
                         if (reader["CompanyName"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["CompanyName"].ToString());
                         } if (reader["MediaDescription"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["MediaDescription"].ToString());
                         } if (reader["MaintananceDueDate"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["MaintananceDueDate"].ToString());
                         }
                         if (reader["Town"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["Town"].ToString());
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.ToString());
             }

         }
         private void ShowAllWorksDueForMaintenanceONEMonth()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASMMaintenance M,ODASPTown T, ODASPPlot P,ODASMJobBriefItems JBI,ODASMJobBrief JB,ODASPPlotSite PS,ODASPAccount A,ODASPmedia ME WHERE JBI.JobBriefItemNo = PS.JobBriefItemNo and JBI.JobBriefNo = JB.JobBriefNo and JB.AccountNo = A.AccountNo and JBI.MediaCode = ME.MediaCode and JBI.MaintananceDueDate> '" + DateTime .Today .ToString ("MMMM dd,yyyy") + "' and JBI.MaintananceDueDate = M.MaintenanceDate and M.Maintained = 'N'and PS.PlotNo = P.PlotNo and P.TownCode = T.TownCode", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT * FROM ODASMMaintenance M,ODASPTown T, ODASPPlot P,ODASMJobBriefItems JBI,ODASMJobBrief JB,ODASPPlotSite PS,ODASPAccount A,ODASPmedia ME WHERE JBI.JobBriefItemNo = PS.JobBriefItemNo and JBI.JobBriefNo = JB.JobBriefNo and JB.AccountNo = A.AccountNo and JBI.MediaCode = ME.MediaCode and JBI.MaintananceDueDate> '" + DateTime .Today .ToString ("MMMM dd,yyyy") + "' and JBI.MaintananceDueDate = M.MaintenanceDate and M.Maintained = 'N'and PS.PlotNo = P.PlotNo and P.TownCode = T.TownCode";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Site No", listView1.Width / 7);
                 listView1.Columns.Add("Site Details", listView1.Width / 7);
                 listView1.Columns.Add("Customer", listView1.Width / 7);
                 listView1.Columns.Add("Media", listView1.Width / 7);
                 listView1.Columns.Add("Product", listView1.Width / 7);
                 listView1.Columns.Add("Maitanance Date", listView1.Width / 7);
                 listView1.Columns.Add("Town", listView1.Width / 7);
                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["SiteNo"].ToString());

                         if (reader["SiteDetails"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["SiteDetails"].ToString());
                         }
                         if (reader["CompanyName"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["CompanyName"].ToString());
                         } if (reader["MediaDescription"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["MediaDescription"].ToString());
                         } if (reader["MaintananceDueDate"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["MaintananceDueDate"].ToString());
                         }
                         if (reader["Town"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["Town"].ToString());
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.ToString());
             }

         }
         private void AllSiteSchedule()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPPlot where status='SITE-ACQUIRED'", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT * FROM ODASPPlot where status='SITE-ACQUIRED'";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Plot No", listView1.Width / 4);
                 listView1.Columns.Add("Plot Name", listView1.Width / 4);
                 listView1.Columns.Add("LRNo", listView1.Width / 4);
                 listView1.Columns.Add("Physical Location", listView1.Width / 4);
                   if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["PlotNo"].ToString());

                         if (reader["PlotName"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PlotName"].ToString());
                         }
                         if (reader["LRNo"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["LRNo"].ToString());
                         } if (reader["PhysicalLocation"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["PhysicalLocation"].ToString());
                         }



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.ToString());
             }

         }
         private void getALLsites()
         {
             try
             {
                 GeneralVariables GeneralVariables = new GeneralVariables();
                 OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                 OdbcCommand cmd1;
                 cnn.Open();
                 cmd1 = new OdbcCommand();
                 cmd1 = new OdbcCommand("SELECT COUNT(*) FROM ODASPPlotSite", cnn);
                 String c;
                 c = cmd1.ExecuteScalar().ToString();
                 progressBar1.Minimum = 0;

                 progressBar1.Maximum = Convert.ToInt32(c);
                 progressBar1.Visible = true;
                 cnn.Close();
                 cnn.Open();
                 string sql = "SELECT * FROM ODASPPlotSite";

                 OdbcCommand cmd = new OdbcCommand(sql, cnn);
                 reader = cmd.ExecuteReader();
                 listView1.Items.Clear();
                 listView1.Columns.Clear();
                 listView1.Columns.Add("Plot No", listView1.Width / 2);
                 listView1.Columns.Add("Site Details", listView1.Width / 2);
              
                 if (reader.HasRows)
                 {

                     while (reader.Read())
                     {

                         ListViewItem lv3 = new ListViewItem(reader["PlotNo"].ToString());

                         if (reader["SiteDetails"].ToString() != "")
                         {
                             lv3.SubItems.Add(reader["SiteDetails"].ToString());
                         }
                       



                         listView1.Items.Add(lv3);

                         progressBar1.Value = progressBar1.Value + 1;


                     }
                 }
                 reader.Close();
                 progressBar1.Value = 0;
                 progressBar1.Visible = false;
                 cnn.Close();

             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.ToString());
             }

         }
       private void SearchPlot(){
             try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();



                string sql = "SELECT * FROM ODASPTown WHERE towncode LIKE '%" + textBox1.Text + "%' OR  town LIKE '%" + textBox1.Text + "%'";
                OdbcCommand cmd = new OdbcCommand(sql, cnn);
                reader = cmd.ExecuteReader();

                listView1.Items.Clear();
                if (reader.HasRows)
                {


                    while (reader.Read())
                    {

                        ListViewItem lv = new ListViewItem(reader.GetString(0).ToString());
                        // lv.SubItems[0].Text = reader.GetString(0).ToString();
                        lv.SubItems.Add(reader.GetString(1));
                        // lv.SubItems.Add(reader.GetString(1));
                        listView1.Items.Add(lv);




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
  
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                GeneralVariables GeneralVariables = new GeneralVariables();
                if (treeView1.SelectedNode.Name == "AssignProperties")
                {
                    SearchshowALLSitesWithoutProperties();
                   
                }
                else if (treeView1.SelectedNode.Name == "SiteAcquisition")
                {
                    SearchPlot();
                    
                }
               
               

            }
            catch (Exception ex) {
                MessageBox.Show(ex.ToString ());
            }
        }     

        private void listView1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            try
            {
                ListView.CheckedListViewItemCollection checkedItems = listView1.CheckedItems;

                foreach (ListViewItem item in checkedItems)
                {
                    if(item .Checked ==true ){
                        textBox2.Text = "";
                        textBox2.Text = item.SubItems[0].Text;
                        GeneralVariables var = new GeneralVariables();
                        if (treeView1.SelectedNode.Name == "AssignProperties")
                        {
                            var.AssignProperties.ShowDialog();
                            item.Checked = false;
                        }
                        else if (treeView1.SelectedNode.Name == "SiteAcquisition")
                        {
                            Cursor.Current = Cursors.WaitCursor;
                            var.SiteAcquisition.ShowDialog();
                            item.Checked = false;
                            Cursor.Current = Cursors.Default;
                        }
                        else if (treeView1.SelectedNode.Name == "PrepareLease")
                        {
                            var.Lease.txtPlotNo.Text = textBox2.Text;
                            var.CurrentUserName = CurrentUserName;
                            Cursor.Current = Cursors.WaitCursor;
                            var.Lease .ShowDialog();
                            item.Checked = false;
                            Cursor.Current = Cursors.Default;
                        }
                        else if (treeView1.SelectedNode.Name == "EditLease")
                        {
                            var.Lease.txtPlotNo.Text = textBox2.Text;

                            Cursor.Current = Cursors.WaitCursor;
                            var.Lease.txtContractNo.Text = item.Text;
                            var.Lease .txtNames.Text =item .SubItems [5].Text ;
                            var.Lease.ShowDialog();
                            item.Checked = false;
                            Cursor.Current = Cursors.Default;
                        }
                        else if (treeView1.SelectedNode.Name == "PrintRentInstallmentsheet")
                        {
                            var.rptODASRentPaymentInstallment.currentRecord=item .Text;
                            var.rptODASRentPaymentInstallment.ShowDialog();

                        }
                        else if (treeView1.SelectedNode.Name == "SetCouncilRates")
                        {
                            var.Councilrates.txtTownCode.Text  = item.Text;
                            var.Councilrates .txtTown .Text =item .SubItems [1].Text ;
                            var.Councilrates.ShowDialog();
                        }
                        else if (treeView1.SelectedNode.Name == "PrintRatesSchedule")
                        {
                            var.rptCouncilRates.ShowDialog();
                        }
                        else if (treeView1.SelectedNode.Name == "PrintSchedule")
                        {
                            var.rptODASAgreementForm.CurrentRecord = item.Text;
                            var.rptODASAgreementForm.ShowDialog();
                        }
                        else if (treeView1.SelectedNode.Name == "FreeAssignedSites")
                        {
                            var.frmFreeAssignedSites.txtPlotNo.Text = item.SubItems[2].Text;
                            var.frmFreeAssignedSites.txtPlotName.Text = item.SubItems[3].Text;
                            var.frmFreeAssignedSites.txtSiteNo.Text = item.Text;
                            var.frmFreeAssignedSites.txtSiteDetails.Text = item.SubItems[2].Text;
                            var.frmFreeAssignedSites.ShowDialog();
                        }
                        else {
                            GeneralVariables GeneralVariables = new GeneralVariables();
                            GeneralVariables.frmSitesReportGroupedByCouncils.strCouncilcode = item.Text;
                            GeneralVariables.frmSitesReportGroupedByCouncils.ShowDialog();
     
                        }
                    }
                   
                   
                  
                   
                }
            }
            catch (Exception ex) {
                MessageBox.Show(ex.ToString ());
            }
           
            
        }

        private void treeView1_Click(object sender, EventArgs e)
        {
           
        }

        private void treeView1_NodeMouseHover(object sender, TreeNodeMouseHoverEventArgs e)
        {
            Cursor.Current = Cursors.Hand;

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void allFreeSitesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();
            GeneralVariables.frmRptODASAllFreeSites.ShowDialog();
        }

        private void asAtSingleDateToolStripMenuItem_Click(object sender, EventArgs e)
        { GeneralVariables GeneralVariables = new GeneralVariables();
        GeneralVariables.frmODASSitesToExpire.strReport = "AsAtASingleDate";
        GeneralVariables.frmODASSitesToExpire.ShowDialog();
        }

        private void withinDateRangeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();
            GeneralVariables.frmODASSitesToExpire.strReport = "";
            GeneralVariables.frmODASSitesToExpire.ShowDialog();
        }

        private void expiredNotRewedToolStripMenuItem_Click(object sender, EventArgs e)
        {
             GeneralVariables GeneralVariables = new GeneralVariables();
             GeneralVariables.frmODASSitesToExpire.strReport = "ExpiredNotRenewed";
            GeneralVariables.frmODASSitesToExpire.ShowDialog();
            
        }

        private void siteAllocationToolStripMenuItem_Click(object sender, EventArgs e)
        {  GeneralVariables GeneralVariables = new GeneralVariables();
        GeneralVariables.frmPlotAllocation.ShowDialog();
            
        }

        private void landlordListingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();
            GeneralVariables.frmLandlordlisting.ShowDialog();
        }

        private void siteDetailsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();
            GeneralVariables.frmSearchSite.ShowDialog();
       

        }

        private void allSitesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();
            GeneralVariables.frmAllSites.ShowDialog();
       
            
        }

        private void allSitesToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();
            GeneralVariables.frmRptODASAllRoadSites.ShowDialog();
       
            
        }

        private void plotsSitesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();
            GeneralVariables.frmPlotAllocation.ShowDialog();
       
        }

        private void queryFreeBillboardsFacesToolStripMenuItem_Click(object sender, EventArgs e)
        {
             GeneralVariables GeneralVariables = new GeneralVariables();
             GeneralVariables.frmFreeBillboards.ShowDialog();
       
            
        }

        private void landlordStatementToolStripMenuItem_Click(object sender, EventArgs e)
        {
              GeneralVariables GeneralVariables = new GeneralVariables();
              GeneralVariables.frmSearchLandlord.ShowDialog();
            
        }

        private void asAtASingleDateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();
            GeneralVariables.frmODASSearchSitesNotPaid.strReport = "PendingPaymentAsAtASingleDate";
            GeneralVariables.frmODASSearchSitesNotPaid.ShowDialog();
        }

        private void withinADateRangeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();
            GeneralVariables.frmODASSearchSitesNotPaid.strReport = "";
            GeneralVariables.frmODASSearchSitesNotPaid.ShowDialog();
        }

        private void paymentVourchersToolStripMenuItem_Click(object sender, EventArgs e)
        {
             GeneralVariables GeneralVariables = new GeneralVariables();
             GeneralVariables.frmUVouchersPrepared.ShowDialog();
     
            
        }

        private void allRightsToolStripMenuItem_Click(object sender, EventArgs e)
        {   
            GeneralVariables GeneralVariables = new GeneralVariables();
            GeneralVariables.frmSitesReportGroupedByCouncils.strCouncilcode = "";
              GeneralVariables.frmSitesReportGroupedByCouncils.ShowDialog();
     
            
        }

        private void filterCouncilToolStripMenuItem_Click(object sender, EventArgs e)
        {
            showALLCOUNCILS();
        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {

        }

        private void clearTheScreenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.Columns.Clear();
            listView1.Items.Clear();
        }

        private void updateFlagToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GeneralVariables vars= new GeneralVariables ();
            vars.frmUpdatePaymentFlag.ShowDialog();
        }

        private void landlordRightsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GeneralVariables vars = new GeneralVariables();
            vars.LandLord.ShowDialog();
        }

        private void aboutTheSystemToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmabout frmabout = new frmabout();
            frmabout.ShowDialog();
        }

        private void sitePropertiesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            showALLSitesWithoutProperties();
        }

        private void sitesReportBasedOnDateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GeneralVariables vars = new GeneralVariables();
            vars.frmUSelectDateRange.ShowDialog();
        }

        private void transactionPendingToolStripMenuItem_Click(object sender, EventArgs e)
        { GeneralVariables vars = new GeneralVariables();
        vars.frmODASSearchSitesNotPaid.strReport = "PendingPayment";
        vars.frmODASSearchSitesNotPaid.ShowDialog();
        }

        private void vourcherPreparedToolStripMenuItem_Click(object sender, EventArgs e)
        {GeneralVariables vars = new GeneralVariables();
        vars.frmODASSearchSitesNotPaid.strReport = "VouchersPrepared";
        vars.frmODASSearchSitesNotPaid.ShowDialog();
           
        }

        private void transactionPendingToolStripMenuItem1_Click(object sender, EventArgs e)
        {GeneralVariables vars = new GeneralVariables();
        vars.frmODASSearchSitesNotPaid.strReport = "PendingConfirmation";
        vars.frmODASSearchSitesNotPaid.ShowDialog();
        }

        private void paymentsConfirmedToolStripMenuItem_Click(object sender, EventArgs e)
        {GeneralVariables vars = new GeneralVariables();
        vars.frmODASSearchSitesNotPaid.strReport = "PaymentsConfirmed";
        vars.frmODASSearchSitesNotPaid.ShowDialog();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            GeneralVariables vars = new GeneralVariables();
            vars.SiteAcquisition.CurrentUserName = CurrentUserName;
            vars.SiteAcquisition.ShowDialog();
            Cursor.Current = Cursors.Default;
        }

        private void rentPaidToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GeneralVariables vars = new GeneralVariables();
            vars.frmODASSearchSitesNotPaid.strReport = "PaymentsConfirmed";
            vars.frmODASSearchSitesNotPaid.ShowDialog();
        }

        private void frmMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();
            Cursor.Current = Cursors.WaitCursor;
            GeneralVariables.PaymentConfirmation.CurrentUserName = CurrentUserName;
            GeneralVariables.PaymentConfirmation.cboPaymentCode.Text = "RENT";

            GeneralVariables.PaymentConfirmation.ShowDialog();
            Cursor.Current = Cursors.Default;
        }

        private void button4_Click(object sender, EventArgs e)
        {
             GeneralVariables GeneralVariables = new GeneralVariables();
             GeneralVariables.frmODASMContractTermination.CurrentUserName = CurrentUserName;
             GeneralVariables.frmODASMContractTermination.ShowDialog();
        }

      

     
    }
}
