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
    public partial class frmFreeAssignedSites : Form
    {
        public frmFreeAssignedSites()
        {
            InitializeComponent();
        }
        OdbcCommand cmd;
        OdbcDataReader reader;
        private void LoadPLOTDetails()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT ODASPPlot.PlotNo,ODASPPlotSite.SiteNo,ODASPPlotSite.JobBriefNo,ODASPPlot.PhysicalLocation,ODASPPlot.CommencementDate,ODASPPlot.expirydate FROM ODASPPlot,ODASPPlotSite where ODASPPlot.PlotNo='" +txtPlotNo.Text  + "' AND ODASPPlotSite.SiteNo='" +txtSiteNo.Text  +"'", cnn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    txtPhysicalLocation.Text = reader["PhysicalLocation"].ToString();
                    txtJobBriefNo.Text = reader["JobBriefNo"].ToString();
                    txtStartDate.Text = Convert.ToDateTime(reader["CommencementDate"].ToString()).ToString("MM/dd/yyyy");
                    txtEndDate.Text =Convert .ToDateTime ( reader["expirydate"].ToString()).ToString ("MM/dd/yyyy");
                       
                }
                reader.Close();
                cnn.Close();
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }
        private void saveRecord()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("Select * from ODASPPlotSite Where SiteNo = '" +txtSiteNo.Text + "'", cnn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    cnn.Close();
                    cnn.Open();
                    cmd = new OdbcCommand("UPDATE ODASPPlotSite SET JobBriefNo='',Status='SITE-AVAILABLE' Where SiteNo = '" + txtSiteNo.Text + "'", cnn);

                }
                reader.Close();
                cnn.Close();
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }
        private void LoadAccount()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT * FROM ODASMJobBrief,ODASPAccount where ODASMJobBrief.AccountNo=ODASPAccount.AccountNo AND ODASMJobBrief.JobBriefNo='" + txtJobBriefNo.Text + "'", cnn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    txtAccountNo.Text = reader["AccountNo"].ToString();
                    txtName.Text = reader["CompanyName"].ToString();

                }
                cnn.Close();
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        }
        private void frmFreeAssignedSites_Load(object sender, EventArgs e)
        {
            LoadPLOTDetails();
            LoadAccount();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            saveRecord();
            MessageBox.Show("Site Freed");
        }
    }
}
