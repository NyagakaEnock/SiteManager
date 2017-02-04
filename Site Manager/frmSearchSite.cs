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
    public partial class frmSearchSite : Form
    {
        public frmSearchSite()
        {
            InitializeComponent();
        }
        OdbcCommand cmd;
        DataSet ds;
        OdbcDataAdapter da;
        DataTable dTable;
        OdbcDataReader reader;
        private void loadPlots() {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();

                ds = new DataSet();
                cmd = new OdbcCommand("SELECT PlotName FROM ODASPPlot", cnn);
                da = new OdbcDataAdapter(cmd);
                dTable = new DataTable();
                da.Fill(ds, "ODASPPlot");
                dTable = ds.Tables[0];
                cbmSearch.Items.Clear();

                foreach (DataRow drow in dTable.Rows)
                {
                    cbmSearch.Items.Add(drow["PlotName"].ToString());


                }
                cnn.Close();
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }
        
        }
        private void frmSearchSite_Load(object sender, EventArgs e)
        {
            loadPlots();
            this.AcceptButton = button1;
        }

        private void cbmSearch_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();
                cmd = new OdbcCommand("SELECT PlotNo FROM ODASPPlot WHERE PlotName='" + cbmSearch.SelectedItem.ToString() + "'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {


                    txtSearch.Text = reader["PlotNo"].ToString();

                }
                cnn.Close();

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.Message);

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GeneralVariables GeneralVariables = new GeneralVariables();
            GeneralVariables.rptPlotSites2.currentRecord = txtSearch.Text;
            GeneralVariables.rptPlotSites2.ShowDialog();
        }

        private void cbmSearch_TextChanged(object sender, EventArgs e)
        {
            
        }
    }
}
