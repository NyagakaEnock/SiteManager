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
    public partial class frmFreeBillboards : Form
    {
        public frmFreeBillboards()
        {
            InitializeComponent();
        }
        OdbcCommand cmd;
        DataSet ds;
        OdbcDataAdapter da;
        DataTable dTable;
        OdbcDataReader reader;
        private void loadPlots()
        {
            try
            {

                GeneralVariables GeneralVariables = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(GeneralVariables.SQLstr);
                cnn.Open();

                ds = new DataSet();
                cmd = new OdbcCommand("SELECT Town FROM ODASPTown", cnn);
                da = new OdbcDataAdapter(cmd);
                dTable = new DataTable();
                da.Fill(ds, "ODASPTown");
                dTable = ds.Tables[0];
                cbmSearch.Items.Clear();

                foreach (DataRow drow in dTable.Rows)
                {
                    cbmSearch.Items.Add(drow["Town"].ToString());


                }
                cnn.Close();
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());

            }

        }
        private void frmFreeBillboards_Load(object sender, EventArgs e)
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
                cmd = new OdbcCommand("SELECT TownCode FROM ODASPTown WHERE Town='" + cbmSearch.SelectedItem.ToString() + "'", cnn);
                reader = cmd.ExecuteReader();

                if (reader.Read())
                {


                    txtSearch.Text = reader["TownCode"].ToString();

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
          
            GeneralVariables vars = new GeneralVariables();
            vars.frmFreeBillboards.Hide();
            vars.MainForm.showALLAvailableFaces();
        }
    }
}
