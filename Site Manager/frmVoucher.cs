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
    public partial class frmVoucher : Form
    {
        public frmVoucher()
        {
            InitializeComponent();
        }

        OdbcCommand cmd;
       OdbcDataReader reader;
        public String currentRecord;
        public string CurrentItem(String currentItem)
        {
            

            return currentItem;
        }

        private void frmVoucher_Load(object sender, EventArgs e)
        {
          


            try {
                rptVoucher vourcher= new rptVoucher();
               
                GeneralVariables vari = new GeneralVariables();
                OdbcConnection cnn = new OdbcConnection(vari.SQLstr);
                cnn.Open();
                OdbcDataAdapter da;
              
                cmd = new OdbcCommand("Select * From ODASMVoucher Where VoucherNo = '" + currentRecord + "' ", cnn);
                DataSet ds;
                ds = new DataSet();
                da = new OdbcDataAdapter(cmd);
                da.Fill(ds, "ODASMVoucher");

               // vourcher.RecordSelectionFormula = "{ODASMVoucher.VoucherNo}= '" + currentRecord + " '";
                vourcher.SetDataSource(ds);
                crystalReportViewer1.ReportSource = vourcher;
           
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
        }
    }
}
