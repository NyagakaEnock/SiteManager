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
    public partial class frmUpdatePaymentFlag : Form
    {
        public frmUpdatePaymentFlag()
        {
            InitializeComponent();
        }
        OdbcCommand CMD;
        OdbcDataReader reader;

        private void frmUpdatePaymentFlag_Load(object sender, EventArgs e)
        {

        }
        private void updatePaymentFlag() {
            try
            {
                
                GeneralVariables vars = new GeneralVariables();
                OdbcConnection con = new OdbcConnection(vars.SQLstr);
                OdbcConnection con2 = new OdbcConnection(vars.SQLstr);
               
                con.Open();
                CMD = new OdbcCommand("SELECT COUNT(*) FROM ODASMInstallment", con);
                int c = Convert.ToInt32(CMD.ExecuteScalar().ToString());
                con.Close();
                con.Open();
                OdbcDataReader odbr;
                CMD = new OdbcCommand("SELECT * FROM ODASMInstallment", con);
                reader = CMD.ExecuteReader();
               
                if (reader.Read())
                {
                   
                    progressBar1.Visible = true;
                    progressBar1.Value = 0;
                    progressBar1.Minimum = 0;
                    progressBar1.Maximum = c;
                    con2.Open();
                    String strSAVE = "Select * From ODASMInstallment where InstallmentNo = '" + reader["InstallmentNo"].ToString() + "'";
                    CMD = new OdbcCommand(strSAVE, con2);
                    odbr = CMD.ExecuteReader();
                  

                        while (reader.Read())
                        {
                           
                            con2.Close();
                            if (con2.State == ConnectionState.Closed)
                            {
                                con2.Open();
                            }

                            if (Convert.ToDouble(reader["PaymentDue"].ToString()) > 0 && Convert.ToDouble(reader["AmountPaid"].ToString()) > 0 && Convert.ToDouble(reader["AmountPaid"].ToString()) < Convert.ToDouble(reader["PaymentDue"].ToString()))
                            {


                                CMD = new OdbcCommand("UPDATE ODASMInstallment SET PaymentFlag='N' where InstallmentNo = '" + reader["InstallmentNo"].ToString() + "'", con2);
                                CMD.ExecuteNonQuery();
                            }
                            else if (Convert.ToDouble(reader["PaymentDue"].ToString()) == 0 && Convert.ToDouble(reader["AmountPaid"].ToString()) == 0)
                            {


                                CMD = new OdbcCommand("UPDATE ODASMInstallment SET PaymentFlag='N' where InstallmentNo = '" + reader["InstallmentNo"].ToString() + "'", con2);
                                CMD.ExecuteNonQuery();

                            }
                            else
                            {
                                if (Convert.ToDouble(reader["AmountPaid"].ToString()) == Convert.ToDouble(reader["PaymentDue"].ToString()))
                                {


                                    CMD = new OdbcCommand("UPDATE ODASMInstallment SET PaymentFlag='Y' where InstallmentNo = '" + reader["InstallmentNo"].ToString() + "'", con2);
                                    CMD.ExecuteNonQuery();
                                }
                            }
                            progressBar1.Value = progressBar1.Value + 1;
                        
                        }
                    
                    progressBar1.Value = 0;
                    progressBar1.Visible = false;
                    MessageBox.Show("Payment Flag upto date", "Updated Successfully", MessageBoxButtons.OK);
                }
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            updatePaymentFlag();
        }
    }
}
