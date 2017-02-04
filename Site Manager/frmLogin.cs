using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Microsoft.Win32;
using System.Security.Principal;
namespace Site_Manager
{
    public partial class frmLogin : Form
    {
        public frmLogin()
        {
            InitializeComponent();
        }
        public String cC, k, xx;
        OdbcDataReader reader;
        OdbcCommand cmd;
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private static string FormRegKey(string sSect)
        {
            return sSect;
        }
        public  string GetSetting(String appName, String section,string key,string Default) {
            try {
                Object obj1 = null;
            if(Default ==null){
                Default = "";

            }
            key = txtPassword.Text.Trim();
            string text2 = FormRegKey(section);
            RegistryKey key1 = Application.UserAppDataRegistry.OpenSubKey(text2);
       
            if(key1 !=null){
                obj1  = key1.GetValue(key, Default);
              
                key1.Close();
                if(obj1 !=null){
                if(!(obj1 is string )){
                    return null;
                }
                return (string)obj1;
                }
                return (string)obj1;
                  }
            return key;
          
        } catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return Default;
            }
        }
        private Boolean validate() {
            try {
                GeneralVariables vars = new GeneralVariables();
                OdbcConnection con = new OdbcConnection(vars.SQLstr2);
                con.Open();
                cmd = new OdbcCommand("SELECT UserMaster.UserName FROM UserMaster WHERE UserName='" + txtUsername.Text.Trim() + "'", con);

                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    return true;
                }
                else {

                    MessageBox.Show("Wrong Username or Password","Access Denied",MessageBoxButtons .OK ,MessageBoxIcon.Error );
                    txtPassword.Text = "";
                    txtPassword.Focus();
                    
                    return false;
                }
                con.Close();
                reader.Close();
            }catch (Exception ex){
               
                MessageBox.Show(ex.ToString ());
                return false;
            }
          
        }
        private Boolean ActiveUserProfile()
        {
            try
            {
                GeneralVariables vars = new GeneralVariables();
                OdbcConnection con = new OdbcConnection(vars.SQLstr2);
                con.Open();
                cmd = new OdbcCommand("SELECT UserMaster.Active FROM UserMaster WHERE UserName='" + txtUsername.Text.Trim() + "'", con);

                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    if (reader["Active"].ToString()=="")
                    {
                        return false;
                    }
                    else if (reader["Active"].ToString() == "I" || reader["Active"].ToString() == "i" || reader["Active"].ToString() == "N")
                    {
                        return false;

                    }
                    else if (reader["Active"].ToString() == "A" || reader["Active"].ToString() == "a" || reader["Active"].ToString() == "Y")
                    {
                        return true;

                    }
                    else {
                        return false;
                    }
                    if (!ActiveUserProfile())
                    {
                        MessageBox.Show("Access Denied. The specified user profile is disabled. Consult your system administrator for assistance...!","Access Denied",MessageBoxButtons .OK ,MessageBoxIcon.Error );
                    }
                }
                else
                {
                    return false;
                }
                con.Close();
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString());
                return false;
            }

        }
        private Boolean UserIsAdministrator() {
            try
            {
                GeneralVariables vars = new GeneralVariables();
                OdbcConnection con = new OdbcConnection(vars.SQLstr2);
                con.Open();
                cmd = new OdbcCommand("SELECT UserMaster.GroupNo FROM UserMaster WHERE UserName='" +txtUsername.Text.Trim ()+ "'", con);

                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    if (reader["GroupNo"].ToString() == "")
                    {
                        return false;
                    }
                    else if (reader["GroupNo"].ToString() == "A")
                    {

                        return true;
                    }
                    else {
                        return false;
                    }
                }
                else {
                    return false;
                }
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
                return false;
            }
        }
        private Boolean HasModuleRights() {
          
            try
            {
                String appName = System.Windows.Forms.Application.ProductName;
                GeneralVariables vars = new GeneralVariables();
                OdbcConnection con = new OdbcConnection(vars.SQLstr2);
                con.Open();
                cmd = new OdbcCommand("SELECT UserModules.allow FROM UserModules WHERE UserName='" + txtUsername.Text.Trim ()+ "' AND exeName ='" + appName + "'", con);
                if (UserIsAdministrator())
                {
                    return true;
                }else {
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        if (reader["Allow"].ToString() == "")
                        {
                            return false;

                        }
                        else if (reader["Allow"].ToString() == "0" ||Convert .ToInt32 ( reader["Allow"].ToString()) == 0 || reader["Allow"].ToString() == "N")
                        {
                            return false;
                        }
                        else if (reader["Allow"].ToString() == "1" || Convert.ToInt32(reader["Allow"].ToString()) == 1 || reader["Allow"].ToString() == "Y")
                        {
                            return true;
                        }
                        else {
                            return false;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
                if (HasModuleRights()==false )
                {
              MessageBox .Show ("Access Denied. You Do NOT have rights to log into this module. Consult your system administrator for assistance...!","Account Disabled",MessageBoxButtons .OK ,MessageBoxIcon.Error );
              }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return false;
            }
        
        }
        public String   GetFullEncryption() {
            
               
                int i, j;
               
                xx = "";
                j = txtPassword.Text.Length;
               
            if(j==0){
                txtPassword.Focus();
              
            }
            for (i = 0; i <= j - 1;i++ )
            {
                txtPassword.SelectionStart = i;
                txtPassword.SelectionLength = 1;
                cC = txtPassword.SelectedText;
                if(cC==""){
                    k = "";
                 
                }else {
                    k = GetSetting("SmallSyzSecure", "SysSecureEncryptor",cC,"");
                    xx = xx + k;
                }
            }
           
            if (xx == "")
            {
                SaveEncryptionCode();
                SaveDecryptionCode();
                GetFullEncryption();
               
                return xx;
               
            }
            else {
               return xx;
            }

              
        }
        public void SaveSetting(String appName ,string Section, string Key, string Setting)
        {
            Key = txtPassword.Text.Trim();
            string text1 = FormRegKey(Section);
            RegistryKey key1 = Application.UserAppDataRegistry.CreateSubKey(text1);
                 
           
            if (key1 == null)
            {
                return;
            }
            try
            {
                key1.SetValue(Key, Setting);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
            finally
            {
                key1.Close();
            }
        }
        public void SaveEncryptionCode() {
            try {
                SaveSetting("SmallSyzSecure", "SysSecureEncryptor", "a", "!");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "b", "@");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "c", "#");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "d", "$");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "e", "%");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "f", "^");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "g", "&");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "h", "*");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "i", "(");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "j", ")");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "k", "-");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "l", "_");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "m", "=");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "n", "+");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "o", "\\");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "p", "|");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "q", "/");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "r", ">");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "s", "<");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "t", "?");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "u", "[");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "v", "]");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "w", "~");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "x", "{");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "y", "}");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "z", ",");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "0", "Z");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "1", "Y");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "2", "X");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "3", "W");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "4", "V");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "5", "U");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "6", "T");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "7", "S");
                SaveSetting ("SmallSyzSecure", "SysSecureEncryptor", "8", "R");
                SaveSetting("SmallSyzSecure", "SysSecureEncryptor", "9", "Q");



            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
        }
        private void SaveDecryptionCode() { 
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "!", "a");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "@", "b");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "#", "c");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "$", "d");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "%", "e");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "^", "f");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "&", "g");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "*", "h");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "(", "i");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", ")", "j");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "-", "k");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "_", "l");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "=", "m");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "+", "n");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "\\", "o");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "|", "p");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "/", "q");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", ">", "r");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "<", "s");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "?", "t");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "[", "u");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "]", "v");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "~", "w");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "{", "x");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "}", "y");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", ",", "z");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "Z", "0");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "Y", "1");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "X", "2");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "W", "3");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "V", "4");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "U", "5");
                        SaveSetting("SmallSyzSecure", "SysSecureDecryptor", "T", "6");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "S", "7");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "R", "8");
                        SaveSetting ("SmallSyzSecure", "SysSecureDecryptor", "Q", "9");
        
        }
        private Boolean ValidPassword() {
                
            GeneralVariables vars = new GeneralVariables();
            OdbcConnection con = new OdbcConnection(vars.SQLstr2);
            con.Open();
            cmd = new OdbcCommand("SELECT UserMaster.password FROM UserMaster WHERE UserName='" + txtUsername.Text.Trim() + "'", con);
            reader = cmd.ExecuteReader();

            try {
                if(reader .Read ()){
                    if (reader["Password"].ToString() == "")
                    {
                        return false;
                    }
                    else if (reader["Password"].ToString() != txtPassword.Text.Trim ())
                    {
                        return false;
                    }
                    else if (reader["Password"].ToString() == txtPassword.Text.Trim())
                    {
                        return true;
                    }
                    else {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
                return false;
            }
        }
        private void SaveLoginRecord() { 
         GeneralVariables vars = new GeneralVariables();
            OdbcConnection con = new OdbcConnection(vars.SQLstr2);
            con.Open();
            try {
                vars.CLoginID = GetNextLoginID();
               
               // System.Environment.MachineName;
                WindowsIdentity.GetCurrent().ToString();
                vars.CurrentUserName = txtUsername.Text.Trim();
                cmd = new OdbcCommand("INSERT INTO UserLog(LoginID,UserName,LoginDate,LoginTime,CompName,SystemUsed)VALUES(" + vars.CLoginID + ",'" + txtUsername.Text + "','" + DateTime.Today.ToString("MMMM dd,yyyy") + "','" + DateTime.Today.ToLongTimeString() + "','" + WindowsIdentity.GetCurrent().Name.ToString() + "','" + Application.ProductName + "')", con);
                
             
                cmd.ExecuteNonQuery();
            } catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
        }
        private long  GetNextLoginID()
        {
            long logid = 1;
            GeneralVariables vars = new GeneralVariables();
            OdbcConnection con = new OdbcConnection(vars.SQLstr2);
           
            try
            {
                con.Open();
             cmd = new OdbcCommand("SELECT MAX(LoginID) AS LastID FROM UserLog WHERE LoginID IS NOT NULL",con);
             reader = cmd.ExecuteReader();
             if (reader.Read())
             {
                 if (reader["lastid"].ToString() == "")
                 {
                     logid = 1;
                 }
                 else {
                     logid = Convert .ToInt64 (reader["lastid"].ToString());
                 }
             }
           
             else {
                 logid = 1;
             }
            
             return logid + 1;
               
            }
               
            catch (Exception ex)
            {
             
               return 1+1;
            }
        }
        private void btnLogin_Click(object sender, EventArgs e)
        {
            GeneralVariables vars = new GeneralVariables();
            encrption passEncrption = new encrption();
            int i=0;
         
            try {
               
            if(txtUsername .Text ==""){
                MessageBox.Show("Enter your Username.","Information Required",MessageBoxButtons .OK,MessageBoxIcon.Error );
                txtUsername.Focus();
            }
            else if (txtPassword.Text == "")
            {
                MessageBox.Show("Enter your Passowrd.", "Information Required", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtPassword.Focus();
            }
            else {
                progressBar1.Visible = true;
               
                if (validate())
                {
                    progressBar1.Value = progressBar1.Value + 1;
                    if (ActiveUserProfile())
                    {
                        progressBar1.Value = progressBar1.Value + 1;
                        if (HasModuleRights())
                        {
                            progressBar1.Value = progressBar1.Value + 1;
                            String Password;
                            Password = txtPassword.Text.Trim();
                            txtPassword.Text =passEncrption .GetFullEncryption (Password);
                            progressBar1.Value = progressBar1.Value + 1;
                            if (ValidPassword())
                            {
                                progressBar1.Value = progressBar1.Value + 1;
                                SaveLoginRecord();
                                vars.MainForm.lblComputer.Text = WindowsIdentity.GetCurrent().Name.ToString();
                                vars.MainForm.lblDate.Text = DateTime.Today.ToString("MM/dd/yyyy");
                                vars.MainForm.lblTime.Text = DateTime.Now.ToString("h:mm:ss tt");
                                vars.MainForm  .CurrentUserName = txtUsername.Text;
                                vars.MainForm.lbluser.Text = txtUsername.Text;
                                progressBar1.Value = 0;
                                progressBar1.Visible = false;
                                this.Hide();
                                vars.MainForm.ShowDialog();
                            }
                            else
                            {
                                txtPassword.Text = "";
                                txtPassword.Focus();
                                progressBar1.Value = 0;
                                progressBar1.Visible = false;
                                i = i + 1;
                                if (i == 3)
                                {
                                    MessageBox.Show("Too many failed logon attempts. The system shuts down...!", "Forced Shudown", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                        }
                        else {
                            i = i + 1;
                            if (i == 3)
                            {
                                MessageBox.Show("Too many failed logon attempts. The system shuts down...!", "Forced Shudown", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                    else
                    {
                        i = i + 1;
                        if (i == 3)
                        {
                            MessageBox.Show("Too many failed logon attempts. The system shuts down...!", "Forced Shudown", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else
                {
                    i = i + 1;
                    if (i == 3)
                    {
                        MessageBox.Show("Too many failed logon attempts. The system shuts down...!", "Forced Shudown", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
               
            
            }


            }catch (Exception ex){
                MessageBox.Show(ex.ToString ());
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void frmLogin_Load(object sender, EventArgs e)
        {
            this.AcceptButton = btnLogin;
           
        }
    }
}
