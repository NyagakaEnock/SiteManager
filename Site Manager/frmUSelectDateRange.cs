using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Site_Manager
{
    public partial class frmUSelectDateRange : Form
    {
        public frmUSelectDateRange()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GeneralVariables vars = new GeneralVariables();
            vars.frmRptODASAllSitesBasedOnDate.strStartDate = DTPickerStartDate.Text ;
            vars.frmRptODASAllSitesBasedOnDate.strLastDate = DTPickerLastDate.Text;
            vars.frmRptODASAllSitesBasedOnDate.ShowDialog();

        }
    }
}
