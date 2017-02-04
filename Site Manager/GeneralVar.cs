using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.Odbc;


namespace Site_Manager
{
    public class GeneralVariables
    {
        public String ContarctNo="8888";
        public frmLogin LoginForm = new frmLogin() ;
        public frmAllSites frmAllSites = new frmAllSites();
        public frmRptODASAllRoadSites frmRptODASAllRoadSites = new frmRptODASAllRoadSites();
        public frmFreeBillboards frmFreeBillboards = new frmFreeBillboards();
        public frmSite_Acquisition SiteAcquisition = new frmSite_Acquisition();
        public frmUVouchersPrepared frmUVouchersPrepared = new frmUVouchersPrepared();
        public frmMain MainForm = new frmMain();
        public int counter=0;
        public frmVoucherPrepare VourcherPrepareForm = new frmVoucherPrepare();
        public frmUpdatePaymentFlag frmUpdatePaymentFlag = new frmUpdatePaymentFlag();
        public frmRptODASAllSitesBasedOnDate frmRptODASAllSitesBasedOnDate = new frmRptODASAllSitesBasedOnDate();
        public frmSearchLandlord frmSearchLandlord = new frmSearchLandlord();
        public frmRStatement frmRStatement = new frmRStatement();
        public frmSitesReportGroupedByCouncils frmSitesReportGroupedByCouncils = new frmSitesReportGroupedByCouncils();
        public frmAgreementForm AgreementForm = new frmAgreementForm();
        public frmLandLord LandLord = new frmLandLord();
        public frmLease Lease = new frmLease();
        public frmassignProperties AssignProperties = new frmassignProperties();
        public frmUSelectDateRange frmUSelectDateRange = new frmUSelectDateRange();
        public frmPaymentConfirmation PaymentConfirmation = new frmPaymentConfirmation();
        public frmCouncilRates Councilrates = new frmCouncilRates();
        public frmVoucher vourcherReport = new frmVoucher();
        public rptODASRentPaymentInstallment rptODASRentPaymentInstallment = new rptODASRentPaymentInstallment();
        public rptContractAgreement rptContractAgreement = new rptContractAgreement();
        public rptODASPPlotSites rptODASPPlotSites = new rptODASPPlotSites();
        public rptPlotSites2 rptPlotSites2 = new rptPlotSites2();
        public bool bBillBoard;
        public frmSearchSite frmSearchSite = new frmSearchSite();
        public frmPlotAllocation frmPlotAllocation = new frmPlotAllocation();
        public bool bStreetSign;
        public frmLandlordlisting frmLandlordlisting = new frmLandlordlisting();
        public frmRptODASAllFreeSites frmRptODASAllFreeSites = new frmRptODASAllFreeSites();
        public frmODASSitesToExpire frmODASSitesToExpire = new frmODASSitesToExpire();
        public Sites_to_Expire Sites_to_Expire = new Sites_to_Expire();
        public rptCouncilRates rptCouncilRates = new rptCouncilRates();
        public rptODASAgreementForm rptODASAgreementForm = new rptODASAgreementForm();
        public frmFreeAssignedSites frmFreeAssignedSites = new frmFreeAssignedSites();
        public frmODASSearchSitesNotPaid frmODASSearchSitesNotPaid = new frmODASSearchSitesNotPaid();
        public rptRentDue rptRentDue = new rptRentDue();
        public frmODASSearchPaidSites frmODASSearchPaidSites = new frmODASSearchPaidSites();
        public Boolean entryinProgress;
        public  OdbcDataReader  reader;
        public String SQLstr = "DSN=ODAS;uid=sa;pwd=administrator$";
        public frmODASMContractTermination frmODASMContractTermination = new frmODASMContractTermination();
        public String SQLstr2 = "DSN=RIGHTS;uid=sa;pwd=administrator$";
        
        public OdbcCommand cmd;
        public OdbcConnection cnn;
        public int C = 0;
        public string towncode;
        public static String str="0";
        public Boolean search;
        public long CLoginID;
        public string CurrentUserName;
        public Boolean bAuthorizeREQUISITION;
        public Boolean bapproveREQUISITION;
        public Boolean bMedicalRequisition;
        public Boolean NewRecord;
        public Boolean bLoadRecord;
        public Boolean bPlotRenewal;
        public Boolean bAllowProcess;
        public Double  PaymentsInAYear;
        public Boolean  valid;
        public  void dbConnect()
        {

            OdbcConnection cnn = new OdbcConnection("dsn=ODAS;uid=sa;pwd=administrator$;");
            cnn.Open();
         
        }
    }
    

}
