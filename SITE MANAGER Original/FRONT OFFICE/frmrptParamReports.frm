VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmParamReports 
   Caption         =   "Form3"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmrptParamReports.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
End
Attribute VB_Name = "frmParamReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New rptparamcity

Option Explicit

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
       '====
       If surrendersetup = True Then
           Dim Reportsurrendersetup As New rptparamsurrender
            CRViewer1.ReportSource = Reportsurrendersetup
            frmParamReports.Caption = "LIST OF SURRENDER SETUP"
            surrendersetup = False
        End If
        If periodssetup = True Then
            Dim Reportperiodssetup As New rptparamperiodssetup
            CRViewer1.ReportSource = Reportperiodssetup
            frmParamReports.Caption = "LIST OF PERIDS SETUP"
            periodssetup = False
        End If
        If companydetails = True Then
            Dim Reportcompanydetails As New rptparamcompanydetails
            CRViewer1.ReportSource = Reportcompanydetails
            frmParamReports.Caption = "LIST OF COMPANY DETAILS"
            companydetails = False
        End If
        If ratetable = True Then
            Dim Reportratetable As New rptparamratetable
            CRViewer1.ReportSource = Reportratetable
            frmParamReports.Caption = "LIST OF RATE TABLE"
            ratetable = False
        End If
        If jointage = True Then
            Dim Reportjointage As New rptparamjointage
            CRViewer1.ReportSource = Reportjointage
            frmParamReports.Caption = "LIST OF JOINT AGE"
            jointage = False
        End If
        If bankssetup = True Then
            Dim Reportbankssetup As New rptparambankssetup
            CRViewer1.ReportSource = Reportbankssetup
            frmParamReports.Caption = "LIST OF BANKS SETUP"
            bankssetup = False
        End If
        If lastnumbers = True Then
            Dim Reportlastnumber As New rptparamlastnumber
            CRViewer1.ReportSource = Reportlastnumber
            frmParamReports.Caption = "LIST OF LAST NUMBER"
            lastnumbers = False
        End If
       '=====
       If agentsbenefits = True Then
            Dim Reportagentsbenefits As New rptparamagentsbenefits
            CRViewer1.ReportSource = Reportagentsbenefits
            frmParamReports.Caption = "LIST OF AGENTS BENEFITS"
            agentsbenefits = False
       End If
       If claimconfig = True Then
            Dim Reportclaimsconfig As New rptclaimsconfiguration
            CRViewer1.ReportSource = Reportclaimsconfig
            frmParamReports.Caption = "LIST OF CLAIMS CONFIGURATION"
            claimconfig = False
        End If
         If loanoptype = True Then
            Dim Reportloansoptype As New rptloanoperationstype
            CRViewer1.ReportSource = Reportloansoptype
            frmParamReports.Caption = "LIST OF LOAANS OPERATION TYPE"
            loanoptype = False
        End If
         If claimcauses = True Then
            Dim Reportclaimcauses As New rptparamclaimcauses
            CRViewer1.ReportSource = Reportclaimcauses
            frmParamReports.Caption = "LIST OF CLAIM CAUSES"
            claimcauses = False
        End If
         If claimsreq = True Then
            Dim Reportclaimreq As New rptparamclaimrequirement
            CRViewer1.ReportSource = Reportclaimreq
            frmParamReports.Caption = "LIST OF CLAIM REQUIREMENT"
            claimsreq = False
        End If
        If loanapprovers = True Then
            Dim Reportloanapprovers As New rptparamloanapprovers
            CRViewer1.ReportSource = Reportloanapprovers
            frmParamReports.Caption = "LIST OF LOAAN APPROVERS"
            loanapprovers = False
        End If
         If loantype = True Then
            Dim Reportloantype As New rptparamloantype
            CRViewer1.ReportSource = Reportloantype
            frmParamReports.Caption = "LIST OF LOAN TYPE"
            loantype = False
        End If
         If receipts = True Then
            Dim Reportreceipts As New rptparamreceipts
            CRViewer1.ReportSource = Reportreceipts
            frmParamReports.Caption = "LIST OF RECEIPTS"
            receipts = False
        End If
        '=====

        If employers = True Then
            Dim Reportemployer As New rptparamemployers
            CRViewer1.ReportSource = Reportemployer
            frmParamReports.Caption = "LIST  OF EMPLOYERS"
            employers = False
        End If

        If tittles = True Then
            Dim Reporttittles As New rptparamtittles
            CRViewer1.ReportSource = Reporttittles
            frmParamReports.Caption = "LIST OF TITTLES"
            tittles = False
        End If

        If cities = True Then
            Dim Reportcity As New rptparamcity
            CRViewer1.ReportSource = Reportcity
            frmParamReports.Caption = "LIST OF CITIES"
            cities = False
        End If

        If accperiods = True Then
            Dim Reportaccountperiods As New rptparamaccountperiods
            CRViewer1.ReportSource = Reportaccountperiods
            frmParamReports.Caption = "LIST OF ACCOUNT PERIODS"
            accperiods = False
        End If

        If countries = True Then
            Dim Reportcountries As New rptparamcountries
            CRViewer1.ReportSource = Reportcountries
            frmParamReports.Caption = "LIST OF COUNTRIES"
            countries = False
        End If

        If currencies = True Then
            Dim Reportcurrency As New rptparamcurrency
            CRViewer1.ReportSource = Reportcurrency
            frmParamReports.Caption = "LIST OF CURRENCY DETAILS"
            currencies = False
        End If
            If feesservices = True Then
            Dim Reportfeesservices As New rptparamfeesservices
            CRViewer1.ReportSource = Reportfeesservices
            frmParamReports.Caption = "LIST OF FEES SERVICES"
            feesservices = False
            End If
            If payinterval = True Then
            Dim Reportpayintervals As New rptparampayintervals
            CRViewer1.ReportSource = Reporttittles
            frmParamReports.Caption = "LIST OF PAYMENT INTERVALS"
            payinterval = False
        End If
        If agents = True Then
            Dim Reportagents As New rptparamagents
            CRViewer1.ReportSource = Reportagents
            frmParamReports.Caption = "LIST OF AGENTS"
            agents = False
        End If

        If agentspay = True Then
            Dim Reportagentspay As New rptparamagentspay

            CRViewer1.ReportSource = Reportagentspay
            frmParamReports.Caption = "LIST OF AGENTS PAY"
            agentspay = False
        End If

        If paymethods = True Then
            Dim Reportpaymethod As New rptparampaymethods
            CRViewer1.ReportSource = Reportpaymethod
            frmParamReports.Caption = "LIST OF PAYMENT METHODS"
            paymethods = False
        End If
        If taxes = True Then
            Dim Reporttaxes As New rptparamtaxes
            CRViewer1.ReportSource = Reporttaxes
            frmParamReports.Caption = "LIST OF TAXES"
            taxes = False
        End If
        '===
        If companydepts = True Then
            Dim Reportcompanydepts As New rptparamcompanydepts
            CRViewer1.ReportSource = Reportcompanydepts
            frmParamReports.Caption = "LIST  OF COMPANY DEPARTMENTS"
            companydepts = False
        End If
        If companybranch = True Then
            Dim Reportcompanybranch As New rptparamcompanybranch
            CRViewer1.ReportSource = Reportcompanybranch
            frmParamReports.Caption = "LIST  OF COMPANY BRANCHES"
            companybranch = False
        End If
        If companymaster = True Then
            Dim Reportcompanymaster As New rptparamcompanymaster
            CRViewer1.ReportSource = Reportcompanymaster
            frmParamReports.Caption = "LIST  OF COMPANY MASTER"
            companymaster = False
        End If
        If moffice = True Then
            Dim Reportmoffice As New rptparammoffice
            CRViewer1.ReportSource = Reportmoffice
            frmParamReports.Caption = "LIST  OF M OFFICE"
            moffice = False
        End If
        If enquiry = True Then
            Dim Reportenquiry As New rptparamenquiry
            CRViewer1.ReportSource = Reportenquiry
            frmParamReports.Caption = "LIST  OF ENQUIRY"
            enquiry = False
        End If
        If empmaster = True Then
            Dim Reportempmaster As New rptparamempmaster
            CRViewer1.ReportSource = Reportempmaster
            frmParamReports.Caption = "LIST  OF EMPLOYERS MASTER"
            empmaster = False
        End If
CRViewer1.ViewReport
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub
