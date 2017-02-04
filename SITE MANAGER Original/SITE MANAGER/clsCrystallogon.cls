VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrystallogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public CrxRep As CRAXDRT.Report


Function setCRLoginInfo()
Dim crxDatabase As CRAXDRT.Database
Dim crxDatabaseTables As CRAXDRT.DatabaseTables
Dim crxDatabaseTable As CRAXDRT.DatabaseTable

Dim crxSubreport As CRAXDRT.Report
Dim crxSubreportObject As SubreportObject
Dim crxSections As CRAXDRT.Sections
Dim crxSection As CRAXDRT.Section
Dim CRXReportObject As Object

Dim strServerName, strDatabaseName, strUserName, strPassword
    
    strServerName = DSN
    strDatabaseName = Database
    strUserName = Uid
    strPassword = Pwd

 'Set your Database object to the Report object's Database object
  Set crxDatabase = CrxRep.Database
  
  'Set your DatabaseTables object to the Database object's Tables object
  Set crxDatabaseTables = crxDatabase.Tables
  
 'Loop through each DatabaseTable object in the DatabaseTables collection and then set the location
    'of the database file for each table
    
    For Each crxDatabaseTable In crxDatabaseTables
        'If you are using physical path of the access database file (Native connection to PC database)
        'crxDatabaseTable.Location = App.Path & "\xtremelite.mdb"

        'Use next line, if you are using Native connection to SQL database
        'crxDatabaseTable.SetLogOnInfo strServerName, strDatabaseName, strUserName, strPassword

        'Use next line, if you are using ODBC connection to a PC or SQL database
        crxDatabaseTable.SetLogOnInfo strServerName, strDatabaseName, strUserName, strPassword
    Next crxDatabaseTable
  

    'set up logon info for the various subreports, if any
     Set crxSections = CrxRep.Sections

    For Each crxSection In crxSections
        For Each CRXReportObject In crxSection.ReportObjects
           If CRXReportObject.Kind = crSubreportObject Then
'             Set crxSubreportObject = CRXReportObject
'             Set crxSubreport = crxSubreportObject.OpenSubreport
'             Set crxDatabase = crxSubreport.Database
'             Set crxDatabaseTables = crxDatabase.Tables
'
'             For Each crxDatabaseTable In crxDatabaseTables
'                'If you are using a  Native connection to PC database
'                'crxDatabaseTable.Location = App.Path & "\xtremelite.mdb"
'                'Use next line, if you are using Native connection to SQL database
'                'crxDatabaseTable.SetLogOnInfo "servername", "databasename", "userid", "password"
'                'Use next line, if you are using ODBC connection to a PC or SQL database
'                'crxDatabaseTable.SetLogOnInfo strServerName, strDatabaseName, strUserName, strPassword
'             Next
'
'             Set crxSubreport = Nothing
           End If
        Next
    Next



End Function
