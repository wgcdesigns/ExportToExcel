@@ -0,0 +1,857 @@
﻿' History of changes
' 02/12/2018 ns
'   - Updated InsertWorksheet to stop corrupting excel sheet when adding additional worksheets
' 11/15/2017 ns
'   - Added version of script for AGENCY BRW with alternate columns and locations
'   - Includes POD Report to BRW in Excel Format with 4 sheets with different reports, DoPODValidationReport Function
' 10/05/2017 ns 
'   - Copy of ImageDBCrossChecker project to use as a custom reporting tool.
'   - Sends ADMIN Reports of Daily Document Uploads and emails results to TNG ADMIN


Option Strict On

Imports System.IO
Imports GdPicture9
Imports System.Text
Imports System.Data.SqlClient
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.ExtendedProperties

''' <summary>
''' Start-up object for
''' </summary>
''' <remarks></remarks>
Public Module StartUp

    Private strAppName As String = My.Application.Info.AssemblyName & " - Version:" & My.Application.Info.Version.ToString
    Private m_LogDirName As String = My.Application.Info.DirectoryPath & "\MARC10_IMAGE_DB_CROSSCHECKER\LOGS"
    Private Const ARG_NAME_PAGE_CHECK As String = "/page_check"
    Private Const ARG_NAME_MAX_ROWS As String = "/max_rows="
    Private Const ARG_NAME_CHK_BY_DEALER As String = "/chk_by_cust_docnum"
    Private Const ARG_NAME_CHK_REPORT_PODVAL As String = "/pod_validation"
    Private Const ARG_NAME_SKIP_CHK_BY_INDEX1 As String = "/skip_chk_by_index1"

    ' TITLES OF MICROSOFT EXCEL SPREADSHEET REPORT SEND TO TNG --------------------------
    Private Const MARC10DOCUMENT_BRW As String = "MARC10PODVAL_{0:D4}_{1:D2}_{2:D2}.xlsx"

    ' --------------------------------------------------------------------------
    ' Variables set by command-line arguments
    ' -------------------------------------------------------------------------
    ' The number of rows from 'Document' table to check. 0=No limit;  Intended use is mostly testing
    Private g_MaximumNumberOfRowsToCheck As Integer = 0

    ' The number of days back from current date to use to select rows in 'Document' table whose images are to validated as still present.
    Private g_NumberOfDaysInPastToCheck As Integer = 1

    ' Controls the maximum number of missing image files to be logged (all will be counted); value of 0 will allow all image files to be logged.
    Private g_MaximumErrorMessagesToLog As Int32 = 0

    ' Controls if ImageDbCrossChecker should compare the number of pages in the PDF file with the value in the column 'page_count' in the document
    ' and update it if the 'document' table is incorrect.
    Private g_Do_Page_Count_Check As Boolean = False

    Private g_Do_Report_PODVAL_Check As Boolean = False
    ' Cross check  presence of images by columns 'index1' and 'index2'
    Private g_chk_by_Index1 As Boolean = True

    ' Cross check  presence of images by columns 'dealer' and 'document_number'
    Private g_chk_by_dealer As Boolean = False

    Private Enum eDocXcheckType As Integer
        Dealer = 0
        Index1 = 1
    End Enum


    Public Sub Main()

        'initialize all the file directories needed
        If InitializeService() = False Then Exit Sub

        'Try
        ' If the command-line arguments are valid 
        Dim Message As String = String.Empty
        If (HandleArgs(Message) = True) Then
            ' -----------------------------------------------------------------------------------------------------------------
            ' Use the information in the row of the Document to be sure we can find the document image file in file system.
            ' -----------------------------------------------------------------------------------------------------------------
            Dim DocXcheckType As Integer = eDocXcheckType.Index1

            ' Check by index1 column
            If (g_chk_by_Index1 = True) Then
                DocXcheckType = eDocXcheckType.Index1
                DoImageDbCrossCheck(g_NumberOfDaysInPastToCheck, g_MaximumNumberOfRowsToCheck, DocXcheckType)
            End If

            If (g_Do_Report_PODVAL_Check = True) Then
                DocXcheckType = eDocXcheckType.Dealer
                DoPODValidationReport(g_NumberOfDaysInPastToCheck, g_MaximumNumberOfRowsToCheck, DocXcheckType)
            End If
        Else
            ' Command line arguments were not correct - log a message and send the alert
            LogMessage(Message, System.Diagnostics.TraceEventType.Error)
        End If

    End Sub

    Private Function InitializeService() As Boolean

        Try

            ' Perform standard start-up procedure for shared routines
            MainStartup()

            ' Initialize standard Marc 10 logging for this program
            InitializeLogging()

            ' Initialize imaging-specific objects.
            ModImaging.ImagingInit()

            Dim Msg As String = My.Application.Info.AssemblyName & " - Version:" & My.Application.Info.Version.ToString & " started."

            ' Append arguments, if any
            If (My.Application.CommandLineArgs.Count > 0) Then
                Msg &= " Arguments: "
                For Each arg As String In My.Application.CommandLineArgs
                    Msg &= " " & arg
                Next
            End If

            LogMessage(Msg, TraceEventType.Start)


            Return True

        Catch ex As Exception
            ' This will be Log Error message
            LogMessage("Error Initializing " & strAppName & vbCrLf & vbCrLf & ex.Message & vbCrLf & vbCrLf & ex.StackTrace, System.Diagnostics.TraceEventType.Error)
            Return False

        End Try


    End Function

    Private Sub InitializeLogging()

        ' -------------------------------------------------------------------------------
        ' Create Log directory if needed and initialize the logging system
        ' -------------------------------------------------------------------------------
        If (Directory.Exists(m_LogDirName) = False) Then
            Directory.CreateDirectory(m_LogDirName)
        End If

        ' Initialize the logging for this program
        InitializeLogFile(My.Application.Log, "ImageDbCrossChecker", m_LogDirName)

    End Sub


    ''' <summary>
    ''' Less crude method to parse command line arguments
    ''' </summary>
    ''' <param name="LogMessage"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function HandleArgs(ByRef LogMessage As String) As Boolean
        Dim RetVal As Boolean = True
        Dim idxLastArg As Integer = My.Application.CommandLineArgs.Count - 1
        Dim ThisArg As String = String.Empty
        Dim idx As Integer = 0
        Dim ThisArgValue As String

        For idx = 0 To idxLastArg

            ThisArg = My.Application.CommandLineArgs(idx).ToLower
            ThisArgValue = String.Empty

            If (ThisArg.IndexOf("/daysback=") = 0) Then
                ' ----------------------------------------------------------------------------------------
                ' Check the parameter for number of days back to use to select documents to validate
                ' ----------------------------------------------------------------------------------------
                If (ThisArg.Length - ThisArg.IndexOf("=") > 1) Then
                    ThisArgValue = ThisArg.Substring(ThisArg.IndexOf("=") + 1).Trim
                    If (IsNumeric(ThisArgValue) = True) Then
                        g_NumberOfDaysInPastToCheck = CInt((ThisArgValue))
                    Else
                        RetVal = False
                    End If
                Else
                    RetVal = False
                End If

            ElseIf (ThisArg.IndexOf("/errors=") = 0) Then
                ' ----------------------------------------------------------------------------------------
                ' Check the parameter to control the number of errors to log
                ' ----------------------------------------------------------------------------------------
                If (ThisArg.Length - ThisArg.IndexOf("=") > 1) Then
                    ThisArgValue = ThisArg.Substring(ThisArg.IndexOf("=") + 1).Trim
                    If (IsNumeric(ThisArgValue) = True) Then
                        g_MaximumErrorMessagesToLog = CInt((ThisArgValue))
                    Else
                        RetVal = False
                    End If
                Else
                    RetVal = False
                End If

            ElseIf (ThisArg.IndexOf(ARG_NAME_MAX_ROWS) = 0) Then
                ' ----------------------------------------------------------------------------------------
                ' Set the global value to control the maximum number of rows to check
                ' ----------------------------------------------------------------------------------------
                If (ThisArg.Length - ThisArg.IndexOf("=") > 1) Then
                    ThisArgValue = ThisArg.Substring(ThisArg.IndexOf("=") + 1).Trim
                    If (IsNumeric(ThisArgValue) = True) Then
                        g_MaximumNumberOfRowsToCheck = CInt((ThisArgValue))
                    Else
                        RetVal = False
                    End If
                Else
                    RetVal = False
                End If

            ElseIf (ThisArg = ARG_NAME_PAGE_CHECK) Then
                ' Set the global value to invoke the page count check method
                g_Do_Page_Count_Check = True

            ElseIf (ThisArg = ARG_NAME_CHK_REPORT_PODVAL) Then
                ' Set the global value to invoke the page count check method
                g_Do_Report_PODVAL_Check = True

            ElseIf (ThisArg = ARG_NAME_CHK_BY_DEALER) Then
                ' Set the global value to invoke cross-check by customer number
                g_chk_by_dealer = True

            ElseIf (ThisArg = ARG_NAME_SKIP_CHK_BY_INDEX1) Then
                ' Set the global value to skip cross-check by index1
                g_chk_by_Index1 = False

            Else

                RetVal = False
            End If


        Next
        LogMessage = ""

        If (RetVal = False) Then
            LogMessage &= "Error parsing command line - " & vbCrLf & "Command line arguments must be as follows: '" & My.Application.Info.AssemblyName & "  [ /daysback=<Days from current day to cross-check>  /errors=<Maximum number of errors to log> " & ARG_NAME_PAGE_CHECK & " ]' " & vbCrLf & _
                "Note: the arguments are optional. If no arguments are specified, all documents are checked and all errors are logged; checking of document.page_count against pdf file is not performed be default. " & vbCrLf & _
                "Example:  '" & My.Application.Info.AssemblyName & "  /daysback=100  /errors=2000 " & ARG_NAME_PAGE_CHECK & "'"
        End If

        Return RetVal
    End Function


    Private Sub DoImageDbCrossCheck(ByVal NumberOfDaysInPastToCheck As Integer, ByVal MaximumNumberOfRowsToCheck As Integer, ByVal DocXcheckType As Integer)
        Dim Message As String = String.Empty

        Dim emailSubject As String = Nothing
        Dim emailtitlelines As String = Nothing
        Dim emailbody As String = Nothing
        Dim strEmailFrom As String = RetrieveValueByContext("TECH_SUPPORT_EMAIL")
        Dim emailUser As String = RetrieveValueByContext("TECH_SUPPORT_USER")
        Dim emailpassword As String = RetrieveValueByContext("TECH_SUPPORT_PASS")
        Dim emailcc As String = ""
        Dim strEmailTo As String = RetrieveValueByContext("Batch_Upload_Verification_Report_TO")
        Dim emailbcc As String = RetrieveValueByContext("Batch_Upload_Verification_Report_BCC")
        emailSubject = RetrieveValueByContext("Batch_Upload_Verification_Report_SUBJECT")

        ' ==============================================================================================

        Try

            Message = "Beginning Batch Upload Verification Report to MARC10 - "
            LogMessage(Message)

            ' ------------------------------------------------------------------------------
            ' Construct mail message to be sent
            ' ------------------------------------------------------------------------------
            Dim Status_Suffix As String = String.Empty
            Dim strSubTitle As String = Nothing
            Dim istring As String = Nothing
            Dim istringonly4 As String = Nothing

            ' BUILD BATCH PART OF REPORT WITH LIVE VS BATCH DOC TOTALS
            ' SEARCH WITH vw_document_upload_verification
            Dim strSQL As String = "SELECT document_batch_id,upload_date,batch_dir,machine_name,document_count,[LiveCount] " & _
            "FROM [venTNG].[dbo].[vw_document_upload_verification]  "

            Dim strSQLMach4Totals As String = "SELECT sum(document_count) " & _
            "FROM [venTNG].[dbo].[vw_document_upload_verification]  "

            Dim strSQLTotals As String = "SELECT sum(LiveCount) " & _
            "FROM [venTNG].[dbo].[vw_document_upload_verification]  "

            Dim iDateFrom As Date = CDate(Now().AddDays(-1).ToString("d"))
            Dim iDateTo As Date = CDate(iDateFrom.ToString("d"))
            Dim SqlParameter As SqlParameter
            Dim strWhere As String = ""
            Dim strWhereOnlyMach4 As String = ""
            Dim myCmd As New SqlCommand

            ' ----------------------------------------------------------------
            Dim dtFrom As Date = Convert.ToDateTime(iDateFrom)
            Dim dtTo As Date = DateAdd(DateInterval.Day, 1, Convert.ToDateTime(iDateTo))
            strWhere &= "(upload_date >= '" & dtFrom & "' AND upload_date < '" & dtTo & "') "
            strWhereOnlyMach4 &= "(upload_date >= '" & dtFrom & "' AND upload_date < '" & dtTo & "') "
            If strWhereOnlyMach4 <> "" Then strWhereOnlyMach4 &= " AND "
            strWhereOnlyMach4 &= " ((machine_id <> '20' AND machine_id <> '2' AND LiveCount <> '0') and machine_id = '4')  "

            If strWhere <> "" Then strWhere &= " AND "
            strWhere &= " ((machine_id <> '20' AND machine_id <> '2' AND LiveCount <> '0') and machine_id <> '4')   "

            'Add the Sql Parameters to complete the construction of the WHERE clause
            SqlParameter = New SqlParameter("from", SqlDbType.DateTime2)
            SqlParameter.Value = dtFrom
            myCmd.Parameters.Add(SqlParameter)

            SqlParameter = New SqlParameter("to", SqlDbType.DateTime2)
            SqlParameter.Value = dtTo
            myCmd.Parameters.Add(SqlParameter)

            strSubTitle &= "Date Range " & iDateFrom & " to " & iDateTo

            ' ----------------------------------------------------------------

            Dim strGroupBy As String = ""
            Dim strOrderBy As String = " ORDER BY document_batch_id ASC "

            ' SHOW SEARCH RESULTS UNDER MISSING FILE MESSAGE WITH TOTALS OF DAILY BATCHES AND PAGE TOTALS

            ' ALSO SHOW FAILED BATCHES WITH 0 DOCUMENT COUNT TOTALS
            If strWhere <> "" Then strWhere = "WHERE " & strWhere
            If strWhereOnlyMach4 <> "" Then strWhereOnlyMach4 = "WHERE " & strWhereOnlyMach4

            'Supply a DataTable corresponding to each report dataset
            'The dataset name must match the name defined in the RDLC report
            myCmd.CommandText = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;" & strSQL & strWhere & strGroupBy & strOrderBy
            Dim myDT As DataTable = GetDataTable(myCmd)

            Dim ilivecountmach4 As String = ""
            Dim ilivecountmach4_TEMP As Object = GetValue(strSQLMach4Totals & strWhere)
            If (ilivecountmach4_TEMP IsNot DBNull.Value) Then
                ilivecountmach4 = CStr(GetValue(strSQLMach4Totals & strWhere))
            Else
                ilivecountmach4 = "0"
            End If

            Dim idocument_countmach4 As String = ""
            Dim idocument_countmach4_TEMP As Object = GetValue(strSQLMach4Totals & strWhere)
            If (idocument_countmach4_TEMP IsNot DBNull.Value) Then
                idocument_countmach4 = CStr(GetValue(strSQLMach4Totals & strWhere))
            Else
                idocument_countmach4 = "0"
            End If

            Dim ilivecountnot4 As String = ""
            Dim ilivecountnot4_TEMP As Object = GetValue(strSQLTotals & strWhereOnlyMach4)
            If (ilivecountnot4_TEMP IsNot DBNull.Value) Then
                ilivecountnot4 = CStr(GetValue(strSQLTotals & strWhereOnlyMach4))
            Else
                ilivecountnot4 = "0"
            End If

            Dim idocument_countnot4 As String = ""
            Dim idocument_countnot4_TEMP As Object = GetValue(strSQLMach4Totals & strWhereOnlyMach4)
            If (idocument_countnot4_TEMP IsNot DBNull.Value) Then
                idocument_countnot4 = CStr(GetValue(strSQLMach4Totals & strWhereOnlyMach4))
            Else
                idocument_countnot4 = "0"
            End If
            istring = BuildReportHTML(myDT)
            istring = "<br /><h2>Batch Details<h2><hr /><h3>Recent Results from Batch Uploads <br />---------- Daily Totals ---------- DOCUMENT TOTAL: " & CInt(idocument_countmach4).ToString("n0") & " ---------- LIVE COUNT: " & CInt(ilivecountmach4).ToString("n0") & "</h3>" & istring

            myCmd.CommandText = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;" & strSQL & strWhereOnlyMach4 & strGroupBy & strOrderBy
            myDT = GetDataTable(myCmd)

            istringonly4 = BuildReportHTML(myDT)
            istringonly4 = "<br /><h3>Recent Results from Agentek / Mobileframe Uploads <br />---------- Daily Totals ---------- DOCUMENT TOTAL: " & CInt(idocument_countnot4).ToString("n0") & " ---------- LIVE COUNT: " & CInt(ilivecountnot4).ToString("n0") & "</h3>" & istringonly4

            emailtitlelines = "<h1>" & emailSubject & "</h1><h2>Summary:</h2><hr />"
            emailtitlelines += "" & CInt(idocument_countmach4) + CInt(idocument_countnot4) & " Total Documents Uploaded on " & iDateFrom & "<br />"
            emailtitlelines += "<h2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
            emailtitlelines += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
            emailtitlelines += "Batch Uploads: " & CInt(idocument_countmach4).ToString("n0") & " --- LIVE COUNT: " & CInt(ilivecountmach4).ToString("n0") & ""

            If (CInt(idocument_countmach4) = CInt(ilivecountmach4)) Then
                emailtitlelines += " --> 0 MISSING<br />"
            Else
                emailtitlelines += " ***** COUNT MISS-MATCHED!<br />"
            End If

            emailtitlelines += "Agentek / Mobileframe Uploads: " & CInt(idocument_countnot4).ToString("n0") & " --- LIVE COUNT: " & CInt(ilivecountnot4).ToString("n0") & ""
            If (CInt(idocument_countnot4) = CInt(ilivecountnot4)) Then
                emailtitlelines += " --> 0 MISSING<br />"
            Else
                emailtitlelines += " ***** COUNT MISS-MATCHED!<br />"
            End If

            emailtitlelines += "</h3>"
            'EmailTechSupport("Cross check of Document list and image files - " & Status_Suffix, Message & " " & istring)
            SendEmail(emailSubject, emailtitlelines & emailbody & istring & istringonly4, strEmailTo, strEmailFrom, emailUser, emailpassword, emailcc, emailbcc, , , True, )
        Catch ex As Exception
            Dim ExceptionType As Integer = Type.GetTypeCode(ex.GetType())
            LogMessage(strAppName & "  - Failure : " & Message & "  Error:'" & ex.Message & "'  Error Type:'" & CStr(ExceptionType) & "'  Trace:" & ex.StackTrace, TraceEventType.Error)

        End Try
        Message = "Report - Batch Upload Verification Report to MARC10 - SENT SUCCESSFULLY"
        LogMessage(Message)

        Message = "DoPODValidationReport NEXT ....... "
        LogMessage(Message)
    End Sub

    Private Sub DoPODValidationReport(ByVal NumberOfDaysInPastToCheck As Integer, ByVal MaximumNumberOfRowsToCheck As Integer, ByVal DocXcheckType As Integer)
        Dim Message As String = String.Empty

        Dim emailSubject As String = Nothing
        Dim emailtitlelines As String = Nothing
        Dim emailbody As String = Nothing
        Dim emailcc As String = ""
        Dim strEmailTo As String = RetrieveValueByContext("POD_Validation_Report_TO")
        Dim emailbcc As String = RetrieveValueByContext("POD_Validation_Report_BCC")
        Dim strEmailFrom As String = RetrieveValueByContext("TECH_SUPPORT_EMAIL")
        Dim emailUser As String = RetrieveValueByContext("TECH_SUPPORT_USER")
        Dim emailpassword As String = RetrieveValueByContext("TECH_SUPPORT_PASS")

        ' ==================================================================================================================
        ' DO POD VALIDATION REPORT =========================================================================================
        ' ==================================================================================================================
        Message = "Begin DoPODValidationReport... "
        LogMessage(Message)
        ' -----------------------------------------------------------------------------------------------------
        ' Format and execute the query using a DataReader
        ' ------------------------------------------------------------------------------------------------------
        Dim Ct_TotalCrossChecked As Integer = 0
        Dim Ct_CrossChecked_OK As Integer = 0
        Dim Ct_CrossCheck_Failed As Integer = 0
        Dim Ct_PageCountCrossCheck_Failed As Integer = 0
        Dim Ct_PDF_Files_Opened As Integer = 0
        Dim Ct_FailedToOpen As Integer = 0
        Dim Ct_OnePageDocs As Integer = 0
        Dim Ct_PageCountCorrections As Integer = 0
        Dim Maximum_Pages_In_Doc = 0
        Dim SqlSelectDocsToCrossCheck As String = String.Empty
        Dim Sql_WHERE_Clause As String = " "
        Dim StoppedLoggingErrors As Boolean = False
        Dim TopClause As String = String.Empty
        Dim iExcelFileLoc As String = "" ' BASE PATH FOR REPORT
        Dim ExportFileName As String = Nothing
        Dim dOptionalReportDate As Date = Nothing

        Const SQL_SELECT_DOCUMENTS_TO_FOR_DUPLICATES As String = _
        " SELECT        legacy_dealer, document_number, vendorid, document_date, COUNT(*) AS count" & _
        " FROM            [document]" & _
        " WHERE        (vendorid <> '0') OR" & _
        "                          (vendorid <> '')" & _
        " GROUP BY legacy_dealer, document_number, vendorid, document_date" & _
        " HAVING        (COUNT(*) > 1) AND (document_date >= CONVERT(DATETIME, '2017-05-01 00:00:00', 102))"

        ' CHECK FOR DOCUMENTS IN DOCUMENT BUT NOT IN DOC NUMBERS
        Const SQL_SELECT_DOCUMENTS_TO_FOR_DOCS_NOT_FOUND_IN_DOCNUMS As String = _
        "SELECT        TOP (10000) transaction_id,  ds_location_id, document_date, legacy_dealer, document_number" & _
        " FROM            document_numbers WHERE        (NOT EXISTS" & _
        " (SELECT        ERPInvoiceID, Document_Number, ds_location_id, Client_id, transaction_id, legacy_dealer" & _
        " FROM vw_document_numberdetail_search" & _
        " WHERE        (Document_Number = document_numbers.document_number))) AND (ds_location_id = 292904)"
        ' SHOW TOTAL DOC DETAIL COUNT IN DOCS_CAN FILE

        ' ------ ADD SQL TO CHECK DOCUMENT_NUMBER_DETAILS FOR DOCNUMS NOT FOUND IN DOCS
        Const SQL_SELECT_DOCUMENTS_TO_FOR_DOCS_NOT_FOUND As String = _
        " SELECT        TOP (100) ERPInvoiceID, Document_Number, legacy_dealer, Client_id AS VendorIDCheck" & _
        " FROM            vw_document_numberdetail_search WHERE        (NOT EXISTS" & _
        " (SELECT        document_id, document_type_id, index1, index2, ds_id, document_date, image_file_name, machine_id, created, created_by, modified, modified_by, " & _
        " verified, update_sent, document_scan_type_id, pending_review, document_batch_id, page_count, legacy_dealer, document_number, is_dsd, " & _
        " is_not_pod_por, vendorid FROM            [document]" & _
        " WHERE        (index2 = vw_document_numberdetail_search.Document_Number)))"
        ' SHOW # OF MISSING SCANNED DOCUMENTS

        ' ---- IF FOUND IN DOC_NUM_DETAILS THEN CONFIRM VENDORID
        Const SQL_SELECT_DOCUMENTS_TO_FOR_INVALID_VENDORIDS As String = _
        "SELECT [document].document_id,  " & _
        " vw_document_numberdetail_search.Document_Number, [document].document_number AS Ddocument_number, " & _
        " [document].legacy_dealer AS Dlegacy_dealer, vw_document_numberdetail_search.legacy_dealer, vw_document_numberdetail_search.ds_location_id," & _
        " vw_document_numberdetail_search.Client_id, [document].vendorid,  [document].document_date" & _
        " FROM            [document] INNER JOIN" & _
        " vw_document_numberdetail_search " & _
        " ON [document].vendorid <> vw_document_numberdetail_search.Client_id " & _
        " AND ([document].document_number = vw_document_numberdetail_search.document_number " & _
        " AND ([document].legacy_dealer = vw_document_numberdetail_search.legacy_dealer ))"
        ' ==============================================================================================

        ' ======================================================================
        ' SETUP FILE NAME TO EXPORT REPORT DETAILS - EXCEL FILE WITH MULTPLE SHEETS
        ' ======================================================================
        dOptionalReportDate = Today()
        ExportFileName = String.Format(MARC10DOCUMENT_BRW, dOptionalReportDate.Year, dOptionalReportDate.Month, dOptionalReportDate.Day)
        Dim iFullExportFilePath As String = iExcelFileLoc & ExportFileName
        ' ======================================================================
        ' REPORT TITLE AND INTRO ===============================================
        ' ======================================================================
        emailSubject = "POD Validation Report to MARC10"
        emailtitlelines = "<h1>" & emailSubject & "</h1><h2>Results from " & Today.ToString() & ", "
        emailtitlelines += "are attached in Excel Spreadsheet.</h2><hr /> "
        ' ======================================================================
        emailtitlelines += "Duplicate Document Checker:<br />"
        Dim iAGetTable As DataTable = GetDataTable(SQL_SELECT_DOCUMENTS_TO_FOR_DUPLICATES)
        Dim iReportA As String = BuildReportHTML(iAGetTable)


        Try
            CreateExcelFileFromDataTable(iExcelFileLoc & ExportFileName, iAGetTable)
        Catch ex As Exception
            Dim ExceptionType As Integer = Type.GetTypeCode(ex.GetType())
            LogMessage(strAppName & "  - Failure : " & iExcelFileLoc & ExportFileName & "  Error:'" & ex.Message & "'  Error Type:'" & CStr(ExceptionType) & "'  Trace:" & ex.StackTrace, TraceEventType.Error)

        End Try

        emailtitlelines += iReportA & "<br /><hr /><br />Additional POD Validation Report Details found in Excel Spreadsheet for:<br />"
        ' ======================================================================
        emailtitlelines += "Missing Scanned Documents Checker:<br />"
        Dim iBGetTable As DataTable = GetDataTable(SQL_SELECT_DOCUMENTS_TO_FOR_DOCS_NOT_FOUND_IN_DOCNUMS)
        Dim iReportB As String = BuildReportHTML(iBGetTable)
        InsertWorksheet(iExcelFileLoc & ExportFileName, iBGetTable, "Missing Scanned Documents", 2)
        'emailtitlelines += iReportB & "<br /><hr />"
        ' ======================================================================

        emailtitlelines += "Missing Document Detail Checker:<br />"
        Dim iCGetTable As DataTable = GetDataTable(SQL_SELECT_DOCUMENTS_TO_FOR_DOCS_NOT_FOUND)
        Dim iReportC As String = BuildReportHTML(iCGetTable)
        InsertWorksheet(iExcelFileLoc & ExportFileName, iCGetTable, "Missing Document Detail", 3)
        'emailtitlelines += iReportC & "<br /><hr />"
        ' ======================================================================

        emailtitlelines += "Invalid Vendor ID Checker:<br />"
        Dim iDGetTable As DataTable = GetDataTable(SQL_SELECT_DOCUMENTS_TO_FOR_INVALID_VENDORIDS)
        Dim iReportD As String = BuildReportHTML(iDGetTable)
        InsertWorksheet(iExcelFileLoc & ExportFileName, iDGetTable, "Invalid Vendor ID", 4)
        'emailtitlelines += iReportD & "<br /><hr />"
        ' ======================================================================
        ' END REPORT - SEND ====================================================
        ' ======================================================================

        Dim arrayList As New System.Collections.ArrayList()
        'For Each filename As String In txtFiles
        arrayList.Add(iExcelFileLoc & ExportFileName)
        'Next

        Message = "DoPODValidationReport Report Built - NOW SENDEMAIL"
        LogMessage(Message)

        Try
            SendEmail(emailSubject, emailtitlelines, strEmailTo, strEmailFrom, emailUser, emailpassword, emailcc, emailbcc, , , True, arrayList)
        Catch ex As Exception
            Dim ExceptionType As Integer = Type.GetTypeCode(ex.GetType())
            LogMessage(strAppName & "  - Failure : " & Message & "  Error:'" & ex.Message & "'  Error Type:'" & CStr(ExceptionType) & "'  Trace:" & ex.StackTrace, TraceEventType.Error)

        End Try
        ' ============= ***** END REPORT ***** =============================================================================

        Message = "DoPODValidationReport COMPLETE!"
        LogMessage(Message)
    End Sub

    ' Given a document name, inserts a new worksheet.
    Public Sub InsertWorksheet(ByVal docName As String, ByVal SQL As DataTable, ByVal sheetName As String, ByVal intSheetId As Integer)
        'Dim sheetName As String
        Dim fileName As String = docName
        ' Open an existing spreadsheet document for editing.
        Dim spreadSheet As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, True)
        Using (spreadSheet)
            ' Add a blank WorksheetPart.
            Dim newWorksheetPart As WorksheetPart = spreadSheet.WorkbookPart.AddNewPart(Of WorksheetPart)()
            newWorksheetPart.Worksheet = New Worksheet(New SheetData())

            ' Create a Sheets object.
            Dim sheets As Sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild(Of Sheets)()
            Dim relationshipId As String = spreadSheet.WorkbookPart.GetIdOfPart(newWorksheetPart)

            ' Get a unique ID for the new worksheet.
            Dim sheetId As UInteger = 1
            If (sheets.Elements(Of Sheet).Count > 0) Then
                sheetId = CUInt(sheets.Elements(Of Sheet).Select(Function(s) s.SheetId.Value).Max + 1)
            End If

            ' Append the new worksheet and associate it with the workbook.
            Dim sheet As Sheet = New Sheet
            sheet.Id = relationshipId
            sheet.SheetId = sheetId
            sheet.Name = sheetName
            sheets.Append(sheet)

            'get the sheetData object so we can add the data table to it
            Dim sheetData As SheetData = newWorksheetPart.Worksheet.GetFirstChild(Of SheetData)()

            'add the data table
            AddDataTable(SQL, sheetData)

            'save the workbook
            newWorksheetPart.Worksheet.Save()

            ' Close the document.
            spreadSheet.Close()

        End Using

    End Sub

    Public Function FindTotalDocsinDB(ByVal iTotalDocsFound As String) As String

        Dim ScriptLine As String = ""
        Dim iStatusMessg As String = Nothing
        ' ---------------------------------

        Using reader As New StreamReader(iTotalDocsFound)
            While Not reader.EndOfStream
                Dim line As String = reader.ReadLine()
                If line.Contains("Count of records cross-checked:") Then
                    ScriptLine = line
                    Exit While
                End If
            End While

            If ScriptLine <> "" Then

                'split the string by equals sign
                Dim ary As String() = ScriptLine.Split(":"c)
                Dim iEndValue As String = Nothing

                'Check the data type of the string we collected is inline with what we are expecting, e.g. numeric
                If IsNumeric(ary(5)) Then

                    'obtain the string after the equals sign
                    Dim value As String = ary(3)
                    Dim phrase As String = value
                    Console.WriteLine("Before: {0}", phrase)
                    phrase = phrase.Replace(" Count of Image files found", "")
                    Console.WriteLine("After: {0}", phrase)

                    ScriptLine = Trim(phrase)
                    'based on the value after the equals sign, do something

                End If

            End If

        End Using


        ' ---------------------------------
        Return ScriptLine
    End Function


    Public Sub CreateFileForExcel(ByVal FilePath As String, myDT As DataTable, ByVal SheetName As String, ByVal SheetNum As String, ByVal SheetStatus As String)

        Dim spreadsheetDocument As SpreadsheetDocument = spreadsheetDocument.Create(FilePath, SpreadsheetDocumentType.Workbook)

        ' Add a WorkbookPart to the document.
        Dim workbookpart As WorkbookPart = spreadsheetDocument.AddWorkbookPart
        workbookpart.Workbook = New Workbook

        ' Add a WorksheetPart to the WorkbookPart.
        Dim worksheetPart As WorksheetPart = workbookpart.AddNewPart(Of WorksheetPart)()
        worksheetPart.Worksheet = New Worksheet(New SheetData())

        ' Add Sheets to the Workbook.
        Dim sheets As Sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(Of Sheets)(New Sheets())

        ' Append a new worksheet and associate it with the workbook.
        Dim sheet As Sheet = New Sheet
        sheet.Id = SheetNum
        sheet.SheetId = 1
        sheet.Name = SheetName

        sheets.Append(sheet)

        'get the sheetData object so we can add the data table to it
        Dim sheetData As SheetData = worksheetPart.Worksheet.GetFirstChild(Of SheetData)()

        'add the data table
        AddDataTable(myDT, sheetData)

        'save the workbook
        workbookpart.Workbook.Save()

        ' Close the document.
        spreadsheetDocument.Close()

    End Sub

    Private Sub AddDataTable(ByRef exportData As DataTable, ByRef sheetdata As SheetData)

        'add column names to the first row   
        Dim Header As Row = New Row()
        Header.RowIndex = 1

        For Each col As DataColumn In exportData.Columns
            Dim headerCell As Cell = createTextCell(exportData.Columns.IndexOf(col) + 1, Convert.ToInt32(Header.RowIndex.Value), col.ColumnName)
            Header.AppendChild(headerCell)
        Next

        sheetdata.AppendChild(Header)

        'loop through each data row   
        Dim contentRow As DataRow
        Dim intStartRow As Int32 = 2

        For intLoop = 0 To exportData.Rows.Count - 1
            contentRow = exportData.Rows(intLoop)
            sheetdata.AppendChild(createContentRow(contentRow, intLoop + intStartRow))
        Next

    End Sub

    Private Function createTextCell(ByVal columnIndex As Integer, ByVal rowIndex As Integer, ByVal cellValue As Object) As Cell

        Dim cell As New Cell

        cell.DataType = CellValues.InlineString
        cell.CellReference = getColumnName(columnIndex) & rowIndex.ToString

        Dim inlineSTring As New InlineString()

        Dim t As New Text()

        t.Text = cellValue.ToString()
        inlineSTring.AppendChild(t)
        cell.AppendChild(inlineSTring)

        Return cell

    End Function

    Private Function getColumnName(ByVal columnIndex As Integer) As String

        Dim strCol As String = ""

        Select Case columnIndex
            Case 1
                strCol = "A"
            Case 2
                strCol = "B"
            Case 3
                strCol = "C"
            Case 4
                strCol = "D"
            Case 5
                strCol = "E"
            Case 6
                strCol = "F"
            Case 7
                strCol = "G"
            Case 8
                strCol = "H"
            Case 9
                strCol = "I"
            Case 10
                strCol = "J"
            Case 11
                strCol = "K"
            Case 12
                strCol = "L"
            Case 13
                strCol = "M"
            Case 14
                strCol = "N"
            Case 15
                strCol = "O"
            Case 16
                strCol = "P"
            Case 17
                strCol = "Q"
            Case 18
                strCol = "R"
            Case 19
                strCol = "S"
            Case 20
                strCol = "T"
            Case 21
                strCol = "U"
            Case 22
                strCol = "V"
            Case 23
                strCol = "W"
            Case 24
                strCol = "X"
            Case 25
                strCol = "Y"
            Case 26
                strCol = "Z"
        End Select

        Return strCol

    End Function
    Private Function createContentRow(ByVal dataRow As DataRow, ByVal rowIndex As Integer) As Row

        Dim row As New Row

        For i As Integer = 0 To dataRow.Table.Columns.Count - 1
            Dim datacell As Cell = createTextCell(i + 1, rowIndex, dataRow(i))
            row.AppendChild(datacell)
        Next

        Return row

    End Function

    Public Function BuildReportHTML(ByVal iRecordSet As DataTable) As String

        Dim dt As DataTable = iRecordSet
        Dim istring As String = Nothing
        Dim itemp As Int64 = 0
        Dim ibgcolor As String = Nothing
        'Populating a DataTable from database.

        'Building an HTML string.
        Dim html As New StringBuilder()

        'Table start.
        html.Append("<table border = '0' width='100%' border='0' cellspacing='0' cellpadding='10'>")

        'Building the Header row.
        html.Append("<tr>")
        For Each column As DataColumn In dt.Columns
            html.Append("<th>")
            html.Append(column.ColumnName)
            html.Append("</th>")
        Next
        html.Append("</tr>")

        'Building the Data rows.
        For Each row As DataRow In dt.Rows
            If itemp Mod 2 = 0 Then
                ibgcolor = "bgcolor=#FFFFFF"
            Else
                ibgcolor = "bgcolor=#eeeeee"

            End If
            html.Append("<tr " & ibgcolor & ">")
            For Each column As DataColumn In dt.Columns

                html.Append("<td >")
                html.Append(row(column.ColumnName))
                html.Append("</td>")

            Next
            html.Append("</tr>")
            itemp += 1
        Next

        'Table end.
        html.Append("</table>")

        'Append the HTML string to Placeholder.
        istring = html.ToString
        Return istring
    End Function

End Module
