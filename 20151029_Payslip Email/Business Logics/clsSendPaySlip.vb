Imports System.IO
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Net
Imports System.Net.Mail

Public Class clsSendPaySlip
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_SendPayslip, frm_Pay_SendPaySlip)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.PaneLevel = 1
        Databind(oForm)
        oForm.Freeze(False)
    End Sub


    Private Sub AddChooseFromList_Conditions(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCombobox = objForm.Items.Item("9").Specific
            oCFL = oCFLs.Item("CFL_2")
            oCons = oCFL.GetConditions()
            oCon = oCons.Item(0)
            oCon.Alias = "U_Z_CompNo"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = oCombobox.Selected.Value
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_3")
            oCons = oCFL.GetConditions()
            oCon = oCons.Item(0)
            oCon.Alias = "U_Z_CompNo"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = oCombobox.Selected.Value
            oCFL.SetConditions(oCons)
            ' oCon = oCons.Add
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#Region "Add Choose From List"

    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFL = oCFLs.Item("CFL_2")

            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.BracketOpenNum = 2
            oCon.Alias = "Active"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCon.BracketCloseNum = 1
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

            oCon = oCons.Add
            '// (CardType = 'S'))
            oCon.BracketOpenNum = 1
            oCon.Alias = "Active"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCon.BracketCloseNum = 2
            'oCon = oCons.Add
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_3")
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.BracketOpenNum = 2
            oCon.Alias = "Active"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCon.BracketCloseNum = 1
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

            oCon = oCons.Add
            '// (CardType = 'S'))
            oCon.BracketOpenNum = 1
            oCon.Alias = "Active"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCon.BracketCloseNum = 2
            oCFL.SetConditions(oCons)


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
  

#End Region
    Private Sub Databind(ByVal aform As SAPbouiCOM.Form)
        aform.DataSources.UserDataSources.Add("intYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        aform.DataSources.UserDataSources.Add("intMonth", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        aform.DataSources.UserDataSources.Add("strComp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        aform.DataSources.UserDataSources.Add("strPost", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

        aform.DataSources.UserDataSources.Add("frmEMP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        aform.DataSources.UserDataSources.Add("toEmp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

        AddChooseFromList(aform)
        oEditText = aform.Items.Item("17").Specific
        oEditText.DataBind.SetBound(True, "", "frmEMP")
        oEditText.ChooseFromListUID = "CFL_2"
        oEditText.ChooseFromListAlias = "empID"

        oEditText = aform.Items.Item("19").Specific
        oEditText.DataBind.SetBound(True, "", "toEmp")
        oEditText.ChooseFromListUID = "CFL_3"
        oEditText.ChooseFromListAlias = "empID"

        oCombobox = aform.Items.Item("5").Specific
        oCombobox.ValidValues.Add("0", "")
        For intRow As Integer = 2010 To 2050
            oCombobox.ValidValues.Add(intRow, intRow)
        Next
        oCombobox.DataBind.SetBound(True, "", "intYear")
        oCombobox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
        aform.Items.Item("5").DisplayDesc = True
        oCombobox = aform.Items.Item("7").Specific
        oCombobox.ValidValues.Add("0", "")
        For intRow As Integer = 1 To 12
            oCombobox.ValidValues.Add(intRow, MonthName(intRow))
        Next

        oCombobox.DataBind.SetBound(True, "", "intMonth")
        oCombobox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
        aform.Items.Item("7").DisplayDesc = True

        oCombobox = aform.Items.Item("9").Specific
        oCombobox.DataBind.SetBound(True, "", "strComp")
        FillCombobox(oCombobox, "Select U_Z_CompCode,U_Z_CompName from [@Z_OADM]")

        oCombobox = aform.Items.Item("11").Specific
        oCombobox.ValidValues.Add("0", "")
        oCombobox.ValidValues.Add("R", "Regular")
        oCombobox.ValidValues.Add("O", "OffCycle")
        oCombobox.DataBind.SetBound(True, "", "strPost")
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        aform.Items.Item("11").DisplayDesc = True
    End Sub
    Public Sub FillCombobox(ByVal aCombo As SAPbouiCOM.ComboBox, ByVal aQuery As String)
        Dim oRS As SAPbobsCOM.Recordset
        oRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRS.DoQuery(aQuery)
        For intRow As Integer = aCombo.ValidValues.Count - 1 To 0 Step -1
            aCombo.ValidValues.Remove(intRow)
        Next
        aCombo.ValidValues.Add("", "")
        For intRow As Integer = 0 To oRS.RecordCount - 1
            Try
                aCombo.ValidValues.Add(oRS.Fields.Item(0).Value, oRS.Fields.Item(1).Value)

            Catch ex As Exception

            End Try
            oRS.MoveNext()
        Next
        aCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
    End Sub
    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim intYear, intMonth As Integer
            Dim strmonth, strPostMethod As String
            oCombobox = aForm.Items.Item("11").Specific
            If oCombobox.Selected.Value = "" Then
                oApplication.Utilities.Message("Select Posting Method", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            Else
                strPostMethod = oCombobox.Selected.Value
                If strPostMethod = "0" Then
                    oApplication.Utilities.Message("Select Posting Method", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Return False
                End If
            End If
            oCombobox = aForm.Items.Item("5").Specific
            If oCombobox.Selected.Value = "" Then
                oApplication.Utilities.Message("Select year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            Else
                intYear = oCombobox.Selected.Value
                If intYear = 0 Then
                    oApplication.Utilities.Message("Select year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Return False
                End If
            End If
            oCombobox = aForm.Items.Item("7").Specific
            If oCombobox.Selected.Value = "" Then
                oApplication.Utilities.Message("Select Month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            Else
                intMonth = oCombobox.Selected.Value
                If intYear = 0 Then
                    oApplication.Utilities.Message("Select Month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Return False
                Else
                    strmonth = oCombobox.Selected.Description
                End If
            End If
            Dim strCompany As String
            oCombobox = aForm.Items.Item("9").Specific
            If oCombobox.Selected.Value = "" Then
                oApplication.Utilities.Message("Select Company Code", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            Else
                strCompany = oCombobox.Selected.Value
            End If

            oCombobox = aForm.Items.Item("11").Specific
            If oCombobox.Selected.Value = "R" Then
                strPostMethod = "N"
            Else
                strPostMethod = "Y"
            End If
            Dim strFrmEmp, strToEmp, strCondition As String
            strFrmEmp = oApplication.Utilities.getEdittextvalue(aForm, "17")
            strToEmp = oApplication.Utilities.getEdittextvalue(aForm, "19")
            If strFrmEmp <> "" Then
                strCondition = " T1.empID >=" & strFrmEmp
            Else
                strCondition = " 1=1"
            End If

            If strToEmp <> "" Then
                strCondition = strCondition & " and T1.empID<=" & strToEmp
            Else
                strCondition = strCondition & " and 1=1"
            End If


            Dim strquery As String
            strquery = "SELECT T0.Code, T0.[U_Z_RefCode], T0.[U_Z_empid],[U_Z_EmpId1], T0.[U_Z_EmpName],T1.[Email],T0.[U_Z_CompNo], T0.[U_Z_NetSalary],T0.[U_Z_MONTH], T0.[U_Z_YEAR] FROM [@Z_PAYROLL1]  T0"
            strquery = strquery & " Inner Join OHEM T1 on T1.empID=T0.[U_Z_empid]  where " & strCondition & " and  U_Z_Posted='Y' and  T0.U_Z_CompNo='" & strCompany & "' and U_Z_OffCycle='" & strPostMethod & "' and U_Z_YEAR=" & intYear & " and U_Z_MONTH=" & intMonth & ""
            Dim otest As SAPbobsCOM.Recordset
            otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otest.DoQuery(strquery)
            If otest.RecordCount <= 0 Then
                oApplication.Utilities.Message("Payroll not posted for selected month and year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else

            End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Sub GridBind(ByVal aForm As SAPbouiCOM.Form)
        Dim intYear, intMonth As Integer
        Dim strCompany, strPostMethod, strquery, strFrmEmp, strToEmp, strCondition As String
        Dim oRecSet As SAPbobsCOM.Recordset
        oGrid = aForm.Items.Item("14").Specific
        Try
            oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCombobox = aForm.Items.Item("5").Specific
            intYear = oCombobox.Selected.Value
            oCombobox = aForm.Items.Item("7").Specific
            intMonth = oCombobox.Selected.Value
            oCombobox = aForm.Items.Item("9").Specific
            strCompany = oCombobox.Selected.Value
            oCombobox = aForm.Items.Item("11").Specific
            If oCombobox.Selected.Value = "R" Then
                strPostMethod = "N"
            Else
                strPostMethod = "Y"
            End If

            strFrmEmp = oApplication.Utilities.getEdittextvalue(aForm, "17")
            strToEmp = oApplication.Utilities.getEdittextvalue(aForm, "19")
            If strFrmEmp <> "" Then
                strCondition = " T1.empID >=" & strFrmEmp
            Else
                strCondition = " 1=1"
            End If

            If strToEmp <> "" Then
                strCondition = strCondition & " and T1.empID<=" & strToEmp
            Else
                strCondition = strCondition & " and 1=1"
            End If

            'strquery = "Select Code, [U_Z_empid],[U_Z_EmpId1],[U_Z_EmpName],[U_Z_CompNo],[U_Z_ExtraSalary],[U_Z_InrAmt],[U_Z_MONTH],[U_Z_YEAR],[U_Z_OffCycle],U_Z_RefCode"
            'strquery += "  from [@Z_PAYROLL1]  where U_Z_CompNo='" & strCompany & "' and U_Z_OffCycle='" & strPostMethod & "' and U_Z_YEAR=" & intYear & " and U_Z_MONTH=" & intMonth & ""


            strquery = "SELECT T0.Code, T0.[U_Z_RefCode], T0.[U_Z_empid],[U_Z_EmpId1],T1.ExtEmpNo 'Batch Number', T0.[U_Z_EmpName],T2.[Name],T1.[Email],T0.[U_Z_CompNo],  T0.[U_Z_NetSalary],T0.[U_Z_MONTH], T0.[U_Z_YEAR] FROM [@Z_PAYROLL1]  T0"

            strquery = strquery & " inner Join OHEM T1 on T1.empID=T0.U_Z_empid  Left Join OUDP T2 on T2.Code =T1.dept where  " & strCondition & " and T0.U_Z_CompNo='" & strCompany & "' and U_Z_OffCycle='" & strPostMethod & "' and U_Z_YEAR=" & intYear & " and U_Z_MONTH=" & intMonth & ""
            oGrid.DataTable.ExecuteQuery(strquery)
            aForm.Freeze(True)
            Formatgrid(oGrid)
            oApplication.Utilities.assignMatrixLineno(oGrid, aForm)
            aForm.Freeze(False)
            Dim oStatic As SAPbouiCOM.StaticText
            oStatic = aForm.Items.Item("15").Specific
            oStatic.Caption = ""
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub Formatgrid(ByVal agrid As SAPbouiCOM.Grid)
        agrid.Columns.Item("Code").TitleObject.Caption = "Code"
        agrid.Columns.Item("U_Z_empid").TitleObject.Caption = "System ID"
        oEditTextColumn = agrid.Columns.Item("U_Z_empid")
        oEditTextColumn.LinkedObjectType = "171"
        agrid.Columns.Item("U_Z_EmpId1").TitleObject.Caption = "Employee ID"
        agrid.Columns.Item("U_Z_EmpName").TitleObject.Caption = "Employee Name"
        agrid.Columns.Item("Name").TitleObject.Caption = "Department"
        agrid.Columns.Item("Email").TitleObject.Caption = "E-Mail ID"
        agrid.Columns.Item("U_Z_CompNo").TitleObject.Caption = "Company Name"
        agrid.Columns.Item("U_Z_MONTH").TitleObject.Caption = "Month"
        agrid.Columns.Item("U_Z_YEAR").TitleObject.Caption = "Year"
        agrid.Columns.Item("U_Z_RefCode").TitleObject.Caption = "Reference Code"
        agrid.Columns.Item("U_Z_NetSalary").TitleObject.Caption = "Net Salary"
        agrid.AutoResizeColumns()
        agrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
    End Sub
    Private Sub PrintPaySlip(ByVal aform As SAPbouiCOM.Form, aReportChoice As String)
        Dim strEmpId, strRefCode, strEmpName, strmonth, strYear As String
        Dim strCompany As String = ""

        oCombobox = aform.Items.Item("9").Specific
        If oCombobox.Selected.Value = "" Then
        Else
            strCompany = oCombobox.Selected.Value
        End If

        Try
            If CheckEmailsetup() = False Then
                Exit Sub
            End If

            Dim strReportFileName As String

            If aReportChoice = "PaySlip" Then
                strReportFileName = System.Windows.Forms.Application.StartupPath & "\Reports\" & strCompany & "_Payslip.rpt"
            ElseIf aReportChoice = "Saving" Then
                strReportFileName = System.Windows.Forms.Application.StartupPath & "\Reports\Saving Scheme Balance Report.rpt"
            Else
                strReportFileName = ""
            End If
            If strReportFileName = "" Then
                oApplication.Utilities.Message("Report does not exists : " & strReportFileName, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            If File.Exists(strReportFileName) = False Then
                oApplication.Utilities.Message("Report does not exists : " & strReportFileName, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oGrid = aform.Items.Item("14").Specific
            Dim oStatic As SAPbouiCOM.StaticText
            oStatic = aform.Items.Item("15").Specific
            Dim strDepartment As String
            If oGrid.Rows.Count > 0 Then
                For intRow As Integer = 0 To oGrid.Rows.Count - 1
                    strEmpId = oGrid.DataTable.GetValue("U_Z_empid", intRow)
                    strRefCode = oGrid.DataTable.GetValue("U_Z_RefCode", intRow)
                    strEmpName = oGrid.DataTable.GetValue("U_Z_EmpName", intRow)
                    strmonth = oGrid.DataTable.GetValue("U_Z_MONTH", intRow)
                    strYear = oGrid.DataTable.GetValue("U_Z_YEAR", intRow)
                    strDepartment = oGrid.DataTable.GetValue("Name", intRow)
                    oStatic.Caption = "Processing Employee ID : " & strEmpId
                    If oGrid.DataTable.GetValue("Email", intRow) <> "" Then
                        Dim oCrystalDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
                        ' Dim strReportFileName As String = System.Windows.Forms.Application.StartupPath & "\Reports\" & "RptMonthPaySlip.rpt"
                        If File.Exists(strReportFileName) Then
                            Dim strServer As String = oApplication.Company.Server ' ConfigurationManager.AppSettings("SAPServer")
                            Dim strDB As String = oApplication.Company.CompanyDB
                            Dim strUser As String = oApplication.Company.DbUserName
                            Dim strPwd As String = oApplication.Company.DbPassword
                            Dim crtableLogoninfos As New TableLogOnInfos
                            Dim crtableLogoninfo As New TableLogOnInfo
                            Dim crConnectionInfo As New ConnectionInfo
                            Dim CrTables As Tables
                            Dim CrTable As Table
                            oCrystalDocument.Load(strReportFileName)
                            With crConnectionInfo
                                .ServerName = strServer
                                .DatabaseName = strDB
                                .UserID = strUser
                                '.Password = strPwd
                                .IntegratedSecurity = True
                            End With
                            CrTables = oCrystalDocument.Database.Tables
                            For Each CrTable In CrTables
                                crtableLogoninfo = CrTable.LogOnInfo
                                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                            Next

                            If strRefCode <> 0 Then
                                If aReportChoice = "PaySlip" Then
                                    oCrystalDocument.SetParameterValue("U_Z_RefCode", strRefCode)
                                    oCrystalDocument.SetParameterValue("U_Z_Empid", strEmpId)
                                Else
                                    oCrystalDocument.SetParameterValue("Month", strmonth)
                                    oCrystalDocument.SetParameterValue("Year", strYear)
                                    oCrystalDocument.SetParameterValue("Department", strDepartment)
                                    oCrystalDocument.SetParameterValue("StaffID", strEmpId)
                                End If
                                
                            End If
                            Dim strFilename As String
                            ' Dim strFilename As String = System.Windows.Forms.Application.StartupPath & "\PaySlip\Payslip_" & strEmpName & "_" & MonthName(CInt(strmonth)) & "_" & strYear & ".pdf"
                            If aReportChoice = "PaySlip" Then

                                If Directory.Exists(strReportFilePah & "\PaySlip") = False Then
                                    Directory.CreateDirectory(strReportFilePah & "\PaySlip")
                                End If
                                strFilename = strReportFilePah & "\PaySlip\Payslip_" & strEmpName & "_" & MonthName(CInt(strmonth)) & "_" & strYear & ".pdf"
                                If File.Exists(strFilename) Then
                                    File.Delete(strFilename)
                                End If
                            Else

                                If Directory.Exists(strReportFilePah & "\PaySlip") = False Then
                                    Directory.CreateDirectory(strReportFilePah & "\PaySlip")
                                End If
                                strFilename = strReportFilePah & "\PaySlip\SavingScheme_" & strEmpName & "_" & MonthName(CInt(strmonth)) & "_" & strYear & ".pdf"
                                If File.Exists(strFilename) Then
                                    File.Delete(strFilename)
                                End If
                            End If

                            Dim CrExportOptions As CrystalDecisions.Shared.ExportOptions
                            Dim CrDiskFileDestinationOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions()
                            Dim CrFormatTypeOptions As New CrystalDecisions.Shared.PdfRtfWordFormatOptions()
                            CrDiskFileDestinationOptions.DiskFileName = strFilename

                            CrExportOptions = oCrystalDocument.ExportOptions
                            With CrExportOptions
                                .ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                                .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                                .DestinationOptions = CrDiskFileDestinationOptions
                                .FormatOptions = CrFormatTypeOptions
                            End With
                            oCrystalDocument.ExportToDisk(ExportFormatType.PortableDocFormat, strFilename)
                            oCrystalDocument.Export()
                            oCrystalDocument.Close()
                            Dim strMessage As String = "Payslip for " & MonthName(CInt(strmonth)) & "_" & strYear
                            SendMail(strEmpId, strFilename, strMessage)
                        End If
                    End If
                Next
            End If
            oStatic = aform.Items.Item("15").Specific
            oStatic.Caption = "Email sending completed"
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Function CheckEmailsetup() As Boolean
        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet.DoQuery("Select U_Z_SMTPSERV,U_Z_SMTPPORT,U_Z_SMTPUSER,U_Z_SMTPPWD,U_Z_SSL,U_Z_FilePath From [@Z_PAY_OMAIL]")
        If oRecordSet.RecordCount <= 0 Then
            oApplication.Utilities.Message("Email setup is not defined", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Else
            If oRecordSet.Fields.Item("U_Z_FilePath").Value = "" Then
                oApplication.Utilities.Message("Report File path not exists in Payslip -Email setup", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                strReportFilePah = oRecordSet.Fields.Item("U_Z_FilePath").Value
            End If
        End If
        Return True
    End Function

    Public Sub SendMail(ByVal EmpNo As String, ByVal strFileName As String, ByVal strSubject As String)
        Dim mailServer, mailPort, mailId, mailPwd, mailSSL, aMail
        Dim strMailContent, strContact, strMessage As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet.DoQuery("Select U_Z_SMTPSERV,U_Z_SMTPPORT,U_Z_SMTPUSER,U_Z_SMTPPWD,U_Z_SSL,U_Z_Text,U_Z_Contact From [@Z_PAY_OMAIL]")
        If Not oRecordSet.EoF Then
            mailServer = oRecordSet.Fields.Item("U_Z_SMTPSERV").Value
            mailPort = oRecordSet.Fields.Item("U_Z_SMTPPORT").Value
            mailId = oRecordSet.Fields.Item("U_Z_SMTPUSER").Value
            mailPwd = oRecordSet.Fields.Item("U_Z_SMTPPWD").Value
            mailSSL = oRecordSet.Fields.Item("U_Z_SSL").Value
            strMailContent = oRecordSet.Fields.Item("U_Z_Text").Value
            strContact = oRecordSet.Fields.Item("U_Z_Contact").Value
            strMailContent = strMailContent '& "-" & strContact


            strMessage = "<!DOCTYPE html><html><head><title></title></head><body>  <span>  " & strMailContent & " .</span> <br /><br />"
            strMessage += " <span> " & strContact & " .</span> <br /><br />"
            strMessage += "   <br /><br /></body></html>"
            If mailServer <> "" And mailId <> "" And mailPwd <> "" Then
                oRecordSet.DoQuery("Select * from OHEM where empID='" & EmpNo & "'")
                aMail = oRecordSet.Fields.Item("email").Value
                If aMail <> "" Then
                    SendMailEmployee(mailServer, mailPort, mailId, mailPwd, mailSSL, aMail, strFileName, strSubject, strMessage)
                End If
            Else
                oApplication.Utilities.Message("Mail Server Details Not Configured...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If
        End If
    End Sub
    Private Sub SendMailEmployee(ByVal mailServer As String, ByVal mailPort As String, ByVal mailId As String, ByVal mailpwd As String, ByVal mailSSL As String, ByVal toId As String, ByVal aPath As String, ByVal Message As String, aBody As String)
        Dim SmtpServer As New Net.Mail.SmtpClient()
        Dim mail As New Net.Mail.MailMessage
        Try
            SmtpServer.Credentials = New Net.NetworkCredential(mailId, mailpwd)
            SmtpServer.Port = mailPort
            SmtpServer.EnableSsl = mailSSL
            SmtpServer.Host = mailServer
            mail = New Net.Mail.MailMessage()
            mail.From = New Net.Mail.MailAddress(mailId, "Payroll")
            mail.To.Add(toId)
            mail.IsBodyHtml = True
            mail.Priority = MailPriority.High
            mail.Subject = Message
            mail.Body = aBody ' Message
            mail.Attachments.Add(New Net.Mail.Attachment(aPath))
            SmtpServer.Send(mail)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
            mail.Dispose()
        End Try
    End Sub
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Pay_SendPaySlip Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And oForm.PaneLevel = 2 Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "9" Then
                                    AddChooseFromList_Conditions(oForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" Then
                                    oForm.PaneLevel = oForm.PaneLevel + 1
                                    If oForm.PaneLevel = 3 Then
                                        GridBind(oForm)
                                    End If
                                ElseIf pVal.ItemUID = "12" Then
                                    oForm.PaneLevel = oForm.PaneLevel - 1
                                ElseIf pVal.ItemUID = "13" Then
                                    If oApplication.SBO_Application.MessageBox("Do you want to send the Payslips in Email?", , "Yes", "No") = 2 Then
                                        Exit Sub
                                    End If
                                    PrintPaySlip(oForm, "PaySlip")
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim oItm As SAPbobsCOM.Items
                                Dim sCHFL_ID, val, Val1 As String
                                Dim intChoice, introw As Integer
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        oForm.Freeze(True)
                                        If pVal.ItemUID = "17" Or pVal.ItemUID = "19" Then
                                            val = oDataTable.GetValue("empID", 0)
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                            Catch ex As Exception
                                            End Try
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                End Try

                        End Select
                End Select
            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_PaySlipEmail
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
