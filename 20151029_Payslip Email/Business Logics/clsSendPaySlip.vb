Imports System.IO
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Net
Imports System.Net.Mail
Imports System.Text
Imports System.Threading

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
    Dim offCycleQueryBuilder As StringBuilder

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

        offCycleQueryBuilder = New StringBuilder("DECLARE @U_Z_RefCode  varchar(1) ")
        offCycleQueryBuilder.Append("SELECT T0.Code, ")
        offCycleQueryBuilder.Append("@U_Z_RefCode as U_Z_RefCode, ")
        offCycleQueryBuilder.Append("T0.[U_Z_empid], ")
        offCycleQueryBuilder.Append("[U_Z_EmpId1], ")
        offCycleQueryBuilder.Append("T1.ExtEmpNo 'Batch Number', ")
        offCycleQueryBuilder.Append("T0.[U_Z_EmpName], ")
        offCycleQueryBuilder.Append("T1.[Email], ")
        offCycleQueryBuilder.Append("T3.[U_Z_CompNo],  ")
        offCycleQueryBuilder.Append("SUM((CASE WHEN T0.[U_Z_Type] = 'D' THEN -T0.[U_Z_Amount] ELSE T0.[U_Z_Amount] END  ))  as U_Z_NetSalary, ")
        offCycleQueryBuilder.Append("T0.[U_Z_MONTH], ")
        offCycleQueryBuilder.Append("T0.[U_Z_YEAR] ")
        offCycleQueryBuilder.Append("FROM [@Z_PAY_TRANS]  T0  ")
        offCycleQueryBuilder.Append("inner Join OHEM T1 on T1.empID=T0.U_Z_empid  ")
        offCycleQueryBuilder.Append("INNER JOIN [dbo].[OHEM] T3 ON T3.empID = T0.U_Z_EMPID  ")
        offCycleQueryBuilder.Append("WHERE  {0} and T3.U_Z_CompNo = '{1}' and U_Z_YEAR = {2} and U_Z_MONTH = {3} AND T0.U_Z_Posted = 'Y' AND T0.U_Z_offTool = 'Y' ")
        offCycleQueryBuilder.Append("GROUP BY T0.Code,  ")
        offCycleQueryBuilder.Append("T0.[U_Z_empid], ")
        offCycleQueryBuilder.Append("t0.[U_Z_EmpId1], ")
        offCycleQueryBuilder.Append("T1.ExtEmpNo, ")
        offCycleQueryBuilder.Append("T0.[U_Z_EmpName], ")
        offCycleQueryBuilder.Append("T1.[Email], ")
        offCycleQueryBuilder.Append("T3.[U_Z_CompNo], ")
        offCycleQueryBuilder.Append("T0.[U_Z_MONTH], ")
        offCycleQueryBuilder.Append("T0.[U_Z_YEAR] ")
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
        oCombobox.ValidValues.Add("T", "OffCycle Transaction")
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

            'Added by Houssam
            offCycleQueryBuilder = New StringBuilder("DECLARE @U_Z_RefCode  varchar(1) ")
            offCycleQueryBuilder.Append("SELECT T0.Code, ")
            offCycleQueryBuilder.Append("@U_Z_RefCode as U_Z_RefCode, ")
            offCycleQueryBuilder.Append("T0.[U_Z_empid], ")
            offCycleQueryBuilder.Append("[U_Z_EmpId1], ")
            offCycleQueryBuilder.Append("T1.ExtEmpNo 'Batch Number', ")
            offCycleQueryBuilder.Append("T0.[U_Z_EmpName], ")
            offCycleQueryBuilder.Append("T1.[Email], ")
            offCycleQueryBuilder.Append("T3.[U_Z_CompNo],  ")
            offCycleQueryBuilder.Append("SUM((CASE WHEN T0.[U_Z_Type] = 'D' THEN -T0.[U_Z_Amount] ELSE T0.[U_Z_Amount] END  ))  as U_Z_NetSalary, ")
            offCycleQueryBuilder.Append("T0.[U_Z_MONTH], ")
            offCycleQueryBuilder.Append("T0.[U_Z_YEAR] ")
            offCycleQueryBuilder.Append("FROM [@Z_PAY_TRANS]  T0  ")
            offCycleQueryBuilder.Append("inner Join OHEM T1 on T1.empID=T0.U_Z_empid  ")
            offCycleQueryBuilder.Append("INNER JOIN [dbo].[OHEM] T3 ON T3.empID = T0.U_Z_EMPID  ")
            offCycleQueryBuilder.Append("WHERE  {0} and T3.U_Z_CompNo = '{1}' and U_Z_YEAR = {2} and U_Z_MONTH = {3} AND T0.U_Z_Posted = 'Y' AND T0.U_Z_offTool = 'Y' ")
            offCycleQueryBuilder.Append("GROUP BY T0.Code,  ")
            offCycleQueryBuilder.Append("T0.[U_Z_empid], ")
            offCycleQueryBuilder.Append("t0.[U_Z_EmpId1], ")
            offCycleQueryBuilder.Append("T1.ExtEmpNo, ")
            offCycleQueryBuilder.Append("T0.[U_Z_EmpName], ")
            offCycleQueryBuilder.Append("T1.[Email], ")
            offCycleQueryBuilder.Append("T3.[U_Z_CompNo], ")
            offCycleQueryBuilder.Append("T0.[U_Z_MONTH], ")
            offCycleQueryBuilder.Append("T0.[U_Z_YEAR] ")

            If oCombobox.Selected.Value = "T" Then
                strquery = String.Format(offCycleQueryBuilder.ToString(), strCondition, strCompany, intYear, intMonth)
            Else
                'Commented By Houssam
                strquery = "SELECT T0.Code, T0.[U_Z_RefCode], T0.[U_Z_empid],[U_Z_EmpId1], T0.[U_Z_EmpName],T1.[Email],T0.[U_Z_CompNo], T0.[U_Z_NetSalary],T0.[U_Z_MONTH], T0.[U_Z_YEAR] FROM [@Z_PAYROLL1]  T0"
                strquery = strquery & " Inner Join OHEM T1 on T1.empID=T0.[U_Z_empid]  where " & strCondition & " and  U_Z_Posted='Y' and  T0.U_Z_CompNo='" & strCompany & "' and U_Z_OffCycle='" & strPostMethod & "' and U_Z_YEAR=" & intYear & " and U_Z_MONTH=" & intMonth & ""
            End If

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

    Dim query As String = String.Empty
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

            If oCombobox.Selected.Value = "T" Then
                strquery = String.Format(offCycleQueryBuilder.ToString(), strCondition, strCompany, intYear, intMonth)
            Else
                strquery = "Select Code, [U_Z_empid],[U_Z_EmpId1],[U_Z_EmpName],[U_Z_CompNo],[U_Z_ExtraSalary],[U_Z_InrAmt],[U_Z_MONTH],[U_Z_YEAR],[U_Z_OffCycle],U_Z_RefCode"
                strquery += "  from [@Z_PAYROLL1]  where U_Z_CompNo='" & strCompany & "' and U_Z_OffCycle='" & strPostMethod & "' and U_Z_YEAR=" & intYear & " and U_Z_MONTH=" & intMonth & ""
            End If
            'strquery = "Select Code, [U_Z_empid],[U_Z_EmpId1],[U_Z_EmpName],[U_Z_CompNo],[U_Z_ExtraSalary],[U_Z_InrAmt],[U_Z_MONTH],[U_Z_YEAR],[U_Z_OffCycle],U_Z_RefCode"
            'strquery += "  from [@Z_PAYROLL1]  where U_Z_CompNo='" & strCompany & "' and U_Z_OffCycle='" & strPostMethod & "' and U_Z_YEAR=" & intYear & " and U_Z_MONTH=" & intMonth & ""
            query = strquery
            oGrid.DataTable.ExecuteQuery(strquery)
            Formatgrid(oGrid)
            oApplication.Utilities.assignMatrixLineno(oGrid, aForm)
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
        agrid.Columns.Item("Email").TitleObject.Caption = "E-Mail ID"
        agrid.Columns.Item("U_Z_CompNo").TitleObject.Caption = "Company Name"
        agrid.Columns.Item("U_Z_MONTH").TitleObject.Caption = "Month"
        agrid.Columns.Item("U_Z_YEAR").TitleObject.Caption = "Year"
        agrid.Columns.Item("U_Z_RefCode").TitleObject.Caption = "Reference Code"
        agrid.Columns.Item("U_Z_NetSalary").TitleObject.Caption = "Net Salary"
        agrid.AutoResizeColumns()
        agrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
    End Sub
    Private Sub PrintPaySlip(ByVal aform As SAPbouiCOM.Form)
        'Dim tw As TextWriter
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
            oCombobox = aform.Items.Item("11").Specific
            Dim offCycleTransactionPath As String = String.Empty
            If oCombobox.Selected.Value = "T" Then
                offCycleTransactionPath = "_Offycle"
            Else
                offCycleTransactionPath = String.Empty
            End If
            Dim path As String = System.Windows.Forms.Application.StartupPath & "\" & Guid.NewGuid().ToString() & ".txt"

            'Dim fs1 As FileStream = New FileStream(path, FileMode.OpenOrCreate, FileAccess.Write)
            'tw = New StreamWriter(fs1)

            'File.Create(path)
            'tw = New StreamWriter(path, True)

            Dim strReportFileName As String = System.Windows.Forms.Application.StartupPath & "\Reports\" & strCompany & offCycleTransactionPath & "_Payslip.rpt"
            'tw.WriteLine("the strReportFileName is  " & strReportFileName)
            If File.Exists(strReportFileName) = False Then
                oApplication.Utilities.Message("Payslip report does not exists : " & strReportFileName, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oGrid = aform.Items.Item("14").Specific
            Dim oStatic As SAPbouiCOM.StaticText
            oStatic = aform.Items.Item("15").Specific

            Dim emailContents As List(Of EmailContent) = New List(Of EmailContent)
            Dim employeesList As List(Of String) = New List(Of String)
            Dim employeesDT As SAPbouiCOM.DataTable = oGrid.DataTable
            Dim oTempForm As SAPbouiCOM.Items = oForm.Items
            'tw.WriteLine("Enter the loop")
            If employeesDT.Rows.Count > 0 Then
                For intRow As Integer = 0 To employeesDT.Rows.Count - 1
                    strEmpId = employeesDT.GetValue("U_Z_empid", intRow)
                    strRefCode = employeesDT.GetValue("U_Z_RefCode", intRow)
                    strEmpName = employeesDT.GetValue("U_Z_EmpName", intRow)
                    strmonth = employeesDT.GetValue("U_Z_MONTH", intRow)
                    strYear = employeesDT.GetValue("U_Z_YEAR", intRow)
                    oStatic.Caption = "Processing Employee ID : " & strEmpId

                    If employeesDT.GetValue("Email", intRow) <> "" Then
                        Dim oCrystalDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument

                        If File.Exists(strReportFileName) Then
                            Dim strServer As String = oApplication.Company.Server
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
                                .IntegratedSecurity = True
                            End With
                            CrTables = oCrystalDocument.Database.Tables
                            For Each CrTable In CrTables
                                crtableLogoninfo = CrTable.LogOnInfo
                                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                            Next

                            oCombobox = aform.Items.Item("11").Specific

                            If oCombobox.Selected.Value = "T" Then
                                oCrystalDocument.SetParameterValue("@U_Z_Empid", strEmpId)
                                oCrystalDocument.SetParameterValue("@U_Month", strmonth)
                                oCrystalDocument.SetParameterValue("@U_Year", strYear)
                            Else
                                If strRefCode <> 0 Then
                                    oCrystalDocument.SetParameterValue("U_Z_RefCode", strRefCode)
                                    oCrystalDocument.SetParameterValue("U_Z_Empid", Convert.ToDouble(strEmpId))
                                End If
                            End If
                            'tw.WriteLine("Check the Directory path")

                            If Directory.Exists(strReportFilePah & "\PaySlip") = False Then
                                Directory.CreateDirectory(strReportFilePah & "\PaySlip")
                            End If
                            'tw.WriteLine(String.Format("strEmpName = {0};strmonth = {1};strYear = {2};", strEmpName, strmonth, strYear))
                            Dim strFilename As String = strReportFilePah & "\PaySlip\Payslip_" & strEmpName.Replace("\", "").Replace("/", "") & "_" & MonthName(CInt(strmonth)) & "_" & strYear & ".pdf"

                            'If File.Exists(strFilename) Then
                            '    File.Delete(strFilename)
                            'End If

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
                            'tw.WriteLine("Check if employee in list")
                            If (employeesList.Contains(strEmpId) = False) Then
                                'tw.WriteLine("Employee id = " & strEmpId)
                                'tw.WriteLine("strFilename = " & strFilename)

                                employeesList.Add(strEmpId)
                                If File.Exists(strFilename) Then
                                    'tw.WriteLine("File Exists ")
                                    File.Delete(strFilename)
                                    'tw.WriteLine("File Deleted")
                                End If
                                'tw.WriteLine("Create PDF File ")
                                oCrystalDocument.ExportToDisk(ExportFormatType.PortableDocFormat, strFilename)
                                oCrystalDocument.Export()
                                'tw.WriteLine("PDF File Created ")
                            End If
                            Dim strMessage As String = "Payslip for " & MonthName(CInt(strmonth)) & "_" & strYear
                            Dim isAvailable As Boolean = IsAvailableInList(emailContents, strEmpId)

                            If (isAvailable = False) Then
                                emailContents.Add(New EmailContent(strEmpId, strFilename, strMessage))
                            End If
                            'SendMail(strEmpId, strFilename, strMessage)
                        End If
                        oCrystalDocument.Close()
                    End If

                Next
            End If

            'tw.WriteLine("End of Loop")
            'tw.Close()
            Dim emailContent As EmailContent
            For Each emailContent In emailContents
                SendMail(emailContent.EmployeeId, emailContent.FileName, emailContent.Message)
            Next

            oStatic = aform.Items.Item("15").Specific
            oStatic.Caption = "Email sending completed"
        Catch ex As Exception
            'tw.Close()
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally

        End Try

    End Sub

    Public Sub ExportToPdf(ByVal reportFileName As String, ByVal offCycleTransactionPath As String, ByVal employeesList As List(Of String), ByVal emailContents As List(Of EmailContent),
                           ByVal strEmpId As String, ByVal strmonth As String, ByVal strYear As String, ByVal strRefCode As String, ByVal strEmpName As String)

        'Dim oRecSet As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'oRecSet.DoQuery(query)

        'While Not oRecSet.EoF

        '    strEmpId = oRecSet.Fields.Item("U_Z_empid").Value
        '    strRefCode = oRecSet.Fields.Item("U_Z_RefCode").Value
        '    strEmpName = oRecSet.Fields.Item("U_Z_EmpName").Value
        '    strmonth = oRecSet.Fields.Item("U_Z_MONTH").Value
        '    strYear = oRecSet.Fields.Item("U_Z_YEAR").Value
        '    oStatic.Caption = "Processing Employee ID : " & strEmpId

        '    If oRecSet.Fields.Item("Email").Value <> "" Then

        '        Dim subThread As Thread = New Thread(Sub() ExportToPdf(strReportFileName, offCycleTransactionPath, employeesList, emailContents, strEmpId, strmonth, strYear, strRefCode, strEmpName))
        '        subThread.SetApartmentState(ApartmentState.STA)
        '        subThread.Start()
        '        Dim oCrystalDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        '        ' Dim strReportFileName As String = System.Windows.Forms.Application.StartupPath & "\Reports\" & "RptMonthPaySlip.rpt"
        '        If File.Exists(strReportFileName) Then
        '            Dim strServer As String = oApplication.Company.Server ' ConfigurationManager.AppSettings("SAPServer")
        '            Dim strDB As String = oApplication.Company.CompanyDB
        '            Dim strUser As String = oApplication.Company.DbUserName
        '            Dim strPwd As String = oApplication.Company.DbPassword
        '            Dim crtableLogoninfos As New TableLogOnInfos
        '            Dim crtableLogoninfo As New TableLogOnInfo
        '            Dim crConnectionInfo As New ConnectionInfo
        '            Dim CrTables As Tables
        '            Dim CrTable As Table
        '            oCrystalDocument.Load(strReportFileName)
        '            With crConnectionInfo
        '                .ServerName = strServer
        '                .DatabaseName = strDB
        '                .UserID = strUser
        '                '.Password = strPwd
        '                .IntegratedSecurity = True
        '            End With
        '            CrTables = oCrystalDocument.Database.Tables
        '            For Each CrTable In CrTables
        '                crtableLogoninfo = CrTable.LogOnInfo
        '                crtableLogoninfo.ConnectionInfo = crConnectionInfo
        '                CrTable.ApplyLogOnInfo(crtableLogoninfo)
        '            Next

        '            'oCombobox = aform.Items.Item("11").Specific

        '            If offCycleTransactionPath = "_Offycle" Then
        '                oCrystalDocument.SetParameterValue("U_Z_Empid", strEmpId)
        '                oCrystalDocument.SetParameterValue("U_Month", strmonth)
        '                oCrystalDocument.SetParameterValue("U_Year", strYear)
        '            Else
        '                If strRefCode <> 0 Then
        '                    oCrystalDocument.SetParameterValue("U_Z_RefCode", strRefCode)
        '                    oCrystalDocument.SetParameterValue("U_Z_Empid", Convert.ToDouble(strEmpId))
        '                End If
        '            End If

        '            ' Dim strFilename As String = System.Windows.Forms.Application.StartupPath & "\PaySlip\Payslip_" & strEmpName & "_" & MonthName(CInt(strmonth)) & "_" & strYear & ".pdf"
        '            If Directory.Exists(strReportFilePah & "\PaySlip") = False Then
        '                Directory.CreateDirectory(strReportFilePah & "\PaySlip")
        '            End If
        '            Dim strFilename As String = strReportFilePah & "\PaySlip\Payslip_" & strEmpName.Replace("/", "") & "_" & MonthName(CInt(strmonth)) & "_" & strYear & ".pdf"
        '            'If File.Exists(strFilename) Then
        '            '    File.Delete(strFilename)
        '            'End If
        '            Dim CrExportOptions As CrystalDecisions.Shared.ExportOptions
        '            Dim CrDiskFileDestinationOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions()
        '            Dim CrFormatTypeOptions As New CrystalDecisions.Shared.PdfRtfWordFormatOptions()
        '            CrDiskFileDestinationOptions.DiskFileName = strFilename

        '            CrExportOptions = oCrystalDocument.ExportOptions
        '            With CrExportOptions
        '                .ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
        '                .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
        '                .DestinationOptions = CrDiskFileDestinationOptions
        '                .FormatOptions = CrFormatTypeOptions
        '            End With
        '            If (employeesList.Contains(strEmpId) = False) Then
        '                employeesList.Add(strEmpId)
        '                If File.Exists(strFilename) Then
        '                    File.Delete(strFilename)
        '                End If
        '                oCrystalDocument.ExportToDisk(ExportFormatType.PortableDocFormat, strFilename)
        '                oCrystalDocument.Export()
        '            End If

        '            oCrystalDocument.Close()
        '            Dim strMessage As String = "Payslip for " & MonthName(CInt(strmonth)) & "_" & strYear

        '            Dim isAvailable As Boolean = IsAvailableInList(emailContents, strEmpId)
        '            If (isAvailable = False) Then
        '                emailContents.Add(New EmailContent(strEmpId, strFilename, strMessage))
        '            End If
        '            'SendMail(strEmpId, strFilename, strMessage)
        '        End If
        '    End If

        'oRecSet.MoveNext()
        'End While
        Dim oCrystalDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        'Dim strReportFileName As String = System.Windows.Forms.Application.StartupPath & "\Reports\" & "RptMonthPaySlip.rpt"
        Dim strReportFileName As String = "C:\SAP\Payslip PDF\PaySlip\Temp\" & Guid.NewGuid().ToString() & ".rpt"
        File.Copy(reportFileName, strReportFileName)
        Dim fileStream As FileStream = New FileStream(strReportFileName, FileMode.Open, FileAccess.Read)
        fileStream.Close()
        If File.Exists(strReportFileName) Then
            Dim strServer As String = oApplication.Company.Server
            Dim strDB As String = oApplication.Company.CompanyDB
            Dim strUser As String = oApplication.Company.DbUserName
            Dim strPwd As String = oApplication.Company.DbPassword
            Dim crtableLogoninfos As New TableLogOnInfos
            Dim crtableLogoninfo As New TableLogOnInfo
            Dim crConnectionInfo As New ConnectionInfo
            Dim CrTables As Tables
            Dim CrTable As Table
            Thread.Sleep(100)

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

            'oCombobox = aform.Items.Item("11").Specific

            If offCycleTransactionPath = "_Offycle" Then
                oCrystalDocument.SetParameterValue("U_Z_Empid", strEmpId)
                oCrystalDocument.SetParameterValue("U_Month", strmonth)
                oCrystalDocument.SetParameterValue("U_Year", strYear)
            Else
                If strRefCode <> 0 Then
                    oCrystalDocument.SetParameterValue("U_Z_RefCode", strRefCode)
                    oCrystalDocument.SetParameterValue("U_Z_Empid", Convert.ToDouble(strEmpId))
                End If
            End If

            ' Dim strFilename As String = System.Windows.Forms.Application.StartupPath & "\PaySlip\Payslip_" & strEmpName & "_" & MonthName(CInt(strmonth)) & "_" & strYear & ".pdf"
            If Directory.Exists(strReportFilePah & "\PaySlip") = False Then
                Directory.CreateDirectory(strReportFilePah & "\PaySlip")
            End If
            Dim strFilename As String = strReportFilePah & "\PaySlip\Payslip_" & strEmpName.Replace("/", "") & "_" & MonthName(CInt(strmonth)) & "_" & strYear & ".pdf"
            'If File.Exists(strFilename) Then
            '    File.Delete(strFilename)
            'End If
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
            If (employeesList.Contains(strEmpId) = False) Then
                employeesList.Add(strEmpId)
                If File.Exists(strFilename) Then
                    File.Delete(strFilename)
                End If
                oCrystalDocument.ExportToDisk(ExportFormatType.PortableDocFormat, strFilename)
                oCrystalDocument.Export()
            End If

            oCrystalDocument.Close()
            Dim strMessage As String = "Payslip for " & MonthName(CInt(strmonth)) & "_" & strYear

            Dim isAvailable As Boolean = IsAvailableInList(emailContents, strEmpId)
            If (isAvailable = False) Then
                emailContents.Add(New EmailContent(strEmpId, strFilename, strMessage))
            End If
        End If
    End Sub

    Public Function IsAvailableInList(ByVal emailContents As List(Of EmailContent), ByVal empId As String) As Boolean
        Dim emilContent As EmailContent
        Dim isAvailable As Boolean = False
        For Each emilContent In emailContents
            If (emilContent.EmployeeId = empId) Then
                isAvailable = True
                Exit For
            End If
        Next
        Return isAvailable
    End Function
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
        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet.DoQuery("Select U_Z_SMTPSERV,U_Z_SMTPPORT,U_Z_SMTPUSER,U_Z_SMTPPWD,U_Z_SSL From [@Z_PAY_OMAIL]")
        If Not oRecordSet.EoF Then
            mailServer = oRecordSet.Fields.Item("U_Z_SMTPSERV").Value
            mailPort = oRecordSet.Fields.Item("U_Z_SMTPPORT").Value
            mailId = oRecordSet.Fields.Item("U_Z_SMTPUSER").Value
            mailPwd = oRecordSet.Fields.Item("U_Z_SMTPPWD").Value
            mailSSL = oRecordSet.Fields.Item("U_Z_SSL").Value
            If mailServer <> "" And mailId <> "" And mailPwd <> "" Then
                oRecordSet.DoQuery("Select * from OHEM where empID='" & EmpNo & "'")
                aMail = oRecordSet.Fields.Item("email").Value
                If aMail <> "" Then
                    SendMailEmployee(mailServer, mailPort, mailId, mailPwd, mailSSL, aMail, strFileName, strSubject)
                End If
            Else
                oApplication.Utilities.Message("Mail Server Details Not Configured...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If
        End If
    End Sub
    Private Sub SendMailEmployee(ByVal mailServer As String, ByVal mailPort As String, ByVal mailId As String, ByVal mailpwd As String, ByVal mailSSL As String, ByVal toId As String, ByVal aPath As String, ByVal Message As String)
        Dim SmtpServer As New Net.Mail.SmtpClient()
        Dim mail As New Net.Mail.MailMessage
        Try
            SmtpServer.Credentials = New Net.NetworkCredential(mailId, mailpwd)
            SmtpServer.Port = mailPort
            SmtpServer.EnableSsl = mailSSL
            SmtpServer.Timeout = 2000000
            SmtpServer.Host = mailServer
            mail = New Net.Mail.MailMessage()
            mail.From = New Net.Mail.MailAddress(mailId, "Payroll")
            mail.To.Add(toId)
            mail.IsBodyHtml = False
            mail.Priority = MailPriority.High
            mail.Subject = Message
            mail.Body = Message
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
                                    Dim newThread As New Thread(Sub() Me.PrintPaySlip(oForm))
                                    newThread.Start()
                                    'PrintPaySlip(oForm)<<======Commented by Houssam
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

Public Class EmailContent
    Public Message As String
    Public EmployeeId As String
    Public FileName As String

    Public Sub New(ByVal empId As String, ByVal fileName As String, ByVal message As String)
        Me.Message = message
        Me.EmployeeId = empId
        Me.FileName = fileName
    End Sub
End Class
