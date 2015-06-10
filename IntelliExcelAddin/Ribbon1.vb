'TODO:  Follow these steps to enable the Ribbon (XML) item:

'1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Ribbon1()
'End Function

'2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
'   actions, such as clicking a button. Note: if you have exported this Ribbon from the
'   Ribbon designer, move your code from the event handlers to the callback methods and
'   modify the code to work with the Ribbon extensibility (RibbonX) programming model.

'3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

'For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.

Imports Microsoft.Office.Interop.Excel
Imports System
Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.IO
Imports MySql.Data.MySqlClient

<Runtime.InteropServices.ComVisible(True)> _
Public Class Ribbon1
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("IntelliExcelAddin.Ribbon1.xml")
    End Function

#Region "Ribbon Callbacks"
    'Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub



#End Region

    '***********************************************************************************************

#Region "Global Variables"
    'Global Variables that need to be saved
    Private mCurrentItemID As Object
    Private xlWks As Worksheet
    Private xlWkb As Workbook
    Private xlFunc As WorksheetFunction
    Private style As Excel.Style
    Private tLen As Integer
    Private XlApp As Excel.Application
    Private Const limit As Integer = 5000 'The number of operations that Population and Validator will do per batch
    'Private Const xLimit As Integer = 1048576
    Private Const xLimit As Integer = 10001
    Private username As String
    Private password As String
#End Region

#Region "My Helpers"
    Private Sub VariableSetup()
        If IsNothing(XlApp) Then
            XlApp = Globals.ThisAddIn.Application
            xlWkb = XlApp.ActiveWorkbook
            xlWks = XlApp.ActiveSheet
            xlFunc = XlApp.WorksheetFunction
        ElseIf XlApp.ActiveWorkbook IsNot xlWkb Then
            xlWkb = XlApp.ActiveWorkbook
            xlWks = XlApp.ActiveSheet
        ElseIf XlApp.ActiveSheet IsNot xlWks Then
            xlWks = XlApp.ActiveSheet
        End If

    End Sub

    Function num2col(num As Integer) As String
        ' Subtract one to make modulo/divide cleaner. '
        num = num - 1
        ' Select return value based on invalid/one-char/two-char input. '
        If num < 0 Or num >= 27 * 26 Then
            ' Return special sentinel value if out of range. '
            num2col = "-"
        Else
            ' Single char, just get the letter. '

            If num < 26 Then
                num2col = Chr(num + 65)
            Else
                ' Double char, get letters based on integer divide and modulus. 
                num2col = Chr(num \ 26 + 64) + Chr(num Mod 26 + 65)
            End If
        End If
    End Function

    Private Sub loginInfo()
        If IsNothing(username) Or IsNothing(password) Then
            Dim loginForm As New MyLogin
            loginForm.ShowDialog()
            If loginForm.DialogResult = DialogResult.OK Then
                username = loginForm.username.Text
                password = loginForm.password.Text
            End If
            loginForm.Close()
        End If
    End Sub

    Private Function SqlCommand() As String
        Dim cForm As New GetCommand
        Dim cmdStr As String = Nothing
        cForm.ShowDialog()
        If cForm.DialogResult = DialogResult.OK Then
            Dim sel As String = cForm.mySelect.Text
            Dim frm As String = cForm.myFrom.Text
            Dim whr As String = cForm.myWhere.Text
            Dim ord As String = cForm.myOrder.Text
            cmdStr = "SELECT " & sel & " FROM " & frm
            If Not IsNothing(whr) Then
                cmdStr += " WHERE " & whr
            ElseIf Not IsNothing(ord) Then
                cmdStr += " ORDER BY " & ord
            End If
        End If
        Return cmdstr
    End Function

    Private Sub Export2Sheet(pForm As ValForm, row As Integer, limit As Integer)
        Dim sWatch As New Stopwatch 'Keep track of how fast the program is running
        Dim sh As Worksheet
        Dim mflag As Boolean = False
        Dim eflag As Boolean = False
        Dim result As Integer
        Dim strFullPath As String
        Dim tempList As New List(Of String)
        Dim MyInput As String
        Dim i As Integer

        MyInput = pForm.lastCol.Text
        Dim InputColumn = pForm.colNum.Value
        pForm.Close()

        For Each sh In xlWkb.Sheets
            If (sh.Name = "Master") Then
                mflag = True
            End If
            If (sh.Name = "Export") Then
                eflag = True
            End If
        Next

        If mflag = False And eflag = False Then
            result = MsgBox("You do not have both the 'Master' and 'Export' sheet!" & vbCrLf & "Do you want me to create it for you?", MsgBoxStyle.YesNo)
            If (result = MsgBoxResult.Yes) Then
                Dim nWks As Worksheet
                nWks = CType(xlWkb.Worksheets.Add(After:=xlWks), Worksheet)
                nWks.Name = "Export"
                xlWks.Name = "Master"
            Else : Exit Sub
            End If
        ElseIf mflag = True And eflag = False Then
            result = MsgBox("You do not have the 'Export' sheet!" & vbCrLf & "Do you want me to create it for you?", MsgBoxStyle.YesNo)
            If (result = MsgBoxResult.Yes) Then
                Dim nWks As Worksheet
                nWks = CType(xlWkb.Worksheets.Add(After:=xlWks), Worksheet)
                nWks.Name = "Export"
            Else : Exit Sub
            End If
        ElseIf mflag = False And eflag = True Then
            result = MsgBox("You do not have the 'Master' sheet!" & vbCrLf & "Do you want me to create it for you?", MsgBoxStyle.YesNo)
            If (result = MsgBoxResult.Yes) Then
                xlWks.Name = "Master"
            Else : Exit Sub
            End If
        End If

        xlWkb.Sheets("Master").Activate()

        '...................................../Format_Checks..............................................

        'Import File
        strFullPath = XlApp.GetOpenFilename("Text Files (*.csv),*.csv", , "Please select IMPORT file...")

        XlApp.StatusBar = "If you smell what The Rock is Cooking!"

        If strFullPath = "False" Then Exit Sub 'User pressed Cancel on the open file dialog

        'Copy & Paste Headers from "Master" to "Export"
        Dim column As String = Nothing
        For Each x In MyInput
            If IsNumeric(x) = False Then
                column += x
            End If
        Next

        xlWkb.Sheets("Export").Range("A1:" & column & "1").Value = xlWkb.Sheets("Master").Range("A1:" & column & "1").Value

        sWatch.Start()

        'Import in csv file
        Using MyReader As New  _
            Microsoft.VisualBasic.FileIO.TextFieldParser(strFullPath)
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(",")

            'Set the array
            Dim currentRow() As String
            Dim MArray(limit, InputColumn) As String
            Dim SFColumn As String = num2col(InputColumn + 1)
            Dim flag As Boolean
            If pForm.headerCheckbox.Checked Then
                flag = True
            Else
                flag = False
            End If


            'This While Loop Grabs batches specified by 'limit'
            While Not MyReader.EndOfData
                i = 0
                While i < limit And Not MyReader.EndOfData
                    If (flag) Then
                        currentRow = MyReader.ReadFields ' Get rid of headers
                        flag = False
                    End If

                    currentRow = MyReader.ReadFields
                    For j As Integer = 0 To currentRow.Length - 1
                        MArray(i, j) = currentRow(j)
                    Next
                    i += 1
                End While

                xlWkb.Sheets("Master").Range("A3:" & num2col(InputColumn) & i + 2).Value = MArray


                '.....................................FormulaDRAG.................................................
                xlWkb.Sheets("Master").Range(SFColumn & "2:" & MyInput).Resize(i + 1).FillDown()

                'Force formula calculations before copy&paste values
                xlWkb.Sheets("Master").Range(SFColumn & "3:" & column & i + 2).Calculate()

                '......................................ExportDATA.................................................
                xlWkb.Sheets("Export").Range("A" & row & ":" & column & row + i - 1).Value = xlWkb.Sheets("Master").Range("A3:" & column & 2 + i).Value

                'Clear PNs from Master
                xlWkb.Sheets("Master").Range("A3:" & column & i + 2).Delete()

                row = row + i

            End While
            MyReader.Close()
            MyReader.Dispose()
        End Using

        sWatch.Stop()
        XlApp.StatusBar = "Layeth The Smackethdown!"
        MsgBox("Sample Data READY! " & vbCrLf & "Execution Time: " & sWatch.ElapsedMilliseconds / 1000 & " s" & vbCrLf & "Processed " & row - 2 & " records @ " & limit & " per batch")

        '...........................................Export_To_CSV.................................................
        result = MsgBox("Do you want to export this to a .CSV file?", MsgBoxStyle.YesNo)
        If result = MsgBoxResult.Yes Then
            Dim rowValues As String = Nothing

            Dim progress = New loading_bar
            progress.Show()
            progress.ProgressBar1.Minimum = 0
            progress.ProgressBar1.Maximum = row - 1

            Dim path1 As String = Path.GetTempFileName()
            Dim fi As FileInfo = New FileInfo(strFullPath)
            Dim oFile As String = fi.DirectoryName & "\" & Path.GetFileNameWithoutExtension(fi.Name) & "_output.csv"


            If File.Exists(oFile) = True Then
                File.Delete(oFile)
            End If

            Dim sw As StreamWriter = File.CreateText(oFile)

            For x As Integer = 1 To row - 1
                rowValues = Nothing
                For y As Integer = 1 To xlWks.Range(MyInput).Column
                    rowValues = rowValues & "," & xlWkb.Sheets("Export").Cells(x, y).Value
                Next
                sw.WriteLine(rowValues.Substring(1))
                progress.ProgressBar1.Value += 1
            Next

            'Clean up
            sw.Flush()
            sw.Close()
            progress.Close()
            MsgBox("Successfully Exported to CSV file (" & oFile)
        End If
    End Sub

    Private Sub Export2CSV(pForm As ValForm, row As Integer, limit As Integer)
        Dim sWatch As New Stopwatch 'Keep track of how fast the program is running
        Dim sh As Worksheet
        Dim mflag As Boolean = False
        Dim result As Integer
        Dim strFullPath As String
        Dim MyInput As String
        Dim i As Integer

        MyInput = pForm.lastCol.Text
        Dim InputColumn = pForm.colNum.Value
        pForm.Close()

        For Each sh In xlWkb.Sheets
            If (sh.Name = "Master") Then
                mflag = True
            End If
        Next

        If mflag = False Then
            result = MsgBox("You do not have the 'Master' sheet!" & vbCrLf & "Do you want me to create it for you?", MsgBoxStyle.YesNo)
            If (result = MsgBoxResult.Yes) Then
                xlWks.Name = "Master"
            Else : Exit Sub
            End If
        End If
        xlWkb.Sheets("Master").Activate()

        strFullPath = XlApp.GetOpenFilename("Text Files (*.csv),*.csv", , "Please select IMPORT file...")

        sWatch.Start()
        XlApp.StatusBar = "If you smell what The Rock is Cooking!"

        If strFullPath = "False" Then Exit Sub 'User pressed Cancel on the open file dialog

        'Dim rowValues As String = Nothing
        Dim rowValues As String
        Dim sList As New System.Text.StringBuilder

        Dim path1 As String = Path.GetTempFileName()
        Dim fi As FileInfo = New FileInfo(strFullPath)
        Dim oFile As String = fi.DirectoryName & "\" & Path.GetFileNameWithoutExtension(fi.Name) & "_output.csv"
        Dim progress = New loading_bar

        Dim max As Integer = File.ReadLines(strFullPath).Count  'Could severely slow down operation time if use on a large file

        progress.ProgressBar1.Minimum = 0
        progress.ProgressBar1.Maximum = max * 2

        If File.Exists(oFile) = True Then
            File.Delete(oFile)
        End If

        'Dim sw As StreamWriter = File.CreateText(oFile)
        Dim hList As New List(Of String)
        Dim column As String = Nothing

        For Each x In MyInput
            If IsNumeric(x) = False Then
                column += x
            End If
        Next

        'Write Headers
        For x As Integer = 1 To xlWks.Range(MyInput).Column
            hList.Add(xlWks.Cells(1, x).Value)
        Next

        'sw.WriteLine(String.Join(",", hList.ToArray()))
        sList.AppendLine(String.Join(",", hList.ToArray()))
        'sList.Add(temp)


        'Import in csv file
        Using MyReader As New  _
            Microsoft.VisualBasic.FileIO.TextFieldParser(strFullPath)
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(",")

            'Set the array
            Dim currentRow() As String
            Dim MArray(limit, InputColumn) As String
            Dim SFColumn As String = num2col(InputColumn + 1)
            Dim flag As Boolean

            If pForm.headerCheckbox.Checked Then
                flag = True
            Else
                flag = False
            End If


            'This While Loop Grabs batches specified by 'limit'
            progress.Show()
            While Not MyReader.EndOfData
                i = 0
                While i < limit And Not MyReader.EndOfData
                    If (flag) Then
                        currentRow = MyReader.ReadFields ' Get rid of headers
                        flag = False
                    End If
                    currentRow = Nothing
                    currentRow = MyReader.ReadFields
                    For j As Integer = 0 To currentRow.Length - 1
                        MArray(i, j) = currentRow(j)
                    Next
                    i += 1
                    progress.ProgressBar1.Value += 1
                End While

                xlWkb.Sheets("Master").Range("A3:" & num2col(InputColumn) & i + 2).Value = MArray


                '.....................................FormulaDRAG.................................................
                xlWkb.Sheets("Master").Range(SFColumn & "2:" & MyInput).Resize(i + 1).FillDown()

                'Force formula calculations before copy&paste values
                xlWkb.Sheets("Master").Range(SFColumn & "3:" & column & i + 2).Calculate()

                '......................................ExportDATA.................................................
                For x As Integer = 3 To i + 2
                    rowValues = Nothing
                    If xlWkb.Sheets("Master").Range(pForm.validColumn.Text & x).Text = "TRUE" Then
                        For y As Integer = 1 To xlWkb.Sheets("Master").Range(pForm.exCol.Text).Column 'xlWks.Range(MyInput).Column
                            If y = 1 Then
                                rowValues = xlWkb.Sheets("Master").Cells(x, y).Value
                            Else
                                rowValues = rowValues & "," & xlWkb.Sheets("Master").Cells(x, y).Value
                            End If
                        Next
                        sList.AppendLine(rowValues)
                    End If
                    progress.ProgressBar1.Value += 1
                    'sw.WriteLine(rowValues.Substring(1))
                    'rowValues += Environment.NewLine


                Next


                'Clear PNs from Master
                xlWkb.Sheets("Master").Range("A3:" & column & i + 2).Delete()
                row += i
            End While
            MyReader.Close()
            MyReader.Dispose()
            progress.Close()
        End Using
        File.WriteAllText(oFile, sList.ToString)
        sWatch.Stop()
        XlApp.StatusBar = "Layeth The Smackethdown!"
        MsgBox("Sample Data READY! " & vbCrLf & "Execution Time: " & sWatch.ElapsedMilliseconds / 1000 & " s" & vbCrLf & "Processed " & row - 2 & " records @ " & limit & " per batch")

        'Clean up
        'sw.Flush()
        'sw.Close()
    End Sub

#End Region

#Region "My Functions"
    '......................................Headers..................................................
    ' This function print out the headers that are most common in certain capacitor families. The
    ' purpose of this function is to have a convenient way to start their series builds.
    ' *Created by Quang
    '   05/28/2015
    '   #Idea: Let it import headers from a text file, which then get its headers from running a
    '   query on the dev to see which headers are not null for the family.
    '   04/09/2015
    '   -Added CFR
    Public Sub Headers(ByVal control As Office.IRibbonControl)
        Dim Cell_row As Integer
        Dim Cell_column As Integer
        Dim temp() As String
        Dim H_range As Range

        'Get the ID of the selected item
        mCurrentItemID = control.Id

        'Prevent from redefining the Activesheet multiple time during one session
        VariableSetup()

        If IsNothing(style) Then
            style = xlWks.Application.ActiveWorkbook.Styles.Add("NewStyle")
        End If
        style.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.ForestGreen)
        style.Font.Bold = True
        style.Font.Size = 11
        style.HorizontalAlignment = XlHAlign.xlHAlignLeft
        style.Interior.Color = XlColorIndex.xlColorIndexNone

        'Select what headers to load up depending of the family
        Select Case mCurrentItemID
            Case "CCA"
                temp = {"Status", "ProdCat", "Mfg", "Series", "PN", "Alias1", "Alias2", "Alias3", "Alias4", "Alias5", "SuperAlias", "Value", "Value_Spec", "Tol", "Voltage", "Voltage_Spec", "TempRange", "TC", "DielectricStrength", "Description", "Style", "Dielectric", "App", "Termination", "Lead", "RoHS", "FailureRate", "DF1KHz", "IR", "IR_Spec", "Pref", "Misc", "PKind", "PType", "PQty", "drawing_file_name", "series_units", "D", "D_tol", "L", "L_tol", "LL", "LL_tol", "F", "F_tol"}
            Case "CCD"
                temp = {"Status", "ProdCat", "Mfg", "Series", "PN", "Alias3", "Value", "Value_Spec", "Tol", "Voltage", "TempRange", "TC", "DielectricStrength", "Style", "Dielectric", "App", "Termination", "Lead", "RoHS", "FailureRate", "DF", "IR", "IR_Spec", "Pref", "PKind", "PType", "drawing_file_name", "series_units", "D", "D_tol", "T", "T_tol", "S", "S_tol", "LL", "LL_tol", "F", "F_tol", "G", "G_tol"}
            Case "CCR"
                temp = {"Status", "ProdCat", "Mfg", "Series", "PN", "Alias1", "Alias2", "Alias3", "Alias4", "Alias5", "SuperAlias", "Value", "Value_Spec", "Tol", "Voltage", "Voltage_Spec", "Voltage_AC", "TempRange", "TC", "DielectricStrength", "Description", "Style", "Dielectric", "App", "Termination", "Lead", "RoHS", "FailureRate", "DF1KHz", "IR", "IR_Spec", "Pref", "Misc", "PKind", "PType", "PQty", "drawing_file_name", "series_units", "L", "L_tol", "H", "H_tol", "T", "T_tol", "S", "S_tol", "LL", "LL_tol", "F", "F_tol", "G", "G_tol"}
            Case "CCS"
                temp = {"Status", "ProdCat", "Mfg", "Series", "PN", "Alias3", "Value", "Value_Spec", "Tol", "Voltage", "Voltage_Spec", "TempRange", "TC", "DielectricStrength", "Style", "Dielectric", "App", "Termination", "RoHS", "DF", "DF1KHz", "QFactor", "IR", "IR_Spec", "Pref", "Misc", "PKind", "PSize", "PType", "PQty", "drawing_file_name", "series_units", "SizeRef", "L", "L_tol", "T", "T_tol", "B", "B_tol", "W", "W_tol"}
            Case "CFR"
                temp = {"Status", "ProdCat", "Mfg", "Series", "PN", "Value", "Value_Spec", "Tol", "RatedTemp", "Voltage_AC", "TempRange", "Style", "Dielectric", "App", "Lead", "RoHS", "Approvals", "DF1KHz", "MaxDvDt", "IR", "IR_Spec", "Pref", "PKind", "PSize", "PType", "PQty", "drawing_file_name", "series_units", "L", "L_tol", "H", "H_tol", "T", "T_tol", "S", "S_tol", "LL", "LL_tol", "F", "F_tol", "H0", "H0_tol", "Marking"}
            Case "CFS"
                temp = {"Status", "ProdCat", "Mfg", "Series", "PN", "Value", "Value_Spec", "Tol", "RatedTemp", "Voltage", "Voltage_Spec", "Voltage_AC", "TempRange", "Style", "Dielectric", "App", "RoHS", "HiTempSolder", "DF", "DF1KHz", "IR", "IR_Spec", "Pref", "PKind", "PSize", "PQty", "drawing_file_name", "series_units", "SizeRef", "L", "L_tol", "T", "T_tol", "F", "F_tol", "W", "W_tol", "Marking"}
            Case Else
                temp = {"Hello"}
        End Select

        If tLen = 0 Then
            tLen = temp.Length
        End If

        xlWks.Range(xlWks.Cells(1, 1), xlWks.Cells(1, tLen)).Clear()

        If tLen <> temp.Length Then
            tLen = temp.Length
        End If
        Cell_row = 1
        Cell_column = temp.Length

        H_range = xlWks.Range(xlWks.Cells(1, 1), xlWks.Cells(Cell_row, Cell_column))

        H_range.Value = temp
        H_range.Style = "NewStyle"

        'For i = 1 To Cell_column
        'T = xlWks.Cells(1, i).Value
        'If (T = "Mfg" Or T = "Series" Or T = "PN" Or T = "App" Or T = "Values" Or T = "Tol" Or T = "L" Or T = "H" Or T = "T" Or T = "PKind" Or T = "RoHS" Or T = "Style" Or T = "LL" Or T = "Voltage" Or T = "Voltage_AC") Then
        'xlWks.Cells(1, i).Style = "NewStyle"
        'End If
        'xlWks.Cells(1, i).HorizontalAlignment = Constants.xlLeft
        'xlWks.Cells(1, i).Font.Bold = True
        'Next i
        'ReleaseComObject terminate the COM object, so it will allow the object to run once and no more until the application is restarted.
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(XlApp)

    End Sub

    '......................................Vlookup..................................................
    ' This function create a VLOOKUP formula using the current active cell location.
    Public Sub Vlookup(ByVal control As Office.IRibbonControl)
        VariableSetup()
        Dim row As Integer = XlApp.ActiveCell.Row
        Dim column As Integer = XlApp.ActiveCell.Column

        Select Case control.Id
            Case "base"
                XlApp.ActiveCell.Value = "=IF(VLOOKUP($E" & row.ToString & ",Base,MATCH(" & num2col(column) & "$1,Base[#Headers],0),FALSE)=0,"""",VLOOKUP($E" & row & ",Base,MATCH(" & num2col(column) & "$1,Base[#Headers],0),FALSE))"
            Case "universal"
                XlApp.ActiveCell.Value = "=VLOOKUP(" & num2col(column) & "$1,Universal,2,FALSE)"
            Case "series"
                XlApp.ActiveCell.Value = "=VLOOKUP($D" & row & ",Series,MATCH(" & num2col(column) & "$1,Series[#Headers],0),FALSE)"
            Case Else
                XlApp.ActiveCell.Value = "Something Went Wrong!"
        End Select
    End Sub

    '......................................ImportMulCSV..............................................
    ' This function take multiple csv files in sequential naming and import them together.
    ' The purpose of this function is to lessen the time it take for importing multiple parts of 
    ' a large csv file.
    ' *Created by Quang
    '   05/28/2015
    '   !Performance drop significantly when given large files
    '   04/21/2015
    '   -Better UI (Can choose the file now)
    '   -Can choose multiple files instead of iterating
    '   -Can import files with different number of columns
    '   04/10/2015
    '   -Errors in enumeration that causes it not to return all of the values
    Public Sub ImportMulCSV(ByVal control As Office.IRibbonControl)
        VariableSetup()

        Dim strFullPath As String
        Dim intChoice As Integer
        Dim c As Integer = XlApp.ActiveCell.Column
        Dim r As Integer = XlApp.ActiveCell.Row
        Dim currentRow() As String
        Dim FDO = XlApp.FileDialog(Microsoft.Office.Core.MsoFileDialogType.msoFileDialogOpen)
        Dim skipped As Boolean = False

        FDO.AllowMultiSelect = True
        FDO.Title = "Please Choose the file(s) you want to import..."
        FDO.Filters.Clear()
        FDO.Filters.Add("CSV Files", "*.csv")
        intChoice = FDO.Show

        If intChoice <> 0 Then
            For i As Integer = 0 To FDO.SelectedItems.Count
                strFullPath = FDO.SelectedItems(i)

                Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(strFullPath)
                    MyReader.TextFieldType = FileIO.FieldType.Delimited
                    MyReader.SetDelimiters(",")
                    If skipped Then
                        currentRow = MyReader.ReadFields
                    End If
                    While Not MyReader.EndOfData
                        currentRow = MyReader.ReadFields
                        For x As Integer = 0 To currentRow.Length - 1
                            xlWks.Cells(r, c + x).Value = currentRow(x)
                        Next
                        r += 1
                    End While
                    skipped = True
                    MyReader.Close()
                    MyReader.Dispose()
                End Using

            Next
        End If
    End Sub


    '......................................Population................................................
    ' This function take in PNs from a csv file and apply user's formulas to them in batches before
    ' pasting the raw values to a new worksheet. The purpose of this function is to reduce calculation
    ' time when applying formulas to large quantity of PNs.
    ' *Created by Luke and Updated by Quang
    '   04/27/2015
    '   -Bring code to be on par with VBA while not using clipboard
    '   -Able to take in multiple inputs
    '   04/17/2015
    '   -Cleaned up code
    '   -Increased calculation w/o using clipboard
    '   04/14/2015
    '   -Added Progress bar
    '   -Improved performance
    '   04/13/2015
    '   -Optimized code to run 4-5x Faster than VBA code
    '   -Tidied up the codes
    '   04/09/2015
    '   -Replaced ADO COM with TextFieldParser to import csv files
    '   -Added Checks and Option to create 'Master' and 'Export' worksheets
    '   -Added CSV export functionality
    '   -Overall convertion and cleaning of VBA code
    '................................................................................................

    Public Sub Population(ByVal control As Office.IRibbonControl)
        ' And it begins...
        VariableSetup()
        '......................................Configurations.............................................
        Dim row As Integer = 2 'Which row in "Export" will the data start pasting from
        XlApp.ScreenUpdating = True 'Set to 'False' to increase performance
        XlApp.Calculation = XlCalculation.xlCalculationManual 'Set to 'Manual' to increase performance
        XlApp.EnableEvents = False 'Set to 'False' to increase performance
        XlApp.DisplayStatusBar = True
        '...................................../Configurations.............................................


        MsgBox("CHECKLIST:" & vbCrLf & "NO SPECIAL CHARACTERS OR SPACES IN FILE NAME!!" & vbCrLf & "Is Your Formula sheet named 'Master'?" & vbCrLf & "Master Format:" & vbCrLf & _
        "   First row contains Headings" & vbCrLf & "   Second Row contains formulas" & vbCrLf & "  Third Row is blank" & vbCrLf & "Is your export sheet named 'Export'?")

        XlApp.StatusBar = "You will go one on one with the Great One!"

        '......................................Format_Checks...............................................

        Dim sWatch As New Stopwatch 'Keep track of how fast the program is running
        Dim sh As Worksheet
        Dim mflag As Boolean = False
        Dim eflag As Boolean = False
        Dim result As Integer
        Dim strFullPath As String
        Dim tempList As New List(Of String)
        Dim MyInput As String
        Dim i As Integer
        Dim pForm As New PopForm
        Dim err As Range = Nothing
        Dim total As Integer
        Dim current As Integer = 1

        pForm.ShowDialog()

        If pForm.DialogResult = DialogResult.OK Then
            MyInput = pForm.lastCol.Text
            Dim InputColumn = pForm.colNum.Value
            pForm.Close()

            For Each sh In xlWkb.Sheets
                If (sh.Name = "Master") Then
                    mflag = True
                End If
                If (sh.Name = "Export") Then
                    eflag = True
                End If
            Next

            If mflag = False And eflag = False Then
                result = MsgBox("You do not have both the 'Master' and 'Export' sheet!" & vbCrLf & "Do you want me to create it for you?", MsgBoxStyle.YesNo)
                If (result = MsgBoxResult.Yes) Then
                    Dim nWks As Worksheet
                    nWks = CType(xlWkb.Worksheets.Add(After:=xlWks), Worksheet)
                    nWks.Name = "Export"
                    xlWks.Name = "Master"
                Else : Exit Sub
                End If
            ElseIf mflag = True And eflag = False Then
                result = MsgBox("You do not have the 'Export' sheet!" & vbCrLf & "Do you want me to create it for you?", MsgBoxStyle.YesNo)
                If (result = MsgBoxResult.Yes) Then
                    Dim nWks As Worksheet
                    nWks = CType(xlWkb.Worksheets.Add(After:=xlWks), Worksheet)
                    nWks.Name = "Export"
                Else : Exit Sub
                End If
            ElseIf mflag = False And eflag = True Then
                result = MsgBox("You do not have the 'Master' sheet!" & vbCrLf & "Do you want me to create it for you?", MsgBoxStyle.YesNo)
                If (result = MsgBoxResult.Yes) Then
                    xlWks.Name = "Master"
                Else : Exit Sub
                End If
            End If

            xlWkb.Sheets("Master").Activate()

            '...................................../Format_Checks..............................................

            'Import File
            strFullPath = XlApp.GetOpenFilename("Text Files (*.csv),*.csv", , "Please select IMPORT file...")

            XlApp.StatusBar = "If you smell what The Rock is Cooking!"

            If strFullPath = "False" Then Exit Sub 'User pressed Cancel on the open file dialog

            'Copy & Paste Headers from "Master" to "Export"
            Dim column As String = Nothing
            For Each x In MyInput
                If IsNumeric(x) = False Then
                    column += x
                End If
            Next

            xlWkb.Sheets("Export").Range("A1:" & column & "1").Value = xlWkb.Sheets("Master").Range("A1:" & column & "1").Value

            Dim rowValues As New StringBuilder

            '*****************************Creating & Opening Txt File****************************************************
            Dim fi As FileInfo = New FileInfo(strFullPath)
            Dim oFile As String = fi.DirectoryName & "\" & Path.GetFileNameWithoutExtension(fi.Name) & "_populated.csv"


            If File.Exists(oFile) = True Then
                File.Delete(oFile)
            End If

            Dim sw As StreamWriter = File.CreateText(oFile)
            '***********************************************************************************************************

            'Checking the number of lines in the input file
            total = File.ReadAllLines(strFullPath).Length
            If total > 1048575 Then
                MsgBox("There are more rows in the input file than Excel can hold, so the output data will be exported to a CSV file.")
                pForm.exportCheckbox.Checked = True
            End If
            total = Math.Ceiling(total / xLimit)

            sWatch.Start()
            'Change the format of 'Export' sheet to Text
            xlWkb.Sheets("Export").Cells.NumberFormat = "@"


            'Write headers to file
            For y As Integer = 1 To xlWks.Range(MyInput).Column
                If y <> 1 And y <> xlWks.Range(MyInput).Column Then
                    rowValues.Append(",")
                End If
                rowValues.Append(xlWkb.Sheets("Export").Cells(1, y).Value2)
            Next
            sw.WriteLine(rowValues.ToString)
            rowValues.Clear()


            'Import in csv file
            Using MyReader As New  _
                Microsoft.VisualBasic.FileIO.TextFieldParser(strFullPath)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")

                'Set the array
                Dim currentRow() As String
                Dim MArray(limit, InputColumn) As String
                Dim SFColumn As String = num2col(InputColumn + 1)
                Dim flag As Boolean
                If pForm.headerCheckbox.Checked Then
                    flag = True
                Else
                    flag = False
                End If


                'This While Loop Grabs batches specified by 'limit'
                While Not MyReader.EndOfData
                    While row + i < xLimit And Not MyReader.EndOfData
                        XlApp.StatusBar = "Beginning of While Loop"
                        i = 0
                        While i < limit And Not MyReader.EndOfData
                            If (flag) Then
                                currentRow = MyReader.ReadFields ' Get rid of headers
                                flag = False
                            End If

                            currentRow = MyReader.ReadFields
                            For j As Integer = 0 To currentRow.Length - 1
                                MArray(i, j) = currentRow(j)
                            Next
                            i += 1
                        End While

                        xlWkb.Sheets("Master").Range("A3:" & num2col(InputColumn) & i + 2).Value = MArray


                        '.....................................FormulaDRAG.................................................
                        xlWkb.Sheets("Master").Range(SFColumn & "2:" & MyInput).Resize(i + 1).FillDown()

                        'Force formula calculations before copy&paste values
                        xlWkb.Sheets("Master").Range(SFColumn & "3:" & column & i + 2).Calculate()

                        '.....................................Error Checking..............................................
                        err = xlWkb.Sheets("Master").Range(SFColumn & "3:" & column & i + 2).Find("#", , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False)
                        If Not IsNothing(err) Then
                            XlApp.StatusBar = "There is an error!"
                            xlWks.Range(err.AddressLocal).Activate()
                            MsgBox("ERROR in cell: " & err.AddressLocal)
                            XlApp.ScreenUpdating = True
                            XlApp.Calculation = XlCalculation.xlCalculationAutomatic
                            XlApp.EnableEvents = True
                            XlApp.DisplayStatusBar = True
                            Exit Sub
                        End If

                        '......................................ExportDATA.................................................
                        xlWkb.Sheets("Export").Range("A" & row & ":" & column & row + i - 1).Value = xlWkb.Sheets("Master").Range("A3:" & column & 2 + i).Value

                        'Clear PNs from Master
                        xlWkb.Sheets("Master").Range("A3:" & column & i + 2).Delete()

                        row = row + i
                        XlApp.StatusBar = "End of While Loop"
                    End While

                    '...........................................Export_To_CSV.................................................
                    If pForm.exportCheckbox.Checked Then
                        Dim progress = New loading_bar
                        progress.Show()
                        progress.ProgressBar1.Minimum = 0
                        progress.ProgressBar1.Maximum = row - 1
                        progress.total.Text = String.Concat(current, "/", total)

                        For x As Integer = 2 To row - 1
                            For y As Integer = 1 To xlWks.Range(MyInput).Column
                                If y <> 1 And y <> xlWks.Range(MyInput).Column Then
                                    rowValues.Append(",")
                                End If
                                rowValues.Append(xlWkb.Sheets("Export").Cells(x, y).Value2)
                            Next
                            sw.WriteLine(rowValues.ToString)
                            progress.ProgressBar1.Value += 1
                            rowValues.Clear()
                        Next
                        'Clean up
                        progress.Close()

                        'Reset row
                        xlWkb.Sheets("Export").Range("A2:" & column & row).Delete()

                    Else
                        MsgBox("ERROR: Can't hold anymore data in Export.")
                        Exit Sub
                    End If

                    current += 1
                    row = 2
                End While
                MyReader.Close()
                MyReader.Dispose()
            End Using

            sWatch.Stop()
            XlApp.StatusBar = "Layeth The Smackethdown!"
            MsgBox("Sample Data READY! " & vbCrLf & "Execution Time: " & sWatch.ElapsedMilliseconds / 1000 & " s" & vbCrLf & "Processed " & row - 2 & " records @ " & limit & " per batch")

            If pForm.exportCheckbox.Checked = False Then
                File.Delete(oFile)
            End If
            sw.Flush()
            sw.Close()
        End If

        'Clean up

        XlApp.ScreenUpdating = True
        XlApp.Calculation = XlCalculation.xlCalculationAutomatic
        XlApp.EnableEvents = True
        XlApp.DisplayStatusBar = True
    End Sub

    '..........................................Validator................................................
    ' This function take in inputs from a CSV file and apply the formulas from the worksheet(s) to them.
    ' The ouputs are then checked whether or not they are valid (using the valid column) and write to a
    ' CSV file.
    ' *Created by Luke and Updated by Quang
    '   04/27/2015
    '   -New Goal: Add stop functionality to the "Stop" button (Using backgroundWorker)
    '   04/24/2015
    '   -Added extra options in the window form
    '   -Used File.WriteAllText rather than StreamWriter to increase performance
    '   -Split Validator into its own function rather than combined with Population
    '   -Replaced List rather with StringBuilder to increase writing speed
    '................................................................................................
    Public Sub Validator(ByVal control As Office.IRibbonControl)
        VariableSetup()

        '......................................Configurations.............................................
        Dim row As Integer = 2 'Which row in "Export" will the data start pasting from
        XlApp.ScreenUpdating = True 'Set to 'False' to increase performance
        XlApp.Calculation = XlCalculation.xlCalculationManual 'Set to 'Manual' to increase performance
        XlApp.EnableEvents = False 'Set to 'False' to increase performance
        XlApp.DisplayStatusBar = True
        '...................................../Configurations.............................................

        MsgBox("CHECKLIST:" & vbCrLf & "NO SPECIAL CHARACTERS OR SPACES IN FILE NAME!!" & vbCrLf & "Is Your Formula sheet named 'Master'?" & vbCrLf & "Master Format:" & vbCrLf & _
        "   First row contains Headings" & vbCrLf & "   Second Row contains formulas" & vbCrLf & "  Third Row is blank" & vbCrLf & "Is your export sheet named 'Export'?")

        XlApp.StatusBar = "You will go one on one with the Great One!"

        Dim sWatch As New Stopwatch 'Keep track of how fast the program is running
        Dim sh As Worksheet
        Dim mflag As Boolean = False
        Dim result As Integer
        Dim strFullPath As String
        Dim MyInput As String
        Dim i As Integer
        Dim vForm As New ValForm

        vForm.ShowDialog()

        If vForm.DialogResult = DialogResult.OK Then
            MyInput = vForm.lastCol.Text
            Dim InputColumn = vForm.colNum.Value
            vForm.Close()

            For Each sh In xlWkb.Sheets
                If (sh.Name = "Master") Then
                    mflag = True
                End If
            Next

            If mflag = False Then
                result = MsgBox("You do not have the 'Master' sheet!" & vbCrLf & "Do you want me to create it for you?", MsgBoxStyle.YesNo)
                If (result = MsgBoxResult.Yes) Then
                    xlWks.Name = "Master"
                Else : Exit Sub
                End If
            End If
            xlWkb.Sheets("Master").Activate()

            strFullPath = XlApp.GetOpenFilename("Text Files (*.csv),*.csv", , "Please select IMPORT file...")

            sWatch.Start()
            XlApp.StatusBar = "If you smell what The Rock is Cooking!"

            If strFullPath = "False" Then Exit Sub 'User pressed Cancel on the open file dialog

            Dim rowValues As String
            Dim sList As New System.Text.StringBuilder

            Dim path1 As String = Path.GetTempFileName()
            Dim fi As FileInfo = New FileInfo(strFullPath)
            Dim oFile As String = fi.DirectoryName & "\" & Path.GetFileNameWithoutExtension(fi.Name) & "_validated.csv"
            Dim progress = New loading_bar

            Dim max As Integer = File.ReadLines(strFullPath).Count  'Could severely slow down operation time if use on a large file

            progress.ProgressBar1.Minimum = 0
            progress.ProgressBar1.Maximum = max * 2

            If File.Exists(oFile) = True Then
                File.Delete(oFile)
            End If

            'Dim sw As StreamWriter = File.CreateText(oFile)
            Dim hList As New List(Of String)
            Dim column As String = Nothing

            For Each x In MyInput
                If IsNumeric(x) = False Then
                    column += x
                End If
            Next

            'Write Headers
            For x As Integer = 1 To xlWks.Range(vForm.exCol.Text).Column
                If x <> xlWkb.Sheets("Master").Range(vForm.validColumn.Text & x).Column Then
                    hList.Add(xlWks.Cells(1, x).Value)
                End If
            Next

            sList.AppendLine(String.Join(",", hList.ToArray()))


            'Import in csv file
            Using MyReader As New  _
                Microsoft.VisualBasic.FileIO.TextFieldParser(strFullPath)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")

                'Set the array
                Dim currentRow() As String
                Dim MArray(limit, InputColumn) As String
                Dim SFColumn As String = num2col(InputColumn + 1)
                Dim flag As Boolean

                If vForm.headerCheckbox.Checked Then
                    flag = True
                Else
                    flag = False
                End If


                'This While Loop Grabs batches specified by 'limit'
                progress.Show()
                While Not MyReader.EndOfData
                    i = 0
                    While i < limit And Not MyReader.EndOfData
                        If (flag) Then
                            currentRow = MyReader.ReadFields ' Get rid of headers
                            flag = False
                        End If
                        currentRow = Nothing
                        currentRow = MyReader.ReadFields
                        For j As Integer = 0 To currentRow.Length - 1
                            MArray(i, j) = currentRow(j)
                        Next
                        i += 1
                        progress.ProgressBar1.Value += 1
                    End While

                    xlWkb.Sheets("Master").Range("A3:" & num2col(InputColumn) & i + 2).Value = MArray


                    '.....................................FormulaDRAG.................................................
                    xlWkb.Sheets("Master").Range(SFColumn & "2:" & MyInput).Resize(i + 1).FillDown()

                    'Force formula calculations before copy&paste values
                    xlWkb.Sheets("Master").Range(SFColumn & "3:" & column & i + 2).Calculate()

                    '......................................ExportDATA.................................................
                    For x As Integer = 3 To i + 2
                        rowValues = Nothing
                        If String.IsNullOrEmpty(vForm.validColumn.Text) Then
                            For y As Integer = 1 To xlWkb.Sheets("Master").Range(vForm.exCol.Text).Column
                                If y = 1 Then
                                    rowValues = xlWkb.Sheets("Master").Cells(x, y).Value
                                Else
                                    rowValues = rowValues & "," & xlWkb.Sheets("Master").Cells(x, y).Value
                                End If
                            Next
                            sList.AppendLine(rowValues)
                            progress.ProgressBar1.Value += 1
                        Else
                            If xlWkb.Sheets("Master").Range(vForm.validColumn.Text & x).Text = "TRUE" Then
                                For y As Integer = 1 To xlWkb.Sheets("Master").Range(vForm.exCol.Text).Column
                                    If y = 1 Then
                                        rowValues = xlWkb.Sheets("Master").Cells(x, y).Value
                                    ElseIf y <> xlWkb.Sheets("Master").Range(vForm.validColumn.Text & x).Column Then
                                        rowValues = rowValues & "," & xlWkb.Sheets("Master").Cells(x, y).Value
                                    End If
                                Next
                                sList.AppendLine(rowValues)
                            End If
                            progress.ProgressBar1.Value += 1

                        End If
                    Next


                    'Clear PNs from Master
                    xlWkb.Sheets("Master").Range("A3:" & column & i + 2).Delete()
                    row += i
                End While
                MyReader.Close()
                MyReader.Dispose()
                progress.Close()
            End Using
            File.WriteAllText(oFile, sList.ToString)
            sWatch.Stop()
            XlApp.StatusBar = "Layeth The Smackethdown!"
            MsgBox("Sample Data READY! " & vbCrLf & "Execution Time: " & sWatch.ElapsedMilliseconds / 1000 & " s" & vbCrLf & "Processed " & row - 2 & " records @ " & limit & " per batch")

        Else
            vForm.Close()
        End If

        'Clean up
        XlApp.ScreenUpdating = True
        XlApp.Calculation = XlCalculation.xlCalculationAutomatic
        XlApp.EnableEvents = True
        XlApp.DisplayStatusBar = True
    End Sub

    '......................................Text_To_Columns...........................................
    ' This function take strings in a range and separate them into individual columns by a delimiter.
    ' The purpose of this function is to make parsing dimensions of capacitors much easier.
    ' *Created by Quang
    '   05/28/2015
    '   -Can use selected range along with specifying the range
    '   04/27/2015
    '   -New Goal: Add the ability to take in multiple columns
    '   04/17/2015
    '   -Added an option windows
    Public Sub Text2Column(ByVal control As Office.IRibbonControl)

        VariableSetup()


        Dim txt2col = New Text2Column
        If txt2col.ShowDialog() = DialogResult.OK Then
            Dim TRange
            If txt2col.TRange.Text = "" Then
                TRange = XlApp.Selection
            Else
                TRange = xlWks.Range(txt2col.TRange.Text)
            End If
            Dim DRange = xlWks.Range(txt2col.DRange.Text)
            Dim delim = txt2col.Delimiter.Text
            Dim tempArray As Object(,)
            tempArray = TRange.Value
            If Not IsNothing(TRange.Find(Environment.NewLine, , XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, False)) Then
                TRange.Replace(Environment.NewLine, delim)
            End If

            TRange.TextToColumns( _
              Destination:=DRange, _
              DataType:=XlTextParsingType.xlDelimited, _
              Tab:=False, _
              Semicolon:=False, _
              Comma:=False, _
              Space:=False, _
              Other:=True, _
              OtherChar:=delim)
            TRange.Value = tempArray
        End If
        txt2col.Close()
    End Sub

    '......................................Delete_Empty_Columns.....................................
    ' This function take strings in a range and separate them into individual columns by a delimiter.
    ' The purpose of this function is to make parsing dimensions of capacitors much easier.
    ' *Created by Quang
    '   04/21/2015
    '   -Created
    Public Sub EmpCol(ByVal control As Office.IRibbonControl)
        VariableSetup()
        Dim result As String
        result = MsgBox("Did you delete all of the 'NULL' from your data?", MsgBoxStyle.YesNo)

        If (result = MsgBoxResult.Yes) Then
            Dim lastRow = XlApp.ActiveCell.SpecialCells(XlCellType.xlCellTypeLastCell).Row
            Dim lastColumn = XlApp.ActiveCell.SpecialCells(XlCellType.xlCellTypeLastCell).Column
            Dim col As String
            Dim count As Integer = 1
            While (xlWks.Range(num2col(count) & "1").Value IsNot vbNullString)
                col = num2col(count)
                If XlApp.WorksheetFunction.CountA(xlWks.Range(col & "2:" & col & lastRow)) = 0 Then
                    xlWks.Columns(count).EntireColumn.Delete()
                    count -= 1
                End If
                count += 1
            End While
        End If
    End Sub

    '......................................Compare Values...........................................
    ' This function connect to the Dev database and get the PN column. The data is then pasted to
    ' the specified column. The purpose of this function is to speed up the process of comparing
    ' the new PN and the existing PN.
    ' *Created by Quang
    '   04/27/2015
    '   -Created
    Public Sub CompareValues(ByVal control As Office.IRibbonControl)
        VariableSetup()
        loginInfo()

        Dim conn As New MySqlConnection
        'Dim stm As String = "SELECT PN FROM capacitors WHERE prodcat='CFP' AND mfg='CDE' AND series='947C'"
        Dim dForm As New DataCheck
        dForm.ShowDialog()
        If dForm.DialogResult = DialogResult.OK Then
            Dim col As String = dForm.Col.Text
            'For i As Integer = 0 To dForm.Col.Text
            'If Not IsNumeric(i) Then
            'col += i
            'End If
            'Next
            Dim header As String = xlWks.Range(col & "1").Value
            Dim stm As String = "SELECT DISTINCT " & header & " FROM capacitors WHERE prodcat='" & dForm.ProdCat.Text & "' AND mfg='" & dForm.Mfg.Text & "';"
            Dim dbList As New List(Of String)
            'Dim exArray(,) As String
            Dim lastRow As Integer = xlWks.Range(col & "1").SpecialCells(XlCellType.xlCellTypeLastCell).Row
            Dim connStr As String = "server=10.176.3.13;user=" & username & ";database=dev;port=3306;password=" & password & ";"
            Dim myData As MySqlDataReader
            conn.ConnectionString = connStr
            Try
                If conn.State = Data.ConnectionState.Closed Then
                    XlApp.StatusBar = "Connecting to MySQL..."
                    conn.Open()
                    Dim cmd As MySqlCommand = New MySqlCommand(stm, conn)
                    myData = cmd.ExecuteReader
                    While myData.Read
                        dbList.Add(myData(0).ToString)
                    End While
                    'exArray = xlWks.Range(col & "2:" & col & lastRow).AdvancedFilter(XlFilterAction.xlFilterInPlace, , , True)
                    MsgBox(lastRow)
                End If
            Catch ex As Exception
                XlApp.StatusBar = ex.ToString
            End Try

            conn.Close()
            XlApp.StatusBar = "Done"
        End If
    End Sub

    Public Sub DBTest(ByVal control As Office.IRibbonControl)
        VariableSetup()
        loginInfo()
        Dim conn As New MySqlConnection
        Dim dForm As New DataCheck
        dForm.ShowDialog()
        If dForm.DialogResult = DialogResult.OK Then
            Dim connStr As String = "server=10.176.3.13;user=" & username & ";database=dev;port=3306;password=" & password & ";"
            conn.ConnectionString = connStr
            Try
                If conn.State = Data.ConnectionState.Closed Then
                    XlApp.StatusBar = "Connecting to MySQL..."
                    conn.Open()
                End If
            Catch ex As Exception
                XlApp.StatusBar = ex.ToString
            End Try

            conn.Close()
            XlApp.StatusBar = "Done"
        End If
    End Sub

    '......................................Remove Blanks...................................................
    ' This function connect to the Dev database and get the PN column. The data is then pasted to
    ' the specified column. The purpose of this function is to speed up the process of comparing
    ' the new PN and the existing PN.
    ' *Created by Quang
    '   04/27/2015
    '   -Created
    Public Sub BlanksRemove(ByVal control As Office.IRibbonControl)
        VariableSetup()
        xlWks.Cells.NumberFormat = "@"
        xlWks.Range("C1:D2").Value = xlWks.Range("A1:B2").Value
        xlWks.Range("C3:D4").Value = xlWks.Range("A3:B4").Value
    End Sub

    '......................................Data Checking...................................................
    ' This function do basic data verification:
    '   -Check for duplicates in the PN column
    '   -Deletion of the ID and ExportDate column
    '   -Check for '0' value in all columns with exception to Pref, Marking, and RoHS
    '   -Format checking on Tol, dimension tols, and TempRange (Regex)
    '   -Check for errors (#NA, #VALUE, etc)
    '   %Create one big loop that goes through all of the rows/columns of the worksheet
    '       and have all of the data checkign done in there. That way there won't be multiple
    '       loops running and thus, cut down on calculation time.
    '   !Performance is extremely slow. Might have to use threading or cut down the amount of columns.
    '       %Solution: Use List.Distinct instead
    ' *Created by Quang
    '   05/28/2015
    '   -Created
    Public Sub DataChecking(ByVal control As Office.IRibbonControl)
        VariableSetup()

        Dim lastRow = XlApp.ActiveCell.SpecialCells(XlCellType.xlCellTypeLastCell).Row
        Dim lastColumn = XlApp.ActiveCell.SpecialCells(XlCellType.xlCellTypeLastCell).Column
        Dim regex_tol As RegularExpressions.Regex = New RegularExpressions.Regex("(\d|[-+]).*(%|pF)")
        Dim temp() As String
        Dim header As String
        Dim list As New List(Of String)
        Dim num As Integer
        Dim progress = New loading_bar
        progress.Show()
        progress.ProgressBar1.Minimum = 0
        progress.ProgressBar1.Maximum = lastColumn * lastRow

        For x As Integer = 1 To lastColumn
            header = xlWks.Range(num2col(x) & 1).Value
            progress.ProgressBar1.Value += 1
            list.Clear()
            For y As Integer = 2 To lastRow
                If IsNothing(xlWks.Range(num2col(x) & y).Value) Then
                    list.Add("NULL")
                Else
                    list.Add(xlWks.Range(num2col(x) & y).Value.ToString)
                End If
                progress.ProgressBar1.Value += 1
            Next

            Select Case header
                Case "PN"
                    If list.Distinct.ToArray.Length < list.Count Then
                        MsgBox("There are duplicates in 'PN'.")
                        Exit Sub
                    End If

                Case "id"
                    MsgBox("Please Delete the 'id' column.")
                    Exit Sub
                Case "ImportDate"
                    MsgBox("Please Delete the 'ImportDate' column.")
                    Exit Sub
                Case "Tol"
                    temp = list.Distinct.ToArray
                    For y As Integer = 0 To temp.Length - 1
                        'MsgBox(temp(y))
                        Dim match As RegularExpressions.Match = regex_tol.Match(temp(y))
                        If Not match.Success Then
                            MsgBox("Errors found in 'Tol': " & temp(y))
                            Exit Sub
                        End If
                    Next
                Case "RoHS"
                    temp = list.Distinct.ToArray
                    For y As Integer = 0 To temp.Length - 1
                        If String.Compare(temp(y), "0") <> 0 And String.Compare(temp(y), "1") <> 0 Then
                            MsgBox("Error in 'RoHS'.")
                            Exit Sub
                        End If
                    Next
                Case "marking"
                    temp = list.Distinct.ToArray
                    For y As Integer = 0 To temp.Length - 1
                        If String.Compare(temp(y), "0") <> 0 And String.Compare(temp(y), "1") <> 0 And String.Compare(temp(y), "NULL") <> 0 Then
                            MsgBox("Error in 'marking'.")
                            Exit Sub
                        End If
                    Next
                Case "Pref"
                    temp = list.Distinct.ToArray
                    For y As Integer = 0 To temp.Length - 1
                        If String.Compare(temp(y), "NULL") = 0 Then
                            MsgBox("There is a null cell in 'Pref'.")
                            Exit Sub
                        End If
                    Next
                Case Else
                    temp = list.Distinct.ToArray
                    For y As Integer = 0 To temp.Length - 1
                        If String.Compare(temp(y), "0") = 0 Then
                            MsgBox("Random '0' Value.")
                            Exit Sub
                        ElseIf Integer.TryParse(temp(y), num) Then
                            If num < -1 Then
                                MsgBox("Error in the data.")
                                Exit Sub
                            End If
                        End If
                    Next
            End Select
            XlApp.StatusBar = num2col(x)
        Next
        progress.Close()
        MsgBox("Everything is good to go.")
    End Sub

    '.......................................Concat Cells................................................

    Public Sub ConcatCells(ByVal control As Office.IRibbonControl)
        VariableSetup()
        Dim concatForm = New ConcatForm
        If concatForm.ShowDialog() = DialogResult.OK Then
            Dim TRange = xlWks.Range(concatForm.InputRange.Text)
            Dim DRange = xlWks.Range(concatForm.OutputRange.Text)

            Dim concatArray As Object(,)
            concatArray = TRange.Value
            Dim output As String = Nothing
            Dim elem As String

            Dim delim As String

            If concatForm.Delimiter.Text = "" Then
                delim = Environment.NewLine
            Else
                delim = concatForm.Delimiter.Text
            End If

            For Each elem In concatArray
                If IsNothing(output) Then
                    output = elem
                Else
                    output += delim & elem
                End If
            Next
            DRange.Value = output
        End If
        concatForm.Close()
    End Sub

    Public Sub PullDown(ByVal control As Office.IRibbonControl)
        VariableSetup()
        XlApp.Calculation = XlCalculation.xlCalculationManual 'Set to 'Manual' to increase performance
        Dim p_form As New PullDownForm

        Dim eCol As String = Nothing

        p_form.ShowDialog()



        If p_form.DialogResult = DialogResult.OK Then
            p_form.Hide()
            Dim col() As String = p_form.p_range.Text.Split(New Char() {":"c})
            Dim eColNum As Integer = xlWks.Range(col(0)).Row + p_form.p_limit.Text
            For Each i As Char In col(1)
                If Not IsNumeric(i) Then
                    eCol += i
                End If
            Next
            xlWks.Range(p_form.p_range.Text).Resize(p_form.p_limit.Text).FillDown()

            'Force formula calculations before copy&paste values
            xlWks.Range(col(0) & ":" & eCol & eColNum).Calculate()
        End If
        p_form.Close()
        XlApp.Calculation = XlCalculation.xlCalculationAutomatic
    End Sub
#End Region

    '***********************************************************************************************

#Region "Helpers"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
