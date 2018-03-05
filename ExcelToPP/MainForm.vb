Imports System.ComponentModel
Imports Excel = Microsoft.Office.Interop.Excel
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint

Public Class FrmMain
    Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long

    Private ExcApp As Excel.Application
    Private PPApp As PowerPoint.Application

    Private Sub BtnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        Dim ExcWbk As Excel._Workbook = Nothing
        Dim ExcWsh As Excel._Worksheet = Nothing
        Dim ExcName As Excel.Name = Nothing
        Dim ExcRng As Excel.Range = Nothing
        Dim PPPrs As PowerPoint._Presentation = Nothing
        Dim PPSld As PowerPoint._Slide = Nothing
        Dim PPShp As PowerPoint.Shape = Nothing
        Dim PPShpNew As PowerPoint.Shape = Nothing

        Dim repCnt As Long
        Dim cnt As Long
        Dim pth As String
        Dim RngDic As New Dictionary(Of String, Excel.Range)

        Const Wtime As Long = 0.5
        Const ADD_LBL As String = "LCP"
        Const DEL_LBL As String = "DEL"
        Const TMP_LBL As String = "TMP"
        'Or Me.TbPPfile.Text = ""
        'path check 
        If Me.TbExcFile.Text = "" Then
            MsgBox("Excel file path is missing", vbCritical + vbOKOnly, "Error")
            Exit Sub
        Else
            If Not System.IO.File.Exists(Me.TbExcFile.Text) Then
                MsgBox("Excel file dont't exist: " & Me.TbExcFile.Text, vbCritical + vbOKOnly, "Error")
                Exit Sub
            End If
        End If

        If Me.TbPPfile.Text = "" Then
            MsgBox("PowerPoint file path is missing", vbCritical + vbOKOnly, "Error")
            Exit Sub
        Else
            If Not System.IO.File.Exists(Me.TbPPfile.Text) Then
                MsgBox("PowerPoint file dont't exist: " & Me.TbPPfile.Text, vbCritical + vbOKOnly, "Error")
                Exit Sub
            End If
        End If

        ExcApp = New Excel.Application
        PPApp = New PowerPoint.Application

        Me.lblInfo.Text = "Opening Excel file"
        ExcWbk = ExcApp.Workbooks.Open(Me.TbExcFile.Text, False, True,,,, True)
        ExcApp.Visible = False
        Me.lblInfo.Text = "Serching for Excel tables"
        'find workbook names

        For Each ExcName In ExcWbk.Names
            If Mid(ExcName.Name, 1, 3) = ADD_LBL Then RngDic.Add(ExcName.Name, ExcName.RefersToRange)
        Next

        Me.lblInfo.Text = "Opening Power Point file"
        PPPrs = PPApp.Presentations.Open(Me.TbPPfile.Text, True,, True)

        'PPPrs = PPApp.Presentations.Open(Me.TbPPfile.Text, True,, false) ' don't use in 2010
        PPApp.Activate()
        PPApp.ActiveWindow.WindowState = PowerPoint.PpWindowState.ppWindowMinimized

        Me.pbCount.Visible = True
        Me.pbCount.Minimum = 0
        Me.pbCount.Maximum = PPPrs.Slides.Count
        Me.pbCount.Step = 1
        For Each PPSld In PPPrs.Slides
            Me.pbCount.Value = PPSld.SlideIndex
            Me.lblInfo.Text = "Adding tabels to slides"
            PPSld.Select()
            For Each PPShp In PPSld.Shapes
                If RngDic.ContainsKey(PPShp.AlternativeText) = True Then
                    repCnt = repCnt + 1
                    ExcRng = RngDic.Item(PPShp.AlternativeText)

                    'copy to clipboard
                    Clipboard.Clear()
                    ExcRng.Copy()
                    Do
                        AppWait(Wtime)
                    Loop Until ChckClip() <> vbNullString
                    cnt = PPSld.Shapes.Count
                    AppWait(Wtime)

                    PPSld.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteEnhancedMetafile)
                    'PPApp.CommandBars.ExecuteMso("PasteExcelTableSourceFormatting")
                    'PPApp.CommandBars.ExecuteMso("PasteExcelTableDestinationTableStyle")
                    ' PPApp.CommandBars.ReleaseFocus()
                    Do
                        AppWait(Wtime)
                    Loop Until PPSld.Shapes.Count > cnt
                    AppWait(Wtime)
                    Clipboard.Clear()

                    'copy of formats
                    PPShpNew = PPSld.Shapes(PPSld.Shapes.Count)
                    PPShpNew.AlternativeText = Replace(PPShp.AlternativeText, ADD_LBL, TMP_LBL)
                    PPShp.AlternativeText = Replace(PPShp.AlternativeText, ADD_LBL, DEL_LBL)
                    'PPShpNew.Table.ApplyStyle(PPShpTrg.Table.Style.Id, True)
                    'PPShpNew.Table.FirstRow = PPShpTrg.Table.FirstRow
                    PPShpNew.LockAspectRatio = False
                    PPShpNew.Top = PPShp.Top
                    PPShpNew.Left = PPShp.Left
                    PPShpNew.Width = PPShp.Width
                    PPShpNew.Height = PPShp.Height
                End If
            Next PPShp
            'remove old shapes
            For Each PPShp In PPSld.Shapes
                If Mid(PPShp.AlternativeText, 1, 3) = DEL_LBL Then
                    cnt = PPSld.Shapes.Count
                    PPShp.Delete()
                    If PPSld.Shapes.Count = cnt Then
                        PPSld.Shapes(PPSld.Shapes.Count).Delete()
                    End If
                Else
                    PPShp.AlternativeText = Replace(PPShp.AlternativeText, TMP_LBL, ADD_LBL)
                End If
            Next PPShp
        Next

        Me.pbCount.Visible = False
        Me.lblInfo.Text = "Saving Power Point file"

        RngDic.Clear()

        ExcApp.DisplayAlerts = False
        ExcWbk.Close(False)
        ExcApp.DisplayAlerts = True

        pth = Replace(Me.TbPPfile.Text, ".pptx", "_New.pptx")
        PPPrs.SaveAs(pth)
        PPPrs.Close()

        Me.lblInfo.Text = ""
        MsgBox("Operation finished. New powerpoint file has been created. " & vbLf & vbLf & "No. of replaced table: " & repCnt, vbOKOnly + vbInformation, "Done")
        'display PP
        'PPApp.ActiveWindow.WindowState = PowerPoint.PpWindowState.ppWindowMaximized
        'PPPrs.Slides(1).Select()
        'PPApp.Activate()
        'clear
        AppClr()
        ExcRng = Nothing
        ExcWsh = Nothing
        ExcWbk = Nothing
        ExcName = Nothing
        PPShp = Nothing
        PPShpNew = Nothing
        PPSld = Nothing
        PPPrs = Nothing
    End Sub
    Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        TbExcFile.Text = My.Settings.ExcPath
        TbPPfile.Text = My.Settings.PPPath
        lblInfo.Text = ""
    End Sub
    Private Sub FrmMain_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
    End Sub
    Private Sub FrmMain_Activated(sender As Object, e As EventArgs) Handles Me.Activated
    End Sub

    Private Sub TbExcFile_MouseClick(sender As Object, e As MouseEventArgs) Handles TbExcFile.MouseClick
        TbExcFile.Text = GetFile("Excel files|*.*xl*")
        If TbExcFile.Text <> vbNullString Then
            My.Settings.ExcPath = TbExcFile.Text
            My.Settings.Save()
        End If
    End Sub
    Private Sub TbPPfile_MouseClick(sender As Object, e As MouseEventArgs) Handles TbPPfile.MouseClick
        TbPPfile.Text = GetFile("PowerPoint file|*.*pp*")
        If TbPPfile.Text <> vbNullString Then
            My.Settings.PPPath = TbPPfile.Text
            My.Settings.Save()
        End If
    End Sub

    Private Sub AppWait(ByVal seconds As Integer)
        For i As Integer = 0 To seconds * 100
            System.Threading.Thread.Sleep(10)
            Application.DoEvents()
        Next
    End Sub
    Private Sub AppClr()
        Dim ExcWbk As Excel._Workbook
        Dim PPPrs As PowerPoint.Presentation

        If Not ExcApp Is Nothing Then
            For Each ExcWbk In ExcApp.Workbooks
                ExcWbk.Close(False)
            Next ExcWbk
            ExcApp.Quit()
        End If
        If Not PPApp Is Nothing Then
            For Each PPPrs In PPApp.Presentations
                PPPrs.Close()
            Next PPPrs
            PPApp.Quit()
        End If

        'KillProcess(ExcApp.Hwnd)
        'KillProcess(PPApp.HWND)
        ReleaseObject(ExcApp)
        ReleaseObject(PPApp)
        ExcApp = Nothing
        PPApp = Nothing

        For Each prog As Process In Process.GetProcesses
            If prog.ProcessName = "EXCEL" Or prog.ProcessName = "POWERPNT" Then
                prog.Kill()
            End If
            'Debug.Print(prog.ProcessName.ToString)
        Next

    End Sub

    Private Function GetFile(flt As String) As String
        Dim fd As OpenFileDialog = New OpenFileDialog() With {
            .Title = "Select file",
            .InitialDirectory = My.Application.Info.DirectoryPath,
            .Filter = flt,
            .FilterIndex = 1,
            .RestoreDirectory = True
        }

        If fd.ShowDialog() = DialogResult.OK Then
            GetFile = fd.FileName
        End If
    End Function
    Private Function ChckClip() As String
        Dim lst() As String = {"Bitmap", "CommaSeparatedValue", "Dib", "Dif", "EnhancedMetafile", "FileDrop", "Html", "Locale", "MetafilePict", "OemText", "Palette", "PenData", "Riff", "Rtf", "Serializable", "StringFormat", "SymbolicLink", "Text", "Tiff", "UnicodeText", "WaveAudio"}

        For Each el As String In lst
            If Clipboard.ContainsData(el) = True Then
                ChckClip = ChckClip & "|" & el
            End If
        Next
    End Function

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            Dim intRel As Integer = 0
            Do
                intRel = Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            Loop While intRel > 0
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub



    Private Sub KillProcess(hwnd As Long)
        Dim CurrentForegroundThreadID As Long
        Dim strComputer As String
        Dim objWMIService
        Dim colProcessList
        Dim objProcess
        Dim ProcIdXL As Long

        ProcIdXL = 0
        CurrentForegroundThreadID = GetWindowThreadProcessId(hwnd, ProcIdXL)

        strComputer = "."

        objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where ProcessID =" & ProcIdXL)
        For Each objProcess In colProcessList
            objProcess.Terminate
        Next

    End Sub

End Class
