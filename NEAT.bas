Attribute VB_Name = "neat"
Sub GetStarted()
'Unprotect sheets and enable editing
'Run this first before editing

ThisWorkbook.IsAddin = False
With ThisWorkbook.Sheets("NEAT")
    .Unprotect "gotcha"
    .Range(Range("22:22"), Range("26:26")).EntireRow.Hidden = False
End With

End Sub
Sub Deploy()
'Run this to "complile" changes for release. Save the new version .xlsm format
    
    With ThisWorkbook.Sheets("NEAT")
        .Range("Checkem") = 0
        .Range("IsInstaller") = 0
        .Range(Range("22:22"), Range("26:26")).EntireRow.Hidden = True
        .Protect "gotcha"
    End With
    'ThisWorkbook.IsAddin = True
End Sub
Function MyVersion() As String

    MyVersion = "0.92"

End Function
Sub Auto_Open()

    'Sets keyboard shortcuts
    
    Application.OnKey "^+s", "Save"
    Application.OnKey "^+b", "SaveA1Active"
    Application.OnKey "^+v", "JumpVersion"
    Application.OnKey "^+z", "Set_Zoom"
    Application.OnKey "^+h", "CantTouchThis"
    Application.OnKey "^+u", "ShowHidden"
    Application.OnKey "^+n", "NewWindow"
    Application.OnKey "^+x", "NewSheet"
    Application.OnKey "^+e", "HideErrors"
    Application.OnKey "^+i", "DefineInputs"
    Application.OnKey "^+m", "CenterAcross"
    Application.OnKey "^+o", "Opposite"
    Application.OnKey "^+k", "xThousand"
    Application.OnKey "^+j", "divThousand"
    Application.OnKey "^+p", "Pct"
    Application.OnKey "^%0", "AddDecimal"
    Application.OnKey "^%9", "RemoveDecimal"
    Application.OnKey "^%i", "InsertCleanPvt"
    Application.OnKey "^%2", "daNumber"
    Application.OnKey "^%a", "AddAllFieldsValues"
    Application.OnKey "^%s", "PivotFieldsToSum"
    Application.OnKey "^%w", "SetColumnWidth"
    Application.OnKey "^%f", "FreezePanels"
    Application.OnKey "^+{F9}", "CalcRange"
    Application.OnKey "^+.", "Separator"
    
    Application.OnKey "{f1}", ""
    
    'Reads Addin cfg txt file
    MyId
    CheckTest
    
End Sub
Sub Authors(control As IRibbonControl)
    
    About.Show

End Sub

Sub ShortList(control As IRibbonControl)

    Shortcuts.Show

End Sub


'INSTALLER


Option Compare Text 'Makes string comparison non-case sensitive
Sub InstallAddIn()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set currWb = ThisWorkbook
    
    aiName = Sheets("NEAT").Range("AddinName").Value
    If IsEmpty(aiName) Then aiName = "NEAT.XLAM"
    aiPath = Application.UserLibraryPath & aiName
    inx = whichIndex(aiName)  'Grabs add-in index

    If Dir(Application.UserLibraryPath & aiName) = Empty And Not (inx) Then 'Checks if NEAT is installed
        Set dummyWb = Application.Workbooks.Add 'Dummy wb is needed as workspace. Took me a while to figure this out
       
        With currWb
            With .Sheets("NEAT")
                .Unprotect "gotcha"
                .Range("AddinName").Value = aiName 'Store Add-in name
                .Protect "gotcha"
            End With
            .Sheets("Tutorial").Move Before:=dummyWb.Sheets(1)
            .IsAddin = True
            .SaveAs aiPath, FileFormat:=xlOpenXMLAddIn, Filename:="NEAT.XLAM"
        End With
                           
        Set myAi = AddIns.Add(Filename:=aiPath, copyfile:=True) 'Creates add-in
        myAi.Installed = True 'Auto opens with Excel
               
        For Each S In dummyWb.Sheets
            If Not S.Name = "Tutorial" Then S.Delete
        Next S
        
        MsgBox ("Neat successfully installed in:" & vbNewLine & aiPath)
        currWb.Close
    Else
        If MsgBox("NEAT is already installed. Uninstall?", vbYesNo, "NEAT") = vbYes Then RemoveAddIn
    End If
      
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
Sub RemoveAddIn(Optional FromRibbon As Boolean)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    currWs = ThisWorkbook.Name

    If FromRibbon Then aiName = Workbooks("NEAT.XLAM").Sheets("NEAT").Range("AddinName").Value
    If IsEmpty(aiName) Then aiName = "NEAT.XLAM"
    aiPath = Application.UserLibraryPath & aiName
    
On Error GoTo NotFound
        If Dir(Application.UserLibraryPath & aiName) <> Empty Then 'Checks if NEAT is installed
        AddIns(whichIndex(aiName)).Installed = False
        
        Set dummyWb = Application.Workbooks.Add 'Dummy wb is needed as workspace
        Kill aiPath
    
    End If
        
    dummyWb.Close
    Workbooks(currWs).Close
    
    MsgBox ("Neat successfully uninstalled")

NotFound:
    MsgBox ("An error occured. Neat was not uninstalled")
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
Sub RemoveAddinFromRibbon(Optional control As IRibbonControl)
    RemoveAddIn
End Sub
Function whichIndex(aiName)
    whichIndex = False
    i = 0
    
    For Each ai In AddIns
        i = i + 1
'        Debug.Print AddIns(i).Name
        If ai.Name = aiName Then
            whichIndex = i
            Exit For
        End If
    Next ai
End Function
Function BringSheets()
Workbooks("NEAT.xlam").IsAddin = False
End Function


'
'
'
'TXT OPS
'
'
'

Function MyZoom() As String

    Dim Cfg As String
    Dim Text As String
    Dim Textline As String
    Dim Text2 As String
    Dim Cp As Integer
    
'Opens and reads cfg txt file
    Cfg = Application.UserLibraryPath & "\BDHR_Cfg.txt"

Open Cfg For Input As #1

    Do Until EOF(1)
    
        Line Input #1, Textline
        
    Text = Text & Textline & "/"
    
    Loop

Close #1

    Cp = InStr(1, Text, "/:02")
    Text = Right(Text, Len(Text) - Cp - 4)
    Cp = InStr(1, Text, "/:03")
    Text = Left(Text, Cp - 1)

MyZoom = Text

End Function
Function MyId(Optional control As IRibbonControl) As String

    Dim Cfg As String
    Dim Text As String
    Dim Textline As String
    Dim ID As String
    Dim Name As String
    Dim InstallDate As String
    Dim Zoom As String
    Dim Text2 As String
    Dim Cp As Integer
    
    On Error GoTo errhandle

'Opens and reads cfg txt file
    Cfg = Application.UserLibraryPath & "\BDHR_Cfg.txt"

Open Cfg For Input As #1

    Do Until EOF(1)
    
        Line Input #1, Textline
        
    Text = Text & Textline & "/"
    
    Loop

Close #1

'On error, writes new cfg txt file
errhandle:
    If Err.Number <> 0 Then
    
    Name = InputBox("Enter the user's initials", title:="Neat Addin", Default:="User's initials")
    InstallDate = Format(Date, "yymmdd")
    Name = Replace(Name, " ", "")
    Name = Replace(Name, "_", "")
    Name = Replace(Name, ":", "")
    Name = Replace(Name, "/", "")
    
    If Len(Name) = 0 Or Name = "User'sinitials" Then Name = "Noid"
    
    If Len(Name) = 20 Then Name = "DudeOfTwentyInitials" Else Name = UCase(Left(Name, 20)) 'easter egg
    
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ts = fso.CreateTextFile(Application.UserLibraryPath & "\BDHR_Cfg.txt", True)
        ts.WriteLine (":01/" & Name)              'First and last name
        ts.WriteLine (":02/" & 85)                'Default zoom
        ts.WriteLine (":03/" & InstallDate)       'Current date
        ts.WriteLine (":04/0.40")                 'Current version
        ts.Close
        
    Open Cfg For Input As #1

    Do Until EOF(1)
    
        Line Input #1, Textline
        
    Text = Text & Textline & "/"
    
    Loop
    
    Close #1
        
    End If
    
    Text2 = Right(Text, Len(Text) - 4)   'trims out :01/
    Cp = InStr(1, Text2, "/:02")         'finds where /:02 is
    ID = Left(Text2, Cp - 1)
    ID = Replace(ID, " ", "")
    ID = Replace(ID, "_", "")

    MyId = ID
 
End Function

Sub ChangeId(Optional control As IRibbonControl)

'edits current user's first and last name

    Dim Name As String
    Dim InstallDate As String
    Dim Zoom As String
    Dim Version As String
    Dim Std As String
    
    Std = MyId

    Name = InputBox("Enter the user's initials", title:="Neat Addin", Default:=MyId)
    InstallDate = Format(Date, "yymmdd")
    Name = Replace(Name, " ", "")
    Name = Replace(Name, "_", "")
    
    Zoom = MyZoom()
    InstallDate = MyInstallDate()
    Version = MyVersion()
    
    If Len(Name) = 0 Or Name = "User'sinitials" Then Name = "Noid"
    
    If Len(Name) = 20 Then Name = "DudeOfTwentyInitials" Else Name = UCase(Left(Name, 20)) 'easter egg
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ts = fso.CreateTextFile(Application.UserLibraryPath & "\BDHR_Cfg.txt", True)
        ts.WriteLine (":01/" & Name)              'First and last name
        ts.WriteLine (":02/" & Zoom)              'Default zoom
        ts.WriteLine (":03/" & InstallDate)       'Current date
        ts.WriteLine (":04/" & Version)           'Current version
        ts.Close
    
End Sub
Function MyInstallDate() As String

    Dim Cfg As String
    Dim Text As String
    Dim Textline As String
    
'Opens and reads cfg txt file
    Cfg = Application.UserLibraryPath & "\BDHR_Cfg.txt"

Open Cfg For Input As #1

    Do Until EOF(1)
    
        Line Input #1, Textline
        
    Text = Text & Textline & "/"
    
    Loop

Close #1

    Cp = InStr(1, Text, "/:03")
    Text = Right(Text, Len(Text) - Cp - 4)
    Cp = InStr(1, Text, "/:04")
    Text = Left(Text, Cp - 1)
    
    MyInstallDate = Text

End Function
Sub CheckTest()
'With ThisWorkbook.Sheets("NEAT")
'   .Unprotect "gotcha"
'
'   If Weekday(Now) = 4 Or .Range("IsInstaller").Value = 0 Then
'
'       If .Range("Checkem").Value = 0 Then
'           CheckVOnline
'           .Range("Checkem").Value = 1
'           .Range("IsInstaller").Value = 1
'       End If
'
'   Else
'   .Range("Checkem").Value = 0
'
'   End If
'
'   .Protect "gotcha"
'End With

End Sub


Sub CheckVOnline()

'Dim title As String
'Dim objHttp As Object
'Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
'objHttp.Open "GET", "http://www.neataddin.com/vcheck", False
'objHttp.Send ""
'
'title = objHttp.ResponseText
'
'If InStr(1, UCase(title), "<TITLE>") Then
'    title = Mid(title, InStr(1, UCase(title), "<TITLE>") + Len("<TITLE>"))
'    title = Mid(title, 1, InStr(1, UCase(title), "</TITLE>") - 1)
'Else
'    title = ""
'End If
'
'If CStr(title) <> CStr(MyVersion) Then MsgBox "A new NEAT version was released! Please visit www.neataddin.com for updating", vbOKOnly, "New version available"

End Sub




'''''''''' SUBS


Sub SaveA1Active(Optional control As IRibbonControl)
    
    Call SaveSub(1)
    
End Sub
Sub Save(Optional control As IRibbonControl)

    Call SaveSub(0)

End Sub

Sub JumpVersion(Optional control As IRibbonControl)
    
    Call SkipVersion(0)
    
End Sub
Sub SaveSub(A1Active As Integer)

    Dim dlgSaveAs As FileDialog
    Dim dlgFile As Object
    Dim mys As String, myPath As String, FileExt As String, CurrentDate As String, _
    Cp1 As String, Cp2 As String, cp3 As String, DateStr As String, docName As String, versionStr As String, _
    subversionStr As String, zeroFlag As String, myFileName As String
    Dim versionVal As Integer, subversionVal As Integer
    Dim ws As Worksheet
    Dim FName As String
    Dim chk As String
    
    CurrentDate = Format(Date, "yymmdd")
    
'Activates A1 cell in all worksheets
If A1Active Then
    
    Application.ScreenUpdating = False
    For Each ws In ActiveWorkbook.Worksheets
      ws.Activate
      Range("A1").Select
    Next ws
    Sheets(1).Activate
    Application.ScreenUpdating = True

End If

'Creates the string for the new name
    On Error GoTo errhandle
    mys = ActiveWorkbook.Name
    myPath = ActiveWorkbook.Path

'DH 08/04/2018
'Implemented to avoid ocasional errors occuring on retrieving the file format
    If InStrRev(mys, ".") = 0 Then
        FileExt = GetFormatType()
        Else
            FileExt = Right(mys, Len(mys) - InStrRev(mys, ".") + 1)
    End If

'Grabs initials from Config file
    MyInit = MyId()
    
'Checks if the filename is already standardized
    chk = CheckFileName()
    Cp1 = InStrRev(mys, "_v")
    
    If chk = 1 Then
    
        Cp2 = InStr(Cp1, mys, ".")
        cp3 = InStr(Cp2, mys, "_")
        
        'Check if there are 6 numbers on the date
            DateStr = 7
            docName = Mid(mys, DateStr, Cp1 - DateStr + 1)
        
        versionStr = Mid(mys, Cp1 + 2, Cp2 - Cp1 - 2)
        
        If cp3 <> 0 Then                            'Check if version is standardized "(v1)"
            subversionStr = Mid(mys, Cp2 + 1, cp3 - Cp2 - 1)
        Else
            subversionStr = "1"
        End If
        
        versionVal = Val(versionStr)
        subversionVal = Val(subversionStr) + 1
        
        myFileName = myPath & "\" & CurrentDate & docName & "v" & versionVal & "." & subversionVal & "_" & MyInit & FileExt
                
                
'If not standardized open the dialog box

    Else
    
        'Requests and formats user input for File Name
        FName = InputBox("Enter file name", title:="Neat Addin", Default:="File name")
        FName = Replace(FName, " ", ".")
        FName = Replace(FName, "_", ".")
        FName = UCase(FName)
        
        If Len(FName) = 0 Or FName = "FILE.NAME" Then FName = "UNTITLED"
        
        Set dlgFile = Application.FileDialog(msoFileDialogSaveAs)
        
        dlgFile.title = "Neat Addin"
        
        With dlgFile
            .InitialFileName = CurrentDate & "_" & FName & "_v1.1_" & MyInit & FileExt
            
            If .Show = -1 Then
                myFileName = CurrentDate & "_" & FName & "_v1.1_" & MyInit & FileExt ' .SelectedItems(1)
            End If
        End With
        
        Set dlgFile = Nothing
  
    'Name checking ends
    End If
    
'Save file
ActiveWorkbook.SaveAs Filename:=myFileName ', FileFormat:=ActiveWorkbook.FileFormat
       
'If error displays
errhandle:
    If Err.Number <> 0 Then MsgBox "The file was not saved correctly", , "Neat Addin"

End Sub
Sub SkipVersion(A1Active As Integer)

    Dim dlgSaveAs As FileDialog
    Dim dlgFile As Object
    Dim mys As String, myPath As String, FileExt As String, CurrentDate As String, _
    Cp1 As String, Cp2 As String, cp3 As String, DateStr As String, docName As String, versionStr As String, _
    subversionStr As String, zeroFlag As String, myFileName As String
    Dim versionVal As Integer, subversionVal As Integer
    Dim ws As Worksheet
    Dim FName As String

    CurrentDate = Format(Date, "yymmdd")
    
'Activates A1 cell in all worksheets
If A1Active Then
    
    Application.ScreenUpdating = False
    For Each ws In ActiveWorkbook.Worksheets
      ws.Activate
      Range("A1").Select
    Next ws
    Sheets(1).Activate
    Application.ScreenUpdating = True

End If

'Creates the string for the new name
    On Error GoTo errhandle
    mys = ActiveWorkbook.Name
    myPath = ActiveWorkbook.Path

'DH 08/04/2018
'Implemented to avoid ocasional errors occuring on retrieving the file format
    If InStrRev(mys, ".") = 0 Then
        FileExt = GetFormatType()
        Else
            FileExt = Right(mys, Len(mys) - InStrRev(mys, ".") + 1)
    End If
    
'Grabs initials from Config file
    MyInit = MyId()

'Checks if the filename is already standardized
    chk = CheckFileName()
    Cp1 = InStrRev(mys, "_v")
    
    If chk = 1 Then
    
        Cp2 = InStr(Cp1, mys, ".")
        cp3 = InStr(Cp2, mys, "_")
        
        'Check if there are 6 numbers on the date
            DateStr = 7
            docName = Mid(mys, DateStr, Cp1 - DateStr + 1)
        
        versionStr = Mid(mys, Cp1 + 2, Cp2 - Cp1 - 2)
        
        If cp3 <> 0 Then                            'Check if version is standardized "(v1)"
            subversionStr = Mid(mys, Cp2 + 1, cp3 - Cp2 - 1)
        Else
            subversionStr = "1"
        End If
        
        versionVal = Val(versionStr) + 1
        subversionVal = 1
        
        myFileName = myPath & "\" & CurrentDate & UCase(docName) & "v" & versionVal & "." & subversionVal & "_" & MyInit & FileExt
                
'Save file
        ActiveWorkbook.SaveAs Filename:=myFileName, FileFormat:=ActiveWorkbook.FileFormat
                        
'If not standardized open the dialog box

    Else
    
        'Requests user to properly save standardized file
        MsgBox "File must be in standardized file name format in onder to jump version", , "Neat Addin"
  
    'Name checking ends
    End If
    
       
'If error displays
errhandle:
    If Err.Number <> 0 Then MsgBox "The file was not saved correctly", , "Neat Addin"
    
End Sub
Function CheckFileName() As Integer

    Dim Name As String          'filename
    Dim ULCount As Integer      'number of "_" occurencies
    Dim Okeys As Integer        'okeys in check
    Dim NotOkeys As Integer     'fails in check
    Dim SpaceCount As Integer   'number of space occurencies
    Dim SecUnd As String        'what follows second underline
    Dim SecUndNoId As String    'what follows second underline but w/o ID
    Dim v As String             'the version (ex: 2.11 - no v)

    Name = ActiveWorkbook.Name
    Okeys = 0
    NotOkeys = 0

'Checks if "_" occurs exactly three times in the filename string

    ULCount = Len(Name) - Len(Replace(Name, "_", ""))

    If ULCount = 3 Then Okeys = Okeys + 1 Else NotOkeys = NotOkeys + 1

'Checks if " " (space) does not occurs in the filename w/o ID

    SpaceCount = Len(Left(Name, InStrRev(Name, "_"))) - Len(Replace(Left(Name, InStrRev(Name, "_")), " ", ""))

    If SpaceCount = 0 Then Okeys = Okeys + 1 Else NotOkeys = NotOkeys + 1
    
'Checks if the 7th char is "_"

    If Right(Left(Name, 7), 1) = "_" Then Okeys = Okeys + 1 Else NotOkeys = NotOkeys + 1

'Checks if the first six chars are numbers
    
    If IsNumeric(Left(Name, 6)) And Left(Name, 1) <> 0 Then Okeys = Okeys + 1 Else NotOkeys = NotOkeys + 1
    
'Checks if the "v##.##" follows the second "_"

SecUnd = Right(Name, Len(Name) - InStr(InStr(Name, "_") + 1, Name, "_"))
    
    SecUndNoId = Left(SecUnd, InStr(SecUnd, "_"))
    
    If Left(SecUndNoId, 1) = "v" And Right(SecUndNoId, 1) = "_" Then
    
        v = Left(Right(SecUndNoId, Len(SecUndNoId) - 1), Len(Right(SecUndNoId, Len(SecUndNoId) - 1)) - 1)
        
        If IsNumeric(Left(v, InStr(v, ".") - 1)) And IsNumeric(Right(v, Len(v) - InStr(v, "."))) Then
            
            Okeys = Okeys + 1
        
        Else
        
            Fails = Fails + 1
            
        End If
        
    Else
        
        Fails = Fails + 1
    
    End If
    
    If Okeys = 5 Then
        
        CheckFileName = 1
    Else
    
        CheckFileName = 0
    End If

End Function
Function GetFormatType() As String

'DH 08/04/2018
'Implemented to correct occasional error on retrieving file format type

Dim FormatCode As Integer
Dim FileExtension As String

    FormatCode = ActiveWorkbook.FileFormat
    
    Select Case FormatCode
        
        Case Is = -4518, 19, 21
            FileExtension = ".txt"
        
        Case Is = 6, 22, 24, 62
            FileExtension = ".csv"
        
        Case Is = 36
            FileExtension = ".prn"
        
        Case Is = 45
            FileExtension = ".mht"
        
        Case Is = 50
            FileExtension = ".xlsb"
        
        Case Is = 51
            FileExtension = ".xlsx"
        
        Case Is = 52
            FileExtension = ".xlsm"
        
        Case Is = 56
            FileExtension = ".xls"
            
    End Select

    GetFormatType = FileExtension

End Function

'New Window
Sub NewWindow(Optional control As IRibbonControl)

    Dim ws As Worksheet, wSCurrent As Worksheet
    Dim Zoom As String

    ActiveWindow.NewWindow
    Set wSCurrent = ActiveSheet
    Application.ScreenUpdating = False
    
    On Error GoTo ErrHandler
    
    Zoom = MyZoom()
        
    For Each ws In Sheets
        ws.Activate
        ActiveWindow.Zoom = Zoom
        ActiveWindow.DisplayGridlines = False
        MsgBox ws
    Next ws
 
    wSCurrent.Select
    Application.ScreenUpdating = True

    ActiveWorkbook.Windows.Arrange ArrangeStyle:=xlVertical
    
ErrHandler:
    Resume Next

End Sub

'Insert new tab
Sub MasterNewSheet(SheetIndex As Integer)

    Dim Zoom As String

    Application.ThisWorkbook.Sheets(SheetIndex).Copy After:=ActiveSheet
    On Error Resume Next
    ActiveWorkbook.Styles("Normal 2").Delete
    ActiveWorkbook.Styles("Normal 3").Delete
    
    'Sets the font of the entire sheet to default
    With ActiveSheet.Cells.Font
        .Size = ThisWorkbook.Styles("Normal").Font.Size
        .Name = ThisWorkbook.Styles("Normal").Font.Name
    End With
    
    Zoom = MyZoom()
    ActiveWindow.Zoom = Zoom
    
End Sub
Sub NewSheet(Optional control As IRibbonControl)

    'Inserts new blank sheet after the active worksheet
    
    Dim Zoom As String
    
    Zoom = MyZoom()
    
    Application.ScreenUpdating = False
    
    Sheets.Add After:=ActiveSheet
    Columns("A:A").Select
    Selection.ColumnWidth = 1.5
    ActiveWindow.Zoom = Zoom
    
    ActiveWindow.DisplayGridlines = False

    Range("A1").Select
    
    Application.ScreenUpdating = True

End Sub
Sub NewFS(Optional control As IRibbonControl)

    Call MasterNewSheet(2)

End Sub
Sub NewEV1(Optional control As IRibbonControl)

    Call MasterNewSheet(3)

End Sub
Sub NewEV2(Optional control As IRibbonControl)

    Call MasterNewSheet(4)

End Sub
Sub Set_Zoom(Optional control As IRibbonControl)

    'Set all tabs zoom to user given value

    Set ogSheet = ActiveSheet
    Set ogRange = Selection
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    'Dim Z As Integer
    Dim Name As String
    Dim InstallDate As String
    Dim Version As String
    Dim CurZoom As String
    Dim SNarray
    
    On Error GoTo ErrHandler
    
    CurZoom = MyZoom()
    
    Z = InputBox("Please enter % zoom for all worksheets", "Neat Addin", CurZoom)
    If Not IsNumeric(Z) Or Z > 400 Or Z < 10 Then Z = CurZoom
    
    ReDim SNarray(1 To ActiveWorkbook.Sheets.Count)
    For i = 1 To ActiveWorkbook.Sheets.Count
        SNarray(i) = ActiveWorkbook.Sheets(i).Name
    Next
    
    ActiveWorkbook.Sheets(SNarray).Select
    ActiveWindow.Zoom = Z
    
    ogSheet.Select
    ogRange.Select
    
    Application.ScreenUpdating = True
    
    'Edits current user's first and last name

    InstallDate = MyInstallDate()
    Name = MyId()
    Version = MyVersion()
  
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ts = fso.CreateTextFile(Application.UserLibraryPath & "\BDHR_Cfg.txt", True)
        ts.WriteLine (":01/" & Name)                'First and last name
        ts.WriteLine (":02/" & Z)                   'Default zoom
        ts.WriteLine (":03/" & InstallDate)         'Current date
        ts.WriteLine (":04/" & Version)             'Current version
        ts.Close
    
ErrHandler:
    Resume Next
    
End Sub

Sub cantTouchThis(Optional control As IRibbonControl)

    On Error Resume Next
    ActiveSheet.Visible = xlVeryHidden

End Sub
Sub ShowHidden(Optional control As IRibbonControl)

    On Error Resume Next
    For Each S In ActiveWorkbook.Sheets
    
        S.Visible = True
        
    Next S

End Sub
Sub HideAllHeaders(Optional control As IRibbonControl)

    If Application.DisplayFormulaBar Then
    
        With Application
            .DisplayFormulaBar = False
            .ShowWindowsInTaskbar = False
        End With
        
        With ActiveWindow
            .DisplayHorizontalScrollBar = False
            .DisplayVerticalScrollBar = False
            .DisplayWorkbookTabs = False
        End With
        
        ActiveWindow.DisplayHeadings = False
        Application.DisplayFunctionToolTips = False
    
    Else
    
        With Application
            .DisplayFormulaBar = True
            .ShowWindowsInTaskbar = True
        End With
        
        With ActiveWindow
            .DisplayHorizontalScrollBar = True
            .DisplayVerticalScrollBar = True
            .DisplayWorkbookTabs = True
        End With
        
        ActiveWindow.DisplayHeadings = True
        Application.DisplayFunctionToolTips = True
        
        
    End If

End Sub
Sub UnHideAllHeaders(Optional control As IRibbonControl)
Application.ScreenUpdating = False
    If Application.DisplayFormulaBar Then
    
        With Application
            .DisplayFormulaBar = True
            .ShowWindowsInTaskbar = True
        End With
        
        With ActiveWindow
            .DisplayHorizontalScrollBar = True
            .DisplayVerticalScrollBar = True
            .DisplayWorkbookTabs = True
        End With
        
        ActiveWindow.DisplayHeadings = True
        Application.DisplayFunctionToolTips = True
    
    End If
Application.ScreenUpdating = True
End Sub


Sub HideErrors(Optional control As IRibbonControl)

    Application.ScreenUpdating = False
    
    errCellValue = InputBox("Type what you want to appear in error cells?", , "-")
    'If IsNumeric(errCellValue) Then errCellValue = Val(errCellValue)
    
    For Each c In Selection
        If c.HasFormula Then
            f = c.Formula
            f = Right(f, Len(f) - 1)
            If IsNumeric(errCellValue) Then
                c.Formula = "=IFERROR(" & f & "," & errCellValue & ")"
            Else
                c.Formula = "=IFERROR(" & f & "," & Chr(34) & errCellValue & Chr(34) & ")"
            End If
        End If
    Next c
    
    Application.ScreenUpdating = True

End Sub
Sub Separator(Optional control As IRibbonControl)
   
    If Application.UseSystemSeparators = True Then
        With Application
            .UseSystemSeparators = False
            .DecimalSeparator = "."
            .ThousandsSeparator = ","
        End With
    
    Else
        With Application
            .UseSystemSeparators = True
            .DecimalSeparator = ","
            .ThousandsSeparator = "."
        End With
    End If

End Sub
Sub DefineInputs(Optional control As IRibbonControl)

'formats selection as input and changes cell content to values

Dim cell As Range

On Error Resume Next
Application.ScreenUpdating = False

For Each cell In Selection
    Selection.Value = Selection.Value
    Selection.Style = "Inputs"
    Selection.Style = "Input"
    Selection.Style = "Entrada"
Next

End Sub

Sub DateStamp(Optional control As IRibbonControl)
    On Error Resume Next
    Application.ScreenUpdating = False
    ActiveCell.FormulaR1C1 = "=TODAY()"
    ActiveCell.Copy
    ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    'Selection.NumberFormat = "m/d/yyyy"

End Sub

Sub TimeStamp(Optional control As IRibbonControl)
    On Error Resume Next
    Application.ScreenUpdating = False
    ActiveCell.FormulaR1C1 = "=NOW()"
    ActiveCell.Copy
    ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
End Sub
Sub CenterAcross(Optional control As IRibbonControl)
    
    On Error Resume Next

' Center Across Selection

    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlBottom
        .MergeCells = False
    End With

End Sub
Sub FlipH(Optional control As IRibbonControl)

    On Error GoTo EndMacro
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set Rng = Selection
    rw = Selection.Rows.Count
    cl = Selection.Columns.Count
    
    If cl = ActiveCell.EntireRow.Cells.Count Then
        MsgBox "You May Not Select An Entire Row", vbExclamation, _
        "Flip Selection"
        Exit Sub
    End If
    
    If rw = ActiveCell.EntireColumn.Cells.Count Then
        MsgBox "You May Not Select An Entire Column", vbExclamation, _
        "Flip Selection"
        Exit Sub
    End If
    
    ReDim arr(rw, cl)
    For cc = 1 To cl ' = Rng.Columns.Count
        For rr = 1 To rw 'rr = Rng.Rows.Count
            arr(rr, cc) = Rng.Cells(rr, cc) '.Formula
            A = arr(rr, cc)
        Next
    Next
    
    'copy arry to range flippingnhorizontal
    cc = cl
    
    For A = 1 To cl ' to loop the columns
        For rr = 1 To rw 'rr = Rng.Rows.Count
            Rng.Cells(rr, cc) = arr(rr, A) '=  .Formula
        Next
        cc = cc - 1
    Next
    
EndMacro:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub



Sub daNumber()

    Application.ScreenUpdating = False
    On Error GoTo multiStyles
    
    Select Case Selection.Style
    
    Case "Normal"
    Selection.Style = "10^3"
    
    Case "10^3"
    Selection.Style = "10^6"
    
    Case "10^6"
    Selection.Style = "Normal"
    
    Case Else
    Selection.Style = "Normal"
    
    End Select
    
multiStyles:
    If Err.Number <> 0 Then Selection.Style = "Normal"
    Application.ScreenUpdating = True


End Sub
Sub xThousand(Optional control As IRibbonControl)
Application.ScreenUpdating = False
    
    For Each c In Selection
        If c.HasFormula Then
            f = c.Formula
            f = Right(f, Len(f) - 1)
            c.Formula = "=(" & f & ")*1000"
        ElseIf c.Value <> "" And IsNumeric(c) Then
            c.Value = c.Value * 1000
        End If
    Next c

Application.ScreenUpdating = True

End Sub
Sub divThousand(Optional control As IRibbonControl)
Application.ScreenUpdating = False
    
    For Each c In Selection
        If c.HasFormula Then
            f = c.Formula
            f = Right(f, Len(f) - 1)
            c.Formula = "=(" & f & ")/1000"
        ElseIf c.Value <> "" And IsNumeric(c) Then
            c.Value = c.Value / 1000
        End If
    Next c

Application.ScreenUpdating = True
End Sub
Sub Opposite(Optional control As IRibbonControl)
Application.ScreenUpdating = False

    For Each c In Selection
        If c.HasFormula Then
            f = c.Formula
            f = Right(f, Len(f) - 1)
            c.Formula = "=-(" & f & ")"
        ElseIf c.Value <> "" And IsNumeric(c) Then
            c.Value = c.Value * -1
        End If
    Next c
    
Application.ScreenUpdating = True
End Sub
Sub Pct(Optional control As IRibbonControl)
Application.ScreenUpdating = True

    For Each c In Selection
    ognf = c.NumberFormat
        If InStr(c.NumberFormat, "%") Then
            If c.HasFormula Then
                f = c.Formula
                f = Right(f, Len(f) - 1)
                c.Formula = "=(" & f & ")*10000"
            ElseIf c.Value <> "" And IsNumeric(c) Then
                c.Value = c.Value * 10000
            End If
            newnf = Replace(ognf, "%", """bps""")
            c.NumberFormat = newnf
        ElseIf InStr(c.NumberFormat, "bps") Then
            If c.HasFormula Then
                f = c.Formula
                f = Right(f, Len(f) - 1)
                c.Formula = "=(" & f & ")/100"
            ElseIf c.Value <> "" And IsNumeric(c) Then
                c.Value = c.Value / 100
            End If
            newnf = Replace(ognf, """bps""", "")
            c.NumberFormat = newnf
        Else
            If c.HasFormula Then
                f = c.Formula
                f = Right(f, Len(f) - 1)
                c.Formula = "=(" & f & ")/100"
            ElseIf c.Value <> "" And IsNumeric(c) Then
                c.Value = c.Value / 100
            End If
            c.NumberFormat = "0%"
        End If
    Next c
    
Application.ScreenUpdating = True
End Sub
Sub AddDecimal(Optional control As IRibbonControl)
For Each c In Selection

    daStr = Split(c.NumberFormat, ";")
    
    For i = LBound(daStr) To UBound(daStr)
        If InStr(daStr(i), ".") Then
            daStr(i) = Replace(daStr(i), ".", ".0")
        ElseIf InStr(daStr(i), "0") Then
            daStr(i) = StrReverse(Replace(StrReverse(daStr(i)), StrReverse("0"), StrReverse("0.0"), , 1))
        ElseIf daStr(i) = "General" Then
            daStr(i) = "0.0"
        End If
    Next i
    
    c.NumberFormat = Join(daStr, ";")

Next c
End Sub
Sub RemoveDecimal(Optional control As IRibbonControl)
For Each c In Selection

    daStr = Split(c.NumberFormat, ";")
    
    For i = LBound(daStr) To UBound(daStr)
        If InStr(daStr(i), ".00") Then
            daStr(i) = Replace(daStr(i), ".00", ".0")
        ElseIf InStr(daStr(i), ".0") Then
            daStr(i) = StrReverse(Replace(StrReverse(daStr(i)), StrReverse(".0"), StrReverse(""), , 1))
        End If
    Next i

    c.NumberFormat = Join(daStr, ";")

Next c
End Sub
Sub MultiplyByRange(Optional control As IRibbonControl)
    Dim CopyRange, PasteRange As Range
        
    On Error GoTo NoSelection
   
    Set CopyRange = Application.InputBox( _
        Prompt:="Select range to be multiplied", _
        title:="", _
        Default:=Selection.Address, _
        Type:=8)

    daSize = 0
    Do While daSize <> 2
    Set PasteRange = Application.InputBox( _
        Prompt:="Select the multiplier (a single cell)", title:="", Default:="", Type:=8)
        
        daSize = PasteRange.Columns.Count + PasteRange.Rows.Count
        If daSize <> 2 Then
            MsgBox "Please select just one cell"
        End If
    Loop
    
Application.ScreenUpdating = False
    If Not CopyRange Is Nothing And Not PasteRange Is Nothing Then
        
        For Each c In CopyRange
            If Not IsEmpty(c.Value) Then
                If c.HasFormula Then
                    c.Formula = "=(" & Mid(c.Formula, 2) & ")*" & PasteRange.Address
                    ElseIf IsNumeric(c.Value) Then
                        c.Formula = "=(" & c.Formula & ")*" & PasteRange.Address
                End If
            End If
        Next c
    End If

Application.ScreenUpdating = True

NoSelection:
End Sub



Sub TraceAllPrec(Optional control As IRibbonControl)

    Application.ScreenUpdating = False
    
    For Each c In Selection
    c.ShowPrecedents
    
    Next c
    Application.ScreenUpdating = True

End Sub
Sub TraceAllDep(Optional control As IRibbonControl)
    
    Application.ScreenUpdating = False
    
    For Each c In Selection
    c.ShowDependents
    
    Next c
    Application.ScreenUpdating = True

End Sub

Sub CleanArrows(Optional control As IRibbonControl)

Application.CommandBars.ExecuteMso ("TraceRemoveAllArrows")

End Sub


Sub InsertCleanPvt(Optional control As IRibbonControl)

    On Error GoTo Errormask

'Turn selection into a readable string for PivotCaches
    srcData = ActiveSheet.Name & "!" & Selection.Address(ReferenceStyle:=xlR1C1)

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcData) _
    .CreatePivotTable TableDestination:=""

    Application.ScreenUpdating = False
    ActiveWindow.Zoom = 85
    ActiveWindow.DisplayGridlines = False
    
'Loop to rename sheet to PVT(i)
    i = 0
    Do
        Err.Clear
        i = i + 1
        sName = "PVT(" & i & ")"
        On Error Resume Next
        ActiveSheet.Name = sName
    Loop Until Err.Number = 0

'Clean PVT
    Set PT = ActiveSheet.PivotTables(1)
    PT.ManualUpdate = True
            
    With PT
        .InGridDropZones = False
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
        .TableStyle2 = "None"
        .ColumnGrand = False
        .RowGrand = False
        .HasAutoFormat = False
    End With
    
Application.CommandBars.ExecuteMso ("PivotTableSubtotalsDoNotShow")
    
    Exit Sub

Errormask:
    MsgBox "Please select a valid range"

End Sub

Sub CleanPivotTable(Optional control As IRibbonControl)

    Dim PT As PivotTable
    On Error GoTo Errormask
      
    Set PT = ActiveCell.PivotTable
        
    Application.ScreenUpdating = False
    ActiveWindow.Zoom = MyZoom()
    ActiveWindow.DisplayGridlines = False
       
    With PT
        .InGridDropZones = False
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
        .TableStyle2 = "None"
        .ColumnGrand = False
        .RowGrand = False
        .HasAutoFormat = False
    End With
    
    Application.CommandBars.ExecuteMso ("PivotTableSubtotalsDoNotShow")

Exit Sub

Errormask:
    MsgBox "The active cell must be over the Pivot Table to be cleaned"

End Sub

Sub AddAllFieldsValues(Optional control As IRibbonControl)

'Add all fields to the Pivot Table Values
    Dim PT As PivotTable
    Dim PF As PivotField
    Dim i As Long
    
    On Error GoTo Errormask
    
    Set PT = ActiveCell.PivotTable
 
        For Each PF In PT.PivotFields
            PF.Orientation = xlHidden
        Next PF
         
        For Each PF In PT.PivotFields
            PF.Orientation = xlDataField
        Next PF

    Exit Sub

Errormask:
    MsgBox "The active cell must be over the Pivot Table where fields will be added"
    
End Sub

Sub PivotFieldsToSum(Optional control As IRibbonControl)

' Cycles through all pivot data fields and sets to sum
    Dim PF As PivotField
    On Error GoTo Errormask
    
    
    With Selection.PivotTable
        .ManualUpdate = True
    For Each PF In .DataFields
    
    With PF
        .Function = xlSum
        .NumberFormat = "#,##0"
    End With
    
    Next PF
        .ManualUpdate = False
    End With
    
    Exit Sub

Errormask:
    MsgBox "The active cell must be over the Pivot Table where fields will be converted"
    
    
End Sub


'''''''''' UDFS AND NON-RIBBON SUBS

Function WAVG(Values As Range, Weights As Range)
    
    MyCounter = 0
    mysumproduct = 0
    MySum = 0
    For Each MyValue In Values
        MyCounter = MyCounter + 1
        mysumproduct = mysumproduct + (MyValue * Weights(MyCounter))
        MySum = MySum + Weights(MyCounter)
    Next
           
    WAVG = mysumproduct / MySum

End Function
Function ReturnLastWord(The_Text As String)

Dim stGotIt As String
    
    stGotIt = StrReverse(The_Text)
    stGotIt = Left(stGotIt, InStr(1, stGotIt, " ", vbTextCompare))
    ReturnLastWord = StrReverse(Trim(stGotIt))

End Function
Function REVERSE(daText As String)

    REVERSE = StrReverse(Trim(daText))

End Function
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)

End Function
Function DOS(Inventory As Range, SalesRange As Range)

    Dim forecast, Weeks, PartialWeek, LastWeek As Double
    Dim flag52, Solutionflag As Boolean
    Dim cell As Range

    forecast = 0
    Weeks = 0
    LastWeek = 0
    partailweek = 0
    flag52 = False

    If Inventory.Value = 0 Then
        DOS = 0
    Else
        For Each cell In SalesRange
            LastWeek = cell.Value
            forecast = forecast + LastWeek
            If forecast > Inventory.Value Then
                Solutionflag = True
                Exit For
            Else
                Weeks = Weeks + 1
                If Weeks >= 52 Then 'greater than 52 weeks -- get out...
                    flag52 = True
                    Exit For
                End If
            End If
    
        Next cell
    
        'Calc DOS
        If flag52 = True Then 'more than a year of inventory
            DOS = "365+"
        ElseIf Solutionflag = False Then  'not enough weeks to consume, so use average demand of those weeks
            If forecast = 0 Then
                DOS = "365+"
            Else
                DOS = Inventory.Value / (forecast / (Weeks * 7))
            End If
        Else
            'Partial Week
            If LastWeek = 0 Then
                PartialWeek = 0
            Else
                PartialWeek = (Inventory.Value - (forecast - LastWeek)) / LastWeek
            End If
    
            DOS = (Weeks + PartialWeek) * 7
    
        End If
    
        If DOS > 365 Then
            DOS = "365+"
        End If

    End If

End Function

Function WOS(Inventory As Range, SalesRange As Range)

    Dim forecast, Weeks, PartialWeek, LastWeek As Double
    Dim flag52, Solutionflag As Boolean
    Dim cell As Range

    forecast = 0
    Weeks = 0
    LastWeek = 0
    partailweek = 0
    flag52 = False

    If Inventory.Value = 0 Then
        WOS = 0
    Else
        For Each cell In SalesRange
            LastWeek = cell.Value
            forecast = forecast + LastWeek
            If forecast > Inventory.Value Then
                Solutionflag = True
                Exit For
            Else
                Weeks = Weeks + 1
                If Weeks >= 26 Then 'greater than 26 weeks -- get out...
                    flag52 = True
                    Exit For
                End If
            End If
    
        Next cell
    
        'Calc DOS
        If flag52 = True Then 'more than a year of inventory
            WOS = "26+"
        ElseIf Solutionflag = False Then  'not enough weeks to consume, so use average demand of those weeks
            If forecast = 0 Then
                WOS = "26+"
            Else
                WOS = Inventory.Value / (forecast / Weeks)
            End If
        Else
            'Partial Week
            If LastWeek = 0 Then
                PartialWeek = 0
            Else
                PartialWeek = (Inventory.Value - (forecast - LastWeek)) / LastWeek
            End If
    
            WOS = Weeks + PartialWeek
    
        End If
    
        If WOS > 26 Then
            WOS = "26+"
        End If
    
    End If

End Function

Sub ClearExcessRowsAndColumns(Optional control As IRibbonControl)
    
    Dim ar As Range, R As Long, c As Long, tr As Long, tc As Long, x As Range
    Dim wksWks As Worksheet, ur As Range, arCount As Integer, i As Integer
    Dim blProtCont As Boolean, blProtScen As Boolean, blProtDO As Boolean
    Dim shp As Shape

    If ActiveWorkbook Is Nothing Then Exit Sub

    On Error Resume Next
    'For Each wksWks In ActiveWorkbook.Worksheets
    With ActiveWorkSheet
        
        Err.Clear
        Set ur = Nothing
        'Store worksheet protection settings and unprotect if protected.
        blProtCont = wksWks.ProtectContents
        blProtDO = wksWks.ProtectDrawingObjects
        blProtScen = wksWks.ProtectScenarios
        wksWks.Unprotect ""
        
        If Err.Number = 1004 Then
            Err.Clear
            MsgBox "'" & wksWks.Name & _
                   "' is protected with a password and cannot be checked." _
                 , vbInformation
        Else
            Application.StatusBar = "Checking " & wksWks.Name & _
                                    ", Please Wait..."
            R = 0
            c = 0

            'Determine if the sheet contains both formulas and constants
            Set ur = Union(wksWks.UsedRange.SpecialCells(xlCellTypeConstants), _
                           wksWks.UsedRange.SpecialCells(xlCellTypeFormulas))
            'If both fails, try constants only
            
            If Err.Number = 1004 Then
                Err.Clear
                Set ur = wksWks.UsedRange.SpecialCells(xlCellTypeConstants)
            End If
            
            'If constants fails then set it to formulas
            If Err.Number = 1004 Then
                Err.Clear
                Set ur = wksWks.UsedRange.SpecialCells(xlCellTypeFormulas)
            End If
            
            'If there is still an error then the worksheet is empty
            If Err.Number <> 0 Then
                Err.Clear
                
                    If wksWks.UsedRange.Address <> "$A$1" Then
                    wksWks.UsedRange.EntireRow.Hidden = False
                    wksWks.UsedRange.EntireColumn.Hidden = False
                    wksWks.UsedRange.EntireRow.RowHeight = _
                    IIf(wksWks.StandardHeight <> 12, 75#, 12, 75#, 13)
                    wksWks.UsedRange.EntireColumn.ColumnWidth = 10
                    wksWks.UsedRange.EntireRow.Clear
                    'Reset column width which can also _
                     cause the lastcell to be innacurate
                    wksWks.UsedRange.EntireColumn.ColumnWidth = _
                    wksWks.StandardWidth
                    'Reset row height which can also cause the _
                     lastcell to be innacurate
                    
                    If wksWks.StandardHeight < 1 Then
                        wksWks.UsedRange.EntireRow.RowHeight = 12,75#
                    
                Else
                        wksWks.UsedRange.EntireRow.RowHeight = _
                        wksWks.StandardHeight
                    End If
                Else
                    Set ur = Nothing
                End If
            End If
            
            'On Error GoTo 0
            If Not ur Is Nothing Then
                arCount = ur.Areas.Count
                
                'determine the last column and row that contains data or formula
                For Each ar In ur.Areas
                    i = i + 1
                    tr = ar.Range("A1").Row + ar.Rows.Count - 1
                    tc = ar.Range("A1").Column + ar.Columns.Count - 1
                    If tc > c Then c = tc
                    If tr > R Then R = tr
                Next
                
                'Determine the area covered by shapes
                'so we don't remove shading behind shapes
                For Each shp In wksWks.Shapes
                    tr = shp.BottomRightCell.Row
                    tc = shp.BottomRightCell.Column
                    If tc > c Then c = tc
                    If tr > R Then R = tr
                Next
                
                Application.StatusBar = "Clearing Excess Cells in " & _
                                        wksWks.Name & ", Please Wait..."
                
                If R < wksWks.Rows.Count Then
                    Set ur = wksWks.Rows(R + 1 & ":" & wksWks.Rows.Count)
                    ur.EntireRow.Hidden = False
                    ur.EntireRow.RowHeight = IIf(wksWks.StandardHeight <> 12, 75#, _
                                                 12, 75#, 13)
                    
                    'Reset row height which can also cause the _
                     lastcell to be innacurate
                    If wksWks.StandardHeight < 1 Then
                        ur.RowHeight = 12,75#
                    Else
                        ur.RowHeight = wksWks.StandardHeight
                    End If
                    
                    Set x = ur.Dependents
                    If 1 = 0 Then
                        ur.Clear
                    Else
                        ur.Delete
                    End If
                End If
                
                If c < wksWks.Columns.Count Then
                    Set ur = wksWks.Range(wksWks.Cells(1, c + 1), _
                                          wksWks.Cells(1, wksWks.Columns.Count)).EntireColumn
                    ur.EntireColumn.Hidden = False
                    ur.ColumnWidth = 18

                    'Reset column width which can _
                     also cause the lastcell to be innacurate
                    ur.EntireColumn.ColumnWidth = _
                    wksWks.StandardWidth

                    Set x = ur.Dependents
                    If Err.Number = 0 Then
                        ur.Clear
                    Else
                        ur.Delete
                    End If
                End If
            End If
        End If
       
       'Reset protection.
        wksWks.Protect "", blProtDO, blProtCont, blProtScen
        Err.Clear
    
    'Next
    End With
    Application.StatusBar = False
    MsgBox "'" & ActiveWorkbook.Name & _
           "' has been cleared of excess formatting." & Chr(13) & _
           "You must save the file to keep the changes.", vbInformation
End Sub

Sub ResetStyles()
'NOT USED IN v1.0
'
'
'
    Dim ss As Style
    Dim MyArray() As String
    Dim i As Long
    
    Application.ScreenUpdating = False
    
    'daWb = Workbooks("NEAT.xlam").Name
    
    'Creates array with target styles
    i = 0
    
    For Each ss In Workbooks("NEAT.XLAM").Styles
        'Debug.Print ss.Name
        ReDim Preserve MyArray(i)
        MyArray(i) = ss.Name
        i = i + 1
    Next ss
      
        
    'Delete each non-compatible style
    For Each ss In ActiveWorkbook.Styles
        If Not IsInArray(ss.Name, MyArray) Then ss.Delete
    Next ss
    
    
    'Paste array in cell range
    'Range("A1:A" & UBound(MyArray) + 1) = WorksheetFunction.Transpose (MyArray)
    
    'Merge styles and color themes
    ActiveWorkbook.Styles.Merge Workbook:=Workbooks("NEAT.xlam")
                       
    'ActiveWorkbook.Theme.ThemeFontScheme.Load _
    '        ("C:\Users\bbagnariolli\AppData\Roaming\Microsoft\Templates\Document Themes\Theme Fonts\Custom 1.xml")
            
    Application.ScreenUpdating = True

End Sub

Sub SetColumnWidth(Optional control As IRibbonControl)

    Application.ScreenUpdating = False
        Cells.EntireColumn.AutoFit
    Application.ScreenUpdating = True
End Sub

Sub FreezePanels()

'freezes/unfreezes panels

On Error GoTo errhandle

    Application.ScreenUpdating = False
        
    If ActiveWindow.FreezePanes = True Then
    
        ActiveWindow.FreezePanes = False
    
    Else
    
        ActiveCell.Select
    
        ActiveWindow.FreezePanes = True

    End If
    
    Application.ScreenUpdating = True

errhandle:
    Resume Next

End Sub


Sub ListSheets()
For Each S In Sheets
Debug.Print S.Name
Next S
End Sub




Sub RedIfNotZero()

cRange = Replace(Selection.Address, "$", "")

cFormula = "=Round(" & cRange & ";2)<>0"

    Selection.FormatConditions.Add Type:=xlExpression, Formula1:=cFormula
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic1
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
  
 
End Sub

Function MRATE(Rate As Variant)
    
MRATE = (1 + Rate) ^ (1 / 12) - 1
           
End Function

Function YRATE(Rate As Variant)
    
YRATE = (1 + Rate) ^ (12) - 1
           
End Function

Sub AttachLabelsToPoints()

   'Dimension variables.
   Dim Counter As Integer, ChartName As String, xVals As String

   ' Disable screen updating while the subroutine is run.
   Application.ScreenUpdating = False

   'Store the formula for the first series in "xVals".
   
For Scounter = 1 To 4

   xVals = ActiveChart.SeriesCollection(Scounter).Formula

   'Extract the range for the data from xVals.
   xVals = Mid(xVals, InStr(InStr(xVals, ","), xVals, _
      Mid(Left(xVals, InStr(xVals, "!") - 1), 9)))
   xVals = Left(xVals, InStr(InStr(xVals, "!"), xVals, ",") - 1)
   Do While Left(xVals, 1) = ","
      xVals = Mid(xVals, 2)
   Loop

   'Attach a label to each data point in the chart.
   For Counter = 1 To Range(xVals).Cells.Count
     ActiveChart.SeriesCollection(Scounter).Points(Counter).HasDataLabel = _
         True
      ActiveChart.SeriesCollection(Scounter).Points(Counter).DataLabel.Text = _
         Range(xVals).Cells(Counter, 1).Offset(0, -1).Value
   Next Counter

Next Scounter

End Sub


Sub Show45Line()

Set daSeries = ActiveChart.SeriesCollection.NewSeries
On Error GoTo 0
'Set daSeries = ActiveChart.SeriesCollection("45degline")

With daSeries
    .Name = "=""45degline"""
    .XValues = "={0,10,100}"
    .Values = "={0,1,10}"
    '.XValues = "={0,1}"
    '.Values = "={0,100}"
    .MarkerStyle = -4142
    .Format.Line.Visible = msoTrue
    .Format.Line.ForeColor.RGB = RGB(191, 191, 191)
    .Format.Line.Weight = 0.75
    .Format.Line.DashStyle = msoLineDash
    End With
    
End Sub


Sub del45line()
While 1
On Error GoTo daExit
ActiveChart.SeriesCollection("45degline").Delete
Wend

daExit:
End Sub


Sub Matritial()

Dim M As Variant
'Redim

Set M = Selection

'M2 = -1 * M
'Debug.Print LBound(M, 1)


For i = 0 To UBound(M, 1)
    For J = LBound(M, 2) To UBound(M, 2)
        M(i, J) = M(i, J) * -1
    Next J
Next i

Range("i13:i22").FormulaArray = M
End Sub



Sub windspeed()

Dim WsRange1 As Range
Dim WsRange2 As Range
Dim lastrowNo As Long
Dim R, J As Long

Dim aryX As Variant

'define the ranges in the worksheet
'Set WsRange1 = Range("f13:g22")
'Set WsRange2 = Range("i13:j22")
Set WsRange1 = Selection
Set WsRange2 = Selection

'transfer the data into a variant array
aryX = WsRange1


'loop throught the array and multiply each element
For R = LBound(aryX, 1) To UBound(aryX, 1)
    For J = LBound(aryX, 2) To UBound(aryX, 2)
'        If aryX(r, j).HasFormula Then
'            f = aryX(r, j).Formula
'            f = Right(f, Len(f) - 1)
'            aryX(r, j) = "=-(" & f & ")"
'        ElseIf aryX(r, j).Value <> "" And IsNumeric(aryX(r, j)) Then
'            aryX(r, j) = aryX(r, j) * -1
'        End If


         aryX(R, J) = aryX(R, J) * -1
         
    Next J
Next R

'transfer variant arry back to sheet
WsRange2.FormulaArray = aryX

End Sub



Sub Calcrange()

Selection.Calculate

End Sub


Function TEXTJOIN(Rng As Range, Optional delimiter As String, Optional ignore_empty As Boolean) As String
Dim compiled As String
For Each cell In Rng
    If ignore_empty And IsEmpty(cell.Value) Then
    'nothing
    Else
    compiled = compiled + IIf(compiled = "", "", delimiter) + CStr(cell.Value)
    End If
Next
TEXTJOIN = compiled
End Function

Function RemoveNotNum(Rng As Range) As String
    xOut = ""
    For i = 1 To Len(Rng.Value)
        xTemp = Mid(Rng.Value, i, 1)
        If xTemp Like "[0-9]" Then
            xStr = xTemp
        Else
            xStr = ""
        End If
        xOut = xOut & xStr
    Next i
    RemoveNotNum = xOut


End Function

Function RemoveChars(Rng As Range) As String
    xOut = ""
    For i = 1 To Len(Rng.Value)
        xTemp = Mid(Rng.Value, i, 1)
        If Not xTemp Like "*[\/:*?<>()|0-9]*" Then
            xStr = xTemp
        Else
            xStr = ""
        End If
        xOut = xOut & xStr
    Next i
    RemoveChars = xOut


End Function


Function RemoveAllButLetters(Rng As Range) As String
    xOut = ""
    For i = 1 To Len(Rng.Value)
        xTemp = Mid(Rng.Value, i, 1)
        If xTemp Like "*[a-z A-Z]*" Then
            xStr = xTemp
        Else
            xStr = ""
        End If
        xOut = xOut & xStr
    Next i
    RemoveAllButLetters = xOut


End Function

Function FINDN(sFindWhat As String, sInputString As String, N As Integer) As Integer
Dim J As Integer

'Application.Volatile
'sInputString = StrReverse(sInputString)

FINDN = 0
    For J = 1 To N
        FINDN = InStr(FINDN + 1, sInputString, sFindWhat)
            If FINDN = 0 Then Exit For
    Next
    
End Function

Function SumByColor(CellColor As Range, rRange As Range)
Dim cSum As Long
Dim ColIndex As Integer
ColIndex = CellColor.Interior.ColorIndex
For Each cl In rRange
  If cl.Interior.ColorIndex = ColIndex Then
    cSum = WorksheetFunction.Sum(cl, cSum)
  End If
Next cl
SumByColor = cSum
End Function
