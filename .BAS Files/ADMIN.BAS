' ### ENUMERATE WORKBOOKS ###
Sub EnumerateWorkbooks()
    '...
    EnumerateWorkbooks_
    '...
End Sub

' ### ENUMERATE WORKBOOKS AND ADDINS ###
Sub EnumerateWorkbooksAndAddins()
    '...
    EnumerateWorkbooks_ True
    '...
End Sub

' ### PRINT UNSAVED WORKBOOKS ###
Sub printUnsavedWorkbooks()
    Dim wb As Workbook
    
    Debug.Print ""
    Debug.Print "Unsaved Workbooks:"
    For Each wb In Workbooks
        If Not wb.Saved Then Debug.Print wb.Name
    Next
End Sub

' ### SAVE UNSAVED WORKBOOKS ###
Sub saveUnsavedWorkbooks()
    Dim wb As Workbook
    For Each wb In Workbooks
        If Not wb.Saved Then wb.Save
    Next
End Sub

Private Sub EnumerateWorkbooks_(Optional AlsoAddins As Boolean)
    For i = 1 To 20
        Debug.Print
    Next
    
    Debug.Print "Enumeration for Workbooks:"
    For i = 1 To Workbooks.count
        Debug.Print i & ": " & Workbooks(i).Name
    Next
    
    If AlsoAddins Then
        Debug.Print
        Debug.Print
        
        Debug.Print "Enumeration for Addins: "
        For i = 1 To AddIns.count
            Debug.Print i & ": " & AddIns(i).Name
        Next
    End If
End Sub

Sub SAVEALL_AND_FULL_REBOOT()
    saveAddin
    Dim wb As Workbook
    For Each wb In Workbooks
        wb.Save
        wb.Saved = True
    Next wb
    
    'QUIT THEN OPEN EXCEL
    '********************
    
    'Wait for 3 seconds then open Excel
    Dim CMD As String: CMD = "CMD /S /C PING 10.0.0.0 -n 1 -w 3000 >NUL & Excel.exe"
    Shell CMD, vbMinimizedNoFocus
    
    'Quit Excel
    Application.Quit
End Sub

Private Sub example_saveAddin()
    'Example of saving an addin with module names:
    'Admin, aMain_CreateTable8, buttons, CheckDuplicates, CPearsonArrayFunctions, table8ScheduledChecker and TestAlternateDataStreams
    'and a form 'BigMsgBox'
    
    'Used in Table8 addin
    
    'Be careful not to use this when isAddin = false
    With ThisWorkbook
        Dim alt As Boolean: alt = .IsAddin
        .IsAddin = True
        .Save
        
        
        exportModule "Admin"
        exportModule "aMain_CreateTable8"
        exportModule "buttons"
        exportModule "CheckDuplicates"
        exportModule "CPearsonArrayFunctions"
        exportModule "table8ScheduledChecker"
        exportModule "TestAlternateDataStreams"
        exportForm "BigMsgBox"
        
        .IsAddin = alt
    End With
    
End Sub

Private Sub example_saveAndReboot()
    'Another example for saving and repooting an AddIn.
    saveAddin
    rebootAddin ThisWorkbook
End Sub

Sub rebootAddin(wb As Workbook)
    Dim fName As String: fName = wb.FullName
    Application.DisplayAlerts = False
        Workbooks.Open fName
        MsgBox "AddIn has rebooted."
    Application.DisplayAlerts = True
End Sub

Private Sub exportModuleByName(moduleName As String, Optional wb As Workbook)
    If wb Is Nothing Then Set wb = ThisWorkbook
    wb.VBProject.VBComponents.Item(moduleName).Export wb.path & "\" & wb.Name & " BAS Files\" & moduleName & ".BAS"
End Sub

Private Sub exportFormByName(formName As String, Optional wb As Workbook)
    If wb Is Nothing Then Set wb = ThisWorkbook
    wb.VBProject.VBComponents.Item(formName).Export wb.path & "\" & wb.Name & " BAS Files\" & formName & ".FRM"
End Sub

Sub compareMacroWorkbooks(wb1 As Workbook, wb2 As Workbook)
    'Dim wb1, wb2 As Workbook
    'Set wb1 = Workbooks(1)
    'Set wb2 = Workbooks(4)
    
    Debug.Print "Comparing VBModules from """ & wb1.Name & """ with VBModules from """ & wb2.Name & """"
    
    'If number of modules in WB1 <> number of modules in WB2
    If wb1.VBProject.VBComponents.count <> wb2.VBProject.VBComponents.count Then
        Dim cModuleList As New Collection, i As Integer
        For i = 1 To wb1.VBProject.VBComponents.count
            cModuleList.Add wb1.VBProject.VBComponents(1).Name
        Next
        '...
        'Exit Sub
    End If
    
    For i = 1 To wb1.VBProject.VBComponents.count
        For j = 1 To wb2.VBProject.VBComponents.count
            wb1_ModuleName = wb1.VBProject.VBComponents(i).Name
            wb2_ModuleName = wb2.VBProject.VBComponents(j).Name
            If wb1_ModuleName = wb2_ModuleName Then
                If vbGetLines(wb1.VBProject.VBComponents(i)) <> vbGetLines(wb2.VBProject.VBComponents(j)) Then
                    Debug.Print "Differences found in    " & wb1_ModuleName & "," & wb2_ModuleName
                Else
                    'Debug.Print "No Differences found in " & wb1_ModuleName & "," & wb2_ModuleName
                End If
                GoTo NextI
            End If
        Next j
NextI:
    DoEvents
    Next i
    
    Debug.Print "--- Comparrison finished ---"
End Sub

Sub findTextInWorkbookModule(wb As Workbook, ByVal strToFind As String)
    Dim wb_ModuleName As String
    For i = 1 To wb.VBProject.VBComponents.count
        wb_ModuleName = ""
        wb_ModuleName = wb.VBProject.VBComponents(i).Name
        If InStr(1, vbGetLines(wb.VBProject.VBComponents(i)), strToFind) > 0 Then
            Debug.Print wb.Name & ":" & wb_ModuleName
        End If
    Next
End Sub

Sub Dev_ReferenceInfo()
    'From f_referenceFunctions
    Dim ref As Object
         
     'Loop Through Each Active Reference (Displays in Immediate Window [ctrl + g])
      For Each ref In ThisWorkbook.VBProject.References
        Debug.Print "{" & Join(Array(ref.Description, ref.Name, ref.GUID, ref.FullPath, ref.Major & "." & ref.Minor), ",") & "}"
      Next ref
  
End Sub

Function checkProcExists(wBook As Workbook, sModuleName As String, sProcName As String, Optional lProcType As Long = 1) As Boolean
    ' ===========================================================================
    ' Based off of code from CPearson:
    ' http://www.cpearson.com at http://www.cpearson.com/excel/vbe.aspx
    '
    ' Requires access to VBProject object model
    ' File > Options > Trust Centre > Trust Centre Settings > Macro Settings > Developer Macro Settings > Trust Access to the VBProject Object Model
    '
    ' USAGE:
    ' To check if a procedure exists within a specific module within a specific workbook
    '
    ' ===========================================================================
    Dim oVBProj As Object 'VBIDE.VBProject
    Dim VBComp As Object 'VBIDE.VBComponent
    Dim CodeMod As Object 'VBIDE.CodeModule
    
    Dim ProcName As String
    Dim ProcKind As Variant 'VBIDE.vbext_ProcKind

    checkProcName = False

    Set VBProj = wBook.VBProject
    Set VBComp = VBProj.VBComponents(sModuleName)
    Set CodeMod = VBComp.CodeModule

    With CodeMod
        LineNum = .CountOfDeclarationLines + 1
        Do Until LineNum >= .CountOfLines
            'VBIDE.vbext_ProcKind
            '0 =  vbext_pk_Proc     A Sub or Function Procedure
            '1 =  vbext_pk_Let      A Property Let procedure
            '2 =  vbext_pk_Set      A Property Set procedure
            '3 =  vbext_pk_Get      A Property Get procedure

            ProcName = .ProcOfLine(LineNum, lProcType)
            If ProcName = sProcName Then
                checkProcExists = True
                Exit Do
            End If
            
            'Debug.Print ProcName
            LineNum = .ProcStartLine(ProcName, ProcKind) + .ProcCountLines(ProcName, ProcKind) + 1
        Loop
    End With
End Function


Function checkProcExists_Global(sProcName As String, Optional lProcType As Long = 1) As Boolean
    ' ===========================================================================
    ' Based off of code from CPearson:
    ' http://www.cpearson.com at http://www.cpearson.com/excel/vbe.aspx
    '
    ' Requires access to VBProject object model
    ' File > Options > Trust Centre > Trust Centre Settings > Macro Settings > Developer Macro Settings > Trust Access to the VBProject Object Model
    '
    ' USAGE:
    ' To check if a procedure exists within any of the currently open workbooks
    ' ===========================================================================
    Dim wBook As Workbook
    Dim oVBProj, VBComp, CodeMod As Object
    Dim i As Long
    Dim ProcName As String

    checkProcName = False
    
    For Each wBook In Workbooks
        Set VBProj = wBook.VBProject
        
        'Debug.Assert False
        'Debug.Print wBook.Name
        
        For i = 1 To VBProj.VBComponents.count
            Set VBComp = VBProj.VBComponents(i)
            
            'Debug.Assert False
            'Debug.Print "    " & VBComp.Name
            
            Set CodeMod = VBComp.CodeModule
        
            With CodeMod
                LineNum = .CountOfDeclarationLines + 1
                Do Until LineNum >= .CountOfLines
                    'VBIDE.vbext_ProcKind
                    '0 =  vbext_pk_Proc     A Sub or Function Procedure
                    '1 =  vbext_pk_Let      A Property Let procedure
                    '2 =  vbext_pk_Set      A Property Set procedure
                    '3 =  vbext_pk_Get      A Property Get procedure
        
                    ProcName = .ProcOfLine(LineNum, lProcType)
                    If ProcName = sProcName Then
                        checkProcExists_Global = True
                        Exit Function
                    End If
                    
                    'Debug.Assert False
                    'Debug.Print "    " & "    " & ProcName
                    LineNum = .ProcStartLine(ProcName, lProcType) + .ProcCountLines(ProcName, lProcType) + 1
                Loop
            End With
        Next i
    Next wBook
End Function


Sub AllProceduresToFile(Optional ProcType As Integer = -1)
    ' ===========================================================================
    ' Based off of code from CPearson:
    ' http://www.cpearson.com at http://www.cpearson.com/excel/vbe.aspx
    '
    ' Requires access to VBProject object model
    ' File > Options > Trust Centre > Trust Centre Settings > Macro Settings > Developer Macro Settings > Trust Access to the VBProject Object Model
    '
    'ProcTypes:
    '        0 =  vbext_pk_Proc     A Sub or Function Procedure
    '        1 =  vbext_pk_Let      A Property Let procedure
    '        2 =  vbext_pk_Set      A Property Set procedure
    '        3 =  vbext_pk_Get      A Property Get procedure
    ' ===========================================================================
    Dim wBook As Workbook
    Dim oVBProj, VBComp, CodeMod As Object
    Dim i, n As Long
    Dim ProcTypeName, ProcName  As String
    Dim LineNum, lProcType As Long
    
    'Currently have to do this because I can't figure out how to print properties
    ProcType = 0
    
    'If ProcType Specified then only export single ProcType, else Export them all.
    If ProcType >= 0 Then
        MinProc = ProcType
        MaxProc = ProcType
    Else
        MinProc = 0
        MaxProc = 3
    End If
    
    'Get a free file number
    n = FreeFile
    
    'Ask the user for save path:
    Dim sFilePath As String
    sFilePath = Application.GetSaveAsFilename("MyProcedures.txt", "Text Files, *.txt", , "Save procedures document")
    
    On Error Resume Next
        Kill sFilePath
    On Error GoTo 0
    
    'Open the text file
    Open sFilePath For Append As #n
        'Loop over all open worbooks
        For Each wBook In Workbooks
            'Print the workbook's name
            Print #n, wBook.Name
            
            'Get the workbook's VBProjects
            Set VBProj = wBook.VBProject
            
            'For each VB Component (Module/Class/Sheet) in the VB Project
            For i = 1 To VBProj.VBComponents.count
                'Get the VB Component
                Set VBComp = VBProj.VBComponents(i)
                
                'Print the name of the component
                Print #n, "    " & VBComp.Name
                
                'Get the code for the vb component
                Set CodeMod = VBComp.CodeModule
                
                'With the code
                With CodeMod
                    'Get the total number of declaration lines
                    LineNum = .CountOfDeclarationLines + 1
                    
                    'loop over all specified procedure types
                    For lProcType = MinProc To MaxProc
                        'Get the total number of declaration lines
                        'LineNum = .CountOfDeclarationLines + 1
                        
                        'If the procedure type is <X> print <Y>
                        Select Case lProcType
                            Case 0
                                Print #n, "    " & "    " & "Sub-Routines and Functions:"
                            Case 1
                                Print #n, "    " & "    " & "Let Properties:"
                            Case 2
                                Print #n, "    " & "    " & "Set Properties:"
                            Case 3
                                Print #n, "    " & "    " & "Get Properties:"
                            Case Else
                                Debug.Print "Errors on deck!"
                        End Select
                        
                        'Loop over all procedures
                        Do Until LineNum >= .CountOfLines
                            
                            'Clear old error handling
                            On Error GoTo -1
                            
                            'New Error Handling
                            On Error GoTo NextlProcType
                            
                                'Get the procedure name
                                ProcName = .ProcOfLine(LineNum, lProcType)
                            
                            On Error GoTo 0
                            
                            'Append it to the file
                            Print #n, "    " & "    " & "    " & ProcName
                            
                            'Get next routine
                            LineNum = .ProcStartLine(ProcName, lProcType) + .ProcCountLines(ProcName, lProcType) + 1
                        Loop
NextlProcType:
                    Next lProcType
                End With
            Next i
        Next wBook
    Close #n
End Sub

Sub ExportOpenMacroWorkbooks(exportDir As String)
    
    'if exportDir doesn't end in a "\" then make it
    If Right(exportDir, 1) <> "\" Then exportDir = exportDir & "\"
    
    'Loop over all open worbooks
    For Each wBook In Workbooks
        'Get the workbook's VBProjects
        Set VBProj = wBook.VBProject
        
        'For each VB Component (Module/Class/Sheet) in the VB Project
        For i = 1 To VBProj.VBComponents.count
            'Get the VB Component
            Set VBComp = VBProj.VBComponents(i)
            
            Select Case VBComp.Type
            
                Case 1
                    VBComp.Export exportDir & wBook.Name & ";" & VBComp.Name & ".bas"
                Case 100
                    VBComp.Export exportDir & wBook.Name & ";EventModule - " & VBComp.Name & ".bas"
                Case 11
                    VBComp.Export exportDir & wBook.Name & ";EventModule - " & VBComp.Name & ".cls"
                Case 25
                    VBComp.Export exportDir & wBook.Name & ";EventModule - " & VBComp.Name & ".frm"
                Case Else
                
            End Select
        Next i
    Next wBook
End Sub

Sub timeSub(subName As String, numTrials As Integer)
    Dim i As Integer, iStart, iEnd, dt, old_dt As Double
    'timeSub "Dev_ReferenceInfo" , 1000
    'Time taken for Dev_ReferenceInfo to run: 0.016015625s
    
    'Remember - Excel can optimise it's inner workings. Therefore it may be best to choose large number of trials.
    'Time taken for Dev_ReferenceInfo to run: 0.00833203125s
    
    'Loop over numtests
    For i = 1 To numTrials
        'Time sub routine
        iStart = Timer()
            'Up to 30 args
            Application.Run subName
        iEnd = Timer()
        
        'Get change in time
        dt = iEnd - iStart
        
        old_dt = (old_dt + dt)
    Next
    
    'Average
    old_dt = old_dt / numTrials
    
    Debug.Print "Time taken for " & subName & " to run: " & old_dt & "s"
End Sub




Function vbGetLines(objVBComp As Object) As String
    'Get lines from VBA Module and return as string
    If objVBComp.Type = 1 Then
        objVBComp.Export "C:\temp\" & objVBComp.Name & ".bas"
    Else
        Exit Function
    End If
    
    'Return file contents as string
    vbGetLines = harvestFile("C:\temp\" & objVBComp.Name & ".bas")
    
    'Delete file
    Kill "C:\temp\" & objVBComp.Name & ".bas"
End Function

Private Function makeFileNameLegal(ByVal fileName As String, Optional replaceChar As String = "¦") As String
    'FROM f_directoryFunctions
    
    Const illegalChars = "\/:*?""<>|"
    If InStr(1, illegalChars, replaceChar) Then
        MsgBox "Replacement character """ & replaceChar & """ is illegal", vbExclamation
        Exit Function
    End If
    
    For i = 1 To Len(illegalChars)
        fileName = Replace(fileName, Mid(illegalChars, i, 1), replaceChar)
    Next
    
    makeFileNameLegal = fileName
End Function


Private Function harvestFile(strPath As String) As String
    'FROM f_directoryFunctions
    'Open file and return contents of the file as string
    
    Open strPath For Input As #1
        'read textlines into text
        Do Until EOF(1)
            Line Input #1, textline
            Text = Text & textline
        Loop
    Close #1
    
    harvestFile = Text
End Function



