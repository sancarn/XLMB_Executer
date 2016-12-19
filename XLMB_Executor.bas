Attribute VB_Name = "Module2"
Sub test()
    'Get commands to run from alternative text of button
    Dim cmds As Range
    Set cmds = Range(ActiveSheet.Shapes(Application.Caller).AlternativeText)
    
    'Loop through all commands and execute
    Dim MI As Object
    Set MI = GetMI()
    
    Dim cmd As Range
    Dim origColor As Long
    For Each cmd In cmds.Cells
        origColor = cmd.Interior.Color
        cmd.Interior.Color = RGB(255, 255, 0)
        
        'Debug.Print "Execute: " & cmd.Value
        Application.Calculate
        MI.Do cmd.Value
        
        DoEvents
        
        cmd.Interior.Color = origColor
        origColor = 0
    Next
End Sub

Function GetMI() As Object
    Dim MapInfo As Object
    On Error Resume Next
        If MapInfo Is Nothing Then Set MapInfo = GetObject(, "MapInfo.Application.x64")
        If MapInfo Is Nothing Then Set MapInfo = GetObject(, "MapInfo.Application")
        If MapInfo Is Nothing Then Set MapInfo = CreateObject("MapInfo.Application.x64"): MapInfo.Visible = 1
        If MapInfo Is Nothing Then Set MapInfo = CreateObject("MapInfo.Application"): MapInfo.Visible = 1
    On Error GoTo 0
    
    'Return retrieved instance of mapinfo
    Set GetMI = MapInfo
End Function

Sub helloworld()
    Dim k As Boolean
    k = True
    If k Then On Error Resume Next
    Debug.Print 1 / 0
End Sub
