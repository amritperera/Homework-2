Attribute VB_Name = "Module1"
Sub totalvolume()

'sets labels
Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Volume"

'using challenge method of one click
'declare various counters for loops
Dim countT As Long
Dim countV As Integer
Dim totvol As Double
Dim tic As String
Dim x As Long
Dim obs As Long
Dim y As Long
Dim overflow As String

'get total observations to know how many times to loop
obs = Application.CountA(Range("A:A"))
overflow = Cells(obs, 1).Value
countT = 2
countV = 2
x = 2

'first ticker counter
Do While Cells(countT, 1) = Cells(x, 1)
    totvol = totvol + Cells(countT, 7).Value
    countT = countT + 1
Loop
'outputs totals from the count and the ticker counted
Cells(countV, 9) = Cells(x, 1)
Cells(countV, 12) = totvol


For y = 1 To obs
    'intialize new starting point and reset total volume
    x = countT
    totvol = 0

    Do While Cells(countT, 1) = Cells(x, 1)
        totvol = totvol + Cells(countT, 7)
        countT = countT + 1
    Loop
    'output of the next totals
    countV = countV + 1
    Cells(countV, 9) = Cells(x, 1)
    Cells(countV, 12) = totvol
    
    'check to prevent overflow
    If Cells(x, 1).Value = overflow Then GoTo line1
   
Next y
line1:
Call yearly_percent_change
Call cond_for
Call hardproblem
End Sub

Sub yearly_percent_change()

'setlabels
Dim count3 As Long
Dim x As Long
Dim y As Long
Dim op As Double
Dim clo As Double
Dim yearchange As Double
Dim perchange As Double
Dim obs As Long
Dim error As Integer
error = 0

obs = Application.CountA(Range("A:A"))
count3 = 2
x = 2
y = 2
op = Cells(count3, 3)

'loop for all observations
For Z = 1 To (obs + 30000)

'if statement to navigate to bottom date of ticker
    If Cells(count3, 1) = Cells(x, 1) Then
        count3 = count3 + 1
        
    Else
'calculations and output
        clo = Cells(count3 - 1, 6)
        yearchange = clo - op
        If op = 0 Then
        Cells(y, 11) = error
        GoTo line1
        Else
        
        perchange = (clo / op - 1) * 100
        End If
        
        Cells(y, 11) = perchange
line1:
        op = Cells(count3, 3)
        x = count3
        
        Cells(y, 10) = yearchange
        
        y = y + 1
        
        
    End If
    
Next Z

End Sub

Sub cond_for()
'
' cond_for Macro
'written with record function

'
    Columns("j:j").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("j1").Select
    Selection.Style = "Normal"
    Selection.Style = "Normal"
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("j1").Select
    Selection.FormatConditions.Delete
    Range("L6").Select
End Sub

Sub hardproblem()

'set labels
Cells(1, 15) = "Ticker"
Cells(1, 16) = "Value"
Cells(2, 14) = "Greatest % Increase"
Cells(3, 14) = "Greatest % Decrease"
Cells(4, 14) = "Greatest Total Volume"

'declare some variables
Dim obs As Integer
Dim count4 As Long
Dim x As Double
Dim currenthighP As Double
Dim cuurenthighV As Double
Dim tick As String


count4 = 2
x = 0

obs = Application.CountA(Range("K:K"))
currenthighP = Cells(count4, 11)


'greatest % increase loop
    For i = 1 To obs
    
    If Cells(count4, 11) > x Then
    currenthighP = Cells(count4, 11)
    x = Cells(count4, 11)
    tick = Cells(count4, 9)
    count4 = count4 + 1
    Else
    count4 = count4 + 1
    End If
    
    Next i
    
    Cells(2, 16) = x
    Cells(2, 15) = tick
    
'------------------------------------

count4 = 2
x = 0

currenthighV = Cells(count4, 11)

'greatest % decrease loop
    For i = 1 To obs
    
    If Cells(count4, 11) < x Then
    currenthighP = Cells(count4, 11)
    x = Cells(count4, 11)
    tick = Cells(count4, 9)
    count4 = count4 + 1
    Else
    count4 = count4 + 1
    End If
    
    Next i
    
    Cells(3, 16) = x
    Cells(3, 15) = tick

'------------------------------------
count4 = 2
x = 0

currenthighP = Cells(count4, 12)

'greatest total volume loop
    For i = 1 To obs
    
    If Cells(count4, 12) > x Then
    currenthighV = Cells(count4, 12)
    x = Cells(count4, 12)
    tick = Cells(count4, 9)
    count4 = count4 + 1
    Else
    count4 = count4 + 1
    End If
    
    Next i
    
    Cells(4, 16) = x
    Cells(4, 15) = tick


End Sub




