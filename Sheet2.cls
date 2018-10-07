VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub Worksheet_Activate()
    
    With ThisWorkbook.Worksheets("Summary")
        If Len(Trim(CStr(.Range("S15").Value))) > 0 _
            And IsNumeric(.Range("S15").Value) = True Then
             
            If .Range("S15").Value >= 2015 Then
                ProgramDriver "year"
            Else
                ProgramDriver "normal"
                MsgBox "Couldn't find enough info on year " & .Range("S15").Value
            End If
        
        Else
            
            ProgramDriver "normal"
        
        End If
    End With

End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
    
    Dim KeyCells As Range
    Set KeyCells = ThisWorkbook.Worksheets("Summary").Range("S15")
    
    If Not Application.Intersect(KeyCells, Range(Target.Address)) Is Nothing Then

        Call Worksheet_Activate
       
    End If
    
End Sub


Sub ProgramDriver(indicator As String)

    Dim leadNames() As String
    Dim lastRow As Integer
    Dim i As Integer
    Dim f As Integer
    Dim numberOfLeads As Integer
    Dim lastPipelineRow As Long
    Dim row As Long
    Dim pipelineYear As String: pipelineYear = Year(Date)
    Dim col As Integer
    Dim filterexists As Boolean: filterexists = True
    Dim yearColStart As Integer
    Dim currentMonth As String: currentMonth = Format(Date, "mmm")
    Dim w As Worksheet: Set w = ThisWorkbook.Worksheets("Pipeline")
    Dim filterArray()
    Dim currentFiltRange As String
    Dim lastArchiveRow As Long
    Dim mainDict As Scripting.Dictionary: Set mainDict = New Scripting.Dictionary
    
    ' Setup per arguments passed
    With ThisWorkbook.Worksheets("Summary")
        If indicator = "normal" Then
            .Range("M18:N19").Interior.Color = RGB(33, 89, 103)
            
        ElseIf indicator = "year" Then
            currentMonth = "Jun"
            pipelineYear = .Range("S15").Value
            .Range("M18:N19").Interior.Color = RGB(0, 133, 130)
            
        End If
    End With
    
    ' Setting up table & dynamic array
    With ThisWorkbook.Worksheets("Summary")
    
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).row
        numberOfLeads = lastRow - 19
    
        .Range("C20:AB100").ClearContents
        .Range("C20:AB100").Interior.ColorIndex = 0
        .Range("C20:AB" & lastRow).Value = 0
        .Range("C20:AB" & lastRow).WrapText = False
        .Range("C20:AB" & lastRow).RowHeight = 18
        .Range("C20:AB" & lastRow).ColumnWidth = 8.11
        .Range("C20:AB50").ClearComments
        
        If indicator = "normal" Then
            .Range("C20:L" & lastRow).Interior.Color = RGB(242, 242, 242)
        End If
        
        ReDim leadNames(numberOfLeads)

        For row = 20 To lastRow
            leadNames(row - 20) = Trim(CStr(.Range("B" & row).Value))
            .Range("AA" & row).Value = "=SUM(C" & row & ",E" & row & ",G" & row & ",I" & row & ",K" & row & ",M" & row & ",O" & row & ",Q" & row & ",S" & row & ",U" & row & ",W" & row & ",Y" & row & ")"
            .Range("AB" & row).Value = "=SUM(D" & row & ",F" & row & ",H" & row & ",J" & row & ",L" & row & ",N" & row & ",P" & row & ",R" & row & ",T" & row & ",V" & row & ",X" & row & ",Z" & row & ")"
        Next
        
    End With
    
    Application.ScreenUpdating = False
    
    ' Capture AutoFilter settings -- doesn't work with date or color filters
    On Error GoTo ErrNoFilter
    With w.AutoFilter
        currentFiltRange = .Range.Address
        With .Filters
            ReDim filterArray(1 To .Count, 1 To 3)
            For f = 1 To .Count
                With .Item(f)
                    If .On Then
                        filterArray(f, 1) = .Criteria1
                        If .Operator Then
                            filterArray(f, 2) = .Operator
                            'filterArray(f, 3) = .Criteria2 'simply delete this line to make it work in Excel 2010
                        End If
                    End If
                End With
            Next f
        End With
    End With

NoFilter:

    ' Filling out table on Summary tab
    On Error GoTo 0
    With ThisWorkbook.Worksheets("Pipeline")
    
        On Error Resume Next
        .AutoFilterMode = False 'Option = .ShowAllData ?
        lastPipelineRow = .Cells(.Rows.Count, "B").End(xlUp).row
        
        ' Searching for months
        On Error GoTo 0
        For col = 15 To 250
            If InStr(.Cells(1, col), currentMonth) > 0 And InStr(.Cells(1, col), pipelineYear) > 0 Then
                yearColStart = col - 5
                Exit For
            End If
            If col = 250 Then
                MsgBox "Unable to find " & currentMonth & "-" & pipelineYear & " column header on sheet 'Pipeline'"
                Application.ScreenUpdating = True
                Exit Sub
            End If
        Next
        
        ' Calling sub to format col headers
        FormatColumnHeaders pipelineYear, yearColStart
        
        ' Adding data to the table
        TableGenerator lastPipelineRow, numberOfLeads, leadNames, pipelineYear, yearColStart, ThisWorkbook.Worksheets("Pipeline"), mainDict
        
    End With
    
    ' Adding archive information to the table
    With ThisWorkbook.Worksheets("Archive")
    
        .AutoFilterMode = False 'Option = .ShowAllData ?
        lastArchiveRow = .Cells(.Rows.Count, "B").End(xlUp).row
        
        TableGenerator lastArchiveRow, numberOfLeads, leadNames, pipelineYear, yearColStart, ThisWorkbook.Worksheets("Archive"), mainDict
    
    End With
    
    ' Adding dictionary information to table
    AddDictionaryToTable lastRow, numberOfLeads, leadNames, mainDict
    
    ' Fixing row heights
    With ThisWorkbook.Worksheets("Summary")
        .Range("C20:AB" & lastRow).RowHeight = 18
    End With
    
    ' Restore Filter settings
    On Error GoTo 0
    If filterexists = True Then
        For col = 1 To UBound(filterArray(), 1)
            If Not IsEmpty(filterArray(col, 1)) Then
                If filterArray(col, 2) Then
                    w.Range(currentFiltRange).AutoFilter field:=col, _
                    Criteria1:=filterArray(col, 1), _
                    Operator:=filterArray(col, 2), _
                    Criteria2:=filterArray(col, 3)
                Else
                    w.Range(currentFiltRange).AutoFilter field:=col, _
                    Criteria1:=filterArray(col, 1)
                End If
            End If
        Next col
    End If
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrNoFilter:

    filterexists = False
    Resume NoFilter

End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    
    Dim lastRow As Integer
    Dim targetCol As Integer
    Dim colArray() As Variant
    Dim col As Integer

    colArray = Array(3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27)
    
    With ThisWorkbook.Worksheets("Summary")
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).row
    
        If Not Application.Intersect(Target, .Range("C20:AA" & lastRow)) Is Nothing Then
            If Not Application.Intersect(Target, Range("C20:C" & lastRow)) Is Nothing _
            Or Not Application.Intersect(Target, Range("E20:E" & lastRow)) Is Nothing _
            Or Not Application.Intersect(Target, Range("G20:G" & lastRow)) Is Nothing _
            Or Not Application.Intersect(Target, Range("I20:I" & lastRow)) Is Nothing _
            Or Not Application.Intersect(Target, Range("K20:K" & lastRow)) Is Nothing _
            Or Not Application.Intersect(Target, Range("M20:M" & lastRow)) Is Nothing _
            Or Not Application.Intersect(Target, Range("O20:O" & lastRow)) Is Nothing _
            Or Not Application.Intersect(Target, Range("Q20:Q" & lastRow)) Is Nothing _
            Or Not Application.Intersect(Target, Range("S20:S" & lastRow)) Is Nothing _
            Or Not Application.Intersect(Target, Range("U20:U" & lastRow)) Is Nothing _
            Or Not Application.Intersect(Target, Range("W20:W" & lastRow)) Is Nothing _
            Or Not Application.Intersect(Target, Range("Y20:Y" & lastRow)) Is Nothing _
            Or Not Application.Intersect(Target, Range("AA20:AA" & lastRow)) Is Nothing Then
            
            targetCol = Target.Column
                If Columns(targetCol).ColumnWidth < 9 Then
                    ' Expanding range to view
                    .Columns(targetCol).ColumnWidth = 30
                    ' Calling function to switch comments & values
                    flipComments targetCol, lastRow, "expand"
                End If
            End If
        Else
            For col = LBound(colArray) To UBound(colArray)
                If .Columns(colArray(col)).ColumnWidth > 9 Then
                   ' Switching back comments & values
                   flipComments CInt(colArray(col)), lastRow, "contract"
                End If
            Next
            ' Resetting size of cells
            .Range("C20:AB" & lastRow).RowHeight = 18
            .Range("C20:AB" & lastRow).ColumnWidth = 8.11
        End If
    End With
    
End Sub

Sub flipComments(targetCol As Integer, lastRow As Integer, indicator As String)

    Dim cell As Variant
    Dim tempString As String
    
    With ThisWorkbook.Worksheets("Summary")
        For Each cell In .Range(.Cells(20, targetCol), .Cells(lastRow, targetCol))
            If Not cell.Comment Is Nothing Then
                tempString = cell.Comment.Text
                cell.ClearComments
                cell.AddComment CStr(cell.Value)
                cell.Value = tempString
                With cell.Comment.Shape.TextFrame
                    .AutoSize = True
                End With
            End If
            If indicator = "expand" Then
                .Rows(cell.row).AutoFit
            End If
        Next
    End With

End Sub
