Attribute VB_Name = "Module1"
Option Explicit 

Sub ClearTable()

    Dim lastRow As Integer

    With ThisWorkbook.Worksheets("Summary")
        
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).row
    
        .Range("C20:AB100").ClearContents
        .Range("C20:AB100").Interior.ColorIndex = 0
        .Range("C20:AB" & lastRow).Value = ""
    End With

End Sub


Sub PopulateArchive()

    Dim lastRow As Integer
    Dim lastArchiveRow As Integer
    Dim row As Integer
    Dim col As Integer
    Dim lastCol As Integer
    Dim archiveSpot As Integer: archiveSpot = 0
    Dim targetRange As Range
    Dim letterArray() As Variant: letterArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N")
    Dim letter As Variant
    
    With ThisWorkbook.Worksheets("Pipeline")
        
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).row
        
        With ThisWorkbook.Worksheets("Archive")
            For Each letter In letterArray
                If archiveSpot < .Cells(.Rows.Count, letter).End(xlUp).row + 1 Then
                    archiveSpot = .Cells(.Rows.Count, letter).End(xlUp).row + 1
                End If
            Next
        End With
        
        ' Finding last used column
        For col = 1 To 1000
            If Len(.Cells(1, col).Value) = 0 Then
                lastCol = col
                Exit For
            End If
        Next
        
        ' Moving sign-off rows to archive .Range(.Cells(1, 1), .Cells(10, 1)) = 5
        For row = 2 To lastRow
            If InStr(CStr(.Range("J" & row).Value), "Sign-Off provided") > 0 _
                Or InStr(CStr(.Range("J" & row).Value), "Production") > 0 _
                Or InStr(CStr(.Range("J" & row).Value), "Prod Deployed") > 0 Then
            
                ' First transferring to archive
                With ThisWorkbook.Worksheets("Archive")
                    Set targetRange = .Range(.Cells(archiveSpot, 1), .Cells(archiveSpot, lastCol))
                End With
                targetRange.Value = .Range(.Cells(row, 1), .Cells(row, lastCol)).Value
                
                archiveSpot = archiveSpot + 1
                
                ' Deleting row from Pipeline
                .Rows(row).EntireRow.Delete
                row = row - 1
            
            End If
        Next
        
    End With
    
    ' Adding borders
    With ThisWorkbook.Worksheets("Archive")
        
        lastRow = 0
        
        For Each letter In letterArray
            If lastRow < .Cells(.Rows.Count, letter).End(xlUp).row Then
                lastRow = .Cells(.Rows.Count, letter).End(xlUp).row
            End If
        Next
        
        .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).Borders.LineStyle = xlContinuous
        
    End With

End Sub

Sub PutArchiveBack()

    Dim lastArchiveSpot As Integer: lastArchiveSpot = 0
    Dim lastPipelineSpot As Integer: lastPipelineSpot = 0
    Dim letterArray() As Variant: letterArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N")
    Dim col As Integer
    Dim lastCol As Integer
    Dim letter As Variant
    
    With ThisWorkbook.Worksheets("Archive")
        
        ' Getting last archive row
        For Each letter In letterArray
            If lastArchiveSpot < .Cells(.Rows.Count, letter).End(xlUp).row Then
                lastArchiveSpot = .Cells(.Rows.Count, letter).End(xlUp).row
            End If
        Next
        
        ' Getting last archive col
        lastCol = .Cells(2, .Columns.Count).End(xlToLeft).Column
    End With
        
    ' Getting last pipeline row
    With ThisWorkbook.Worksheets("Pipeline")
        
        For Each letter In letterArray
            If lastPipelineSpot < .Cells(.Rows.Count, letter).End(xlUp).row + 1 Then
                lastPipelineSpot = .Cells(.Rows.Count, letter).End(xlUp).row + 1
            End If
        Next
        
        ' Getting last pipeline col
        If lastCol < .Cells(1, .Columns.Count).End(xlToLeft).Column Then
            lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        End If
    End With
        
    With ThisWorkbook.Worksheets("Archive")
        
        ' Transferring to pipeline
        If lastArchiveSpot > 2 Then
            .Range(.Cells(3, 1), .Cells(lastArchiveSpot, lastCol)).Copy Destination:=ThisWorkbook.Worksheets("Pipeline").Range("A" & lastPipelineSpot)
            .Range(.Cells(3, 1), .Cells(lastArchiveSpot, lastCol)).ClearContents
            .Range(.Cells(3, 1), .Cells(lastArchiveSpot, lastCol)).Borders.LineStyle = xlNone
        End If
    End With

End Sub

Public Sub FormatColumnHeaders(pipelineYear As String, yearColStart As Integer)

    Dim i As Integer
    Dim col As Integer
    Dim stub As String

    ' Setting column headers on summary tab
    With ThisWorkbook.Worksheets("Pipeline")
        i = 0
        For col = 3 To 25 Step 2
            ' getting left 3 letters of pipelne col for summary col
            ThisWorkbook.Worksheets("Summary").Cells(18, col).Value = .Cells(1, yearColStart + i).Value ' Left(CStr(.Cells(1, yearColStart + i).Value), 3)
            If Len(CStr(ThisWorkbook.Worksheets("Summary").Cells(18, col).Value)) = 0 Then
                Select Case Left(CStr(ThisWorkbook.Worksheets("Summary").Cells(18, col - 2).Value), 3)
                    Case "Jan"
                        ThisWorkbook.Worksheets("Summary").Cells(18, col).Value = "Feb-" & (CInt(pipelineYear) + 1)
                    Case "Feb"
                        ThisWorkbook.Worksheets("Summary").Cells(18, col).Value = "Mar-" & (CInt(pipelineYear) + 1)
                    Case "Mar"
                        ThisWorkbook.Worksheets("Summary").Cells(18, col).Value = "Apr-" & (CInt(pipelineYear) + 1)
                    Case "Apr"
                        ThisWorkbook.Worksheets("Summary").Cells(18, col).Value = "May-" & (CInt(pipelineYear) + 1)
                    Case "May"
                        ThisWorkbook.Worksheets("Summary").Cells(18, col).Value = "Jun-" & (CInt(pipelineYear) + 1)
                    Case "Jun"
                        ThisWorkbook.Worksheets("Summary").Cells(18, col).Value = "Jul-" & (CInt(pipelineYear) + 1)
                    Case "Jul"
                        ThisWorkbook.Worksheets("Summary").Cells(18, col).Value = "Aug-" & (CInt(pipelineYear) + 1)
                    Case "Aug"
                        ThisWorkbook.Worksheets("Summary").Cells(18, col).Value = "Sep-" & (CInt(pipelineYear) + 1)
                    Case "Sep"
                        ThisWorkbook.Worksheets("Summary").Cells(18, col).Value = "Oct-" & (CInt(pipelineYear) + 1)
                    Case "Oct"
                        ThisWorkbook.Worksheets("Summary").Cells(18, col).Value = "Nov-" & (CInt(pipelineYear) + 1)
                    Case "Nov"
                        ThisWorkbook.Worksheets("Summary").Cells(18, col).Value = "Dec-" & (CInt(pipelineYear) + 1)
                    Case "Dec"
                        ThisWorkbook.Worksheets("Summary").Cells(18, col).Value = "Jan-" & (CInt(pipelineYear) + 1)
                    Case Else
                        ThisWorkbook.Worksheets("Summary").Cells(18, col).Value = "Jan-" & (CInt(pipelineYear))
                End Select
            End If
            i = i + 1
        Next
    End With

End Sub

Public Sub TableGenerator(lastRow As Long, numberOfLeads As Integer, leadNames() As String, pipelineYear As String, yearColStart As Integer, targetWorksheet As Worksheet, ByRef mainDict As Scripting.Dictionary)

    Dim searchString As String
    Dim searchYear As String
    Dim targetCol As Integer
    Dim col As Integer
    Dim tableColSpot As Integer: tableColSpot = 4
    Dim row As Long
    Dim trimStatus As String
    Dim monthDict As Scripting.Dictionary: Set monthDict = New Scripting.Dictionary
    Dim statusDict As Scripting.Dictionary: Set statusDict = New Scripting.Dictionary
    Dim i As Long
    Dim keyString As String
    Dim monthExisted As Boolean

    With targetWorksheet
    
        For row = 2 To lastRow
            For i = 0 To numberOfLeads - 1
            
                If mainDict.Exists(leadNames(i)) Then
                    Set monthDict = mainDict(leadNames(i))
                    monthExisted = True
                Else
                    Set monthDict = New Scripting.Dictionary
                    monthExisted = False
                End If
            
                ' Optional use lowercase - If InStr(LCase(CStr(.Range("F" & row).Value)), LCase(leadNames(i))) > 0 Then
                If InStr(CStr(.Range("F" & row).Value), leadNames(i)) > 0 Then
                
                    ' Summing WR count by month
                    If Right(Trim(.Range("E" & row).Value), 4) = pipelineYear _
                        Or Right(Trim(.Range("E" & row).Value), 4) = CStr(CInt(pipelineYear) + 1) _
                        Or Right(Trim(.Range("E" & row).Value), 4) = CStr(CInt(pipelineYear) - 1) Then
                        
                        Select Case Left(Trim(.Range("E" & row).Value), InStr(Trim(.Range("E" & row).Value), "/") - 1)
                            Case "1"
                                searchString = "Jan"
                            Case "2"
                                searchString = "Feb"
                            Case "3"
                                searchString = "Mar"
                            Case "4"
                                searchString = "Apr"
                            Case "5"
                                searchString = "May"
                            Case "6"
                                searchString = "Jun"
                            Case "7"
                                searchString = "Jul"
                            Case "8"
                                searchString = "Aug"
                            Case "9"
                                searchString = "Sep"
                            Case "10"
                                searchString = "Oct"
                            Case "11"
                                searchString = "Nov"
                            Case "12"
                                searchString = "Dec"
                            Case Else
                                MsgBox "Error with release date: " & Trim(.Range("E" & row).Value) & ", cannot continue updating table."
                        End Select
                    
                        searchYear = Right(Trim(.Range("E" & row).Value), 4)
                        With ThisWorkbook.Worksheets("Summary")
                            For col = 3 To 27 Step 2
                                If InStr(.Cells(18, col), searchString) > 0 And InStr(.Cells(18, col), searchYear) > 0 Then
                                    targetCol = col
                                    Exit For
                                End If
                                If col = 27 Then
                                    targetCol = 0
                                End If
                            Next
                            
                            ' Updating Table
                            If targetCol <> 0 Then
                                .Cells(i + 20, targetCol).Value = CInt(.Cells(i + 20, targetCol).Value) + 1
                                
                                With targetWorksheet
                                    ' Counting statuses - dictionary
                                    trimStatus = Trim(CStr(.Range("J" & row).Value))
                                    keyString = searchString & CStr(pipelineYear)
                                    
                                    If monthDict.Exists(keyString) Then
                                        Set statusDict = monthDict(keyString)
                                        If statusDict.Exists(trimStatus) Then
                                            statusDict(trimStatus) = CInt(statusDict(trimStatus) + 1)
                                        Else
                                            statusDict.Add trimStatus, CInt(1)
                                        End If
                                        Set monthDict(keyString) = statusDict
                                    Else
                                        Set statusDict = New Scripting.Dictionary
                                        statusDict.Add trimStatus, CInt(1)
                                        monthDict.Add keyString, statusDict
                                    End If
                                    
                                    If monthExisted = False Then
                                        mainDict.Add leadNames(i), monthDict
                                    Else
                                        Set mainDict(leadNames(i)) = monthDict
                                    End If
                                End With
                            End If
                        End With
                    End If
                    
                    ' Summing month hours
                    For col = yearColStart To yearColStart + 12
                        
                        If IsNumeric(ThisWorkbook.Worksheets("Summary").Cells(i + 20, tableColSpot).Value) And IsNumeric(Trim(.Cells(row, col))) Then
                            ThisWorkbook.Worksheets("Summary").Cells(i + 20, tableColSpot).Value = CInt(ThisWorkbook.Worksheets("Summary").Cells(i + 20, tableColSpot).Value) + CInt(Trim(.Cells(row, col)))
                        End If
                        
                        tableColSpot = tableColSpot + 2
                    Next
                    tableColSpot = 4
                    
                    Exit For
                End If
            Next
        Next
    End With

End Sub

Sub AddDictionaryToTable(lastRow As Integer, numberOfLeads As Integer, leadNames() As String, mainDict As Scripting.Dictionary)

    Dim row As Integer
    Dim col As Integer
    Dim i As Integer
    Dim key As Variant
    Dim leftMonthKey As String
    Dim rightMonthKey As String
    Dim holderString As String
    Dim monthDict As Scripting.Dictionary: Set monthDict = New Scripting.Dictionary
    Dim statusDict As Scripting.Dictionary: Set statusDict = New Scripting.Dictionary
    Dim commentString As String
    Dim cmt As Comment
    Dim keyHolder As String

    With ThisWorkbook.Worksheets("Summary")
        For row = 20 To 20 + numberOfLeads
            i = row - 20
            On Error GoTo LeadHandler
            Set monthDict = mainDict(leadNames(i))
            For col = 3 To 25 Step 2
                leftMonthKey = Left(.Cells(18, col).Value, 3)
                rightMonthKey = Right(.Cells(18, col).Value, 4)
                
                On Error GoTo Handler
                Set statusDict = monthDict(leftMonthKey & rightMonthKey)
                
                On Error GoTo 0
                .Range(.Cells(20, col), .Cells(lastRow, col)).HorizontalAlignment = xlRight
                .Range(.Cells(20, col), .Cells(lastRow, col)).VerticalAlignment = xlTop
                For Each key In statusDict.Keys
                    If Len(Trim(CStr(key))) > 0 Then
                        keyHolder = key
                    Else
                        keyHolder = "(blank)"
                    End If
                    If Len(commentString) > 0 Then
                        commentString = commentString & vbCrLf & statusDict(key) & " - " & keyHolder
                    Else
                        commentString = statusDict(key) & " - " & keyHolder
                    End If
                Next
                
                .Cells(row, col).AddComment commentString
                
                Set cmt = .Cells(row, col).Comment
                With cmt.Shape
                  .TextFrame.Characters.Font.Size = 11
                  .TextFrame.AutoSize = True
                End With
                commentString = ""
NextMonth:
            Next
NextLead:
        Next
    End With
    
    Exit Sub
    
LeadHandler:
    Resume NextLead
    
Handler:
    Resume NextMonth

End Sub

