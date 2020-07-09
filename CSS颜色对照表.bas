Attribute VB_Name = "Ä£¿é1"
Sub csscolors()
    Dim Hex As String, r As Integer, g As Integer, b As Integer, Rowscout As Integer
    
    Rowscout = ActiveSheet.[A65536].End(xlUp).Row
    For i = 2 To Rowscout
        Hex = Worksheets("Sheet1").Cells(i, 2)
        r = H2D(Mid(Hex, 2, 2))
        g = H2D(Mid(Hex, 4, 2))
        b = H2D(Mid(Hex, 6, 2))
        Worksheets("Sheet1").Cells(i, 3).Interior.Color = RGB(r, g, b)
    Next i
    
End Sub

Public Function H2D(ByVal Hex As String) As Long
     Dim i As Long
     Dim b As Long
    
    Hex = UCase(Hex)
     For i = 1 To Len(Hex)
         Select Case Mid(Hex, Len(Hex) - i + 1, 1)
             Case "0": b = b + 16 ^ (i - 1) * 0
             Case "1": b = b + 16 ^ (i - 1) * 1
             Case "2": b = b + 16 ^ (i - 1) * 2
             Case "3": b = b + 16 ^ (i - 1) * 3
             Case "4": b = b + 16 ^ (i - 1) * 4
             Case "5": b = b + 16 ^ (i - 1) * 5
             Case "6": b = b + 16 ^ (i - 1) * 6
             Case "7": b = b + 16 ^ (i - 1) * 7
             Case "8": b = b + 16 ^ (i - 1) * 8
             Case "9": b = b + 16 ^ (i - 1) * 9
             Case "A": b = b + 16 ^ (i - 1) * 10
             Case "B": b = b + 16 ^ (i - 1) * 11
             Case "C": b = b + 16 ^ (i - 1) * 12
             Case "D": b = b + 16 ^ (i - 1) * 13
             Case "E": b = b + 16 ^ (i - 1) * 14
             Case "F": b = b + 16 ^ (i - 1) * 15
         End Select
     Next i
     H2D = b
End Function

