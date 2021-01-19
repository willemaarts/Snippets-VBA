Sub ArrayTest()
'https://excelmacromastery.com/excel-vba-array/

    'Dim arrMarks(0 To 3) As Variant
    
    Dim arrMarks() As Variant
    ReDim arrMarks(0 To 3)
    
    Dim i As Long
    
'    For i = 0 To 3
'        arrMarks(i) = Range("A1").Offset(i).Value
'    Next i
    
    For i = LBound(arrMarks) To UBound(arrMarks)
        Debug.Print arrMarks(i)
    Next i
    
    For i = 0 To 3
        Debug.Print arrMarks(i)
    Next i
    
'    Debug.Print arrMarks(0)
'
'    Debug.Print arrMarks(1)
'
'    Debug.Print arrMarks(2)
'
'    Debug.Print arrMarks(3)

End Sub

