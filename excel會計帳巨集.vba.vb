Sub 巨集1()
'
' 巨集1 巨集
'
'
Dim I As Integer
I = 4

    
   Do While (Range("C" & I).Value <> "")
   
   
   
   If (Range("C" & I).Value = "存出保證金") Then
    Range("P" & I) = choose_ex(I)
    End If
    
   If (Range("C" & I).Value = "暫付款") Then
    Range("R" & I) = choose_ex(I)
     Range("T" & I) = choose_ex(I)
  End If
   If (Range("C" & I).Value = "運輸設備") Then
    Range("V" & I) = choose_ex(I)
    End If
    
     If (Range("C" & I).Value = "稅金") Then
    Range("X" & I) = choose_ex(I)
    End If
     If (Range("C" & I).Value = "薪資支出") Then
    Range("Z" & I) = choose_ex(I)
    End If
    If (Range("C" & I).Value = "雜費") Then
    Range("AB" & I) = choose_ex(I)
    End If
      If (Range("C" & I).Value = "郵電費") Then
    Range("AF" & I) = choose_ex(I)
    End If
      If (Range("C" & I).Value = "文具用品") Then
    Range("AH" & I) = choose_ex(I)
    End If
      If (Range("C" & I).Value = "修繕費") Then
    Range("AJ" & I) = choose_ex(I)
    End If
      If (Range("C" & I).Value = "燃料費") Then
    Range("AD" & I) = choose_ex(I)
    End If
     If (Range("C" & I).Value = "租金") Then
    Range("AL" & I) = choose_ex(I)
    End If
     If (Range("C" & I).Value = "雜項購置") Then
    Range("AN" & I) = choose_ex(I)
    End If
    If (Range("C" & I).Value = "書報雜誌") Then
    Range("AP" & I) = choose_ex(I)
    End If
    If (Range("C" & I).Value = "保險費") Then
    Range("AR" & I) = choose_ex(I)
    End If
    If (Range("C" & I).Value = "交際費") Then
    Range("AT" & I) = choose_ex(I)
    End If
    If (Range("C" & I).Value = "水電費") Then
    Range("AV" & I) = choose_ex(I)
    End If
    If (Range("C" & I).Value = "快遞費") Then
    Range("AX" & I) = choose_ex(I)
    End If
    If (Range("C" & I).Value = "訓練費") Then
    Range("AZ" & I) = choose_ex(I)
    End If
    If (Range("C" & I).Value = "交通費") Then
    Range("BB" & I) = choose_ex(I)
    End If
    
    
    If (Range("C" & I).Value = "廣告費") Then
    Range("BD" & I) = choose_ex(I)
    End If
    
    If (Range("C" & I).Value = "運費") Then
    Range("BF" & I) = choose_ex(I)
    End If
    
    If (Range("C" & I).Value = "旅費") Then
    Range("BH" & I) = choose_ex(I)
    End If
    
     I = I + 1
   Loop
End Sub
   
   Function choose_ex(ByVal j As Integer) As Long
      Dim ex As Long
   
      If Range("F" & j) <> 0 Then
         ex = CLng(Range("F" & j))
     
      End If
      
      If Range("I" & j) <> 0 Then
         ex = CLng(Range("I" & j))
      Else
         ex = CLng(Range("L" & j))
      End If
    
    Application.Goto Reference:="巨集1"
         choose_ex = ex
   
   End Function
   
   
   
   
