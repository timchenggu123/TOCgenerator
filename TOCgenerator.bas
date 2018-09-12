Attribute VB_Name = "NewMacros"
Sub test()
'
' test Macro
'
'
MsgBox (Selection.Start)
End Sub
Sub getTOC()
Dim strFirstWord As String
Dim strIn As String
Dim oPara As Word.Paragraph
Dim tableofcontent As Range
Dim textline As String

Set tableofcontent = Selection.Range
a = 1

For Each oPara In ActiveDocument.Paragraphs
    On Error Resume Next
    strFirstWord = oPara.Range.Characters(1).text
    strSecondWord = oPara.Range.Characters(2).text
    strThirdWord = oPara.Range.Characters(3).text
    strFourthWord = oPara.Range.Characters(4).text
    strWords = strFirstWord & strSecondWord & strThirdWord
    'MsgBox (oPara.Range.Text & " " & strWords)

        If Len(strWords) < 4 And Len(strWords) > 2 And IsNumeric(Left(strWords, 1)) And IsNumeric(Mid(strWords, 3, 1)) And Val(strFourthWord) = 0 Then
            textline = oPara.Range.text
            textline = Left(textline, Len(textline) - 1)
            oPara.Range.Select
            pg = Selection.Information(wdActiveEndAdjustedPageNumber)
            
            tabs = TOClevel(textline)
            
            Do While tabs > 0
                textline = "    " & textline
                tabs = tabs - 1
            Loop
            
            b = 175 - Len(textline) * 2
            dots = ""
            flag = True
            Do
            tableofcontent.Select
            Selection.MoveDown Unit:=wdLine, count:=a
            Selection.Expand wdLine
            
            For i = 1 To b
                dots = dots & "."
            Next i
            
            Selection.text = textline & dots & pg & Chr(10) & Chr(10)
            tableofcontent.Select
            Selection.MoveDown Unit:=wdLine, count:=a + 1
            Selection.Expand wdLine
            
            b = b - 1
            If Len(Selection.text) > 1 Then
                If Len(Selection.text) > 155 Then
                    b = b - 3
                End If
                
                Selection.Expand wdParagraph
                Selection.text = ""
                dots = ""
            Else
                Exit Do
            End If
            Loop
            a = a + 1
        End If
Next oPara

End Sub
Sub ConvertAutoNumbers()
'
' ConvertAutoNumbers Macro

    If ActiveDocument.Lists.count > 0 Then
        Dim lisAutoNumList As List

        For Each lisAutoNumList In ActiveDocument.Lists
            lisAutoNumList.ConvertNumbersToText
        Next
    Else
              
    End If

End Sub
Sub moveline()
'
' moveline Macro
'
'
Selection.MoveDown Unit:=wdLine, count:=1
Selection.Expand wdLine
MsgBox (Selection.text)
End Sub
Sub count()
'
' count Macro

MsgBox (Len(Selection.text))
End Sub

Function TOClevel(text As String) As Integer

sp = InStr(text, Chr(32))
ID = Left(text, sp - 1)
c = 0

If IsNumeric(Right(ID, 1)) And Val(Right(ID, 1)) = 0 Then
    c = 1
End If
    
TOClevel = (Len(ID) - 1) / 2 - c
    
End Function

Sub figuresTOC()
Dim strFirstWord As String
Dim strIn As String
Dim oPara As Word.Paragraph
Dim tableofcontent As Range
Dim textline As String

Set tableofcontent = Selection.Range
a = 1

For Each shp In ActiveDocument.Shapes
     shp.Select
     Selection.Expand wdParagraph
     pg = Selection.Information(wdActiveEndAdjustedPageNumber)
     
     textline = Selection.Range.text
     fig = InStr(textline, "Figure")
     textline = Mid(textline, fig)
     textline = Replace(textline, Chr(13), "")
     textline = Replace(textline, Chr(10), "")
     
     b = 175 - Len(textline) * 2
     dots = ""
     flag = True
     forcestop = 1
     Do
         tableofcontent.Select
         Selection.MoveDown Unit:=wdLine, count:=a
         Selection.Expand wdLine
        
         For i = 1 To b
             dots = dots & "."
         Next i
        
        Selection.text = textline & dots & pg & Chr(10) & Chr(10)
        tableofcontent.Select
        Selection.MoveDown Unit:=wdLine, count:=a + 1
        Selection.Expand wdLine
        
        b = b - 1
        If Len(Selection.text) > 1 Then
            If Len(Selection.text) > 155 Then
                b = b - 3
            End If
            
            Selection.Expand wdParagraph
            Selection.text = ""
            dots = ""
        Else
            Exit Do
        End If
        forcestop = forcestop + 1
        If forcestop > 500 Then
            MsgBox ("text too long")
            Exit Do
        End If
        
    Loop
    a = a + 1

Next shp
    
 
End Sub
Sub testsel()
'
' testsel Macro
'
'
MsgBox TypeName(Selection)
End Sub
