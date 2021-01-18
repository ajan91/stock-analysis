Sub GradeBook()

  If Cells(2, 2).Value >= 90 Then

      Cells(2, 3).Value = "Pass"

      Cells(2, 3).Interior.Color = vbGreen

      Cells(2, 4).Value = "A"

  ElseIf Cells(2, 2).Value >= 80 Then

      Cells(2, 3).Value = "Pass"

      Cells(2, 3).Interior.Color = vbGreen

      Cells(2, 4).Value = "B"

  ElseIf Cells(2, 2).Value >= 70 Then

      Cells(2, 3).Value = "Warning"

      Cells(2, 3).Interior.Color = vbYellow

      Cells(2, 4).Value = "C"


  Else

    Cells(2, 3).Value = "Fail"

      Cells(2, 3).Interior.Color = vbRed

      Cells(2, 4).Value = "F"

  End If

End Sub
