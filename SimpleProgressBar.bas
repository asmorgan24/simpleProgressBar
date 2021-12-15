Attribute VB_Name = "Module1"
'' Add progress bar and page numbers to all non-hidden pages
Sub AddProgressBar()
    On Error Resume Next
        With ActivePresentation
              sHeight = .PageSetup.SlideHeight - 12
              n = 0
              j = 0
              vert_loc = 0 ' .PageSetup.SlideHeight - 12
              For i = 1 To .Slides.Count
                If .Slides(i).SlideShowTransition.Hidden Then j = j + 1
              Next i:
              For i = 2 To .Slides.Count
                .Slides(i).Shapes("progressBar").Delete
                .Slides(i).Shapes("pageNumber").Delete
                If .Slides(i).SlideShowTransition.Hidden = msoFalse Then
                  Set slider = .Slides(i).Shapes.AddShape(msoShapeRectangle, 0, vert_loc, (i - n) * .PageSetup.SlideWidth / (.Slides.Count - j), 12)
                  With slider
                      .Fill.ForeColor.RGB = RGB(15, 77, 146)
                      .Name = "progressBar"
                  End With
                  Set pageNumber = .Slides(i).Shapes.AddTextbox(msoTextOrientationHorizontal, ((i - n) * .PageSetup.SlideWidth / (.Slides.Count - j)) - 40, vert_loc - 3, 100, 10)
                  With pageNumber
                      .TextFrame.TextRange.Text = Str(i - n) & "/" & Str(ActivePresentation.Slides.Count - j)
                       With .TextFrame.TextRange.Font
                           .Bold = msoTrue
                           .Size = 10
                           .Color = RGB(255, 255, 255)
                       End With
                       .Name = "pageNumber"
                   End With
                  Else
                    n = n + 1
                  End If
              Next i:
        End With
End Sub

'' Macro to remove the progress bar from all the slides
Sub RemoveProgressBar()
    On Error Resume Next
        With ActivePresentation
              For i = 1 To .Slides.Count
              .Slides(i).Shapes("progressBar").Delete
              .Slides(i).Shapes("pageNumber").Delete
              Next i:
        End With
End Sub
