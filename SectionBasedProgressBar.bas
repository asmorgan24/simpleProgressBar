Attribute VB_Name = "Module1"
'' Add progress bar and page numbers to all non-hidden pages
Sub AddProgressBar()
    On Error Resume Next
        With ActivePresentation
              n = 0
              j = 0 'number of active slides
              bar_height = 40 'height of the progress bar
              vert_loc = 0 ' 0 for top, .PageSetup.SlideHeight - bar_height for bottom
              num_sects = .SectionProperties.Count 'total number of sections in the ppt
              one = .Slides(12).sectionIndex
              nam = .SectionProperties.Name(one)
              sec_rec_size = Int(.PageSetup.SlideWidth / num_sects) 'determines size of textFrame
              
              rect_color = RGB(15, 77, 146)
              
              For i = 1 To .Slides.Count
                If .Slides(i).SlideShowTransition.Hidden Then j = j + 1
              Next i:
              
              RemoveProgressBar
              
              For i = 2 To .Slides.Count
                If .Slides(i).SlideShowTransition.Hidden = msoFalse Then
                
                  s_num = .Slides(i).sectionIndex 'section_number
                  num_slides_in_section = .SectionProperties.SlidesCount(s_num)
                  first_slide_num = .SectionProperties.FirstSlide(s_num)
                  
                  Set full_bar = .Slides(i).Shapes.AddShape(msoShapeRectangle, 0, vert_loc, .PageSetup.SlideWidth, bar_height)
                  With full_bar
                      .Fill.ForeColor.RGB = rect_color
                      .Name = "progressBar"
                      .Fill.Transparency = 0.2
                  End With
                
                  Set slider = .Slides(i).Shapes.AddShape(msoShapeRectangle, 0, vert_loc, (s_num - 1) * sec_rec_size + Int(((i - first_slide_num + 1) / num_slides_in_section) * sec_rec_size), bar_height)
                  With slider
                      .Fill.ForeColor.RGB = rect_color
                      .Name = "progressBar" 'name them the same so they delete easily
                      '.Fill.Transparency = 0.5
                  End With
                  
                  For k = 1 To num_sects:
                         Set section_box = .Slides(i).Shapes.AddTextbox(msoTextOrientationHorizontal, (k - 1) * sec_rec_size, vert_loc - 3, sec_rec_size, bar_height - 2)
                         With section_box
                           .TextFrame.TextRange.Text = ActivePresentation.SectionProperties.Name(k)  'nam 'Str(.SectionProperties.Name(j))
                           .TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
                            With .TextFrame.TextRange.Font
                                If ActivePresentation.Slides(i).sectionIndex = k Then .Bold = msoTrue
                                .Size = bar_height - Int(bar_height / 4)
                                .Color = RGB(255, 255, 255)
                            End With
                            .Name = "sectionBox"
                          End With
                          
                        If ActivePresentation.Slides(i).sectionIndex = k Then
                            Set HeaderBOX = .Slides(i).Shapes.AddShape(msoShapeRectangle, (k - 1) * sec_rec_size, vert_loc, sec_rec_size, bar_height)
                            'or Set HeaderBOX = ppSlide2.AddTextbox(msoTextOrientationHorizontal, 75, 150, 800, 700)
                            With HeaderBOX
                                .Name = "sectionBox"
                                .Fill.Visible = False   'make it transparent
                                .Line.Visible = True
                                .Line.ForeColor.RGB = rgbBlack 'RGB(0,0,0)
                                .Line.Weight = 2
                            End With
                        End If
                                                                                  
                          
                  Next k:
                                    

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
                    For j = 1 To 20
                        .Slides(i).Shapes("progressBar").Delete
                        .Slides(i).Shapes("pageNumber").Delete
                        .Slides(i).Shapes("sectionBox").Delete
                    Next j:
              Next i:
        End With
End Sub
