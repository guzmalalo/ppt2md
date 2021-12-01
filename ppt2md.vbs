' this code extracts objets from a ppt(x) file and saves to markdown
' Author: Guzman
' Adapated from : https://jasonkerwin.com/Files/ConvertToBeamer.txt
' version: 0.1.0

Public Sub ppt2md()
    ' Current presentation
    Dim objPresentation As Presentation
    Set objPresentation = Application.ActivePresentation

    ' Temporal PowerPoint objects
    Dim objSlide As Slide
    Dim objshape As Shape
    Dim objShape4Note As Shape
    Dim objFileSystem
    Dim objTextFile
    Dim objGrpItem As Shape

    ' File variables
    Dim Name As String      ' File Name with extension
    Dim Pth As String       ' Full Path
    Dim BaseName As String  ' Base name
    Dim Dest As String      ' Destination Path
    Dim OutName As String   ' Output filename
    Dim MedPath As String   ' Media Path
    Dim hght As Long        ' Slide height
    Dim wdth As Long        ' Slide widht

    ' Other temporal variables
    Dim IName As String
    Dim ln As String
    Dim ttl As String
    Dim txt As String
    Dim p As Integer, l As Integer, ctr As Integer, i As Integer, j As Integer
    Dim il As Long, cl As Long
    Dim Pgh As TextRange

    ' Get the path
    Pth = Application.ActivePresentation.Path

    ' Get the name
    Name = Application.ActivePresentation.Name

    ' Get the base name
    BaseName = Left(Name, InStrRev(Name, ".") - 1)

    ' Set the output path and name
    Dest = Pth & "\build\"
    If Dir(Dest, vbDirectory) = vbNullString Then
        MkDir (Dest)
    End If
    OutName = Dest & BaseName & ".md"

    ' Set the media path
    MedPath = Dest & "media\"
    If Dir(MedPath, vbDirectory) = vbNullString Then
        MkDir (MedPath)
    End If

    ' Create the output file
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    Set objTextFile = objFileSystem.CreateTextFile(OutName, True, True)

    ' Write metadata
    objTextFile.WriteLine "---"
    objTextFile.WriteLine "marp: true"
    objTextFile.WriteLine "paginate: true"
    objTextFile.WriteLine "theme: gaia"
    objTextFile.WriteLine "backgroundColor: #fff"
    objTextFile.WriteLine "title : " & Name
    With Application.ActivePresentation.PageSetup
        wdth = .SlideWidth
        hght = .SlideHeight
    End With
    objTextFile.WriteLine "wdth : " & wdth
    objTextFile.WriteLine "hght : " & hght
    objTextFile.WriteLine "---"

    ' Initialize counter for objets names
    ctr = 0

    ' Loop over Slides
    For Each objSlide In objPresentation.Slides

        ' TODO -> Check for layout

        ' Write the slide title
        objTextFile.WriteLine ""
        ttl = "# No Title"
        If objSlide.Shapes.HasTitle Then
            ttl = objSlide.Shapes.Title.TextFrame.TextRange.Text
        End If
        objTextFile.WriteLine "## " & ttl
        objTextFile.WriteLine "<!--- Slide Nr:" & objSlide.SlideIndex & "--->"

        ' Loop over the objets shapes on the slides
        For Each objshape In objSlide.Shapes

            ' Text box
            If objshape.HasTextFrame = True Then
                If Not objshape.TextFrame.TextRange Is Nothing Then
                    il = 0
                    For Each Pgh In objshape.TextFrame.TextRange.Paragraphs

                        If Not objshape.TextFrame.TextRange.Text = ttl Then
                            cl = Pgh.Paragraphs.IndentLevel
                            txt = Pgh.TrimText
                            If cl > il Then
                                il = cl
                                ElseIf cl < il Then
                                il = cl
                            End If
                            If txt <> "" Then
                                If il = 0 Then
                                    objTextFile.WriteLine txt
                                    Else
                                    For i = 1 To il - 1
                                        objTextFile.Write "  "
                                    Next i
                                    If Pgh.ParagraphFormat.Bullet.Visible = msoTrue Then
                                        If Pgh.ParagraphFormat.Bullet.Type = ppBulletUnnumbered Then
                                            objTextFile.WriteLine "- " + txt
                                            ElseIf Pgh.ParagraphFormat.Bullet.Type = ppBulletNumbered Then
                                            objTextFile.WriteLine "1. " + txt
                                        End If
                                        Else
                                        objTextFile.WriteLine txt
                                    End If
                                End If
                            End If
                        End If
                    Next Pgh
                    objTextFile.WriteLine
                End If

                ' TODO -> Adapt to markdown
                ' Tables
                ElseIf objshape.HasTable Then
                ln = "\begin{tabular}{|"
                For j = 1 To objshape.Table.Columns.Count
                    ln = ln & "l|"
                Next j
                ln = ln & "} \hline"
                objTextFile.WriteLine ln
                With objshape.Table
                    For i = 1 To .Rows.Count
                        If .Cell(i, 1).Shape.HasTextFrame Then
                            ln = .Cell(i, 1).Shape.TextFrame.TextRange.Text
                        End If

                        For j = 2 To .Columns.Count
                            If .Cell(i, j).Shape.HasTextFrame Then
                                ln = ln & " & " & .Cell(i, j).Shape.TextFrame.TextRange.Text
                            End If
                        Next j
                        ln = ln & "  \\ \hline"
                        objTextFile.WriteLine ln
                    Next i
                    objTextFile.WriteLine "\end{tabular}" & vbCrLf
                End With

                ' TODO -> Export to latex
                ' Equations
                ElseIf objshape.Type = msoEmbeddedOLEObject Then
                If objshape.OLEFormat.ProgID = "Equation.3" Then
                    IName = BaseName + "-img" & Format(ctr, "0000") & ".png"
                    objTextFile.WriteLine "![" & IName & "](media\" & IName & ")"
                    Call objshape.Export(MedPath & IName, ppShapeFormatPNG, , , ppRelativeToSlide)
                    ctr = ctr + 1
                    ElseIf objshape.OLEFormat.ProgID = "Equation.DSMT4" Then
                    IName = BaseName + "-img" & Format(ctr, "0000") & ".png"
                    objTextFile.WriteLine "![" & IName & "](media\" & IName & ")"
                    Call objshape.Export(MedPath & IName, ppShapeFormatPNG, , , ppRelativeToSlide)
                    ctr = ctr + 1
                End If

                ' Pictures
                ElseIf (objshape.Type = msoPicture) Then
                IName = BaseName + "-img" & Format(ctr, "0000") & ".png"
                objTextFile.WriteLine "![" & IName & "](media\" & IName & ")"
                Call objshape.Export(MedPath & IName, ppShapeFormatPNG, , , ppRelativeToSlide)
                ctr = ctr + 1

                ' TODO -> Export as svg or pdf
                ' Groups
                ElseIf (objshape.Type = msoGroup) Then
                IName = BaseName + "-img" & Format(ctr, "0000") & ".png"
                objTextFile.WriteLine "![" & IName & "](media\" & IName & ")"
                Call objshape.Export(MedPath & IName, ppShapeFormatPNG, , , ppRelativeToSlide)
                ctr = ctr + 1
                For Each objGrpItem In objshape.GroupItems
                    If objGrpItem.HasTextFrame = True Then
                        If Not objGrpItem.TextFrame.TextRange Is Nothing Then
                            shpx = objGrpItem.Top / hght
                            shpy = objGrpItem.Left / wdth
                            ' this could need adjustment (Footers textblocks)
                            If objGrpItem.TextFrame.TextRange.Text <> "" Then
                                If shpx < 0.1 And shpy > 0.5 Then
                                    objTextFile.WriteLine ("? " & objGrpItem.TextFrame.TextRange.Text)
                                    ElseIf shpx < 0.1 And shpy < 0.5 Then
                                    objTextFile.WriteLine ("? " & objGrpItem.TextFrame.TextRange.Text)
                                    Else
                                    objTextFile.WriteLine (" Text : " & objGrpItem.TextFrame.TextRange.Text)
                                End If
                            End If
                        End If
                    End If
                Next objGrpItem
            End If
        Next objshape

        ' Export Notes
        Set objShape4Note = objSlide.NotesPage.Shapes(2)
        If objShape4Note.HasTextFrame = True Then
            If Not objShape4Note.TextFrame.TextRange Is Nothing Then
                objTextFile.WriteLine "<!-- Notes: "
                objTextFile.WriteLine objShape4Note.TextFrame.TextRange.Text
                objTextFile.WriteLine "--> "
            End If
        End If

        ' End frame
        objTextFile.WriteLine ""
        objTextFile.WriteLine "---"
        objTextFile.WriteLine ""

    Next objSlide

    objTextFile.Close
    Set objTextFile = Nothing
    Set objFileSystem = Nothing
End Sub


