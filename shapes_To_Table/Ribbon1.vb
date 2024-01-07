Imports Microsoft.Office.Core
Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim app As PowerPoint.Application
        app = Globals.ThisAddIn.Application
        Dim tyape As PowerPoint.Selection

        If app.ActiveWindow.Selection.Type <> tyape.Type.ppSelectionShapes Then


            System.Windows.Forms.MessageBox.Show("Please Select Shapes First")
        Else
            Dim shp As PowerPoint.Shape
            Dim sld As PowerPoint.Slide
            Dim ss As PowerPoint.Shape
            Dim maxrow As Integer
            Dim maxcol As Integer
            Dim lowestshape As PowerPoint.Shape
            Dim highestshape As PowerPoint.Shape
            Dim val()
            Dim arr()
            Dim arr2()
            Dim colr As Integer
            Dim r As Integer
            Dim g As Integer
            Dim b As Integer
            Dim otbl As PowerPoint.Table
            Dim flag As Boolean
            Dim myshape As PowerPoint.Shape
            Dim ww As Double
            Dim hh As Double
            maxrow = 0
            maxcol = 0
            sld = app.ActiveWindow.View.Slide

            Call FixeShapes(app, sld)
            For Each shp In app.ActiveWindow.Selection.ShapeRange

                If InStr(shp.Name, ",") > 0 Then
                    If InStr(shp.Name, "&") > 0 Then
                        val = Split(shp.Name, "&")
                        arr = Split(val(UBound(val)), ",")
                    Else
                        arr = Split(shp.Name, ",")
                        ww = shp.Width
                        hh = shp.Height

                    End If


                    If CInt(arr(0)) > maxrow Then
                        maxrow = CInt(arr(0))
                    End If

                    If CInt(arr(1)) > maxcol Then
                        maxcol = CInt(arr(1))
                    End If

                    If InStr(shp.Name, "1,1") > 0 Then
                        lowestshape = shp
                    ElseIf InStr(shp.Name, maxcol & "," & maxcol) > 0 Then
                        highestshape = shp
                    End If


                End If

            Next

            ss = sld.Shapes.AddTable(maxrow, maxcol, lowestshape.Left, lowestshape.Top, -1, -1)
            otbl = ss.Table
            With otbl
                For lRow = 1 To maxrow
                    For lCol = 1 To maxcol
                        flag = False
                        On Error Resume Next
                        With .Cell(lRow, lCol).Shape
                            For Each shp In sld.Shapes
                                If shp.Name = lRow & "," & lCol Then
                                    myshape = shp
                                    Exit For
                                ElseIf InStr(shp.Name, "&") > 0 Then
                                    val = Split(shp.Name, "&")
                                    For v = LBound(val) To UBound(val)
                                        If val(v) = lRow & "," & lCol Then
                                            myshape = shp
                                            flag = True
                                            Exit For
                                        End If
                                    Next


                                End If
                                If flag Then
                                    Exit For
                                End If
                            Next


                            colr = myshape.Fill.ForeColor.RGB
                            r = colr Mod 256
                            g = (colr \ 256) Mod 256
                            b = (colr \ 65536) Mod 256
                            .Fill.ForeColor.RGB = RGB(r, g, b)
                            .TextFrame.TextRange.Font.Size = myshape.TextFrame.TextRange.Font.Size
                            colr = myshape.TextFrame.TextRange.Font.Color.RGB
                            r = colr Mod 256
                            g = (colr \ 256) Mod 256
                            b = (colr \ 65536) Mod 256
                            .TextFrame.TextRange.Font.Color.RGB = RGB(r, g, b)
                            .TextFrame.TextRange.Text = myshape.TextFrame.TextRange.Text



                            colr = myshape.Line.ForeColor.RGB
                            r = colr Mod 256
                            g = (colr \ 256) Mod 256
                            b = (colr \ 65536) Mod 256


                            otbl.Cell(lRow, lCol).Borders(1).ForeColor.RGB = RGB(r, g, b)
                            otbl.Cell(lRow, lCol).Borders(2).ForeColor.RGB = RGB(r, g, b)
                            otbl.Cell(lRow, lCol).Borders(3).ForeColor.RGB = RGB(r, g, b)
                            otbl.Cell(lRow, lCol).Borders(4).ForeColor.RGB = RGB(r, g, b)

                            If InStr(myshape.Name, "&") < 1 Then
                                myshape.Delete()

                            End If
                        End With
                    Next    ' column
                Next    ' row

            End With

            For Each shp In sld.Shapes
                If InStr(shp.Name, "&") Then
                    val = Split(shp.Name, "&")

                    arr = Split(val(0), ",")
                    arr2 = Split(val(UBound(val)), ",")

                    otbl.Cell(CInt(arr(0)), CInt(arr(1))).Merge(MergeTo:=otbl.Cell(CInt(arr2(0)), CInt(arr2(1))))
                    otbl.Cell(CInt(arr(0)), CInt(arr(1))).Shape.TextFrame.TextRange.Text = shp.TextFrame.TextRange.Text

                    otbl.Cell(CInt(arr(0)), CInt(arr(1))).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = shp.TextFrame.TextRange.ParagraphFormat.Alignment

                End If
            Next
            For Each shp In sld.Shapes
                If InStr(shp.Name, "&") Then
                    shp.Delete()
                End If
            Next
            For i = 1 To otbl.Columns.Count
                otbl.Columns(i).Width = ww
            Next
            For i = 1 To otbl.Rows.Count
                otbl.Rows(i).Height = hh
            Next
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        Dim app As PowerPoint.Application
        app = Globals.ThisAddIn.Application
        Dim tyape As PowerPoint.PpSelectionType
        Dim colr As Integer
        Dim r As Integer
        Dim g As Integer
        Dim b As Integer
        Dim oTbl As PowerPoint.Table
        Dim lRow As Long
        Dim lCol As Long
        Dim sld As PowerPoint.Slide
        Dim ss As PowerPoint.Shape
        Dim shp As PowerPoint.Shape
        Dim sh As PowerPoint.Shape
        On Error GoTo Handle
        If app.ActiveWindow.Selection.Type = tyape.ppSelectionNone Then

            System.Windows.Forms.MessageBox.Show("Please Select a Table Shape First ")
        Else
            If app.ActiveWindow.Selection.ShapeRange(1).HasTable Then




                sld = app.ActiveWindow.View.Slide
                ' Get a reference to a table either programmatically or
                ' for demonstration purposes, by referencing the currently
                ' selected table:
                ss = app.ActiveWindow.Selection.ShapeRange(1)
                oTbl = app.ActiveWindow.Selection.ShapeRange(1).Table
                With oTbl
                    For lRow = 1 To .Rows.Count
                        For lCol = 1 To .Columns.Count
                            On Error Resume Next
                            With .Cell(lRow, lCol).Shape

                                shp = sld.Shapes.AddShape(1, .Left, .Top, .Width, .Height)


                                shp.Fill.BackColor.RGB = RGB(r, g, b)
                                colr = .Fill.ForeColor.RGB

                                r = colr Mod 256
                                g = (colr \ 256) Mod 256
                                b = (colr \ 65536) Mod 256
                                shp.Fill.ForeColor.RGB = RGB(r, g, b)
                                ''shp.Name = lRow & "," & lCol
                                shp.TextFrame.TextRange.Font.Size = .TextFrame.TextRange.Font.Size
                                colr = .TextFrame.TextRange.Font.Color.RGB
                                r = colr Mod 256
                                g = (colr \ 256) Mod 256
                                b = (colr \ 65536) Mod 256
                                shp.TextFrame.TextRange.Font.Color.RGB = RGB(r, g, b)
                                shp.TextFrame.TextRange.Text = .TextFrame.TextRange.Text


                                colr = oTbl.Cell(lRow, lCol).Borders(4).ForeColor.RGB

                                r = colr Mod 256
                                g = (colr \ 256) Mod 256
                                b = (colr \ 65536) Mod 256

                                shp.Line.ForeColor.RGB = RGB(r, g, b)

                                For Each sh In sld.Shapes
                                    If sh.Name <> shp.Name And (shp.Top = sh.Top And sh.Left = shp.Left And sh.Width = shp.Width And shp.Height = sh.Height) Then
                                        '' sh.Name = sh.Name & "&" & shp.Name
                                        shp.Delete()
                                        Exit For


                                    End If
                                Next

                            End With
                        Next    ' column
                    Next    ' row

                End With


                oTbl = Nothing
                ss.Delete()




            Else
                System.Windows.Forms.MessageBox.Show("Selected Shape is not a Table")
            End If
        End If

        Exit Sub
Handle:
        System.Windows.Forms.MessageBox.Show("Selected Shape is not a Table Or Some other issue.Please Contact Developer if you are unable to fix it.")
    End Sub





























    Public Function FixeShapes(app As PowerPoint.Application, sld As PowerPoint.Slide)

        Dim shp As PowerPoint.Shape
        Dim tops As Scripting.Dictionary
        Dim lefts As Scripting.Dictionary
        Dim lowWidth As Integer
        Dim lowHeight As Integer
        Dim numberarray()
        Dim numbers()
        Dim ro As Integer
        Dim col As Integer
        Dim r As Integer, c As Integer, kk As Integer, coloumns As Integer, rows As Integer
        lowWidth = 9999
        lowHeight = 999
        tops = New Scripting.Dictionary
        lefts = New Scripting.Dictionary


        For Each shp In app.ActiveWindow.Selection.ShapeRange
            kk = myExsists(lefts, CInt(shp.Left))
            If kk <> -1 Then
                lefts(kk) = lefts(kk) & "," & shp.Id


            Else
                lefts(CInt(shp.Left)) = shp.Id

            End If
            kk = myExsists(tops, CInt(shp.Top))
            If kk <> -1 Then
                tops(kk) = tops(kk) & "," & shp.Id

            Else
                tops(CInt(shp.Top)) = shp.Id
            End If

        Next
        coloumns = lefts.Count
        rows = tops.Count
        Dim numbering As String
        numbering = ""
        For i = 1 To rows
            For j = 1 To coloumns
                numbering = numbering & i & "," & j & "&"
            Next
        Next
        numbering = Left(numbering, Len(numbering) - 1)
        numberarray = Split(numbering, "&")

        Dim sortedLefts()
        Dim sortedTops()
        Dim lowest As Integer
        ReDim sortedLefts(lefts.Count - 1)
        ReDim sortedTops(tops.Count - 1)

        'Sorting Lefts
        For i = LBound(sortedLefts) To UBound(sortedLefts)
            lowest = 9999999
            For Each key In lefts.Keys
                If CInt(key) < lowest And Not checkPresent(sortedLefts, CInt(key)) Then
                    lowest = CInt(key)
                End If

            Next
            sortedLefts(i) = lowest
        Next



        'Sorting Tops
        For i = LBound(sortedTops) To UBound(sortedTops)
            lowest = 9999999
            For Each key In tops.Keys
                If CInt(key) < lowest And Not checkPresent(sortedTops, CInt(key)) Then
                    lowest = CInt(key)
                End If

            Next
            sortedTops(i) = lowest
        Next


        For Each shp In app.ActiveWindow.Selection.ShapeRange
            r = -1
            c = -1
            For i = 0 To UBound(sortedTops)
                If Absmatch(CInt(shp.Top), CInt(sortedTops(i))) Then
                    r = i + 1
                    Exit For
                End If
            Next
            For i = 0 To UBound(sortedLefts)
                If Absmatch(CInt(shp.Left), CInt(sortedLefts(i))) Then
                    c = i + 1
                    Exit For
                End If
            Next
            shp.Name = r & "," & c
        Next



        numbers = Split(numbering, "&")
        Dim arr()

        For i = 0 To UBound(numbers)
            If Not shapePresent(sld, numbers(i)) Then
                arr = Split(numbers(i), ",")
                col = CInt(arr(1)) - 1
                ro = CInt(arr(0)) - 1
                For Each shp In app.ActiveWindow.Selection.ShapeRange

                    If shp.Left < sortedLefts(col) And (shp.Left + shp.Width) > (sortedLefts(col) + 15) And Absmatch(CInt(shp.Top), CInt(sortedTops(ro))) Then
                        shp.Name = shp.Name & "&" & numbers(i)
                        Exit For
                    ElseIf shp.Top < sortedTops(ro) And shp.Top + shp.Height > (sortedTops(ro) + 15) And Absmatch(CInt(shp.Left), CInt(sortedLefts(col))) Then
                        shp.Name = shp.Name & "&" & numbers(i)
                        Exit For
                    End If
                Next

            End If

        Next


    End Function



    Private Function shapePresent(ByVal sl As PowerPoint.Slide, ByVal myshapeName As String) As Boolean

        Dim myShape As PowerPoint.Shape

        On Error Resume Next

        myShape = sl.Shapes(myshapeName)

        On Error GoTo 0

        shapePresent = Not myShape Is Nothing

    End Function


    Public Function getMatch(s1 As String, s2 As String)
        Dim arr()
        Dim arr2()
        arr = Split(s1, ",")
        arr2 = Split(s2, ",")
    End Function
    Public Function checkPresent(arr As Array, val As Integer) As Boolean
        Dim res As Boolean
        res = False

        For i = LBound(arr) To UBound(arr)
            If arr(i) = val Then
                res = True
                Exit For
            End If
        Next
        checkPresent = res
    End Function



    Public Function myExsists(dic As Scripting.Dictionary, key As Integer) As Integer
        Dim res As Double

        res = -1

        For Each valv In dic.Keys
            If Absmatch(CInt(valv), key) Then
                res = valv
                Exit For
            End If
        Next

        myExsists = res
    End Function


    Function Absmatch(a As Integer, b As Integer)
        If Math.Abs(a - b) < 5 Then
            Absmatch = True
        Else
            Absmatch = False
        End If

    End Function




End Class




