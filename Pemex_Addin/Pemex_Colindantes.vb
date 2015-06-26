Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants

Public Class Pemex_Colindantes

    Private x1stC = 12
    Private x2ndC = 57
    Private x3rdC = 24
    Private yTitleR = 7
    Private yRow = 5
    Private iCurOverlay
    Private angle As Double

    Public Sub New()

    End Sub

    Public Sub Colindantes_new()

        Try

            'create filter
            Loader.SIS.CreatePropertyFilter("fColindantes", "_FC&=1005")
            Loader.SIS.CreatePropertyFilter("f1055", "_FC&=1055")
            Loader.SIS.CreatePropertyFilter("fTxtHeight", "_character_height#=3.175")
            Loader.SIS.CombineFilter("fColTxt", "f1055", "fTxtHeight", SIS_BOOLEAN_AND)
            Loader.SIS.CreateListFromSelection("lSelection")
            Loader.SIS.OpenList("lSelection", 0)
            Dim bdt = False
            If Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "AnchoDdv#") = 0 Then bdt = True
            Loader.SIS.CreateLocusFromItem("Locus", SIS_GT_TOUCH, SIS_GM_GEOMETRY)
            Loader.SIS.CreateLocusFromItem("LOCUSINTERSECT", SIS_GT_INTERSECT, SIS_GM_GEOMETRY)
            Loader.SIS.CreateLocusFromItem("LOCUSWITHN", SIS_GT_WITHIN, SIS_GM_GEOMETRY)
            Loader.SIS.CreateBufferLocusFromItems("lSelection", False, "LOCUSBUFFER", Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "Buffer&"), 0)

            'set current overlay / remove col txt
            For i = 0 To Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1

                If Loader.SIS.GetInt(SIS_OT_OVERLAY, i, "_nDataset&") = Loader.SIS.GetDataset() Then

                    Loader.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", i)
                    iCurOverlay = i
                    If Loader.SIS.ScanOverlay("lDelete", iCurOverlay, "fColTxt", "") > 0 Then Loader.SIS.Delete("lDelete")
                    Exit For

                End If

            Next

            'find intersecting affectacions
            For i As Integer = 0 To (Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)

                If Loader.SIS.GetStr(SIS_OT_OVERLAY, i, "_name$") = "AFECTACIONES" Then

                    If Loader.SIS.ScanOverlay("List", i, "", "LOCUSINTERSECT") > 0 Then

                        Loader.SIS.CopyListItems("List")
                        Loader.SIS.SetListInt("List", "_FC&", 1005)
                        Exit For

                    End If

                End If

            Next

            'cut col areas to buffer
            Loader.SIS.CreateItemFromLocus("LOCUSBUFFER")
            Loader.SIS.SelectItem()
            Loader.SIS.DoCommand("AComBoundary")
            Loader.SIS.DoCommand("AComFillGeometry")
            Loader.SIS.OpenSel(0)
            Loader.SIS.SnipGeometry("List", 0)
            Loader.SIS.DeleteItem()

            'copy master bdt area
            If bdt = True Then

                For i As Integer = 0 To (Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)

                    If Loader.SIS.GetStr(SIS_OT_OVERLAY, i, "_name$") = "BDT" Then

                        Loader.SIS.CreateListFromOverlay(i, "lBDT")
                        Loader.SIS.CopyListItems("lBDT")
                        Exit For

                    End If

                Next

                'boolean BDT with affectacion area
                Loader.SIS.ScanList("lWithin", "List", "", "LOCUSWITHN")
                Loader.SIS.OpenList("lWithin", 0)
                Loader.SIS.AddToList("lBDT")
                Try
                    Loader.SIS.CreateBoolean("lBDT", SIS_BOOLEAN_DIFF)
                Catch
                End Try
                Loader.SIS.OpenList("lWithin", 0)

            Else

                Loader.SIS.ScanList("lWithin", "List", "", "LOCUSWITHN")
                Loader.SIS.Delete("lWithin")
                Loader.SIS.OpenList("lSelection", 0)

            End If

            If bdt = True Then Loader.SIS.Delete("lBDT")

            'decompose colindantes
            Loader.SIS.ScanOverlay("lColindantes", iCurOverlay, "fColindantes", "")
            Loader.SIS.DeselectAll()
            Loader.SIS.SelectList("lColindantes")
            Loader.SIS.DoCommand("AComDecompose")

            'clean colindantes area
            Loader.SIS.ScanOverlay("lColindantes", iCurOverlay, "fColindantes", "")
            CleanAreaGeometry("lColindantes")

            Loader.SIS.ScanOverlay("lColindantes", iCurOverlay, "fColindantes", "")

            'get outside colindantes
            EmptyList("List")
            Loader.SIS.OpenList("lSelection", 0)
            Dim nombre = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "Nombre$")
            Loader.SIS.CreateBufferLocusFromItems("lSelection", False, "LOCUSBUFFER", 1, 0)

            Loader.SIS.OpenList("lSelection", 0)
            Loader.SIS.AddToList("lColindantes")
            Loader.SIS.CreateBoolean("lColindantes", SIS_BOOLEAN_OR)
            Loader.SIS.SelectItem()
            Loader.SIS.DoCommand("AComBoundary")
            Loader.SIS.DoCommand("AComFillGeometry")
            Loader.SIS.OpenSel(0)
            Loader.SIS.AddToList("List")

            Loader.SIS.CreateItemFromLocus("LOCUSBUFFER")
            Loader.SIS.SelectItem()
            Loader.SIS.DoCommand("AComBoundary")
            Loader.SIS.DoCommand("AComFillGeometry")
            Loader.SIS.OpenSel(0)
            Loader.SIS.AddToList("List")

            Try
                Loader.SIS.CreateBoolean("List", SIS_BOOLEAN_DIFF)
                Loader.SIS.DeselectAll()
                Loader.SIS.Delete("List")
                Loader.SIS.AddToList("List")
                Loader.SIS.SelectList("List")
                Loader.SIS.DoCommand("AComDecompose")
                Loader.SIS.CreateListFromSelection("List")
                Loader.SIS.SetListStr("List", "Nombre$", nombre)
                Loader.SIS.SetListStr("List", "_featureTable$", "PEMEX")
                Loader.SIS.SetListInt("List", "_FC&", 1005)
                CleanAreaGeometry("List")
            Catch
                Loader.SIS.Delete("List")
            End Try

            'get colindantes list
            Loader.SIS.ScanOverlay("lColindantes", iCurOverlay, "fColindantes", "")

            'mostly old code here
            Dim x, y, z As Double
            EmptyList("lColTable")
            EmptyList("lPoints")
            EmptyList("lLines")

            Loader.SIS.OpenList("lSelection", 0)
            Dim sTitle = "DEL " & Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "KMI$") & " AL " & Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "KMF$")
            Dim nVertices = Loader.SIS.GetGeomNumPt(0)

            'create vertices
            For i = 0 To nVertices - 2

                Loader.SIS.OpenList("lSelection", 0)
                Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPt(0, i))
                Loader.SIS.CreatePoint(x, y, z, "Circle", 0, 1)
                Loader.SIS.UpdateItem()
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "i&", i)
                Loader.SIS.UpdateItem()
                Loader.SIS.AddToList("lPoints")

            Next

            For i = 0 To Loader.SIS.GetListSize("lColindantes") - 1

                Loader.SIS.OpenList("lColindantes", i)
                Loader.SIS.CreateLocusFromItem("Locus", SIS_GT_TOUCH, SIS_GM_GEOMETRY)

                If Loader.SIS.ScanList("lColPoints", "lPoints", "", "Locus") > 1 Then

                    Loader.SIS.OpenListCursor("cColPoints", "lColPoints", "i&" & vbTab & "_ox#" & vbTab & "_oy#")
                    Loader.SIS.MoveCursorToBegin("cColPoints")

                    'find inipoint
                    If Loader.SIS.GetCursorFieldInt("cColPoints", 0) = 0 Then

                        For ii = 0 To Loader.SIS.GetListSize("lColPoints") - 1

                            If Loader.SIS.GetCursorFieldInt("cColPoints", 0) > ii Then Exit For
                            Loader.SIS.MoveCursor("cColPoints", 1)
                            If ii = Loader.SIS.GetListSize("lColPoints") - 1 Then Loader.SIS.MoveCursorToBegin("cColPoints")

                        Next

                    End If

                    Loader.SIS.MoveTo(Loader.SIS.GetCursorFieldFlt("cColPoints", 1), Loader.SIS.GetCursorFieldFlt("cColPoints", 2), 0)

                    For ii = 0 To Loader.SIS.GetListSize("lColPoints") - 2

                        If Loader.SIS.MoveCursor("cColPoints", 1) = 0 Then Loader.SIS.MoveCursorToBegin("cColPoints")
                        Loader.SIS.LineTo(Loader.SIS.GetCursorFieldFlt("cColPoints", 1), Loader.SIS.GetCursorFieldFlt("cColPoints", 2), 0)

                    Next

                    Loader.SIS.UpdateItem()
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1007)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.AddToList("lLines")

                End If

            Next

            Loader.SIS.Delete("lPoints")

            'create table
            Col_CreateTable(Loader.SIS.GetListSize("lLines"), sTitle)

            'find NE line
            For i = 0 To Loader.SIS.GetListSize("lLines") - 1

                Loader.SIS.OpenList("lSelection", 0)
                Loader.SIS.MoveTo(Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#"), Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#"), 0)
                Loader.SIS.OpenList("lLines", i)
                Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_length#") / 2))
                Loader.SIS.LineTo(x, y, 0)
                Loader.SIS.UpdateItem()
                Dim lineangle = Loader.SIS.GetGeomAngleFromLength(0, 0) * 180 / Math.PI
                lineangle -= 90
                If lineangle < 0 Then lineangle += 450
                Loader.SIS.DeleteItem()
                Loader.SIS.OpenList("lLines", i)
                Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "lineangle#", lineangle)
                Loader.SIS.UpdateItem()

            Next

            Loader.SIS.OpenListCursor("cursor", "lLines", "lineangle#")
            Loader.SIS.OpenSortedCursor("sorted_cursor", "cursor", 0, False)
            Dim n As Integer = 0
            Do

                Loader.SIS.OpenCursorItem("sorted_cursor")
                Dim sLado = Col_GetLado_from_Centre()
                Loader.SIS.CreateBoxText((x1stC / 2), -(yTitleR + yRow + (yRow / 2) + (yRow * n)), 0, 2, sLado)
                SetTxtProperties(1019)
                Loader.SIS.AddToList("lColTable")

                Loader.SIS.OpenCursorItem("sorted_cursor")
                Dim sNombre = Col_GetNeighbour()
                Loader.SIS.CreateBoxText(x1stC + (x2ndC / 2), -(yTitleR + yRow + (yRow / 2) + (yRow * n)), 0, 2, sNombre)
                SetTxtProperties(1019)
                Loader.SIS.AddToList("lColTable")

                Loader.SIS.OpenCursorItem("sorted_cursor")
                Dim sLength = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_length#").ToString("F3")
                Loader.SIS.CreateBoxText(x1stC + x2ndC + (x3rdC / 2), -(yTitleR + yRow + (yRow / 2) + (yRow * n)), 0, 2, sLength)
                SetTxtProperties(1019)
                Loader.SIS.AddToList("lColTable")

                n += 1
            Loop Until Loader.SIS.MoveCursor("sorted_cursor", 1) = 0

            Dim lResponse As Integer
            Do

                lResponse = Loader.SIS.GetPosEx(x, y, z)
                Select Case lResponse

                    Case SIS_ARG_ENTER

                        angle = Loader.SIS.GetFlt(SIS_OT_WINDOW, 0, "_displayAngle#")
                        Loader.SIS.MoveList("lColTable", 0, 0, 0, angle, 1)
                        Loader.SIS.MoveList("lColTable", x, y, z, 0, 1)
                        Loader.SIS.CreateGroupFromItems("lColTable", True, "")
                        Exit Do

                    Case SIS_ARG_ESCAPE

                        Loader.SIS.Delete("lColTable")
                        Exit Do

                    Case SIS_ARG_BACKSPACE

                        Loader.SIS.Delete("lColTable")
                        Exit Do

                    Case SIS_ARG_POSITION

                        angle = Loader.SIS.GetFlt(SIS_OT_WINDOW, 0, "_displayAngle#")
                        Loader.SIS.MoveList("lColTable", 0, 0, 0, angle, 1)
                        Loader.SIS.MoveList("lColTable", x, y, z, 0, 1)
                        Loader.SIS.CreateGroupFromItems("lColTable", True, "gColTable")

                        PropertyTexts()

                        Exit Do

                End Select

            Loop

            Loader.SIS.DeselectAll()
            Loader.SIS.CombineLists("lSelect", "lLines", "lColindantes", SIS_BOOLEAN_OR)
            Loader.SIS.SelectList("lSelect")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub PropertyTexts()

        Try

            For i = 0 To Loader.SIS.GetListSize("lColindantes") - 1

                Loader.SIS.OpenList("lColindantes", i)
                Dim x = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                Dim y = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
                Try

                    Loader.SIS.CreateText(x, y, 0, Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "Nombre$"))
                    SetTxtProperties(1055)
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_font$", "Century Gothic")
                    Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", angle * 180 / Math.PI)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_point_height&", 9)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.SelectItem()
                    Loader.SIS.DoCommand("AComTextToBox")

                Catch
                End Try

            Next

        Catch ex As Exception

        End Try

    End Sub

    Private Function Col_GetLado() As String

        Try

            Dim xs, ys, xe, ye, z As Double

            Loader.SIS.SplitPos(xe, ye, z, Loader.SIS.GetGeomPt(0, 0))
            Loader.SIS.SplitPos(xs, ys, z, Loader.SIS.GetGeomPt(0, Loader.SIS.GetGeomNumPt(0) - 1))

            Loader.SIS.MoveTo(xs, ys, 0)
            Loader.SIS.LineTo(xe, ye, 0)
            Loader.SIS.UpdateItem()

            Dim angle = Loader.SIS.GetGeomAngleFromLength(0, 0) * 180 / Math.PI

            Loader.SIS.DeleteItem()

            angle -= 90
            If angle < 0 Then angle += 360

            Dim sLado As String = ""

            Select Case angle

                Case 0 To 22.5

                    sLado = "E"

                Case 22.5 To 77.5

                    sLado = "N-E"

                Case 77.5 To 122.5

                    sLado = "N"

                Case 122.5 To 167.5

                    sLado = "N-W"

                Case 167.5 To 212.5

                    sLado = "W"

                Case 212.5 To 257.5

                    sLado = "S-W"

                Case 257.5 To 302.5

                    sLado = "S"

                Case 302.5 To 347.5

                    sLado = "S-E"

                Case 347.5 To 360

                    sLado = "E"

            End Select

            Return sLado

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return ""
        End Try

    End Function

    Private Function Col_GetLado_from_Centre() As String

        Try

            angle = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "lineangle#")

            Dim sLado As String = ""

            Select Case angle

                Case 0 To 22.5

                    sLado = "N"

                Case 22.5 To 77.5

                    sLado = "N-W"

                Case 77.5 To 90

                    sLado = "W"

                Case 180 To 202.5

                    sLado = "W"

                Case 202.5 To 247.5

                    sLado = "S-W"

                Case 247.5 To 292.5

                    sLado = "S"

                Case 292.5 To 337.5

                    sLado = "S-E"

                Case 337.5 To 382.5

                    sLado = "E"

                Case 382.5 To 427.5

                    sLado = "N-E"

                Case 427.5 To 450

                    sLado = "N"

            End Select

            Return sLado

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return ""
        End Try

    End Function

    Private Function Col_GetNeighbour() As String

        Try

            Dim x, y, z As Double
            Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_length#") / 2))
            Loader.SIS.CreatePoint(x, y, z, "Circle", 0, 1)
            Loader.SIS.UpdateItem()
            Loader.SIS.CreateLocusFromItem("Locus", SIS_GT_TOUCH, SIS_GM_GEOMETRY)
            Loader.SIS.DeleteItem()
            If Loader.SIS.ScanList("lColindante", "lColindantes", "", "Locus") = 1 Then

                Loader.SIS.OpenList("lColindante", 0)
                Return Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "Nombre$")

            End If
            Return ""

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return ""
        End Try

    End Function

    Private Sub Col_CreateTable(ByVal nColindantes As Integer, ByVal sTitle As String)

        Try

            Dim xTotal = x1stC + x2ndC + x3rdC

            'box
            Loader.SIS.CreateRectangle(0, 0, xTotal, -(yTitleR + yRow + (nColindantes * yRow)))
            Loader.SIS.UpdateItem()
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1016)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lColTable")

            'row lines
            For i = 0 To nColindantes

                Loader.SIS.MoveTo(0, -(yTitleR + (i * yRow)), 0)
                Loader.SIS.LineTo(xTotal, -(yTitleR + (i * yRow)), 0)
                Loader.SIS.UpdateItem()
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1017)
                Loader.SIS.UpdateItem()
                Loader.SIS.AddToList("lColTable")

            Next i

            '1st column
            Loader.SIS.MoveTo(x1stC, -(yTitleR), 0)
            Loader.SIS.LineTo(x1stC, -(yTitleR + yRow + (nColindantes * yRow)), 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1017)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lColTable")

            '2nd column
            Loader.SIS.MoveTo(x1stC + x2ndC, -(yTitleR), 0)
            Loader.SIS.LineTo(x1stC + x2ndC, -(yTitleR + yRow + (nColindantes * yRow)), 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1017)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lColTable")

            'title
            Loader.SIS.CreateBoxText((xTotal / 2), -(yTitleR * 0.25), 0, 2.7, "CUADRO DE COLINDANCIAS DEL POLIGONO")
            SetTxtProperties(1018)
            Loader.SIS.AddToList("lColTable")
            Loader.SIS.CreateBoxText((xTotal / 2), -(yTitleR * 0.75), 0, 2.7, sTitle)
            SetTxtProperties(1018)
            Loader.SIS.AddToList("lColTable")

            'row header
            Loader.SIS.CreateBoxText((x1stC / 2), -(yTitleR + (yRow / 2)), 0, 2, "LADO")
            SetTxtProperties(1019)
            Loader.SIS.AddToList("lColTable")
            Loader.SIS.CreateBoxText(x1stC + (x2ndC / 2), -(yTitleR + (yRow / 2)), 0, 2, "COLINDANTES")
            SetTxtProperties(1019)
            Loader.SIS.AddToList("lColTable")
            Loader.SIS.CreateBoxText(x1stC + x2ndC + (x3rdC / 2), -(yTitleR + (yRow / 2)), 0, 2, "DISTANCIA (m)")
            SetTxtProperties(1019)
            Loader.SIS.AddToList("lColTable")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub SetTxtProperties(ByVal FC As Integer)

        Try

            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", FC)
            Loader.SIS.UpdateItem()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub SetCurrentOverlay()

        Try

            For i = 0 To Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1

                If Loader.SIS.GetInt(SIS_OT_OVERLAY, i, "_nDataset&") = Loader.SIS.GetDataset() Then

                    Loader.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", i)
                    iCurOverlay = i
                    Exit For

                End If

            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub CleanAreaGeometry(ByVal List As String)

        Try

            For i = 0 To Loader.SIS.GetListSize(List) - 1

                Loader.SIS.DeselectAll()
                Loader.SIS.OpenList(List, i)
                Loader.SIS.SelectItem()
                Loader.SIS.DoCommand("AComBoundary")
                Loader.SIS.CreateListFromSelection("lClean")
                Loader.SIS.CleanLines("lClean", 0, SIS_CLEAN_LINE_NONE + SIS_CLEAN_LINE_REMOVE_180 + SIS_CLEAN_LINE_REMOVE_SELF)
                If Loader.SIS.GetListSize("lClean") > 0 Then Loader.SIS.DoCommand("AComFillGeometry")

            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub EmptyList(ByVal List As String)

        Try

            Loader.SIS.EmptyList(List)

        Catch
        End Try

    End Sub

End Class
