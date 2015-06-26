Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants

Public Class Pemex_Vertices

    Private x1stC = 6
    Private x2ndC = 6
    Private x3rdC = 24
    Private x4thC = 18
    Private x5thC = 6
    Private x6thC = 16.5
    Private x7thC = 16.5
    Private yTitleR = 7
    Private yRow = 5
    Private vx() As Double
    Private vy() As Double
    Private vangle() As Double
    Private vsector() As Double
    Private vdist() As Double
    Private sArea As Double
    Private iCurOverlay
    Private angle As Double
    Private bdt As Boolean

    Public Sub New()

    End Sub

    Public Sub Construccion()

        Try

            angle = Loader.SIS.GetFlt(SIS_OT_WINDOW, 0, "_displayAngle#")
            EmptyList("lVerTable")
            Loader.SIS.CreateClassTreeFilter("FLine", "-Item +Line")
            Dim x, y, z As Double

            Loader.SIS.CreateListFromSelection("lSelection")
            Loader.SIS.OpenList("lSelection", 0)
            SetCurrentOverlay()
            sArea = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_area#")
            bdt = False
            If Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "AnchoDdv#") = 0 Then bdt = True

            Dim nVertices = Loader.SIS.GetGeomNumPt(0)
            ReDim vx(nVertices - 2)
            ReDim vy(nVertices - 2)
            ReDim vangle(nVertices - 2)
            ReDim vsector(nVertices - 2)
            ReDim vdist(nVertices - 2)
            For i As Integer = 0 To nVertices - 2

                Loader.SIS.OpenList("lSelection", 0)
                Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPt(0, i))
                vx(i) = x
                vy(i) = y

            Next

            ' check whether points are clockwise
            Dim Clockwise = 1
            Dim SignedArea As Double = 0
            For i As Integer = 0 To nVertices - 3

                SignedArea += (vx(i) * vy(i + 1) - vx(i + 1) * vy(i))

            Next
            SignedArea += (vx(nVertices - 2) * vy(0) - vx(0) * vy(nVertices - 2))
            If SignedArea > 0 Then Clockwise = -1

            ' check north eastern most point
            Dim NE As Double = 0
            Dim iNE As Integer = -1
            For i As Integer = 0 To nVertices - 2
                If vx(i) + vy(i) > NE Then
                    NE = vx(i) + vy(i)
                    iNE = i
                End If
            Next

            ' reorganiza los vertices
            Dim ii = iNE
            For i As Integer = 0 To nVertices - 2

                'n
                If i = 0 Then
                    Loader.SIS.MoveTo(vx(ii), vy(ii), 0)
                Else
                    Loader.SIS.LineTo(vx(ii), vy(ii), 0)
                End If

                'Loader.SIS.SetGeomPt(0, i, vx(ii), vy(ii), 0)
                ii += Clockwise
                If ii > (nVertices - 2) Then ii = 0
                If ii < 0 Then ii = (nVertices - 2)

            Next
            Loader.SIS.UpdateItem()

            'n
            EmptyList("lTmp")
            Loader.SIS.AddToList("lTmp")

            ' get array of ordered vertices
            For i As Integer = 0 To nVertices - 2

                'Loader.SIS.OpenList("lSelection", 0)
                Loader.SIS.OpenList("lTmp", 0)
                Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPt(0, i))
                vx(i) = x
                vy(i) = y

            Next

            'n
            Loader.SIS.Delete("lTmp")

            ' get angle and distances
            For i As Integer = 0 To nVertices - 2

                If Not i = nVertices - 2 Then
                    Dim ydif = vy(i + 1) - vy(i)
                    Dim xdif = vx(i + 1) - vx(i)
                    vangle(i) = Math.Atan2(ydif, xdif) * (180 / Math.PI)
                    vdist(i) = Distance(vx(i), vy(i), vx(i + 1), vy(i + 1))

                Else
                    Dim ydif = vy(0) - vy(i)
                    Dim xdif = vx(0) - vx(i)
                    vangle(i) = Math.Atan2(ydif, xdif) * (180 / Math.PI)
                    vdist(i) = Distance(vx(i), vy(i), vx(0), vy(0))

                End If

                Loader.SIS.CreatePoint(vx(i), vy(i), 0, "", 0, 1)
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1059)
                If vdist(i) < 0.2 Then Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1065)
                Loader.SIS.UpdateItem()
                Loader.SIS.SelectItem()
                Loader.SIS.DoCommand("AComExplodeShape")
                Loader.SIS.CreateText(vx(i), vy(i), 0, Str(i + 1))
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1059)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_point_height&", 6)
                Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", angle * 180 / Math.PI)
                Loader.SIS.UpdateItem()
                Loader.SIS.SelectItem()
                Loader.SIS.DoCommand("AComTextToBox")

            Next

            Ver_CreateTable(nVertices + 1)

            Dim lResponse As Integer
            Do

                lResponse = Loader.SIS.GetPosEx(x, y, z)
                Select Case lResponse

                    Case SIS_ARG_ENTER

                        Loader.SIS.Delete("lVerTable")
                        Exit Do

                    Case SIS_ARG_ESCAPE

                        ' Ending on Escape (this can be caused by starting another non-transparent command)
                        Loader.SIS.Delete("lVerTable")
                        Exit Do

                    Case SIS_ARG_BACKSPACE

                        ' Backspace key pressed
                        Loader.SIS.Delete("lVerTable")
                        Exit Do

                    Case SIS_ARG_POSITION

                        Loader.SIS.MoveList("lVerTable", 0, 0, 0, angle, 1)
                        Loader.SIS.MoveList("lVerTable", x, y, 0, 0, 1)
                        Loader.SIS.CreateGroupFromItems("lVerTable", True, "gVerTable")
                        Exit Do

                End Select

            Loop

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Function Distance(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Single

        Dim X As Double
        Dim Y As Double
        Dim Dist As Double
        X = (x2 - x1) * (x2 - x1)
        Y = (y2 - y1) * (y2 - y1)
        Dist = Math.Sqrt(X + Y)

        Return Dist

    End Function

    Private Sub Ver_CreateTable(ByVal n As Integer)

        Try

            Dim xTotal = x1stC + x2ndC + x3rdC + x4thC + x5thC + x6thC + x7thC

            'box
            Loader.SIS.CreateRectangle(0, 0, xTotal, -(yTitleR + yRow + (n * yRow)))
            Loader.SIS.UpdateItem()
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1016)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lVerTable")

            'row lines
            For i = 0 To n
                Loader.SIS.MoveTo(0, -(yTitleR + (i * yRow)), 0)
                Loader.SIS.LineTo(xTotal, -(yTitleR + (i * yRow)), 0)
                Loader.SIS.UpdateItem()
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1017)
                Loader.SIS.UpdateItem()
                Loader.SIS.AddToList("lVerTable")
            Next i

            'sub header lado
            Loader.SIS.MoveTo(0, -(yTitleR + (yRow / 2)), 0)
            Loader.SIS.LineTo(x1stC + x2ndC, -(yTitleR + (yRow / 2)), 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1017)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lVerTable")

            'sub header coordenadas
            Loader.SIS.MoveTo(x1stC + x2ndC + x3rdC + x4thC + x5thC, -(yTitleR + (yRow / 2)), 0)
            Loader.SIS.LineTo(xTotal, -(yTitleR + (yRow / 2)), 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1017)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lVerTable")

            '1st column
            Loader.SIS.MoveTo(x1stC, -(yTitleR + (yRow / 2)), 0)
            Loader.SIS.LineTo(x1stC, -(yTitleR + yRow + ((n - 1) * yRow)), 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1017)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lVerTable")

            '2nd column
            Loader.SIS.MoveTo(x1stC + x2ndC, -(yTitleR), 0)
            Loader.SIS.LineTo(x1stC + x2ndC, -(yTitleR + yRow + ((n - 1) * yRow)), 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1017)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lVerTable")

            '3rd column
            Loader.SIS.MoveTo(x1stC + x2ndC + x3rdC, -(yTitleR), 0)
            Loader.SIS.LineTo(x1stC + x2ndC + x3rdC, -(yTitleR + yRow + ((n - 1) * yRow)), 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1017)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lVerTable")

            '4th column
            Loader.SIS.MoveTo(x1stC + x2ndC + x3rdC + x4thC, -(yTitleR), 0)
            Loader.SIS.LineTo(x1stC + x2ndC + x3rdC + x4thC, -(yTitleR + yRow + ((n - 1) * yRow)), 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1017)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lVerTable")

            '5th column
            Loader.SIS.MoveTo(x1stC + x2ndC + x3rdC + x4thC + x5thC, -(yTitleR), 0)
            Loader.SIS.LineTo(x1stC + x2ndC + x3rdC + x4thC + x5thC, -(yTitleR + yRow + ((n - 1) * yRow)), 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1017)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lVerTable")

            '6th column
            Loader.SIS.MoveTo(x1stC + x2ndC + x3rdC + x4thC + x5thC + x6thC, -(yTitleR + (yRow / 2)), 0)
            Loader.SIS.LineTo(x1stC + x2ndC + x3rdC + x4thC + x5thC + x6thC, -(yTitleR + yRow + ((n - 1) * yRow)), 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1017)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lVerTable")

            'title
            Loader.SIS.CreateBoxText((xTotal / 2), -(yTitleR / 2), 0, 3, "CUADRO DE CONSTRUCCIÓN DE SUPERFICIE A LEGALIZAR")
            If bdt = True Then Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", "CUADRO DE CONSTRUCCIÓN DE PAGO BDT")
            SetTxtProperties(1018)
            Loader.SIS.AddToList("lVerTable")

            'row header
            Loader.SIS.CreateBoxText((x1stC + x2ndC) / 2, -(yTitleR + (yRow * 0.25)), 0, 2, "LADO")
            SetTxtProperties(1019)
            Loader.SIS.AddToList("lVerTable")
            Loader.SIS.CreateBoxText((x1stC / 2), -(yTitleR + (yRow * 0.75)), 0, 2, "EST")
            SetTxtProperties(1019)
            Loader.SIS.AddToList("lVerTable")
            Loader.SIS.CreateBoxText(x1stC + (x2ndC / 2), -(yTitleR + (yRow * 0.75)), 0, 2, "PV")
            SetTxtProperties(1019)
            Loader.SIS.AddToList("lVerTable")

            Loader.SIS.CreateBoxText(x1stC + x2ndC + (x3rdC / 2), -(yTitleR + (yRow / 2)), 0, 2, "RUMBO")
            SetTxtProperties(1019)
            Loader.SIS.AddToList("lVerTable")
            Loader.SIS.CreateBoxText(x1stC + x2ndC + x3rdC + (x4thC / 2), -(yTitleR + (yRow / 2)), 0, 2, "DISTANCIA")
            SetTxtProperties(1019)
            Loader.SIS.AddToList("lVerTable")
            Loader.SIS.CreateBoxText(x1stC + x2ndC + x3rdC + x4thC + (x5thC / 2), -(yTitleR + (yRow / 2)), 0, 2, "V")
            SetTxtProperties(1019)
            Loader.SIS.AddToList("lVerTable")

            Loader.SIS.CreateBoxText(x1stC + x2ndC + x3rdC + x4thC + x5thC + (x6thC + x7thC) / 2, -(yTitleR + (yRow * 0.25)), 0, 2, "C O O R D E N A D A S")
            SetTxtProperties(1019)
            Loader.SIS.AddToList("lVerTable")
            Loader.SIS.CreateBoxText(x1stC + x2ndC + x3rdC + x4thC + x5thC + (x6thC / 2), -(yTitleR + (yRow * 0.75)), 0, 2, "X")
            SetTxtProperties(1019)
            Loader.SIS.AddToList("lVerTable")
            Loader.SIS.CreateBoxText(x1stC + x2ndC + x3rdC + x4thC + x5thC + x6thC + (x7thC / 2), -(yTitleR + (yRow * 0.75)), 0, 2, "Y")
            SetTxtProperties(1019)
            Loader.SIS.AddToList("lVerTable")

            Dim ii = 0
            For i = 1 To n - 1
                If ii = n - 2 Then ii = 0
                'PV
                Loader.SIS.CreateBoxText(x1stC + (x2ndC / 2), -(yTitleR + (yRow * i) + (yRow / 2)), 0, 2, Str(ii + 1))
                SetTxtProperties(1019)
                Loader.SIS.AddToList("lVerTable")
                'V
                Loader.SIS.CreateBoxText(x1stC + x2ndC + x3rdC + x4thC + (x5thC / 2), -(yTitleR + (yRow * i) + (yRow / 2)), 0, 2, Str(ii + 1))
                SetTxtProperties(1019)
                Loader.SIS.AddToList("lVerTable")
                'X
                Loader.SIS.CreateBoxText(x1stC + x2ndC + x3rdC + x4thC + x5thC + (x6thC / 2), -(yTitleR + (yRow * i) + (yRow / 2)), 0, 2, String.Format("{0:##,###,###.000}", vx(ii)))
                SetTxtProperties(1019)
                Loader.SIS.AddToList("lVerTable")
                'Y
                Loader.SIS.CreateBoxText(x1stC + x2ndC + x3rdC + x4thC + x5thC + x6thC + (x7thC / 2), -(yTitleR + (yRow * i) + (yRow / 2)), 0, 2, String.Format("{0:##,###,###.000}", vy(ii)))
                SetTxtProperties(1019)
                Loader.SIS.AddToList("lVerTable")
                If Not i = n - 1 Then
                    'EST
                    Loader.SIS.CreateBoxText((x1stC / 2), -(yTitleR + yRow + (yRow * i) + (yRow / 2)), 0, 2, Str(ii + 1))
                    SetTxtProperties(1019)
                    Loader.SIS.AddToList("lVerTable")
                    'RUMBO
                    Loader.SIS.CreateBoxText(x1stC + x2ndC + (x3rdC / 2), -(yTitleR + yRow + (yRow * i) + (yRow / 2)), 0, 2, GetRumbo(vangle(ii)))
                    SetTxtProperties(1019)
                    Loader.SIS.AddToList("lVerTable")
                    'DISTANCIA
                    Loader.SIS.CreateBoxText(x1stC + x2ndC + x3rdC + (x4thC / 2), -(yTitleR + yRow + (yRow * i) + (yRow / 2)), 0, 2, String.Format("{0:###,###,##0.000}", vdist(ii)))
                    SetTxtProperties(1019)
                    Loader.SIS.AddToList("lVerTable")
                End If
                ii += 1
            Next i

            'SUPERFICIE
            Loader.SIS.CreateBoxText((xTotal / 2), -(yTitleR + (yRow * n) + (yRow / 2)), 0, 3, String.Format("SUPERFICIE = {0:###,###,###.000} m²", sArea))
            SetTxtProperties(1018)
            Loader.SIS.AddToList("lVerTable")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Function GetRumbo(ByVal _Angle As Single) As String

        Dim Degrees As Single
        Dim Minutes As Single
        Dim Seconds As Single

        If _Angle < 0 Then _Angle += 360
        Dim Rumbo As String = ""
        Select Case _Angle

            Case 0 To 90

                _Angle = Math.Abs(90 - _Angle)
                Degrees = Int(_Angle)
                Minutes = (_Angle - Degrees) * 60
                Seconds = (Minutes - Int(Minutes)) * 60
                Return String.Format("N{0:00}°{1:00}'{2:00.00}''E", Degrees, Int(Minutes), Seconds)

            Case 90 To 180

                _Angle = Math.Abs(90 - _Angle)
                Degrees = Int(_Angle)
                Minutes = (_Angle - Degrees) * 60
                Seconds = (Minutes - Int(Minutes)) * 60
                Return String.Format("N{0:00}°{1:00}'{2:00.00}''W", Degrees, Int(Minutes), Seconds)

            Case 180 To 270
                _Angle = Math.Abs(270 - _Angle)
                Degrees = Int(_Angle)
                Minutes = (_Angle - Degrees) * 60
                Seconds = (Minutes - Int(Minutes)) * 60
                Return String.Format("S{0:00}°{1:00}'{2:00.00}''W", Degrees, Int(Minutes), Seconds)

            Case 270 To 360
                _Angle = Math.Abs(270 - _Angle)
                Degrees = Int(_Angle)
                Minutes = (_Angle - Degrees) * 60
                Seconds = (Minutes - Int(Minutes)) * 60
                Return String.Format("S{0:00}°{1:00}'{2:00.00}''E", Degrees, Int(Minutes), Seconds)

            Case Else

                Return ""

        End Select

    End Function

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

    Public Sub Vertices()

        Try

            Loader.SIS.DeselectAll()
            Loader.SIS.CreatePropertyFilter("fVertices", "_FC&=1059")
            Loader.SIS.CreateClassTreeFilter("fText", "-Item +BoxText")
            Loader.SIS.CombineFilter("fVerticesText", "fVertices", "fText", SIS_BOOLEAN_AND)
            If Loader.SIS.Scan("lVertices", "E", "fVerticesText", "") > 0 Then

                Loader.SIS.OpenList("lVertices", 0)
                SetCurrentOverlay()

                Dim x, y, z As Double
                Loader.SIS.DefineNolView("view")

                For i = 0 To Loader.SIS.GetListSize("lVertices") - 1

                    Loader.SIS.OpenList("lVertices", i)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1060)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.SelectItem()
                    Loader.SIS.DoCommand("AComZoomSelect")
                    Loader.SIS.SetFlt(SIS_OT_WINDOW, 0, "_displayScale#", 250)

                    Dim lResponse As Integer
                    Do

                        lResponse = Loader.SIS.GetPosEx(x, y, z)
                        Select Case lResponse

                            Case SIS_ARG_ENTER

                                Loader.SIS.OpenList("lVertices", i)
                                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1059)
                                Loader.SIS.UpdateItem()
                                Loader.SIS.DeselectAll()
                                Exit Do

                            Case SIS_ARG_ESCAPE

                                Loader.SIS.OpenList("lVertices", i)
                                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1059)
                                Loader.SIS.UpdateItem()
                                Loader.SIS.DeselectAll()
                                Exit For

                            Case SIS_ARG_BACKSPACE

                            Case SIS_ARG_POSITION

                                Loader.SIS.OpenList("lVertices", i)
                                Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_ox#", x)
                                Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_oy#", y)
                                Loader.SIS.UpdateItem()

                        End Select

                    Loop

                Next

                Loader.SIS.RecallNolView("view")

            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
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

    Private Sub EmptyList(ByVal List As String)

        Try

            Loader.SIS.EmptyList(List)

        Catch
        End Try

    End Sub

End Class