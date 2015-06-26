Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants

Public Class Pemex_Seccion

    Public Sub New()

    End Sub

    Public Sub Secciones()

        Try

            EmptyList("lSeccion")
            EmptyList("lCota")
            Dim x, y, z As Double
            Dim scale = 2

            'measure intersections
            Loader.SIS.CreateListFromSelection("lOleo")
            Loader.SIS.OpenList("lOleo", 0)
            Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_length#") / 2))
            Dim lineangle = Loader.SIS.GetGeomAngleFromLength(0, Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_length#") / 2) * (180 / Math.PI)
            Dim reverse = Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "reverse&")

            Loader.SIS.MoveTo(x + 50 * Math.Cos((lineangle + 90) * (Math.PI / 180)), y + 50 * Math.Sin((lineangle + 90) * (Math.PI / 180)), 0)
            Loader.SIS.LineTo(x + 50 * Math.Cos((lineangle + 270) * (Math.PI / 180)), y + 50 * Math.Sin((lineangle + 270) * (Math.PI / 180)), 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lCota")
            If Loader.SIS.ScanGeometry("lIntersect", SIS_GT_INTERSECT, SIS_GM_GEOMETRY, "fPemexCotaLines", "") > 2 Then

                Dim aLineType(Loader.SIS.GetListSize("lIntersect") - 1) As String
                Dim aDistance(Loader.SIS.GetListSize("lIntersect") - 1) As Double

                Dim DistMeasure() As String = Split(Loader.SIS.GetGeomIntersections(0, "lIntersect"), ",")

                'get distances and line types
                If Not DistMeasure(0) = -1 Then

                    For i = 0 To DistMeasure.Length - 1

                        Loader.SIS.OpenList("lCota", 0)
                        Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, CDbl(DistMeasure(i))))
                        Loader.SIS.Snap2D(x, y, 1, False, "L", "fPemexCotaLines", "")
                        aLineType(i) = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "LineType$")
                        aDistance(i) = CDbl(DistMeasure(i)) - CDbl(DistMeasure(0))

                    Next

                End If

                Dim TotalWidth As Double = CDbl(DistMeasure(DistMeasure.Length - 1)) - CDbl(DistMeasure(0))

                'draw first arrows, line and text
                Loader.SIS.CreatePoint(0, 0, 0, "izqARROW", 0, 0.5)
                Explode()
                Loader.SIS.AddToList("lSeccion")
                Loader.SIS.CreatePoint(TotalWidth * scale, 0, 0, "derARROW", 0, 0.5)
                Explode()
                Loader.SIS.AddToList("lSeccion")

                Loader.SIS.MoveTo(0, 0, 0)
                Loader.SIS.LineTo(TotalWidth * scale, 0, 0)
                Loader.SIS.UpdateItem()
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                Loader.SIS.UpdateItem()
                Loader.SIS.AddToList("lSeccion")

                Loader.SIS.CreateText((TotalWidth / 2) * scale, 0, 0, String.Format("{0:###.00}", TotalWidth))
                Loader.SIS.UpdateItem()
                Loader.SIS.SelectItem()
                Loader.SIS.DoCommand("AComTextToBox")
                Loader.SIS.OpenSel(0)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_character_height#", 1.84)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_opaque&", True)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_level&", 1)
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Solid</Style><Colour><RGBA>255 255 255 0</RGBA></Colour><BackgroundColour><RGBA>255 255 0 255</RGBA></BackgroundColour></Brush>")
                Loader.SIS.UpdateItem()
                Loader.SIS.AddToList("lSeccion")

                For i = 0 To aDistance.Length - 2

                    'draw arrows
                    Loader.SIS.CreatePoint(aDistance(i) * scale, -5.5, 0, "izqARROW", 0, 0.5)
                    Explode()
                    Loader.SIS.AddToList("lSeccion")
                    Loader.SIS.CreatePoint(aDistance(i + 1) * scale, -5.5, 0, "derARROW", 0, 0.5)
                    Explode()
                    Loader.SIS.AddToList("lSeccion")

                    'draw line
                    Loader.SIS.MoveTo(aDistance(i) * scale, -5.5, 0)
                    Loader.SIS.LineTo(aDistance(i + 1) * scale, -5.5, 0)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                    Loader.SIS.UpdateItem()
                    Loader.SIS.AddToList("lSeccion")

                    'draw line symbol
                    Loader.SIS.CreatePoint(aDistance(i) * scale, -5.5, 0, aLineType(i), 0, 1)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.SelectItem()
                    Loader.SIS.DoCommand("AComExplodeShape")
                    Loader.SIS.CreateListFromSelection("List")
                    Loader.SIS.CombineLists("lSeccion", "lSeccion", "List", SIS_BOOLEAN_OR)

                    'calc distance
                    Dim distance = aDistance(i + 1) - aDistance(i)

                    'draw distance textbox
                    Loader.SIS.CreateText((aDistance(i) + (distance / 2)) * scale, -5.5, 0, String.Format("{0:###.00}", distance))
                    Loader.SIS.UpdateItem()
                    Loader.SIS.SelectItem()
                    Loader.SIS.DoCommand("AComTextToBox")
                    Loader.SIS.OpenSel(0)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                    Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_character_height#", 1.84)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_opaque&", True)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_level&", 1)
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Solid</Style><Colour><RGBA>255 255 255 0</RGBA></Colour><BackgroundColour><RGBA>255 255 0 255</RGBA></BackgroundColour></Brush>")
                    Loader.SIS.UpdateItem()
                    Loader.SIS.AddToList("lSeccion")

                Next

                ' left-most line symbol
                Loader.SIS.CreatePoint(aDistance(aDistance.Length - 1) * scale, -5.5, 0, aLineType(aLineType.Length - 1), 0, 1)
                Loader.SIS.UpdateItem()
                Loader.SIS.SelectItem()
                Loader.SIS.DoCommand("AComExplodeShape")
                Loader.SIS.CreateListFromSelection("List")
                Loader.SIS.CombineLists("lSeccion", "lSeccion", "List", SIS_BOOLEAN_OR)

                'drop lines
                Loader.SIS.MoveTo(0, 0.25, 0)
                Loader.SIS.LineTo(0, -5.75, 0)
                Loader.SIS.UpdateItem()
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                Loader.SIS.UpdateItem()
                Loader.SIS.AddToList("lSeccion")
                Loader.SIS.MoveTo(TotalWidth * scale, 0.25, 0)
                Loader.SIS.LineTo(TotalWidth * scale, -5.75, 0)
                Loader.SIS.UpdateItem()
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                Loader.SIS.UpdateItem()
                Loader.SIS.AddToList("lSeccion")

                'soil divider line
                Loader.SIS.MoveTo(0, -27, 0)
                Loader.SIS.LineTo(TotalWidth * scale, -27, 0)
                Loader.SIS.UpdateItem()
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                Loader.SIS.UpdateItem()
                Loader.SIS.AddToList("lSeccion")

                'soil pattern
                Loader.SIS.CreatePoint(0, -27, 0, "SOIL_PATTERN", 0, 0)
                Loader.SIS.UpdateItem()
                Loader.SIS.SelectItem()
                Loader.SIS.DoCommand("AComExplodeShape")
                Loader.SIS.CreateListFromSelection("List")
                Loader.SIS.CreateRectangle(0, -27, TotalWidth * 2, -29)
                Loader.SIS.UpdateItem()
                Loader.SIS.SnipGeometry("List", 0)
                Loader.SIS.DeleteItem()
                Loader.SIS.CombineLists("lSeccion", "lSeccion", "List", SIS_BOOLEAN_OR)

                'place group
                Dim lResponse As Integer
                Do

                    lResponse = Loader.SIS.GetPosEx(x, y, z)
                    Select Case lResponse

                        Case SIS_ARG_ESCAPE

                            Loader.SIS.Delete("lSeccion")
                            Exit Do

                        Case SIS_ARG_POSITION

                            Dim angle = Loader.SIS.GetFlt(SIS_OT_WINDOW, 0, "_displayAngle#")
                            Loader.SIS.MoveList("lSeccion", 0, 0, 0, angle, 1)
                            Loader.SIS.MoveList("lSeccion", x, y, z, 0, 1)
                            Loader.SIS.CreateGroupFromItems("lSeccion", True, "gSeccion")

                            Exit Do

                    End Select

                Loop

            End If

            Loader.SIS.Delete("lCota")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Sub Distancias()

        Try

            Dim x, y, z, xsnap, ysnap As Double
            Dim lResponse As Integer
            EmptyList("lTmp")

            Do

                lResponse = Loader.SIS.GetPosEx(x, y, z)
                Select Case lResponse

                    Case SIS_ARG_ESCAPE

                        Exit Sub

                    Case SIS_ARG_POSITION

                        Try
                            Loader.SIS.SplitPos(xsnap, ysnap, z, Loader.SIS.Snap2D(x, y, 3, True, "L", "", ""))
                            Exit Do
                        Catch ex As Exception

                            MsgBox("No L Snap", MsgBoxStyle.Exclamation, "Distancias")

                        End Try

                End Select

            Loop

            Dim lineangle = Loader.SIS.GetGeomAngleFromLength(0, Loader.SIS.GetGeomLengthUpto(0, 0, xsnap, ysnap, z)) * (180 / Math.PI)

            'second line
            Dim x2, y2, z2 As Double
            Dim lineangle2 = Loader.SIS.GetGeomAngleFromLength(0, Loader.SIS.GetGeomLengthUpto(0, 0, xsnap, ysnap, z) + 5) * (180 / Math.PI)
            Loader.SIS.SplitPos(x2, y2, z2, Loader.SIS.GetGeomPosFromLength(0, Loader.SIS.GetGeomLengthUpto(0, 0, xsnap, ysnap, z) + 5))

            Loader.SIS.MoveTo(xsnap + 50 * Math.Cos((lineangle + 90) * (Math.PI / 180)), ysnap + 50 * Math.Sin((lineangle + 90) * (Math.PI / 180)), 0)
            Loader.SIS.LineTo(xsnap + 50 * Math.Cos((lineangle + 270) * (Math.PI / 180)), ysnap + 50 * Math.Sin((lineangle + 270) * (Math.PI / 180)), 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTmp")

            'second line
            Loader.SIS.MoveTo(x2 + 50 * Math.Cos((lineangle + 90) * (Math.PI / 180)), y2 + 50 * Math.Sin((lineangle + 90) * (Math.PI / 180)), 0)
            Loader.SIS.LineTo(x2 + 50 * Math.Cos((lineangle + 270) * (Math.PI / 180)), y2 + 50 * Math.Sin((lineangle + 270) * (Math.PI / 180)), 0)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lTmp")

            Loader.SIS.OpenList("lTmp", 0)
            If Loader.SIS.ScanGeometry("lIntersect", SIS_GT_INTERSECT, SIS_GM_GEOMETRY, "fPemexCotaLines", "") > 1 Then

                Dim DistMeasure() As String = Split(Loader.SIS.GetGeomIntersections(0, "lIntersect"), ",")

                For i = 0 To DistMeasure.Length - 2

                    Loader.SIS.OpenList("lTmp", 0)
                    Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, CDbl(DistMeasure(i))))
                    Loader.SIS.CreatePoint(x, y, z, "izqARROW", 0, 0.5)
                    Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", lineangle - 90)
                    Explode()
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 255 0 0</RGBA></Colour><RoundCaps>true</RoundCaps></Pen>")
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Solid</Style><Colour><RGBA>0 255 0 0</RGBA></Colour></Brush>")
                    Loader.SIS.UpdateItem()

                    Loader.SIS.OpenList("lTmp", 0)
                    Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, CDbl(DistMeasure(i + 1))))
                    Loader.SIS.CreatePoint(x, y, z, "derARROW", 0, 0.5)
                    Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", lineangle - 90)
                    Explode()
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 255 0 0</RGBA></Colour><RoundCaps>true</RoundCaps></Pen>")
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Solid</Style><Colour><RGBA>0 255 0 0</RGBA></Colour></Brush>")
                    Loader.SIS.UpdateItem()

                    Dim distance = DistMeasure(i + 1) - DistMeasure(i)

                    Loader.SIS.OpenList("lTmp", 0)
                    Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, CDbl(DistMeasure(i) + (distance / 2))))
                    Loader.SIS.CreateText(x, y, z, String.Format("{0:###.00}", distance))
                    Loader.SIS.UpdateItem()
                    Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", lineangle + 90)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.SelectItem()
                    Loader.SIS.DoCommand("AComTextToBox")
                    Loader.SIS.OpenSel(0)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_BOTTOM)
                    Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_character_height#", 1.76)
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 255 0 0</RGBA></Colour><RoundCaps>true</RoundCaps></Pen>")
                    Loader.SIS.UpdateItem()

                Next

                Loader.SIS.OpenList("lTmp", 0)
                Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, CDbl(DistMeasure(0))))
                Loader.SIS.MoveTo(x, y, z)
                Loader.SIS.OpenList("lTmp", 0)
                Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, CDbl(DistMeasure(DistMeasure.Length - 1))))
                Loader.SIS.LineTo(x, y, z)
                Loader.SIS.UpdateItem()
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 255 0 0</RGBA></Colour><RoundCaps>true</RoundCaps></Pen>")
                Loader.SIS.UpdateItem()

                'second line
                Loader.SIS.OpenList("lTmp", 1)
                If Loader.SIS.ScanGeometry("lIntersect", SIS_GT_INTERSECT, SIS_GM_GEOMETRY, "fPemexCotaLines", "") > 1 Then

                    Dim DistSecond() As String = Split(Loader.SIS.GetGeomIntersections(0, "lIntersect"), ",")

                    Loader.SIS.OpenList("lTmp", 1)
                    Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, CDbl(DistSecond(0))))
                    Loader.SIS.CreatePoint(x, y, z, "izqARROW", 0, 0.5)
                    Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", lineangle - 90)
                    Explode()
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 255 0 0</RGBA></Colour><RoundCaps>true</RoundCaps></Pen>")
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Solid</Style><Colour><RGBA>0 255 0 0</RGBA></Colour></Brush>")
                    Loader.SIS.UpdateItem()

                    Loader.SIS.OpenList("lTmp", 1)
                    Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, CDbl(DistSecond(DistSecond.Length - 1))))
                    Loader.SIS.CreatePoint(x, y, z, "derARROW", 0, 0.5)
                    Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", lineangle - 90)
                    Explode()
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 255 0 0</RGBA></Colour><RoundCaps>true</RoundCaps></Pen>")
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Solid</Style><Colour><RGBA>0 255 0 0</RGBA></Colour></Brush>")
                    Loader.SIS.UpdateItem()

                    Dim distance = DistSecond(DistSecond.Length - 1) - DistSecond(0)

                    Loader.SIS.OpenList("lTmp", 1)
                    Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, CDbl(DistSecond(0) + (distance / 2))))
                    Loader.SIS.CreateText(x, y, z, String.Format("{0:###.00}", distance))
                    Loader.SIS.UpdateItem()
                    Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", lineangle + 90)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.SelectItem()
                    Loader.SIS.DoCommand("AComTextToBox")
                    Loader.SIS.OpenSel(0)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_BOTTOM)
                    Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_character_height#", 1.76)
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 255 0 0</RGBA></Colour><RoundCaps>true</RoundCaps></Pen>")
                    Loader.SIS.UpdateItem()

                    Loader.SIS.OpenList("lTmp", 1)
                    Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, CDbl(DistSecond(0))))
                    Loader.SIS.MoveTo(x, y, z)
                    Loader.SIS.OpenList("lTmp", 1)
                    Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, CDbl(DistSecond(DistMeasure.Length - 1))))
                    Loader.SIS.LineTo(x, y, z)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 255 0 0</RGBA></Colour><RoundCaps>true</RoundCaps></Pen>")
                    Loader.SIS.UpdateItem()

                End If

            End If

            Loader.SIS.Delete("lTmp")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub Explode()
        Try
            Loader.SIS.UpdateItem()
            Loader.SIS.SelectItem()
            Loader.SIS.DoCommand("AComExplodeShape")
            Loader.SIS.OpenSel(0)
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
