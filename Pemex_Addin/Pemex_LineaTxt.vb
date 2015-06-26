Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants

Public Class Pemex_LineaTxt

    Public Sub New()

    End Sub

    Public Sub LineaTxt()

        Try

            Dim x, y, z, frameangle As Double

            Loader.SIS.DeselectAll()
            Loader.SIS.CreatePropertyFilter("fAngle", "_FC&=1100")

            If Loader.SIS.Scan("lAngle", "E", "fAngle", "") = 1 Then

                Loader.SIS.OpenList("lAngle", 0)
                frameangle = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_angleDeg#")

            End If

            Dim lResponse As Integer

            Do

                lResponse = Loader.SIS.GetPosEx(x, y, z)
                Select Case lResponse

                    Case SIS_ARG_ESCAPE

                        Exit Do

                    Case SIS_ARG_POSITION

                        Try

                            Loader.SIS.SplitPos(x, y, z, (Loader.SIS.Snap2D(x, y, 3, True, "L", "fPemexLines", ""))) 
                            Loader.SIS.GetGeomLengthUpto(0, 0, x, y, 0)
                            Dim lineangle = Loader.SIS.GetGeomAngleFromLength(0, Loader.SIS.GetGeomLengthUpto(0, 0, x, y, 0)) * 180 / Math.PI

                            If (lineangle - frameangle > 90 And lineangle - frameangle < 270) Or (lineangle - frameangle < -90 And lineangle - frameangle > -270) Then lineangle += 180

                            If Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "_FC&") = 1033 Then

                                Loader.SIS.CreateText(x, y, 0, "LÍMITE DE PROPIEDAD")

                            ElseIf Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "_FC&") = 1021 Then

                                Loader.SIS.CreateText(x, y, 0, "LDDV PEMEX")


                            ElseIf Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "_FC&") = 1023 Then

                                Loader.SIS.CreateText(x, y, 0, "LDDV CFE")

                            End If

                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_point_height&", 7)
                            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1055)
                            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", lineangle)
                            Loader.SIS.UpdateItem()
                            Loader.SIS.SelectItem()
                            Loader.SIS.DoCommand("AComTextToBox")

                        Catch ex As Exception

                            MsgBox("No L Snap", MsgBoxStyle.Exclamation, "LineaTxt")

                        End Try

                End Select

            Loop

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Sub KMITxt()

        Try

            Dim x, y, z, frameangle As Double
            Dim KMI = ""

            Loader.SIS.DeselectAll()
            Loader.SIS.CreatePropertyFilter("fAngle", "_FC&=1100")
            If Loader.SIS.Scan("lAngle", "E", "fAngle", "") = 1 Then

                Loader.SIS.OpenList("lAngle", 0)
                frameangle = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_angleDeg#")

            End If

            Try

                Loader.SIS.CreatePropertyFilter("fAfectacion", "_FC&=1004")
                If Loader.SIS.Scan("lAfectacion", "E", "fAfectacion", "") = 1 Then

                    Loader.SIS.OpenList("lAfectacion", 0)
                    KMI = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "KMI$")
                    If KMI = "" Then Exit Sub

                Else

                    Exit Sub

                End If

            Catch ex As Exception
                MsgBox("No KMI", MsgBoxStyle.Exclamation, "KMI TXT")
                Exit Sub
            End Try

            Dim lResponse As Integer

            Do

                lResponse = Loader.SIS.GetPosEx(x, y, z)
                Select Case lResponse

                    Case SIS_ARG_ESCAPE

                        Exit Do

                    Case SIS_ARG_POSITION

                        Try

                            Loader.SIS.SplitPos(x, y, z, (Loader.SIS.Snap2D(x, y, 3, True, "L", "", "")))
                            Loader.SIS.GetGeomLengthUpto(0, 0, x, y, 0)
                            Dim lineangle = Loader.SIS.GetGeomAngleFromLength(0, Loader.SIS.GetGeomLengthUpto(0, 0, x, y, 0)) * 180 / Math.PI

                            If (lineangle + frameangle > 90 And lineangle + frameangle < 270) Or (lineangle + frameangle < -90 And lineangle + frameangle > -270) Then lineangle += 180

                            Loader.SIS.CreateText(x, y, 0, KMI)
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_point_height&", 7)
                            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1055)
                            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", lineangle)
                            Loader.SIS.UpdateItem()
                            Loader.SIS.SelectItem()
                            Loader.SIS.DoCommand("AComTextToBox")
                            Exit Do

                        Catch ex As Exception

                            MsgBox("No L Snap", MsgBoxStyle.Exclamation, "KMI TXT")

                        End Try

                End Select

            Loop

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Sub KMFTxt()

        Try

            Dim x, y, z, frameangle As Double
            Dim KMF = ""

            Loader.SIS.DeselectAll()
            Loader.SIS.CreatePropertyFilter("fAngle", "_FC&=1100")
            If Loader.SIS.Scan("lAngle", "E", "fAngle", "") = 1 Then

                Loader.SIS.OpenList("lAngle", 0)
                frameangle = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_angleDeg#")

            End If

            Try

                Loader.SIS.CreatePropertyFilter("fAfectacion", "_FC&=1004")
                If Loader.SIS.Scan("lAfectacion", "E", "fAfectacion", "") = 1 Then

                    Loader.SIS.OpenList("lAfectacion", 0)
                    KMF = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "KMF$")
                    If KMF = "" Then Exit Sub

                Else

                    Exit Sub

                End If

            Catch ex As Exception
                MsgBox("No KMF", MsgBoxStyle.Exclamation, "KMF TXT")
                Exit Sub
            End Try

            Dim lResponse As Integer

            Do

                lResponse = Loader.SIS.GetPosEx(x, y, z)
                Select Case lResponse

                    Case SIS_ARG_ESCAPE

                        Exit Do

                    Case SIS_ARG_POSITION

                        Try

                            Loader.SIS.SplitPos(x, y, z, (Loader.SIS.Snap2D(x, y, 3, True, "L", "", "")))
                            Loader.SIS.GetGeomLengthUpto(0, 0, x, y, 0)
                            Dim lineangle = Loader.SIS.GetGeomAngleFromLength(0, Loader.SIS.GetGeomLengthUpto(0, 0, x, y, 0)) * 180 / Math.PI

                            If (lineangle + frameangle > 90 And lineangle + frameangle < 270) Or (lineangle + frameangle < -90 And lineangle + frameangle > -270) Then lineangle += 180

                            Loader.SIS.CreateText(x, y, 0, KMF)
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_point_height&", 7)
                            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1055)
                            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", lineangle)
                            Loader.SIS.UpdateItem()
                            Loader.SIS.SelectItem()
                            Loader.SIS.DoCommand("AComTextToBox")
                            Exit Do

                        Catch ex As Exception

                            MsgBox("No L Snap", MsgBoxStyle.Exclamation, "KMF TXT")

                        End Try

                End Select

            Loop

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

End Class
