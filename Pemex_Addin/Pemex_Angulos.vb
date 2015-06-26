Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants

Public Class Pemex_Angulos

    Public Sub New()

    End Sub

    Public Sub Angulos()

        Try

            Dim x, y, z, xA, yA, xS, yS, xE, yE, angle, length As Double

            Dim lResponse As Integer

            'snap x
            Do

                lResponse = Loader.SIS.GetPosEx(x, y, z)
                Select Case lResponse

                    Case SIS_ARG_ENTER

                    Case SIS_ARG_ESCAPE

                        Exit Sub

                    Case SIS_ARG_BACKSPACE

                    Case SIS_ARG_POSITION

                        Try
                            Loader.SIS.SplitPos(xA, yA, z, (Loader.SIS.Snap2D(x, y, 3, True, "X", "fPemexAngle", "")))
                            Exit Do
                        Catch ex As Exception

                            MsgBox("No X Snap", MsgBoxStyle.Exclamation, "Angulos")

                        End Try

                End Select

            Loop

            'snap l1
            Do

                lResponse = Loader.SIS.GetPosEx(x, y, z)
                Select Case lResponse

                    Case SIS_ARG_ENTER

                    Case SIS_ARG_ESCAPE

                        Exit Sub

                    Case SIS_ARG_BACKSPACE

                    Case SIS_ARG_POSITION

                        Try

                            Loader.SIS.MoveTo(xA, yA, z)
                            Loader.SIS.SplitPos(xS, yS, z, (Loader.SIS.Snap2D(x, y, 3, True, "L", "fPemexAngle", "")))

                            Dim lengthS As Double = -1
                            Dim lengthA As Double = -1
                            Dim i As Integer = 0
                            Do While lengthA = -1
                                lengthS = Loader.SIS.GetGeomLengthUpto(i, 0, xS, yS, 0)
                                lengthA = Loader.SIS.GetGeomLengthUpto(i, 0, xA, yA, 0)
                                i += 1
                            Loop
                            Dim lengthmod As Double = -0.001
                            If lengthS > lengthA Then lengthmod = +0.001
                            Loader.SIS.SplitPos(xS, yS, z, (Loader.SIS.GetGeomPosFromLength(i - 1, lengthA + lengthmod)))

                            Loader.SIS.LineTo(xS, yS, z)
                            Loader.SIS.UpdateItem()
                            angle = Loader.SIS.GetGeomAngleFromLength(0, 0.0001)
                            Loader.SIS.DeleteItem()
                            xS = xA + 5 * Math.Cos(angle)
                            yS = yA + 5 * Math.Sin(angle)
                            Exit Do

                        Catch ex As Exception

                            MsgBox("No L Snap", MsgBoxStyle.Exclamation, "Angulos")

                        End Try

                End Select

            Loop

            'snap l2
            Do

                lResponse = Loader.SIS.GetPosEx(x, y, z)
                Select Case lResponse

                    Case SIS_ARG_ENTER

                    Case SIS_ARG_ESCAPE

                        Exit Sub

                    Case SIS_ARG_BACKSPACE

                    Case SIS_ARG_POSITION

                        Try

                            Loader.SIS.MoveTo(xA, yA, z)
                            Loader.SIS.SplitPos(xE, yE, z, (Loader.SIS.Snap2D(x, y, 1, True, "L", "fPemexAngle", "")))

                            Dim lengthE As Double = -1
                            Dim lengthA As Double = -1
                            Dim i As Integer = 0
                            Do While lengthA = -1
                                lengthE = Loader.SIS.GetGeomLengthUpto(i, 0, xE, yE, 0)
                                lengthA = Loader.SIS.GetGeomLengthUpto(i, 0, xA, yA, 0)
                                i += 1
                            Loop
                            Dim lengthmod As Double = -0.001
                            If lengthE > lengthA Then lengthmod = +0.001
                            Loader.SIS.SplitPos(xE, yE, z, (Loader.SIS.GetGeomPosFromLength(i - 1, lengthA + lengthmod)))

                            Loader.SIS.LineTo(xE, yE, z)
                            Loader.SIS.UpdateItem()
                            angle = Loader.SIS.GetGeomAngleFromLength(0, 0.0001)
                            Loader.SIS.DeleteItem()
                            xE = xA + 5 * Math.Cos(angle)
                            yE = yA + 5 * Math.Sin(angle)
                            Exit Do

                        Catch ex As Exception

                            MsgBox("No L Snap", MsgBoxStyle.Exclamation, "Angulos")

                        End Try

                End Select

            Loop

            Loader.SIS.MoveTo(xS, yS, 0)
            Loader.SIS.LineTo(xE, yE, 0)
            Loader.SIS.UpdateItem()
            length = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_length#")
            Loader.SIS.DeleteItem()
            angle = 2 * Math.Asin(length / 10)
            Loader.SIS.MoveTo(xS, yS, 0)
            Loader.SIS.BulgeTo(angle, xE, yE)
            Loader.SIS.UpdateItem()
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1002)
            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "angle#", angle * (180 / Math.PI))
            Loader.SIS.UpdateItem()

            'create text
            Do

                lResponse = Loader.SIS.GetPosEx(x, y, z)
                Select Case lResponse

                    Case SIS_ARG_ENTER

                    Case SIS_ARG_ESCAPE

                        Exit Sub

                    Case SIS_ARG_BACKSPACE

                    Case SIS_ARG_POSITION

                        Dim displayangle = Loader.SIS.GetFlt(SIS_OT_WINDOW, 0, "_displayAngle#")
                        Loader.SIS.CreateText(x, y, 0, DDtoDMS(angle * (180 / Math.PI)).replace(" ", ""))
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_point_height&", 6)
                        Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", displayangle * 180 / Math.PI)
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1002)
                        Loader.SIS.UpdateItem()
                        Loader.SIS.SelectItem()
                        Loader.SIS.DoCommand("AComTextToBox")
                        Exit Do

                End Select

            Loop

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Function DDtoDMS(ByVal Decimal_Deg As Double)

        Try

            Dim Degrees = Int(Decimal_Deg)
            Dim Minutes = (Decimal_Deg - Degrees) * 60
            Dim Seconds = Format(((Minutes - Int(Minutes)) * 60), "0")
            Return Str(Degrees) & "°" & Str(Int(Minutes)) & "'" & Str(Seconds) + Chr(34)

        Catch

            Return Nothing

        End Try

    End Function

    Private Sub EmptyList(ByVal List As String)

        Try

            Loader.SIS.EmptyList(List)

        Catch
        End Try

    End Sub

End Class
