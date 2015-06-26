Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants

Public Class Pemex_Plano

    Private iCurOverlay As Integer
    Private angle As Double
    Private secc As String
    Private reverse As Boolean
    Private bdt As Boolean

    Public Sub New()

    End Sub

    Public Sub Plano()

        Try

            EmptyList("lFrame")
            Dim x, y, z As Double
            Loader.SIS.CreateListFromSelection("lOleo")
            angle = GetAngle()
            Loader.SIS.SetListInt("lOleo", "reverse&", reverse)

            Loader.SIS.OpenList("lOleo", 0)
            SetCurrentOverlay()

            Loader.SIS.OpenList("lOleo", 0)
            Dim length = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_length#")
            Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_length#") / 2))

            ' check arroyo
            Dim arroyo = False
            Loader.SIS.CreatePropertyFilter("fAfectacion", "_FC&=1012")
            If Loader.SIS.ScanOverlay("lAfectacion", iCurOverlay, "fAfectacion", "") > 0 Then arroyo = True
            Dim poli = False
            Loader.SIS.CreatePropertyFilter("fAfectacion", "_FC&=1203")
            If Loader.SIS.ScanOverlay("lAfectacion", iCurOverlay, "fAfectacion", "") > 0 Then poli = True

            ' get afectacion
            Loader.SIS.CreatePropertyFilter("fAfectacion", "_FC&=1004")

            If Loader.SIS.ScanOverlay("lAfectacion", iCurOverlay, "fAfectacion", "") = 1 Then

                Loader.SIS.OpenList("lAfectacion", 0)
                secc = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "Secc$")
                bdt = True

                Try
                    length += Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "addlength#")

                Catch
                End Try

                If Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "AnchoDdv#") > 0 Then

                    bdt = False
                    If length > 660 Then

                        If poli = True Then
                            Loader.SIS.RecallNolItem("Pemex_Frame_XXXL_Poli")
                        Else
                            Loader.SIS.RecallNolItem("Pemex_Frame_XXXL")
                        End If

                    ElseIf length > 460 Then

                        If poli = True Then
                            Loader.SIS.RecallNolItem("Pemex_Frame_XXL_Poli")
                        Else
                            Loader.SIS.RecallNolItem("Pemex_Frame_XXL")
                        End If

                    ElseIf length > 240 Then

                        If poli = True Then
                            Loader.SIS.RecallNolItem("Pemex_Frame_XL_Poli")
                        Else
                            Loader.SIS.RecallNolItem("Pemex_Frame_XL")
                        End If

                    Else

                        If poli = True Then
                            Loader.SIS.RecallNolItem("Pemex_Frame_Poli")
                        Else
                            Loader.SIS.RecallNolItem("Pemex_Frame")
                        End If

                    End If

                Else

                    If length > 660 Then

                        If poli = True Then
                            Loader.SIS.RecallNolItem("Pemex_Frame_BDT_XXXL_Poli")
                        Else
                            Loader.SIS.RecallNolItem("Pemex_Frame_BDT_XXXL")
                        End If

                    ElseIf length > 460 Then

                        If poli = True Then
                            Loader.SIS.RecallNolItem("Pemex_Frame_BDT_XXL_Poli")
                        Else
                            Loader.SIS.RecallNolItem("Pemex_Frame_BDT_XXL")
                        End If

                    ElseIf length > 240 Then

                        If poli = True Then
                            Loader.SIS.RecallNolItem("Pemex_Frame_BDT_XL_Poli")
                        Else
                            Loader.SIS.RecallNolItem("Pemex_Frame_BDT_XL")
                        End If

                    Else

                        If poli = True Then
                            Loader.SIS.RecallNolItem("Pemex_Frame_BDT_POLI")
                        Else
                            Loader.SIS.RecallNolItem("Pemex_Frame_BDT")
                        End If

                    End If

                End If

                Loader.SIS.AddToList("lFrame")
                Loader.SIS.MoveList("lFrame", 0, 0, 0, angle * Math.PI / 180, 1)
                Loader.SIS.MoveList("lFrame", x, y, 0, 0, 1)
                Loader.SIS.SetFlt(SIS_OT_WINDOW, 0, "_displayAngle#", angle * Math.PI / 180)
                Loader.SIS.DeselectAll()
                Loader.SIS.SelectList("lFrame")
                Loader.SIS.DoCommand("AComExplodeGroup")
                Loader.SIS.CreateListFromSelection("lFrame")

                ' NOTAS
                Notas()

                ' arroyo
                If arroyo = False Then

                    Loader.SIS.CreatePropertyFilter("fIDtext", "IDtext$='arroyo'")
                    If Loader.SIS.ScanList("lIDtext", "lFrame", "fIDtext", "") > 0 Then Loader.SIS.Delete("lIDtext")

                End If

                ' DATOS PROPIETARIO
                Try
                    Loader.SIS.OpenList("lAfectacion", 0)
                    SetFrameText("Nombre", Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "Nombre$"))
                Catch
                End Try
                Try
                    Loader.SIS.OpenList("lAfectacion", 0)
                    SetFrameText("Regimen", Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "Regimen$"))
                Catch
                End Try
                Try
                    Loader.SIS.OpenList("lAfectacion", 0)
                    SetFrameText("NombrePropiedad", Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "NombrePropiedad$"))
                Catch
                End Try
                Try
                    Loader.SIS.OpenList("lAfectacion", 0)
                    SetFrameText("Municipio", Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "Municipio$"))
                Catch
                End Try
                Try
                    Loader.SIS.OpenList("lAfectacion", 0)
                    SetFrameText("Estado", Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "Estado$"))
                Catch
                End Try

                ' Lugar
                Try
                    Loader.SIS.OpenList("lAfectacion", 0)
                    SetFrameText("Lugar", Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "Estado$"))
                Catch
                End Try

                ' KM
                Try
                    Loader.SIS.OpenList("lAfectacion", 0)
                    SetFrameText("KM", "DEL " & Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "KMI$") & " AL " & Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "KMF$"))
                Catch
                End Try

                ' ID
                Try
                    Loader.SIS.OpenList("lAfectacion", 0)
                    If poli = True Then
                        SetFrameText("ID", "(B-010-PL-" & Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "CLAVE_CAT$").Substring(2, 2) & ")")
                    Else
                        SetFrameText("ID", "(B-010-OL-" & Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "CLAVE_CAT$").Substring(2, 2) & ")")
                    End If
                Catch
                End Try

                ' Clave
                Try
                    Loader.SIS.OpenList("lAfectacion", 0)
                    SetFrameText("Clave", "Q-" & Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "CLAVE_CAT$"))
                Catch
                End Try

                ' IDplan
                Try
                    Loader.SIS.OpenList("lAfectacion", 0)
                    If poli = True Then
                        SetFrameText("IDplan", "Q-" & Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "CLAVE_CAT$") & " (B-010-PL-" & Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "CLAVE_CAT$").Substring(2, 2) & ")-A-DWG")
                    Else
                        SetFrameText("IDplan", "Q-" & Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "CLAVE_CAT$") & " (B-010-OL-" & Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "CLAVE_CAT$").Substring(2, 2) & ")-A-DWG")
                    End If
                Catch
                End Try

                ' FechaHoy
                Try
                    Loader.SIS.OpenList("lAfectacion", 0)
                    SetFrameText("FechaHoy", DateTime.Now().ToString("dd-MMM-yyyy").ToUpper.Replace(".", ""))
                Catch
                End Try

                ' FechaRev
                Try
                    Loader.SIS.OpenList("lAfectacion", 0)
                    SetFrameText("FechaRev", DateTime.Now().ToString("dd/MMM/yy").ToUpper.Replace(".", ""))
                Catch
                End Try

                ' NUMreferencia
                Try
                    Loader.SIS.OpenList("lAfectacion", 0)
                    SetFrameText("NUMreferencia", "Q-" & Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "NUMreferencia$"))
                Catch
                End Try

                ' BDT
                Try
                    Loader.SIS.OpenList("lAfectacion", 0)
                    SetFrameText("BDT_DDV", String.Format("{0:###,###,###.000} m", Math.Round(Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "AnchoBdt#"), 4)))
                Catch
                End Try

                ' LONGITUD
                Try
                    Loader.SIS.OpenList("lAfectacion", 0)
                    SetFrameText("DDV_LONGITUD", String.Format("{0:###,###,##0.000} m", Math.Round(Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "Longitud#"), 4)))
                Catch
                End Try

                ' AREA BDT
                Try
                    Loader.SIS.OpenList("lAfectacion", 0)
                    SetFrameText("BDT_AREA", String.Format("{0:###,###,###.000} m²", Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "areaBDT#")))
                Catch
                End Try

                ' AREA BDT HA
                Try
                    Loader.SIS.OpenList("lAfectacion", 0)
                    SetFrameText("BDT_areaHA", String.Format("{0:00-00-00.000}" & " ha", Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "areaBDT#")))
                Catch
                End Try

                If bdt = False Then

                    'DDV
                    Try
                        Loader.SIS.OpenList("lAfectacion", 0)
                        SetFrameText("DDV_DDV", String.Format("{0:###,###,###.000} m", Math.Round(Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "AnchoDdv#"), 4)))
                    Catch
                    End Try
                    Try
                        Loader.SIS.OpenList("lAfectacion", 0)
                        SetFrameText("DDV_AREA", String.Format("{0:###,###,###.000} m²", Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_area#")))
                    Catch
                    End Try
                    Try
                        Loader.SIS.OpenList("lAfectacion", 0)
                        SetFrameText("DDV_areaHA", String.Format("{0:00-00-00.000}" & " ha", Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_area#")))
                    Catch
                    End Try

                End If

                SetNorte()

                Firma()

                Grid()

                Croquis()

                Toponomia()

                Loader.SIS.DeselectAll()
                Loader.SIS.OpenList("lAfectacion", 0)
                Loader.SIS.CreateBufferFromItems("lAfectacion", Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "Buffer&"), 0)
                Loader.SIS.SelectItem()
                Loader.SIS.DoCommand("AComBoundary")
                Loader.SIS.DoCommand("AComFillGeometry")
                Loader.SIS.OpenSel(0)
                Loader.SIS.SnipGeometry("lOleo", False)
                Loader.SIS.DeleteItem()

            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Function GetAngle()

        Try

            Dim xs, ys, xe, ye, z As Double
            Loader.SIS.SplitPos(xs, ys, z, Loader.SIS.GetGeomPt(0, 0))
            Loader.SIS.SplitPos(xe, ye, z, Loader.SIS.GetGeomPt(Loader.SIS.GetNumGeom() - 1, Loader.SIS.GetGeomNumPt(Loader.SIS.GetNumGeom() - 1) - 1))
            Loader.SIS.MoveTo(xs, ys, 0)
            Loader.SIS.LineTo(xe, ye, 0)
            Loader.SIS.UpdateItem()
            Dim _angle = Loader.SIS.GetGeomAngleFromLength(0, 0) * 180 / Math.PI
            Loader.SIS.DeleteItem()

            Select Case _angle

                Case -90 To 90

                    reverse = True
                    Return _angle

                Case 90 To 180

                    reverse = False
                    Return _angle - 180

                Case -180 To -90

                    reverse = False
                    Return _angle + 180

                Case Else

                    Return 0

            End Select

            Loader.SIS.DeleteItem()

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return 0
        End Try

    End Function

    Private Sub MultiLineGeometry()

        Try

            Dim x, y, z As Double
            Loader.SIS.OpenList("lOleo", 0)
            Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPt(0, 0))
            Loader.SIS.MoveTo(x, y, 0)
            For i = 0 To Loader.SIS.GetNumGeom() - 1

                For ii = 0 To Loader.SIS.GetGeomNumPt(i) - 1

                    Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPt(i, ii))
                    Loader.SIS.LineTo(x, y, 0)

                Next

            Next
            Loader.SIS.UpdateItem()
            Loader.SIS.EmptyList("lOleo")
            Loader.SIS.AddToList("lOleo")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub SetFrameText(ByVal IDtext As String, ByVal text As String)

        Try

            Loader.SIS.CreatePropertyFilter("fIDtext", "IDtext$='" & IDtext & "'")
            If Loader.SIS.ScanList("lIDtext", "lFrame", "fIDtext", "") = 1 Then

                Loader.SIS.OpenList("lIDtext", 0)
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", text)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1056)

            End If

        Catch
        End Try

    End Sub

    Private Sub Notas()

        Try

            Loader.SIS.DeselectAll()
            Loader.SIS.OpenList("lAfectacion", 0)
            Dim note = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "Note$")
            Dim x, y, z As Double
            Loader.SIS.CreatePropertyFilter("fNote", "_FC&=1058")

            If Loader.SIS.ScanList("lNote", "lFrame", "fNote", "") = 1 Then

                Loader.SIS.OpenList("lNote", 0)
                Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPt(0, 0))
                Loader.SIS.Delete("lNote")

                If bdt = True Then note = note & "_BDT"
                Loader.SIS.CreatePoint(x, y, 0, "NOTAS." & note, 0, 1)
                Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", angle)
                Loader.SIS.UpdateItem()
                Loader.SIS.SelectItem()
                Loader.SIS.DoCommand("AComExplodeShape")
                Loader.SIS.OpenSel(0)
                Loader.SIS.AddToList("lFrame")

            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub SetNorte()

        Try

            Dim x, y, z As Double
            Loader.SIS.CreatePropertyFilter("fNorte", "_FC&=1045")

            If Loader.SIS.ScanList("lNorte", "lFrame", "fNorte", "") = 1 Then
                Loader.SIS.OpenList("lNorte", 0)
                Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPt(0, 0))
                Loader.SIS.Delete("lNorte")

                If angle > 0 Then

                    Loader.SIS.CreatePoint(x, y, 0, "NORTE_R", 0, 1)

                Else

                    Loader.SIS.CreatePoint(x, y, 0, "NORTE", 0, 1)

                End If

                Loader.SIS.UpdateItem()

            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub Firma()

        Try

            Dim x, y, z As Double
            Loader.SIS.CreatePropertyFilter("fProperty", "_shape$='FIRMA_NOMBRE'")
            If Loader.SIS.ScanList("lFirma", "lFrame", "fProperty", "") = 1 Then
                Loader.SIS.OpenList("lFirma", 0)
                Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPt(0, 0))
                Loader.SIS.Delete("lFirma")
                Loader.SIS.OpenList("lAfectacion", 0)
                Loader.SIS.CreatePoint(x, y, z, Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "Firma$"), 0, 1)
                Loader.SIS.UpdateItem()
                Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", angle)
                Loader.SIS.UpdateItem()

            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub Grid()

        Try

            EmptyList("lMalla")
            Loader.SIS.CreatePropertyFilter("fMallaExtent", "_FC&=1101")
            If Loader.SIS.ScanList("lMallaExtent", "lFrame", "fMallaExtent", "") = 1 Then

                Loader.SIS.OpenList("lMallaExtent", 0)
                Dim x = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                Dim y = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
                Loader.SIS.CreateLocusFromItem("LocusExtent", SIS_GT_DISJOINT, SIS_GM_GEOMETRY)
                Loader.SIS.Delete("lMallaExtent")
                Dim xround = Math.Round(x / 1000, 1) * 1000
                Dim yround = Math.Round(y / 1000, 1) * 1000

                For xi = xround - 500 To xround + 500 Step 100

                    For yi = yround - 500 To yround + 500 Step 100

                        Loader.SIS.CreatePoint(xi, yi, 0, "MALLA", 0, 1)
                        Loader.SIS.UpdateItem()
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1041)
                        Loader.SIS.UpdateItem()
                        Loader.SIS.AddToList("lMalla")

                    Next

                Next

                If Loader.SIS.ScanList("lDelete", "lMalla", "", "LocusExtent") > 0 Then

                    Loader.SIS.Delete("lDelete")

                End If

                Loader.SIS.CreatePropertyFilter("fMalla", "_FC&=1041")
                If Loader.SIS.ScanOverlay("lMalla", iCurOverlay, "fMalla", "") > 0 Then

                    For i = 0 To Loader.SIS.GetListSize("lMalla") - 1

                        Loader.SIS.OpenList("lMalla", i)
                        x = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                        y = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
                        Loader.SIS.CreateText(x, y, 0, "     Y=" & String.Format("{0:#,###,###}", y))
                        Loader.SIS.UpdateItem()
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1041)
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_point_height&", 7)
                        Loader.SIS.UpdateItem()
                        Loader.SIS.SelectItem()
                        Loader.SIS.DoCommand("AComTextToBox")
                        Loader.SIS.CreateText(x, y, 0, "     X=" & String.Format("{0:#,###,###}", x))
                        Loader.SIS.UpdateItem()
                        If angle < 0 Then
                            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", 270)
                        Else
                            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", 90)
                        End If
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1041)
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_point_height&", 7)
                        Loader.SIS.UpdateItem()
                        Loader.SIS.SelectItem()
                        Loader.SIS.DoCommand("AComTextToBox")

                    Next

                End If

            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Sub Croquis()

        Try

            Loader.SIS.DeselectAll()
            Loader.SIS.OpenList("lAfectacion", 0)
            Dim ox = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
            Dim oy = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
            Dim KMI = Int(Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "KMI$").Substring(2, Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "KMI$").IndexOf("+") - 2))

            Dim xsFrame = 58
            Dim ysFrame = 36

            Loader.SIS.CreatePropertyFilter("fCroquisExtent", "_FC&=1102")
            If Loader.SIS.ScanList("lCroquisExtent", "lFrame", "fCroquisExtent", "") = 1 Then

                Loader.SIS.OpenList("lCroquisExtent", 0)
                Dim xoFrame = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                Dim yoFrame = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
                Loader.SIS.Delete("lCroquisExtent")

                Select Case KMI

                    Case 0 To 20

                        Loader.SIS.RecallNolItem("CROQUIS_000-020")

                    Case 20 To 35

                        Loader.SIS.RecallNolItem("CROQUIS_020-035")

                    Case 35 To 50

                        Loader.SIS.RecallNolItem("CROQUIS_035-050")

                    Case 50 To 65

                        Loader.SIS.RecallNolItem("CROQUIS_050-065")

                    Case 65 To 90

                        Loader.SIS.RecallNolItem("CROQUIS_065-090")

                    Case 300 To 314

                        Loader.SIS.RecallNolItem("CROQUIS_300-315")

                    Case 315 To 329

                        Loader.SIS.RecallNolItem("CROQUIS_315-330")

                    Case 330 To 344

                        Loader.SIS.RecallNolItem("CROQUIS_330-345")

                    Case 345 To 359

                        Loader.SIS.RecallNolItem("CROQUIS_345-360")

                    Case 360 To 374

                        Loader.SIS.RecallNolItem("CROQUIS_360-375")

                    Case 375 To 389

                        Loader.SIS.RecallNolItem("CROQUIS_375-390")

                    Case 390 To 404

                        Loader.SIS.RecallNolItem("CROQUIS_390-405")

                    Case 405 To 419

                        Loader.SIS.RecallNolItem("CROQUIS_405-420")

                    Case 420 To 434

                        Loader.SIS.RecallNolItem("CROQUIS_420-435")

                    Case 435 To 449

                        Loader.SIS.RecallNolItem("CROQUIS_435-450")

                    Case 450 To 464

                        Loader.SIS.RecallNolItem("CROQUIS_450-465")

                    Case 465 To 479


                        Loader.SIS.RecallNolItem("CROQUIS_465-480")

                    Case 480 To 494

                        Loader.SIS.RecallNolItem("CROQUIS_480-495")

                    Case 495 To 509

                        Loader.SIS.RecallNolItem("CROQUIS_495-510")

                    Case 510 To 524

                        Loader.SIS.RecallNolItem("CROQUIS_510-525")

                    Case 525 To 539

                        Loader.SIS.RecallNolItem("CROQUIS_525-540")

                    Case 540 To 554

                        Loader.SIS.RecallNolItem("CROQUIS_540-555")

                    Case 555 To 569

                        Loader.SIS.RecallNolItem("CROQUIS_555-570")

                    Case 570 To 584

                        Loader.SIS.RecallNolItem("CROQUIS_570-585")

                    Case 585 To 599

                        Loader.SIS.RecallNolItem("CROQUIS_585-600")

                    Case 600 To 614

                        Loader.SIS.RecallNolItem("CROQUIS_600-615")

                    Case 615 To 629

                        Loader.SIS.RecallNolItem("CROQUIS_615-630")

                    Case 630 To 660

                        Loader.SIS.RecallNolItem("CROQUIS_630-645")

                End Select

                Loader.SIS.DeselectAll()
                EmptyList("lCroquis")
                Loader.SIS.AddToList("lCroquis")
                Loader.SIS.SelectList("lCroquis")
                Loader.SIS.DoCommand("AComExplodeGroup")
                Loader.SIS.CreateListFromSelection("lCroquis")

                Dim x1, x2, y1, y2, z As Double

                Loader.SIS.CreatePropertyFilter("fCroquisLocation", "_shape$='CROQUIS_LOCATION'")
                If Loader.SIS.ScanList("lCroquisLocation", "lCroquis", "fCroquisLocation", "") = 1 Then
                    Loader.SIS.OpenList("lCroquisLocation", 0)
                    Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_ox#", ox)
                    Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_oy#", oy)
                    Loader.SIS.UpdateItem()

                End If

                Loader.SIS.SplitExtent(x1, y1, z, x2, y2, z, Loader.SIS.GetListExtent("lCroquis"))

                Dim xoCroquis = (x1 + x2) / 2
                Dim yoCroquis = (y1 + y2) / 2
                Dim xsCroquis = x2 - x1
                Dim ysCroquis = y2 - y1
                Dim scaleCroquis As Double

                If xsFrame / xsCroquis < ysFrame / ysCroquis Then

                    scaleCroquis = xsFrame / xsCroquis

                Else

                    scaleCroquis = ysFrame / ysCroquis

                End If

                Loader.SIS.MoveList("lCroquis", 0, 0, 0, angle * Math.PI / 180, scaleCroquis)
                Loader.SIS.SplitExtent(x1, y1, z, x2, y2, z, Loader.SIS.GetListExtent("lCroquis"))
                xoCroquis = (x1 + x2) / 2
                yoCroquis = (y1 + y2) / 2

                Dim xm = xoFrame - xoCroquis
                Dim ym = yoFrame - yoCroquis

                Loader.SIS.MoveList("lCroquis", xm, ym, 0, 0, 1)

                ' ClaveCroquis
                Try
                    Loader.SIS.OpenList("lAfectacion", 0)
                    Dim clave = "Q-" & Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "CLAVE_CAT$")
                    Loader.SIS.CreatePropertyFilter("fIDtext", "IDtext$='ClaveCroquis'")
                    If Loader.SIS.ScanList("lIDtext", "lCroquis", "fIDtext", "") = 1 Then

                        Loader.SIS.OpenList("lIDtext", 0)
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", clave)
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1056)

                    End If
                Catch
                End Try

                If Loader.SIS.ScanList("lCroquisExtent", "lCroquis", "fCroquisExtent", "") = 1 Then Loader.SIS.Delete("lCroquisExtent")

            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub Toponomia()

        Try

            Dim x, y, z, lineangle As Double

            Loader.SIS.OpenList("lOleo", 0)
            lineangle = Loader.SIS.GetGeomAngleFromLength(0, 1) * 180 / Math.PI
            Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, 1))

            If reverse = True Then
                Loader.SIS.CreatePoint(x, y, 0, "TOPONOMIA.DE_" & secc & "_R", 0, 1)
            Else
                Loader.SIS.CreatePoint(x, y, 0, "TOPONOMIA.DE_" & secc, 0, 1)
            End If

            Loader.SIS.UpdateItem()
            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", 180 + lineangle)
            Loader.SIS.UpdateItem()

            Loader.SIS.OpenList("lOleo", 0)
            lineangle = Loader.SIS.GetGeomAngleFromLength(0, Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_length#") - 1) * 180 / Math.PI
            Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_length#")))

            If reverse = True Then
                Loader.SIS.CreatePoint(x, y, 0, "TOPONOMIA.A_" & secc & "_R", 0, 1)
            Else
                Loader.SIS.CreatePoint(x, y, 0, "TOPONOMIA.A_" & secc, 0, 1)
            End If

            Loader.SIS.UpdateItem()
            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", 180 + lineangle)
            Loader.SIS.UpdateItem()

            Loader.SIS.OpenList("lOleo", 0)
            lineangle = Loader.SIS.GetGeomAngleFromLength(0, 5) * 180 / Math.PI
            If reverse = True Then
                Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, 10))
            Else
                Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, 5))
            End If
            Loader.SIS.CreatePoint(x, y, 0, "DISTANCIAS." & secc, 0, 1)
            Loader.SIS.UpdateItem()
            If reverse = True Then
                Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", lineangle)
            Else
                Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", 180 + lineangle)
            End If
            Loader.SIS.UpdateItem()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Sub Seccion()

        Try

            EmptyList("lExplode")
            Dim x, y, z As Double
            Loader.SIS.OpenSel(0)
            Dim reverse = Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "reverse&")
            Loader.SIS.CreateListFromSelection("lOleo")
            SetCurrentOverlay()

            MultiLineGeometry()

            Dim lineangle As Double = 0
            Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_length#") / 2))
            lineangle = Loader.SIS.GetGeomAngleFromLength(0, Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_length#") / 2) * 180 / Math.PI

            Loader.SIS.CreatePropertyFilter("fAfectacion", "_FC&=1004")

            If Loader.SIS.ScanOverlay("lAfectacion", iCurOverlay, "fAfectacion", "") = 1 Then

                Loader.SIS.OpenList("lAfectacion", 0)
                secc = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "Secc$")

                If reverse = True Then
                    Loader.SIS.CreatePoint(x, y, z, "LINEA SECCION A_R", 0, 1)
                Else
                    Loader.SIS.CreatePoint(x, y, z, "LINEA SECCION A", 0, 1)
                End If
                Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", 180 + lineangle)
                Loader.SIS.UpdateItem()

                Dim lResponse As Integer
                Do

                    lResponse = Loader.SIS.GetPosEx(x, y, z)
                    Select Case lResponse

                        Case SIS_ARG_POSITION

                            angle = Loader.SIS.GetFlt(SIS_OT_WINDOW, 0, "_displayAngle#")
                            Loader.SIS.CreatePoint(x, y, z, "COTAS." & secc, 0, 1)
                            Loader.SIS.UpdateItem()
                            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", angle * 180 / Math.PI)
                            Loader.SIS.UpdateItem()
                            Exit Do

                    End Select

                Loop

            End If

            Loader.SIS.Delete("lOleo")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Sub SetViewAngle()

        Try

            Loader.SIS.DeselectAll()

            Loader.SIS.CreatePropertyFilter("fIDtext", "IDtext$='Clave'")
            If Loader.SIS.Scan("lIDtext", "E", "fIDtext", "") = 1 Then

                Loader.SIS.OpenList("lIDtext", 0)
                Loader.SIS.SelectItem()
                Loader.SIS.DoCommand("AComBoxToText")
                Loader.SIS.SetFlt(SIS_OT_WINDOW, 0, "_displayAngle#", Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_angleDeg#") * Math.PI / 180)
                Loader.SIS.DoCommand("AComTextToBox")
                Loader.SIS.DeselectAll()

            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Sub ResetViewAngle()

        Try

            Loader.SIS.SetFlt(SIS_OT_WINDOW, 0, "_displayAngle#", 0)

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