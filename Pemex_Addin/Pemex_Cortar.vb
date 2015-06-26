Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants

Public Class Pemex_Cortar

    Private sClave As String
    Private iCurOverlay As Integer

    Public Sub New()

    End Sub

    Public Sub Cortar()

        Try

            Dim Pla As Pemex_Plano = New Pemex_Plano
            Loader.SIS.CreatePropertyFilter("fOleo", "_FC&=1202")
            Loader.SIS.CreateListFromSelection("lSelection")
            Loader.SIS.DeselectAll()

            For i = 0 To Loader.SIS.GetListSize("lSelection") - 1

                Loader.SIS.OpenList("lSelection", i)
                If Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "_FC&") = 1004 Then

                    Loader.SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", False)
                    SwitchOffOverlays()
                    sClave = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "CLAVE_CAT$")
                    iCurOverlay = Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&")
                    Loader.SIS.CreateInternalOverlay(sClave, iCurOverlay)
                    Loader.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", iCurOverlay)
                    Loader.SIS.SetFlt(SIS_OT_DATASET, Loader.SIS.GetInt(SIS_OT_OVERLAY, iCurOverlay, "_nDataset&"), "_scale#", 1000)

                    MeasureKM()
                    Loader.SIS.OpenList("lSelection", i)
                    CopyTopo()
                    Loader.SIS.OpenList("lSelection", i)
                    GetBDT()
                    'CopyColindantes()

                    Loader.SIS.DeselectAll()

                    If Loader.SIS.ScanOverlay("lOleo", iCurOverlay, "fOleo", "") = 1 Then

                        Loader.SIS.OpenList("lOleo", 0)
                        Loader.SIS.SelectItem()
                        Pla.Plano()
                        Loader.SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", True)

                    End If

                End If

            Next

        Catch ex As Exception
            Loader.SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", True)
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub SwitchOffOverlays()

        Try

            For i As Integer = 0 To (Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)

                Loader.SIS.SetInt(SIS_OT_OVERLAY, i, "_status&", 0)

            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub MeasureKM()

        Try

            EmptyList("lAfectacion")
            Loader.SIS.AddToList("lAfectacion")
            'Dim AddMeasure As Integer = Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "AddMeasure&")
            Dim AddMeasure As Integer = 0
            Loader.SIS.CreateLocusFromItem("LocusCrossBy", SIS_GT_CROSSBY, SIS_GM_GEOMETRY)
            EmptyList("lTmp")
            Loader.SIS.AddToList("lTmp")
            Loader.SIS.CopyListItems("lTmp")
            Loader.SIS.OpenList("lTmp", 0)
            Loader.SIS.DeselectAll()
            Loader.SIS.SelectItem()
            Loader.SIS.DoCommand("AComBoundary")
            Loader.SIS.CreateListFromSelection("lBoundary")

            ' get MEASURE overlay
            For i As Integer = 0 To (Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)

                If Loader.SIS.GetStr(SIS_OT_OVERLAY, i, "_name$") = "MEASURE" Then

                    Loader.SIS.CreateListFromOverlay(i, "lMeasure")
                    Exit For

                End If

            Next

            ' get DDV overlay
            For i As Integer = 0 To (Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)

                If Loader.SIS.GetStr(SIS_OT_OVERLAY, i, "_name$") = "DDV" Then

                    Loader.SIS.CreateListFromOverlay(i, "lDDV")
                    Exit For

                End If

            Next

            ' measure first and last intersection
            If Loader.SIS.ScanList("lMeasure1", "lMeasure", "", "LocusCrossBy") > 0 Then

                Loader.SIS.OpenList("lMeasure1", 0)
                Dim Dist() As String = Split(Loader.SIS.GetGeomIntersections(0, "lBoundary"), ",")
                Loader.SIS.Delete("lTmp")

                Dim KMI_measure = CDbl(Dist(0))
                Dim KMI_KM = Int(KMI_measure / 1000) + AddMeasure
                Dim KMI_M = String.Format("{0:000.000}", Math.Round(KMI_measure - (Int(KMI_measure / 1000) * 1000), 3))
                Dim KMI As String = ("K-" + Str(KMI_KM) + "+" + KMI_M).ToString.Replace(" ", "")

                Dim KMF_measure = CDbl(Dist(Dist.Length - 1))
                Dim KMF_KM = Int(KMF_measure / 1000) + AddMeasure
                Dim KMF_M = String.Format("{0:000.000}", Math.Round(KMF_measure - (Int(KMF_measure / 1000) * 1000), 3))
                Dim KMF As String = ("K-" + Str(KMF_KM) + "+" + KMF_M).ToString.Replace(" ", "")

                Loader.SIS.CopyListItems("lAfectacion")
                Loader.SIS.OpenList("lAfectacion", 0)
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "KMI$", KMI)
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "KMF$", KMF)

                Dim Longitud As Double = 0
                For i = 1 To Dist.Length - 1 Step 2

                    Longitud += CDbl(Dist(i)) - CDbl(Dist(i - 1))

                Next
                Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "Longitud#", Longitud)

            Else

                EmptyList("lLines")
                Dim x, y, z, lineangle As Double
                Loader.SIS.OpenList("lAfectacion", 0)
                For i = 0 To Loader.SIS.GetGeomNumPt(0) - 1

                    Loader.SIS.OpenList("lAfectacion", 0)
                    Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPt(0, i))
                    Loader.SIS.CreatePoint(x, y, z, "_HATCH", 0, 1)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.SelectItem()
                    Loader.SIS.DoCommand("AComExplodeShape")
                    Loader.SIS.OpenSel(0)
                    Loader.SIS.AddToList("lPoints")

                Next

                Loader.SIS.OpenList("lAfectacion", 0)
                Loader.SIS.CreateLocusFromItem("LOCUS", SIS_GT_TOUCH, SIS_GM_GEOMETRY)
                Loader.SIS.ScanList("lDDVtouch", "lDDV", "", "LOCUS")
                Loader.SIS.CopyListItems("lDDVtouch")
                Loader.SIS.OpenList("lDDVtouch", 0)
                Dim DistDDV() As String = Split(Loader.SIS.GetGeomIntersections(0, "lPoints"), ",")
                Loader.SIS.Delete("lPoints")

                If Not DistDDV(0) = -1 Then

                    Loader.SIS.OpenList("lDDVtouch", 0)
                    Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, CDbl(DistDDV(0))))
                    lineangle = Loader.SIS.GetGeomAngleFromLength(0, CDbl(DistDDV(0))) * (180 / Math.PI)
                    Loader.SIS.MoveTo(x + 50 * Math.Cos((lineangle + 90) * (Math.PI / 180)), y + 50 * Math.Sin((lineangle + 90) * (Math.PI / 180)), 0)
                    Loader.SIS.LineTo(x + 50 * Math.Cos((lineangle + 270) * (Math.PI / 180)), y + 50 * Math.Sin((lineangle + 270) * (Math.PI / 180)), 0)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.AddToList("lLines")

                    Loader.SIS.OpenList("lDDVtouch", 0)
                    Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, CDbl(DistDDV(DistDDV.Length - 1))))
                    lineangle = Loader.SIS.GetGeomAngleFromLength(0, CDbl(DistDDV(DistDDV.Length - 1))) * (180 / Math.PI)
                    Loader.SIS.MoveTo(x + 50 * Math.Cos((lineangle + 90) * (Math.PI / 180)), y + 50 * Math.Sin((lineangle + 90) * (Math.PI / 180)), 0)
                    Loader.SIS.LineTo(x + 50 * Math.Cos((lineangle + 270) * (Math.PI / 180)), y + 50 * Math.Sin((lineangle + 270) * (Math.PI / 180)), 0)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.CreateLocusFromItem("LocusCrossBy", SIS_GT_CROSSBY, SIS_GM_GEOMETRY)
                    Loader.SIS.AddToList("lLines")

                End If

                Loader.SIS.Delete("lDDVtouch")

                Loader.SIS.ScanList("lMeasure", "lMeasure", "", "LocusCrossBy")

                Loader.SIS.OpenList("lMeasure", 0)
                Dim DistMeasure() As String = Split(Loader.SIS.GetGeomIntersections(0, "lLines"), ",")
                Loader.SIS.Delete("lLines")
                If Not DistMeasure(0) = -1 Then

                    Dim KMI_measure = CDbl(DistMeasure(0))
                    Dim KMI_KM = Int(KMI_measure / 1000) + AddMeasure
                    Dim KMI_M = String.Format("{0:000.000}", Math.Round(KMI_measure - (Int(KMI_measure / 1000) * 1000), 3))
                    Dim KMI As String = ("K-" + Str(KMI_KM) + "+" + KMI_M).ToString.Replace(" ", "")

                    Dim KMF_measure = CDbl(DistMeasure(DistMeasure.Length - 1))
                    Dim KMF_KM = Int(KMF_measure / 1000) + AddMeasure
                    Dim KMF_M = String.Format("{0:000.000}", Math.Round(KMF_measure - (Int(KMF_measure / 1000) * 1000), 3))
                    Dim KMF As String = ("K-" + Str(KMF_KM) + "+" + KMF_M).ToString.Replace(" ", "")

                    Loader.SIS.CopyListItems("lAfectacion")
                    Loader.SIS.OpenList("lAfectacion", 0)
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "KMI$", KMI)
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "KMF$", KMF)
                    Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "Longitud#", 0)

                End If

            End If

            Try
                Loader.SIS.Delete("lBoundary")
            Catch
            End Try

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub CopyTopo()

        Try

            EmptyList("lTmp")
            Loader.SIS.AddToList("lTmp")
            Loader.SIS.CreateBufferFromItems("lTmp", Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "Buffer&"), 0)
            Loader.SIS.CreateLocusFromItem("LOCUSBUFFER", SIS_GT_INTERSECT, SIS_GM_GEOMETRY)
            Loader.SIS.DeleteItem()

            ' get overlays
            For i As Integer = 0 To (Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)

                If Loader.SIS.GetStr(SIS_OT_OVERLAY, i, "_name$") = "LIMITES DE PROPIEDAD" _
                    Or Loader.SIS.GetStr(SIS_OT_OVERLAY, i, "_name$") = "DDV" _
                    Or Loader.SIS.GetStr(SIS_OT_OVERLAY, i, "_name$") = "TOPOGRAFIA" _
                    Or Loader.SIS.GetStr(SIS_OT_OVERLAY, i, "_name$") = "DUCTOS" Then

                    If Loader.SIS.ScanOverlay("lScan", i, "", "LOCUSBUFFER") > 0 Then

                        Loader.SIS.CopyListItems("lScan")

                    End If

                End If

            Next

            Loader.SIS.CreateListFromOverlay(iCurOverlay, "lSnip")
            If Loader.SIS.ScanList("lremove", "lSnip", "fOleo", "") = 1 Then

                Loader.SIS.CreateItemFromLocus("LOCUSBUFFER")
                Loader.SIS.SelectItem()
                Loader.SIS.DoCommand("AComBoundary")
                Loader.SIS.DoCommand("AComFillGeometry")
                Loader.SIS.OpenSel(0)
                EmptyList("lLOCUSBUFFER")
                Loader.SIS.AddToList("lLOCUSBUFFER")

                Loader.SIS.OpenList("lremove", 0)
                Dim Dist() As String = Split(Loader.SIS.GetGeomIntersections(0, "lLOCUSBUFFER"), ",")
                Loader.SIS.TraceGeom(0, CDbl(Dist(0)), CDbl(Dist(Dist.Length - 1)), 0, False)
                Loader.SIS.UpdateItem()
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "PEMEX")
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1202)
                Loader.SIS.UpdateItem()
                Loader.SIS.Delete("lremove")

                Loader.SIS.OpenList("lLOCUSBUFFER", 0)
                Loader.SIS.SnipGeometry("lSnip", False)
                Loader.SIS.DeleteItem()

            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub CopyColindantes()

        Try

            Loader.SIS.OpenList("lAfectacion", 0)
            EmptyList("lTmp")
            Loader.SIS.AddToList("lTmp")

            Loader.SIS.CreateLocusFromItem("LOCUSTOUCH", SIS_GT_TOUCH, SIS_GM_GEOMETRY)
            For i As Integer = 0 To (Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)

                If Loader.SIS.GetStr(SIS_OT_OVERLAY, i, "_name$") = "AFECTACIONES" Then

                    If Loader.SIS.ScanOverlay("List", i, "", "LOCUSTOUCH") > 0 Then

                        Loader.SIS.CopyListItems("List")

                        'replace affectacion with possible bdt area
                        Loader.SIS.OpenList("lTmp", 0)
                        Loader.SIS.CreateLocusFromItem("LOCUSWITHIN", SIS_GT_WITHIN, SIS_GM_GEOMETRY)
                        If Loader.SIS.ScanList("lDelete", "List", "", "LOCUSWITHIN") > 0 Then Loader.SIS.Delete("lDelete")
                        Loader.SIS.CopyListItems("lTmp")
                        Loader.SIS.CombineLists("List", "List", "lTmp", SIS_BOOLEAN_OR)

                        Loader.SIS.SetListInt("List", "_FC&", 1005)
                        Loader.SIS.CreateItemFromLocus("LOCUSBUFFER")
                        Loader.SIS.SelectItem()
                        Loader.SIS.DoCommand("AComBoundary")
                        Loader.SIS.DoCommand("AComFillGeometry")
                        Loader.SIS.OpenSel(0)
                        Loader.SIS.SnipGeometry("List", 0)
                        Loader.SIS.DeleteItem()
                        Loader.SIS.OpenList("lTmp", 0)
                        Dim nombre = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "Nombre$")
                        Loader.SIS.CreateBufferLocusFromItems("lTmp", False, "LOCUSBUFFER", 1, 0)
                        Loader.SIS.AddToList("List")
                        Loader.SIS.CopyListItems("List")

                        ' union items to be taken off / unfill fill to répair boolean geometry
                        Loader.SIS.CreateBoolean("List", SIS_BOOLEAN_OR)
                        Loader.SIS.SelectItem()
                        Loader.SIS.DoCommand("AComBoundary")
                        Loader.SIS.DoCommand("AComFillGeometry")
                        Loader.SIS.OpenSel(0)

                        Loader.SIS.Delete("List")
                        Loader.SIS.AddToList("List")

                        ' add buffer as primary item to the list of areas
                        Loader.SIS.CreateItemFromLocus("LOCUSBUFFER")
                        Loader.SIS.SelectItem()
                        Loader.SIS.DoCommand("AComBoundary")
                        Loader.SIS.DoCommand("AComFillGeometry")
                        Loader.SIS.OpenSel(0)
                        Loader.SIS.AddToList("List")

                        ' create a donut around the original clave
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
                        Loader.SIS.Delete("lTmp")
                        CleanAreaGeometry("List")

                    End If

                End If

            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub GetBDT()

        Try

            Dim areaBDT As Double = 0
            EmptyList("lTmp")
            Loader.SIS.CreateLocusFromItem("LocusIntersect", SIS_GT_INTERSECT, SIS_GM_GEOMETRY)
            Loader.SIS.AddToList("lTmp")

            For i As Integer = 0 To (Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)

                If Loader.SIS.GetStr(SIS_OT_OVERLAY, i, "_name$") = "BDT" Then

                    Loader.SIS.ScanOverlay("lBDT", i, "", "LocusIntersect")
                    Loader.SIS.OpenList("lTmp", 0)
                    Loader.SIS.AddToList("lBDT")
                    Exit For

                End If

            Next

            Loader.SIS.CreateBoolean("lBDT", SIS_BOOLEAN_AND)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 1001)
            areaBDT = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_area#")

            If Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "AnchoDdv#") > 0 Then

                Loader.SIS.DeleteItem()

            Else

                EmptyList("lBDT")
                Loader.SIS.AddToList("lBDT")
                Loader.SIS.OpenList("lAfectacion", 0)
                Loader.SIS.AddToList("lBDT")
                Loader.SIS.CreateBoolean("lBDT", SIS_BOOLEAN_AND)
                Loader.SIS.Delete("lAfectacion")
                Loader.SIS.AddToList("lAfectacion")
                Loader.SIS.Delete("lBDT")

            End If

            Loader.SIS.OpenList("lAfectacion", 0)
            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "areaBDT#", areaBDT)

        Catch ex As Exception
            'MsgBox(ex.ToString)
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