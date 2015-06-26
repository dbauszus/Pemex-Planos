Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants

Public Class MeasureIntersections

    Public Sub New()

    End Sub

    Public Sub SplitLines()

        Try

            Loader.SIS.CreateListFromSelection("list")

            Dim x1, y1, z1, x2, y2, z2 As Double
            For i = 0 To Loader.SIS.GetListSize("list") - 1

                Loader.SIS.OpenList("list", i)
                For ii = 0 To (Loader.SIS.GetGeomNumPt(0) - 2)

                    Loader.SIS.OpenList("list", i)
                    Loader.SIS.SplitPos(x1, y1, z1, Loader.SIS.GetGeomPt(0, ii))
                    Loader.SIS.SplitPos(x2, y2, z2, Loader.SIS.GetGeomPt(0, ii + 1))
                    Loader.SIS.MoveTo(x1, y1, z1)
                    Loader.SIS.LineTo(x2, y2, z2)
                    Loader.SIS.UpdateItem()

                Next

            Next
            Loader.SIS.Delete("list")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Sub Duplicates()

        Try

            Loader.SIS.CreateListFromSelection("list")
            Loader.SIS.SetListInt("list", "DUPLICATE&", 0)

            Dim n = 1
            For i = 0 To Loader.SIS.GetListSize("list") - 1

                Loader.SIS.OpenList("list", i)
                If Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "DUPLICATE&") = 0 Then

                    If Loader.SIS.ScanGeometry("lDuplicate", SIS_GT_CONTAIN, SIS_GM_GEOMETRY, "", "") > 1 Then

                        Loader.SIS.SetListInt("lDuplicate", "DUPLICATE&", n)
                        n += 1

                    End If

                End If

            Next

            For i = 1 To n

                Loader.SIS.CreatePropertyFilter("filter", "DUPLICATE&=" & Str(i))
                If Loader.SIS.Scan("scan", "E", "filter", "") > 1 Then

                    Loader.SIS.OpenList("scan", 0)
                    Loader.SIS.RemoveFromList("scan")
                    Loader.SIS.Delete("scan")

                End If

            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

End Class
