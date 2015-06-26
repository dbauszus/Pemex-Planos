Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants
Imports System.IO
Imports System.Windows.Forms

<GisLinkProgram("Pemex Addin")> _
Public Class Loader

    Private Shared APP As SisApplication
    Private Shared _sis As MapModeller

    Public Shared Property SIS As MapModeller
        Get
            If _sis Is Nothing Then _sis = APP.TakeoverMapManager
            Return _sis
        End Get
        Set(ByVal value As MapModeller)
            _sis = value
        End Set
    End Property

    Public Sub New(ByVal SISApplication As SisApplication)
        APP = SISApplication

        Dim group As SisRibbonGroup = APP.RibbonGroup
        group.Text = "PEMEX"

        Dim btnPemex_Cortar As SisRibbonButton = New SisRibbonButton("CORTAR", New SisClickHandler(AddressOf Pemex_Cortar))
        group.Controls.Add(btnPemex_Cortar)

        Dim btnPemex_KMITxt As SisRibbonButton = New SisRibbonButton("KMI TXT", New SisClickHandler(AddressOf Pemex_KMITxt))
        group.Controls.Add(btnPemex_KMITxt)

        Dim btnPemex_KMFTxt As SisRibbonButton = New SisRibbonButton("KMF TXT", New SisClickHandler(AddressOf Pemex_KMFTxt))
        group.Controls.Add(btnPemex_KMFTxt)

        Dim btnPemex_LineaTxt As SisRibbonButton = New SisRibbonButton("LINEA TXT", New SisClickHandler(AddressOf Pemex_LineaTxt))
        group.Controls.Add(btnPemex_LineaTxt)

        Dim btnPemex_Colindantes_new As SisRibbonButton = New SisRibbonButton("COLINDANTES", New SisClickHandler(AddressOf Pemex_Colindantes_new))
        btnPemex_Colindantes_new.MinSelection = 1
        btnPemex_Colindantes_new.MaxSelection = 1
        group.Controls.Add(btnPemex_Colindantes_new)

        Dim btnPemex_Construccion As SisRibbonButton = New SisRibbonButton("CONSTRUCCIÓN", New SisClickHandler(AddressOf Pemex_Construccion))
        btnPemex_Construccion.MinSelection = 1
        btnPemex_Construccion.MaxSelection = 1
        group.Controls.Add(btnPemex_Construccion)

        Dim btnPemex_Vertices As SisRibbonButton = New SisRibbonButton("VERTICES", New SisClickHandler(AddressOf Pemex_Vertices))
        group.Controls.Add(btnPemex_Vertices)

        Dim btnPemex_Seccion As SisRibbonButton = New SisRibbonButton("SECCION", New SisClickHandler(AddressOf Pemex_Seccion))
        btnPemex_Seccion.MinSelection = 1
        btnPemex_Seccion.MaxSelection = 1
        group.Controls.Add(btnPemex_Seccion)

        Dim btnPemex_Angulos As SisRibbonButton = New SisRibbonButton("ANGULOS", New SisClickHandler(AddressOf Pemex_Angulos))
        btnPemex_Angulos.MinSelection = 1
        btnPemex_Angulos.MaxSelection = 1
        group.Controls.Add(btnPemex_Angulos)

        Dim btnPemex_View As SisRibbonButton = New SisRibbonButton("SET VIEW", New SisClickHandler(AddressOf Pemex_View))
        group.Controls.Add(btnPemex_View)

        Dim btnPemex_ResetView As SisRibbonButton = New SisRibbonButton("VIEW 0", New SisClickHandler(AddressOf Pemex_ResetView))
        group.Controls.Add(btnPemex_ResetView)

        'Dim btnPemex_SplitLines As SisRibbonButton = New SisRibbonButton("SPLIT LINES", New SisClickHandler(AddressOf SplitLines))
        'group.Controls.Add(btnPemex_SplitLines)

        'Dim btnPemex_Duplicates As SisRibbonButton = New SisRibbonButton("DUPLICATES", New SisClickHandler(AddressOf Duplicates))
        'group.Controls.Add(btnPemex_Duplicates)

        'Dim btnPemex_NewSeccion As SisRibbonButton = New SisRibbonButton("NEW SECCION", New SisClickHandler(AddressOf Pemex_NewSeccion))
        'group.Controls.Add(btnPemex_NewSeccion)

        'Dim btnPemex_Distancias As SisRibbonButton = New SisRibbonButton("DISTANCIAS", New SisClickHandler(AddressOf Pemex_Distancias))
        'group.Controls.Add(btnPemex_Distancias)

        'Dim btnCreateFeatureFilter As SisRibbonButton = New SisRibbonButton("feature filter", New SisClickHandler(AddressOf feature_filter))
        'group.Controls.Add(btnCreateFeatureFilter)

        'Dim btnCreateUnFill_Fill As SisRibbonButton = New SisRibbonButton("unfill_fill", New SisClickHandler(AddressOf unfill_fill))
        'group.Controls.Add(btnCreateUnFill_Fill)

    End Sub

    Private Sub Pemex_Cortar(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            Loader.SIS = e.MapModeller
            Dim Cor As Pemex_Cortar = New Pemex_Cortar
            Cor.Cortar()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Pemex_Colindantes_new(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            Loader.SIS = e.MapModeller
            Dim Col As Pemex_Colindantes = New Pemex_Colindantes
            Col.Colindantes_new()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Pemex_Construccion(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            Loader.SIS = e.MapModeller
            Dim Ver As Pemex_Vertices = New Pemex_Vertices
            Ver.Construccion()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Pemex_Vertices(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            Loader.SIS = e.MapModeller
            Dim Ver As Pemex_Vertices = New Pemex_Vertices
            Ver.Vertices()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Pemex_Seccion(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            Loader.SIS = e.MapModeller
            Dim Pla As Pemex_Plano = New Pemex_Plano
            SIS.OpenSel(0)
            If SIS.GetInt(SIS_OT_CURITEM, 0, "_FC&") = 1202 Then
                Pla.Seccion()
            Else
                MsgBox("No seleccion linea Oleoducto 36", MsgBoxStyle.Exclamation)
            End If
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Pemex_Angulos(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            Loader.SIS = e.MapModeller
            Dim Ang As Pemex_Angulos = New Pemex_Angulos
            SIS.OpenSel(0)
            If SIS.GetInt(SIS_OT_CURITEM, 0, "_FC&") = 1202 Then
                Ang.Angulos()
            Else
                MsgBox("No seleccion linea Oleoducto 36", MsgBoxStyle.Exclamation)
            End If
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Pemex_LineaTxt(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            Loader.SIS = e.MapModeller
            Dim Lin As Pemex_LineaTxt = New Pemex_LineaTxt
            Lin.LineaTxt()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Pemex_KMITxt(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            Loader.SIS = e.MapModeller
            Dim Lin As Pemex_LineaTxt = New Pemex_LineaTxt
            Lin.KMITxt()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Pemex_KMFTxt(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            Loader.SIS = e.MapModeller
            Dim Lin As Pemex_LineaTxt = New Pemex_LineaTxt
            Lin.KMFTxt()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Pemex_View(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            Loader.SIS = e.MapModeller
            Dim Pla As Pemex_Plano = New Pemex_Plano
            Pla.SetViewAngle()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Pemex_ResetView(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            Loader.SIS = e.MapModeller
            Dim Pla As Pemex_Plano = New Pemex_Plano
            Pla.ResetViewAngle()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Pemex_NewSeccion(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            Loader.SIS = e.MapModeller
            Dim Sec As Pemex_Seccion = New Pemex_Seccion
            Sec.Secciones()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Pemex_Distancias(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            Loader.SIS = e.MapModeller
            Dim Sec As Pemex_Seccion = New Pemex_Seccion
            Sec.Distancias()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub SplitLines(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            Loader.SIS = e.MapModeller
            Dim Mea As MeasureIntersections = New MeasureIntersections
            Mea.SplitLines()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Duplicates(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            Loader.SIS = e.MapModeller
            Dim Mea As MeasureIntersections = New MeasureIntersections
            Mea.Duplicates()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub feature_filter(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            Loader.SIS = e.MapModeller
            Loader.SIS.CreateFeatureFilter("fPEMEX", "PEMEX")
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub number_points(ByVal sender As Object, ByVal e As SisClickArgs)

        Try
            Loader.SIS = e.MapModeller

            Dim n As Integer = 0
            Loader.SIS.CreateListFromSelection("list")

            For i = 0 To Loader.SIS.GetListSize("list") - 1

                Loader.SIS.OpenList("list", i)

                For ii = 0 To Loader.SIS.GetNumGeom() - 1

                    n += Loader.SIS.GetGeomNumPt(ii)

                Next

            Next

            MsgBox(n.ToString)

            SIS.Dispose()
            SIS = Nothing

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub unfill_fill(ByVal sender As Object, ByVal e As SisClickArgs)

        Try
            Loader.SIS = e.MapModeller

            Dim n As Integer = 0
            Loader.SIS.CreateListFromSelection("list")

            For i = 0 To Loader.SIS.GetListSize("list") - 1

                Loader.SIS.DeselectAll()
                Loader.SIS.OpenList("list", i)
                Loader.SIS.SelectItem()
                Loader.SIS.DoCommand("AComBoundary")
                Loader.SIS.DoCommand("AComFillGeometry")

            Next
            SIS.Dispose()
            SIS = Nothing

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

End Class