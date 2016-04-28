Imports System
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports DataDynamics.ActiveReports
Imports DataDynamics.ActiveReports.Document
Imports Intelexion.Nomina
Imports Nomina
Imports System.IO
'Imports System.Web
'Imports Intelexion.Framework.Model
'Imports Intelexion.Framework.View
'Imports System.Exception



Public Class ComparativoNominaMensual
    Inherits DataDynamics.ActiveReports.ActiveReport3

    Dim RPT As New ActiveReport3

    Dim sConnection As String = Intelexion.Framework.Model.SQLConnectionProvider.getConnection("default").getConnection.ConnectionString
    Public Sub New()

        'This call is required by the ActiveReports Designer.
        InitializeComponent()


    End Sub

    Public Sub New(ByVal reporte As Entities.ReportesProceso)

        'This call is required by the ActiveReports Designer.
        InitializeComponent()

    End Sub




    Public Sub CartaFiniquito_ReportStart(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.ReportStart
        Me.ShowParameterUI = False
        Me.SetLicense("Hector Ramirez,Intelexion,DD-APN-30-C000260,S4VKSH448MS77HKH9FH9")

        'Dim Path As String = Intelexion.Framework.ApplicationConfiguration.ConfigurationSettings.GetConfigurationSettings("ApplicationPath")
        'Dim picRazonSocial As String = "/logos/Logoempresa" + Me.Parameters("IdRazonSocial").Value + ".jpg"
        'Dim archivo As New FileInfo(System.Web.HttpContext.Current.Server.MapPath(Path) + picRazonSocial)


        'Picture1.Image = Image.FromFile(archivo.FullName)
        'Picture1.Visible = True



        Me.DataSource = New DataDynamics.ActiveReports.DataSources.OleDBDataSource("Provider=SQLOLEDB.1;" & _
                 sConnection, "sp_Reporte_ListadoNomina_DatosOrg_Mes_Aliatec @IdRazonSocial = " + Me.Parameters("IdRazonSocial").Value + "," & _
                 "@IdTipoNominaAsig = " + Me.Parameters("IdTipoNominaAsig").Value + "," & _
                 "@folioDesde = " + Me.Parameters("folioDesde").Value + "," & _
                 "@folioHasta = " + Me.Parameters("folioHasta").Value + "," & _
                 "@UID = " + Me.Parameters("UID").Value + "," & _
                 "@LID = " + Me.Parameters("LID").Value + "," & _
                 "@idAccion = " + Me.Parameters("idAccion").Value + "", 90)


        'parametros que se envian al subreporte abajo mencionados
        Dim ParamIdRazonSocial As New Parameter
        ParamIdRazonSocial.Key = "IdRazonSocial"

        Dim ParamIdTipoNominaAsig As New Parameter
        ParamIdTipoNominaAsig.Key = "IdTipoNominaAsig"
        ParamIdTipoNominaAsig.Type = Parameter.DataType.String

        Dim ParamfolioDesde As New Parameter
        ParamfolioDesde.Key = "folioDesde"
        ParamfolioDesde.Type = Parameter.DataType.String

        Dim ParamfolioHasta As New Parameter
        ParamfolioHasta.Key = "folioHasta"
        ParamfolioHasta.Type = Parameter.DataType.String

        Dim ParamUID As New Parameter
        ParamUID.Key = "UID"
        ParamUID.Type = Parameter.DataType.String

        Dim ParamLID As New Parameter
        ParamLID.Key = "LID"
        ParamLID.Type = Parameter.DataType.String

        Dim ParamidAccion As New Parameter
        ParamidAccion.Key = "idAccion"

        ''******* SUBREPORTE PERCEPCIONES **************/
        '                                  nombre del subreporte
        Dim SubReportePercepciones As New ComparativoMensualSubReporte
        SubReportePercepciones.Parameters.Clear()

        'aqui se agregan
        SubReportePercepciones.Parameters.Add(ParamIdRazonSocial)
        SubReportePercepciones.Parameters.Add(ParamIdTipoNominaAsig)
        SubReportePercepciones.Parameters.Add(ParamfolioDesde)
        SubReportePercepciones.Parameters.Add(ParamfolioHasta)
        SubReportePercepciones.Parameters.Add(ParamUID)
        SubReportePercepciones.Parameters.Add(ParamLID)
        SubReportePercepciones.Parameters.Add(ParamidAccion)


        'aqui se les asigna el valor
        SubReportePercepciones.Parameters("IdRazonSocial").Value = Me.Parameters("IdRazonSocial").Value
        SubReportePercepciones.Parameters("IdTipoNominaAsig").Value = Me.Parameters("IdTipoNominaAsig").Value
        SubReportePercepciones.Parameters("folioDesde").Value = Me.Parameters("folioDesde").Value
        SubReportePercepciones.Parameters("folioHasta").Value = Me.Parameters("folioHasta").Value
        SubReportePercepciones.Parameters("UID").Value = Me.Parameters("UID").Value
        SubReportePercepciones.Parameters("LID").Value = Me.Parameters("LID").Value
        SubReportePercepciones.Parameters("idAccion").Value = Me.Parameters("idAccion").Value
        Me.SubReport3.Report = SubReportePercepciones

    End Sub


    Private Sub Detail1_Format(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Detail1.Format

    End Sub

    Private Sub PageHeader1_Format(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class
