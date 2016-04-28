Imports System
Imports System.IO
Imports System.Web
Imports DataDynamics.ActiveReports
Imports Intelexion.Framework.Model
Imports Intelexion.Nomina
Imports Intelexion.Framework.View
Imports System.Collections.Generic
Imports System.Web.HttpContext


Partial Class ComparativoMensual

    Inherits System.Web.UI.Page

    Public RPT As New ComparativoNominaMensual

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub


    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim sConnection As String = Intelexion.Framework.Model.SQLConnectionProvider.getConnection("default").getConnection.ConnectionString
        'Dim sConnection As String = "Data Source=DCW319V1\MSSQLSERVER8; Initial Catalog=V5McGrawHillNominaTest; User Id=ITXTV5; Password=ITXTV5; Connection Lifetime=60; Max Pool Size=50; Min Pool Size=3"



        RPT.DataSource = New DataDynamics.ActiveReports.DataSources.OleDBDataSource("Provider=SQLOLEDB.1;" + sConnection, "sp_Reporte_ListadoNomina_DatosOrg_Mes_Aliatec @IdRazonSocial = <%IdRazonSocial%>,@IdTipoNominaAsig=<" & _
                "%IdTipoNominaAsig%>,@folioDesde = <%folioDesde%>,@folioHasta = <%folioHasta%>,@UID = <%UID%>, @LID = <%LID%>, @idAccion = <%idAccion%>", 90)



        Dim ParamIdRazonSocial As New Parameter
        ParamIdRazonSocial.Key = "IdRazonSocial"
        
        Dim ParamIdTipoNominaAsig As New Parameter
        ParamIdTipoNominaAsig.Key = "IdTipoNominaAsig"
        ParamIdTipoNominaAsig.Type = Parameter.DataType.String

        Dim ParamfolioHasta As New Parameter
        ParamfolioHasta.Key = "folioHasta"
        ParamIdTipoNominaAsig.Type = Parameter.DataType.String

        Dim ParamfolioDesde As New Parameter
        ParamfolioDesde.Key = "folioDesde"
        ParamIdTipoNominaAsig.Type = Parameter.DataType.String

        Dim ParamUID As New Parameter
        ParamUID.Key = "UID"
        ParamUID.Type = Parameter.DataType.String

        Dim ParamLID As New Parameter
        ParamLID.Key = "LID"
        ParamLID.Type = Parameter.DataType.String

        Dim ParamidAccion As New Parameter
        ParamidAccion.Key = "idAccion"

        RPT.Parameters.Clear()
        RPT.Parameters.Add(ParamIdRazonSocial)
        RPT.Parameters.Add(ParamIdTipoNominaAsig)
        RPT.Parameters.Add(ParamfolioHasta)
        RPT.Parameters.Add(ParamfolioDesde)
        RPT.Parameters.Add(ParamUID)
        RPT.Parameters.Add(ParamLID)
        RPT.Parameters.Add(ParamidAccion)


        RPT.Parameters("IdRazonSocial").Value = Request.Params("IdRazonSocial")
        RPT.Parameters("folioHasta").Value = Request.Params("folioHasta")
        RPT.Parameters("folioDesde").Value = Request.Params("folioDesde")
        RPT.Parameters("UID").Value = "'" + Context.Session("UID") + "'"
        RPT.Parameters("LID").Value = "'" + Context.Session("LID") + "'"
        RPT.Parameters("idAccion").Value = Request.Params("idAccion")
        RPT.Parameters("IdTipoNominaAsig").Value = "'" + Request.Params("IdTipoNominaAsig") + "'"

        'LogoEmpresa("4", RPT)

        Try
            RPT.Run()
            WebViewer1.Report = RPT
        Catch ex As Exception

        End Try
        Try
            Dim pathApp = MapPath("~")
            pathApp = pathApp + "\ArchivosTemp\"
            Dim strName As String
            strName = "ComparativoNominaMensual" + Context.Session("UID").trim + Date.Now.Day.ToString + Date.Now.Month.ToString + Date.Now.Year.ToString + Date.Now.Minute.ToString + Date.Now.Second.ToString + ".xls"
            Dim objXls As New DataDynamics.ActiveReports.Export.Xls.XlsExport
            objXls.Export(RPT.Document, pathApp + strName)
            ligaExcel.Text = "<a href='" + Intelexion.Framework.ApplicationConfiguration.ConfigurationSettings.GetConfigurationSettings("ApplicationPath") + "/ArchivosTemp/" + strName + "' >" + "Exportar Excel" + "</a>"
        Catch ex As Exception



        End Try



        'RPT.Document.Save(context.Server.MapPath(path + "\ReportOutput\axreport.rdf"), DataDynamics.ActiveReports.Document.RdfFormat.AR20)

    End Sub



End Class




