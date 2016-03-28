﻿Imports Intelexion.Framework.Model
Imports System.Data
Imports Intelexion.Nomina
Imports System.Web
Imports System.Collections
Imports System.Collections.Specialized
Imports Intelexion.Framework.Controller
Imports Intelexion.Framework.View
Imports System.IO
Imports System.Data.SqlClient

Public Class DAO
    Inherits Intelexion.Framework.Model.DAO

    Private Const ValesDespensaQNC As String = "sp_ValesDespensa_CargaSaldos '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@Periodo','@UID','@LID','@idAccion'"
    Private Const ValesDespensaTarjetas As String = "sp_ValesDespensa_SolicitudTarjetas '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@Periodo','@UID','@LID','@idAccion'"
    Private Const polizacr As String = "sp_Poliza_ContableCR '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@Periodo','@UID','@LID','@idAccion'"
    Private Const polizascc As String = "sp_Poliza_ContableSCC '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@Periodo','@UID','@LID','@idAccion'"
    Private Const LayoutSolicitudSantander As String = "sp_Reporte_LayoutSolicitudSantander '@IdRazonSocial','@IdTipoNominaAsig','@UID','@LID','@idAccion'"
    Private Const polizaPor As String = "sp_Poliza_ContableCR_Porciento '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@Periodo','@UID','@LID','@idAccion'"
    Private Const Declaracion As String = "sp_LayoutConstancia '@IdRazonSocial','@UID','@LID','@idAccion'"
    Private Const Acumulado As String = "sp_ReporteAcumuladosCR '@IdRazonSocial','@UID','@LID','@idAccion'"





    Public Sub New(ByVal DataConnection As SQLDataConnection)
        MyBase.New(DataConnection)
    End Sub
    Public Function Layout(ByVal ReportesProceso As EntitiesITX.ReportesProceso, ByVal opL As String) As DataSet
        Dim ds As New DataSet
        Dim resultado As Boolean
        Dim comandstr As String
        Try
            Select Case opL


                Case "polizacr"
                    comandstr = polizacr
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "polizascc"
                    comandstr = polizascc
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "ValesDespensaTarjetas"
                    comandstr = ValesDespensaTarjetas
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "ValesDespensaQNC"
                    comandstr = ValesDespensaQNC
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "polizaPor"
                    comandstr = polizaPor
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "Declaracion"
                    comandstr = Declaracion
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "Acumulado"
                    comandstr = Acumulado
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds



            End Select
        Catch e As Exception
        End Try
        Return ds
    End Function

    Public Function LayoutTxt(ByVal ReportesProceso As EntitiesITX.ReportesProceso, ByVal opL As String, ByVal context As System.Web.HttpContext) As String
        Dim ds As New DataSet
        Dim sFile As String
        Dim sPathApp As String = Intelexion.Framework.ApplicationConfiguration.ConfigurationSettings.GetConfigurationSettings("ApplicationPath")
        Dim sPathArchivosTemp As String = Intelexion.Framework.ApplicationConfiguration.ConfigurationSettings.GetConfigurationSettings("ArchivosTemporales")
        'Dim ruta As String
        Try
            Select Case opL

                


                Case "LayoutSolicitudSantander"
                    Dim results As ResultCollection
                    Dim objLayoutDispersion As Entities.LayoutDispersion
                    Dim dTotalImporte As Decimal
                    Dim sCadena As String
                    Dim i As Integer
                    results = New ResultCollection
                    ReportesProceso.tipoArchivo = 0
                    objLayoutDispersion = New Entities.LayoutDispersion
                    objLayoutDispersion.IdRazonSocial = context.Session("IdRazonSocial")
                    objLayoutDispersion.UID = context.Session("UID")
                    objLayoutDispersion.LID = context.Session("LID")
                    objLayoutDispersion.idAccion = context.Items.Item("idAccion")
                    Dim UserId As String
                    UserId = ReportesProceso.UID.ToString
                    UserId = UserId.Replace("/", "")
                    UserId = UserId.Replace("\", "")
                    UserId = UserId.Replace("%", "")
                    UserId = UserId.Replace("_", "")
                    UserId = UserId.Replace("-", "")
                    sFile = "\SolicitudSantander" + ReportesProceso.IdRazonSocial.ToString + UserId + Date.Now.Second.ToString + ".txt"

                    results.getEntitiesFromDataReader(objLayoutDispersion, Me.ReporteSolicitudSantander(ReportesProceso))
                    dTotalImporte = 0
                    If results.Count > 0 Then
                        Dim sb As New FileStream(context.Server.MapPath(sPathApp + sPathArchivosTemp) + sFile, FileMode.Create)
                        Dim sw As New StreamWriter(sb)

                        For i = 0 To results.Count - 1
                            sCadena = results(i).centroCosto
                            sw.WriteLine(sCadena)
                        Next i

                        sw.Close()

                    End If

                    Return sPathApp + sPathArchivosTemp + sFile



            End Select
        Catch e As Exception
        End Try
        Return ""
    End Function
    Public Function ReporteSolicitudSantander(ByVal ReportesProceso As EntitiesITX.ReportesProceso) As SqlDataReader
        Dim data As SqlDataReader = Nothing
        Dim resultado As Boolean
        Dim comandstr As String
        Try
            comandstr = LayoutSolicitudSantander
            resultado = Me.ExecuteQuery(comandstr, data, ReportesProceso)
            Return data
        Catch e As Exception
        End Try
        Return data
    End Function

   




End Class
