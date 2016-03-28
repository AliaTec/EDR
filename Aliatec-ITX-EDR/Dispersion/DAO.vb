Imports Intelexion.Framework.Model
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
    Private Const ImpuestoSobreNominaAgu As String = "sp_Reporte_Impuestos_Sobre_Nomina_Agu '@IdRazonSocial','@IdTipoNominaAsig','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const ImpuestoSobreNominaCan As String = "sp_Reporte_Impuestos_Sobre_Nomina_Can '@IdRazonSocial','@IdTipoNominaAsig','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const ImpuestoSobreNominaJUCH As String = "sp_Reporte_Impuestos_Sobre_Nomina_JUCH '@IdRazonSocial','@IdTipoNominaAsig','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const ImpuestoSobreNominaCUE As String = "sp_Reporte_Impuestos_Sobre_Nomina_CUE '@IdRazonSocial','@IdTipoNominaAsig','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const ImpuestoSobreNominaGUA As String = "sp_Reporte_Impuestos_Sobre_Nomina_GUA '@IdRazonSocial','@IdTipoNominaAsig','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const ImpuestoSobreNominaHER As String = "sp_Reporte_Impuestos_Sobre_Nomina_HER '@IdRazonSocial','@IdTipoNominaAsig','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const ImpuestoSobreNominaLEO As String = "sp_Reporte_Impuestos_Sobre_Nomina_LEO '@IdRazonSocial','@IdTipoNominaAsig','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const ImpuestoSobreNominaMER As String = "sp_Reporte_Impuestos_Sobre_Nomina_MER '@IdRazonSocial','@IdTipoNominaAsig','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const ImpuestoSobreNominaMEX As String = "sp_Reporte_Impuestos_Sobre_Nomina_MEX '@IdRazonSocial','@IdTipoNominaAsig','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const ImpuestoSobreNominaMTY As String = "sp_Reporte_Impuestos_Sobre_Nomina_MTY '@IdRazonSocial','@IdTipoNominaAsig','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const ImpuestoSobreNominaPN As String = "sp_Reporte_Impuestos_Sobre_Nomina_PN '@IdRazonSocial','@IdTipoNominaAsig','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const ImpuestoSobreNominaPUE As String = "sp_Reporte_Impuestos_Sobre_Nomina_PUE '@IdRazonSocial','@IdTipoNominaAsig','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const ImpuestoSobreNominaQRO As String = "sp_Reporte_Impuestos_Sobre_Nomina_QRO '@IdRazonSocial','@IdTipoNominaAsig','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const ImpuestoSobreNominaSLP As String = "sp_Reporte_Impuestos_Sobre_Nomina_SLP '@IdRazonSocial','@IdTipoNominaAsig','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const ImpuestoSobreNominaTIJ As String = "sp_Reporte_Impuestos_Sobre_Nomina_TIJ '@IdRazonSocial','@IdTipoNominaAsig','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const ImpuestoSobreNominaVER As String = "sp_Reporte_Impuestos_Sobre_Nomina_VER '@IdRazonSocial','@IdTipoNominaAsig','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const MenosSalarioMinimoDF As String = "sp_Reporte_MenosSalarioMinimo_Df '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@Periodo','@UID','@LID','@idAccion'"
    Private Const LayoutPolizaContable As String = "spq_nomPolizaEdenred '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@Periodo','@UID','@LID','@idAccion'"
    Private Const LayoutPolizaContabletext As String = "spq_nomPolizaEdenredLayout '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@Periodo','@UID','@LID','@idAccion'"
    Private Const ImpuestoSobreNominaSucursales As String = "spq_nomFinLiq '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const ImpuestoSobreNominaDf As String = "sp_Reporte_Impuestos_Sobre_Nomina_DF '@IdRazonSocial','@IdTipoNominaAsig','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const LayoutTreinta As String = "sp_Reporte_Listado_Treinta '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@Periodo','@UID','@LID','@idAccion'"
    Private Const LayoutDetallada As String = "sp_Reporte_Nomina_Detallada '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@Periodo','@UID','@LID','@idAccion'"
    Private Const LayoutFormaPago As String = "sp_Reporte_Forma_Pago '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@Periodo','@UID','@LID','@idAccion'"
    Private Const LayoutNetosPago As String = "sp_Reporte_Netos_Pago '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@Periodo','@UID','@LID','@idAccion'"
    Private Const LayoutExcepciones As String = "sp_Reporte_Nomina_Excepciones '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@Periodo','@UID','@LID','@idAccion'"
    Private Const LayoutHSBC As String = "sp_Reporte_LayoutDispersionHSBC_Edenred '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@Periodo','@UID','@LID','@idAccion'"
    Private Const LayoutCentroCos As String = "spq_nomCentroCostos '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@Periodo','@UID','@LID','@idAccion'"
    Private Const LayoutFinLiq As String = "spq_nomFinLiq '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const LayoutFinLiqDet As String = "spq_nomFinLiqDet '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const LayoutGraExenDF As String = "sp_Reporte_Grav_ExenDF '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const LayoutGraExenSuc As String = "sp_Reporte_Grav_ExenSUC '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@mesDesde','@mesHasta','@UID','@LID','@idAccion'"
    Private Const LayoutDetalleCC As String = "spq_nomDetalleCC '@IdRazonSocial','@IdTipoNominaAsig','@IdTipoNominaProc','@Anio','@Periodo','@UID','@LID','@idAccion'"





    Public Sub New(ByVal DataConnection As SQLDataConnection)
        MyBase.New(DataConnection)
    End Sub
    Public Function Layout(ByVal ReportesProceso As EntitiesITX.ReportesProceso, ByVal opL As String) As DataSet
        Dim ds As New DataSet
        Dim resultado As Boolean
        Dim comandstr As String
        Try
            Select Case opL
                Case "ImpuestoSobreNominaAgu"
                    comandstr = ImpuestoSobreNominaAgu
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds
                Case "ImpuestoSobreNominaCAN"
                    comandstr = ImpuestoSobreNominaCan
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds
                Case "ImpuestoSobreNominaJUCH"
                    comandstr = ImpuestoSobreNominaJUCH
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds
                Case "ImpuestoSobreNominaCUE"
                    comandstr = ImpuestoSobreNominaCUE
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds
                Case "ImpuestoSobreNominaGUA"
                    comandstr = ImpuestoSobreNominaGUA
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds
                Case "ImpuestoSobreNominaHER"
                    comandstr = ImpuestoSobreNominaHER
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds
                Case "ImpuestoSobreNominaLEO"
                    comandstr = ImpuestoSobreNominaLEO
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds
                Case "ImpuestoSobreNominaMER"
                    comandstr = ImpuestoSobreNominaMER
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds
                Case "ImpuestoSobreNominaMEX"
                    comandstr = ImpuestoSobreNominaMEX
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds
                Case "ImpuestoSobreNominaMTY"
                    comandstr = ImpuestoSobreNominaMTY
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds
                Case "ImpuestoSobreNominaPN"
                    comandstr = ImpuestoSobreNominaPN
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds
                Case "ImpuestoSobreNominaPUE"
                    comandstr = ImpuestoSobreNominaPUE
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds
                Case "ImpuestoSobreNominaQRO"
                    comandstr = ImpuestoSobreNominaQRO
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds
                Case "ImpuestoSobreNominaSLP"
                    comandstr = ImpuestoSobreNominaSLP
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds
                Case "ImpuestoSobreNominaTIJ"
                    comandstr = ImpuestoSobreNominaTIJ
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds
                Case "ImpuestoSobreNominaVER"
                    comandstr = ImpuestoSobreNominaVER
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "MenosSalarioMinimoDF"
                    comandstr = MenosSalarioMinimoDF
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "LayoutPolizaContable"
                    comandstr = LayoutPolizaContable
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "ImpuestoSobreNominaSucursales"
                    comandstr = ImpuestoSobreNominaSucursales
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "ImpuestoSobreNominaDf"
                    comandstr = ImpuestoSobreNominaDf
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "LayoutTreinta"
                    comandstr = LayoutTreinta
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds


                Case "LayoutDetallada"
                    comandstr = LayoutDetallada
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "LayoutFormaPago"
                    comandstr = LayoutFormaPago
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "LayoutNetosPago"
                    comandstr = LayoutNetosPago
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "LayoutExcepciones"
                    comandstr = LayoutExcepciones
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "LayoutCentroCos"
                    comandstr = LayoutCentroCos
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "LayoutFinLiqDet"
                    comandstr = LayoutFinLiqDet
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "LayoutGraExenDF"
                    comandstr = LayoutGraExenDF
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "LayoutGraExenSuc"
                    comandstr = LayoutGraExenSuc
                    resultado = Me.ExecuteQuery(comandstr, ds, ReportesProceso)
                    Return ds

                Case "LayoutDetalleCC"
                    comandstr = LayoutDetalleCC
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
                Case "LayoutHSBC"
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
                    sFile = "\LayoutHSBC" + ReportesProceso.IdRazonSocial.ToString + UserId + Date.Now.Second.ToString + ".txt"

                    results.getEntitiesFromDataReader(objLayoutDispersion, Me.ReporteLayoutHSBC(ReportesProceso))
                    dTotalImporte = 0
                    If results.Count > 0 Then
                        Dim sb As New FileStream(context.Server.MapPath(sPathApp + sPathArchivosTemp) + sFile, FileMode.Create)
                        Dim sw As New StreamWriter(sb)

                        For i = 0 To results.Count - 1
                            sCadena = results(i).importe
                            sw.WriteLine(sCadena)
                        Next i

                        sw.Close()

                    End If

                    Return sPathApp + sPathArchivosTemp + sFile

                Case "LayoutPolizaContabletext"
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
                    sFile = "\LayoutPolizaContabletext" + ReportesProceso.IdRazonSocial.ToString + UserId + Date.Now.Second.ToString + ".txt"

                    results.getEntitiesFromDataReader(objLayoutDispersion, Me.ReporteLayoutPoliza(ReportesProceso))
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



    Public Function ReporteLayoutHSBC(ByVal ReportesProceso As EntitiesITX.ReportesProceso) As SqlDataReader
        Dim data As SqlDataReader = Nothing
        Dim resultado As Boolean
        Dim comandstr As String
        Try
            comandstr = LayoutHSBC
            resultado = Me.ExecuteQuery(comandstr, data, ReportesProceso)
            Return data
        Catch e As Exception
        End Try
        Return data


    End Function

    Public Function ReporteLayoutPoliza(ByVal ReportesProceso As EntitiesITX.ReportesProceso) As SqlDataReader
        Dim data As SqlDataReader = Nothing
        Dim resultado As Boolean
        Dim comandstr As String
        Try
            comandstr = LayoutPolizaContabletext
            resultado = Me.ExecuteQuery(comandstr, data, ReportesProceso)
            Return data
        Catch e As Exception
        End Try
        Return data

    End Function


End Class
