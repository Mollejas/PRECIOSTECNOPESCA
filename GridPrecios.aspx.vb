' ============================================================
' GridPrecios.aspx.vb - Code Behind
' ============================================================
' Acceso DIRECTO a la DBF vía OLEDB sin API, sin Web.config
' Solo PageMethods llamados desde JavaScript
' ============================================================

Imports System.Web.Services
Imports System.Data.OleDb

Public Class GridPrecios
    Inherits System.Web.UI.Page

    ' ╔══════════════════════════════════════════════════════════╗
    ' ║  CAMBIAR ESTA RUTA A DONDE ESTÁ TU ARCHIVO fcuapr1.DBF ║
    ' ╚══════════════════════════════════════════════════════════╝
    Private Const DBF_FOLDER As String = "Z:\"
    Private Const DBF_TABLE As String = "fcuapr1"

    Private Shared Function ConnStr() As String
        Return "Provider=VFPOLEDB.1;" &
               "Data Source=" & DBF_FOLDER & ";" &
               "Collating Sequence=MACHINE;" &
               "Exclusive=No;" &
               "NULL=No;"
    End Function

    ' ===========================
    ' Clases para retornar datos
    ' ===========================
    Public Class ItemPrecio
        Public Property Clave As String
        Public Property Lista As String
        Public Property Precio As String
        Public Property Desc As String
    End Class

    Public Class ResultBusqueda
        Public Property Clave As String
        Public Property Lista As String
        Public Property Encontrado As Boolean
        Public Property Precio As String
        Public Property Desc As String
    End Class

    Public Class ItemBusqueda
        Public Property Clave As String
        Public Property Lista As String
    End Class

    Public Class ItemGuardar
        Public Property Clave As String
        Public Property Lista As String
        Public Property Precio As String
    End Class

    Public Class ResultGuardar
        Public Property Actualizados As Integer
        Public Property Total As Integer
        Public Property Errores As List(Of String)
    End Class

    ' ===========================
    ' BUSCAR UN PRECIO (clave + lista)
    ' Llamado automáticamente al pegar o salir del campo
    ' ===========================
    <WebMethod>
    Public Shared Function BuscarPrecio(clave As String, lista As String) As ResultBusqueda
        Dim res As New ResultBusqueda() With {
            .Clave = clave,
            .Lista = lista,
            .Encontrado = False,
            .Precio = "",
            .Desc = ""
        }

        Try
            Using conn As New OleDbConnection(ConnStr())
                conn.Open()

                Dim sql As String = String.Format(
                    "SELECT APRPRC, APRDESC FROM {0} WHERE APRCLAVE = ? AND APRLISTA = ?",
                    DBF_TABLE)

                Using cmd As New OleDbCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@p1", clave.Trim().PadRight(25))
                    cmd.Parameters.AddWithValue("@p2", lista.Trim().PadRight(3))

                    Using rd = cmd.ExecuteReader()
                        If rd.Read() Then
                            res.Encontrado = True
                            res.Precio = Convert.ToDecimal(rd("APRPRC")).ToString("F2")
                            res.Desc = rd("APRDESC").ToString().Trim()
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' Silenciar para no romper el JS
            res.Encontrado = False
            res.Precio = "ERR: " & ex.Message
        End Try

        Return res
    End Function

    ' ===========================
    ' BUSCAR MÚLTIPLES PRECIOS (bulk)
    ' ===========================
    <WebMethod>
    Public Shared Function BuscarPrecios(items As List(Of ItemBusqueda)) As List(Of ResultBusqueda)
        Dim resultados As New List(Of ResultBusqueda)

        If items Is Nothing OrElse items.Count = 0 Then
            Return resultados
        End If

        Dim filtros As New List(Of String)
        Dim valores As New List(Of Object)

        For Each item In items
            filtros.Add("(APRCLAVE = ? AND APRLISTA = ?)")
            valores.Add(item.Clave.Trim().PadRight(25))
            valores.Add(item.Lista.Trim().PadRight(3))
        Next

        Dim sql As String = String.Format(
            "SELECT APRCLAVE, APRLISTA, APRPRC, APRDESC FROM {0} WHERE {1}",
            DBF_TABLE,
            String.Join(" OR ", filtros))

        Dim encontrados As New Dictionary(Of String, ResultBusqueda)(StringComparer.OrdinalIgnoreCase)

        Try
            Using conn As New OleDbConnection(ConnStr())
                conn.Open()
                Using cmd As New OleDbCommand(sql, conn)
                    For Each paramValue In valores
                        cmd.Parameters.AddWithValue("@p", paramValue)
                    Next

                    Using rd = cmd.ExecuteReader()
                        While rd.Read()
                            Dim clave As String = rd("APRCLAVE").ToString().Trim()
                            Dim lista As String = rd("APRLISTA").ToString().Trim()
                            Dim key As String = clave & "|" & lista

                            If Not encontrados.ContainsKey(key) Then
                                encontrados(key) = New ResultBusqueda() With {
                                    .Clave = clave,
                                    .Lista = lista,
                                    .Encontrado = True,
                                    .Precio = Convert.ToDecimal(rd("APRPRC")).ToString("F2"),
                                    .Desc = rd("APRDESC").ToString().Trim()
                                }
                            End If
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Throw New Exception("Error al buscar precios: " & ex.Message)
        End Try

        For Each item In items
            Dim key As String = item.Clave.Trim() & "|" & item.Lista.Trim()
            If encontrados.ContainsKey(key) Then
                resultados.Add(encontrados(key))
            Else
                resultados.Add(New ResultBusqueda() With {
                    .Clave = item.Clave,
                    .Lista = item.Lista,
                    .Encontrado = False,
                    .Precio = "",
                    .Desc = ""
                })
            End If
        Next

        Return resultados
    End Function

    ' ===========================
    ' CARGAR TODOS LOS REGISTROS
    ' ===========================
    <WebMethod>
    Public Shared Function CargarTodos() As List(Of ItemPrecio)
        Dim items As New List(Of ItemPrecio)

        Try
            Using conn As New OleDbConnection(ConnStr())
                conn.Open()

                Dim sql As String = String.Format(
                    "SELECT APRCLAVE, APRLISTA, APRPRC, APRDESC FROM {0} ORDER BY APRCLAVE, APRLISTA",
                    DBF_TABLE)

                Using cmd As New OleDbCommand(sql, conn)
                    Using rd = cmd.ExecuteReader()
                        While rd.Read()
                            items.Add(New ItemPrecio() With {
                                .Clave = rd("APRCLAVE").ToString().Trim(),
                                .Lista = rd("APRLISTA").ToString().Trim(),
                                .Precio = Convert.ToDecimal(rd("APRPRC")).ToString("F2"),
                                .Desc = rd("APRDESC").ToString().Trim()
                            })
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Throw New Exception("Error al cargar DBF: " & ex.Message)
        End Try

        Return items
    End Function

    ' ===========================
    ' Sanitizar texto para incrustar en script FoxPro
    ' ===========================
    Private Shared Function FoxSafe(s As String) As String
        Return s.Trim() _
               .Replace("'", "''") _
               .Replace(vbCr, "") _
               .Replace(vbLf, "") _
               .Replace(vbTab, "")
    End Function

    ' ===========================
    ' GUARDAR PRECIOS (masivo vía EXECSCRIPT)
    ' Envía un solo script FoxPro con todos los UPDATEs
    ' ===========================
    <WebMethod>
    Public Shared Function GuardarPrecios(items As List(Of ItemGuardar)) As ResultGuardar
        Dim res As New ResultGuardar() With {
            .Actualizados = 0,
            .Total = items.Count,
            .Errores = New List(Of String)
        }

        ' Validar precios y construir lista limpia
        Dim validos As New List(Of Tuple(Of String, String, Decimal))
        For Each item In items
            Dim precio As Decimal = 0
            If Not Decimal.TryParse(item.Precio, precio) Then
                res.Errores.Add(String.Format("{0}: precio inválido '{1}'", item.Clave, item.Precio))
                Continue For
            End If
            validos.Add(Tuple.Create(item.Clave, item.Lista, precio))
        Next

        If validos.Count = 0 Then
            Return res
        End If

        Try
            Using conn As New OleDbConnection(ConnStr())
                conn.Open()

                ' Procesar en lotes de 50 para no exceder límites de EXECSCRIPT
                Dim batchSize As Integer = 50
                For b As Integer = 0 To validos.Count - 1 Step batchSize
                    Dim chunk = validos.GetRange(b, Math.Min(batchSize, validos.Count - b))

                    Dim script As New System.Text.StringBuilder()
                    script.AppendLine("LOCAL lnOk")
                    script.AppendLine("lnOk = 0")

                    For Each t In chunk
                        Dim safeClave = FoxSafe(t.Item1)
                        Dim safeLista = FoxSafe(t.Item2)
                        Dim priceStr = t.Item3.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)

                        script.AppendLine(String.Format(
                            "UPDATE {0} SET APRPRC = {1} WHERE ALLTRIM(APRCLAVE) = '{2}' AND ALLTRIM(APRLISTA) = '{3}'",
                            DBF_TABLE, priceStr, safeClave, safeLista))
                        script.AppendLine("lnOk = lnOk + _TALLY")
                    Next

                    script.AppendLine("RETURN lnOk")

                    Using cmd As New OleDbCommand("EXECSCRIPT(?)", conn)
                        cmd.Parameters.AddWithValue("@s", script.ToString())
                        Dim resultado = cmd.ExecuteScalar()
                        res.Actualizados += Convert.ToInt32(resultado)
                    End Using
                Next

                Dim noEncontrados = validos.Count - res.Actualizados
                If noEncontrados > 0 Then
                    res.Errores.Add(String.Format("{0} registro(s) no encontrados en la DBF", noEncontrados))
                End If

            End Using
        Catch ex As Exception
            Throw New Exception("Error al guardar: " & ex.Message)
        End Try

        Return res
    End Function

End Class