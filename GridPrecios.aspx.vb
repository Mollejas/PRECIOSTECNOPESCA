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
    ' GUARDAR PRECIOS (UPDATE directo por registro)
    ' Usa AddWithValue igual que BuscarPrecio (que sí funciona)
    ' ===========================
    <WebMethod>
    Public Shared Function GuardarPrecios(items As List(Of ItemGuardar)) As ResultGuardar
        Dim res As New ResultGuardar() With {
            .Actualizados = 0,
            .Total = If(items IsNot Nothing, items.Count, 0),
            .Errores = New List(Of String)
        }

        If items Is Nothing OrElse items.Count = 0 Then
            res.Errores.Add("No se recibieron items del navegador")
            Return res
        End If

        Try
            Using conn As New OleDbConnection(ConnStr())
                conn.Open()

                For Each item In items
                    Dim precio As Decimal = 0
                    If Not Decimal.TryParse(item.Precio,
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture,
                            precio) Then
                        res.Errores.Add(String.Format("{0}: precio inválido '{1}'", item.Clave, item.Precio))
                        Continue For
                    End If

                    Dim sql As String = String.Format(
                        "UPDATE {0} SET APRPRC = ? WHERE APRCLAVE = ? AND APRLISTA = ?",
                        DBF_TABLE)

                    Using cmd As New OleDbCommand(sql, conn)
                        cmd.Parameters.AddWithValue("@precio", CDbl(precio))
                        cmd.Parameters.AddWithValue("@clave", item.Clave.Trim().PadRight(25))
                        cmd.Parameters.AddWithValue("@lista", item.Lista.Trim().PadRight(3))

                        Dim affected As Integer = cmd.ExecuteNonQuery()
                        If affected > 0 Then
                            res.Actualizados += 1
                        Else
                            res.Errores.Add(String.Format("No encontrado: [{0}] Lista [{1}]", item.Clave.Trim(), item.Lista.Trim()))
                        End If
                    End Using
                Next
            End Using
        Catch ex As Exception
            Throw New Exception("Error al guardar: " & ex.Message & " | StackTrace: " & ex.StackTrace)
        End Try

        Return res
    End Function

End Class