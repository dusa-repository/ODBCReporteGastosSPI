Imports System.Xml
Imports System.IO
Imports System.Data.SqlClient
Imports System.Data
Imports System.Net
Module Principal

    Sub Main()

        cargar_parametros()

        'If correrInterfaz() Then
        ' GERENCIA FALTA
        'cambiarEstado()
        importarCargo()
        importarUnidad()
        importarEmpleado()
        'End If

    End Sub

    Public conexionString As String = ""
    Dim lineaLogger As String
    Dim logger As StreamWriter
    Dim lineaLoggerE As String
    Dim loggerE As StreamWriter
    Dim Conn400 As ADODB.Connection
    Dim Rst400 As ADODB.Recordset


    Dim RstSQLAS As ADODB.Recordset
    Dim CmdSQLAS As ADODB.Command
    Dim ConnSQLAS As ADODB.Connection

    Dim cadenaSQL As String
    Dim cadenaAS400_DTA As String
    Dim cadenaAS400_CTL As String
    Dim cadenaAS400_SPI As String
    Dim unaVez As Boolean
    Dim server As String
    Dim database As String
    Dim uid As String
    Dim pwd As String

    Private Sub cambiarEstado()



        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand

        Try

            connSQL.ConnectionString = conexionString
            connSQL.Open()

            cmdSQL.Connection = connSQL

            If valido() Then
                cmdSQL.CommandText = "UPDATE  evaluacion   SET estado_evaluacion = '" + obtenerEstado() + "'  WHERE id_revision='" + buscarRevisionActiva() + "'"
                cmdSQL.ExecuteNonQuery()
                connSQL.Close()
            End If

            

        Catch oe As Exception
            Console.WriteLine(oe.StackTrace.ToString & "-" & oe.Message.ToString)
            escribirLog(oe.StackTrace.ToString & "-" & oe.Message.ToString, "(MATMAS) ")
        Finally

            logger.Close()
        End Try


    End Sub


    Private Function valido() As Boolean

        Dim esValido As Boolean
        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand
        connSQL.ConnectionString = conexionString
        connSQL.Open()

        Dim format As String = "dd/MM/yyyy"

        esValido = True

        Try

            cmdSQL.Connection = connSQL
            cmdSQL.CommandText = "select count(*) from fecha_validez_estado where convert(datetime,'" & DateTime.Now.ToString(format) & "',103)>=convert(datetime,fecha_desde,103) and convert(datetime,'" & DateTime.Now.ToString(format) & "',103)<=convert(datetime,fecha_hasta,103)  "
            Console.WriteLine(cmdSQL.CommandText)
            Dim lrdSQL As SqlDataReader = cmdSQL.ExecuteReader()

            While lrdSQL.Read()
                If lrdSQL.GetInt32(0) = 1 Then
                    esValido = True
                Else
                    esValido = False
                End If


            End While

            lrdSQL.Close()
            connSQL.Close()

        Catch oe As Exception
            escribirLog(oe.StackTrace.ToString & "-" & oe.Message.ToString, "(MATMAS) ")
        End Try
       

        Return esValido
    End Function


    Private Function obtenerEstado() As String

        Dim estado As String
        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand
        connSQL.ConnectionString = conexionString
        connSQL.Open()

        Dim format As String = "dd/MM/yyyy"

        Try

            cmdSQL.Connection = connSQL
            cmdSQL.CommandText = "select estado from fecha_validez_estado where convert(datetime,'" & DateTime.Now.ToString(format) & "',103)>=convert(datetime,fecha_desde,103) and convert(datetime,'" & DateTime.Now.ToString(format) & "',103)<=convert(datetime,fecha_hasta,103)  "
            Dim lrdSQL As SqlDataReader = cmdSQL.ExecuteReader()
            estado = ""
            While lrdSQL.Read()
                estado = lrdSQL.GetString(0)
            End While

            lrdSQL.Close()
            connSQL.Close()
        Catch oe As Exception
            escribirLog(oe.StackTrace.ToString & "-" & oe.Message.ToString, "(MATMAS) ")
        End Try

       

        Return estado
    End Function

    Private Function buscarRevisionActiva() As String

        Dim estado As String
        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand

        Try

            connSQL.ConnectionString = conexionString
            connSQL.Open()

            cmdSQL.Connection = connSQL
            cmdSQL.CommandText = "SELECT [id_revision] FROM [dusa_evaluacion].[dbo].[revision]   where estado_revision='ACTIVO'"
            Dim lrdSQL As SqlDataReader = cmdSQL.ExecuteReader()
            estado = "0"
            While lrdSQL.Read()
                estado = lrdSQL.GetInt32(0)
            End While

            lrdSQL.Close()
            connSQL.Close()

        Catch oe As Exception
            escribirLog(oe.StackTrace.ToString & "-" & oe.Message.ToString, "(MATMAS) ")
        End Try

       

        Return estado
    End Function


   
    Private Sub cargar_parametros()

        Dim host As String
        Dim database As String
        Dim user As String
        Dim password As String

        Dim conexion As New SqlConnection
        Dim comando As New SqlClient.SqlCommand

        Dim diccionario As New Dictionary(Of String, String)
        Dim xmldoc As New XmlDataDocument()
        Dim file_log_path As String

        Try
            file_log_path = Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "")
            If System.IO.File.Exists(file_log_path & "\log.log") Then
            Else
                Dim fs1 As FileStream = File.Create(file_log_path & "\log.log")
                fs1.Close()
            End If

            Try
                logger = New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\log.log", True)
            Catch ex As Exception

            End Try


            Dim fs As New FileStream(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\integrador.xml", FileMode.Open, FileAccess.Read)
            xmldoc = New XmlDataDocument()
            xmldoc.Load(fs)
            diccionario = obtenerNodosHijosDePadre("parametros", xmldoc)
            host = diccionario.Item("host")
            database = diccionario.Item("database")
            user = diccionario.Item("user")
            password = diccionario.Item("password")

            conexionString = "Data Source=" & host & ";Database=" & database & ";User ID=" & user & ";Password=" & password & ";"

        Catch oe As Exception
            escribirLog(oe.StackTrace.ToString & "-" & oe.Message.ToString, "(MATMAS) ")
        Finally

            logger.Close()
        End Try

    End Sub


    Public Sub escribirLog(ByVal mensaje As String, ByVal proceso As String)

        Dim time As DateTime = DateTime.Now
        Dim format As String = "dd/MM/yyyy HH:mm "



        Try


            lineaLogger = proceso & time.ToString(format) & ":" & mensaje & vbNewLine
            logger.WriteLine(lineaLogger)
            logger.Flush()

        Catch ex As Exception

            Try
                logger = New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\log.log", True)


                lineaLogger = proceso & time.ToString(format) & ":" & mensaje & vbNewLine
                logger.WriteLine(lineaLogger)
                logger.Flush()
            Catch ex1 As Exception


            End Try

        End Try


    End Sub

    Public Function obtenerNodosHijosDePadre(ByVal nombreNodoPadre As String, ByVal xmldoc As XmlDataDocument) As Dictionary(Of String, String)
        Dim diccionario As New Dictionary(Of String, String)
        Dim nodoPadre As XmlNodeList
        Dim i As Integer
        Dim h As Integer
        nodoPadre = xmldoc.GetElementsByTagName(nombreNodoPadre)
        For i = 0 To nodoPadre.Count - 1
            For h = 0 To nodoPadre(i).ChildNodes.Count - 1
                If Not diccionario.ContainsKey(nodoPadre(i).ChildNodes.Item(h).Name.Trim()) Then
                    diccionario.Add(nodoPadre(i).ChildNodes.Item(h).Name.Trim(), nodoPadre(i).ChildNodes.Item(h).InnerText.Trim())
                End If
            Next
        Next
        Return diccionario
    End Function

    Private Function correrInterfaz() As Boolean

        Dim correr As Boolean
        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand
        connSQL.ConnectionString = conexionString
        connSQL.Open()

        cmdSQL.Connection = connSQL
        cmdSQL.CommandText = "select interfaz from configuracion_general "
        Dim lrdSQL As SqlDataReader = cmdSQL.ExecuteReader()
        correr = False
        While lrdSQL.Read()
            If lrdSQL.GetString(0).Trim() = "SI" Then
                correr = True
            End If


        End While

        lrdSQL.Close()
        connSQL.Close()

        Return correr
    End Function


    Private Sub importarCargo()


        Try

            cadenaAS400_SPI = "DSN=SPI;uid=TRANSFTP;pwd=TRANSFTP;"
            Dim cnn400 As New Odbc.OdbcConnection(cadenaAS400_SPI)
            Dim rs400 As New Odbc.OdbcCommand("SELECT CIACGO, CODCGO, DESCGO FROM NMPP004 WHERE CIACGO IN (7,14) ", cnn400)
            Dim reader400 As Odbc.OdbcDataReader

            cnn400.Open()
            reader400 = rs400.ExecuteReader

            While reader400.Read()

                If existeCargo(Trim(reader400("CODCGO")), Trim(reader400("CIACGO"))) Then
                    guardarCargo(True, Trim(reader400("DESCGO")), Trim(reader400("CODCGO")), Trim(reader400("CIACGO")), obtenerIdEmpresa(Trim(reader400("CIACGO"))))
                Else
                    guardarCargo(False, Trim(reader400("DESCGO")), Trim(reader400("CODCGO")), Trim(reader400("CIACGO")), obtenerIdEmpresa(Trim(reader400("CIACGO"))))
                End If

                Console.WriteLine(reader400("CODCGO"))

            End While

            reader400.Close()
            cnn400.Close()

        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, "(MATMAS) ")
        End Try

       

    End Sub


    Private Function existeCargo(ByVal codigoCargo As String, ByVal codigoEmpresa As String) As Boolean

        Dim existe As Boolean
        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand

        Try

            connSQL.ConnectionString = conexionString
            connSQL.Open()

            cmdSQL.Connection = connSQL
            cmdSQL.CommandText = "select * from cargo where id_cargo_auxiliar='" + codigoCargo + "' and id_empresa_auxiliar='" + codigoEmpresa + "' "
            Dim lrdSQL As SqlDataReader = cmdSQL.ExecuteReader()
            existe = False
            While lrdSQL.Read()
                existe = True
            End While

            lrdSQL.Close()
            connSQL.Close()

        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, "(MATMAS) ")
        End Try

       

        Return existe
    End Function


    Private Sub guardarCargo(ByVal existe As Boolean, ByVal descripcion As String, ByVal id_cargo_auxiliar As String, ByVal id_empresa_auxiliar As String, ByVal id_empresa As String)


        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand

        Try

            connSQL.ConnectionString = conexionString
            connSQL.Open()

            cmdSQL.Connection = connSQL

            If existe Then
                cmdSQL.CommandText = "UPDATE [dbo].[cargo]   SET [descripcion] = '" + descripcion + "'  WHERE id_cargo_auxiliar='" + id_cargo_auxiliar + "' AND id_empresa_auxiliar='" + id_empresa_auxiliar + "'"
            Else
                cmdSQL.CommandText = "INSERT INTO [dbo].[cargo]([descripcion],[fecha_auditoria],[hora_auditoria],[id_cargo_auxiliar],[nomina],[id_empresa_auxiliar],[usuario],[id_empresa]) VALUES('" + descripcion + "',convert(datetime,'" & Format(Now(), "dd/MM/yyyy") & "',103),'" + Format(Now(), "hh:mm:ss") + "' ,'" + id_cargo_auxiliar + "',0,'" + id_empresa_auxiliar + "','interfaz','" + id_empresa + "')"
            End If

            cmdSQL.ExecuteNonQuery()
            connSQL.Close()

        Catch ex As Exception

        End Try

       


    End Sub

    Private Sub importarUnidad()

        cadenaAS400_SPI = "DSN=SPI;uid=TRANSFTP;pwd=TRANSFTP;"
        Dim cnn400 As New Odbc.OdbcConnection(cadenaAS400_SPI)
        Dim rs400 As New Odbc.OdbcCommand("SELECT CIADPT, CODDPT, DESDPT FROM NMPP019 WHERE CIADPT IN (7,14) ", cnn400)
        Dim reader400 As Odbc.OdbcDataReader

        Try
            cnn400.Open()
            reader400 = rs400.ExecuteReader

            While reader400.Read()

                If existeUnidad(Trim(reader400("CODDPT")), Trim(reader400("CIADPT"))) Then
                    guardarUnidad(True, Trim(reader400("DESDPT")), Trim(reader400("CODDPT")), Trim(reader400("CIADPT")), obtenerIdEmpresa(Trim(reader400("CIADPT"))))
                Else
                    guardarUnidad(False, Trim(reader400("DESDPT")), Trim(reader400("CODDPT")), Trim(reader400("CIADPT")), obtenerIdEmpresa(Trim(reader400("CIADPT"))))
                End If

                Console.WriteLine(reader400("CODDPT"))

            End While

            reader400.Close()
            cnn400.Close()


        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, "(MATMAS) ")
        End Try

       

    End Sub

    Private Function existeUnidad(ByVal codigoUnidad As String, ByVal codigoEmpresa As String) As Boolean

        Dim existe As Boolean
        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand

        Try
            connSQL.ConnectionString = conexionString
            connSQL.Open()

            cmdSQL.Connection = connSQL
            cmdSQL.CommandText = "select * from unidad_organizativa where id_unidad_organizativa_auxiliar='" + codigoUnidad + "' and id_empresa_auxiliar='" + codigoEmpresa + "' "
            Dim lrdSQL As SqlDataReader = cmdSQL.ExecuteReader()
            existe = False
            While lrdSQL.Read()
                existe = True
            End While

            lrdSQL.Close()
            connSQL.Close()
        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, "(MATMAS) ")
        End Try


       

        Return existe
    End Function

    Private Sub guardarUnidad(ByVal existe As Boolean, ByVal descripcion As String, ByVal id_unidad_auxiliar As String, ByVal id_empresa_auxiliar As String, ByVal id_empresa As String)


        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand


        Try

            connSQL.ConnectionString = conexionString
            connSQL.Open()

            cmdSQL.Connection = connSQL

            If existe Then
                cmdSQL.CommandText = "UPDATE [dbo].[unidad_organizativa]   SET [descripcion] = '" + descripcion + "'  WHERE id_unidad_organizativa_auxiliar='" + id_unidad_auxiliar + "' AND id_empresa_auxiliar='" + id_empresa_auxiliar + "'"
            Else
                cmdSQL.CommandText = "INSERT INTO [dbo].[unidad_organizativa]([id_gerencia],[id_usuario],[descripcion],[fecha_auditoria],[hora_auditoria],[nivel],[sub_nivel],[id_empresa_auxiliar],[id_unidad_organizativa_auxiliar],[usuario])     VALUES(16,1,convert(datetime,'" & Format(Now(), "dd/MM/yyyy") & "',103),'" + Format(Now(), "hh:mm:ss") + "',0,0,'" + id_empresa_auxiliar + "','" + id_unidad_auxiliar + "','interfaz')"
            End If

            cmdSQL.ExecuteNonQuery()
            connSQL.Close()


        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, "(MATMAS) ")
        End Try

       

    End Sub

    Private Sub importarEmpleado()

        cadenaAS400_SPI = "DSN=SPI;uid=TRANSFTP;pwd=TRANSFTP;"
        Dim cnn400 As New Odbc.OdbcConnection(cadenaAS400_SPI)
        Dim rs400 As New Odbc.OdbcCommand("SELECT CIAFIC,CODFIC,NOMFI1,APEFI1,CTGFI2,CEDMEXFIC,FECRET,DPTFIC, CGOFIC FROM NMPP060 WHERE  CIAFIC IN (7,14) AND CGOFIC NOT IN ('2AI001','2AI002') AND TPNFIC IN ('1106','2106') ", cnn400)
        Dim reader400 As Odbc.OdbcDataReader
        Dim insertar As Boolean
        Dim estadoEmpleado As String

        Try

            'loggerE = New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\ficha.log", True)


           

            limpiarError()

            cnn400.Open()
            reader400 = rs400.ExecuteReader

            While reader400.Read()

                Dim fichaAux = reader400("CODFIC")


                If reader400("CGOFIC").ToString.Substring(0, 1) <> "3" Then

                    insertar = True
                    estadoEmpleado = "INACTIVO"
                    If Trim(reader400("FECRET")) = "0" Then
                        estadoEmpleado = "ACTIVO"
                    Else
                        estadoEmpleado = "INACTIVO"
                    End If

                    If estadoEmpleado = "ACTIVO" Then


                        Dim grado As Integer
                        If Replace(Trim(reader400("CTGFI2")), "G", "") = "" Then
                            grado = 0
                        Else
                            Try
                                grado = Int32.Parse(Replace(Trim(reader400("CTGFI2")), "G", ""))
                            Catch ex As Exception
                                grado = 0
                            End Try
                        End If

                        Dim ficha_supervisor As String

                        If Trim(reader400("CEDMEXFIC")) = "" Then
                            ficha_supervisor = 0
                        Else
                            ficha_supervisor = Trim(reader400("CEDMEXFIC"))
                        End If

                        If grado = 0 Then
                            insertarError(Trim(reader400("CODFIC")), Trim(reader400("NOMFI1")) + " " + Trim(reader400("APEFI1")), "GRADO DEL EMPLEADO NO VALIDO")
                            insertar = False
                        End If

                        If ficha_supervisor = "0" Then
                            insertarError(Trim(reader400("CODFIC")), Trim(reader400("NOMFI1")) + " " + Trim(reader400("APEFI1")), "FICHA DEL SUPERVISOR  DEL EMPLEADO NO VALIDO")
                            insertar = False
                        End If

                        If insertar Then
                            If existeEmpleado(Trim(reader400("CODFIC")), Trim(reader400("CIAFIC"))) Then
                                guardarEmpleado(True, Trim(reader400("CIAFIC")), Trim(reader400("DPTFIC")), Trim(reader400("CGOFIC")), Trim(reader400("NOMFI1")) + " " + Trim(reader400("APEFI1")), Trim(reader400("NOMFI1")), Trim(reader400("APEFI1")), Trim(reader400("CODFIC")), ficha_supervisor, grado, estadoEmpleado)
                            Else
                                If estadoEmpleado = "ACTIVO" Then
                                    guardarEmpleado(False, Trim(reader400("CIAFIC")), Trim(reader400("DPTFIC")), Trim(reader400("CGOFIC")), Trim(reader400("NOMFI1")) + " " + Trim(reader400("APEFI1")), Trim(reader400("NOMFI1")), Trim(reader400("APEFI1")), Trim(reader400("CODFIC")), ficha_supervisor, grado, estadoEmpleado)
                                End If
                            End If
                        End If

                        Console.WriteLine(reader400("CODFIC"))
                        'lineaLoggerE = lineaLoggerE + "'" & Trim(reader400("CODFIC")) & "',"


                    End If
                End If

            End While

            'loggerE.WriteLine(lineaLoggerE)
            'loggerE.Flush()
            'loggerE.Close()

            reader400.Close()
            cnn400.Close()

        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, "(MATMAS) ")
        End Try

       

    End Sub

    Private Function existeEmpleado(ByVal ficha As String, ByVal codigoEmpresa As String) As Boolean

        Dim existe As Boolean
        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand

        Try

            connSQL.ConnectionString = conexionString
            connSQL.Open()

            cmdSQL.Connection = connSQL
            'cmdSQL.CommandText = "select * from empleado where ficha='" + ficha + "' and id_empresa_auxiliar='" + codigoEmpresa + "' "
            cmdSQL.CommandText = "select * from empleado where ficha='" + ficha + "'"
            Dim lrdSQL As SqlDataReader = cmdSQL.ExecuteReader()
            existe = False
            While lrdSQL.Read()
                existe = True
            End While

            lrdSQL.Close()
            connSQL.Close()

        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, "(MATMAS) ")
        End Try

        

        Return existe
    End Function

    Private Sub guardarEmpleado(ByVal existe As Boolean, ByVal id_empresa_auxiliar As String, ByVal id_unidad_auxiliar As String, ByVal id_cargo_auxiliar As String, ByVal nombre As String, ByVal nombreS As String, ByVal apellido As String, ByVal ficha As String, ByVal ficha_supervisor As String, ByVal grado As String, ByVal estado As String)


        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand

        Try

            connSQL.ConnectionString = conexionString
            connSQL.Open()

            cmdSQL.Connection = connSQL

            'If existe Then
            ' cmdSQL.CommandText = "UPDATE [dbo].[empleado]   SET [nombre] = '" + nombre + "',[ficha_supervisor]='" + ficha_supervisor + "',[grado_auxiliar]=" + grado + ",id_unidad_organizativa_auxiliar='" + id_unidad_auxiliar + "', id_cargo_auxiliar= '" + id_cargo_auxiliar + "' ,[id_cargo]='" + obtenerIdCargo(Trim(id_cargo_auxiliar)).ToString() + "',[id_unidad_organizativa]='" + obtenerIdUnidad(Trim(id_unidad_auxiliar)).ToString() + "'  WHERE ficha='" + ficha + "' AND id_empresa_auxiliar='" + id_empresa_auxiliar + "'"
            ' Else
            ' cmdSQL.CommandText = "INSERT INTO [dbo].[empleado]([id_empresa],[id_cargo],[id_unidad_organizativa],[nombre],[ficha],[fecha_auditoria],[hora_auditoria],[ficha_supervisor],[grado_auxiliar],[usuario],[id_cargo_auxiliar],[id_unidad_organizativa_auxiliar],[id_empresa_auxiliar]) VALUES ('" + obtenerIdEmpresa(Trim(id_empresa_auxiliar)).ToString() + "','" + obtenerIdCargo(Trim(id_cargo_auxiliar)).ToString() + "','" + obtenerIdUnidad(Trim(id_unidad_auxiliar)).ToString() + "','" + nombre + "','" + ficha + "',convert(datetime,'" & Format(Now(), "dd/MM/yyyy") & "',103),'" + Format(Now(), "hh:mm:ss") + "','" + ficha_supervisor + "'," + grado + ",'interfaz','" + id_cargo_auxiliar + "','" + id_unidad_auxiliar + "','" + id_empresa_auxiliar + "')"
            ' End If


            If existe Then
                cmdSQL.CommandText = "UPDATE [dbo].[empleado]   SET [nombre] = '" + nombre + "',[ficha_supervisor]='" + ficha_supervisor + "',[grado_auxiliar]=" + grado + ",id_unidad_organizativa_auxiliar='" + id_unidad_auxiliar + "', id_cargo_auxiliar= '" + id_cargo_auxiliar + "' ,[id_cargo]='" + obtenerIdCargo(Trim(id_cargo_auxiliar), Trim(id_empresa_auxiliar)).ToString() + "',[id_unidad_organizativa]='" + obtenerIdUnidad(Trim(id_unidad_auxiliar), Trim(id_empresa_auxiliar)).ToString() + "',estado='" + estado + "'  WHERE ficha='" + ficha + "' AND id_empresa_auxiliar='" + id_empresa_auxiliar + "'"
            Else
                cmdSQL.CommandText = "INSERT INTO [dbo].[empleado]([id_empresa],[id_cargo],[id_unidad_organizativa],[nombre],[ficha],[fecha_auditoria],[hora_auditoria],[ficha_supervisor],[grado_auxiliar],[usuario],[id_cargo_auxiliar],[id_unidad_organizativa_auxiliar],[id_empresa_auxiliar],[estado]) VALUES ('" + obtenerIdEmpresa(Trim(id_empresa_auxiliar)).ToString() + "','" + obtenerIdCargo(Trim(id_cargo_auxiliar), Trim(id_empresa_auxiliar)).ToString() + "','" + obtenerIdUnidad(Trim(id_unidad_auxiliar), Trim(id_empresa_auxiliar)).ToString() + "','" + nombre + "','" + ficha + "',convert(datetime,'" & Format(Now(), "dd/MM/yyyy") & "',103),'" + Format(Now(), "hh:mm:ss") + "','" + ficha_supervisor + "'," + grado + ",'interfaz','" + id_cargo_auxiliar + "','" + id_unidad_auxiliar + "','" + id_empresa_auxiliar + "','" + estado + "')"
            End If

            cmdSQL.ExecuteNonQuery()


            If existe = False Then
                cmdSQL.CommandText = "INSERT INTO [dbo].[usuario_seguridad]([login],[apellido],[email],[estado],[fecha_auditoria],[hora_auditoria],[nombre],[password],[usuario_auditoria])     VALUES ('" + ficha + "','" + apellido + "','" + Left(Trim(nombreS), 1) + Trim(apellido) + "@dusa.com.ve',1,convert(datetime,'" & Format(Now(), "dd/MM/yyyy") & "',103),'" + Format(Now(), "hh:mm:ss") + "','" + nombreS + "','" + ficha + "','admin')"
                cmdSQL.ExecuteNonQuery()

                cmdSQL.CommandText = "INSERT INTO [dbo].[usuario]([id_rol],[nombre],[apellido],[cedula],[email],[login],[password],[fecha_auditoria],[hora_auditoria],[estado],[ficha])     VALUES (4,'" + nombreS + "','" + apellido + "','" + ficha + "','','" + ficha + "','" + ficha + "',convert(datetime,'" & Format(Now(), "dd/MM/yyyy") & "',103),'" + Format(Now(), "hh:mm:ss") + "',1,'" + ficha + "')"
                cmdSQL.ExecuteNonQuery()

            End If

            cmdSQL.CommandText = "insert into grupo_usuario select login, 4 from usuario_seguridad where cast(login as varchar) not in (select id_usuario from grupo_usuario) and cast(login as varchar)<>'admin'"
            cmdSQL.ExecuteNonQuery()

            connSQL.Close()

        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, "(MATMAS) ")
        End Try


       


    End Sub

    Private Sub insertarError(ByVal ficha As String, ByVal nombre As String, ByVal errorD As String)

        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand

        Try

            connSQL.ConnectionString = conexionString
            connSQL.Open()

            cmdSQL.Connection = connSQL
            cmdSQL.CommandText = "INSERT INTO [dbo].[error_interfaz] ([fecha],[hora],[ficha],[nombre],[error]) VALUES (convert(datetime,'" & Format(Now(), "dd/MM/yyyy") & "',103),'" + Format(Now(), "hh:mm:ss") + "','" + ficha + "','" + nombre + "','" + errorD + "')"
            cmdSQL.ExecuteNonQuery()
            connSQL.Close()

        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, "(MATMAS) ")
        End Try

       

    End Sub

    Private Sub limpiarError()

        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand

        Try
            connSQL.ConnectionString = conexionString
            connSQL.Open()

            cmdSQL.Connection = connSQL
            cmdSQL.CommandText = "DELETE FROM [dbo].[error_interfaz] "
            cmdSQL.ExecuteNonQuery()
            connSQL.Close()

        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, "(MATMAS) ")
        End Try

       

    End Sub

    Private Function obtenerIdEmpresa(ByVal id_empresa_auxiliar As String) As Integer

        Dim idEmpresa As Integer
        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand
        Try

            connSQL.ConnectionString = conexionString
            connSQL.Open()

            cmdSQL.Connection = connSQL
            cmdSQL.CommandText = "select id_empresa from empresa where id_empresa_auxiliar='" + id_empresa_auxiliar + "' "
            Dim lrdSQL As SqlDataReader = cmdSQL.ExecuteReader()

            idEmpresa = "00"

            While lrdSQL.Read()
                idEmpresa = lrdSQL.GetInt32(0)
            End While

            lrdSQL.Close()
            connSQL.Close()


        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, "(MATMAS) ")
        End Try

        
        Return idEmpresa
    End Function

    Private Function obtenerIdCargo(ByVal id_cargo_auxiliar As String, ByVal id_empresa_auxiliar As String) As Integer

        Dim idCargo As Integer
        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand

        Try

            connSQL.ConnectionString = conexionString
            connSQL.Open()

            cmdSQL.Connection = connSQL
            cmdSQL.CommandText = "select id_cargo from cargo where id_cargo_auxiliar='" + id_cargo_auxiliar + "' and id_empresa_auxiliar='" + id_empresa_auxiliar + "' "
            Dim lrdSQL As SqlDataReader = cmdSQL.ExecuteReader()

            idCargo = "00"

            While lrdSQL.Read()
                idCargo = lrdSQL.GetInt32(0)
            End While

            lrdSQL.Close()
            connSQL.Close()

        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, "(MATMAS) ")
        End Try

       

        Return idCargo
    End Function

    Private Function obtenerIdUnidad(ByVal id_unidad_auxiliar As String, ByVal id_empresa_auxiliar As String) As Integer

        Dim idUnidad As Integer
        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand

        Try

            connSQL.ConnectionString = conexionString
            connSQL.Open()

            cmdSQL.Connection = connSQL
            cmdSQL.CommandText = "select id_unidad_organizativa from unidad_organizativa where id_unidad_organizativa_auxiliar='" + id_unidad_auxiliar + "'  and id_empresa_auxiliar='" + id_empresa_auxiliar + "' "
            Dim lrdSQL As SqlDataReader = cmdSQL.ExecuteReader()

            idUnidad = "00"

            While lrdSQL.Read()
                idUnidad = lrdSQL.GetInt32(0)
            End While

            lrdSQL.Close()
            connSQL.Close()

        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, "(MATMAS) ")
        End Try

        

        Return idUnidad
    End Function

End Module
