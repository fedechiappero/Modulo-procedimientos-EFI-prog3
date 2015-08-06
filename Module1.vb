Imports MySql.Data.MySqlClient


Module Module1

    'declaro cadena de conexion
    Dim cone As MySqlConnection = New MySqlConnection("server=localhost;user id=root;password=root;database=programacion3efi")


    Class CFormaPago

        Public Sub Agregar(ByVal nombre As String)
            Try
                cone.Open()
                Using cone
                    'defino parametros
                    Dim prmNombre As New MySqlParameter("@FoPaNomb", MySqlDbType.Text)
                    Dim prmEstado As New MySqlParameter("@FoPaEsta", MySqlDbType.Int32)
                    'asigno valores a los parametros
                    prmNombre.Value = StrConv(nombre, 3)
                    prmEstado.Value = 1
                    'defino el comando sql
                    Dim instruccion As New MySqlCommand("INSERT INTO formapago (FoPaNomb,FoPaEsta) VALUES (@FoPaNomb,@FoPaEsta)", cone)
                    'agrego paramtros al comando
                    instruccion.Parameters.Add(prmNombre)
                    instruccion.Parameters.Add(prmEstado)
                    'ejecucion del comando
                    instruccion.ExecuteNonQuery()
                    MessageBox.Show("Agregado")
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub Modificar(ByVal idTipo As String, ByVal nombre As String, ByVal estado As String)
            Try
                cone.Open()
                Using cone
                    Dim prmId As New MySqlParameter("@FoPaCodi", MySqlDbType.Int32)
                    Dim prmNombre As New MySqlParameter("@FoPaNomb", MySqlDbType.Text)
                    Dim prmEstado As New MySqlParameter("@FoPaEsta", MySqlDbType.Int32)

                    prmId.Value = CType(idTipo, Int32)
                    prmNombre.Value = StrConv(nombre, 3)
                    If estado = "Activo" Then 'si estado = activo asigno 1, sino 0
                        prmEstado.Value = 1
                    Else
                        prmEstado.Value = 0
                    End If

                    Dim instruccion As New MySqlCommand("UPDATE formapago SET FoPaNomb=@FoPaNomb,FoPaEsta=@FoPaEsta WHERE FoPaCodi = @FoPaCodi", cone)

                    instruccion.Parameters.Add(prmId)
                    instruccion.Parameters.Add(prmNombre)
                    instruccion.Parameters.Add(prmEstado)

                    instruccion.ExecuteNonQuery()
                End Using
                MessageBox.Show("Modificado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        'eliminacion fisica
        Public Sub EliminarF(ByVal idTipo As String)
            Try
                cone.Open()
                Using cone
                    Dim prmId As New MySqlParameter("@FoPaCodi", MySqlDbType.Int32)

                    prmId.Value = CType(idTipo, Int32)

                    Dim instruccion As New MySqlCommand("DELETE FROM formapago WHERE FoPaCodi = @FoPaCodi", cone)

                    instruccion.Parameters.Add(prmId)

                    instruccion.ExecuteNonQuery()
                End Using
                MessageBox.Show("Eliminado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        'eliminacion logica
        Public Sub EliminarL(ByVal idFormaPago As String)
            Try
                cone.Open()
                Using cone
                    Dim prmId As New MySqlParameter("@FoPaCodi", MySqlDbType.Int32)
                    Dim prmEstado As New MySqlParameter("@FoPaEsta", MySqlDbType.Int16)

                    prmId.Value = CType(idFormaPago, Int32)
                    prmEstado.Value = 0

                    Dim instruccion As New MySqlCommand("UPDATE formapago SET FoPaEsta=@FoPaEsta WHERE FoPaCodi = @FoPaCodi", cone)

                    instruccion.Parameters.Add(prmId)
                    instruccion.Parameters.Add(prmEstado)

                    instruccion.ExecuteNonQuery()
                End Using
                MessageBox.Show("Eliminado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

    End Class


    Class CTipoHabitacion

        Public Sub Agregar(ByVal nombre As String, ByVal cantidad As String)
            Try
                cone.Open()
                Using cone

                    Dim prmNombre As New MySqlParameter("@TiHaNomb", MySqlDbType.Text)
                    Dim prmCantidad As New MySqlParameter("@TiHaCant", MySqlDbType.Int32)

                    prmNombre.Value = StrConv(nombre, 3)
                    prmCantidad.Value = CType(cantidad, Int32)

                    Dim instruccion As New MySqlCommand("INSERT INTO tipohabitacion (TiHaNomb,TiHaCant) VALUES (@TiHaNomb,@TiHaCant)", cone)

                    instruccion.Parameters.Add(prmNombre)
                    instruccion.Parameters.Add(prmCantidad)

                    instruccion.ExecuteNonQuery()
                    MessageBox.Show("Agregado")
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub Modificar(ByVal idTipo As String, ByVal nombre As String, ByVal cantidad As String)
            Try
                cone.Open()
                Using cone
                    Dim prmId As New MySqlParameter("@TiHaCodi", MySqlDbType.Int32)
                    Dim prmNombre As New MySqlParameter("@TiHaNomb", MySqlDbType.Text)
                    Dim prmCantidad As New MySqlParameter("@TiHaCant", MySqlDbType.Int32)

                    prmId.Value = CType(idTipo, Int32)
                    prmNombre.Value = StrConv(nombre, 3)
                    prmCantidad.Value = CType(cantidad, Int32)

                    Dim instruccion As New MySqlCommand("UPDATE tipohabitacion SET TiHaNomb=@TiHaNomb,TiHaCant=@TiHaCant WHERE TiHaCodi=@TiHaCodi", cone)

                    instruccion.Parameters.Add(prmId)
                    instruccion.Parameters.Add(prmNombre)
                    instruccion.Parameters.Add(prmCantidad)

                    instruccion.ExecuteNonQuery()
                End Using
                MessageBox.Show("Modificado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

    End Class


    Class CHabitacion

        Public Sub Agregar(ByVal numero As String, ByVal disponible As String, ByVal llena As String, ByVal tipo As String, ByVal limpia As String, ByVal estado As String)
            Try
                cone.Open()
                Using cone

                    Dim prmTipo As New MySqlParameter("@HabiTipo", MySqlDbType.Int32)
                    Dim prmNumero As New MySqlParameter("@HabiNume", MySqlDbType.Int32)
                    Dim prmDisponible As New MySqlParameter("@HabiDisp", MySqlDbType.Int16)
                    Dim prmLlena As New MySqlParameter("@HabiLlen", MySqlDbType.Int16)
                    Dim prmLimpia As New MySqlParameter("@HabiLimp", MySqlDbType.Int16)
                    Dim prmEstado As New MySqlParameter("@HabiEsta", MySqlDbType.Int16)

                    prmTipo.Value = CType(tipo, Int32)
                    prmNumero.Value = CType(numero, Int32)
                    prmDisponible.Value = CType(disponible, Int16)
                    prmLlena.Value = CType(llena, Int16)
                    prmLimpia.Value = CType(limpia, Int16)
                    prmEstado.Value = CType(estado, Int16)

                    Dim instruccion As New MySqlCommand("INSERT INTO habitacion (HabiTipo,HabiNume,HabiDisp,HabiLlen,HabiLimp,TariEsta) VALUES (@HabiTipo,@HabiNume,@HabiDisp,@HabiLlen,@HabiLimp,@TariEsta)", cone)

                    instruccion.Parameters.Add(prmTipo)
                    instruccion.Parameters.Add(prmEstado)
                    instruccion.Parameters.Add(prmDisponible)
                    instruccion.Parameters.Add(prmLlena)
                    instruccion.Parameters.Add(prmLimpia)
                    instruccion.Parameters.Add(prmEstado)

                    instruccion.ExecuteNonQuery()
                    MessageBox.Show("Agregada")
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub Modificar(ByVal idHabitacion As String, ByVal numero As String, ByVal disponible As String, ByVal Llena As String, ByVal tipo As String, ByVal limpia As String, ByVal estado As String)
            Try
                cone.Open()
                Using cone
                    Dim prmId As New MySqlParameter("@HabiCodi", MySqlDbType.Int32)
                    Dim prmTipo As New MySqlParameter("@HabiTipo", MySqlDbType.Int32)
                    Dim prmNumero As New MySqlParameter("@HabiNume", MySqlDbType.Int16)
                    Dim prmDisponible As New MySqlParameter("@HabiDisp", MySqlDbType.Int16)
                    Dim prmLlena As New MySqlParameter("@HabiLlen", MySqlDbType.Int16)
                    Dim prmLimpia As New MySqlParameter("@HabiEsta", MySqlDbType.Int16)
                    Dim prmEstado As New MySqlParameter("@HabiEsta", MySqlDbType.Int16)

                    prmId.Value = CType(idHabitacion, Int32)
                    prmNumero.Value = CType(numero, Int16)
                    prmDisponible.Value = CType(disponible, Int16)
                    prmLlena.Value = CType(Llena, Int16)
                    prmTipo.Value = CType(tipo, Int32)
                    prmLimpia.Value = CType(limpia, Int16)
                    prmEstado.Value = CType(estado, Int16)

                    Dim instruccion As New MySqlCommand("UPDATE habitacion SET HabiTipo=@HabiTipo,HabiNume=@HabiNume,HabiDisp=@HabiDisp,HabiLlen=@HabiLlen,HabiLimp,HabiEsta=@HabiEsta WHERE HabiCodi = @HabiCodi", cone)

                    instruccion.Parameters.Add(prmId)
                    instruccion.Parameters.Add(prmEstado)
                    instruccion.Parameters.Add(prmDisponible)
                    instruccion.Parameters.Add(prmLlena)
                    instruccion.Parameters.Add(prmTipo)
                    instruccion.Parameters.Add(prmLimpia)
                    instruccion.Parameters.Add(prmEstado)

                    instruccion.ExecuteNonQuery()
                End Using
                MessageBox.Show("Modificado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        'eliminacion fisica
        Public Sub EliminarF(ByVal idHabitacion As String)
            Try
                cone.Open()
                Using cone
                    Dim prmId As New MySqlParameter("@HabiCodi", MySqlDbType.Int32)

                    prmId.Value = CType(idHabitacion, Int32)

                    Dim instruccion As New MySqlCommand("DELETE FROM habitacion WHERE HabiCodi = @HabiCodi", cone)

                    instruccion.Parameters.Add(prmId)

                    instruccion.ExecuteNonQuery()
                End Using
                MessageBox.Show("Eliminado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        'eliminacion logica
        Public Sub EliminarL(ByVal idhabitacion As String, ByVal estado As String)
            Try
                cone.Open()
                Using cone
                    Dim prmId As New MySqlParameter("@HabiCodi", MySqlDbType.Int32)
                    Dim prmEstado As New MySqlParameter("@HabiEsta", MySqlDbType.Int16)

                    prmId.Value = CType(idhabitacion, Int32)
                    prmEstado.Value = CType(estado, Int16)

                    Dim instruccion As New MySqlCommand("UPDATE habitacion SET HabiEsta=@HabiEsta WHERE HabiCodi = @HabiCodi", cone)

                    instruccion.Parameters.Add(prmId)
                    instruccion.Parameters.Add(prmEstado)

                    instruccion.ExecuteNonQuery()
                End Using
                MessageBox.Show("Eliminado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub RecuperarLimpia(ByRef label As Label, ByVal numero As String)
            Try
                cone.Open()
                Using cone
                    Dim prmHabitacion As New MySqlParameter("@HabiNume", MySqlDbType.Text)

                    prmHabitacion.Value = CType(numero, String)

                    Dim recuperar As New MySqlCommand("SELECT HabiLimp FROM habitacion WHERE HabiNume=@HabiNume")
                    Dim res = recuperar.ExecuteScalar()

                    If res = 1 Then
                        label.Text = "Limpia"
                    ElseIf res = 0 Then
                        label.Text = "Sucia"
                    End If

                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub ModificarLimpia(ByVal idHabitacion As String, ByVal estado As String)
            Try
                cone.Open()
                Using cone
                    Dim prmHabitacion As New MySqlParameter("@HabiCodi", MySqlDbType.Int32)
                    Dim prmEstado As New MySqlParameter("@HabiLimp", MySqlDbType.Int16)

                    prmHabitacion.Value = CType(idHabitacion, Int32)
                    prmEstado.Value = CType(estado, Int16)

                    Dim instruccion As New MySqlCommand("UPDATE habitacion SET HabiLimp=@HabiLimp WHERE HabiCodi=@HabiCodi", cone)

                    instruccion.Parameters.Add(prmHabitacion)
                    instruccion.Parameters.Add(prmEstado)

                    instruccion.ExecuteNonQuery()
                End Using
                MessageBox.Show("Modificado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub

        Public Sub ModifiarDisponible(ByVal idHabitacion As String, ByVal disponible As String)
            Try
                cone.Open()
                Using cone
                    Dim prmHabitacion As New MySqlParameter("@HabiCodi", MySqlDbType.Int32)
                    Dim prmDisponible As New MySqlParameter("@HabiLimp", MySqlDbType.Int16)

                    prmHabitacion.Value = CType(idHabitacion, Int32)
                    prmDisponible.Value = CType(disponible, Int16)

                    Dim instruccion As New MySqlCommand("UPDATE habitacion SET HabiDisp=@HabiDisp WHERE HabiCodi=@HabiCodi", cone)

                    instruccion.Parameters.Add(prmHabitacion)
                    instruccion.Parameters.Add(prmDisponible)

                    instruccion.ExecuteNonQuery()
                End Using
                MessageBox.Show("Modificado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub

    End Class


    Class CAlquiler

        Public Sub Agregar(ByVal idReserva As String, ByVal idFormaPago As String, ByVal mantenimiento As String)
            Try
                cone.Open()
                Using cone

                    Dim prmReserva As New MySqlParameter("@ReseCodi", MySqlDbType.Int32)
                    Dim prmFormaPago As New MySqlParameter("@FoPaCodi", MySqlDbType.Int32)
                    Dim prmHabilitacionMantenimiento As New MySqlParameter("@AlquMant", MySqlDbType.Int16)

                    prmReserva.Value = CType(idReserva, Int32)
                    prmFormaPago.Value = CType(idFormaPago, Int32)
                    prmHabilitacionMantenimiento.Value = CType(mantenimiento, Int16)

                    Dim instruccion As New MySqlCommand("INSERT INTO alquiler (ReseCodi,FoPaCodi,AlquMant) VALUES (@ReseCodi,@FoPaCodi@AlquMant)", cone)

                    instruccion.Parameters.Add(prmReserva)
                    instruccion.Parameters.Add(prmFormaPago)
                    instruccion.Parameters.Add(prmHabilitacionMantenimiento)

                    instruccion.ExecuteNonQuery()
                    MessageBox.Show("Agregado")
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub Modificar(ByVal idReserva As String, ByVal idFormaPago As String, ByVal mantenimiento As String)
            Try
                cone.Open()
                Using cone
                    Dim prmReserva As New MySqlParameter("@ReseCodi", MySqlDbType.Int32)
                    Dim prmFormaPago As New MySqlParameter("@FoPaNomb", MySqlDbType.Int32)
                    Dim prmHabilitacionMantenimiento As New MySqlParameter("@AlquMant", MySqlDbType.Int16)

                    prmReserva.Value = CType(idReserva, Int32)
                    prmFormaPago.Value = CType(idFormaPago, Int32)
                    prmHabilitacionMantenimiento.Value = CType(mantenimiento, Int16)

                    Dim instruccion As New MySqlCommand("UPDATE alquiler SET FoPaCodi=@FoPaCodi, AlquMant=@AlquMant WHERE ReseCodi=@ReseCodi", cone)

                    instruccion.Parameters.Add(prmReserva)
                    instruccion.Parameters.Add(prmFormaPago)
                    instruccion.Parameters.Add(prmHabilitacionMantenimiento)

                    instruccion.ExecuteNonQuery()
                End Using
                MessageBox.Show("Modificado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub CheckOut(ByVal idReserva As String)
            Try
                cone.Open()
                Using cone
                    Dim prmReserva As New MySqlParameter("@ReseCodi", MySqlDbType.Int32)
                    Dim prmEstado As New MySqlParameter("@AlquEsta", MySqlDbType.Int16)

                    prmReserva.Value = CType(idReserva, Int32)
                    prmEstado.Value = 0

                    Dim instruccion As New MySqlCommand("UPDATE alquiler SET AlquEsta=@AlquEsta WHERE ReseCodi=@ReseCodi", cone)

                    instruccion.Parameters.Add(prmReserva)
                    instruccion.Parameters.Add(prmEstado)

                    instruccion.ExecuteNonQuery()
                End Using
                MessageBox.Show("Modificado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub


    End Class


    Class CReserva

        Public Sub Agregar(ByVal persona As String, ByVal habitacion As String, ByVal fechaIn As String, ByVal fechaOut As String, ByVal monto As String)
            Try
                cone.Open()
                Using cone

                    Dim prmIdPersona As New MySqlParameter("@PersCodi", MySqlDbType.Int32)
                    Dim prmIdHabitacion As New MySqlParameter("@HabiCodi", MySqlDbType.Int32)
                    Dim prmFechaIn As New MySqlParameter("@ReseFeIn", MySqlDbType.Date)
                    Dim prmFechaOut As New MySqlParameter("@ReseFeSa", MySqlDbType.Date)
                    Dim prmMonto As New MySqlParameter("@ReseMont", MySqlDbType.Int64)

                    prmIdPersona.Value = CType(persona, Int32)
                    prmIdHabitacion.Value = CType(habitacion, Int32)
                    prmFechaIn.Value = CType(fechaIn, Date)
                    prmFechaOut.Value = CType(fechaOut, Date)
                    prmMonto.Value = CType(monto, Int64)

                    Dim instruccion As New MySqlCommand("INSERT INTO reserva (PersCodi,HabiCodi,ReseFeIn,ReseFeSa,ReseMont) VALUES (@PersCodi,@HabiCodi,@ReseFeIn,@ReseFeSa,@ReseMont)", cone)

                    instruccion.Parameters.Add(prmIdPersona)
                    instruccion.Parameters.Add(prmIdHabitacion)
                    instruccion.Parameters.Add(prmFechaIn)
                    instruccion.Parameters.Add(prmFechaOut)
                    instruccion.Parameters.Add(prmMonto)

                    instruccion.ExecuteNonQuery()
                    MessageBox.Show("Agregada")
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub Modificar(ByVal idReserva As String, ByVal persona As String, ByVal habitacion As String, ByVal fechaIn As String, ByVal fechaOut As String, ByVal monto As String)
            Try
                cone.Open()
                Using cone
                    Dim prmIdReserva As New MySqlParameter("@ReseCodi", MySqlDbType.Int32)
                    Dim prmIdPersona As New MySqlParameter("@PersCodi", MySqlDbType.Int32)
                    Dim prmIdHabitacion As New MySqlParameter("@HabiCodi", MySqlDbType.Int32)
                    Dim prmFechaIn As New MySqlParameter("@ReseFeIn", MySqlDbType.Date)
                    Dim prmFechaOut As New MySqlParameter("@ReseFeSa", MySqlDbType.Date)
                    Dim prmMonto As New MySqlParameter("@ReseMont", MySqlDbType.Int64)

                    prmIdReserva.Value = CType(idReserva, Int32)
                    prmIdPersona.Value = CType(persona, Int32)
                    prmIdHabitacion.Value = CType(habitacion, Int32)
                    prmFechaIn.Value = CType(fechaIn, Date)
                    prmFechaOut.Value = CType(fechaOut, Date)
                    prmMonto.Value = CType(monto, Int64)

                    Dim instruccion As New MySqlCommand("UPDATE reserva SET PersCodi=@PersCodi,HabiCodi=@HabiCodi,ReseFeIn=@ReseFeIn,ReseFeSa=@ReseFeSa,ReseMont=@ReseMont WHERE ReseCodi = @ReseCodi", cone)

                    instruccion.Parameters.Add(prmIdReserva)
                    instruccion.Parameters.Add(prmIdPersona)
                    instruccion.Parameters.Add(prmIdHabitacion)
                    instruccion.Parameters.Add(prmFechaIn)
                    instruccion.Parameters.Add(prmFechaOut)
                    instruccion.Parameters.Add(prmMonto)

                    instruccion.ExecuteNonQuery()
                End Using
                MessageBox.Show("Modificado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        'eliminacion logica
        Public Sub EliminarL(ByVal idhabitacion As String, ByVal estado As String)
            Try
                cone.Open()
                Using cone
                    Dim prmId As New MySqlParameter("@HabiCodi", MySqlDbType.Int32)
                    Dim prmEstado As New MySqlParameter("@HabiEsta", MySqlDbType.Int16)

                    prmId.Value = CType(idhabitacion, Int32)
                    prmEstado.Value = CType(estado, Int16)

                    Dim instruccion As New MySqlCommand("UPDATE habitacion SET HabiEsta=@HabiEsta WHERE HabiCodi = @HabiCodi", cone)

                    instruccion.Parameters.Add(prmId)
                    instruccion.Parameters.Add(prmEstado)

                    instruccion.ExecuteNonQuery()
                End Using
                MessageBox.Show("Eliminado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub DiasRestantes(ByRef label As Label, ByVal numero As String)
            Try
                cone.Open()
                Using cone
                    Dim prmNumero As New MySqlParameter("@HabiNume", MySqlDbType.Text)

                    prmNumero.Value = CType(numero, String)

                    Dim recuperar As New MySqlCommand("SELECT reserva.ReseFeSa FROM reserva INNER JOIN habitacion ON (reserva.HabiCodi=habitacion.HabiCodi) WHERE habitacion.HabiNume=@HabiNume", cone)
                    Dim res = recuperar.ExecuteScalar()

                    label.Text = DateDiff(DateInterval.Day, res, Date.Today)

                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub

        Public Sub RecuperarTitular(ByVal idHabitacion As String, ByRef label As Label)
            Try
                cone.Open()
                Using cone
                    Dim prmHabitacion As New MySqlParameter("@HabiCodi", MySqlDbType.Int32)

                    prmHabitacion.Value = CType(idHabitacion, Int32)

                    Dim instruccion As New MySqlCommand("SELECT persona.PersApel,persona.PersNomb FROM alquiler INNER JOIN reserva ON (alquiler.ReseCodi=reserva.ReseCodi) INNER JOIN huesped ON (reserva.PersCodi=huesped.PersCodi) INNER JOIN persona ON (huesped.PersCodi=persona.PersCodi) WHERE alquiler.HabiCodi=@HabiCodi", cone)

                    instruccion.Parameters.Add(prmHabitacion)

                    Dim res As MySqlDataReader = instruccion.ExecuteReader()

                    While res.Read()
                        label.Text = res.GetString(0) & " " & res.GetString(1)
                    End While
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub

    End Class


    Class CDetalleAlquiler
        Public Sub ContarOcupantes(ByRef label As Label, ByVal numero As String)
            Try
                cone.Open()
                Using cone
                    Dim prmNumero As New MySqlParameter("@HabiNume", MySqlDbType.Text)

                    prmNumero.Value = CType(numero, String)

                    Dim recuperar As New MySqlCommand("SELECT COUNT(detallealquiler.PersCodi) FROM detallealquiler INNER JOIN alquiler ON (detallealquiler.AlquCodi=alquiler.AlquCodi) INNER JOIN habitacion ON (alquiler.HabiCodi=habitacion.HabiCodi) WHERE habitacion.HabiNume=@HabiNume AND NOW() BETWEEN detallealquiler.DetaFeIn AND detallealquiler.DetaFeSa", cone)
                    Dim res = recuperar.ExecuteScalar()

                    label.Text = res

                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub

        Public Sub Agregar(ByVal idPersona As String, ByVal idAlquiler As String, ByVal fechaEntrada As String, ByVal fechaSalida As String)
            Try
                cone.Open()
                Using cone
                    Dim prmPersona As New MySqlParameter("@PersCodi", MySqlDbType.Int32)
                    Dim prmAlquiler As New MySqlParameter("@AlquCodi", MySqlDbType.Int32)
                    Dim prmFechaIn As New MySqlParameter("@DetaFeIn", MySqlDbType.Date)
                    Dim prmFechaOut As New MySqlParameter("@DetaFeSa", MySqlDbType.Date)

                    prmPersona.Value = CType(idPersona, Int32)
                    prmAlquiler.Value = CType(idAlquiler, Int32)
                    prmFechaIn.Value = CType(fechaEntrada, Date)
                    prmFechaOut.Value = CType(fechaSalida, Date)

                    Dim instruccion As New MySqlCommand("INSERT INTO detallealquiler (PersCodi,AlquCodi,DetaFeIn,DetaFeSa) VALUES (@PersCodi,@AlquCodi,@DetaFeIn,@DetaFeSa)", cone)

                    instruccion.Parameters.Add(prmPersona)
                    instruccion.Parameters.Add(prmAlquiler)
                    instruccion.Parameters.Add(prmFechaIn)
                    instruccion.Parameters.Add(prmFechaOut)

                    instruccion.ExecuteNonQuery()

                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub Modificar(ByVal idDetalle As String, ByVal idPersona As String, ByVal idAlquiler As String, ByVal fechaEntrada As String, ByVal fechaSalida As String)
            Try
                cone.Open()
                Using cone
                    Dim prmDetalle As New MySqlParameter("DetaCodi", MySqlDbType.Int32)
                    Dim prmPersona As New MySqlParameter("@PersCodi", MySqlDbType.Int32)
                    Dim prmAlquiler As New MySqlParameter("@AlquCodi", MySqlDbType.Int32)
                    Dim prmFechaIn As New MySqlParameter("@DetaFeIn", MySqlDbType.Date)
                    Dim prmFechaOut As New MySqlParameter("@DetaFeSa", MySqlDbType.Date)

                    prmDetalle.Value = CType(idDetalle, Int32)
                    prmPersona.Value = CType(idPersona, Int32)
                    prmAlquiler.Value = CType(idAlquiler, Int32)
                    prmFechaIn.Value = CType(fechaEntrada, Date)
                    prmFechaOut.Value = CType(fechaSalida, Date)

                    Dim instruccion As New MySqlCommand("UPDATE detallealquiler SET PersCodi=@PersCodi,AlquCodi=@AlquCodi,DetaFeIn=@DetaFeIn,DetaFeSa=@DetaFeSa WHERE DetaCodi=@DetaCodi", cone)

                    instruccion.Parameters.Add(prmDetalle)
                    instruccion.Parameters.Add(prmPersona)
                    instruccion.Parameters.Add(prmAlquiler)
                    instruccion.Parameters.Add(prmFechaIn)
                    instruccion.Parameters.Add(prmFechaOut)

                    instruccion.ExecuteNonQuery()

                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub Eliminiar(ByVal id As String)
            Try
                cone.Open()
                Using cone
                    Dim prmId As New MySqlParameter("@DetaCodi", MySqlDbType.Int32)

                    prmId.Value = CType(id, Int32)

                    Dim instruccion As New MySqlCommand("DELETE FROM detallealquiler WHERE DetaCodi = @DetaCodi", cone)

                    instruccion.Parameters.Add(prmId)

                    instruccion.ExecuteNonQuery()
                End Using
                MessageBox.Show("Eliminado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub CargarGrilla(ByVal idAlquiler As String, ByRef grilla As DataGridView)
            Try
                cone.Open()
                Using cone
                    Dim prmId As New MySqlParameter("@AlquCodi", MySqlDbType.Int32)

                    prmId.Value = CType(idAlquiler, Int32)

                    Dim recuperar As New MySqlCommand("SELECT persona.PersCodi,persona.PersApel,persona.PersNomb,persona.PersDni,persona.PersTele FROM persona INNER JOIN huesped ON (persona.PersCodi=huesped.PersCodi) INNER JOIN detallealquiler ON (huesped.PersCodi=detallealquiler.PersCodi) WHERE detallealquiler.AlquCodi=@AlquCodi AND NOW() BETWEEN detallealquiler.DetaFeIn AND detallealquiler.DetaFeSa", cone)
                    Dim res = recuperar.ExecuteReader()

                    While res.Read()
                        grilla.Rows.Add(res.GetString(0), res.GetString(1), res.GetString(2), res.GetString(3), res.GetString(4))
                    End While

                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub BajaOcupante(ByVal idAlquiler As String, ByVal idPersona As String)
            Try
                cone.Open()
                Using cone
                    Dim prmAlquiler As New MySqlParameter("AlquCodi", MySqlDbType.Int32)
                    Dim prmPersona As New MySqlParameter("PersCodi", MySqlDbType.Int32)
                    Dim prmFechaOut As New MySqlParameter("@DetaFeSa", MySqlDbType.Date)

                    prmAlquiler.Value = CType(idAlquiler, Int32)
                    prmPersona.Value = CType(idPersona, Int32)
                    prmFechaOut.Value = CType(FechaVBaSQL(Date.Today), Date)

                    Dim instruccion As New MySqlCommand("UPDATE detallealquiler SET DetaFeSa=@DetaFeSa WHERE AlquCodi=@AlquCodi AND PersCodi=@PersCodi", cone)

                    instruccion.Parameters.Add(prmAlquiler)
                    instruccion.Parameters.Add(prmPersona)
                    instruccion.Parameters.Add(prmFechaOut)

                    instruccion.ExecuteNonQuery()

                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

    End Class


    Class CPersona

        Public Sub Agregar(ByVal apellido As String, ByVal nombre As String, ByVal fechaNac As String, ByVal dni As String, ByVal direccion As String, ByVal telefono As String, ByVal clave As String)
            Try
                cone.Open()
                Using cone

                    Dim prmApellido As New MySqlParameter("@PersApel", MySqlDbType.Text)
                    Dim prmNombre As New MySqlParameter("@PersNomb", MySqlDbType.Text)
                    Dim prmFechaNac As New MySqlParameter("@PersFeNa", MySqlDbType.Date)
                    Dim prmDni As New MySqlParameter("@PersDni", MySqlDbType.Text)
                    Dim prmDireccion As New MySqlParameter("@PersDire", MySqlDbType.Text)
                    Dim prmTelefono As New MySqlParameter("@PersTele", MySqlDbType.Text)
                    Dim prmClave As New MySqlParameter("@PersClav", MySqlDbType.Text)
                    Dim prmEstado As New MySqlParameter("@PersEsta", MySqlDbType.Int16)

                    prmApellido.Value = StrConv(apellido, 3)
                    prmNombre.Value = StrConv(nombre, 3)
                    prmFechaNac.Value = FechaVBaSQL(fechaNac)
                    prmDni.Value = CType(dni, String)
                    prmDireccion.Value = CType(direccion, String)
                    prmTelefono.Value = CType(telefono, String)
                    prmClave.Value = CType(clave, String)
                    prmEstado.Value = 1

                    Dim instruccion As New MySqlCommand("INSERT INTO persona (PersApel,PersNomb,PersFeNa,PersDni,PersDire,PersTele,PersClav,PersEsta) VALUES (@PersApel,@PersNomb,@PersFeNa,@PersDni,@PersDire,@PersTele,@PersClav,@PersEsta)", cone)

                    instruccion.Parameters.Add(prmApellido)
                    instruccion.Parameters.Add(prmNombre)
                    instruccion.Parameters.Add(prmFechaNac)
                    instruccion.Parameters.Add(prmDni)
                    instruccion.Parameters.Add(prmDireccion)
                    instruccion.Parameters.Add(prmTelefono)
                    instruccion.Parameters.Add(prmClave)
                    instruccion.Parameters.Add(prmEstado)

                    instruccion.ExecuteNonQuery()
                    MessageBox.Show("Agregada")
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub Modificar(ByVal id As String, ByVal apellido As String, ByVal nombre As String, ByVal fechaNac As String, ByVal dni As String, ByVal direccion As String, ByVal telefono As String, ByVal clave As String, ByVal estado As String)
            Try
                cone.Open()
                Using cone

                    Dim prmId As New MySqlParameter("@PersCodi", MySqlDbType.Int32)
                    Dim prmApellido As New MySqlParameter("@PersApel", MySqlDbType.Text)
                    Dim prmNombre As New MySqlParameter("@PersNomb", MySqlDbType.Text)
                    Dim prmFechaNac As New MySqlParameter("@PersFeNa", MySqlDbType.Date)
                    Dim prmDni As New MySqlParameter("@PersDni", MySqlDbType.Text)
                    Dim prmDireccion As New MySqlParameter("@PersDire", MySqlDbType.Text)
                    Dim prmTelefono As New MySqlParameter("@PersTele", MySqlDbType.Text)
                    Dim prmClave As New MySqlParameter("@PersClav", MySqlDbType.Text)
                    Dim prmEstado As New MySqlParameter("@PersEsta", MySqlDbType.Int16)

                    prmApellido.Value = StrConv(apellido, 3)
                    prmNombre.Value = StrConv(nombre, 3)
                    prmFechaNac.Value = fechaNac
                    prmDni.Value = CType(dni, String)
                    prmDireccion.Value = CType(direccion, String)
                    prmTelefono.Value = CType(telefono, String)
                    prmClave.Value = CType(clave, String)
                    prmEstado.Value = CType(estado, Int16)

                    Dim instruccion As New MySqlCommand("UPDATE persona SET (PersApel=@PersApel,PersNomb=@PersNomb,PersFeNa=@PersFeNa,PersDni=@PersDni,PersDire=@PersDire,PersTele=@PersTele,PersClav=@PersClav,PersEsta=@PersEsta) WHERE PersCodi=@PersCodi", cone)

                    instruccion.Parameters.Add(prmId)
                    instruccion.Parameters.Add(prmApellido)
                    instruccion.Parameters.Add(prmNombre)
                    instruccion.Parameters.Add(prmFechaNac)
                    instruccion.Parameters.Add(prmDni)
                    instruccion.Parameters.Add(prmDireccion)
                    instruccion.Parameters.Add(prmTelefono)
                    instruccion.Parameters.Add(prmClave)
                    instruccion.Parameters.Add(prmEstado)

                    instruccion.ExecuteNonQuery()
                    MessageBox.Show("Modificado")
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        '// Carga grilla Empleados.
        Public Sub CargarGrilla(ByRef grilla As DataGridView)
            Try
                cone.Open()
                Using cone
                    Dim instruccion As New MySqlCommand("SELECT * FROM persona INNER JOIN empleado ON (persona.PersCodi=empleado.PersCodi) INNER JOIN tipousuario ON (empleado.PersTiUs=tipousuario.TiUsCodi) WHERE persona.PersEsta = 1", cone)
                    Dim res = instruccion.ExecuteReader()
                    While res.Read() '    'id               'apellido         'nombre          'telefono         'tipoUsuario      'fechaNac           'dni             'direccion        'seguroSocial       'fechaIngreso     'clave          
                        grilla.Rows.Add(res.GetString(0), res.GetString(1), res.GetString(2), res.GetString(7), res.GetString(16), res.GetString(4), res.GetString(5), res.GetString(6), res.GetString(14), res.GetString(13), res.GetString(8))
                    End While

                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub


        Public Sub Eliminar(ByVal id As String)
            Try
                cone.Open()
                Using cone

                    Dim prmIdPersona As New MySqlParameter("@PersCodi", MySqlDbType.Int32)
                    Dim prmEstado As New MySqlParameter("@PersEsta", MySqlDbType.Int16)

                    prmIdPersona.Value = CType(id, Int32)
                    prmEstado.Value = 0

                    Dim instruccion As New MySqlCommand("UPDATE persona SET PersEsta=@PersEsta WHERE PersCodi=@PersCodi", cone)

                    instruccion.Parameters.Add(prmIdPersona)
                    instruccion.Parameters.Add(prmEstado)

                    instruccion.ExecuteNonQuery()
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

    End Class


    Class CHuesped

        Public Sub Agregar(ByVal idPersona As String, ByVal frecuente As String, ByVal confiable As String, ByVal observacion As String)
            Try
                cone.Open()
                Using cone

                    Dim prmIdPersona As New MySqlParameter("@PersCodi", MySqlDbType.Text)
                    Dim prmFrecuente As New MySqlParameter("@PersFrec", MySqlDbType.Text)
                    Dim prmConfiable As New MySqlParameter("@PersConf", MySqlDbType.Date)
                    Dim prmObservacion As New MySqlParameter("@PersObse", MySqlDbType.Text)

                    prmIdPersona.Value = CType(idPersona, Int32)
                    prmFrecuente.Value = CType(frecuente, Int16)
                    prmConfiable.Value = CType(confiable, Int16)
                    prmObservacion.Value = CType(observacion, String)

                    Dim instruccion As New MySqlCommand("INSERT INTO huesped (PersCodi,PersFrec,PersConf,PersObse) VALUES (@PersCodi,@PersFrec,@PersConf,@PersObse)", cone)

                    instruccion.Parameters.Add(prmIdPersona)
                    instruccion.Parameters.Add(prmFrecuente)
                    instruccion.Parameters.Add(prmConfiable)
                    instruccion.Parameters.Add(prmObservacion)

                    instruccion.ExecuteNonQuery()
                    MessageBox.Show("Agregada")
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub Modificar(ByVal idPersona As String, ByVal frecuente As String, ByVal confiable As String, ByVal observacion As String)
            Try
                cone.Open()
                Using cone

                    Dim prmIdPersona As New MySqlParameter("@PersCodi", MySqlDbType.Int32)
                    Dim prmFrecuente As New MySqlParameter("@PersFrec", MySqlDbType.Int16)
                    Dim prmConfiable As New MySqlParameter("@PersConf", MySqlDbType.Int16)
                    Dim prmObservacion As New MySqlParameter("@PersObse", MySqlDbType.Text)

                    prmIdPersona.Value = CType(idPersona, Int32)
                    prmFrecuente.Value = CType(frecuente, Int16)
                    prmConfiable.Value = CType(confiable, Int16)
                    prmObservacion.Value = CType(observacion, String)

                    Dim instruccion As New MySqlCommand("UPDATE huesped SET PersFrec=@PersFrec, PersConf=@PersConf, PersObse=@PersObse WHERE PersCodi=,PersCodi", cone)

                    instruccion.Parameters.Add(prmIdPersona)
                    instruccion.Parameters.Add(prmFrecuente)
                    instruccion.Parameters.Add(prmConfiable)
                    instruccion.Parameters.Add(prmObservacion)

                    instruccion.ExecuteNonQuery()
                    MessageBox.Show("modificado")
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

    End Class


    Class CEmpleado

        Public Sub Agregar(ByVal idPersona As String, ByVal tipoUsuario As String, ByVal fechaIngreso As String, ByVal numeroSeguro As String)
            Try
                cone.Open()
                Using cone

                    Dim prmIdPersona As New MySqlParameter("@PersCodi", MySqlDbType.Int32)
                    Dim prmTipoUsuario As New MySqlParameter("@PersTiUs", MySqlDbType.Int32)
                    Dim prmFechaIngreso As New MySqlParameter("@PersFeIn", MySqlDbType.Date)
                    Dim prmNumeroSeguro As New MySqlParameter("@PersNSS", MySqlDbType.Text)

                    prmIdPersona.Value = CType(idPersona, Int32)
                    prmTipoUsuario.Value = CType(tipoUsuario, Int32)
                    prmFechaIngreso.Value = CType(fechaIngreso, Date)
                    prmNumeroSeguro.Value = CType(numeroSeguro, String)

                    Dim instruccion As New MySqlCommand("INSERT INTO empleado (PersCodi,PersTiUs,PersFeIn,PersNSS) VALUES (@PersCodi,@PersTiUs,@PersFeIn,@PersNSS)", cone)

                    instruccion.Parameters.Add(prmIdPersona)
                    instruccion.Parameters.Add(prmTipoUsuario)
                    instruccion.Parameters.Add(prmFechaIngreso)
                    instruccion.Parameters.Add(prmNumeroSeguro)

                    instruccion.ExecuteNonQuery()
                    MessageBox.Show("Agregada")
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub Modificar(ByVal idPersona As String, ByVal tipoUsuario As String, ByVal fechaIngreso As String, ByVal numeroSeguro As String)
            Try
                cone.Open()
                Using cone

                    Dim prmIdPersona As New MySqlParameter("@PersCodi", MySqlDbType.Int32)
                    Dim prmTipoUsuario As New MySqlParameter("@PersTiUs", MySqlDbType.Int32)
                    Dim prmFechaIngreso As New MySqlParameter("@PersFeIn", MySqlDbType.Date)
                    Dim prmNumeroSeguro As New MySqlParameter("@PersNSS", MySqlDbType.Text)

                    prmIdPersona.Value = CType(idPersona, Int32)
                    prmTipoUsuario.Value = CType(tipoUsuario, Int32)
                    prmFechaIngreso.Value = CType(fechaIngreso, Date)
                    prmNumeroSeguro.Value = CType(numeroSeguro, String)

                    Dim instruccion As New MySqlCommand("UPDATE empleado SET (PersCodi=@PersCodi,PersTiUs=@PersTiUs,PersFeIn=@PersFeIn,PersNSS=@PersNSS) WHERE PersCodi=@PersCodi", cone)

                    instruccion.Parameters.Add(prmIdPersona)
                    instruccion.Parameters.Add(prmTipoUsuario)
                    instruccion.Parameters.Add(prmFechaIngreso)
                    instruccion.Parameters.Add(prmNumeroSeguro)

                    instruccion.ExecuteNonQuery()
                    MessageBox.Show("Modificado")
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

    End Class


    Class CTipoUsuario

        Public Sub Agregar(ByVal nombre As String, ByVal estado As String)  'CTipoUsuario
            Try
                'coneccion con la base de datos
                cone.Open()
                Using cone
                    'parametro nombre                'nombre fantasia 'tipo de dato de la BD
                    Dim prmNombre As New MySqlParameter("@TiUsNomb", MySqlDbType.Text)
                    Dim prmEstado As New MySqlParameter("@TiUsEsta", MySqlDbType.Int16)
                    'conversion, primera letra en mayuscula, el resto en minuscula
                    prmNombre.Value = StrConv(nombre, 3)

                    prmEstado.Value = 0
                    If estado = "Activo" Then
                        prmEstado.Value = 1
                    End If

                    'definicion del comando MySql Insert en tabla tipoUsuario, atributo TiUsNomb, el valor fantasia @TiUsNomb con la coneccion
                    Dim insertar As New MySqlCommand("INSERT INTO tipousuario (TiUsNomb,TiUsEsta) VALUES (@TiUsNomb,@TiUsEsta)", cone)
                    'agrega a insertar, el parametro prmNombre
                    insertar.Parameters.Add(prmNombre)
                    insertar.Parameters.Add(prmEstado)
                    'en este punto se hace la insercion en la BD
                    insertar.ExecuteNonQuery()  'NonQuery para (Insert, Update, Delete), para Select: executeScalar(0 o 1 resultado) o executeReader(0 o muchos resultados)
                End Using ' cierre de coneccion con la BD
                MessageBox.Show("Agregado") 'mensaje de confirmacion, todo OK!
            Catch ex As Exception   'mensaje de error
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub CargaGrilla(ByVal grilla As DataGridView) 'CTipoUsuario
            Dim userType As MySqlDataReader
            Try
                cone.Open()
                Using cone
                    Dim instruction As New MySqlCommand("SELECT * FROM tipousuario", cone)
                    userType = instruction.ExecuteReader()
                    While userType.Read()
                        grilla.Rows.Add(userType.GetString(0), userType.GetString(1))
                    End While
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub Modificar(ByVal id As String, ByVal nombre As String, ByVal estado As String) 'CTipoUsuario
            Try
                cone.Open()
                Using cone
                    Dim prmId As New MySqlParameter("@TiUsCodi", MySqlDbType.Int32)
                    Dim prmNombre As New MySqlParameter("@TiUsNomb", MySqlDbType.Text)
                    Dim prmEstado As New MySqlParameter("@TiUsEsta", MySqlDbType.Int16)

                    prmId.Value = CType(id, Int32)
                    prmNombre.Value = nombre
                    prmEstado.Value = CType(estado, Int16)

                    Dim modificar As New MySqlCommand("UPDATE tipousuario SET TiUsNomb=@TiUsNomb, TiUsEsta=@TiUsEsta WHERE TiUsCodi = @TiUsCodi", cone)

                    modificar.Parameters.Add(prmId)
                    modificar.Parameters.Add(prmNombre)
                    modificar.Parameters.Add(prmEstado)

                    modificar.ExecuteNonQuery()
                End Using
                MessageBox.Show("Agregado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub CargarTxt(ByRef id As Label, ByRef nombre As TextBox, ByVal estado As TextBox, ByVal grilla As DataGridView) 'CTipoUsuario
            Dim dgvrow As DataGridViewRow
            For Each dgvrow In grilla.SelectedRows
                id.Text = dgvrow.Cells(0).Value()
                nombre.Text = dgvrow.Cells(1).Value()
                estado.Text = dgvrow.Cells(2).Value()
            Next

        End Sub

        Public Sub Eliminar(ByVal id As String) ' eliminacion Logica
            Try
                cone.Open()
                Using cone
                    Dim prmId As New MySqlParameter("@TiUsCodi", MySqlDbType.Int32)
                    Dim prmEstado As New MySqlParameter("@TiUsEsta", MySqlDbType.Int16)

                    prmId.Value = CType(id, Int32)
                    prmEstado.Value = 0

                    Dim instruccion As New MySqlCommand("UPDATE tipousuario SET TiUsEsta=@TiUsEsta WHERE TiUsCodi = @TiUsCodi", cone)

                    instruccion.Parameters.Add(prmId)
                    instruccion.Parameters.Add(prmEstado)

                    instruccion.ExecuteNonQuery()
                End Using
                MessageBox.Show("Eliminado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub BuscarTipo(ByRef grilla As DataGridView, ByVal busqueda As String)
            Try
                cone.Open()
                Using cone
                    Dim prmBusqueda As New MySqlParameter("@TiUsNomb", MySqlDbType.Text)
                    prmBusqueda.Value = busqueda
                    Dim recuperar As New MySqlCommand("SELECT * FROM tipousuario WHERE TiUsNomb = @TiUsNomb%", cone)
                    Dim res As MySqlDataReader
                    recuperar.Parameters.Add(prmBusqueda)
                    res = recuperar.ExecuteReader()
                    While res.Read()
                        grilla.Rows.Add(res.GetString(0), res.GetString(1))
                    End While
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub ObtenerIdTipo(ByVal nombre As String, ByRef label As Label)
            Dim res As MySqlDataReader
            Try
                cone.Open()
                Using cone
                    Dim prmNombre As New MySqlParameter("@TiUsNomb", MySqlDbType.Text)
                    prmNombre.Value = nombre
                    Dim instruction As New MySqlCommand("SELECT TiUsCodi FROM tipousuario WHERE TiUsNomb LIKE @TiUsNomb", cone)
                    instruction.Parameters.Add(prmNombre)
                    res = instruction.ExecuteReader()
                    While res.Read()
                        label.Text = res.GetString(0)
                    End While
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

    End Class


    Public Class CTipoRotura

        Public Sub Agregar(ByVal nombre As String)
            Try
                cone.Open()
                Using cone
                    Dim prmNombre As New MySqlParameter("@TiRoNomb", MySqlDbType.Text)

                    prmNombre.Value = StrConv(nombre, 3)

                    Dim insertar As New MySqlCommand("INSERT INTO tiporotura TiRoNomb VALUES TiRoNomb", cone)

                    insertar.Parameters.Add(prmNombre)

                    insertar.ExecuteNonQuery()
                End Using
                MessageBox.Show("Agregado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub Modificar(ByVal id As String, ByVal nombre As String)
            Try
                cone.Open()
                Using cone
                    Dim prmId As New MySqlParameter("@TiRoCodi", MySqlDbType.Int32)
                    Dim prmNombre As New MySqlParameter("@TiRoNomb", MySqlDbType.Text)

                    prmId.Value = CType(id, Int32)
                    prmNombre.Value = StrConv(nombre, 3)

                    Dim modificar As New MySqlCommand("UPDATE tiporotura SET TiRoNomb=@TiRoNomb WHERE TiRoCodi = @TiRoCodi", cone)

                    modificar.Parameters.Add(prmId)
                    modificar.Parameters.Add(prmNombre)

                    modificar.ExecuteNonQuery()
                End Using
                MessageBox.Show("Agregado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

    End Class


    Public Class CMantenimiento
        Public Sub Agregar(ByVal idHabitacion As String, ByVal idRotura As String, ByVal idPersona As String)
            Try
                cone.Open()
                Using cone
                    Dim prmHabitacion As New MySqlParameter("@HabiCodi", MySqlDbType.Int32)
                    Dim prmRotura As New MySqlParameter("@TiRoCodi", MySqlDbType.Int32)
                    Dim prmPersona As New MySqlParameter("@PersCodi", MySqlDbType.Int32)
                    Dim prmFinalizado As New MySqlParameter("@MantFina", MySqlDbType.Int16)

                    prmHabitacion.Value = CType(idHabitacion, Int32)
                    prmRotura.Value = CType(idRotura, Int32)
                    prmPersona.Value = CType(idPersona, Int32)
                    prmFinalizado.Value = 0

                    Dim insertar As New MySqlCommand("INSERT INTO mantenimiento HabiCodi,PersCodi,TiRoCodi,MantFina VALUES @HabiCodi,@PersCodi,@TiRoNomb,@MantFina", cone)

                    insertar.Parameters.Add(prmHabitacion)
                    insertar.Parameters.Add(prmRotura)
                    insertar.Parameters.Add(prmPersona)
                    insertar.Parameters.Add(prmFinalizado)

                    insertar.ExecuteNonQuery()
                End Using
                MessageBox.Show("Agregado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub

        Public Sub Modificar(ByVal idMantenimiento As String, idHabitacion As String, ByVal idRotura As String, ByVal idPersona As String, ByVal finalizado As String)
            Try
                cone.Open()
                Using cone
                    Dim prmMantenimiento As New MySqlParameter("@MantCodi", MySqlDbType.Int32)
                    Dim prmHabitacion As New MySqlParameter("@HabiCodi", MySqlDbType.Int32)
                    Dim prmRotura As New MySqlParameter("@TiRoCodi", MySqlDbType.Int32)
                    Dim prmPersona As New MySqlParameter("@PersCodi", MySqlDbType.Int32)
                    Dim prmFinalizado As New MySqlParameter("@MantFina", MySqlDbType.Int16)

                    prmRotura.Value = CType(idMantenimiento, Int32)
                    prmHabitacion.Value = CType(idHabitacion, Int32)
                    prmRotura.Value = CType(idRotura, Int32)
                    prmPersona.Value = CType(idPersona, Int32)
                    prmFinalizado.Value = CType(finalizado, Int16)

                    Dim modificar As New MySqlCommand("UPDATE mantenimiento SET HabiCodi=@HabiCodi,TiRoCodi=@TiRoCodi,PersCodi=@PersCodi,MantFina=@MantFina WHERE MantTiRoCodi = @MantCodi", cone)

                    modificar.Parameters.Add(prmMantenimiento)
                    modificar.Parameters.Add(prmHabitacion)
                    modificar.Parameters.Add(prmRotura)
                    modificar.Parameters.Add(prmPersona)
                    modificar.Parameters.Add(prmFinalizado)

                    modificar.ExecuteNonQuery()
                End Using
                MessageBox.Show("Agregado")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub CargarGrilla(ByRef grilla As DataGridView)
            Try
                cone.Open()
                Using cone
                    Dim recuperar As New MySqlCommand("SELECT mantenimiento.MantCodi,habitacion.HabiNume,tiporotura.TiRoNomb,empleado.PersCodi,mantenimiento.MantFina FROM mantenimiento INNER JOIN habitacion ON (mantenimiento.HabiCodi=habitacion.HabiCodi) INNER JOIN empleado ON (mantenimiento.PersCodi=empleado.PersCodi) INNER JOIN tiporotura ON (mantenimiento.TiRoCodi=tiporotura.TiRoCodi)")
                    Dim res = recuperar.ExecuteReader()

                    While res.Read()
                        grilla.Rows.Add(res.GetString(0), res.GetString(1), res.GetString(2), res.GetString(3), res.GetString(4))
                    End While

                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

    End Class


    Public Class CVigilancia

        Public Sub CargarGrilla(ByRef grilla As DataGridView, ByVal fechaInicio As String, ByVal fechaFin As String)
            Try
                cone.Open()
                Using cone
                    Dim prmFechaInicio As New MySqlParameter("@VigiFeEn", MySqlDbType.Date)
                    Dim prmFechaFin As New MySqlParameter("@VigiFeSa", MySqlDbType.Date)

                    prmFechaInicio.Value = FechaVBaSQL(fechaInicio)
                    prmFechaFin.Value = FechaVBaSQL(fechaFin)

                    Dim recuperar As New MySqlCommand("SELECT persona.PersApel, persona.PersNomb, habitacion.HabiNume, vigilancia.VigiFeEn, vigilancia.VigiFeSa, vigilancia.VigiObse FROM vigilancia INNER JOIN habitacion ON (vigilancia.HabiCodi=habitacion.HabiCodi) INNER JOIN persona ON (vigilancia.PersCodi=persona.PersCodi) WHERE VigiFeEn BETWEEN @VigiFeEn AND @VigiFeSa", cone)

                    recuperar.Parameters.Add(prmFechaInicio)
                    recuperar.Parameters.Add(prmFechaFin)

                    Dim res = recuperar.ExecuteReader()

                    While res.Read()
                        grilla.Rows.Add(res.GetString(0) & " " & res.GetString(1), res.GetString(2), res.GetString(3), res.GetString(4), res.GetString(5))
                    End While

                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

    End Class

    '                   'lugar donde guardar la clave    'tipo de usuario al que se le asigna la clave
    Public Sub GenerarClave(ByRef text As TextBox, ByRef tipoUsuario As String)

        Randomize()

        Dim ok = False
        Dim rand As New Random(DateTime.Now.Millisecond)
        Dim clave As String = ""

        While Not ok
            clave = ""

            If tipoUsuario = "Administrador" Then
                clave += "1"
            ElseIf tipoUsuario = "Recepcionista" Then
                clave += "2"
            ElseIf tipoUsuario = "Limpieza" Then
                clave += "3"
            ElseIf tipoUsuario = "Mantenimiento" Then
                clave += "4"
            Else    'clave para el huesped, no tiene tipo de usuario
                clave += "5"
            End If

            While Len(clave) < 6
                If rand.Next(0, 2) = 0 Then 'para alternar entre generacion de letra y numero
                    clave += CType(rand.Next(0, 9), String) 'generacion de numeros
                Else
                    clave += Convert.ToChar(rand.Next(65, 90))  'generacion de numeros correspondientes a letras de la tabla ascii, conversion a caracter
                End If
            End While

            Try
                cone.Open()
                Using cone
                    Dim prmBusqueda As New MySqlParameter("@PersClav", MySqlDbType.Text)
                    prmBusqueda.Value = clave
                    Dim recuperar As New MySqlCommand("SELECT PersCodi FROM persona WHERE PersClav=@PersClav", cone)
                    Dim res As MySqlDataReader
                    recuperar.Parameters.Add(prmBusqueda)
                    res = recuperar.ExecuteScalar()
                    If res Is Nothing Then  'si res is nothing la clave esta sin usar, ok=true y finaliza el while, sino, se genera de nuevo
                        ok = True
                    End If

                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try


        End While


        text.Text = clave

    End Sub

    Public Sub CalcularReserva(ByVal tipoHabitacion As String, ByVal fechaIn As String, ByVal fechaOut As String, ByRef label As Label)
        Dim res As MySqlDataReader
        Try
            cone.Open()
            Using cone
                Dim prmNombre As New MySqlParameter("@TiHaNomb", MySqlDbType.Text)
                Dim prmFechaIngreso As New MySqlParameter("@ReseFeIn", MySqlDbType.Date)
                Dim prmFechaSalida As New MySqlParameter("@ReseFeSa", MySqlDbType.Date)

                prmNombre.Value = tipoHabitacion
                prmFechaIngreso.Value = CType(fechaIn, Date)
                prmFechaSalida.Value = CType(fechaOut, Date)

                Dim instruction As New MySqlCommand("SELECT COUNT(reserva.TiHaCodi) FROM reserva INNER JOIN tipohabitacion ON (reserva.TiHaCodi=tipohabitacion.TiHaCodi) WHERE tipohabitacion.TiHaNomb=@TiHaNomb AND ((reserva.ReseFeIn<@ReseFeIn AND reserva.ReseFeSa<@ReseFeIn) OR (reserva.ReseFeIn>@ReseFeSa AND reserva.ReseFeSa>@ReseFeSa)) ", cone)

                instruction.Parameters.Add(prmNombre)
                instruction.Parameters.Add(prmFechaIngreso)
                instruction.Parameters.Add(prmFechaSalida)

                res = instruction.ExecuteReader()

                While res.Read()
                    Label.Text = res.GetString(0)
                End While
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    'conversion de fecha formato visual basic a sql
    Public Function FechaVBaSQL(ByVal fecha As String) As String
        '         año                       'mes                       'dia
        Return Mid(fecha, 7, 10) + "-" + Mid(fecha, 4, 2) + "-" + Mid(fecha, 1, 2)
    End Function

    Public Function SeleccionarHabitacion(ByVal habitacion As Button) As Boolean
        Dim res As Boolean = False
        If habitacion.BackColor = Color.LimeGreen Then
            res = True
        End If
        Return res
    End Function

    'procedimiento para permitir solo ingreso de letras en cajas de texto
    Public Sub SoloLetras(ByVal caracter As KeyPressEventArgs, ByRef ctrl As Boolean) 'control para permitir solo el ingreso de letras
        If Not (Char.IsLetter(caracter.KeyChar) Or caracter.KeyChar = Chr(8) Or caracter.KeyChar = Chr(32)) Then
            caracter.Handled = True
            If Not (ctrl) Then
                ctrl = True
                MessageBox.Show("Solo puede ingresar letras", "Aviso!")
            End If
        End If
    End Sub


End Module
