VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccessTier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'========================================================================
' Copyright © CHMR 2007 All rights reserved.
'========================================================================
'
'   Module         - cAccessTier
'   Version Number - 1.0
'   Last Updated   - January 18th 2007 - 5:34 PM
'   Author         - Algaze, Gastón
'
'
'
'
'
'========================================================================
' This file contains trade secrets of CHMR No part
' may be reproduced or transmitted in any form by any means or for any purpose
' without the express written permission of CHMR.
'========================================================================

'========================================================================
'Variables miembro
'========================================================================
Private moConn As ADODB.Connection
Private mrsTemp As ADODB.Recordset
Private msDSN As String

'========================================================================
'Variables auxiliares
'========================================================================
Private strSQL As String
Private strXML As String


'========================================================================
'Inicialización de clase
'========================================================================
Private Sub Class_Initialize()
Dim oRs As Recordset
    
    Set moConn = New ADODB.Connection
    moConn.CursorLocation = adUseServer
    
    Set oRs = RecordsetFromXMLFile(App.Path & "\Biblos_DA.Config.xml")
    msDSN = oRs(0)
    Set oRs = Nothing
    
End Sub
'========================================================================
'Finalización de clase
'========================================================================
Private Sub Class_Terminate()
    Set moConn = Nothing
End Sub


'========================================================================
'Métodos y Funciones
'========================================================================

'========================================================================
'
'   Title          - Función Connect()
'   Description    - Funcion para conectarse a la base de datos.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================

Private Function Connect(ByVal sConn As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo ConnectErr
    
    Connect = False
    
    If moConn.State = adStateClosed Then
        moConn.Open sConn
    End If
    
    Connect = True
    
    Exit Function
    
ConnectErr:
    ShowError "cAccessTier.Connect", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'
'   Title          - Función Disconnect()
'   Description    - Funcion para desconectarse a la base de datos.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Private Function Disconnect(Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo DisconnectErr
    
    Disconnect = False
    
    If moConn.State = adStateOpen Then
        moConn.Close
    End If
    
    Disconnect = True
    
    Exit Function
    
DisconnectErr:
    ShowError "cAccessTier.Disconnect", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'
'   Title          - Función RetrieveData()
'   Description    - Funcion para obtener datos de la base de datos, y
'                    generar un XML con el resultado.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Private Function RetrieveData(ByVal sQuery As String, ByRef sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo retrieveDataErr
    
    Dim oDOM As New MSXML2.DOMDocument
    
    RetrieveData = False
    
    'Query the db in order to fill in the rs
    Set mrsTemp = moConn.Execute(sQuery)
    'Save the data of the rs into a XML DOMDocument
    mrsTemp.Save oDOM, adPersistXML
    sXML = oDOM.xml
    
    mrsTemp.Close
    
    Set mrsTemp = Nothing
    
    RetrieveData = True
    
    Set oDOM = Nothing
    
    Exit Function
    
retrieveDataErr:
    ShowError "cAccessTier.RetrieveData", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'
'   Title          - Función Query()
'   Description    - Ejecuta una consulta en la base de datos.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Private Function Query(ByVal sQuery As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo QueryErr
    
    Query = False
    
    moConn.Execute sQuery
    
    Query = True
    Exit Function
    
QueryErr:
    ShowError "cAccessTier.Query", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'
'   Title          - Función GetRights()
'   Description    - Devuelve un XML con los permisos del usuario, para
'                    una determinada accion
'   Version Number - 1.0
'   Last Updated   - Fecruary 25th 2007 - 8:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function GetRights(sXML As String, sRightsXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo GetRightsErr

    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    GetRights = False
    
    If Not oRs.EOF Then
    
        'strSQL = "SELECT FROM '" & oRs("descripcion") & "', '" & oRs("titulo") & "' )"
        
        strSQL = "SELECT Roles.descripcion AS Rol, funciones.descripcion AS Funcion, " & _
                 "tablas.nombre AS Tabla, permisos.descripcion AS Permiso, " & _
                 "restricciones.valor, operaciones.descripcion, campos.nombre " & _
                 "FROM Roles " & _
                 "INNER JOIN roles_usuarios ON (Roles.id_rol = roles_usuarios.id_rol) " & _
                 "LEFT OUTER JOIN roles_funciones ON (Roles.id_rol = roles_funciones.id_rol) " & _
                 "INNER JOIN funciones ON (roles_funciones.id_funcion = funciones.id_funcion) " & _
                 "INNER JOIN funciones_tablas_permisos ON (funciones.id_funcion = funciones_tablas_permisos.id_funcion) " & _
                 "INNER JOIN permisos ON (funciones_tablas_permisos.id_permiso = permisos.id_permiso) " & _
                 "INNER JOIN tablas ON (funciones_tablas_permisos.id_tabla = tablas.id_tabla) " & _
                 "LEFT OUTER JOIN roles_restricciones ON (Roles.id_rol = roles_restricciones.id_rol) " & _
                 "LEFT OUTER JOIN restricciones ON (roles_restricciones.id_restriccion = restricciones.id_restriccion) " & _
                 "LEFT OUTER JOIN operaciones ON (restricciones.id_operacion = operaciones.id_operacion) " & _
                 "LEFT OUTER JOIN campos ON (restricciones.id_campo = campos.id_campo) " & _
                 "WHERE roles_usuarios.id_usuario = " & oRs("userID")
  
  
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If RetrieveData(strSQL, sRightsXML, lErrNum, sErrDesc, sErrSource) Then
                GetRights = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.GetRights", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function
    
GetRightsErr:
    ShowError "cAccessTier.GetRights", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'
'   Title          - Ubicaciones
'   Description    - Metodos y funciones de trabajo para la clase cUbicacion
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SaveUbicacion(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SaveUbicacionErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    SaveUbicacion = False
    
    If Not oRs.EOF Then
    
        strSQL = "INSERT INTO ubicaciones (descripcion, titulo, fecha_alta, fecha_ult_act) VALUES ( '" & oRs("descripcion") & "', '" & oRs("titulo") & "', " & Format(Now(), "yyyymmdd") & ", " & Format(Now(), "yyyymmdd") & " )"
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                SaveUbicacion = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.SaveUbicacion", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function
    
SaveUbicacionErr:
    ShowError "cAccessTier.SaveUbicacion", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function UpdateUbicacion(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo UpdateUbicacionErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    UpdateUbicacion = False
    
    If Not oRs.EOF Then
    
        strSQL = "UPDATE ubicaciones SET descripcion = '" & oRs("descripcion") & "', titulo = '" & oRs("titulo") & "', fecha_ult_act = " & Format(Now(), "yyyymmdd") & " WHERE id_ubicacion = " & oRs("id")
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                UpdateUbicacion = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.UpdateUbicacion", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

UpdateUbicacionErr:
    ShowError "cAccessTier.UpdateUbicacion", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function DeleteUbicacion(sXML As String, Optional iForever As Integer = 0, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo DeleteUbicacionErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    DeleteUbicacion = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            If iForever = 0 Then
                strSQL = "UPDATE ubicaciones SET fecha_baja = " & Format(Now(), "yyyymmdd") & " WHERE id_ubicacion = " & oRs("id")
            Else
                strSQL = "DELETE FROM ubicaciones WHERE id_ubicacion = " & oRs("id")
            End If
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                DeleteUbicacion = True
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.DeleteUbicacion", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

DeleteUbicacionErr:
    ShowError "cAccessTier.DeleteUbicacion", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function RecoverUbicacion(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo RecoverUbicacionErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    RecoverUbicacion = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            strSQL = "UPDATE ubicaciones SET fecha_baja = NULL WHERE id_ubicacion = " & oRs("id")
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                RecoverUbicacion = True
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.RecoverUbicacion", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

RecoverUbicacionErr:
    ShowError "cAccessTier.RecoverUbicacion", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function SearchUbicacion(sXML As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchUbicacionErr
    
    Dim i As Integer
    Dim strAux As String
    
    SearchUbicacion = False
    
    strSQL = "SELECT id_ubicacion as ""ID"", titulo as ""Titulo"", descripcion as ""Descripcion"", fecha_alta, fecha_ult_act, fecha_baja as ""Fecha_Baja"" FROM ubicaciones "
    
    'caso particular...
    strAux = ""
    'Fin caso particular
    
    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & strAux & " AND " & sSearch
            Else
                strSQL = strSQL & "WHERE " & sSearch
            End If
        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & " " & strAux
            End If
        End If
    End If

        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchUbicacion = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchUbicacionErr:
    ShowError "cAccessTier.SearchUbicacion", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'Métodos y Funciones para la clase cUsuario
'========================================================================
'========================================================================
'
'   Title          - Usuarios
'   Description    - Metodos y funciones de trabajo para la clase cUsuario
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SaveUsuario(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SaveUsuarioErr
    
    Dim oRs As New Recordset
    Dim oRsAux As New Recordset
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    SaveUsuario = False
    
    If Not oRs.EOF Then
    
        strSQL = "INSERT INTO usuarios ( " & _
            "username, nombre, apellido, mail, dni, matricula, " & _
            "fecha_nacimiento, domicilio_calle, domicilio_nro, domicilio_piso, " & _
            "domicilio_unidad, domicilio_cod_postal, tel1, tel2, fecha_alta, fecha_ult_act ) " & _
            "VALUES " & _
            "( '" & oRs("Username") & "', '" & oRs("Nombre") & "', " & _
            "'" & oRs("Apellido") & "', '" & oRs("Mail") & "', '" & oRs("DNI") & "', '" & oRs("Matricula") & "', " & _
            "'" & oRs("Fecha_Nacimiento") & "', '" & oRs("Domicilio_Calle") & "', '" & oRs("Domicilio_Nro") & "', " & _
            "'" & oRs("Domicilio_Piso") & "', '" & oRs("Domicilio_Unidad") & "', " & _
            "'" & oRs("Domicilio_Cod_Postal") & "', '" & oRs("Tel1") & "', '" & oRs("Tel2") & "', " & oRs("fecha_alta") & ", " & oRs("fecha_ult_act") & " )"
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
            
                strSQL = "SELECT id_usuario FROM usuarios WHERE username = '" & oRs("Username") & "'"
                If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
                    Set oRsAux = RecordsetFromXMLString(sXML)
                    If Not oRsAux.EOF Then
                        strSQL = "INSERT INTO roles_usuarios ( " & _
                              "id_usuario, id_rol ) " & _
                              "VALUES " & _
                              "( " & oRsAux("id_usuario") & ", " & oRs("rolID") & " )"
                        If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                            strSQL = "UPDATE usuarios SET password = MD5('" & oRsAux("id_usuario") & oRs("Password") & "') WHERE id_usuario = " & oRsAux("id_usuario")
                            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                                SaveUsuario = True
                            End If
                        End If
                    End If
                End If
                
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.SaveUsuario", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Set oRs = Nothing
    Set oRsAux = Nothing
    
    Exit Function
    
SaveUsuarioErr:
    ShowError "cAccessTier.SaveUsuario", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function UpdateUsuario(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo UpdateUsuarioErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    UpdateUsuario = False
    
    If Not oRs.EOF Then
    
        strSQL = "UPDATE usuarios SET " & _
                "username = '" & oRs("Username") & "', "
                If Len(Trim(oRs("Password"))) > 0 Then
                   strSQL = strSQL & "password = MD5('" & oRs("ID") & oRs("Password") & "'), "
                End If
        strSQL = strSQL & "nombre = '" & oRs("Nombre") & "', " & _
                "apellido = '" & oRs("Apellido") & "', " & _
                "mail = '" & oRs("Mail") & "', " & _
                "dni = '" & oRs("DNI") & "', " & _
                "matricula = '" & oRs("Matricula") & "', " & _
                "fecha_nacimiento = '" & oRs("Fecha_Nacimiento") & "', " & _
                "domicilio_calle = '" & oRs("Domicilio_Calle") & "', " & _
                "domicilio_nro = '" & oRs("Domicilio_Nro") & "', " & _
                "domicilio_piso = '" & oRs("Domicilio_Piso") & "', " & _
                "domicilio_unidad = '" & oRs("Domicilio_Unidad") & "', " & _
                "tel1 = '" & oRs("Tel1") & "', " & _
                "tel2 = '" & oRs("Tel2") & "', fecha_ult_act = " & oRs("fecha_ult_act") & " WHERE id_usuario = " & oRs("ID")
                    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                strSQL = "DELETE FROM roles_usuarios WHERE id_usuario = " & oRs("ID")
                If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                    strSQL = "INSERT INTO roles_usuarios ( " & _
                             "id_usuario, id_rol ) " & _
                             "VALUES " & _
                             "( " & oRs("ID") & ", " & oRs("rolID") & " )"
                    If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                        UpdateUsuario = True
                    End If
                End If
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.UpdateUsuario", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

UpdateUsuarioErr:
    ShowError "cAccessTier.UpdateUsuario", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function DeleteUsuario(sXML As String, Optional iForever As Integer = 0, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo DeleteUsuarioErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    DeleteUsuario = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            If iForever = 0 Then
                strSQL = "UPDATE usuarios SET fecha_baja = " & Format(Now(), "yyyymmdd") & " WHERE id_usuario = " & oRs("id")
            Else
                strSQL = "DELETE FROM usuarios WHERE id_usuario = " & oRs("id")
            End If
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                DeleteUsuario = True
                If iForever > 0 Then
                    strSQL = "DELETE FROM usuarios_roles WHERE id_usuario = " & oRs("id")
                End If
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.DeleteUsuario", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

DeleteUsuarioErr:
    ShowError "cAccessTier.DeleteUsuario", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function RecoverUsuario(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo RecoverUsuarioErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    RecoverUsuario = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            strSQL = "UPDATE usuarios SET fecha_baja = NULL WHERE id_usuario = " & oRs("id")
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                RecoverUsuario = True
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.RecoverUsuario", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

RecoverUsuarioErr:
    If CLng(Err.Number) <> 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "cAccessTier.RecoverUsuario"
    End If
End Function

Public Function SearchUsuario(sXML As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    'On Error GoTo SearchUsuarioErr
    
    Dim i As Integer
    Dim strAux As String
    
    SearchUsuario = False
    
    strSQL = "SELECT usuarios.id_usuario as ""ID"", username as ""Username"", " & _
        "nombre as ""Nombre"", apellido as ""Apellido"", mail as ""Mail"", " & _
        "dni as ""DNI"", " & _
        "matricula as ""Mat."", fecha_nacimiento as ""Nac."", " & _
        "domicilio_calle as ""Calle"", domicilio_nro as ""Nro."", " & _
        "domicilio_piso as ""Piso"", domicilio_unidad as ""Depto."", domicilio_cod_postal as ""C.P."", " & _
        "tel1 as ""Tel1"", tel2 as ""Tel2"", " & _
        "roles.id_rol as ""RolID"", roles.descripcion as ""Rol"", " & _
        "usuarios.fecha_alta as ""Fecha_Alta"", usuarios.fecha_ult_act as ""Fecha_Ult_Act"", usuarios.fecha_baja as ""Fecha_Baja"" " & _
        "FROM usuarios, roles, roles_usuarios "
    
    'caso particular...
    strAux = "WHERE roles.id_rol = roles_usuarios.id_rol AND usuarios.id_usuario = roles_usuarios.id_usuario "
    'Fin caso particular
    
    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & strAux & " AND " & sSearch
            Else
                strSQL = strSQL & "WHERE " & sSearch
            End If
        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & " " & strAux
            End If
        End If
    End If
        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchUsuario = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchUsuarioErr:
    ShowError "cAccessTier.SearchUsuario", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function ChangePasswordUsuario(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo ChangePasswordUsuarioErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    ChangePasswordUsuario = False
    
    If Not oRs.EOF Then
    
        strSQL = "UPDATE usuarios SET " & _
                "password = MD5('" & oRs("userID") & oRs("pwdnew1") & "') " & _
                "WHERE id_usuario = " & oRs("userID")
                    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                ChangePasswordUsuario = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.ChangePasswordUsuario", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

ChangePasswordUsuarioErr:
    ShowError "cAccessTier.ChangePasswordUsuario", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'
'   Title          - Roles
'   Description    - Metodos y funciones de trabajo para la clase cRol
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SaveRol(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SaveRolErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    SaveRol = False
    
    If Not oRs.EOF Then
    
        strSQL = "INSERT INTO Roles (descripcion, fecha_alta, fecha_ult_act) VALUES ( '" & oRs("descripcion") & "', " & oRs("fecha_alta") & ", " & oRs("fecha_ult_act") & " )"
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                SaveRol = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.SaveRol", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function
    
SaveRolErr:
    If CLng(Err.Number) <> 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "cAccessTier.SaveRol"
    End If
End Function

Public Function UpdateRol(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo UpdateRolErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    UpdateRol = False
    
    If Not oRs.EOF Then
    
        strSQL = "UPDATE Roles SET descripcion = '" & oRs("descripcion") & "', privilegios_globales = " & oRs("privilegios_globales") & ", fecha_ult_act = " & oRs("fecha_ult_act") & " WHERE id_Rol = " & oRs("id")
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                UpdateRol = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.UpdateRol", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

UpdateRolErr:
    ShowError "cAccessTier.UpdateRol", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function DeleteRol(sXML As String, Optional iForever As Integer = 0, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo DeleteRolErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    DeleteRol = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            If iForever = 0 Then
                strSQL = "UPDATE Roles SET fecha_baja = " & Format(Now(), "yyyymmdd") & " WHERE id_Rol = " & oRs("id")
            Else
                strSQL = "DELETE FROM Roles WHERE id_Rol = " & oRs("id")
            End If
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                If iForever = 1 Then
                    strSQL = "DELETE FROM Roles_tablas_permisos WHERE id_Rol = " & oRs("id")
                    If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                        DeleteRol = True
                    End If
                Else
                    DeleteRol = True
                End If
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.DeleteRol", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

DeleteRolErr:
    If CLng(Err.Number) <> 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "cAccessTier.DeleteRol"
    End If
End Function

Public Function RecoverRol(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo RecoverRolErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    RecoverRol = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            strSQL = "UPDATE Roles SET fecha_baja = NULL WHERE id_Rol = " & oRs("id")
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                RecoverRol = True
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.RecoverRol", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

RecoverRolErr:
    ShowError "cAccessTier.RecoverRol", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function SearchRol(sXML As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchRolErr
    
    Dim i As Integer
    Dim strAux As String
    
    SearchRol = False
    
    strSQL = "SELECT id_Rol as ""ID"", descripcion as ""Descripcion"", fecha_alta as ""Fecha_Alta"", fecha_ult_act as ""Fecha_Ult_Act"", privilegios_globales as ""Privilegios_Globales"", fecha_baja as ""Fecha_Baja"" FROM Roles "
    
    'caso particular...
    strAux = ""
    'Fin caso particular
    
    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & strAux & " AND " & sSearch
            Else
                strSQL = strSQL & "WHERE " & sSearch
            End If
        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & " " & strAux
            End If
        End If
    End If

        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchRol = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchRolErr:
    ShowError "cAccessTier.SearchRol", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
 
End Function



'========================================================================
'
'   Title          - Restricciones
'   Description    - Metodos y funciones de trabajo para la clase cRestriccion
'   Version Number - 1.0
'   Last Updated   - February 26th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SaveRestriccion(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SaveRestriccionErr
    
    Dim oRs As New Recordset
    Dim oRsAux As New Recordset
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    SaveRestriccion = False
    
    If Not oRs.EOF Then
    
        strSQL = "INSERT INTO Restricciones (valor, campo, tabla, id_operacion, fecha_alta, fecha_ult_act) VALUES ( '" & oRs("valor") & "', '" & oRs("campo") & "', '" & oRs("tabla") & "', " & oRs("id_operacion") & ", " & oRs("fecha_alta") & ", " & oRs("fecha_ult_act") & " )"
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                strSQL = "SELECT id_restriccion FROM restricciones WHERE valor = '" & oRs("valor") & "' AND campo = '" & oRs("campo") & "' AND tabla = '" & oRs("tabla") & "' AND id_operacion = " & oRs("id_operacion")
                If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
                    Set oRsAux = RecordsetFromXMLString(sXML)
                    If Not oRsAux.EOF Then
                        strSQL = "INSERT INTO Roles_Restricciones (id_restriccion, id_rol, fecha_alta, fecha_ult_act) VALUES ( " & oRsAux("id_restriccion") & ", " & oRs("id_rol") & ", " & oRs("fecha_alta") & ", " & oRs("fecha_ult_act") & " )"
                        If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                            SaveRestriccion = True
                        End If
                    End If
                End If
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.SaveRestriccion", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Set oRs = Nothing
    Set oRsAux = Nothing
    
    Exit Function
    
SaveRestriccionErr:
    ShowError "cAccessTier.SaveRestriccion", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    
    Set oRs = Nothing
    Set oRsAux = Nothing
    
End Function

Public Function UpdateRestriccion(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo UpdateRestriccionErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    UpdateRestriccion = False
    
    If Not oRs.EOF Then
    
        strSQL = "UPDATE Restricciones SET valor = '" & oRs("valor") & "', campo = '" & oRs("campo") & "', tabla = '" & oRs("tabla") & "', id_operacion = " & oRs("id_operacion") & ", fecha_ult_act = " & oRs("fecha_ult_act") & " WHERE id_Restriccion = " & oRs("id")
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                UpdateRestriccion = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.UpdateRestriccion", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

UpdateRestriccionErr:
    If CLng(Err.Number) <> 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "cAccessTier.UpdateRestriccion"
    End If
End Function

Public Function DeleteRestriccion(sXML As String, Optional iForever As Integer = 0, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo DeleteRestriccionErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    DeleteRestriccion = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            If iForever = 0 Then
                strSQL = "UPDATE Restricciones SET fecha_baja = " & Format(Now(), "yyyymmdd") & " WHERE id_Restriccion = " & oRs("id")
            Else
                strSQL = "DELETE FROM Restricciones WHERE id_Restriccion = " & oRs("id")
            End If
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                If iForever = 1 Then
                    strSQL = "DELETE FROM Roles_Restricciones WHERE id_Restriccion = " & oRs("id")
                    If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                        DeleteRestriccion = True
                    End If
                Else
                    DeleteRestriccion = True
                End If
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.DeleteRestriccion", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

DeleteRestriccionErr:
    ShowError "cAccessTier.DeleteRestriccion", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function RecoverRestriccion(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo RecoverRestriccionErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    RecoverRestriccion = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            strSQL = "UPDATE Restricciones SET fecha_baja = NULL WHERE id_Restriccion = " & oRs("id")
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                RecoverRestriccion = True
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.RecoverRestriccion", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

RecoverRestriccionErr:
    If CLng(Err.Number) <> 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "cAccessTier.RecoverRestriccion"
    End If
End Function

Public Function SearchRestriccion(sXML As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchRestriccionErr
    Dim strAux
    
    Dim i As Integer
    
    SearchRestriccion = False
    
    'strSQL = "SELECT id_Restriccion as ""ID"", restricciones.tabla, restricciones.campo, id_operacion, valor as ""valor"", fecha_alta as ""Fecha_Alta"", fecha_ult_act as ""Fecha_Ult_Act"", fecha_baja as ""Fecha_Baja"" FROM Restricciones "
    strSQL = "SELECT restricciones.id_restriccion as ""ID"", roles.id_rol as ""RolID"", restricciones.campo, restricciones.tabla, " & _
             "restricciones.id_operacion as ""operacionID"", " & _
             "operaciones.descripcion as Operacion, restricciones.valor as ""Valor"", restricciones.fecha_alta, " & _
             "restricciones.fecha_ult_act, restricciones.fecha_baja " & _
             "FROM roles_restricciones " & _
             "INNER JOIN roles ON (roles_restricciones.id_rol = roles.id_rol) " & _
             "INNER JOIN restricciones ON (roles_restricciones.id_restriccion = restricciones.id_restriccion) " & _
             "INNER JOIN operaciones ON (restricciones.id_operacion = `operaciones`.id_operacion) "
    
    'caso particular...
    strAux = "WHERE roles.id_rol = roles_Restricciones.id_rol "
    'Fin caso particular
    
    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & strAux & " AND " & sSearch
            Else
                strSQL = strSQL & "WHERE " & sSearch
            End If
        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & " " & strAux
            End If
        End If
    End If

        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchRestriccion = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchRestriccionErr:
    ShowError "cAccessTier.SearchRestriccion", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function SearchRestriccionByRol(sXML As String, iRol As Integer, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchRestriccionByRolErr
    
    Dim i As Integer
    
    SearchRestriccionByRol = False
    
    strSQL = "SELECT restricciones.id_restriccion as ""ID"", roles.id_rol as ""RolID"", restricciones.campo, restricciones.tabla, " & _
             "restricciones.id_operacion as ""operacionID"", " & _
             "operaciones.descripcion as Operacion, restricciones.valor as ""Valor"", restricciones.fecha_alta, " & _
             "restricciones.fecha_ult_act, restricciones.fecha_baja " & _
             "FROM roles_restricciones " & _
             "INNER JOIN roles ON (roles_restricciones.id_rol = roles.id_rol) " & _
             "INNER JOIN restricciones ON (roles_restricciones.id_restriccion = restricciones.id_restriccion) " & _
             "INNER JOIN operaciones ON (restricciones.id_operacion = `operaciones`.id_operacion) " & _
             "WHERE roles.id_rol = roles_Restricciones.id_rol AND " & _
             "roles.id_rol = " & CStr(iRol)
        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchRestriccionByRol = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchRestriccionByRolErr:
    ShowError "cAccessTier.SearchRestriccionByRol", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'
'   Title          - Campos
'   Description    - Metodos y funciones de trabajo para la clase cCampo
'   Version Number - 1.0
'   Last Updated   - February 26th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================

Public Function SearchCampoByTabla(sXML As String, sTable As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchCampoByTablaErr
    
    SearchCampoByTabla = False
    
    strSQL = "SHOW COLUMNS FROM " & sTable
           
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchCampoByTabla = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchCampoByTablaErr:
    ShowError "cAccessTier.SearchCampoByTabla", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function


'========================================================================
'
'   Title          - Operaciones
'   Description    - Metodos y funciones de trabajo para la clase cOperacion
'   Version Number - 1.0
'   Last Updated   - February 26th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================

Public Function SearchOperacion(sXML As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchOperacionErr
    
    Dim i As Integer
    Dim strAux As String
    
    SearchOperacion = False
    
    strSQL = "SELECT id_Operacion as ""ID"", descripcion FROM Operaciones"
    
    'caso particular...
    strAux = ""
    'Fin caso particular
    
    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & strAux & " AND " & sSearch
            Else
                strSQL = strSQL & "WHERE " & sSearch
            End If
        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & " " & strAux
            End If
        End If
    End If

        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchOperacion = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchOperacionErr:
    ShowError "cAccessTier.SearchOperacion", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'
'   Title          - Permisos
'   Description    - Metodos y Funciones de trabajo para la clase cPermiso
'   Version Number - 1.0
'   Last Updated   - February 26th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SearchPermisoByRol(sXML As String, iRol As Integer, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchPermisoByRolErr
    
    Dim i As Integer
    
    SearchPermisoByRol = False
    
    strSQL = "SELECT Roles_tablas_permisos.tabla, Roles_tablas_permisos.permiso as Permiso " & _
             "FROM Roles " & _
             "INNER JOIN Roles_tablas_permisos ON (Roles.id_Rol = Roles_tablas_permisos.id_Rol) " & _
             "WHERE Roles.id_Rol = Roles_tablas_permisos.id_Rol AND " & _
             "Roles.id_Rol = " & CStr(iRol)
        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchPermisoByRol = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchPermisoByRolErr:
    ShowError "cAccessTier.SearchPermisoByRol", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function SavePermisoByRol(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SavePermisoByRolErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    SavePermisoByRol = False
    
    If Not oRs.EOF Then
    
        If DeletePermisoByRol(oRs("RolID"), lErrNum, sErrDesc, sErrSource) Then
        
            While Not oRs.EOF
            
                strSQL = "INSERT INTO roles_tablas_permisos (id_rol, tabla, permiso, fecha_alta) VALUES ( " & oRs("RolID") & ", '" & oRs("Tabla") & "', " & oRs("permiso") & ", " & oRs("fecha_alta") & " )"
                
                If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
                    If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                        SavePermisoByRol = True
                    End If
                End If
                
                oRs.MoveNext
                
            Wend
            
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        Else
            Err.Raise -1, "cAccessTier.SavePermisoByRol", "No se pudieron borrar los permisos anteriores."
        End If
    
    Else
        Err.Raise -1, "cAccessTier.SavePermisoByRol", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function
    
SavePermisoByRolErr:
    ShowError "cAccessTier.SavePermisoByRol", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Private Function DeletePermisoByRol(iRolID As Integer, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo DeletePermisoByRolErr
    
    DeletePermisoByRol = False
    
    strSQL = "DELETE FROM roles_tablas_permisos WHERE id_rol = " & iRolID
    
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
            DeletePermisoByRol = True
        End If
    
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
DeletePermisoByRolErr:
    ShowError "cAccessTier.DeletePermisoByRol", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'Métodos y Funciones para la clase cCategoria
'========================================================================
'========================================================================
'
'   Title          - Categorias
'   Description    - Metodos y funciones de trabajo para la clase cCategorias
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================

Public Function SaveCategoria(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SaveCategoriaErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    SaveCategoria = False
    
    If Not oRs.EOF Then
    
        strSQL = "INSERT INTO Categorias (descripcion, id_item_tipo, fecha_alta, fecha_ult_act) VALUES ( '" & oRs("descripcion") & "', " & oRs("itemtipoID") & ", " & oRs("fecha_alta") & ", " & oRs("fecha_ult_act") & " )"
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                SaveCategoria = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.SaveCategoria", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Set oRs = Nothing
    
    Exit Function
    
SaveCategoriaErr:
    ShowError "cAccessTier.SaveCategoria", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    
    Set oRs = Nothing
    
End Function

Public Function UpdateCategoria(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo UpdateCategoriaErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    UpdateCategoria = False
    
    If Not oRs.EOF Then
    
        strSQL = "UPDATE Categorias SET descripcion = '" & oRs("descripcion") & "', fecha_ult_act = " & oRs("fecha_ult_act") & " WHERE id_Categoria = " & oRs("id")
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                UpdateCategoria = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.UpdateCategoria", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

UpdateCategoriaErr:
    If CLng(Err.Number) <> 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "cAccessTier.UpdateCategoria"
    End If
End Function

Public Function DeleteCategoria(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo DeleteCategoriaErr
    
    Dim oRs As New Recordset
    Dim oRsAux As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    DeleteCategoria = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            strSQL = "DELETE FROM Categorias WHERE id_Categoria = " & oRs("id")
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                DeleteCategoria = True
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.DeleteCategoria", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Set oRs = Nothing
    Set oRsAux = Nothing
    
    Exit Function

DeleteCategoriaErr:
    ShowError "cAccessTier.DeleteCategoria", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function SearchCategoria(sXML As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchCategoriaErr
    Dim strAux
    
    Dim i As Integer
    
    SearchCategoria = False
    
    strSQL = "SELECT id_Categoria as ""ID"", categorias.descripcion as ""Descripcion"", items_tipos.id_item_tipo as ItemTipoID, items_tipos.descripcion as Tipo, fecha_alta as ""Fecha_Alta"", fecha_ult_act as ""Fecha_Ult_Act"" FROM Categorias, items_tipos "
    
    'caso particular...
    strAux = "WHERE items_tipos.id_item_tipo = categorias.id_item_tipo "
    'Fin caso particular
    
    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & strAux & " AND " & sSearch
            Else
                strSQL = strSQL & "WHERE " & sSearch
            End If
        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & " " & strAux
            End If
        End If
    End If

        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchCategoria = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchCategoriaErr:
    ShowError "cAccessTier.SearchCategoria", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'Métodos y Funciones para la clase cSubCategoria
'========================================================================
'========================================================================
'
'   Title          - SubCategorias
'   Description    - Metodos y funciones de trabajo para la clase cSubCategorias
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================

Public Function SaveSubCategoria(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SaveSubCategoriaErr
    
    Dim oRs As New Recordset
    Dim oRsAux As New Recordset
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    SaveSubCategoria = False
    
    If Not oRs.EOF Then
    
        strSQL = "INSERT INTO SubCategorias (descripcion, fecha_alta, fecha_ult_act) VALUES ( '" & oRs("descripcion") & "', " & oRs("fecha_alta") & ", " & oRs("fecha_ult_act") & " )"
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                strSQL = "SELECT id_SubCategoria FROM SubCategorias WHERE descripcion = '" & oRs("descripcion") & "'"
        
                If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
                    Set oRsAux = RecordsetFromXMLString(sXML)
                    If Not oRsAux.EOF Then
                        strSQL = "INSERT INTO categorias_subcategorias ( " & _
                          "id_categoria, id_subcategoria ) " & _
                          "VALUES " & _
                          "( " & oRs("categoriaID") & ", " & oRsAux("id_subcategoria") & " )"
                        If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                            SaveSubCategoria = True
                        End If
                    End If
                End If
                SaveSubCategoria = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.SaveSubCategoria", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Set oRs = Nothing
    Set oRsAux = Nothing
    
    Exit Function
    
SaveSubCategoriaErr:
    ShowError "cAccessTier.SaveSubCategoria", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    
    Set oRs = Nothing
    Set oRsAux = Nothing
    
End Function

Public Function UpdateSubCategoria(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo UpdateSubCategoriaErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    UpdateSubCategoria = False
    
    If Not oRs.EOF Then
    
        strSQL = "UPDATE SubCategorias SET descripcion = '" & oRs("descripcion") & "', fecha_ult_act = " & oRs("fecha_ult_act") & " WHERE id_SubCategoria = " & oRs("id")
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                UpdateSubCategoria = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.UpdateSubCategoria", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

UpdateSubCategoriaErr:
    If CLng(Err.Number) <> 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "cAccessTier.UpdateSubCategoria"
    End If
End Function

Public Function DeleteSubCategoria(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo DeleteSubCategoriaErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    DeleteSubCategoria = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            strSQL = "DELETE FROM SubCategorias WHERE id_SubCategoria = " & oRs("id")
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                strSQL = "DELETE FROM categorias_SubCategorias WHERE id_SubCategoria = " & oRs("id")
                If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                    DeleteSubCategoria = True
                End If
            End If
            
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.DeleteSubCategoria", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Set oRs = Nothing
    
    Exit Function

DeleteSubCategoriaErr:
    ShowError "cAccessTier.DeleteSubCategoria", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function SearchSubCategoria(sXML As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchSubCategoriaErr
    Dim strAux
    
    Dim i As Integer
    
    SearchSubCategoria = False
    
    strSQL = "SELECT SubCategorias.id_SubCategoria as ""ID"", " & _
             "SubCategorias.descripcion as ""Descripcion"", " & _
             "Categorias.descripcion as ""Categoria"", " & _
             "Categorias.id_categoria as ""CategoriaID"", SubCategorias.fecha_alta as ""Fecha_Alta"", SubCategorias.fecha_ult_act as ""Fecha_Ult_Act"" FROM SubCategorias, Categorias, categorias_subcategorias "
    
    'caso particular...
    strAux = "WHERE categorias.id_categoria = categorias_subcategorias.id_categoria AND subcategorias.id_subcategoria = categorias_subcategorias.id_subcategoria "
    'Fin caso particular
    
    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & strAux & " AND " & sSearch
            Else
                strSQL = strSQL & "WHERE " & sSearch
            End If
        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & " " & strAux
            End If
        End If
    End If

        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchSubCategoria = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchSubCategoriaErr:
    ShowError "cAccessTier.SearchSubCategoria", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'Public Function SearchSubCategoriaByCategoria(sXML As String, iCategoria As Integer, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
'    On Error GoTo SearchSubCategoriaByCategoriaErr
'
'    Dim i As Integer
'
'    SearchSubCategoriaByCategoria = False
'
'    strSQL = "SELECT SubCategorias.id_SubCategoria as ""ID"", " & _
'             "SubCategorias.descripcion as ""Descripcion""" & _
'             "FROM Categorias_SubCategorias " & _
'             "INNER JOIN Categorias ON (Categorias_SubCategorias.id_Categoria = Categorias.id_Categoria) " & _
'             "INNER JOIN SubCategorias ON (Categorias_SubCategorias.id_SubCategoria = SubCategorias.id_SubCategoria) " & _
'             "WHERE Categorias.id_Categoria = Categorias_SubCategorias.id_Categoria AND " & _
'             "Categorias.id_Categoria = " & CStr(iCategoria)
'
'    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
'
'        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
'            SearchSubCategoriaByCategoria = True
'        End If
'
'    End If
'
'    Call Disconnect(lErrNum, sErrDesc, sErrSource)
'
'    Exit Function
'
'
'SearchSubCategoriaByCategoriaErr:
'    ShowError "cAccessTier.SearchSubCategoriaByCategoria", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
'End Function

'========================================================================
'Métodos y Funciones para la clase cItem
'========================================================================
'========================================================================
'
'   Title          - Item
'   Description    - Metodos y funciones de trabajo para la clase cItem
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================

Public Function SaveItem(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SaveItemErr
    Dim oRs As New Recordset
    Dim oRsAux As New Recordset
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    SaveItem = False
    
    If Not oRs.EOF Then
    
        strSQL = "INSERT INTO Items ( titulo, autor, ISBN, " & _
                 "anno, id_editorial, id_ubicacion, id_item_tipo, fecha_alta, " & _
                 "fecha_ult_act ) VALUES ( " & _
                 "'" & oRs("titulo") & "', " & _
                 "'" & oRs("autor") & "', '" & oRs("ISBN") & "', " & _
                 "'" & oRs("anno") & "', " & oRs("editorialID") & ", " & _
                 oRs("ubicacionID") & ", " & oRs("itemtipoID") & ", " & oRs("fecha_alta") & ", " & oRs("fecha_ult_act") & " )"
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                strSQL = "SELECT id_Item FROM Items WHERE titulo = '" & oRs("titulo") & "' AND autor = '" & oRs("autor") & "' AND id_item_tipo = " & oRs("itemtipoID")
                If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
                    Set oRsAux = RecordsetFromXMLString(sXML)
                    If Not oRsAux.EOF Then
                        strSQL = "DELETE FROM Items_categorias_subcategorias WHERE id_Item = " & oRsAux("id_Item")
                        If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                            While Not oRsAux.EOF
                                strSQL = "INSERT INTO Items_categorias_subcategorias ( " & _
                                      "id_Item, id_categoria, id_subcategoria, fecha_alta, fecha_ult_act ) " & _
                                      "VALUES " & _
                                      "( " & oRsAux("id_Item") & ", " & oRs("categoriaID") & ", " & oRs("subcategoriaID") & ", " & oRs("fecha_alta") & ", " & oRs("fecha_ult_act") & ")"
                                SaveItem = Query(strSQL, lErrNum, sErrDesc, sErrSource)
                                If Not SaveItem Then oRs.MoveLast
                                oRsAux.MoveNext
                            Wend
                        End If
                    End If
                End If
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.SaveItem", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function
    
SaveItemErr:
    ShowError "cAccessTier.SaveItem", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function UpdateItem(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo UpdateItemErr
   
    Dim oRs As New Recordset
    Dim oRsAux As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    UpdateItem = False
    
    If Not oRs.EOF Then
    
        strSQL = "SELECT titulo, autor FROM Items WHERE id_Item = " & oRs("id")
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
                Set oRsAux = RecordsetFromXMLString(sXML)
            
                strSQL = "UPDATE Items SET " & _
                "titulo = '" & oRs("titulo") & "', " & _
                "autor = '" & oRs("autor") & "', ISBN = '" & oRs("ISBN") & "', anno = '" & oRs("anno") & "', " & _
                "id_item_tipo = " & oRs("itemtipoID") & ", id_editorial = " & oRs("editorialID") & ", id_ubicacion = " & oRs("ubicacionID") & ", " & _
                "fecha_ult_act = '" & oRs("fecha_ult_act") & "' " & _
                "WHERE titulo = '" & oRsAux("titulo") & "' AND autor = '" & oRsAux("autor") & "'"
        
                If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                    oRsAux.Close
                    
                    strSQL = "SELECT id_Item FROM Items WHERE titulo = '" & oRs("titulo") & "' AND autor = '" & oRs("autor") & "' AND id_item_tipo = " & oRs("itemtipoID")
                    If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
                        Set oRsAux = RecordsetFromXMLString(sXML)
                        If Not oRsAux.EOF Then
                            While Not oRsAux.EOF
                                strSQL = "DELETE FROM Items_categorias_subcategorias WHERE id_Item = " & oRsAux("id_Item")
                                If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                                    strSQL = "INSERT INTO Items_categorias_subcategorias ( " & _
                                          "id_Item, id_categoria, id_subcategoria, fecha_alta, fecha_ult_act ) " & _
                                          "VALUES " & _
                                          "( " & oRsAux("id_Item") & ", " & oRs("categoriaID") & ", " & oRs("subcategoriaID") & ", " & oRs("fecha_ult_act") & ", " & oRs("fecha_ult_act") & ")"
                                    UpdateItem = Query(strSQL, lErrNum, sErrDesc, sErrSource)
                                End If
                                If Not UpdateItem Then oRs.MoveLast
                                oRsAux.MoveNext
                            Wend
                        End If
                    End If
                End If
            End If
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.UpdateItem", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

UpdateItemErr:
    ShowError "cAccessTier.UpdateItem", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function UpdateStatusItem(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo UpdateStatusItemErr
   
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    UpdateStatusItem = False
    
    If Not oRs.EOF Then
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
                strSQL = "UPDATE Items SET " & _
                "prestado = " & IIf(oRs("prestado") = 0, "NULL", "1") & " " & _
                "WHERE id_Item = " & oRs("ItemID")
        
                UpdateStatusItem = Query(strSQL, lErrNum, sErrDesc, sErrSource)
                
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.UpdateStatusItem", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

UpdateStatusItemErr:
    ShowError "cAccessTier.UpdateStatusItem", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function DeleteItem(sXML As String, Optional iForever As Integer = 0, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo DeleteItemErr
    
    Dim oRs As New Recordset
    Dim oRsAux As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    DeleteItem = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            strSQL = "SELECT id_item as ID, titulo, autor, id_item_tipo FROM Items WHERE id_Item = " & oRs("id")
            If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
                
                Set oRsAux = RecordsetFromXMLString(CStr(sXML))
                
                If Not oRsAux.EOF Then
                    If iForever = 0 Then
                        strSQL = "UPDATE Items SET fecha_baja = " & Format(Now(), "yyyymmdd") & " WHERE titulo = '" & oRsAux("titulo") & "' AND autor = '" & oRsAux("autor") & "' AND id_item_tipo = " & oRsAux("id_item_tipo")
                    Else
                        strSQL = "DELETE FROM Items WHERE titulo = '" & oRsAux("titulo") & "' AND autor = '" & oRsAux("autor") & "' AND id_item_tipo = " & oRsAux("id_item_tipo")
                    End If
                    
                    If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                        If iForever = 1 Then
                            strSQL = "DELETE FROM Items_categorias_subcategorias WHERE id_Item = " & oRs("ID")
                            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                                DeleteItem = True
                            End If
                        Else
                            DeleteItem = True
                        End If
                    End If
                End If
                
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.DeleteItem", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Set oRs = Nothing
    Set oRsAux = Nothing
    
    Exit Function

DeleteItemErr:
    ShowError "cAccessTier.DeleteItem", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'   Función DeleteItemCopy()
'   Description    - Borra una copia del Item
'   Version Number - 1.0
'   Last Updated   - 21.03.2007
'   Author         - Algaze, Gastón
'
'   21.03.2007 - Creacion de la funcion.
'========================================================================
Public Function DeleteItemCopy(sXML As String, Optional iForever As Integer = 0, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo DeleteItemCopyErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    DeleteItemCopy = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            If iForever = 0 Then
                strSQL = "UPDATE Items SET fecha_baja = " & Format(Now(), "yyyymmdd") & " WHERE id_Item = " & oRs("id")
            Else
                strSQL = "DELETE FROM Items WHERE id_Item = " & oRs("id")
            End If
            
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                If iForever = 1 Then
                    strSQL = "DELETE FROM Items_categorias_subcategorias WHERE id_Item = " & oRs("ID")
                    If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                        DeleteItemCopy = True
                    End If
                Else
                    DeleteItemCopy = True
                End If
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.DeleteItemCopy", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

DeleteItemCopyErr:
    ShowError "cAccessTier.DeleteItemCopy", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function


Public Function RecoverItem(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo RecoverItemErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    RecoverItem = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            strSQL = "UPDATE Items SET fecha_baja = NULL WHERE id_Item = " & oRs("id")
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                RecoverItem = True
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.RecoverItem", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

RecoverItemErr:
    ShowError "cAccessTier.RecoverItem", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function SearchItem(sXML As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchItemErr
    
    Dim i As Integer
    Dim strAux_B As String
    Dim strSQL_B_1 As String
    Dim strSQL_B_2 As String
    
    SearchItem = False
    
    strSQL_B_1 = "SELECT Items.id_Item AS ID, items_tipos.Descripcion as Tipo, Items.titulo, Items.autor, " & _
             "Items.ISBN, Items.anno as Año, COUNT(*)  AS Copias, " & _
             "Items.prestado, items_tipos.id_item_tipo as ItemTipoID, Items.id_editorial AS editorialID,  Items.id_ubicacion AS ubicacionID, " & _
             "ubicaciones.descripcion AS ubicacion, editoriales.nombre AS editorial, " & _
             "categorias.descripcion AS Categoria, categorias.id_categoria AS categoriaID, " & _
             "subcategorias.descripcion AS subcategoria, subcategorias.id_subcategoria AS subcategoriaID, " & _
             "Items.fecha_alta, Items.fecha_ult_act, Items.fecha_baja " & _
             "FROM Items, ubicaciones, editoriales, categorias, subcategorias, Items_categorias_subcategorias, items_tipos "
    'caso particular A...
    strAux_B = "WHERE Items.id_ubicacion = ubicaciones.id_ubicacion AND " & _
             "Items.id_editorial = editoriales.id_editorial AND " & _
             "Items.id_Item = Items_categorias_subcategorias.id_Item AND " & _
             "categorias.id_categoria = Items_categorias_subcategorias.id_categoria AND " & _
             "subcategorias.id_subcategoria = Items_categorias_subcategorias.id_subcategoria AND " & _
             "items_tipos.id_item_tipo = items.id_item_tipo "
    'Fin caso particular
    
    strSQL_B_2 = "GROUP BY Items.titulo "


    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                
                    strSQL = strSQL_B_1 & strAux_B & " AND " & sSearch & " " & strSQL_B_2 & " ORDER BY " & sOrderBy & " " & sSort
            Else
                
                    strSQL = strSQL_B_1 & strAux_B & " AND " & sSearch & " " & strSQL_B_2 & " ORDER BY " & sOrderBy
                
            End If
        Else

                strSQL = strSQL_B_1 & strAux_B & " AND " & sSearch & " " & strSQL_B_2 & " "

        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
            
                    strSQL = strSQL_B_1 & strAux_B & " " & strSQL_B_2 & " " & " ORDER BY " & sOrderBy & " " & sSort
                
            Else
                
                    strSQL = strSQL_B_1 & strAux_B & " " & strSQL_B_2 & " " & " ORDER BY " & sOrderBy & " "

            End If
        Else
            
                strSQL = strSQL_B_1 & strAux_B & " " & strSQL_B_2 & " "

        End If
    End If

    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchItem = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchItemErr:
    ShowError "cAccessTier.SearchItem", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function SearchItemForBorrow(sXML As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchItemForBorrowErr
    
    Dim i As Integer
    Dim strAux_A As String
    Dim strAux_B As String
    Dim strAux_C As String
    Dim strSQL_A_1 As String
    Dim strSQL_A_2 As String
    Dim strSQL_B_1 As String
    Dim strSQL_B_2 As String
    Dim strSQL_C_1 As String
    Dim strSQL_C_2 As String
    
    SearchItemForBorrow = False
    
    strSQL_A_1 = "SELECT Items.id_Item AS ID, items_tipos.Descripcion as Tipo, Items.titulo, Items.autor, " & _
            "Items.ISBN, Items.anno as Año, COUNT(*)  AS Copias, " & _
            "Items.prestado, items_tipos.id_item_tipo as ItemTipoID, Items.id_editorial AS editorialID,  Items.id_ubicacion AS ubicacionID, " & _
            "ubicaciones.descripcion AS ubicacion, editoriales.nombre AS editorial, " & _
            "categorias.descripcion AS Categoria, categorias.id_categoria AS categoriaID, " & _
            "subcategorias.descripcion AS subcategoria, subcategorias.id_subcategoria AS subcategoriaID, " & _
            "Items.fecha_alta, Items.fecha_ult_act, Items.fecha_baja " & _
            "FROM Items, ubicaciones, editoriales, categorias, subcategorias, Items_categorias_subcategorias, items_tipos "
    
    'caso particular A...
    strAux_A = "WHERE Items.id_ubicacion = ubicaciones.id_ubicacion AND " & _
             "Items.id_editorial = editoriales.id_editorial AND " & _
             "Items.id_Item = Items_categorias_subcategorias.id_Item AND " & _
             "categorias.id_categoria = Items_categorias_subcategorias.id_categoria AND " & _
             "subcategorias.id_subcategoria = Items_categorias_subcategorias.id_subcategoria AND " & _
             "items_tipos.id_item_tipo = items.id_item_tipo AND " & _
             "Items.prestado IS NULL AND " & _
             "Items.id_Item NOT IN (SELECT reservas.id_Item FROM reservas WHERE fecha_reserva BETWEEN " & Format(DateAdd("ww", -1, Date), "yyyymmdd") & " AND " & Format(DateAdd("ww", 1, Date), "yyyymmdd") & " AND fecha_baja IS NULL ) "
    'Fin caso particular
    
    strSQL_A_2 = "GROUP BY Items.titulo UNION "

    strSQL_B_1 = "SELECT Items.id_Item AS ID, items_tipos.Descripcion as Tipo, Items.titulo, Items.autor, " & _
            "Items.ISBN, Items.anno as Año, -COUNT(*)  AS Copias, " & _
            "Items.prestado, items_tipos.id_item_tipo as ItemTipoID, Items.id_editorial AS editorialID,  Items.id_ubicacion AS ubicacionID, " & _
            "ubicaciones.descripcion AS ubicacion, editoriales.nombre AS editorial, " & _
            "categorias.descripcion AS Categoria, categorias.id_categoria AS categoriaID, " & _
            "subcategorias.descripcion AS subcategoria, subcategorias.id_subcategoria AS subcategoriaID, " & _
            "Items.fecha_alta, Items.fecha_ult_act, Items.fecha_baja " & _
            "FROM Items, ubicaciones, editoriales, categorias, subcategorias, Items_categorias_subcategorias, items_tipos "
    
    'caso particular B...
    strAux_B = "WHERE Items.id_ubicacion = ubicaciones.id_ubicacion AND " & _
            "Items.id_editorial = editoriales.id_editorial AND " & _
            "Items.id_Item = Items_categorias_subcategorias.id_Item AND " & _
            "categorias.id_categoria = Items_categorias_subcategorias.id_categoria AND " & _
            "subcategorias.id_subcategoria = Items_categorias_subcategorias.id_subcategoria AND " & _
            "items_tipos.id_item_tipo = items.id_item_tipo AND " & _
            "Items.prestado IS NOT NULL AND " & _
            "Items.id_Item IN (SELECT prestamos.id_Item FROM prestamos WHERE fecha_hasta > " & Format(DateAdd("ww", -1, Date), "yyyymmdd") & " AND fecha_baja IS NULL ) AND " & _
            "Items.id_Item IN (SELECT reservas.id_Item FROM reservas WHERE fecha_reserva BETWEEN " & Format(DateAdd("ww", -1, Date), "yyyymmdd") & " AND " & Format(DateAdd("ww", 1, Date), "yyyymmdd") & " AND fecha_baja IS NULL ) "
    'Fin caso particular
    
    strSQL_B_2 = "GROUP BY Items.titulo UNION "
    
    strSQL_C_1 = " SELECT Items.id_Item AS ID, items_tipos.Descripcion as Tipo, Items.titulo, Items.autor, " & _
            "Items.ISBN, Items.anno as Año, -COUNT(*)  AS Copias, " & _
            "Items.prestado, items_tipos.id_item_tipo as ItemTipoID, Items.id_editorial AS editorialID,  Items.id_ubicacion AS ubicacionID, " & _
            "ubicaciones.descripcion AS ubicacion, editoriales.nombre AS editorial, " & _
            "categorias.descripcion AS Categoria, categorias.id_categoria AS categoriaID, " & _
            "subcategorias.descripcion AS subcategoria, subcategorias.id_subcategoria AS subcategoriaID, " & _
            "Items.fecha_alta, Items.fecha_ult_act, Items.fecha_baja " & _
            "FROM Items, ubicaciones, editoriales, categorias, subcategorias, Items_categorias_subcategorias, items_tipos "
    
    'caso particular B...
    strAux_C = "WHERE Items.id_ubicacion = ubicaciones.id_ubicacion AND " & _
            "Items.id_editorial = editoriales.id_editorial AND " & _
            "Items.id_Item = Items_categorias_subcategorias.id_Item AND " & _
            "categorias.id_categoria = Items_categorias_subcategorias.id_categoria AND " & _
            "subcategorias.id_subcategoria = Items_categorias_subcategorias.id_subcategoria AND " & _
            "items_tipos.id_item_tipo = items.id_item_tipo AND " & _
            "Items.id_Item NOT IN ( " & _
            "SELECT Items.id_Item FROM Items " & _
            "WHERE Items.prestado IS NULL AND " & _
            "Items.id_Item NOT IN (SELECT reservas.id_Item FROM reservas WHERE fecha_reserva BETWEEN " & Format(DateAdd("ww", -1, Date), "yyyymmdd") & " AND " & Format(DateAdd("ww", 1, Date), "yyyymmdd") & " AND fecha_baja IS NULL ) " & _
            "UNION SELECT Items.id_Item FROM Items " & _
            "WHERE Items.prestado IS NOT NULL AND " & _
            "Items.id_Item IN (SELECT prestamos.id_Item FROM prestamos WHERE fecha_hasta > " & Format(DateAdd("ww", -1, Date), "yyyymmdd") & ") AND " & _
            "Items.id_Item IN (SELECT reservas.id_Item FROM reservas WHERE fecha_reserva BETWEEN " & Format(DateAdd("ww", -1, Date), "yyyymmdd") & " AND " & Format(DateAdd("ww", 1, Date), "yyyymmdd") & " AND fecha_baja IS NULL ) ) "
    'Fin caso particular
            
    strSQL_C_2 = "GROUP BY Items.titulo "

            


    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                strSQL = strSQL_A_1 & strAux_A & " AND " & sSearch & " " & strSQL_A_2 & " "
                strSQL = strSQL & strSQL_B_1 & strAux_B & " AND " & sSearch & " " & strSQL_B_2 & " "
                strSQL = strSQL & strSQL_C_1 & strAux_C & " AND " & sSearch & " " & strSQL_C_2 & " ORDER BY " & sOrderBy & " " & sSort
            Else
                strSQL = strSQL_A_1 & strAux_A & " AND " & sSearch & " " & strSQL_A_2 & " "
                strSQL = strSQL & strSQL_B_1 & strAux_B & " AND " & sSearch & " " & strSQL_B_2 & " "
                strSQL = strSQL & strSQL_C_1 & strAux_C & " AND " & sSearch & " " & strSQL_C_2 & " ORDER BY " & sOrderBy
            End If
        Else
            strSQL = strSQL_A_1 & strAux_A & " AND " & sSearch & " " & strSQL_A_2 & " "
            strSQL = strSQL & strSQL_B_1 & strAux_B & " AND " & sSearch & " " & strSQL_B_2 & " "
            strSQL = strSQL & strSQL_C_1 & strAux_C & " AND " & sSearch & " " & strSQL_C_2 & " ORDER BY " & sOrderBy
        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                strSQL = strSQL_A_1 & strAux_A & " " & strSQL_A_2 & " "
                strSQL = strSQL & strSQL_B_1 & strAux_B & " " & strSQL_B_2 & " "
                strSQL = strSQL & strSQL_C_1 & strAux_C & " " & strSQL_C_2 & " ORDER BY " & sOrderBy & " " & sSort
            Else
                strSQL = strSQL_A_1 & strAux_A & " " & strSQL_A_2 & " "
                strSQL = strSQL & strSQL_B_1 & strAux_B & " " & strSQL_B_2 & " "
                strSQL = strSQL & strSQL_C_1 & strAux_C & " " & strSQL_C_2 & " ORDER BY " & sOrderBy
            End If
        Else
            strSQL = strSQL_A_1 & strAux_A & " " & strSQL_A_2 & " "
            strSQL = strSQL & strSQL_B_1 & strAux_B & " " & strSQL_B_2 & " "
            strSQL = strSQL & strSQL_C_1 & strAux_C & " " & strSQL_C_2
        End If
    End If

    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchItemForBorrow = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchItemForBorrowErr:
    ShowError "cAccessTier.SearchItemForBorrow", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function SearchItemForReserve(sXML As String, sXMLData As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchItemForReserveErr
    
    Dim i As Integer
    Dim strAux_B As String
    Dim strSQL_B_1 As String
    Dim strSQL_B_2 As String
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXMLData))
    
    SearchItemForReserve = False
    
    strSQL_B_1 = "SELECT Items.id_Item AS ID, items_tipos.Descripcion as Tipo, Items.titulo, Items.autor, " & _
             "Items.ISBN, Items.anno as Año, COUNT(*)  AS Copias, " & _
             "Items.prestado, items_tipos.id_item_tipo as ItemTipoID, Items.id_editorial AS editorialID,  Items.id_ubicacion AS ubicacionID, " & _
             "ubicaciones.descripcion AS ubicacion, editoriales.nombre AS editorial, " & _
             "categorias.descripcion AS Categoria, categorias.id_categoria AS categoriaID, " & _
             "subcategorias.descripcion AS subcategoria, subcategorias.id_subcategoria AS subcategoriaID, " & _
             "Items.fecha_alta, Items.fecha_ult_act, Items.fecha_baja " & _
             "FROM Items, ubicaciones, editoriales, categorias, subcategorias, Items_categorias_subcategorias, items_tipos "
    'caso particular A...
    strAux_B = "WHERE Items.id_ubicacion = ubicaciones.id_ubicacion AND " & _
             "Items.id_editorial = editoriales.id_editorial AND " & _
             "Items.id_Item = Items_categorias_subcategorias.id_Item AND " & _
             "categorias.id_categoria = Items_categorias_subcategorias.id_categoria AND " & _
             "subcategorias.id_subcategoria = Items_categorias_subcategorias.id_subcategoria AND " & _
             "items_tipos.id_item_tipo = items.id_item_tipo AND " & _
             "Items.id_Item not in (select id_Item from prestamos WHERE fecha_hasta > " & Format(oRs("Fecha_Reserva"), "yyyymmdd") & " AND fecha_baja IS NULL ) AND " & _
             "Items.id_Item not IN ( " & _
             "SELECT reservas.id_Item FROM reservas WHERE " & _
             IIf(Not IsNull(oRs("usuarioID")), "id_usuario <> " & oRs("usuarioID"), "") & " AND fecha_reserva BETWEEN " & Format(DateAdd("ww", -1, oRs("Fecha_Reserva")), "yyyymmdd") & " AND " & Format(DateAdd("ww", 1, oRs("Fecha_Reserva")), "yyyymmdd") & " AND fecha_baja IS NULL ) AND " & _
             "Items.titulo = '" & oRs("titulo") & "' AND Items.autor = '" & oRs("autor") & "' AND Items.id_item_tipo = " & oRs("itemtipoID") & " "
             
    'Fin caso particular
    
    strSQL_B_2 = "GROUP BY Items.titulo "


    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                
                    strSQL = strSQL_B_1 & strAux_B & " AND " & sSearch & " " & strSQL_B_2 & " ORDER BY " & sOrderBy & " " & sSort
            Else
                
                    strSQL = strSQL_B_1 & strAux_B & " AND " & sSearch & " " & strSQL_B_2 & " ORDER BY " & sOrderBy
                
            End If
        Else

                strSQL = strSQL_B_1 & strAux_B & " AND " & sSearch & " " & strSQL_B_2 & " "

        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
            
                    strSQL = strSQL_B_1 & strAux_B & " " & strSQL_B_2 & " " & " ORDER BY " & sOrderBy & " " & sSort
                
            Else
                
                    strSQL = strSQL_B_1 & strAux_B & " " & strSQL_B_2 & " " & " ORDER BY " & sOrderBy & " "

            End If
        Else
            
                strSQL = strSQL_B_1 & strAux_B & " " & strSQL_B_2 & " "

        End If
    End If

    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchItemForReserve = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchItemForReserveErr:
    ShowError "cAccessTier.SearchItemForReserve", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function SearchItemForDelete(sXML As String, sXMLData As String, Optional iAll As Integer = 0, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchItemForDeleteErr
    
    Dim i As Integer
    Dim strAux_B As String
    Dim strSQL_B_1 As String
    Dim strSQL_B_2 As String
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXMLData))
    
    SearchItemForDelete = False
    
    If iAll = 0 Then
    
        strSQL_B_1 = "SELECT Items.id_Item AS ID, items_tipos.Descripcion as Tipo, Items.titulo, Items.autor, " & _
                 "Items.ISBN, Items.anno as Año, COUNT(*)  AS Copias, " & _
                 "Items.prestado, items_tipos.id_item_tipo as ItemTipoID, Items.id_editorial AS editorialID,  Items.id_ubicacion AS ubicacionID, " & _
                 "ubicaciones.descripcion AS ubicacion, editoriales.nombre AS editorial, " & _
                 "categorias.descripcion AS Categoria, categorias.id_categoria AS categoriaID, " & _
                 "subcategorias.descripcion AS subcategoria, subcategorias.id_subcategoria AS subcategoriaID, " & _
                 "Items.fecha_alta, Items.fecha_ult_act, Items.fecha_baja " & _
                 "FROM Items, ubicaciones, editoriales, categorias, subcategorias, Items_categorias_subcategorias, items_tipos "
        'caso particular A...
        strAux_B = "WHERE Items.id_ubicacion = ubicaciones.id_ubicacion AND " & _
                 "Items.id_editorial = editoriales.id_editorial AND " & _
                 "Items.id_Item = Items_categorias_subcategorias.id_Item AND " & _
                 "categorias.id_categoria = Items_categorias_subcategorias.id_categoria AND " & _
                 "subcategorias.id_subcategoria = Items_categorias_subcategorias.id_subcategoria AND " & _
                 "items_tipos.id_item_tipo = items.id_item_tipo AND " & _
                 "Items.id_Item not in (select id_Item from prestamos WHERE fecha_hasta > " & Format(Date, "yyyymmdd") & " AND fecha_baja IS NULL) AND " & _
                 "Items.id_Item not IN ( " & _
                 "SELECT reservas.id_Item FROM reservas WHERE " & _
                 " fecha_reserva > " & Format(Date, "yyyymmdd") & " AND fecha_baja IS NULL) AND " & _
                 "Items.titulo = '" & oRs("titulo") & "' AND Items.autor = '" & oRs("autor") & "' AND Items.id_item_tipo = " & oRs("itemtipoID") & " "
                 
        'Fin caso particular
        
        strSQL_B_2 = "GROUP BY Items.titulo "
    
    Else 'Borro todas las copias
    
        strSQL_B_1 = "SELECT Items.id_Item AS ID, items_tipos.Descripcion as Tipo, Items.titulo, Items.autor, " & _
                 "Items.ISBN, Items.anno as Año, COUNT(*)  AS Copias, " & _
                 "Items.prestado, items_tipos.id_item_tipo as ItemTipoID, Items.id_editorial AS editorialID,  Items.id_ubicacion AS ubicacionID, " & _
                 "ubicaciones.descripcion AS ubicacion, editoriales.nombre AS editorial, " & _
                 "categorias.descripcion AS Categoria, categorias.id_categoria AS categoriaID, " & _
                 "subcategorias.descripcion AS subcategoria, subcategorias.id_subcategoria AS subcategoriaID, " & _
                 "Items.fecha_alta, Items.fecha_ult_act, Items.fecha_baja " & _
                 "FROM Items, ubicaciones, editoriales, categorias, subcategorias, Items_categorias_subcategorias, items_tipos "
        'caso particular A...
        strAux_B = "WHERE Items.id_ubicacion = ubicaciones.id_ubicacion AND " & _
                 "Items.id_editorial = editoriales.id_editorial AND " & _
                 "Items.id_Item = Items_categorias_subcategorias.id_Item AND " & _
                 "categorias.id_categoria = Items_categorias_subcategorias.id_categoria AND " & _
                 "subcategorias.id_subcategoria = Items_categorias_subcategorias.id_subcategoria AND " & _
                 "items_tipos.id_item_tipo = items.id_item_tipo AND (" & _
                 "Items.id_Item IN (select id_Item from prestamos WHERE fecha_hasta > " & Format(Date, "yyyymmdd") & " AND fecha_baja IS NULL) OR " & _
                 "Items.id_Item IN ( " & _
                 "SELECT reservas.id_Item FROM reservas WHERE " & _
                 " fecha_reserva > " & Format(Date, "yyyymmdd") & " AND fecha_baja IS NULL) ) AND " & _
                 "Items.titulo = '" & oRs("titulo") & "' AND Items.autor = '" & oRs("autor") & "' AND Items.id_item_tipo = " & oRs("itemtipoID") & " "
                 
        'Fin caso particular
        
        strSQL_B_2 = "GROUP BY Items.titulo "
    
    End If


    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                
                    strSQL = strSQL_B_1 & strAux_B & " AND " & sSearch & " " & strSQL_B_2 & " ORDER BY " & sOrderBy & " " & sSort
            Else
                
                    strSQL = strSQL_B_1 & strAux_B & " AND " & sSearch & " " & strSQL_B_2 & " ORDER BY " & sOrderBy
                
            End If
        Else

                strSQL = strSQL_B_1 & strAux_B & " AND " & sSearch & " " & strSQL_B_2 & " "

        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
            
                    strSQL = strSQL_B_1 & strAux_B & " " & strSQL_B_2 & " " & " ORDER BY " & sOrderBy & " " & sSort
                
            Else
                
                    strSQL = strSQL_B_1 & strAux_B & " " & strSQL_B_2 & " " & " ORDER BY " & sOrderBy & " "

            End If
        Else
            
                strSQL = strSQL_B_1 & strAux_B & " " & strSQL_B_2 & " "

        End If
    End If

    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchItemForDelete = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchItemForDeleteErr:
    ShowError "cAccessTier.SearchItemForDelete", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'Métodos y Funciones para la clase cEditorial
'========================================================================
'========================================================================
'
'   Title          - Editoriales
'   Description    - Metodos y funciones de trabajo para la clase cEditorial
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SaveEditorial(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SaveEditorialErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    SaveEditorial = False
    
    If Not oRs.EOF Then
        strSQL = "INSERT INTO editoriales ( " & _
                 "nombre, mail, web, " & _
                 "domicilio_calle, domicilio_nro, domicilio_piso, " & _
                 "domicilio_unidad, domicilio_cod_postal, tel1, tel2, fecha_alta, " & _
                 "fecha_ult_act ) " & _
                 "VALUES ( '" & oRs("nombre") & "', " & _
                 "'" & oRs("mail") & "', '" & oRs("web") & "', " & _
                 "'" & oRs("domicilio_calle") & "', '" & oRs("domicilio_nro") & "', " & _
                 "'" & oRs("domicilio_piso") & "', '" & oRs("domicilio_unidad") & "', " & _
                 "'" & oRs("domicilio_cod_postal") & "', " & _
                 "'" & oRs("tel1") & "', '" & oRs("tel2") & "', " & _
                  oRs("fecha_alta") & ", " & oRs("fecha_ult_act") & " )"
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                SaveEditorial = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.SaveEditorial", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Set oRs = Nothing
    
    Exit Function
    
SaveEditorialErr:
    ShowError "cAccessTier.SaveEditorial", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function UpdateEditorial(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo UpdateEditorialErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    UpdateEditorial = False
    
    If Not oRs.EOF Then
     
        strSQL = "UPDATE Editoriales SET " & _
                "nombre = '" & oRs("nombre") & "', " & _
                "tel1 = '" & oRs("tel1") & "', " & _
                "tel2 = '" & oRs("tel2") & "', " & _
                "mail = '" & oRs("mail") & "', " & _
                "web = '" & oRs("web") & "', " & _
                "domicilio_calle = '" & oRs("domicilio_calle") & "', " & _
                "domicilio_nro = '" & oRs("domicilio_nro") & "', " & _
                "domicilio_piso = '" & oRs("domicilio_piso") & "', " & _
                "domicilio_unidad = '" & oRs("domicilio_unidad") & "', " & _
                "domicilio_cod_postal = '" & oRs("domicilio_cod_postal") & "', " & _
                "fecha_ult_act = " & oRs("fecha_ult_act") & " WHERE id_Editorial = " & oRs("ID")
                    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                    UpdateEditorial = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.UpdateEditorial", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

UpdateEditorialErr:
    ShowError "cAccessTier.UpdateEditorial", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function DeleteEditorial(sXML As String, Optional iForever As Integer = 0, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo DeleteEditorialErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    DeleteEditorial = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            If iForever = 0 Then
                strSQL = "UPDATE Editoriales SET fecha_baja = " & Format(Now(), "yyyymmdd") & " WHERE id_Editorial = " & oRs("id")
            Else
                strSQL = "DELETE FROM Editoriales WHERE id_Editorial = " & oRs("id")
            End If
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                DeleteEditorial = True
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.DeleteEditorial", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

DeleteEditorialErr:
    ShowError "cAccessTier.DeleteEditorial", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function RecoverEditorial(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo RecoverEditorialErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    RecoverEditorial = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            strSQL = "UPDATE Editoriales SET fecha_baja = NULL WHERE id_Editorial = " & oRs("id")
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                RecoverEditorial = True
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.RecoverEditorial", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

RecoverEditorialErr:
    If CLng(Err.Number) <> 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "cAccessTier.RecoverEditorial"
    End If
End Function

Public Function SearchEditorial(sXML As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchEditorialErr
    
    Dim i As Integer
    Dim strAux As String
    
    SearchEditorial = False
    
    strSQL = "SELECT editoriales.id_editorial as ""ID"", editoriales.nombre, " & _
             "editoriales.mail, " & _
             "editoriales.web, editoriales.domicilio_calle, " & _
             "editoriales.domicilio_nro, editoriales.domicilio_piso, " & _
             "editoriales.domicilio_unidad, " & _
             "editoriales.domicilio_cod_postal, editoriales.tel1, editoriales.tel2, editoriales.fecha_alta, " & _
             "editoriales.fecha_ult_act, editoriales.fecha_baja " & _
             "FROM editoriales "
    
    'caso particular...
    strAux = ""
    'Fin caso particular
    
    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & strAux & " AND " & sSearch
            Else
                strSQL = strSQL & "WHERE " & sSearch
            End If
        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & " " & strAux
            End If
        End If
    End If
        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchEditorial = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchEditorialErr:
    ShowError "cAccessTier.SearchEditorial", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'
'   Title          - Prestamos
'   Description    - Metodos y funciones de trabajo para la clase cPrestamo
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SavePrestamo(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SavePrestamoErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    SavePrestamo = False
    
    If Not oRs.EOF Then
    
        strSQL = "INSERT INTO prestamos( fecha_desde,  fecha_hasta, " & _
                 "id_usuario,  id_bibliotecaria,  id_Item, id_prestamos_tipos  " & _
                 ") VALUES (  " & oRs("fecha_desde") & ",  " & _
                 oRs("fecha_hasta") & ",  " & oRs("usuarioID") & ",  " & _
                 oRs("bibliotecariaID") & ",  " & oRs("ItemID") & ",  " & oRs("tipo_prestamoID") & ")"
                 
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                strSQL = "UPDATE Items SET " & _
                "prestado = 1 " & _
                "WHERE id_Item = " & oRs("ItemID")
        
                If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                    SavePrestamo = True
                End If
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.SavePrestamo", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function
    
SavePrestamoErr:
    ShowError "cAccessTier.SavePrestamo", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function UpdatePrestamo(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo UpdatePrestamoErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    UpdatePrestamo = False
    
    If Not oRs.EOF Then
        strSQL = "UPDATE prestamos SET  fecha_desde = " & oRs("fecha_desde") & _
        ",  fecha_hasta = " & oRs("fecha_hasta") & _
        ",  id_usuario = " & oRs("usuarioID") & _
        ",  id_bibliotecaria = " & oRs("bibliotecariaID") & _
        ",  id_Item = " & oRs("ItemID") & _
        ",  id_prestamos_tipos = " & oRs("tipo_prestamoID") & _
        ",  fecha_devolucion = " & oRs("fecha_devolucion") & _
        " WHERE id_Prestamo = " & oRs("id")
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                UpdatePrestamo = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.UpdatePrestamo", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

UpdatePrestamoErr:
    ShowError "cAccessTier.UpdatePrestamo", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function SearchPrestamo(sXML As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchPrestamoErr
    
    Dim i As Integer
    Dim strAux As String
    
    SearchPrestamo = False
    strSQL = "SELECT prestamos.id_prestamo as ID, Items_Tipos.Descripcion as Tipo, Items.titulo, Items.autor, prestamos.fecha_desde, prestamos.fecha_hasta, CONCAT(usuarios.nombre,' ',usuarios.apellido) as Nombre, usuarios.mail, usuarios.tel1, items_tipos.id_item_tipo as ItemTipoID, prestamos.id_usuario as ""usuarioID"",  prestamos.id_bibliotecaria as ""bibliotecariaID"",  prestamos.id_Item as ""ItemID"",  prestamos.fecha_devolucion, prestamos.id_prestamos_tipos as ""tipo_prestamoID"" FROM prestamos INNER JOIN usuarios ON (prestamos.id_usuario = usuarios.id_usuario) INNER JOIN usuarios bibliotecarios ON (prestamos.id_bibliotecaria = bibliotecarios.id_usuario) INNER JOIN Items ON (prestamos.id_Item = Items.id_Item) INNER JOIN items_tipos ON (items_tipos.id_item_tipo = items.id_item_tipo) "
    
    'caso particular...
    strAux = ""
    'Fin caso particular
    
    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & strAux & " AND " & sSearch
            Else
                strSQL = strSQL & "WHERE " & sSearch
            End If
        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & " " & strAux
            End If
        End If
    End If

        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchPrestamo = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchPrestamoErr:
    ShowError "cAccessTier.SearchPrestamo", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function SearchPrestamoForReport(sXML As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchPrestamoForReportErr
    
    Dim i As Integer
    Dim strAux As String
    
    SearchPrestamoForReport = False
    strSQL = "SELECT Items_Tipos.Descripcion as Tipo, Items.titulo as Título, Items.autor as Autor, items.ISBN, prestamos.fecha_desde as 'Prestado Desde', prestamos.fecha_hasta as 'Prestado Hasta', CONCAT(usuarios.nombre,' ',usuarios.apellido) as Usuario, usuarios.mail as Email, usuarios.tel1 as Telefono  FROM prestamos INNER JOIN usuarios ON (prestamos.id_usuario = usuarios.id_usuario) INNER JOIN usuarios bibliotecarios ON (prestamos.id_bibliotecaria = bibliotecarios.id_usuario) INNER JOIN Items ON (prestamos.id_Item = Items.id_Item) INNER JOIN items_tipos ON (items_tipos.id_item_tipo = items.id_item_tipo) "
    
    'caso particular...
    strAux = ""
    'Fin caso particular
    
    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & strAux & " AND " & sSearch
            Else
                strSQL = strSQL & "WHERE " & sSearch
            End If
        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & " " & strAux
            End If
        End If
    End If

        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchPrestamoForReport = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchPrestamoForReportErr:
    ShowError "cAccessTier.SearchPrestamoForReport", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'
'   Title          - Seguridad
'   Description    - Metodos y Funciones de trabajo para la clase cSecurityAgent
'   Version Number - 1.0
'   Last Updated   - February 26th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'   Notas          - 21.03.2007
'                    Se comienza la modificacion del modulo de seguridad
'                    para evitar hardcodeos.
'
'========================================================================
Public Function LoginUsuario(sXML As String, sUsername As String, sPassword As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo LoginUsuarioErr
    
    LoginUsuario = False
    
    strSQL = "SELECT usuarios.id_usuario as ""ID"", CONCAT(usuarios.nombre,' ',usuarios.apellido) as nombre_completo, roles.descripcion as Rol FROM usuarios, roles_usuarios, roles " & _
             "WHERE usuarios.username = '" & sUsername & "' AND usuarios.password = MD5(CONCAT(usuarios.id_usuario, '" & sPassword & "')) AND roles.id_rol = roles_usuarios.id_rol AND roles_usuarios.id_usuario = usuarios.id_usuario"
        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            LoginUsuario = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function

LoginUsuarioErr:
    ShowError "cAccessTier.LoginUsuario", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function GetPrivileges(sXML As String, sPrivilegesXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo GetPrivilegesErr
    Dim strXML As String
    Dim oRs As New Recordset
    
    GetPrivileges = False
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    strSQL = "SELECT roles_tablas_permisos.permiso " & _
             "FROM roles_tablas_permisos " & _
             "INNER JOIN roles ON (roles_tablas_permisos.id_rol = roles.id_rol) " & _
             "INNER JOIN roles_usuarios ON (roles.id_rol = roles_usuarios.id_rol) " & _
             "INNER JOIN usuarios ON (roles_usuarios.id_usuario = usuarios.id_usuario) " & _
             "WHERE usuarios.id_usuario = " & oRs("userID") & " AND roles_tablas_permisos.tabla = '" & oRs("tabla") & "'"
        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, strXML, lErrNum, sErrDesc, sErrSource) Then
            sPrivilegesXML = strXML
            GetPrivileges = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Set oRs = Nothing
    
    Exit Function

GetPrivilegesErr:
    ShowError "cAccessTier.GetPrivileges", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function GetGlobalPrivileges(sXML As String, sPrivilegesXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo GetGlobalPrivilegesErr
    Dim strXML As String
    Dim oRs As New Recordset
    
    GetGlobalPrivileges = False
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    strSQL = "SELECT roles.privilegios_globales as permiso " & _
             "FROM roles " & _
             "INNER JOIN roles_usuarios ON (roles.id_rol = roles_usuarios.id_rol) " & _
             "INNER JOIN usuarios ON (roles_usuarios.id_usuario = usuarios.id_usuario) " & _
             "WHERE usuarios.id_usuario = " & oRs("userID")
        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, strXML, lErrNum, sErrDesc, sErrSource) Then
            sPrivilegesXML = strXML
            GetGlobalPrivileges = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Set oRs = Nothing
    
    Exit Function

GetGlobalPrivilegesErr:
    ShowError "cAccessTier.GetGlobalPrivileges", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function GetRestrictions(sXML As String, sRestrictionsXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo GetRestrictionsErr
    Dim strXML As String
    Dim oRs As New Recordset
    
    GetRestrictions = False
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    If oRs("tabla") = "All" Then
    
        strSQL = "SELECT restricciones.tabla, restricciones.campo, `operaciones`.descripcion as Operacion, " & _
                 "restricciones.valor FROM `operaciones`  " & _
                 "INNER JOIN `restricciones` ON (`operaciones`.id_operacion = `restricciones`.id_operacion) " & _
                 "INNER JOIN `roles_restricciones` ON (`restricciones`.id_restriccion = `roles_restricciones`.id_restriccion) " & _
                 "INNER JOIN `roles` ON (`roles_restricciones`.id_rol = `roles`.id_rol) " & _
                 "INNER JOIN `roles_usuarios` ON (`roles`.id_rol = `roles_usuarios`.id_rol) " & _
                 "INNER JOIN usuarios ON (`roles_usuarios`.id_usuario = usuarios.id_usuario) " & _
                 "WHERE usuarios.id_usuario = " & oRs("userID")
    Else
        strSQL = "SELECT restricciones.tabla, restricciones.campo, `operaciones`.descripcion as Operacion, " & _
                 "restricciones.valor FROM `operaciones`  " & _
                 "INNER JOIN `restricciones` ON (`operaciones`.id_operacion = `restricciones`.id_operacion) " & _
                 "INNER JOIN `roles_restricciones` ON (`restricciones`.id_restriccion = `roles_restricciones`.id_restriccion) " & _
                 "INNER JOIN `roles` ON (`roles_restricciones`.id_rol = `roles`.id_rol) " & _
                 "INNER JOIN `roles_usuarios` ON (`roles`.id_rol = `roles_usuarios`.id_rol) " & _
                 "INNER JOIN usuarios ON (`roles_usuarios`.id_usuario = usuarios.id_usuario) " & _
                 "WHERE usuarios.id_usuario = " & oRs("userID") & " AND " & _
                 "restricciones.tabla = '" & oRs("tabla") & "'"
    End If
        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, strXML, lErrNum, sErrDesc, sErrSource) Then
            sRestrictionsXML = strXML
            GetRestrictions = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Set oRs = Nothing
    
    Exit Function

GetRestrictionsErr:
    ShowError "cAccessTier.GetRestrictions", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'   Función GetColumns()
'   Description    - Trae todos los campos de una determinada tabla
'   Version Number - 1.0
'   Last Updated   - 21.03.2007
'   Author         - Algaze, Gastón
'
'   21.03.2007 - Creacion de la funcion.
'========================================================================
Public Function GetColumns(sTable As String, sColumnsXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo GetColumnsErr
    Dim strXML As String
    
    GetColumns = False
    
    strSQL = "SHOW COLUMNS FROM " & sTable & " FROM biblos"
        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, strXML, lErrNum, sErrDesc, sErrSource) Then
            sColumnsXML = strXML
            GetColumns = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function

GetColumnsErr:
    ShowError "cAccessTier.GetColumns", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'   Función GetTables()
'   Description    - Trae todas las tablas
'   Version Number - 1.0
'   Last Updated   - 21.03.2007
'   Author         - Algaze, Gastón
'
'   21.03.2007 - Creacion de la funcion.
'========================================================================
Public Function GetTables(sTablesXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo GetTablesErr
    Dim strXML As String
    
    GetTables = False
    
    strSQL = "SHOW Tables FROM biblos"
        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, strXML, lErrNum, sErrDesc, sErrSource) Then
            sTablesXML = strXML
            GetTables = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function

GetTablesErr:
    ShowError "cAccessTier.GetTables", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'Métodos y Funciones para la clase cPrestamoTipo
'========================================================================
'========================================================================
'
'   Title          - PrestamoTipo
'   Description    - Metodos y funciones de trabajo para la clase cPrestamoTipo
'   Version Number - 1.0
'   Last Updated   - March 21th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================

Public Function SavePrestamoTipo(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SavePrestamoTipoErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    SavePrestamoTipo = False
    
    If Not oRs.EOF Then
    
        strSQL = "INSERT INTO Prestamos_Tipos (descripcion, duracion, fecha_alta, fecha_ult_act) VALUES ( '" & oRs("descripcion") & "', " & oRs("duracion") & ", " & oRs("fecha_alta") & ", " & oRs("fecha_ult_act") & " )"
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                SavePrestamoTipo = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.SavePrestamoTipo", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Set oRs = Nothing
    
    Exit Function
    
SavePrestamoTipoErr:
    ShowError "cAccessTier.SavePrestamoTipo", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    
    Set oRs = Nothing
    
End Function

Public Function UpdatePrestamoTipo(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo UpdatePrestamoTipoErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    UpdatePrestamoTipo = False
    
    If Not oRs.EOF Then
    
        strSQL = "UPDATE prestamos_tipos SET descripcion = '" & oRs("descripcion") & "', duracion = " & oRs("duracion") & ", fecha_ult_act = " & oRs("fecha_ult_act") & " WHERE id_Prestamos_Tipos = " & oRs("id")
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                UpdatePrestamoTipo = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.UpdatePrestamoTipo", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

UpdatePrestamoTipoErr:
    If CLng(Err.Number) <> 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "cAccessTier.UpdatePrestamoTipo"
    End If
End Function

Public Function DeletePrestamoTipo(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo DeletePrestamoTipoErr
    
    Dim oRs As New Recordset
    Dim oRsAux As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    DeletePrestamoTipo = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            strSQL = "DELETE FROM prestamos_tipos WHERE id_Prestamos_Tipos = " & oRs("id")
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                DeletePrestamoTipo = True
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.DeletePrestamoTipo", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Set oRs = Nothing
    Set oRsAux = Nothing
    
    Exit Function

DeletePrestamoTipoErr:
    ShowError "cAccessTier.DeletePrestamoTipo", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function SearchPrestamoTipo(sXML As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchPrestamoTipoErr
    Dim strAux
    
    Dim i As Integer
    
    SearchPrestamoTipo = False
    
    strSQL = "SELECT id_Prestamos_Tipos as ""ID"", descripcion as ""Descripcion"", duracion as ""Duracion"", fecha_alta as ""Fecha_Alta"", fecha_ult_act as ""Fecha_Ult_Act"" FROM prestamos_tipos "
    
    'caso particular...
    strAux = ""
    'Fin caso particular
    
    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & strAux & " AND " & sSearch
            Else
                strSQL = strSQL & "WHERE " & sSearch
            End If
        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & " " & strAux
            End If
        End If
    End If

        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchPrestamoTipo = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchPrestamoTipoErr:
    ShowError "cAccessTier.SearchPrestamoTipo", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function DeletePrestamo(sXML As String, Optional iForever As Integer = 0, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo DeletePrestamoErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    DeletePrestamo = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            If iForever = 0 Then
                strSQL = "UPDATE prestamos SET fecha_baja = " & Format(Now(), "yyyymmdd") & " WHERE id_prestamo = " & oRs("id")
            Else
                strSQL = "DELETE FROM prestamos WHERE id_prestamo = " & oRs("id")
            End If
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                DeletePrestamo = True
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.DeletePrestamo", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

DeletePrestamoErr:
    ShowError "cAccessTier.DeletePrestamo", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'
'   Title          - Reservas
'   Description    - Metodos y funciones de trabajo para la clase cReserva
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SaveReserva(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SaveReservaErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    SaveReserva = False
    
    If Not oRs.EOF Then
    
        strSQL = "INSERT INTO Reservas( fecha_reserva, " & _
                 "id_usuario, id_Item " & _
                 ") VALUES (  " & oRs("fecha_reserva") & ",  " & _
                 oRs("usuarioID") & ",  " & _
                 oRs("ItemID") & " )"
                 
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            SaveReserva = Query(strSQL, lErrNum, sErrDesc, sErrSource)
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.SaveReserva", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function
    
SaveReservaErr:
    ShowError "cAccessTier.SaveReserva", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function UpdateReserva(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo UpdateReservaErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    UpdateReserva = False
    
    If Not oRs.EOF Then
        strSQL = "UPDATE Reservas SET  fecha_reserva = " & oRs("fecha_reserva") & _
        ",  id_usuario = " & oRs("usuarioID") & _
        ",  id_Item = " & oRs("ItemID") & _
        " WHERE id_Reserva = " & oRs("id")
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                UpdateReserva = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.UpdateReserva", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

UpdateReservaErr:
    ShowError "cAccessTier.UpdateReserva", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function SearchReserva(sXML As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchReservaErr
    
    Dim i As Integer
    Dim strAux As String
    
    SearchReserva = False
    strSQL = "SELECT  Reservas.id_Reserva as ""ID"", Items.titulo, Items.autor, Items.ISBN, Reservas.fecha_reserva, CONCAT(usuarios.nombre,' ',usuarios.apellido) as Nombre, usuarios.matricula, usuarios.apellido, Reservas.id_usuario as ""usuarioID"", Reservas.id_Item as ""ItemID"", items_tipos.id_item_tipo as ItemTipoID FROM Reservas, Items, usuarios, items_tipos "
    
    'caso particular...
    strAux = "WHERE Items.id_Item = reservas.id_Item AND usuarios.id_usuario = reservas.id_usuario AND items_tipos.id_item_tipo = items.id_item_tipo "
    'Fin caso particular
    
    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & strAux & " AND " & sSearch
            Else
                strSQL = strSQL & "WHERE " & sSearch
            End If
        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & " " & strAux
            End If
        End If
    End If

        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchReserva = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchReservaErr:
    ShowError "cAccessTier.SearchReserva", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function DeleteReserva(sXML As String, Optional iForever As Integer = 0, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo DeleteReservaErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    DeleteReserva = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            If iForever = 0 Then
                strSQL = "UPDATE reservas SET fecha_baja = " & Format(Now(), "yyyymmdd") & " WHERE id_reserva = " & oRs("id")
            Else
                strSQL = "DELETE FROM reservas WHERE id_reserva = " & oRs("id")
            End If
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                DeleteReserva = True
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.DeleteReserva", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

DeleteReservaErr:
    ShowError "cAccessTier.DeleteReserva", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'
'   Title          - Item_Tipos
'   Description    - Metodos y funciones de trabajo para la clase cItemTipo
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SaveItemTipo(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SaveItemTipoErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    SaveItemTipo = False
    
    If Not oRs.EOF Then
    
        strSQL = "INSERT INTO Items_tipos (descripcion) VALUES ( '" & oRs("descripcion") & "' )"
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                SaveItemTipo = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.SaveItemTipo", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function
    
SaveItemTipoErr:
    ShowError "cAccessTier.SaveItemTipo", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function UpdateItemTipo(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo UpdateItemTipoErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    UpdateItemTipo = False
    
    If Not oRs.EOF Then
    
        strSQL = "UPDATE Items_tipos SET descripcion = '" & oRs("descripcion") & "' WHERE id_item_tipo = " & oRs("id")
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                UpdateItemTipo = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.UpdateItemTipo", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

UpdateItemTipoErr:
    ShowError "cAccessTier.UpdateItemTipo", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function DeleteItemTipo(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo DeleteItemTipoErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    DeleteItemTipo = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            
            strSQL = "DELETE FROM Items_tipos WHERE id_item_tipo = " & oRs("id")
          
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                DeleteItemTipo = True
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.DeleteItemTipo", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

DeleteItemTipoErr:
    ShowError "cAccessTier.DeleteItemTipo", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function SearchItemTipo(sXML As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchItemTipoErr
    
    Dim i As Integer
    Dim strAux As String
    
    SearchItemTipo = False
    
    strSQL = "SELECT id_item_tipo as ""ID"", descripcion as ""Descripcion"" FROM Items_tipos "
    
    'caso particular...
    strAux = ""
    'Fin caso particular
    
    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & strAux & " AND " & sSearch
            Else
                strSQL = strSQL & "WHERE " & sSearch
            End If
        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & " " & strAux
            End If
        End If
    End If

        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchItemTipo = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchItemTipoErr:
    ShowError "cAccessTier.SearchItemTipo", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'
'   Title          - Fichas
'   Description    - Metodos y funciones de trabajo para la clase cFicha
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SaveFicha(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SaveFichaErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    SaveFicha = False
    
    If Not oRs.EOF Then
    
        strSQL = "INSERT INTO Fichas (archivo, titulo, archivo_nombre, id_usuario, fecha_alta, fecha_ult_act) VALUES ( '" & oRs("archivo") & "', '" & oRs("titulo") & "', '" & oRs("archivo_nombre") & "', " & oRs("usuarioID") & ", " & Format(Now(), "yyyymmdd") & ", " & Format(Now(), "yyyymmdd") & " )"
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                SaveFicha = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.SaveFicha", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function
    
SaveFichaErr:
    ShowError "cAccessTier.SaveFicha", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function UpdateFicha(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo UpdateFichaErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    UpdateFicha = False
    
    If Not oRs.EOF Then
    
        strSQL = "UPDATE Fichas SET titulo = '" & oRs("titulo") & "', id_usuario = " & oRs("usuarioID") & ", fecha_ult_act = " & Format(Now(), "yyyymmdd") & " WHERE id_Ficha = " & oRs("id")
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                UpdateFicha = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.UpdateFicha", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

UpdateFichaErr:
    ShowError "cAccessTier.UpdateFicha", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function DeleteFicha(sXML As String, Optional iForever As Integer = 0, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo DeleteFichaErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    DeleteFicha = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            If iForever = 0 Then
                strSQL = "UPDATE Fichas SET fecha_baja = " & Format(Now(), "yyyymmdd") & " WHERE id_Ficha = " & oRs("id")
            Else
                strSQL = "DELETE FROM Fichas WHERE id_Ficha = " & oRs("id")
            End If
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                DeleteFicha = True
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.DeleteFicha", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

DeleteFichaErr:
    ShowError "cAccessTier.DeleteFicha", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function RecoverFicha(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo RecoverFichaErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    RecoverFicha = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            strSQL = "UPDATE Fichas SET fecha_baja = NULL WHERE id_Ficha = " & oRs("id")
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                RecoverFicha = True
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.RecoverFicha", "El archivo XML de entrada se encuentra vacío."
    End If
    
    Exit Function

RecoverFichaErr:
    ShowError "cAccessTier.RecoverFicha", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function SearchFicha(sXML As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchFichaErr
    
    Dim i As Integer
    Dim strAux As String
    
    SearchFicha = False
    
    strSQL = "SELECT id_Ficha as ""ID"", titulo, archivo_nombre as 'Nombre del Archivo', id_usuario as usuarioID, NULL as archivo, archivo_nombre, fecha_alta, fecha_ult_act, fecha_baja as ""Fecha_Baja"" FROM Fichas "
    
    'caso particular...
    strAux = ""
    'Fin caso particular
    
    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & strAux & " AND " & sSearch
            Else
                strSQL = strSQL & "WHERE " & sSearch
            End If
        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & " " & strAux
            End If
        End If
    End If

        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchFicha = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchFichaErr:
    ShowError "cAccessTier.SearchFicha", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function GetFileFicha(sXML As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo GetFileFichaErr
    
    Dim i As Integer
    Dim strAux As String
    
    GetFileFicha = False
    
    strSQL = "SELECT id_Ficha as ""ID"", archivo, titulo, archivo_nombre, id_usuario as usuarioID, fecha_alta, fecha_ult_act, fecha_baja as ""Fecha_Baja"" FROM Fichas "
    
    'caso particular...
    strAux = ""
    'Fin caso particular
    
    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & strAux & " AND " & sSearch
            Else
                strSQL = strSQL & "WHERE " & sSearch
            End If
        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & " " & strAux
            End If
        End If
    End If

        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            GetFileFicha = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

GetFileFichaErr:
    ShowError "cAccessTier.GetFileFicha", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

'========================================================================
'
'   Title          - Links
'   Description    - Metodos y funciones de trabajo para la clase cLink
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SaveLink(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SaveLinkErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(sXML)
    
    SaveLink = False
    
    If Not oRs.EOF Then
    
        strSQL = "INSERT INTO Links (id_subcategoria, link, descripcion, id_usuario, fecha_alta, fecha_ult_act) VALUES ( " & oRs("subcategoriaID") & ", '" & oRs("link") & "', '" & oRs("descripcion") & "', " & oRs("usuarioID") & ", " & Format(Now(), "yyyymmdd") & ", " & Format(Now(), "yyyymmdd") & " )"
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                SaveLink = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.SaveLink", "El id_subcategoria XML de entrada se encuentra vacío."
    End If
    
    Exit Function
    
SaveLinkErr:
    ShowError "cAccessTier.SaveLink", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function UpdateLink(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo UpdateLinkErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    UpdateLink = False
    
    If Not oRs.EOF Then
    
        strSQL = "UPDATE Links SET link = '" & oRs("link") & "', descripcion = '" & oRs("descripcion") & "', id_usuario = " & oRs("usuarioID") & ", id_subcategoria = " & oRs("subcategoriaID") & ", fecha_ult_act = " & Format(Now(), "yyyymmdd") & " WHERE id_Link = " & oRs("id")
        
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                UpdateLink = True
            End If
        
        End If
        
        Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Else
        Err.Raise -1, "cAccessTier.UpdateLink", "El id_subcategoria XML de entrada se encuentra vacío."
    End If
    
    Exit Function

UpdateLinkErr:
    ShowError "cAccessTier.UpdateLink", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function DeleteLink(sXML As String, Optional iForever As Integer = 0, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo DeleteLinkErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    DeleteLink = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            If iForever = 0 Then
                strSQL = "UPDATE Links SET fecha_baja = " & Format(Now(), "yyyymmdd") & " WHERE id_Link = " & oRs("id")
            Else
                strSQL = "DELETE FROM Links WHERE id_Link = " & oRs("id")
            End If
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                DeleteLink = True
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.DeleteLink", "El id_subcategoria XML de entrada se encuentra vacío."
    End If
    
    Exit Function

DeleteLinkErr:
    ShowError "cAccessTier.DeleteLink", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function RecoverLink(sXML As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo RecoverLinkErr
    
    Dim oRs As New Recordset
    
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    
    RecoverLink = False
    
    If Not oRs.EOF Then
    
        If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
            
            strSQL = "UPDATE Links SET fecha_baja = NULL WHERE id_Link = " & oRs("id")
            
            If Query(strSQL, lErrNum, sErrDesc, sErrSource) Then
                RecoverLink = True
            End If
        
            Call Disconnect(lErrNum, sErrDesc, sErrSource)
        
        End If
    
    Else
        Err.Raise -1, "cAccessTier.RecoverLink", "El id_subcategoria XML de entrada se encuentra vacío."
    End If
    
    Exit Function

RecoverLinkErr:
    ShowError "cAccessTier.RecoverLink", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

Public Function SearchLink(sXML As String, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    On Error GoTo SearchLinkErr
    
    Dim i As Integer
    Dim strAux As String
    
    SearchLink = False
    
    strSQL = "SELECT id_Link as ""ID"", links.descripcion, link, items_tipos.descripcion as tipo, categorias.descripcion as categoria, subcategorias.descripcion as subcategoria, links.id_usuario as usuarioID, items_tipos.id_item_tipo as itemtipoID, categorias.id_categoria as categoriaID, links.id_subcategoria as subcategoriaID, links.fecha_alta, links.fecha_ult_act, links.fecha_baja FROM Links, categorias, subcategorias, categorias_subcategorias, items_tipos "
    
    'caso particular...
    strAux = "WHERE categorias.id_item_tipo = items_tipos.id_item_tipo AND links.id_subcategoria = subcategorias.id_subcategoria AND categorias_subcategorias.id_categoria = categorias.id_categoria AND categorias_subcategorias.id_subcategoria = links.id_subcategoria"
    'Fin caso particular
    
    If Len(sSearch) > 0 Then
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & strAux & " AND " & sSearch & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & "WHERE " & sSearch & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & strAux & " AND " & sSearch
            Else
                strSQL = strSQL & "WHERE " & sSearch
            End If
        End If
    Else
        If Len(sOrderBy) > 0 Then
            If Len(sSort) > 0 Then
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy & " " & sSort
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy & " " & sSort
                End If
            Else
                If Len(strAux) > 0 Then
                    strSQL = strSQL & " " & strAux & " ORDER BY " & sOrderBy
                Else
                    strSQL = strSQL & " ORDER BY " & sOrderBy
                End If
            End If
        Else
            If Len(strAux) > 0 Then
                strSQL = strSQL & " " & strAux
            End If
        End If
    End If

        
    If Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
    
        If RetrieveData(strSQL, sXML, lErrNum, sErrDesc, sErrSource) Then
            SearchLink = True
        End If
        
    End If
    
    Call Disconnect(lErrNum, sErrDesc, sErrSource)
    
    Exit Function
   

SearchLinkErr:
    ShowError "cAccessTier.SearchLink", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End Function

