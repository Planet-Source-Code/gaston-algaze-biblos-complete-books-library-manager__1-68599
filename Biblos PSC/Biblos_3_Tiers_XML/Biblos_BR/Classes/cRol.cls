VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRol"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'========================================================================
' Copyright © CHMR 2007 All rights reserved.
'========================================================================
'
'   Module         - cRol
'   Version Number - 1.0
'   Last Updated   - February 25th 2007 - 4:34 PM
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
Private mID As Integer
Private mDescripcion As String
Private mPrivilegiosGlobales As Integer
Private mFecha_Alta As String
Private mFecha_Ult_Act As String
Private mFecha_Baja As String

'========================================================================
'Propiedades
'========================================================================
Public Property Get ID() As Integer
    ID = mID
End Property

Public Property Let ID(ByVal vNewValue As Integer)
    mID = vNewValue
End Property

Public Property Get PrivilegiosGlobales() As Integer
    PrivilegiosGlobales = mPrivilegiosGlobales
End Property

Public Property Let PrivilegiosGlobales(ByVal vNewValue As Integer)
    mPrivilegiosGlobales = vNewValue
End Property

Public Property Get Descripcion() As String
    Descripcion = mDescripcion
End Property

Public Property Let Descripcion(ByVal vNewValue As String)
    mDescripcion = vNewValue
End Property

Public Property Get Fecha_Alta() As String
    Fecha_Alta = mFecha_Alta
End Property

Public Property Let Fecha_Alta(ByVal vNewValue As String)
    mFecha_Alta = vNewValue
End Property

Public Property Get Fecha_Ult_Act() As String
    Fecha_Ult_Act = mFecha_Ult_Act
End Property

Public Property Let Fecha_Ult_Act(ByVal vNewValue As String)
    mFecha_Ult_Act = vNewValue
End Property

Public Property Get Fecha_Baja() As String
    Fecha_Baja = mFecha_Baja
End Property

Public Property Let Fecha_Baja(ByVal vNewValue As String)
    mFecha_Baja = vNewValue
End Property

'========================================================================
'Métodos y Funciones
'========================================================================
'========================================================================
'
'   Title          - Función ValidateMe()
'   Description    - Valida que las propiedades del objeto sean correctas
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Private Function ValidateMe(Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    ValidateMe = False
    If Len(CStr(ID)) > 0 Then
        If Len(Descripcion) > 255 Or _
           Len(Descripcion) = 0 Or Not IsNumeric(ID) Or _
           Not IsNumeric(PrivilegiosGlobales) Or _
           PrivilegiosGlobales > 15 Then
            ValidateMe = False
            Err.Raise -1, "cRol.ValidateMe", "Longitud o tipo de parámetros incorrecto."
        Else
            ValidateMe = True
        End If
    End If
End Function

'========================================================================
'
'   Title          - Función Add()
'   Description    - Inserta un nuevo elemento
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Add(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo AddErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oHelper As cRol
Dim oSecAgent As New cSecurityAgent

    Add = False
    
    If Not ValidateMe Then
        Err.Raise -1, "cRol.Add", "Longitud de parámetros incorrecta."
    End If
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cRol", opInsert, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    'Primero me fijo que no exista el nuevo objeto en el sistema
    If Search(iUserID, strXML, "descripcion = '" & Descripcion & "' AND fecha_baja IS NULL", , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cRol
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cRol.Add", "El nuevo objeto ya existe en el sistema.<BR>Por favor verifique los datos."
        Else
            If lErrNum = "-2" Then
                Err.Raise -1, "cRol.Read", "Hay mas de un registro para leer en el archivo XML"
            End If
            
            'Si no existe lo creo...
            oRs.fields.Append "descripcion", adBSTR
            oRs.fields.Append "fecha_alta", adBSTR
            oRs.fields.Append "fecha_ult_act", adBSTR
            oRs.fields.Append "privilegios_globales", adBSTR
            
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = Descripcion
            oRs(1) = Format(Now(), "yyyymmdd")
            oRs(2) = Format(Now(), "yyyymmdd")
            oRs(3) = PrivilegiosGlobales
            
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Add = oDataAccessTier.SaveRol(oDOM.xml, lErrNum, sErrDesc, sErrSource)
            
            Set oHelper = Nothing
            Set oRs = Nothing
            Set oDOM = Nothing
        End If
    End If
    
    If Add Then
        Exit Function
    Else
    
AddErr:
        
        Set oHelper = Nothing
        Set oRs = Nothing
        Set oDOM = Nothing
    
        ShowError "cRol.Add", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Read()
'   Description    - Carga al objeto, leyendo desde un archivo XML
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Read(iUserID As Variant, sXML As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo ReadErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oSecAgent As New cSecurityAgent

    Read = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cRol", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    strXML = CStr(sXML)
    
    Set oRs = RecordsetFromXMLString(strXML)
    
    If Not oRs.EOF Then
        If oRs.RecordCount > 1 Then
            Err.Raise -2, "cRol.Read", "Hay mas de un registro para leer en el archivo XML"
        Else
            ID = oRs("id")
            Descripcion = oRs("descripcion")
            PrivilegiosGlobales = IIf(IsNull(oRs("privilegios_globales")), 0, oRs("privilegios_globales"))
            Fecha_Alta = IIf(IsNull(oRs("fecha_alta")), "", oRs("fecha_alta"))
            Fecha_Ult_Act = IIf(IsNull(oRs("fecha_ult_act")), "", oRs("fecha_ult_act"))
            Fecha_Baja = IIf(IsNull(oRs("fecha_baja")), "", oRs("fecha_baja"))
            Read = True
        End If
    Else
        Err.Raise -1, "cRol.Read", "No se encontro ningún registro para leer en el archivo XML"
    Set oRs = Nothing
    End If
    
    If Read Then
        Exit Function
    Else
    
ReadErr:
        Set oRs = Nothing
        ShowError "cRol.Read", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Search()
'   Description    - Realiza busquedas, devolviendo un archivo XML con el
'                    resultado.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Search(iUserID As Variant, sXML As Variant, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo searchErr

Dim strXML As String
Dim strRestricciones As String
Dim oSecAgent As New cSecurityAgent

    Search = False
    
    sSearch = Replace(sSearch, "id ", "roles.id_Rol ", , , vbTextCompare)
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cRol", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    Else ' me fijo si tiene algun restriccion
        If oSecAgent.GetRestrictions(CInt(iUserID), "cRol", strRestricciones) Then
            If (Len(sSearch) > 0) And (Len(strRestricciones) > 0) Then
                sSearch = strRestricciones & "AND " & sSearch
            Else
                If Len(sSearch) = 0 Then
                    sSearch = strRestricciones
                End If
            End If
        End If
    End If
    
    
    
        If Len(sOrderBy) > 0 Then
            If oSecAgent.ValidateOrderBy(sOrderBy, "cRol") Then
                If Len(sSort) > 0 Then
                    If Not oSecAgent.ValidateSort(sSort) Then
                        Err.Raise -1, "cRol.Search", """SORT"" inválido, solo se permite ASC o DESC."
                    End If
                End If
            Else
                Err.Raise -1, "cRol.Search", """ORDER BY"" inválido, los campos no corresponden."
            End If
        End If

    Search = oDataAccessTier.SearchRol(strXML, sSearch, sOrderBy, sSort, lErrNum, sErrDesc, sErrSource)
    
    If Search Then
        sXML = strXML
        Exit Function
    Else
    
searchErr:
        ShowError "cRol.Search", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
        End If
End Function

'========================================================================
'
'   Title          - Función Delete()
'   Description    - Elimina lógica o físicamente al objeto.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Delete(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo DeleteErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim oSecAgent As New cSecurityAgent

    Delete = False
   
    If Not IsNumeric(ID) Then
        Err.Raise -1, "cRol.Update", "Longitud o tipo de parámetro incorrecto."
    End If
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cRol", opDelete, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    oRs.fields.Append "id", adBSTR
    
    oRs.open
    
    oRs.AddNew

    oRs(0) = ID

    oRs.Update
    
    oRs.save oDOM, adPersistXML
    
    Delete = oDataAccessTier.DeleteRol(oDOM.xml, 1, lErrNum, sErrDesc, sErrSource)
    
    Set oRs = Nothing
    Set oDOM = Nothing
    
    If Delete Then
        Exit Function
    Else
    
DeleteErr:
        Set oRs = Nothing
        Set oDOM = Nothing
    
        ShowError "cRol.Delete", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Update()
'   Description    - Actualiza al objeto.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Update(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo UpdateErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oHelper As cRol
Dim oSecAgent As New cSecurityAgent

    Update = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cRol", opUpdate, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Search(iUserID, strXML, "id <> " & CStr(ID) & " AND descripcion = '" & Descripcion & "' AND fecha_baja IS NULL", , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cRol
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cRol.Update", "El nuevo objeto ya existe en el sistema.<BR>Por favor verifique los datos."
        Else
            If lErrNum = "-2" Then
                Err.Raise -1, "cRol.Update", "Hay mas de un registro para leer en el archivo XML"
            End If
            If Not ValidateMe Then
                Err.Raise -1, "cRol.Update", "Longitud o tipo de parámetro incorrecto."
            End If
            
            oRs.fields.Append "descripcion", adBSTR
            oRs.fields.Append "fecha_ult_act", adBSTR
            oRs.fields.Append "fecha_baja", adBSTR, , adFldIsNullable
            oRs.fields.Append "id", adBSTR
            oRs.fields.Append "privilegios_globales", adBSTR
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = Descripcion
            oRs(1) = Format(Now(), "yyyymmdd")
            oRs(2) = Fecha_Baja
            oRs(3) = ID
            oRs(4) = PrivilegiosGlobales
            
        
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Update = oDataAccessTier.UpdateRol(oDOM.xml, lErrNum, sErrDesc, sErrSource)
                            
            Set oHelper = Nothing
            Set oRs = Nothing
            Set oDOM = Nothing
        End If
    Else
        Err.Raise -1, "cRol.Update", "No se pudo comprobar la integridad referencial."
    End If
    
    If Update Then
        Exit Function
    Else
    
UpdateErr:
        Set oHelper = Nothing
        Set oRs = Nothing
        Set oDOM = Nothing

        ShowError "cRol.Update", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Recover()
'   Description    - Recupera lógicamente al objeto.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Recover(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo RecoverErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oHelper As cRol
Dim oSecAgent As New cSecurityAgent

    Recover = False
        
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cRol", opUpdate, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
        
    Set oHelper = New cRol
        
    If Not IsNumeric(ID) Then
        Err.Raise -1, "cRol.Recover", "Longitud o tipo de parámetro incorrecto."
    End If
    
    If Search(iUserID, strXML, "id = " & CStr(ID), , , lErrNum, sErrDesc, sErrSource) Then
        oHelper.Read strXML, lErrNum, sErrDesc, sErrSource
    End If
    
    If Search(iUserID, strXML, "id <> " & CStr(ID) & " AND descripcion = '" & oHelper.Descripcion & "'", , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cRol
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cRol.Recover", "El nuevo objeto ya existe en el sistema.<BR>Por favor verifique los datos."
        Else
            If lErrNum = "-2" Then
                Err.Raise -1, "cRol.Read", "Hay mas de un registro para leer en el archivo XML"
            End If
            
            oRs.fields.Append "id", adBSTR
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = ID
        
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Recover = oDataAccessTier.RecoverRol(oDOM.xml, lErrNum, sErrDesc, sErrSource)
            
            Set oHelper = Nothing
            Set oRs = Nothing
            Set oDOM = Nothing
        End If
    End If
    
    If Recover Then
        Exit Function
    Else
    
RecoverErr:
        
        Set oHelper = Nothing
        Set oRs = Nothing
        Set oDOM = Nothing
    
        ShowError "cRol.Recover", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función GetRestricciones()
'   Description    - Carga las Restricciones, para ese rol
'   Version Number - 1.0
'   Last Updated   - 27/02/2007
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function GetRestricciones(iUserID As Variant, sXML As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo GetRestriccionesErr

Dim oRestricciones As New cRestriccion
Dim strXML As String
Dim oSecAgent As New cSecurityAgent
    
    GetRestricciones = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cRestriccion", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If oRestricciones.SearchByRol(iUserID, sXML, ID, lErrNum, sErrDesc, sErrSource) Then
    
        GetRestricciones = True
        
    End If
    
    Set oRestricciones = Nothing
    
    If GetRestricciones Then
        Exit Function
    Else
    
GetRestriccionesErr:
        
        ShowError "cRol.GetRestricciones", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource

    End If
    
End Function

'========================================================================
'
'   Title          - Función GetPermisos()
'   Description    - Carga los permisos que tiene el rol
'   Version Number - 1.0
'   Last Updated   - 27/02/2007
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function GetPermisos(sXML As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo GetPermisosErr

Dim oPermisos As New cPermiso
    
    GetPermisos = False
    
    If oPermisos.SearchByRol(sXML, ID, lErrNum, sErrDesc, sErrSource) Then
            
        GetPermisos = True
        
    End If
    
    Set oPermisos = Nothing
    
    If GetPermisos Then
        Exit Function
    Else
    
GetPermisosErr:
        
        ShowError "cRol.GetPermisos", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función SetPermisos()
'   Description    - Setea los vPermisos que tiene el rol
'   Version Number - 1.0
'   Last Updated   - 27/02/2007
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SetPermisos(sXML As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo SetPermisosErr
Dim oRs As New Recordset
    
    SetPermisos = False
    Set oRs = RecordsetFromXMLString(CStr(sXML))
    If Not oRs.EOF Then
        SetPermisos = oDataAccessTier.SavePermisoByRol(CStr(sXML), lErrNum, sErrDesc, sErrSource)
    Else
        SetPermisos = True
    End If
    Set oRs = Nothing
    If SetPermisos Then
        Exit Function
    Else
    
SetPermisosErr:
    
        ShowError "cRol.SetPermisos", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource

    End If
    
End Function



