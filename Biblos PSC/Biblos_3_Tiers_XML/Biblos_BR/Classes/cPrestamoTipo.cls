VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrestamoTipo"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'========================================================================
' Copyright © CHMR 2007 All rights reserved.
'========================================================================
'
'   Module         - cPrestamoTipo
'   Version Number - 1.0
'   Last Updated   - March 21th 2007 - 4:34 PM
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
Private mDuracion As String
Private mFecha_Alta As String
Private mFecha_Ult_Act As String

'========================================================================
'Propiedades
'========================================================================
Public Property Get ID() As Integer
    ID = mID
End Property

Public Property Let ID(ByVal vNewValue As Integer)
    mID = vNewValue
End Property

Public Property Get Descripcion() As String
    Descripcion = mDescripcion
End Property

Public Property Let Descripcion(ByVal vNewValue As String)
    mDescripcion = vNewValue
End Property

Public Property Get Duracion() As Integer
    Duracion = mDuracion
End Property

Public Property Let Duracion(ByVal vNewValue As Integer)
    mDuracion = vNewValue
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

'========================================================================
'Métodos y Funciones
'========================================================================
'========================================================================
'
'   Title          - Función ValidateMe()
'   Description    - Valida que las propiedades del objeto sean correctas
'   Version Number - 1.0
'   Last Updated   - February 26th 2007 - 6:49 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Private Function ValidateMe(Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    ValidateMe = False
    If Len(CStr(ID)) > 0 Then
        If Len(Descripcion) > 255 Or _
           Len(Descripcion) = 0 Or _
           Not IsNumeric(ID) Then
            ValidateMe = False
            Err.Raise -1, "cPrestamoTipo.ValidateMe", "Longitud o tipo de parámetros incorrecto."
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
'   Last Updated   - February 26th 2007 - 6:49 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Add(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo AddErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oHelper As cPrestamoTipo
Dim oSecAgent As New cSecurityAgent

    Add = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cPrestamoTipo", opInsert, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Not ValidateMe Then
        Err.Raise -1, "cPrestamoTipo.Add", "Longitud de parámetros incorrecta."
    End If
    
    'Primero me fijo que no exista el nuevo objeto en el sistema
    If Search(iUserID, strXML, "descripcion = '" & Descripcion & "' AND id <> " & ID, , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cPrestamoTipo
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cPrestamoTipo.Add", "El nuevo objeto ya existe en el sistema.<BR>Por favor verifique los datos."
        Else
            If lErrNum = "-2" Then
                Err.Raise -1, "cPrestamoTipo.Read", "Hay mas de un registro para leer en el archivo XML"
            End If
            
            'Si no existe lo creo...
            oRs.fields.Append "descripcion", adBSTR
            oRs.fields.Append "duracion", adBSTR
            oRs.fields.Append "fecha_alta", adBSTR
            oRs.fields.Append "fecha_ult_act", adBSTR
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = Descripcion
            oRs(1) = Duracion
            oRs(2) = Format(Now(), "yyyymmdd")
            oRs(3) = Format(Now(), "yyyymmdd")
            
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Add = oDataAccessTier.SavePrestamoTipo(oDOM.xml, lErrNum, sErrDesc, sErrSource)
            
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
    
        ShowError "cPrestamoTipo.Add", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Read()
'   Description    - Carga al objeto, leyendo desde un archivo XML
'   Version Number - 1.0
'   Last Updated   - February 26th 2007 - 6:49 PM
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
    
    strXML = CStr(sXML)
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cPrestamoTipo", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    Set oRs = RecordsetFromXMLString(strXML)
    
    If Not oRs.EOF Then
        If oRs.RecordCount > 1 Then
            Err.Raise -2, "cPrestamoTipo.Read", "Hay mas de un registro para leer en el archivo XML"
        Else
            ID = oRs("id")
            Descripcion = oRs("descripcion")
            Duracion = oRs("duracion")
            Fecha_Alta = IIf(IsNull(oRs("fecha_alta")), "", oRs("fecha_alta"))
            Fecha_Ult_Act = IIf(IsNull(oRs("fecha_ult_act")), "", oRs("fecha_ult_act"))
            Read = True
        End If
    Else
        Err.Raise -1, "cPrestamoTipo.Read", "No se encontro ningún registro para leer en el archivo XML"
    Set oRs = Nothing
    End If
    
    If Read Then
        Exit Function
    Else
    
ReadErr:
        
        Set oRs = Nothing
    
        ShowError "cPrestamoTipo.Read", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Search()
'   Description    - Realiza busquedas, devolviendo un archivo XML con el
'                    resultado.
'   Version Number - 1.0
'   Last Updated   - February 26th 2007 - 6:49 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Search(iUserID As Variant, sXML As Variant, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo searchErr

Dim strXML As String
Dim strRestricciones As String
Dim oSecAgent As New cSecurityAgent

    Search = False
    
    sSearch = Replace(sSearch, "id ", "prestamo_tipos.id_Prestamos_Tipos ", , , vbTextCompare)
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cPrestamoTipo", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    Else ' me fijo si tiene algun restriccion
        If oSecAgent.GetRestrictions(CInt(iUserID), "cPrestamoTipo", strRestricciones, lErrNum, sErrDesc, sErrSource) Then
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
            If oSecAgent.ValidateOrderBy(sOrderBy, "cPrestamoTipo", lErrNum, sErrDesc, sErrSource) Then
                If Len(sSort) > 0 Then
                    If Not oSecAgent.ValidateSort(sSort) Then
                        Err.Raise -1, "cPrestamoTipo.Search", """SORT"" inválido, solo se permite ASC o DESC."
                    End If
                End If
            Else
                Err.Raise -1, "cPrestamoTipo.Search", """ORDER BY"" inválido, los campos no corresponden."
            End If
        End If

    Search = oDataAccessTier.SearchPrestamoTipo(strXML, sSearch, sOrderBy, sSort, lErrNum, sErrDesc, sErrSource)
    
    If Search Then
        sXML = strXML
        Exit Function
    Else
    
searchErr:
    
        ShowError "cPrestamoTipo.Search", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Delete()
'   Description    - Elimina lógica o físicamente al objeto.
'   Version Number - 1.0
'   Last Updated   - February 26th 2007 - 6:49 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Delete(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo DeleteErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim oPrestamo As New cPrestamo
Dim strXML As String
Dim oSecAgent As New cSecurityAgent

    Delete = False
   
    If Not IsNumeric(ID) Then
        Err.Raise -1, "cPrestamoTipo.Delete", "Longitud o tipo de parámetro incorrecto."
    End If
    
    oRs.fields.Append "id", adBSTR
    
    oRs.open
    
    oRs.AddNew

    oRs(0) = ID

    oRs.Update
    
    oRs.save oDOM, adPersistXML
    
    If oPrestamo.Search(iUserID, strXML, "PrestamoTipoID = " & ID, , , lErrNum, sErrDesc, sErrSource) Then
        If oPrestamo.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cPrestamoTipo.Delete", "La PrestamoTipo no se puede eliminar existen Items que la utilizan."
        Else
            Delete = oDataAccessTier.DeletePrestamoTipo(oDOM.xml, lErrNum, sErrDesc, sErrSource)
        End If
    End If
    
    Set oRs = Nothing
    Set oDOM = Nothing
    
    If Delete Then
        Exit Function
    Else
    
DeleteErr:
        Set oRs = Nothing
        Set oDOM = Nothing
    
        ShowError "cPrestamoTipo.Delete", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Update()
'   Description    - Actualiza al objeto.
'   Version Number - 1.0
'   Last Updated   - February 26th 2007 - 6:49 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Update(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo UpdateErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oHelper As cPrestamoTipo
Dim oItem As New cItem
Dim oSecAgent As New cSecurityAgent

    Update = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cPrestamoTipo", opUpdate, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Search(iUserID, strXML, "id <> " & CStr(ID) & " AND descripcion = '" & Descripcion & "'", , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cPrestamoTipo
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cPrestamoTipo.Update", "El nuevo objeto ya existe en el sistema.<BR>Por favor verifique los datos."
        Else
            If lErrNum = "-2" Then
                Err.Raise -1, "cPrestamoTipo.Update", "Hay mas de un registro para leer en el archivo XML"
            End If
            If Not ValidateMe Then
                Err.Raise -1, "cPrestamoTipo.Update", "Longitud o tipo de parámetro incorrecto."
            End If
            
            oRs.fields.Append "descripcion", adBSTR
            oRs.fields.Append "duracion", adBSTR
            oRs.fields.Append "fecha_ult_act", adBSTR
            oRs.fields.Append "fecha_baja", adBSTR, , adFldIsNullable
            oRs.fields.Append "id", adBSTR
               
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = Descripcion
            oRs(1) = Duracion
            oRs(2) = Format(Now(), "yyyymmdd")
            oRs(3) = Format(Now(), "yyyymmdd")
            oRs(4) = ID
        
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Update = oDataAccessTier.UpdatePrestamoTipo(oDOM.xml, lErrNum, sErrDesc, sErrSource)
                            
            Set oHelper = Nothing
            Set oRs = Nothing
            Set oDOM = Nothing
        End If
    Else
        Err.Raise -1, "cPrestamoTipo.Update", "No se pudo comprobar la integridad referencial."
    End If
    
    If Update Then
        Exit Function
    Else
    
UpdateErr:
        Set oHelper = Nothing
        Set oRs = Nothing
        Set oDOM = Nothing

        ShowError "cPrestamoTipo.Update", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function
