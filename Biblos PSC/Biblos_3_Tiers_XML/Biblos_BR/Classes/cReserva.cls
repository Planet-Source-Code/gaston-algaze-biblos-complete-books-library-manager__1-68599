VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cReserva"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'========================================================================
' Copyright © CHMR 2007 All rights reserved.
'========================================================================
'
'   Module         - cReserva
'   Version Number - 1.0
'   Last Updated   - March 22 2007 - 5:34 PM
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
Private mFecha_Reserva As String
Private mUsuarioID As String
Private mItemID As String
Private mItemTipoID As String

'========================================================================
'Propiedades
'========================================================================
Public Property Get ID() As Integer
    ID = mID
End Property

Public Property Let ID(ByVal vNewValue As Integer)
    mID = vNewValue
End Property

Public Property Get Fecha_Reserva() As String
    Fecha_Reserva = mFecha_Reserva
End Property

Public Property Let Fecha_Reserva(ByVal vNewValue As String)
    mFecha_Reserva = vNewValue
End Property

Public Property Get UsuarioID() As Integer
    UsuarioID = mUsuarioID
End Property

Public Property Let UsuarioID(ByVal vNewValue As Integer)
    mUsuarioID = vNewValue
End Property

Public Property Get ItemID() As Integer
    ItemID = mItemID
End Property

Public Property Let ItemID(ByVal vNewValue As Integer)
    mItemID = vNewValue
End Property

Public Property Get ItemTipoID() As Integer
    ItemTipoID = mItemTipoID
End Property

Public Property Let ItemTipoID(ByVal vNewValue As Integer)
    mItemTipoID = vNewValue
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
        If Not IsNumeric(ID) And Format(Fecha_Reserva, "yyyymmdd") < Format(Now, "yyyymmdd") Then
            ValidateMe = False
            Err.Raise -1, "cReserva.ValidateMe", "Longitud o tipo de parámetros incorrecto."
        Else
            ValidateMe = True
        End If
    Else
        'If Len(Descripcion) > 255 Or Len(Titulo) > 30 Or _
        '   Len(Titulo) = 0 Then
        '    ValidateMe = False
        '    Err.Raise -1, "cReserva.ValidateMe", "Longitud o tipo de parámetros incorrecto."
        'Else
            ValidateMe = True
        'End If
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
Dim oSecAgent As New cSecurityAgent

    Add = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cReserva", opInsert, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Not ValidateMe Then
        Err.Raise -1, "cReserva.Add", "Longitud de parámetros incorrecta."
    End If

    oRs.fields.Append "fecha_reserva", adBSTR
    oRs.fields.Append "usuarioID", adBSTR
    oRs.fields.Append "ItemID", adBSTR
    oRs.fields.Append "ItemTipoID", adBSTR
    
    
    oRs.open
    
    oRs.AddNew

    oRs(0) = Format(Fecha_Reserva, "yyyymmdd")
    oRs(1) = UsuarioID
    oRs(2) = ItemID
    oRs(3) = ItemTipoID
    

    oRs.Update
    
    oRs.save oDOM, adPersistXML
    
    Add = oDataAccessTier.SaveReserva(oDOM.xml, lErrNum, sErrDesc, sErrSource)
    
    Set oRs = Nothing
    Set oDOM = Nothing
        
    
    If Add Then
        Exit Function
    Else
    
AddErr:
        
        Set oRs = Nothing
        Set oDOM = Nothing
    
        ShowError "cReserva.Add", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
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
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cReserva", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    strXML = CStr(sXML)
    
    Set oRs = RecordsetFromXMLString(strXML)
    
    If Not oRs.EOF Then
        'If oRs.RecordCount > 1 Then
        '    Err.Raise -2, "cReserva.Read", "Hay mas de un registro para leer en el archivo XML"
        'Else
            ID = oRs("id")
            Fecha_Reserva = oRs("Fecha_reserva")
            UsuarioID = oRs("UsuarioID")
            ItemID = oRs("ItemID")
            ItemTipoID = oRs("itemtipoID")
            
            Read = True
        'End If
    Else
        Err.Raise -1, "cReserva.Read", "No se encontro ningún registro para leer en el archivo XML"
    Set oRs = Nothing
    End If
    
    If Read Then
        Exit Function
    Else
    
ReadErr:
        
        Set oRs = Nothing
    
        ShowError "cReserva.Read", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
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
Dim oSecAgent As New cSecurityAgent
Dim strRestricciones As String

    Search = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cReserva", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    Else ' me fijo si tiene algun restriccion
        If oSecAgent.GetRestrictions(CInt(iUserID), "cReserva", strRestricciones) Then
            If (Len(sSearch) > 0) And (Len(strRestricciones) > 0) Then
                sSearch = strRestricciones & "AND " & sSearch
            Else
                If Len(sSearch) = 0 Then
                    sSearch = strRestricciones
                End If
            End If
        End If
    End If
    
    sSearch = Replace(sSearch, "ItemID ", "Reservas.id_item ", , , vbTextCompare)
    sSearch = Replace(sSearch, "fecha_baja ", "Reservas.fecha_baja ", , , vbTextCompare)
    sSearch = Replace(sSearch, "usuarioID ", "Reservas.id_usuario ", , , vbTextCompare)
    sSearch = Replace(sSearch, "itemtipoID ", "items_tipos.id_item_tipo ", , , vbTextCompare)
    sSearch = Replace(sSearch, "id ", "Reservas.id_Reserva ", , , vbTextCompare)
    
    
        If Len(sOrderBy) > 0 Then
            
            If oSecAgent.ValidateOrderBy(sOrderBy, "cReserva") Then
                If Len(sSort) > 0 Then
                    If Not oSecAgent.ValidateSort(sSort) Then
                        Err.Raise -1, "cReserva.Search", """SORT"" inválido, solo se permite ASC o DESC."
                    End If
                End If
            Else
                Err.Raise -1, "cReserva.Search", """ORDER BY"" inválido, los campos no corresponden."
            End If
        End If

    Search = oDataAccessTier.SearchReserva(strXML, sSearch, sOrderBy, sSort, lErrNum, sErrDesc, sErrSource)
    
    If Search Then
        sXML = strXML
        Exit Function
    Else
    
searchErr:
    
        ShowError "cReserva.Search", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
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
Dim oSecAgent As New cSecurityAgent

    Update = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cReserva", opUpdate, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    
            If Not ValidateMe Then
                Err.Raise -1, "cReserva.Update", "Longitud o tipo de parámetro incorrecto."
            End If
            
            oRs.fields.Append "fecha_reserva", adBSTR
            oRs.fields.Append "usuarioID", adBSTR
            oRs.fields.Append "ItemID", adBSTR
            oRs.fields.Append "id", adBSTR
            oRs.fields.Append "ItemTipoID", adBSTR
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = Format(Fecha_Reserva, "yyyymmdd")
            oRs(1) = UsuarioID
            oRs(2) = ItemID
            oRs(3) = ID
            oRs(4) = ItemTipoID
        
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Update = oDataAccessTier.UpdateReserva(oDOM.xml, lErrNum, sErrDesc, sErrSource)
                            
            Set oRs = Nothing
            Set oDOM = Nothing
        
    
    If Update Then
        Exit Function
    Else
    
UpdateErr:
        Set oRs = Nothing
        Set oDOM = Nothing

        ShowError "cReserva.Update", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Delete()
'   Description    - Elimina al objeto.
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
Dim oItem As New cItem
Dim strXML As String

    Delete = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cReserva", opDelete) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
   
    If Not IsNumeric(ID) Then
        Err.Raise -1, "cReserva.Delete", "Longitud o tipo de parámetro incorrecto."
    End If
    
    oRs.fields.Append "id", adBSTR
    
    oRs.open
    
    oRs.AddNew

    oRs(0) = ID

    oRs.Update
    
    oRs.save oDOM, adPersistXML
                    
    Delete = oDataAccessTier.DeleteReserva(oDOM.xml, 0, lErrNum, sErrDesc, sErrSource)

    Set oRs = Nothing
    Set oDOM = Nothing
    
    If Delete Then
        Exit Function
    Else
    
DeleteErr:
        Set oRs = Nothing
        Set oDOM = Nothing
    
        ShowError "cReserva.Delete", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function
