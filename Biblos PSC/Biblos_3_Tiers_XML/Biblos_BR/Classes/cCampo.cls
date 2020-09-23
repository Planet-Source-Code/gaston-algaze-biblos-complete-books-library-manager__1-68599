VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCampo"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'========================================================================
' Copyright © CHMR 2007 All rights reserved.
'========================================================================
'
'   Module         - cCampo
'   Version Number - 1.0
'   Last Updated   - March 25th 2007 - 4:34 PM
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
Private mTabla As String
Private mCampo As String

'========================================================================
'Propiedades
'========================================================================
Public Property Get Tabla() As String
    Tabla = mTabla
End Property

Public Property Let Tabla(ByVal vNewValue As String)
    mTabla = vNewValue
End Property

Public Property Get Campo() As String
    Campo = mCampo
End Property

Public Property Let Campo(ByVal vNewValue As String)
    mCampo = vNewValue
End Property

'========================================================================
'
'   Title          - Función SearchByTabla()
'   Description    - Realiza busquedas, devolviendo un archivo XML con el
'                    resultado.
'   Version Number - 1.0
'   Last Updated   - March 26th 2007 - 7:15 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SearchByTabla(iUserID As Variant, sTable As Variant, sXML As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo SearchByTablaErr

Dim strXML As String
Dim oSecAgent As New cSecurityAgent

    SearchByTabla = False

    SearchByTabla = oDataAccessTier.SearchCampoByTabla(strXML, CStr(sTable), lErrNum, sErrDesc, sErrSource)
    
    If SearchByTabla Then
        sXML = strXML
        Exit Function
    Else
    
SearchByTablaErr:
    
        ShowError "cCampo.SearchByTabla", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

