VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBiff12RDCols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements cBiff12RDData
Private mCol As cBiff12RDData


Private Sub Class_Initialize()
    Set mCol = New cBiff12RDData
End Sub

Private Sub Class_Terminate()
    mCol.Clear
    Set mCol = Nothing
End Sub

Private Sub cBiff12RDData_Add(newObj As Object, Optional Key As Variant, Optional Before As Variant, Optional After As Variant)
    ' Don't delete me you stupid compiler!
    ' We don't want to use this method.
End Sub

Private Sub cBiff12RDData_Clear()
    mCol.Clear
End Sub

Private Property Get cBiff12RDData_Count() As Long
    cBiff12RDData_Count = mCol.Count
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Private Property Get cBiff12RDData_Item(Index As Variant) As Object
    Set cBiff12RDData_Item = mCol.Item(Index)
End Property

Private Function cBiff12RDData_NewEnum() As IUnknown
    Set cBiff12RDData_NewEnum = mCol.NewEnum
End Function

Private Sub cBiff12RDData_Remove(Index As Variant)
    mCol.Remove Index
End Sub

Public Sub Remove(Index As Variant)
    mCol.Remove Index
End Sub

' A rather simplied interface, just to be quick!
Public Sub Add(ByVal oCell As cBiff12RDCell)
Attribute Add.VB_UserMemId = -4
    mCol.Add oCell
End Sub

Public Sub Clear()
    mCol.Clear
End Sub

Public Property Get Cols(Index As Variant) As cBiff12RDCell
    Set Cols = mCol.Item(Index)
End Property

Public Function NewEnum() As IUnknown
    Set NewEnum = mCol.NewEnum
End Function


