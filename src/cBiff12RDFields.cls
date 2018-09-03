VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBiff12RDFields"
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
Public Sub Add(ByVal sFieldArray As String)
    Dim oField As cBiff12RDField, lCols As Long
    Dim tmp1() As String, tmp2() As String, tmp3() As String
    Dim tmp4() As String, tmp5() As String
    If InStr(sFieldArray, vbTab) > 0 Then
      tmp1 = Split(sFieldArray, vbTab)
    Else
      ReDim tmp1(0) As String
      tmp1(0) = sFieldArray
    End If
    For lCols = 0 To UBound(tmp1)      'Fieldcount = Maximum
      tmp3 = Split(tmp1(lCols), "=")
      tmp4 = Split(tmp3(0), ",")       'FieldNumber, Fieldname
      tmp5 = Split(tmp3(1), ",")
      Set oField = New cBiff12RDField  'DataType, DataTypeNumber
      With oField
        .FieldIndex = CLng(tmp4(0))
        .FieldName = tmp4(1)
        .FieldType = tmp5(0)
        .TypeNumber = CLng(tmp5(1))
      End With
      mCol.Add oField
    Next
End Sub

Public Sub Clear()
    mCol.Clear
End Sub

Public Property Get Fields(Index As Variant) As cBiff12RDField
    Set Fields = mCol.Item(Index)
End Property

Public Function NewEnum() As IUnknown
    Set NewEnum = mCol.NewEnum
End Function
