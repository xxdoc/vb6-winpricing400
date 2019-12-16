VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTotalCommission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public EMP_ID As Long
Public PARENT_ID As Long
Public SALE_CODE As String
Public SALE_NAME  As String

Public YYYYMM As String

Public TOTAL_TARGET As Double
Public TOTAL_SALE As Double
Public TOTAL_SALE_AMOUNT As Double

Public TOTAL_RCP As Double

Public SALE_PERCENT As Double

Public COM_NOW As Double
Public COM_RCP As Double
Public COM_INTEREST As Double
 
Public Function GetKey() As String
   GetKey = Trim(EMP_ID & "-" & YYYYMM)
End Function

