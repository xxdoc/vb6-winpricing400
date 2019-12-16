VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProductionWH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ModuleName As String
Private m_ADOConn As ADODB.Connection
Private m_Rs1 As ADODB.Recordset
Private m_Rs2 As ADODB.Recordset
Private m_Rs3 As ADODB.Recordset

Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   m_ModuleName = "clsProduction"
   Set m_Rs1 = New ADODB.Recordset
   Set m_Rs2 = New ADODB.Recordset
   Set m_Rs3 = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   If m_Rs1.State = adStateOpen Then
      m_Rs1.Close
   End If
   Set m_Rs1 = Nothing

   If m_Rs2.State = adStateOpen Then
      m_Rs2.Close
   End If
   Set m_Rs2 = Nothing

   If m_Rs3.State = adStateOpen Then
      m_Rs3.Close
   End If
   Set m_Rs3 = Nothing
End Sub
' change 1/07/2005
'==
Public Function QueryJob(Ua As CJobWareHouse, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryJob"
   QueryJob = False
   
   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   If Not Rs.EOF Then
      Ua.JOB_ID = Rs("JOB_ID")
   End If
   If Ua.QueryFlag = 1 Then
    'INPUT ++++++++++++++++++++++++++
      Dim Inp As CJobInputWarehouse
      Set Inp = New CJobInputWarehouse
      Inp.JOB_INOUT_ID = -1
      Inp.JOB_ID = Ua.JOB_ID
      Call Inp.QueryData(1, m_Rs2, iCount)
      Set Inp = Nothing
      
      Set Ua.Inputs = Nothing
      Set Ua.Inputs = New Collection
      While Not m_Rs2.EOF
       Set Inp = New CJobInputWarehouse
         Call Inp.PopulateFromRS(1, m_Rs2)
         Inp.Flag = "I"
         If Inp.TX_TYPE = "E" Then
            Call Ua.Inputs.add(Inp)
         End If
         Set Inp = Nothing
         m_Rs2.MoveNext
      Wend
'INPUT ++++++++++++++++++++++++++

'OUTPUT ++++++++++++++++++++++++++
      Dim Op As CJobInputWarehouse
      Set Op = New CJobInputWarehouse
      Op.JOB_INOUT_ID = -1
      Op.JOB_ID = Ua.JOB_ID
      Call Op.QueryData(1, m_Rs2, iCount)
      Set Op = Nothing
      
      Set Ua.Outputs = Nothing
      Set Ua.Outputs = New Collection
      While Not m_Rs2.EOF
       Set Op = New CJobInputWarehouse
         Call Op.PopulateFromRS(1, m_Rs2)
         Op.Flag = "I"
         If Op.TX_TYPE = "I" Then
            Call Ua.Outputs.add(Op)
         End If
         Set Op = Nothing
         m_Rs2.MoveNext
      Wend
'OUTPUT ++++++++++++++++++++++++++

''MACHINE USED TIME ++++++++++++++++++++++++++
'      Dim EH As CJobResource
'     Set EH = New CJobResource
'      EH.JOB_ID = Ua.JOB_ID
'      Call EH.QueryData(m_Rs1, iCount)
'      Set EH = Nothing
'
'      Set Ua.Machines = Nothing
'      Set Ua.Machines = New Collection
'      While Not m_Rs1.EOF
'         Set EH = New CJobResource
'         Call EH.PopulateFromRS(1, m_Rs1)
'
'         EH.Flag = "I"
'         If EH.MACHINE_NO <> "" Then
'         Call Ua.Machines.add(EH)
'         End If
'         Set EH = Nothing
'         m_Rs1.MoveNext
'      Wend
'      'MACHINE TIME USED ++++++++++++++++++++++++++
'
''PERSON USED TIME ++++++++++++++++++++++++++
'      Dim Ep As CJobResource
'     Set Ep = New CJobResource
'      Ep.JOB_ID = Ua.JOB_ID
'      Call Ep.QueryData(m_Rs1, iCount)
'      Set Ep = Nothing
'
'      Set Ua.Peoples = Nothing
'      Set Ua.Peoples = New Collection
'      While Not m_Rs1.EOF
'         Set Ep = New CJobResource
'         Call Ep.PopulateFromRS(2, m_Rs1)
'
'         Ep.Flag = "I"
'         If Ep.EMP_ID > 0 Then
'         Call Ua.Peoples.add(Ep)
'         End If
'         Set Ep = Nothing
'         m_Rs1.MoveNext
'      Wend
'      'PERSON TIME USED ++++++++++++++++++++++++++
'
'
'     'PARAMETER TIME ++++++++++++++++++++++++++
'      Dim PP As CJobParameter
'     Set PP = New CJobParameter
'      PP.JOB_ID = Ua.JOB_ID
'      Call PP.QueryData(1, m_Rs1, iCount)
'      Set PP = Nothing
'
'      Set Ua.Parameters = Nothing
'      Set Ua.Parameters = New Collection
'      While Not m_Rs1.EOF
'         Set PP = New CJobParameter
'         Call PP.PopulateFromRS(1, m_Rs1)
'
'         PP.Flag = "I"
'         Call Ua.Parameters.add(PP)
'         Set PP = Nothing
'         m_Rs1.MoveNext
'      Wend
'      'PARAMETER USED ++++++++++++++++++++++++++
            
      Dim Jv As CJobVerify
     Set Jv = New CJobVerify
      Jv.JOB_ID = Ua.JOB_ID
      Call Jv.QueryData(m_Rs1, iCount)
      Set Jv = Nothing
            
      'Job verify
      Set Ua.Verifies = Nothing
      Set Ua.Verifies = New Collection
      While Not m_Rs1.EOF
         Set Jv = New CJobVerify
         Call Jv.PopulateFromRS(1, m_Rs1)
      
         Jv.Flag = "I"
         Call Ua.Verifies.add(Jv)
         Set Jv = Nothing
         m_Rs1.MoveNext
      Wend
      'Job verify
   End If
   
   QueryJob = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryJob = False
End Function
Public Function QueryJobWareHouse(Ua As CJobWareHouse, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryJobWareHouse"
   QueryJobWareHouse = False
   
   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   If Not Rs.EOF Then
      Ua.JOB_ID = Rs("JOB_ID")
   End If
   If Ua.QueryFlag = 1 Then
    'INPUT ++++++++++++++++++++++++++
      Dim Inp As CJobInputWarehouse
      Set Inp = New CJobInputWarehouse
      Inp.JOB_INOUT_ID = -1
      Inp.JOB_ID = Ua.JOB_ID
      Call Inp.QueryData(1, m_Rs2, iCount)
      Set Inp = Nothing
      
      Set Ua.Inputs = Nothing
      Set Ua.Inputs = New Collection
      While Not m_Rs2.EOF
       Set Inp = New CJobInput
         Call Inp.PopulateFromRS(1, m_Rs2)
         Inp.Flag = "I"
         If Inp.TX_TYPE = "E" Then
            Call Ua.Inputs.add(Inp)
         End If
         Set Inp = Nothing
         m_Rs2.MoveNext
      Wend
'INPUT ++++++++++++++++++++++++++

'OUTPUT ++++++++++++++++++++++++++
      Dim Op As CJobInputWarehouse
      Set Op = New CJobInputWarehouse
      Op.JOB_INOUT_ID = -1
      Op.JOB_ID = Ua.JOB_ID
      Call Op.QueryData(1, m_Rs2, iCount)
      Set Op = Nothing
      
      Set Ua.Outputs = Nothing
      Set Ua.Outputs = New Collection
      While Not m_Rs2.EOF
       Set Op = New CJobInput
         Call Op.PopulateFromRS(1, m_Rs2)
         Op.Flag = "I"
         If Op.TX_TYPE = "I" Then
            Call Ua.Outputs.add(Op)
         End If
         Set Op = Nothing
         m_Rs2.MoveNext
      Wend
'OUTPUT ++++++++++++++++++++++++++

''MACHINE USED TIME ++++++++++++++++++++++++++
'      Dim EH As CJobResource
'     Set EH = New CJobResource
'      EH.JOB_ID = Ua.JOB_ID
'      Call EH.QueryData(m_Rs1, iCount)
'      Set EH = Nothing
'
'      Set Ua.Machines = Nothing
'      Set Ua.Machines = New Collection
'      While Not m_Rs1.EOF
'         Set EH = New CJobResource
'         Call EH.PopulateFromRS(1, m_Rs1)
'
'         EH.Flag = "I"
'         If EH.MACHINE_NO <> "" Then
'         Call Ua.Machines.add(EH)
'         End If
'         Set EH = Nothing
'         m_Rs1.MoveNext
'      Wend
'      'MACHINE TIME USED ++++++++++++++++++++++++++
'
''PERSON USED TIME ++++++++++++++++++++++++++
'      Dim Ep As CJobResource
'     Set Ep = New CJobResource
'      Ep.JOB_ID = Ua.JOB_ID
'      Call Ep.QueryData(m_Rs1, iCount)
'      Set Ep = Nothing
'
'      Set Ua.Peoples = Nothing
'      Set Ua.Peoples = New Collection
'      While Not m_Rs1.EOF
'         Set Ep = New CJobResource
'         Call Ep.PopulateFromRS(2, m_Rs1)
'
'         Ep.Flag = "I"
'         If Ep.EMP_ID > 0 Then
'         Call Ua.Peoples.add(Ep)
'         End If
'         Set Ep = Nothing
'         m_Rs1.MoveNext
'      Wend
'      'PERSON TIME USED ++++++++++++++++++++++++++
'
'
'     'PARAMETER TIME ++++++++++++++++++++++++++
'      Dim PP As CJobParameter
'     Set PP = New CJobParameter
'      PP.JOB_ID = Ua.JOB_ID
'      Call PP.QueryData(1, m_Rs1, iCount)
'      Set PP = Nothing
'
'      Set Ua.Parameters = Nothing
'      Set Ua.Parameters = New Collection
'      While Not m_Rs1.EOF
'         Set PP = New CJobParameter
'         Call PP.PopulateFromRS(1, m_Rs1)
'
'         PP.Flag = "I"
'         Call Ua.Parameters.add(PP)
'         Set PP = Nothing
'         m_Rs1.MoveNext
'      Wend
'      'PARAMETER USED ++++++++++++++++++++++++++
'
'      Dim Jv As CJobVerify
'     Set Jv = New CJobVerify
'      Jv.JOB_ID = Ua.JOB_ID
'      Call Jv.QueryData(m_Rs1, iCount)
'      Set Jv = Nothing
'
'      'Job verify
'      Set Ua.Verifies = Nothing
'      Set Ua.Verifies = New Collection
'      While Not m_Rs1.EOF
'         Set Jv = New CJobVerify
'         Call Jv.PopulateFromRS(1, m_Rs1)
'
'         Jv.Flag = "I"
'         Call Ua.Verifies.add(Jv)
'         Set Jv = Nothing
'         m_Rs1.MoveNext
'      Wend
'      'Job verify
   End If
   
   QueryJobWareHouse = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryJobWareHouse = False
End Function
' 3/01/2005
Public Function AddEditJob(Ug As CJobWareHouse, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditJob"
   AddEditJob = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
   
'   Dim Jv As CJobVerify
'   For Each Jv In Ug.Verifies
'      If Jv.Flag = "A" Then
'         Jv.JOB_ID = Ug.JOB_ID
'         Jv.AddEditMode = SHOW_ADD
'         Jv.AddEditData
'      ElseIf Jv.Flag = "E" Then
'         Jv.AddEditMode = SHOW_EDIT
'         Jv.AddEditData
'      ElseIf Jv.Flag = "D" Then
'         Call Jv.DeleteData
'      End If
'   Next Jv
   
   Dim EH As CJobInputWarehouse
   For Each EH In Ug.Inputs
      If EH.Flag = "A" Then
         EH.JOB_ID = Ug.JOB_ID
         EH.AddEditMode = SHOW_ADD
         EH.AddEditData
      ElseIf EH.Flag = "E" Then
         EH.AddEditMode = SHOW_EDIT
         EH.AddEditData
      ElseIf EH.Flag = "D" Then
         Call EH.DeleteData
      End If
   Next EH
  
   Dim EO As CJobInputWarehouse
   For Each EO In Ug.Outputs
      If EO.Flag = "A" Then
         EO.JOB_ID = Ug.JOB_ID
         EO.AddEditMode = SHOW_ADD
         EO.AddEditData
      ElseIf EO.Flag = "E" Then
         EO.AddEditMode = SHOW_EDIT
         EO.AddEditData
      ElseIf EO.Flag = "D" Then
         Call EO.DeleteData
      End If
   Next EO
  
  
'   Dim Ei As CJobResource
'   For Each Ei In Ug.Peoples
'      If Ei.Flag = "A" Then
'         Ei.JOB_ID = Ug.JOB_ID
'         Ei.AddEditMode = SHOW_ADD
'         Ei.AddEditData
'      ElseIf Ei.Flag = "E" Then
'         Ei.AddEditMode = SHOW_EDIT
'         Ei.AddEditData
'      ElseIf Ei.Flag = "D" Then
'         Call Ei.DeleteData
'      End If
'   Next Ei
  
  
'   Dim EA As CJobResource
'   For Each EA In Ug.Machines
'      If EA.Flag = "A" Then
'         EA.JOB_ID = Ug.JOB_ID
'         EA.AddEditMode = SHOW_ADD
'         EA.AddEditData
'      ElseIf EA.Flag = "E" Then
'         EA.AddEditMode = SHOW_EDIT
'         EA.AddEditData
'      ElseIf EA.Flag = "D" Then
'         Call EA.DeleteData
'      End If
'   Next EA
   
'   Dim Ep As CJobParameter
'   For Each Ep In Ug.Parameters
'      If Ep.Flag = "A" Then
'         Ep.JOB_ID = Ug.JOB_ID
'         Ep.AddEditMode = SHOW_ADD
'         Ep.AddEditData
'      ElseIf Ep.Flag = "E" Then
'         Ep.AddEditMode = SHOW_EDIT
'         Ep.AddEditData
'      ElseIf Ep.Flag = "D" Then
'         Call Ep.DeleteData
'      End If
'   Next Ep
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditJob = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditJob = False
End Function

Public Function DeleteJob(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CJobWareHouse

   RName = "DeleteJob"
   DeleteJob = False
   HasBegin = False

   Set B = New CJobWareHouse

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.JOB_ID = UID
   Call B.QueryData(1, m_Rs1, ItemCount)
   
   If Not m_Rs1.EOF Then
      Call B.PopulateFromRS(1, m_Rs1)
   End If
   
   B.JOB_ID = UID
   Call B.DeleteData

   If B.INVENTORY_DOC_ID > 0 Then
      Call glbDaily.DeleteInventoryDoc(B.INVENTORY_DOC_ID, IsOK, False, ErrorObj)
   End If

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteJob = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteJob = True
End Function

Public Function QueryFormula(Ua As CFormula, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryFormula"
   QueryFormula = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
    'INPUT ++++++++++++++++++++++++++
      Dim Inp As CFormulaItem
      Set Inp = New CFormulaItem
      Inp.FORMULA_ITEM_ID = -1
      Inp.FORMULA_ID = Ua.FORMULA_ID
      Call Inp.QueryData(1, m_Rs2, iCount)
      Set Inp = Nothing
 
      Set Ua.Inputs = Nothing
      Set Ua.Inputs = New Collection
      While Not m_Rs2.EOF
       Set Inp = New CFormulaItem
         Call Inp.PopulateFromRS(1, m_Rs2)
         Inp.Flag = "I"
         Call Ua.Inputs.add(Inp)
          Set Inp = Nothing
         m_Rs2.MoveNext
      Wend
   'INPUT ++++++++++++++++++++++++++
            
    'VARIABLE ++++++++++++++++++++++++++
      Dim Fv As CFormulaVariable
      Set Fv = New CFormulaVariable
      Fv.FORMULA_VARIABLE_ID = -1
      Fv.FORMULA_ID = Ua.FORMULA_ID
      Call Fv.QueryData(1, m_Rs2, iCount)
      Set Fv = Nothing
 
      Set Ua.FormulaVariables = Nothing
      Set Ua.FormulaVariables = New Collection
      While Not m_Rs2.EOF
          Set Fv = New CFormulaVariable
         Call Fv.PopulateFromRS(1, m_Rs2)
         Fv.Flag = "I"
         Call Ua.FormulaVariables.add(Fv, Trim(str(Fv.VARIABLE_ID)))
         Set Fv = Nothing
         m_Rs2.MoveNext
      Wend
      'VARIABLE ++++++++++++++++++++++++++
   End If
   
   QueryFormula = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryFormula = False
End Function

Public Sub QueryCostItem(Ua As CCostPrdItem)
Dim TempRs As ADODB.Recordset
Dim Ci As CCostItem
Dim iCount As Long
Dim CR As CCostItemRaw

   Set TempRs = New ADODB.Recordset
   
   Set Ci = New CCostItem
   Ci.COST_ITEM_ID = -1
   Ci.COSTPRD_ITEM_ID = Ua.COSTPRD_ITEM_ID
   Call Ci.QueryData(1, TempRs, iCount)
   Set Ci = Nothing
   
   Set Ua.CostItems = Nothing
   Set Ua.CostItems = New Collection
   While Not TempRs.EOF
      Set Ci = New CCostItem
      Call Ci.PopulateFromRS(1, TempRs)
      Ci.Flag = "I"
      
      Set CR = New CCostItemRaw
      Call CR.SetFieldValue("CSTITEM_RAW_ID", -1)
      Call CR.SetFieldValue("COST_ITEM_ID", Ci.COST_ITEM_ID)
      Call CR.QueryData(1, m_Rs3, iCount)
      Set CR = Nothing
      
      Set Ci.CostItemRaws = Nothing
      Set Ci.CostItemRaws = New Collection
      While Not m_Rs3.EOF
         Set CR = New CCostItemRaw
         Call CR.PopulateFromRS(1, m_Rs3)
         CR.Flag = "I"
         Call Ci.CostItemRaws.add(CR, Trim(Ci.COST_ITEM_ID & "-" & CR.GetFieldValue("PART_ITEM_ID")))
         Set CR = Nothing
         
         m_Rs3.MoveNext
      Wend
         
      Call Ua.CostItems.add(Ci)
      Set Ci = Nothing
      
      TempRs.MoveNext
   Wend
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Public Function QueryCostProduction(Ua As CCostProduction, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryCostProduction"
   QueryCostProduction = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
      Dim Inp As CCostPrdItem
      Set Inp = New CCostPrdItem
      Inp.COSTPRD_ITEM_ID = -1
      Inp.COST_PRODUCTION_ID = Ua.COST_PRODUCTION_ID
      Call Inp.QueryData(1, m_Rs2, iCount)
      Set Inp = Nothing
 
      Set Ua.CostItems = Nothing
      Set Ua.CostItems = New Collection
      While Not m_Rs2.EOF
       Set Inp = New CCostPrdItem
         Call Inp.PopulateFromRS(1, m_Rs2)
         Inp.Flag = "I"
         Call Ua.CostItems.add(Inp)
         Call QueryCostItem(Inp)
          Set Inp = Nothing
         m_Rs2.MoveNext
      Wend
      
      Dim Ce As CCostExpense
      Dim CR As CCostRaw
      Set Ce = New CCostExpense
      Ce.COST_EXPENSE_ID = -1
      Ce.COST_PRODUCTION_ID = Ua.COST_PRODUCTION_ID
      Call Ce.QueryData(1, m_Rs2, iCount)
      Set Ce = Nothing
 
      Set Ua.ExpenseItem = Nothing
      Set Ua.ExpenseItem = New Collection
      While Not m_Rs2.EOF
         Set Ce = New CCostExpense
         Call Ce.PopulateFromRS(1, m_Rs2)
         Ce.Flag = "I"
         
         Set CR = New CCostRaw
         Call CR.SetFieldValue("COST_RAW_ID", -1)
         Call CR.SetFieldValue("COST_EXPENSE_ID", Ce.COST_EXPENSE_ID)
         Call CR.QueryData(1, m_Rs3, iCount)
         Set CR = Nothing
         
         Set Ce.CostRaws = Nothing
         Set Ce.CostRaws = New Collection
         While Not m_Rs3.EOF
            Set CR = New CCostRaw
            Call CR.PopulateFromRS(1, m_Rs3)
            CR.Flag = "I"
            Call Ce.CostRaws.add(CR)
            Set CR = Nothing
            
            m_Rs3.MoveNext
         Wend
         Call Ua.ExpenseItem.add(Ce)
          Set Ce = Nothing
          
         m_Rs2.MoveNext
      Wend
   End If
   
   QueryCostProduction = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCostProduction = False
End Function


Public Function AddEditFormula(Ug As CFormula, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditFormula"
   AddEditFormula = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
   
   Dim EH As CFormulaItem
   For Each EH In Ug.Inputs
      If EH.Flag = "A" Then
         EH.FORMULA_ID = Ug.FORMULA_ID
         EH.AddEditMode = SHOW_ADD
         EH.AddEditData
      ElseIf EH.Flag = "E" Then
         EH.AddEditMode = SHOW_EDIT
         EH.AddEditData
      ElseIf EH.Flag = "D" Then
         Call EH.DeleteData
      End If
   Next EH
  
   Dim Fv As CFormulaVariable
   For Each Fv In Ug.FormulaVariables
      If Fv.Flag = "A" Then
         Fv.FORMULA_ID = Ug.FORMULA_ID
         Fv.AddEditMode = SHOW_ADD
         Fv.AddEditData
      ElseIf Fv.Flag = "E" Then
         Fv.AddEditMode = SHOW_EDIT
         Fv.AddEditData
      ElseIf Fv.Flag = "D" Then
         Call Fv.DeleteData
      End If
   Next Fv
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditFormula = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditFormula = False
End Function

Private Sub AddEditCostItem(Ug As CCostPrdItem)
Dim Ci As CCostItem

   For Each Ci In Ug.CostItems
      If Ci.Flag = "A" Then
         Ci.COSTPRD_ITEM_ID = Ug.COSTPRD_ITEM_ID
         Ci.AddEditMode = SHOW_ADD
         Ci.AddEditData
         
         Call AddEditCostItemRaw(Ci)
      ElseIf Ci.Flag = "E" Then
         Ci.AddEditMode = SHOW_EDIT
         Ci.AddEditData
         
         Call AddEditCostItemRaw(Ci)
      ElseIf Ci.Flag = "D" Then
         Call Ci.DeleteData
      End If
   Next Ci
End Sub

Private Sub AddEditCostItemRaw(Ug As CCostItem)
Dim Ci As CCostItemRaw

   For Each Ci In Ug.CostItemRaws
      If Ci.Flag = "A" Then
         Call Ci.SetFieldValue("COST_ITEM_ID", Ug.COST_ITEM_ID)
         Ci.ShowMode = SHOW_ADD
         Ci.AddEditData
      ElseIf Ci.Flag = "E" Then
         Ci.ShowMode = SHOW_EDIT
         Ci.AddEditData
      ElseIf Ci.Flag = "D" Then
         Call Ci.DeleteData
      End If
   Next Ci
End Sub

Public Function AddEditCostProduction(Ug As CCostProduction, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCostProduction"
   AddEditCostProduction = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
   
   Dim EH As CCostPrdItem
   For Each EH In Ug.CostItems
      If EH.Flag = "A" Then
         EH.COST_PRODUCTION_ID = Ug.COST_PRODUCTION_ID
         EH.AddEditMode = SHOW_ADD
         EH.AddEditData
         
         Call AddEditCostItem(EH)
      ElseIf EH.Flag = "E" Then
         EH.AddEditMode = SHOW_EDIT
         EH.COST_PRODUCTION_ID = Ug.COST_PRODUCTION_ID
         EH.AddEditData
         
         Call AddEditCostItem(EH)
      ElseIf EH.Flag = "D" Then
         Call EH.DeleteData
      End If
   Next EH
   
   Dim Ce As CCostExpense
   Dim CR As CCostRaw
   For Each Ce In Ug.ExpenseItem
      If Ce.Flag = "A" Then
         Ce.COST_PRODUCTION_ID = Ug.COST_PRODUCTION_ID
         Ce.AddEditMode = SHOW_ADD
         Ce.AddEditData
         For Each CR In Ce.CostRaws
            If CR.Flag = "A" Then
               Call CR.SetFieldValue("COST_EXPENSE_ID", Ce.COST_EXPENSE_ID)
               CR.ShowMode = SHOW_ADD
               CR.AddEditData
            ElseIf CR.Flag = "E" Then
               CR.ShowMode = SHOW_EDIT
               CR.AddEditData
            ElseIf CR.Flag = "D" Then
               CR.DeleteData
            End If
         Next CR
      ElseIf Ce.Flag = "E" Then
         Ce.AddEditMode = SHOW_EDIT
         Ce.COST_PRODUCTION_ID = Ug.COST_PRODUCTION_ID
         Ce.AddEditData
         
         For Each CR In Ce.CostRaws
            If CR.Flag = "A" Then
               Call CR.SetFieldValue("COST_EXPENSE_ID", Ce.COST_EXPENSE_ID)
               CR.ShowMode = SHOW_ADD
               CR.AddEditData
            ElseIf CR.Flag = "E" Then
               CR.ShowMode = SHOW_EDIT
               CR.AddEditData
            ElseIf CR.Flag = "D" Then
               CR.DeleteData
            End If
         Next CR
      ElseIf Ce.Flag = "D" Then
         Call Ce.DeleteData
      End If
   Next Ce
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCostProduction = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCostProduction = False
End Function


Public Function DeleteFormula(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CFormula

   RName = "DeleteFormula"
   DeleteFormula = False
   HasBegin = False

   Set B = New CFormula

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.FORMULA_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteFormula = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteFormula = True
End Function

Public Function DeleteCostProduction(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CCostProduction

   RName = "DeleteCostProduction"
   DeleteCostProduction = False
   HasBegin = False

   Set B = New CCostProduction

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.COST_PRODUCTION_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCostProduction = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteCostProduction = True
End Function

Public Function CalPMC1(LossPercent As Double, Can As Double, OverHead As Double, Sum As Double, RMC As Double, PMC As Double) As Long
Dim Devider As Double
Dim SumMarkup As Double

   Devider = 100 - LossPercent
   If Devider <= 0 Then
      Devider = 1
   End If
   RMC = Sum / Devider
   SumMarkup = Can + OverHead
   
   PMC = RMC + SumMarkup
End Function

Public Function CalPMC2(LossPercent As Double, Can As Double, OverHead As Double, Sum As Double, RMC As Double, PMC As Double) As Long
Dim SumMarkup As Double

   RMC = Sum / 100
   SumMarkup = Can + OverHead
   PMC = RMC * (1 + (LossPercent / 100)) + SumMarkup
End Function

Public Function GetWeightedAvgPrice(PID As Long, LID As Long, UseAmount As Double, PriceType As Long) As Double
Dim TempRs As ADODB.Recordset
Dim PL As CPartLocation
Dim Li As CLotItem
Dim iCount As Long
Dim Price1 As Double
Dim NeedAmount As Double

   Set TempRs = New ADODB.Recordset

   If PriceType = 1 Then
      'ค่าเฉลี่ยถ่วงน้ำหนัก
      Set PL = New CPartLocation
      PL.PART_LOCATION_ID = -1
      PL.PART_ITEM_ID = PID
      PL.LOCATION_ID = LID
      Call PL.QueryData(1, TempRs, iCount)
      If Not TempRs.EOF Then
         Call PL.PopulateFromRS(TempRs)
      End If
      Price1 = PL.AVG_PRICE
      Set PL = Nothing
   ElseIf PriceType = 2 Then
      'ค่าเฉลี่ยแบบ FIFO
      Set Li = New CLotItem
      Li.LOT_ITEM_ID = -1
      Li.PART_ITEM_ID = PID
      Li.LOCATION_ID = LID
      Li.OrderBy = -1
      Li.OrderType = 1
      Call Li.QueryData(2, TempRs, iCount)
      NeedAmount = UseAmount
      Price1 = 0
      While (Not TempRs.EOF) And (NeedAmount > 0)
         Call Li.PopulateFromRS(2, TempRs)
         
         If NeedAmount > Li.LEFT_AMOUNT Then
            Price1 = Price1 + (NeedAmount * Li.INCLUDE_UNIT_PRICE)
            NeedAmount = NeedAmount - Li.LEFT_AMOUNT
         Else
            Price1 = Price1 + (Li.LEFT_AMOUNT * Li.INCLUDE_UNIT_PRICE)
            NeedAmount = 0
         End If
         
         TempRs.MoveNext
      Wend
      
      If NeedAmount > 0 Then
         Price1 = Price1 + (NeedAmount * Li.INCLUDE_UNIT_PRICE)
      End If
      Price1 = MyDiff(Price1, UseAmount)
      Set Li = Nothing
   End If
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   
   GetWeightedAvgPrice = Price1
End Function

Public Function GetCalculatedPrice(FormulaID As Long, AvgPMC As Double, PriceType As Long, UseAmount As Double, IsOK As Boolean, ErrorObj As clsErrorLog) As Long
Dim F As CFormula
Dim F2 As CFormula
Dim Fi As CFormulaItem
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim SumPrice1 As Double
Dim Price1 As Double
Dim SumPrice2 As Double
Dim Price2 As Double
Dim IntermediatFlag As Boolean
Dim Devider As Double
Dim Rmc1 As Double
Dim Rmc2 As Double
Dim PL As CPartLocation
Dim Fv As CFormulaVariable
Dim Loss As Double
Dim Can As Double
Dim OverHead As Double

   Set TempRs = New ADODB.Recordset
   Set F = New CFormula
   
   F.FORMULA_ID = FormulaID
   F.QueryFlag = 1
   Call QueryFormula(F, TempRs, iCount, IsOK, ErrorObj)
   If Not TempRs.EOF Then
      Call F.PopulateFromRS(1, TempRs)
   End If
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   
   IntermediatFlag = (F.INTERMEDIAT_FLAG = "Y")
'IntermediatFlag = False
   SumPrice1 = 0
   SumPrice2 = 0
   
   For Each Fi In F.Inputs
      If Fi.FROM_FORMULA > 0 Then
'         If Not CheckRepeatRecursive(Fi.PART_FORMULA_ID, Col, F2) Then
            Call GetCalculatedPrice(Fi.FROM_FORMULA, Price1, PriceType, UseAmount, IsOK, ErrorObj)
'         Else
'            Price1 = 0
'            Price2 = 0
'         End If
      Else
         Price1 = GetWeightedAvgPrice(Fi.PART_ITEM_ID, Fi.LOCATION_ID, UseAmount, PriceType)
      End If
      
      SumPrice1 = SumPrice1 + Fi.ITEM_PERCENT * Price1
      
      Fi.AVG_PRICE = Price1
      Fi.Flag = "E"
   Next Fi
   
   Set Fv = F.FormulaVariables("1")
   Loss = Fv.VARIABLE_VALUE
   Set Fv = F.FormulaVariables("2")
   OverHead = Fv.VARIABLE_VALUE
   Set Fv = F.FormulaVariables("3")
   Can = Fv.VARIABLE_VALUE
      
   If IntermediatFlag Then
      Call CalPMC1(Loss, Can, OverHead, SumPrice1, Rmc1, AvgPMC)
   Else
      Call CalPMC2(Loss, Can, OverHead, SumPrice1, Rmc1, AvgPMC)
   End If
   
   F.AddEditMode = SHOW_EDIT
   F.PMC = AvgPMC
   F.RMC = Rmc1
'   F.Feature.PartItem.AddEditMode = SHOW_EDIT
'   F.Feature.PartItem.CURRENT_PRICE = F.UNIT_PRICE
   
'   Call Col.Add(F)
   
   Set F = Nothing
End Function

Public Function CopyFormula(Ug As CFormula, IsOK As Boolean, AutoCommit As Boolean, Area As Long, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim NewPrice As Double
Dim O As Object
Dim TempRs As ADODB.Recordset

   RName = "CopyFormula"
   CopyFormula = False
   HasBegin = False

   Set TempRs = New ADODB.Recordset
   
   Ug.QueryFlag = 1
   Call QueryFormula(Ug, TempRs, iCount, IsOK, ErrorObj)
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.AddEditMode = SHOW_ADD
   Ug.FORMULA_ID = -1
   Ug.FORMULA_NO = "**" & Ug.FORMULA_NO
   
   Dim CstName As CFormulaItem
   For Each CstName In Ug.Inputs
      CstName.Flag = "A"
      CstName.AddEditMode = SHOW_ADD
   Next CstName

   Dim RT As CFormulaVariable
   For Each RT In Ug.FormulaVariables
      RT.Flag = "A"
      RT.AddEditMode = SHOW_ADD
   Next RT
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditFormula(Ug, IsOK, False, ErrorObj)
   
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   HasBegin = False
   IsOK = True

   CopyFormula = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyFormula = False
End Function

