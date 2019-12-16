VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPartMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member
'
Public PART_MASTER_ID  As Long
Public PART_MASTER_NO  As String
Public PART_MASTER_NAME As String
Public PART_MASTER_TYPE As Long
Public CANCEL_FLAG As String
Public ANIMAL_TYPE As Long

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public CREATE_DATE As Date
Public MODIFY_DATE As Date
Public CREATE_BY As Long
Public MODIFY_BY As Long

Public PartItems As Collection
Public CusGroups As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set PartItems = New Collection
   Set CusGroups = New Collection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim SubLen As Long
Dim NewStr As String
Dim TempDate As String

If Ind = 1 Then
   SelectStr1 = "SELECT PM.*,PM.CREATE_DATE CD,PM.MODIFY_DATE MD "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PART_MASTER PM "
'   SelectStr1 = "SELECT PM.*,PM.CREATE_DATE CD,PM.MODIFY_DATE MD, PT.PART_TYPE_NO,PT.PART_TYPE_NAME,PT.RAW_FLAG,PT.PART_GROUP_ID,PT.COMMISSION_FLAG,UN.UNIT_NAME,UN.UNIT_NO,UN.PERIOD_TYPE,LC.LOCATION_NAME DEFAULT_LOCATION_NAME,MR2.KEY_NAME ANIMAL_NAME  "
'   SelectStr2 = "SELECT COUNT(*)  "
'
'   FromStr = "FROM PART_MASTER PM "
'   FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PM.PART_MASTER_ID = PI.PART_MASTER_ID) "
'   FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
'   FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PG.PART_GROUP_ID = PT.PART_GROUP_ID) "
'   FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
'   FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (PI.DEFAULT_LOCATION = LC.LOCATION_ID) "
'   FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR2.KEY_ID = PI.ANIMAL_TYPE) "
'ElseIf Ind = 2 Then
'   SelectStr1 = "SELECT PM.* "
'   SelectStr2 = "SELECT COUNT(*)  "
'
'   FromStr = "FROM PART_MASTER PM "
   
End If
  
   
   WhereStr = ""
   
   If PART_MASTER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PM.PART_MASTER_ID = " & PART_MASTER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PM.PART_MASTER_ID = " & PART_MASTER_ID & ") "
      End If
   End If
   
   If Len(PART_MASTER_NO) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PM.PART_MASTER_NO = '" & PART_MASTER_NO & "') "
      Else
         WhereStr = WhereStr & "AND (PM.PART_MASTER_NO = '" & PART_MASTER_NO & "') "
      End If
   End If
   
   If Len(PART_MASTER_NAME) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PM.PART_MASTER_NAME = '" & PART_MASTER_NAME & "') "
      Else
         WhereStr = WhereStr & "AND (PM.PART_MASTER_NAME = '" & PART_MASTER_NAME & "') "
      End If
   End If
   
      If Len(CANCEL_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PM.CANCEL_FLAG = '" & CANCEL_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (PM.CANCEL_FLAG = '" & CANCEL_FLAG & "') "
      End If
   End If
   
   If ANIMAL_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PM.ANIMAL_TYPE = " & ANIMAL_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (PM.ANIMAL_TYPE = " & ANIMAL_TYPE & ") "
      End If
   End If
   
   If PART_MASTER_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PM.PART_MASTER_TYPE = " & PART_MASTER_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (PM.PART_MASTER_TYPE = " & PART_MASTER_TYPE & ") "
      End If
   End If
         
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Or Ind = 2 Then
     If OrderBy = 1 Then
         OrderByStr = "ORDER BY  PM.PART_MASTER_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY  PM.PART_MASTER_ID,PM.PART_MASTER_NO "
      Else
         OrderByStr = "ORDER BY  PM.PART_MASTER_ID "
      End If
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
  If Ind = 1 Then
      PART_MASTER_ID = NVLI(Rs("PART_MASTER_ID"), -1)
      PART_MASTER_NO = NVLS(Rs("PART_MASTER_NO"), "")
      PART_MASTER_NAME = NVLS(Rs("PART_MASTER_NAME"), "")
      PART_MASTER_TYPE = NVLI(Rs("PART_MASTER_TYPE"), -1)
      CANCEL_FLAG = NVLS(Rs("CANCEL_FLAG"), "")
      ANIMAL_TYPE = NVLI(Rs("ANIMAL_TYPE"), -1)
      
      KEY_ID = PART_MASTER_ID
      KEY_LOOKUP = PART_MASTER_NO
      
      CREATE_BY = NVLI(Rs("CREATE_BY"), -1)
      MODIFY_BY = NVLI(Rs("MODIFY_BY"), -1)
      CREATE_DATE = InternalDateToDate(NVLS(Rs("CD"), ""))
      MODIFY_DATE = InternalDateToDate(NVLS(Rs("MD"), ""))
'   ElseIf Ind = 2 Then
'      PART_MASTER_ID = NVLI(Rs("PART_MASTER_ID"), -1)
'      PART_MASTER_NO = NVLS(Rs("PART_MASTER_NO"), "")
'      PART_MASTER_NAME = NVLS(Rs("PART_MASTER_NAME"), "")
'      PART_MASTER_TYPE = NVLI(Rs("PART_MASTER_TYPE"), -1)
'      CANCEL_FLAG = NVLS(Rs("CANCEL_FLAG"), "")
'      ANIMAL_TYPE = NVLI(Rs("ANIMAL_TYPE"), -1)
'
'      KEY_ID = PART_MASTER_ID
'      KEY_LOOKUP = PART_MASTER_NO
   End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PART_MASTER_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO PART_MASTER " & _
                     "( " & _
                     "   PART_MASTER_ID, " & _
                     "   PART_MASTER_NO, " & _
                     "   PART_MASTER_NAME, " & _
                     "   PART_MASTER_TYPE, " & _
                     "   CANCEL_FLAG, " & _
                     "   ANIMAL_TYPE," & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY  " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(PART_MASTER_NO)) & "'," & _
                          "'" & ChangeQuote(Trim(PART_MASTER_NAME)) & "'," & _
                         "" & PART_MASTER_TYPE & "," & _
                         "'" & ChangeQuote(Trim(CANCEL_FLAG)) & "'," & _
                          ANIMAL_TYPE & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         PART_MASTER_ID = ID
   Else 'Show edit ANIMAL_TYPE
      SQL1 = "UPDATE PART_MASTER SET " & _
                     "   PART_MASTER_NO = '" & ChangeQuote(Trim(PART_MASTER_NO)) & "', " & _
                     "   PART_MASTER_NAME = '" & ChangeQuote(Trim(PART_MASTER_NAME)) & "', " & _
                     "   PART_MASTER_TYPE = " & PART_MASTER_TYPE & ", " & _
                     "   CANCEL_FLAG = '" & CANCEL_FLAG & "', " & _
                     "   ANIMAL_TYPE = " & ANIMAL_TYPE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & _
                     " WHERE PART_MASTER_ID = " & PART_MASTER_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM PART_MASTER WHERE PART_MASTER_ID = " & PART_MASTER_ID
      m_Conn.Execute (SQL1)
      
End Function

Private Sub Class_Terminate()
   Set PartItems = Nothing
   Set CusGroups = Nothing
End Sub
Public Function UpdateCancelFlag()
Dim SQL1 As String
      
      SQL1 = "UPDATE PART_MASTER SET CANCEL_FLAG =  '" & ChangeQuote(Trim(CANCEL_FLAG)) & "' WHERE PART_MASTER_ID = " & PART_MASTER_ID
      m_Conn.Execute (SQL1)
      
End Function

