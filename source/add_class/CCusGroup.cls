VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCusGroup"
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

Public PART_CUS_GROUPS_ID  As Long
Public CUS_GROUPS_ID As Long
Public CUS_GROUPS_NO As String
Public CUS_GROUPS_NAME As String
Public PART_MASTER_ID  As Long

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public CREATE_DATE As Date
Public MODIFY_DATE As Date
Public CREATE_BY As Long
Public MODIFY_BY As Long

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
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
   SelectStr1 = "SELECT PCG.*,CT.CSTTYPE_NO CUS_GROUPS_NO,CT.CSTTYPE_NAME CUS_GROUPS_NAME "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PART_CUS_GROUPS PCG "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_TYPE CT ON (CT.CSTTYPE_ID = PCG.CUS_GROUPS_ID) "
   
End If
  
   
   WhereStr = ""
   
   If PART_CUS_GROUPS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PCG.PART_CUS_GROUPS_ID = " & PART_CUS_GROUPS_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PCG.PART_CUS_GROUPS_ID = " & PART_CUS_GROUPS_ID & ") "
      End If
   End If
   
   If CUS_GROUPS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PCG.CUS_GROUPS_ID = " & CUS_GROUPS_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PCG.CUS_GROUPS_ID = " & CUS_GROUPS_ID & ") "
      End If
   End If
   
   If PART_MASTER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PCG.PART_MASTER_ID = " & PART_MASTER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PCG.PART_MASTER_ID = " & PART_MASTER_ID & ") "
      End If
   End If
 
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
     If OrderBy = 1 Then
         OrderByStr = "ORDER BY  PCG.PART_CUS_GROUPS_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY  PCG.CUS_GROUPS_NAME "
      Else
         OrderByStr = "ORDER BY  PCG.PART_CUS_GROUPS_ID "
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
      PART_CUS_GROUPS_ID = NVLI(Rs("PART_CUS_GROUPS_ID"), -1)
      CUS_GROUPS_ID = NVLI(Rs("CUS_GROUPS_ID"), -1)
      CUS_GROUPS_NO = NVLS(Rs("CUS_GROUPS_NO"), "")
      CUS_GROUPS_NAME = NVLS(Rs("CUS_GROUPS_NAME"), "")
      PART_MASTER_ID = NVLI(Rs("PART_MASTER_ID"), -1)
   
      
      CREATE_BY = NVLI(Rs("CREATE_BY"), -1)
      MODIFY_BY = NVLI(Rs("MODIFY_BY"), -1)
   End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PART_CUS_GROUPS_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO PART_CUS_GROUPS " & _
                     "( " & _
                     "   PART_CUS_GROUPS_ID, " & _
                     "   CUS_GROUPS_ID, " & _
                     "   PART_MASTER_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY  " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         CUS_GROUPS_ID & "," & _
                         PART_MASTER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         PART_CUS_GROUPS_ID = ID
   Else 'Show edit ANIMAL_TYPE
      SQL1 = "UPDATE PART_CUS_GROUPS SET " & _
                     "   CUS_GROUPS_ID = " & CUS_GROUPS_ID & ", " & _
                     "   PART_MASTER_ID = " & PART_MASTER_ID & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & _
                     " WHERE PART_CUS_GROUPS_ID = " & PART_CUS_GROUPS_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM PART_CUS_GROUPS  WHERE PART_CUS_GROUPS_ID = " & PART_CUS_GROUPS_ID
      m_Conn.Execute (SQL1)

End Function


