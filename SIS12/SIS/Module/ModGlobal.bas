Attribute VB_Name = "ModGlobal"
Option Explicit
Public myConnectionString As String
Public libcon As New ADODB.Connection
Public MyRsUser As New ADODB.Recordset
Public MyRsProf As New ADODB.Recordset
Public MyRsCourse As New ADODB.Recordset
Public MyRsSearch As New ADODB.Recordset
Public SqlStr As String
Public Mypass As String
'this is my code to connect to DB SIS
Sub MyDBcon()
Set libcon = New ADODB.Connection
      libcon.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & _
      "SERVER=localhost;" & _
      "DATABASE=sis;" & _
      "UID=root;" & _
      "PASSWORD= root;" & _
      "OPTION=3;"
      
      libcon.Open
      
End Sub

'code to center all MDIchild forms
Sub FormCenter(Frm As Form)
    Frm.Top = (Screen.Height * 0.75) / 2 - Frm.Height / 2
    Frm.Left = Screen.Width / 2 - Frm.Width / 2
End Sub

Public Function MyTxtEmpty(ByRef MyText As Variant) As Boolean
If MyText.Text = "" Then
    MyTxtEmpty = True
    MsgBox "The field is required.Please check it!", vbExclamation, "SIS"
    MyText.SetFocus
Else
    MyTxtEmpty = False
End If

End Function
