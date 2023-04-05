VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Escalation Matrix Automated"
   ClientHeight    =   11415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20280
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
'ME CAGO EN BRYAAAAN
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Delete_Click()
    Dim answer As Integer
    answer = MsgBox("Are you sure to Delete this Info?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")

    If answer = vbYes Then
        MsgBox "Done!"
        Range("A" & row_texto).EntireRow.Delete
    Else
        MsgBox "Cancelled!"
    End If
    
    Call Save
    Call NamedRangeDynamic
End Sub

Private Sub btn_Edit_Click()
    ListBox1.Locked = True
    txt_Search.Locked = True
    btn_Insert.Enabled = False
    btn_Delete.Enabled = False
    btn_Edit.Enabled = False
    lbl_EditMode.Visible = True
    btn_Save.Visible = True
    
    Call UnLockedBoxes
End Sub

Private Sub btn_Insert_Click()
    txt_Search.Text = ""
    txt_Search.Locked = True
    btn_SaveInsert.Visible = True
    lbl_InsertMode.Visible = True
    btn_Insert.Enabled = False
    ListBox1.Locked = True
    btn_Delete.Enabled = False
    btn_Edit.Enabled = False
    
    Call NamedRangeDynamic
    ListBox1.Selected(ListBox1.ListIndex) = False
    Call ClearBoxes
    Call UnLockedBoxes
End Sub

Private Sub btn_Save_Click()
Dim answer As Integer
    answer = MsgBox("Are you sure to Edit " & txt_Component.Text & " Info?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")

    If answer = vbYes Then
        MsgBox "Done!"
        Range("A" & row_texto).Value = UserForm1.txt_Component.Value
        Range("B" & row_texto).Value = UserForm1.txt_Environment.Value
        Range("E" & row_texto).Value = UserForm1.txt_SevType.Value
        Range("F" & row_texto).Value = UserForm1.txt_EscalateTo.Value
        Range("G" & row_texto).Value = UserForm1.txt_Access.Value
        Range("C" & row_texto).Value = UserForm1.txt_URL.Value
        Range("H" & row_texto).Value = UserForm1.txt_Description.Value
    Else
        MsgBox "Cancelled!"
    End If

    btn_Save.Visible = False
    lbl_EditMode.Visible = False
    btn_Delete.Enabled = True
    btn_Insert.Enabled = True
    btn_Edit.Enabled = True
    ListBox1.Locked = False
    txt_Search.Locked = False
    Call ClearBoxes
    Call LockedBoxes
    Call Save
    Call NamedRangeDynamic
End Sub

Private Sub btn_SaveInsert_Click()
Dim answer As Integer
answer = MsgBox("Are you sure to Insert this Info?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")

    If answer = vbYes Then
        MsgBox "Done!"
        Range("A2").EntireRow.Insert
        Range("A2").EntireRow.Interior.ColorIndex = 0
        Range("A2").EntireRow.Font.Bold = False
        Range("A2").Value = UserForm1.txt_Component.Value
        Range("B2").Value = UserForm1.txt_Environment.Value
        Range("E2").Value = UserForm1.txt_SevType.Value
        Range("F2").Value = UserForm1.txt_EscalateTo.Value
        Range("G2").Value = UserForm1.txt_Access.Value
        Range("C2").Value = UserForm1.txt_URL.Value
        Range("H2").Value = UserForm1.txt_Description.Value
    Else
        MsgBox "Cancelled!"
    End If
    
    Call LockedBoxes
    Call ClearBoxes
    lbl_InsertMode.Visible = True
    btn_SaveInsert.Visible = False
    btn_Delete.Enabled = True
    btn_Edit.Enabled = True
    btn_Insert.Enabled = True
    Call Save
    Call NamedRangeDynamic
    ListBox1.Locked = False
    txt_Search.Locked = False
    lbl_InsertMode.Visible = False
End Sub

Private Sub Image1_Click()
Unload Me
frm_Menu.Show
End Sub

Private Sub Label1_Click()

End Sub

Private Sub ListBox1_Change()
Dim texto As String

texto = ListBox1.Text
row_texto = ThisWorkbook.Sheets("Mio").Range("A:A").Find(texto, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Row

    txt_Component.Text = Range("A" & row_texto)
    txt_Environment.Text = Range("B" & row_texto)
    txt_SevType.Text = Range("E" & row_texto)
    txt_EscalateTo.Text = Range("F" & row_texto)
    txt_Access.Text = Range("G" & row_texto)
    txt_URL.Text = Range("C" & row_texto)
    txt_Description.Text = Range("H" & row_texto)

    If ListBox1.ListIndex = -1 Then
       btn_Edit.Enabled = False
       btn_Delete.Enabled = False
    Else
        btn_Edit.Enabled = True
        btn_Delete.Enabled = True
    End If
    
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub txt_Search_Change()
If txt_Search.Text <> "" Then
    Label6.Visible = False
    Label7.Visible = False
    Dim i As Long
    Dim arrList As Variant

    Me.ListBox1.Clear
    If Sheets("Mio").Range("A" & Sheets("Mio").Rows.Count).End(xlUp).Row > 1 And Trim(Me.txt_Search.Value) <> vbNullString Then
        arrList = Sheets("Mio").Range("A1:A" & Sheets("Mio").Range("A" & Sheets("Mio").Rows.Count).End(xlUp).Row).Value
        For i = LBound(arrList) To UBound(arrList)
            If InStr(1, arrList(i, 1), Trim(Me.txt_Search.Value), vbTextCompare) Then
                Me.ListBox1.AddItem arrList(i, 1)
            End If
        Next i
    End If
    If Me.ListBox1.ListCount = 1 Then Me.ListBox1.Selected(0) = True
    
    If ListBox1.ListIndex = -1 Then
       btn_Edit.Enabled = False
       btn_Delete.Enabled = False
    Else
        btn_Edit.Enabled = True
        btn_Delete.Enabled = True
    End If
    
Else
    Label6.Visible = True
    Label7.Visible = True
    Call NamedRangeDynamic
End If

End Sub

Private Sub txt_Search_Enter()

End Sub

Private Sub UserForm_Click()

End Sub



Private Sub UserForm_Initialize()
    Label3.Caption = Application.UserName
    Label4.Caption = Format(Date, "dddd, mmmm dd, yyyy")
    Call NamedRangeDynamic
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   If CloseMode = VbQueryClose.vbFormControlMenu Then Cancel = True
End Sub

'Private Sub UserForm_Terminate()
'    UserForm1.Hide
'    frm_Menu.Show
    'ThisWorkbook.Saved = True
   ' Application.Quit
'End Sub
