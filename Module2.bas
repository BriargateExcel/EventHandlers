Attribute VB_Name = "Module2"
Option Explicit
 
    
Sub MakeForm()
    Dim Btn As New EventHandler
    Dim Frm As New EventHandler
    Dim Txt As New EventHandler
    
'   Delete any existing forms
    Dim UserFrm As Object
    For Each UserFrm In ThisWorkbook.VBProject.VBComponents
        If UserFrm.Type = vbext_ct_MSForm Then
            ThisWorkbook.VBProject.VBComponents.Remove UserFrm
        End If
    Next UserFrm
   
'   Create the UserForm
    Dim TempForm As Object
    Dim pForm As Object
    Set TempForm = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    Set pForm = VBA.UserForms.Add(TempForm.Name)
  
    'Set Properties for the new form
    With pForm
        .Caption = "Temporary Form"
        .Width = 200
        .Height = 100
    End With
    With Frm
        Set .FormEvent = pForm
        Set .FormObj = pForm
    End With
    
'   Add a CommandButton
    Dim Ctl As Object
    Set Ctl = pForm.Controls.Add("forms.CommandButton.1")
    With Ctl
        .Caption = "Close"
        .Left = 60
        .Top = 40
        .ForeColor = vbBlack
        .BackColor = vbWhite
    End With
    Set Frm.ButtonObj = Ctl
    With Btn
        Set .ButtonEvent = Ctl
        Set .ButtonObj = Ctl
        Set .FormObj = pForm
    End With
    
'   Add a text field
    Set Ctl = pForm.Controls.Add("Forms.TextBox.1", "textfield", True)
    With Ctl
        .Top = 15
        .Left = 80
        .Height = 20
        .Width = 60
        .ForeColor = vbBlack
        .BackColor = vbWhite
        .Value = "My Text"
    End With
    Set Frm.TextObj = Ctl
    With Txt
        Set .TextEvent = Ctl
        Set .TextObj = Ctl
        Set .FormObj = pForm
    End With
        
'   Show the form
    pForm.Show

'   Delete the form
    ThisWorkbook.VBProject.VBComponents.Remove VBComponent:=TempForm
End Sub

