VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ButtonEvent As msforms.CommandButton
Attribute ButtonEvent.VB_VarHelpID = -1
Private pButtonObj As msforms.CommandButton
Private pButtonName As String

Public WithEvents FormEvent As msforms.UserForm
Attribute FormEvent.VB_VarHelpID = -1
Private pFormObj As msforms.UserForm
Private pFormName As String

Public WithEvents TextEvent As msforms.TextBox
Attribute TextEvent.VB_VarHelpID = -1
Private pTextObj As msforms.TextBox
Private pTextName As String

Private Sub ButtonEvent_Click()
    Unload pFormObj
End Sub

Private Sub ButtonEvent_MouseMove(ByVal Button As Integer, _
                                  ByVal Shift As Integer, _
                                  ByVal X As Single, _
                                  ByVal Y As Single)
    With pButtonObj
        .ForeColor = vbWhite
        .BackColor = vbBlack
    End With
End Sub

Private Sub TextEvent_MouseMove(ByVal Button As Integer, _
                                ByVal Shift As Integer, _
                                ByVal X As Single, _
                                ByVal Y As Single)
    With pTextObj
        .ForeColor = vbWhite
        .BackColor = vbBlack
    End With
End Sub



Private Sub FormEvent_MouseMove(ByVal Button As Integer, _
                                ByVal Shift As Integer, _
                                ByVal X As Single, _
                                ByVal Y As Single)
    With pButtonObj
        .ForeColor = vbBlack
        .BackColor = vbWhite
    End With
    With pTextObj
        .ForeColor = vbBlack
        .BackColor = vbWhite
    End With
End Sub


Public Property Get FormName() As String: FormName = pFormName: End Property
Public Property Let FormName(ByVal BF As String): pFormName = BF: End Property

Public Property Get FormObj() As msforms.UserForm: Set FormObj = pFormObj: End Property
Public Property Set FormObj(ByVal Txt As msforms.UserForm): Set pFormObj = Txt: End Property


Public Property Get ButtonName() As String: ButtonName = pButtonName: End Property
Public Property Let ButtonName(ByVal BN As String): pButtonName = BN: End Property

Public Property Get ButtonObj() As msforms.CommandButton: Set ButtonObj = pButtonObj: End Property
Public Property Set ButtonObj(ByVal Btn As msforms.CommandButton): Set pButtonObj = Btn: End Property


Public Property Get TextName() As String: TextName = pTextName: End Property
Public Property Let TextName(ByVal BN As String): pTextName = BN: End Property

Public Property Get TextObj() As msforms.TextBox: Set TextObj = pTextObj: End Property
Public Property Set TextObj(ByVal Txt As msforms.TextBox): Set pTextObj = Txt: End Property


