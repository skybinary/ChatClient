VERSION 5.00
Begin VB.UserControl ChatMsg 
   BackColor       =   &H0000FFFF&
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7155
   ScaleHeight     =   1500
   ScaleWidth      =   7155
   ToolboxBitmap   =   "ChatMsg.ctx":0000
   Begin VB.Label Message 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4935
   End
End
Attribute VB_Name = "ChatMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private p_Testindex As Integer

Const minWidth As Long = 480
Const minHeight As Long = 480

Private iCount As Integer

Public Event msgClicky()

Public Property Let TestIndex(ByVal newVal As Integer)
    p_Testindex = newVal
    Message.Caption = "Test " & CStr(p_Testindex)
End Property

Public Property Get TestIndex() As Integer
    TestIndex = p_Testindex
End Property

Private Sub Message_Click()
    RaiseEvent msgClicky
End Sub

Private Sub UserControl_Initialize()
    iCount = iCount + 1
    Me.TestIndex = iCount
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width < minWidth Then UserControl.Width = minWidth
    If UserControl.Height < minHeight Then UserControl.Height = minHeight
'    Debug.Print UserControl.Width & "," & UserControl.Height
    Message.Width = UserControl.Width - 60
    Message.Height = UserControl.Height - 60
End Sub

