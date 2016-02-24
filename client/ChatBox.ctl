VERSION 5.00
Begin VB.UserControl ChatBox 
   BackStyle       =   0  'Transparent
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   ScaleHeight     =   4290
   ScaleWidth      =   9000
   ToolboxBitmap   =   "ChatBox.ctx":0000
   Begin VB.VScrollBar VScroll1 
      Height          =   3855
      Left            =   6480
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Wrapper 
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3915
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin prjFN33Client.ChatMsg ChatMsgList 
         Height          =   615
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1085
      End
      Begin prjFN33Client.ChatMsg ChatMsgList 
         Height          =   615
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   600
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1085
      End
   End
End
Attribute VB_Name = "ChatBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const minWidth As Long = 480
Const minHeight As Long = 480

Public Event clicky(ByVal Index As Integer)

Private Sub ChatMsg1_GotFocus(Index As Integer)

End Sub

Private Sub ChatMsgList_msgClicky(Index As Integer)
    RaiseEvent clicky(ByVal Index)
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width < minWidth Then UserControl.Width = minWidth
    If UserControl.Height < minHeight Then UserControl.Height = minHeight
  '  Debug.Print UserControl.Width & "," & UserControl.Height
    Wrapper.Width = UserControl.Width - 240
    Wrapper.Height = UserControl.Height - 240
    
End Sub

Private Sub Wrapper_Resize()
    Dim thisWidth As Long
    thisWidth = Wrapper.Width - 60
    Dim msg As chatmsg
    Dim x As Integer
    For x = 0 To ChatMsgList.UBound
        ChatMsgList(x).Width = thisWidth
    Next
End Sub

'////////// functions ?

Public Sub addNew(ByVal contents As String)
    Load ChatMsgList(ChatMsgList.UBound + 1)
    Dim num As Long
    num = 615 * ChatMsgList.UBound
    
    ChatMsgList(ChatMsgList.UBound).Top = num
    ChatMsgList(ChatMsgList.UBound).Visible = True
    
    showHideScroll ((num + 615) > Wrapper.Height)
End Sub

Public Sub removeMsg(ByVal Index As Integer)
    Unload ChatMsgList(Index)
End Sub

Sub showHideScroll(ByVal tf As Boolean)
    VScroll1.Visible = tf
    
End Sub
