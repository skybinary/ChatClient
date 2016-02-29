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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Wrapper 
      BackColor       =   &H8000000C&
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
         Top             =   -615
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

Private topOffset As Integer

Private Sub ChatMsg1_GotFocus(Index As Integer)

End Sub

Private Sub ChatMsgList_msgClicky(Index As Integer)
    RaiseEvent clicky(ByVal Index)
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width < minWidth Then UserControl.Width = minWidth
    If UserControl.Height < minHeight Then UserControl.Height = minHeight
    Wrapper.Width = UserControl.Width - 240
    Wrapper.Height = UserControl.Height - 240
    
    VScroll1.Left = (Wrapper.Width + 240) - VScroll1.Width
    VScroll1.Height = Wrapper.Height
End Sub

Private Sub VScroll1_Change()
    topOffset = VScroll1.Value
    Wrapper_Resize
End Sub

Private Sub Wrapper_Resize()
    Dim thisWidth As Long
    Dim vw As Integer
    If VScroll1.Visible Then
        vw = VScroll1.Width - 60
    Else
        vw = 60
        topOffset = 0
    End If
    thisWidth = Wrapper.Width - vw
    Dim msg As ChatMsg
    Dim x As Integer
    For x = 0 To ChatMsgList.UBound
        ChatMsgList(x).Width = thisWidth
        Dim num As Long
        num = 615 * ((x - 1) - topOffset)
        
        ChatMsgList(x).Top = num
        ChatMsgList(x).Visible = True
    Next
End Sub

'////////// functions ?

Public Sub addNew(ByVal contents As String)
    Load ChatMsgList(ChatMsgList.UBound + 1)
    ChatMsgList(ChatMsgList.UBound).TestIndex = ChatMsgList.UBound
    updateBox
End Sub

Public Sub removeMsg(ByVal Index As Integer)
    Dim mcount As Integer
    Dim a As Integer
    For mcount = Index To ChatMsgList.UBound - 1
        swapEm mcount, mcount + 1
        'Set ChatMsgList(mcount) = ChatMsgList(mcount + 1)
    Next
    Unload ChatMsgList(ChatMsgList.UBound)
    updateBox
End Sub

Private Sub updateBox()
    Dim num As Long
    num = 615 * (ChatMsgList.UBound - 1)
    showHideScroll ((num + 615) > Wrapper.Height)
    If VScroll1.Visible Then
        Dim calcsize As Integer
        calcsize = UserControl.Height / 615 ' how many chatmsgs can fit in chatbox
        Dim calctwo As Integer
        calctwo = ChatMsgList.UBound - calcsize
        VScroll1.Min = 0
        VScroll1.Max = calctwo
    End If
    Wrapper_Resize
    Debug.Print calcsize
End Sub

Private Sub swapEm(ByVal dest As Integer, src As Integer)
    ChatMsgList(dest).TestIndex = ChatMsgList(src).TestIndex
    
End Sub

Sub showHideScroll(ByVal tf As Boolean)
    VScroll1.Visible = tf
    
End Sub
