VERSION 5.00
Begin VB.Form msv_vf01 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vast Fame original release variables"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4095
   Icon            =   "msv_vf01.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton vf01ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox initvalue 
      Height          =   315
      ItemData        =   "msv_vf01.frx":1542
      Left            =   1560
      List            =   "msv_vf01.frx":154C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label labelv10 
      Caption         =   "Initial running value"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "msv_vf01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vf0110 As Byte
Private origcount As Byte

Private Sub Form_Load()
On Error GoTo ErrHandler
    origcount = initvalue.ListCount
    If vf0110 = 16 Then
        initvalue.ListIndex = 1
    Else
        initvalue.ListIndex = vf0110
    End If
    Exit Sub
ErrHandler:
    initvalue.AddItem Hex$(vf0110)
    initvalue.ListIndex = (initvalue.ListCount - 1)
End Sub

Private Sub vf01ok_Click()
    If origcount <> initvalue.ListCount Then If initvalue.ListIndex = (initvalue.ListCount - 1) Then GoTo Finish
    If initvalue.ListIndex = 0 Then
        config.vf0110 = 0
    ElseIf initvalue.ListIndex = 1 Then
        config.vf0110 = 16
    End If
Finish:
    Unload msv_vf01
End Sub
