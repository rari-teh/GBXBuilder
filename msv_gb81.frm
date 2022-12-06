VERSION 5.00
Begin VB.Form msv_gb81 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vast Fame secondary release variables"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4095
   Icon            =   "msv_gb81.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton gb81ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox pcbtype 
      Height          =   315
      ItemData        =   "msv_gb81.frx":1542
      Left            =   600
      List            =   "msv_gb81.frx":154C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label labelv10 
      Caption         =   "PCB"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "msv_gb81"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gb8110 As Byte
Private origcount As Byte

Private Sub Form_Load()
On Error GoTo ErrHandler
    origcount = pcbtype.ListCount
    pcbtype.ListIndex = gb8110
    Exit Sub
ErrHandler:
    pcbtype.AddItem Right("00" & Hex$(gb8110), 2)
    pcbtype.ListIndex = (pcbtype.ListCount - 1)
End Sub

Private Sub gb81ok_Click()
    If origcount <> pcbtype.ListCount Then If pcbtype.ListIndex = (pcbtype.ListCount - 1) Then GoTo Finish
    config.gb8110 = pcbtype.ListIndex
Finish:
    Unload msv_gb81
End Sub
