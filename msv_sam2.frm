VERSION 5.00
Begin VB.Form msv_sam2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sachen MMC2 variables"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3465
   Icon            =   "msv_sam2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   3465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton msvsam2ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CheckBox var1b2 
      Caption         =   "OPT1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.CheckBox var1b1 
      Caption         =   "Return open bus on 512 KiB outer banks"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.CheckBox var1b0 
      Caption         =   "Return open bus on 1 MiB outer banks"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "msv_sam2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sam210 As Byte
Public sam211 As Byte
Public sam212 As Byte

Private Sub Form_Load()
    var1b0.Value = sam210
    var1b1.Value = sam211
    var1b2.Value = sam212
End Sub

Private Sub msvsam2ok_Click()
    sam210 = var1b0.Value
    sam211 = var1b1.Value
    sam212 = var1b2.Value
    config.sam210 = sam210
    config.sam211 = sam211
    config.sam212 = sam212
    Unload msv_sam2
End Sub
