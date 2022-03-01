VERSION 5.00
Begin VB.Form about 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About GBXBuilder"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2880
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton baboutok 
      Caption         =   "OK"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "GBX file format by taizou"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Build date: Mar 1, 2022"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "GBX version: 1.0"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Author: rari_teh"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "GBXBuilder 1.0 Rev G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub baboutok_Click()
    Unload about
End Sub
