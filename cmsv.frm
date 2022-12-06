VERSION 5.00
Begin VB.Form cmsv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set variables"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4935
   Icon            =   "cmsv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmsv_ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox sv8 
      Height          =   285
      Left            =   3720
      MaxLength       =   8
      TabIndex        =   7
      Text            =   "00000000"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox sv7 
      Height          =   285
      Left            =   2520
      MaxLength       =   8
      TabIndex        =   6
      Text            =   "00000000"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox sv6 
      Height          =   285
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   5
      Text            =   "00000000"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox sv5 
      Height          =   285
      Left            =   120
      MaxLength       =   8
      TabIndex        =   4
      Text            =   "00000000"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox sv4 
      Height          =   285
      Left            =   3720
      MaxLength       =   8
      TabIndex        =   3
      Text            =   "00000000"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox sv3 
      Height          =   285
      Left            =   2520
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "00000000"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox sv2 
      Height          =   285
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "00000000"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox sv1 
      Height          =   285
      Left            =   120
      MaxLength       =   8
      TabIndex        =   0
      Text            =   "00000000"
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "cmsv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cvar1 As Long
Public cvar2 As Long
Public cvar3 As Long
Public cvar4 As Long
Public cvar5 As Long
Public cvar6 As Long
Public cvar7 As Long
Public cvar8 As Long

Private Sub Form_Load()
    Dim signtrip As Boolean
    signtrip = False
    cvar1 = config.cvar1
    cvar2 = config.cvar2
    cvar3 = config.cvar3
    cvar4 = config.cvar4
    cvar5 = config.cvar5
    cvar6 = config.cvar6
    cvar7 = config.cvar7
    cvar8 = config.cvar8
    If (cvar1 < 0) Then
            signtrip = True
            cvar1 = cvar1 * -1
    End If
    sv1.Text = Right("00000000" & Hex$(cvar1), 8)
    If signtrip Then
            sv1.Text = Hex(CByte("&H" & Left(sv1.Text, 1)) + 8) & Right(sv1.Text, 7)
            signtrip = False
    End If
    If (cvar2 < 0) Then
            signtrip = True
            cvar2 = cvar2 * -1
    End If
    sv2.Text = Right("00000000" & Hex$(cvar2), 8)
    If signtrip Then
            sv2.Text = Hex(CByte("&H" & Left(sv2.Text, 1)) + 8) & Right(sv2.Text, 7)
            signtrip = False
    End If
    If (cvar3 < 0) Then
            signtrip = True
            cvar3 = cvar3 * -1
    End If
    sv3.Text = Right("00000000" & Hex$(cvar3), 8)
    If signtrip Then
            sv3.Text = Hex(CByte("&H" & Left(sv3.Text, 1)) + 8) & Right(sv3.Text, 7)
            signtrip = False
    End If
    If (cvar4 < 0) Then
            signtrip = True
            cvar4 = cvar4 * -1
    End If
    sv4.Text = Right("00000000" & Hex$(cvar4), 8)
    If signtrip Then
            sv4.Text = Hex(CByte("&H" & Left(sv4.Text, 1)) + 8) & Right(sv4.Text, 7)
            signtrip = False
    End If
    If (cvar5 < 0) Then
            signtrip = True
            cvar5 = cvar5 * -1
    End If
    sv5.Text = Right("00000000" & Hex$(cvar5), 8)
    If signtrip Then
            sv5.Text = Hex(CByte("&H" & Left(sv5.Text, 1)) + 8) & Right(sv5.Text, 7)
            signtrip = False
    End If
    If (cvar6 < 0) Then
            signtrip = True
            cvar6 = cvar6 * -1
    End If
    sv6.Text = Right("00000000" & Hex$(cvar6), 8)
    If signtrip Then
            sv6.Text = Hex(CByte("&H" & Left(sv6.Text, 1)) + 8) & Right(sv6.Text, 7)
            signtrip = False
    End If
    If (cvar7 < 0) Then
            signtrip = True
            cvar7 = cvar7 * -1
    End If
    sv7.Text = Right("00000000" & Hex$(cvar7), 8)
    If signtrip Then
            sv7.Text = Hex(CByte("&H" & Left(sv7.Text, 1)) + 8) & Right(sv7.Text, 7)
            signtrip = False
    End If
    If (cvar8 < 0) Then
            signtrip = True
            cvar8 = cvar8 * -1
    End If
    sv8.Text = Right("00000000" & Hex$(cvar8), 8)
    If signtrip Then
            sv8.Text = Hex(CByte("&H" & Left(sv8.Text, 1)) + 8) & Right(sv8.Text, 7)
            signtrip = False
    End If
End Sub

Private Sub cmsv_ok_Click()
On Error GoTo ErrHandler
    Dim signtrip As Boolean
    Dim highnybl As Byte
    signtrip = False
    highnybl = CByte("&H" & Left(sv1.Text, 1))
    If (highnybl > 7) Then
        signtrip = True
        sv1.Text = Hex$(highnybl - 8) & Right(sv1.Text, 7)
    End If
    config.cvar1 = CLng("&H" & sv1.Text)
    If signtrip Then
        config.cvar1 = config.cvar1 * -1
        signtrip = False
    End If
    highnybl = CByte("&H" & Left(sv2.Text, 1))
    If (highnybl > 7) Then
        signtrip = True
        sv2.Text = Hex$(highnybl - 8) & Right(sv2.Text, 7)
    End If
    config.cvar2 = CLng("&H" & sv2.Text)
    If signtrip Then
        config.cvar2 = config.cvar2 * -1
        signtrip = False
    End If
    highnybl = CByte("&H" & Left(sv3.Text, 1))
    If (highnybl > 7) Then
        signtrip = True
        sv3.Text = Hex$(highnybl - 8) & Right(sv3.Text, 7)
    End If
    config.cvar3 = CLng("&H" & sv3.Text)
    If signtrip Then
        config.cvar3 = config.cvar3 * -1
        signtrip = False
    End If
    highnybl = CByte("&H" & Left(sv4.Text, 1))
    If (highnybl > 7) Then
        signtrip = True
        sv4.Text = Hex$(highnybl - 8) & Right(sv4.Text, 7)
    End If
    config.cvar4 = CLng("&H" & sv4.Text)
    If signtrip Then
        config.cvar4 = config.cvar4 * -1
        signtrip = False
    End If
    highnybl = CByte("&H" & Left(sv5.Text, 1))
    If (highnybl > 7) Then
        signtrip = True
        sv5.Text = Hex$(highnybl - 8) & Right(sv5.Text, 7)
    End If
    config.cvar5 = CLng("&H" & sv5.Text)
    If signtrip Then
        config.cvar5 = config.cvar5 * -1
        signtrip = False
    End If
    highnybl = CByte("&H" & Left(sv6.Text, 1))
    If (highnybl > 7) Then
        signtrip = True
        sv6.Text = Hex$(highnybl - 8) & Right(sv6.Text, 7)
    End If
    config.cvar6 = CLng("&H" & sv6.Text)
    If signtrip Then
        config.cvar6 = config.cvar6 * -1
        signtrip = False
    End If
    highnybl = CByte("&H" & Left(sv7.Text, 1))
    If (highnybl > 7) Then
        signtrip = True
        sv7.Text = Hex$(highnybl - 8) & Right(sv7.Text, 7)
    End If
    config.cvar7 = CLng("&H" & sv7.Text)
    If signtrip Then
        config.cvar7 = config.cvar7 * -1
        signtrip = False
    End If
    highnybl = CByte("&H" & Left(sv8.Text, 1))
    If (highnybl > 7) Then
        signtrip = True
        sv8.Text = Hex$(highnybl - 8) & Right(sv8.Text, 7)
    End If
    config.cvar8 = CLng("&H" & sv8.Text)
    If signtrip Then
        config.cvar8 = config.cvar8 * -1
        signtrip = False
    End If
    Unload cmsv
    Exit Sub
ErrHandler:
    MsgBox "All variables must be set to a valid hexadecimal number.", vbOKOnly, "GBXBuilder Error"
    Err.Clear
End Sub

Private Sub sv1_GotFocus()
    sv1.SelStart = 0
    sv1.SelLength = Len(sv1.Text)
End Sub

Private Sub sv1_LostFocus()
    sv1.Text = Right("00000000" & sv1.Text, 8)
End Sub

Private Sub sv2_GotFocus()
    sv2.SelStart = 0
    sv2.SelLength = Len(sv2.Text)
End Sub

Private Sub sv2_LostFocus()
    sv2.Text = Right("00000000" & sv2.Text, 8)
End Sub

Private Sub sv3_GotFocus()
    sv3.SelStart = 0
    sv3.SelLength = Len(sv3.Text)
End Sub

Private Sub sv3_LostFocus()
    sv3.Text = Right("00000000" & sv3.Text, 8)
End Sub

Private Sub sv4_GotFocus()
    sv4.SelStart = 0
    sv4.SelLength = Len(sv4.Text)
End Sub

Private Sub sv4_LostFocus()
    sv4.Text = Right("00000000" & sv4.Text, 8)
End Sub

Private Sub sv5_GotFocus()
    sv5.SelStart = 0
    sv5.SelLength = Len(sv5.Text)
End Sub

Private Sub sv5_LostFocus()
    sv5.Text = Right("00000000" & sv5.Text, 8)
End Sub

Private Sub sv6_GotFocus()
    sv6.SelStart = 0
    sv6.SelLength = Len(sv6.Text)
End Sub

Private Sub sv6_LostFocus()
    sv6.Text = Right("00000000" & sv6.Text, 8)
End Sub

Private Sub sv7_GotFocus()
    sv7.SelStart = 0
    sv7.SelLength = Len(sv7.Text)
End Sub

Private Sub sv7_LostFocus()
    sv7.Text = Right("00000000" & sv7.Text, 8)
End Sub

Private Sub sv8_GotFocus()
    sv8.SelStart = 0
    sv8.SelLength = Len(sv8.Text)
End Sub

Private Sub sv8_LostFocus()
    sv8.Text = Right("00000000" & sv8.Text, 8)
End Sub
