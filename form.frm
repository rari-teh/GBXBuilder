VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form config 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GBXBuilder 1.0"
   ClientHeight    =   1995
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   3600
   Icon            =   "form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton bcancel 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton bsave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox ramtext 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      ForeColor       =   &H80000011&
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Text            =   "0"
      Top             =   480
      Width           =   1095
   End
   Begin VB.CheckBox istimer 
      Caption         =   "Timer"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CheckBox isrumble 
      Caption         =   "Rumble"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   3015
   End
   Begin VB.CheckBox isram 
      Caption         =   "Battery (RAM)"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.ComboBox mappertype 
      Height          =   315
      ItemData        =   "form.frx":1542
      Left            =   960
      List            =   "form.frx":1597
      TabIndex        =   1
      Text            =   "(no mapper)"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label labelb 
      Caption         =   "bytes"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   480
      Width           =   375
   End
   Begin VB.Label labelm 
      Caption         =   "Mapper"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu menutop 
      Caption         =   "&File"
      Index           =   1
      Begin VB.Menu menuopen 
         Caption         =   "&Open…"
         Index           =   11
         Shortcut        =   ^O
      End
      Begin VB.Menu menuabout 
         Caption         =   "&About GBXBuilder"
         Index           =   12
         Shortcut        =   {F1}
      End
      Begin VB.Menu menuexit 
         Caption         =   "E&xit"
         Index           =   13
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem ToByteArray dependency
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Rem Filename buffers
Private original As String
Private suggestion As String

Rem Initial declarations
Private mapper(3) As Byte
Private battery As Byte
Private rumble As Byte
Private timer As Byte
Private romsize As Long
Private romsizeba() As Byte
Private ramsize As Long
Private ramsizeba() As Byte
Rem Fixed declarations
Private footersize As Byte
Public major As Byte
Public minor As Byte
Rem Array declaration
Private footer(63) As Byte

Private Sub Form_Load()
On Error GoTo ErrHandler
    loadfile
    Rem Variable init
    mapper(0) = 82
    mapper(1) = 79
    mapper(2) = 77
    mapper(3) = 0
    battery = 0
    rumble = 0
    timer = 0
    romsize = LOF(1)
    ReDim romsizeba(3)
    romsizeba(0) = 0
    romsizeba(1) = 0
    romsizeba(2) = 0
    romsizeba(3) = 0
    ramsize = 0
    ReDim ramsizeba(3)
    ramsizeba(0) = 0
    ramsizeba(1) = 0
    ramsizeba(2) = 0
    ramsizeba(3) = 0
    Rem Set fixed values
    footersize = 64
    major = 1
    minor = 0
    Rem Array init
    Dim counter As Byte
    counter = 0
    Do While counter < 64
        footer(counter) = 0
        counter = counter + 1
    Loop
    Close
    Exit Sub
ErrHandler:
    Err.Clear
    Close
    Unload config
    Set config = Nothing
    Exit Sub
End Sub

Private Sub isram_Click()
    If isram.Value = 0 Then
        ramtext.Enabled = False
        ramtext.ForeColor = &H80000011
        ramtext.BackColor = &HE0E0E0
    Else
        ramtext.Enabled = True
        ramtext.ForeColor = &H80000007
        ramtext.BackColor = &H80000005
    End If
End Sub

Private Sub bsave_Click()
On Error GoTo ErrHandler
    If isram = 1 Then
        If ramtext.Text = 0 Then
            MsgBox "RAM size cannot be 0 bytes if battery is present.", vbOKOnly, "GBXBuilder Error"
            Exit Sub
        End If
    End If
    With CommonDialog1
        .CancelError = True
        .DialogTitle = "Save GBX file"
        .FileName = suggestion
        .Filter = "GBX format ROM (*.gbx)|*.gbx|"
        .ShowSave
    End With
    FileCopy original, CommonDialog1.FileName
    ramsize = ramtext.Text
    If ramtext.Enabled = False Then
        ramsize = 0
    End If
    battery = isram.Value
    rumble = isrumble.Value
    timer = istimer.Value
    Rem Figure mapper
    If mappertype.ListIndex = 0 Then
        mapper(0) = 82
        mapper(1) = 79
        mapper(2) = 77
        mapper(3) = 0
    ElseIf mappertype.ListIndex = 1 Then
        mapper(0) = 77
        mapper(1) = 66
        mapper(2) = 67
        mapper(3) = 49
    ElseIf mappertype.ListIndex = 2 Then
        mapper(0) = 77
        mapper(1) = 66
        mapper(2) = 67
        mapper(3) = 50
    ElseIf mappertype.ListIndex = 3 Then
        mapper(0) = 77
        mapper(1) = 66
        mapper(2) = 67
        mapper(3) = 51
    ElseIf mappertype.ListIndex = 4 Then
        mapper(0) = 77
        mapper(1) = 66
        mapper(2) = 67
        mapper(3) = 53
    ElseIf mappertype.ListIndex = 5 Then
        mapper(0) = 77
        mapper(1) = 66
        mapper(2) = 67
        mapper(3) = 55
    ElseIf mappertype.ListIndex = 6 Then
        mapper(0) = 77
        mapper(1) = 66
        mapper(2) = 49
        mapper(3) = 77
    ElseIf mappertype.ListIndex = 7 Then
        mapper(0) = 77
        mapper(1) = 77
        mapper(2) = 77
        mapper(3) = 49
    ElseIf mappertype.ListIndex = 8 Then
        mapper(0) = 67
        mapper(1) = 65
        mapper(2) = 77
        mapper(3) = 82
    ElseIf mappertype.ListIndex = 9 Then
        mapper(0) = 72
        mapper(1) = 85
        mapper(2) = 67
        mapper(3) = 49
    ElseIf mappertype.ListIndex = 10 Then
        mapper(0) = 72
        mapper(1) = 85
        mapper(2) = 67
        mapper(3) = 51
    ElseIf mappertype.ListIndex = 11 Then
        mapper(0) = 84
        mapper(1) = 65
        mapper(2) = 77
        mapper(3) = 53
    ElseIf mappertype.ListIndex = 12 Then
        mapper(0) = 77
        mapper(1) = 49
        mapper(2) = 54
        mapper(3) = 49
    ElseIf mappertype.ListIndex = 13 Then
        mapper(0) = 66
        mapper(1) = 66
        mapper(2) = 68
        mapper(3) = 0
    ElseIf mappertype.ListIndex = 14 Then
        mapper(0) = 72
        mapper(1) = 73
        mapper(2) = 84
        mapper(3) = 75
    ElseIf mappertype.ListIndex = 15 Then
        mapper(0) = 83
        mapper(1) = 78
        mapper(2) = 84
        mapper(3) = 88
    ElseIf mappertype.ListIndex = 16 Then
        mapper(0) = 78
        mapper(1) = 84
        mapper(2) = 79
        mapper(3) = 49
    ElseIf mappertype.ListIndex = 17 Then
        mapper(0) = 78
        mapper(1) = 84
        mapper(2) = 79
        mapper(3) = 50
    ElseIf mappertype.ListIndex = 18 Then
        mapper(0) = 78
        mapper(1) = 84
        mapper(2) = 78
        mapper(3) = 0
    ElseIf mappertype.ListIndex = 19 Then
        mapper(0) = 76
        mapper(1) = 73
        mapper(2) = 67
        mapper(3) = 72
    ElseIf mappertype.ListIndex = 20 Then
        mapper(0) = 76
        mapper(1) = 66
        mapper(2) = 77
        mapper(3) = 67
    ElseIf mappertype.ListIndex = 21 Then
        mapper(0) = 76
        mapper(1) = 73
        mapper(2) = 66
        mapper(3) = 65
    ElseIf mappertype.ListIndex = 22 Then
        mapper(0) = 80
        mapper(1) = 75
        mapper(2) = 74
        mapper(3) = 68
    ElseIf mappertype.ListIndex = 23 Then
        mapper(0) = 87
        mapper(1) = 73
        mapper(2) = 83
        mapper(3) = 68
    ElseIf mappertype.ListIndex = 24 Then
        mapper(0) = 83
        mapper(1) = 65
        mapper(2) = 77
        mapper(3) = 49
    ElseIf mappertype.ListIndex = 25 Then
        mapper(0) = 83
        mapper(1) = 65
        mapper(2) = 77
        mapper(3) = 50
    ElseIf mappertype.ListIndex = 26 Then
        mapper(0) = 82
        mapper(1) = 79
        mapper(2) = 67
        mapper(3) = 75
    End If
    Rem Mount array begin
    Rem Mapper
    footer(0) = mapper(0)
    footer(1) = mapper(1)
    footer(2) = mapper(2)
    footer(3) = mapper(3)
    Rem Battery
    footer(4) = battery
    Rem Rumble
    footer(5) = rumble
    Rem Timer
    footer(6) = timer
    Rem ROM size
    romsizeba = ToByteArray(romsize)
    footer(8) = romsizeba(3)
    footer(9) = romsizeba(2)
    footer(10) = romsizeba(1)
    footer(11) = romsizeba(0)
    Rem RAM size
    ramsizeba = ToByteArray(ramsize)
    footer(12) = ramsizeba(3)
    footer(13) = ramsizeba(2)
    footer(14) = ramsizeba(1)
    footer(15) = ramsizeba(0)
    Rem Footer
    footer(51) = footersize
    footer(55) = major
    footer(59) = minor
    footer(60) = 71
    footer(61) = 66
    footer(62) = 88
    footer(63) = 33
    Close
    Open CommonDialog1.FileName For Binary Access Read Write As #10
    Dim pos As Long
    pos = LOF(10) + 1
    Seek #10, pos
    Put #10, , footer
    Close
    Unload config
    Set config = Nothing
    Exit Sub
ErrHandler:
    Err.Clear
    Close
    Exit Sub
End Sub

Private Sub bcancel_Click()
    Close
    Unload config
    Set config = Nothing
End Sub

Private Sub menuexit_Click(Index As Integer)
    Close
    Unload config
    Set config = Nothing
End Sub

Private Sub menuopen_Click(Index As Integer)
On Error GoTo ErrHandler
    loadfile
ErrHandler:
    Err.Clear
    Close
    Exit Sub
End Sub

Private Sub menuabout_Click(Index As Integer)
    about.Show
End Sub

Public Sub loadfile()
        With CommonDialog1
        .CancelError = True
        .DialogTitle = "Open Game Boy (Color) ROM"
        .Filter = "Game Boy ROMs (*.gb; *.gbc; *.bin)|*.gb;*.gbc;*.bin|"
        .ShowOpen
    End With
    Open CommonDialog1.FileName For Binary As #1
    original = CommonDialog1.FileName
    Dim tempst As String
    tempst = Mid$(original, InStrRev(original, "\") + 1)
    suggestion = Left$(tempst, InStrRev(tempst, ".") - 1)
End Sub

Public Function ToByteArray(ByVal lng As Long) As Byte()
    Dim ByteArray(3) As Byte
    CopyMemory ByteArray(0), ByVal VarPtr(lng), Len(lng)
    ToByteArray = ByteArray
End Function
