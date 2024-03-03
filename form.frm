VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form config 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GBXBuilder 1.0"
   ClientHeight    =   1560
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   3600
   Icon            =   "form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton button_msv 
      Caption         =   "Mapper variables..."
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Width           =   735
   End
   Begin VB.CheckBox isrumble 
      Caption         =   "Rumble"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   855
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
      ItemData        =   "form.frx":1042
      Left            =   960
      List            =   "form.frx":10A6
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
      Begin VB.Menu menuopen 
         Caption         =   "Change R&OM…"
         Shortcut        =   ^O
      End
      Begin VB.Menu menusave 
         Caption         =   "&Save GBX file…"
         Shortcut        =   ^S
      End
      Begin VB.Menu menustrip 
         Caption         =   "Save ra&w image…"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu menuexit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu menuopt 
      Caption         =   "Op&tions"
      Begin VB.Menu menucmsv 
         Caption         =   "&Custom mapper variables…"
         Shortcut        =   {F12}
      End
      Begin VB.Menu menureset 
         Caption         =   "&Reset mapper-specific variables"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu menuhelp 
      Caption         =   "&Help"
      Begin VB.Menu menuabout 
         Caption         =   "&About GBXBuilder…"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem Dependencies
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hfile As Long) As Long

Rem Buffers
Private original As String
Private suggestion As String
Private prevoriginal As String
Private prevsuggestion As String
Private previsgbx As Boolean
Private arg As String

Rem Flags
Private isgbx As Boolean
Private gbx1 As Boolean

Rem Initial declarations
Private mapper(3) As Byte
Private battery As Byte
Private rumble As Byte
Private timer As Byte
Private romsize As Long
Private romsizeba() As Byte
Private ramsize As Long
Private ramsizeba() As Byte
Private var1 As Long
Private var1ba() As Byte
Private var2 As Long
Private var2ba() As Byte
Private var3 As Long
Private var3ba() As Byte
Private var4 As Long
Private var4ba() As Byte
Private var5 As Long
Private var5ba() As Byte
Private var6 As Long
Private var6ba() As Byte
Private var7 As Long
Private var7ba() As Byte
Private var8 As Long
Private var8ba() As Byte
Rem Fixed declarations
Private footersize As Byte
Public major As Byte
Public minor As Byte
Private magic() As Byte
Rem Array declaration
Private footer(63) As Byte

Rem MSV config forms communication variables
Public sam210 As Byte
Public sam211 As Byte
Public sam212 As Byte
Public vf0110 As Byte
Public gb8110 As Byte
Public cvar1 As Long
Public cvar2 As Long
Public cvar3 As Long
Public cvar4 As Long
Public cvar5 As Long
Public cvar6 As Long
Public cvar7 As Long
Public cvar8 As Long

Private Sub button_msv_Click()
    Dim signtrip As Boolean
    signtrip = False
    If mappertype.ListIndex = 26 Then
        var1 = cvar1
        var1ba = ToByteArray(var1)
        msv_sam2.sam210 = var1ba(3) Mod 2
        msv_sam2.sam211 = (var1ba(3) \ 2) Mod 2
        msv_sam2.sam212 = (var1ba(3) \ 4) Mod 2
        msv_sam2.Show 1
        var1ba(3) = sam210 + (2 * sam211) + (4 * sam212)
        var1 = ToLong(var1ba)
        refreshmsvs
        verifymsvs
    ElseIf mappertype.ListIndex = 29 Then
        GoTo ListTypeInit
Ret29:
        vf0110 = var1ba(3)
        msv_vf01.vf0110 = var1ba(3)
        msv_vf01.Show 1
        var1ba(3) = vf0110
        GoTo ListTypeEnd
    ElseIf mappertype.ListIndex = 30 Then
        GoTo ListTypeInit
Ret30:
        gb8110 = var1ba(3)
        msv_gb81.gb8110 = var1ba(3)
        msv_gb81.Show 1
        var1ba(3) = gb8110
        GoTo ListTypeEnd
    Else
        MsgBox "This button should be disabled. Please file a bug on GitHub. (ERROR 1046)", vbOKOnly, "GBXBuilder Error"
    End If
    Exit Sub
ListTypeInit:
    var1 = cvar1
    If (var1 < 0) Then
        signtrip = True
        var1 = var1 * -1
    End If
    var1ba = ToByteArray(var1)
    If signtrip Then
        var1ba(3) = var1ba(3) + 128
        signtrip = False
    End If
    If (mappertype.ListIndex = 29) Then GoTo Ret29
    If (mappertype.ListIndex = 30) Then GoTo Ret30
    MsgBox "If this popup appears, please close the program and file a bug on GitHub. (ERROR 1046-2)", vbOKOnly, "GBXBuilder Error"
        Exit Sub
ListTypeEnd:
    If (var1ba(3) > 127) Then
        signtrip = True
        var1ba(3) = var1ba(3) - 128
    End If
    var1 = ToLong(var1ba)
    If signtrip Then
        var1 = var1 * -1
        signtrip = False
    End If
    refreshmsvs
    verifymsvs
        Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    arg = Command
    gbx1 = True
    isgbx = False
    original = ""
    suggestion = ""
    Rem Variable init
    mapper(0) = 82
    mapper(1) = 79
    mapper(2) = 77
    mapper(3) = 0
    battery = 0
    rumble = 0
    timer = 0
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
    ReDim var1ba(3)
    ReDim var2ba(3)
    ReDim var3ba(3)
    ReDim var4ba(3)
    ReDim var5ba(3)
    ReDim var6ba(3)
    ReDim var7ba(3)
    ReDim var8ba(3)
    clearmsvs
    sam210 = 0
    sam211 = 0
    sam212 = 0
    gb8110 = 0
    Rem Set fixed values
    footersize = 64
    major = 1
    minor = 0
    ReDim magic(3)
    magic(0) = 71
    magic(1) = 66
    magic(2) = 88
    magic(3) = 33
    Rem Array init
    Dim counter As Byte
    counter = 0
    Do While counter < 64
        footer(counter) = 0
        counter = counter + 1
    Loop
    Rem Load file
BackToLoad:
    loadfile
    If gbx1 = False Then
        MsgBox "This GBX file is in a newer format than this version of GBXBuilder can handle.", vbOKOnly, "GBXBuilder Error"
        gbx1 = True
        GoTo BackToLoad
    End If
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

Private Sub mappertype_Change()
    If mappertype.ListIndex = 26 Or mappertype.ListIndex = 29 Or mappertype.ListIndex = 30 Then button_msv.Enabled = True Else button_msv.Enabled = False
End Sub

Private Sub mappertype_Click()
    If mappertype.ListIndex = 26 Or mappertype.ListIndex = 29 Or mappertype.ListIndex = 30 Then button_msv.Enabled = True Else button_msv.Enabled = False
End Sub

Private Sub menucmsv_Click()
    cmsv.Show 1
    var1 = cvar1
    var2 = cvar2
    var3 = cvar3
    var4 = cvar4
    var5 = cvar5
    var6 = cvar6
    var7 = cvar7
    var8 = cvar8
    verifymsvs
End Sub

Private Sub verifymsvs()
    Dim msvset As Boolean
    msvset = False
    If var1 <> 0 Then msvset = True
    If var2 <> 0 Then msvset = True
    If var3 <> 0 Then msvset = True
    If var4 <> 0 Then msvset = True
    If var5 <> 0 Then msvset = True
    If var6 <> 0 Then msvset = True
    If var7 <> 0 Then msvset = True
    If var8 <> 0 Then msvset = True
    If msvset = True Then
        menureset.Checked = False
        menureset.Enabled = True
    Else
        menureset.Checked = True
        menureset.Enabled = False
    End If
End Sub

Private Sub menusave_Click()
On Error GoTo ErrHandler
    If isram = 1 Then
        If ramtext.Text = 0 Then
            MsgBox "RAM size cannot be 0 bytes if battery is present.", vbOKOnly, "GBXBuilder Error"
            Exit Sub
        End If
    End If
    If mappertype.ListIndex < 0 Then
        If Len(mappertype.Text) > 4 Then
            If mappertype.Text = "(no mapper)" Then
                mappertype.Text = "ROM"
            Else
                MsgBox "Custom mapper identifier cannot have more than four characters.", vbOKOnly, "GBXBuilder Error"
                Exit Sub
            End If
        End If
        If Len(mappertype.Text) = 0 Then
            mappertype.ListIndex = 0
        Else
            Dim char As Byte
            char = 0
            Do
                mapper(char) = FromAscii(mappertype.Text, char + 1)
                char = char + 1
            Loop While char < 4
        End If
    End If
    With CommonDialog1
        .CancelError = True
        .DialogTitle = "Save GBX file"
        .filename = suggestion
        .Filter = "GBX format ROM (*.gbx)|*.gbx|"
        .ShowSave
    End With
    continuesave
ErrHandler:
    Err.Clear
    Close
    Exit Sub
End Sub

Private Sub continuesave()
On Error GoTo ErrHandler
    FileCopy original, CommonDialog1.filename
    ramsize = ramtext.Text
    If ramtext.Enabled = False Then
        ramsize = 0
    End If
    battery = isram.Value
    rumble = isrumble.Value
    timer = istimer.Value
    Rem Figure mapper
    Rem ROM
    If mappertype.ListIndex = 0 Then
        mapper(0) = 82
        mapper(1) = 79
        mapper(2) = 77
        mapper(3) = 0
    Rem MBC1
    ElseIf mappertype.ListIndex = 1 Then
        mapper(0) = 77
        mapper(1) = 66
        mapper(2) = 67
        mapper(3) = 49
    Rem MBC2
    ElseIf mappertype.ListIndex = 2 Then
        mapper(0) = 77
        mapper(1) = 66
        mapper(2) = 67
        mapper(3) = 50
    Rem MBC3
    ElseIf mappertype.ListIndex = 3 Then
        mapper(0) = 77
        mapper(1) = 66
        mapper(2) = 67
        mapper(3) = 51
    Rem MBC5
    ElseIf mappertype.ListIndex = 4 Then
        mapper(0) = 77
        mapper(1) = 66
        mapper(2) = 67
        mapper(3) = 53
    Rem MBC6
    ElseIf mappertype.ListIndex = 5 Then
        mapper(0) = 77
        mapper(1) = 66
        mapper(2) = 67
        mapper(3) = 54
    Rem MBC7
    ElseIf mappertype.ListIndex = 6 Then
        mapper(0) = 77
        mapper(1) = 66
        mapper(2) = 67
        mapper(3) = 55
    Rem MB1M
    ElseIf mappertype.ListIndex = 7 Then
        mapper(0) = 77
        mapper(1) = 66
        mapper(2) = 49
        mapper(3) = 77
    Rem MMM1
    ElseIf mappertype.ListIndex = 8 Then
        mapper(0) = 77
        mapper(1) = 77
        mapper(2) = 77
        mapper(3) = 49
    Rem CAMR
    ElseIf mappertype.ListIndex = 9 Then
        mapper(0) = 67
        mapper(1) = 65
        mapper(2) = 77
        mapper(3) = 82
    Rem HUC1
    ElseIf mappertype.ListIndex = 10 Then
        mapper(0) = 72
        mapper(1) = 85
        mapper(2) = 67
        mapper(3) = 49
    Rem HUC3
    ElseIf mappertype.ListIndex = 11 Then
        mapper(0) = 72
        mapper(1) = 85
        mapper(2) = 67
        mapper(3) = 51
    Rem TAM5
    ElseIf mappertype.ListIndex = 12 Then
        mapper(0) = 84
        mapper(1) = 65
        mapper(2) = 77
        mapper(3) = 53
    Rem M161
    ElseIf mappertype.ListIndex = 13 Then
        mapper(0) = 77
        mapper(1) = 49
        mapper(2) = 54
        mapper(3) = 49
    Rem BBD
    ElseIf mappertype.ListIndex = 14 Then
        mapper(0) = 66
        mapper(1) = 66
        mapper(2) = 68
        mapper(3) = 0
    Rem HITK
    ElseIf mappertype.ListIndex = 15 Then
        mapper(0) = 72
        mapper(1) = 73
        mapper(2) = 84
        mapper(3) = 75
    Rem SNTX
    ElseIf mappertype.ListIndex = 16 Then
        mapper(0) = 83
        mapper(1) = 78
        mapper(2) = 84
        mapper(3) = 88
    Rem NTO1
    ElseIf mappertype.ListIndex = 17 Then
        mapper(0) = 78
        mapper(1) = 84
        mapper(2) = 79
        mapper(3) = 49
    Rem NTO2
    ElseIf mappertype.ListIndex = 18 Then
        mapper(0) = 78
        mapper(1) = 84
        mapper(2) = 79
        mapper(3) = 50
    Rem NTN
    ElseIf mappertype.ListIndex = 19 Then
        mapper(0) = 78
        mapper(1) = 84
        mapper(2) = 78
        mapper(3) = 0
    Rem LICH
    ElseIf mappertype.ListIndex = 20 Then
        mapper(0) = 76
        mapper(1) = 73
        mapper(2) = 67
        mapper(3) = 72
    Rem LBMC
    ElseIf mappertype.ListIndex = 21 Then
        mapper(0) = 76
        mapper(1) = 66
        mapper(2) = 77
        mapper(3) = 67
    Rem LIBA
    ElseIf mappertype.ListIndex = 22 Then
        mapper(0) = 76
        mapper(1) = 73
        mapper(2) = 66
        mapper(3) = 65
    Rem PKJD
    ElseIf mappertype.ListIndex = 23 Then
        mapper(0) = 80
        mapper(1) = 75
        mapper(2) = 74
        mapper(3) = 68
    Rem WISD
    ElseIf mappertype.ListIndex = 24 Then
        mapper(0) = 87
        mapper(1) = 73
        mapper(2) = 83
        mapper(3) = 68
    Rem SAM1
    ElseIf mappertype.ListIndex = 25 Then
        mapper(0) = 83
        mapper(1) = 65
        mapper(2) = 77
        mapper(3) = 49
    Rem SAM2
    ElseIf mappertype.ListIndex = 26 Then
        mapper(0) = 83
        mapper(1) = 65
        mapper(2) = 77
        mapper(3) = 50
    Rem ROCK
    ElseIf mappertype.ListIndex = 27 Then
        mapper(0) = 82
        mapper(1) = 79
        mapper(2) = 67
        mapper(3) = 75
    Rem NGHK
    ElseIf mappertype.ListIndex = 28 Then
        mapper(0) = 78
        mapper(1) = 71
        mapper(2) = 72
        mapper(3) = 75
    Rem VF01
    ElseIf mappertype.ListIndex = 29 Then
        mapper(0) = 86
        mapper(1) = 70
        mapper(2) = 48
        mapper(3) = 49
    Rem GB81
    ElseIf mappertype.ListIndex = 30 Then
        mapper(0) = 71
        mapper(1) = 66
        mapper(2) = 56
        mapper(3) = 49
    Rem TPP1
    ElseIf mappertype.ListIndex = 31 Then
        mapper(0) = 84
        mapper(1) = 80
        mapper(2) = 80
        mapper(3) = 49
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
    Rem MSVs
    If var1 < 0 Then
        var1 = var1 * (-1)
        var1ba = ToByteArray(var1)
        var1ba(3) = var1ba(3) + 128
    Else
        var1ba = ToByteArray(var1)
    End If
    footer(16) = var1ba(3)
    footer(17) = var1ba(2)
    footer(18) = var1ba(1)
    footer(19) = var1ba(0)
    If var2 < 0 Then
        var2 = var2 * (-1)
        var2ba = ToByteArray(var2)
        var2ba(3) = var2ba(3) + 128
    Else
        var2ba = ToByteArray(var2)
    End If
    footer(20) = var2ba(3)
    footer(21) = var2ba(2)
    footer(22) = var2ba(1)
    footer(23) = var2ba(0)
    If var3 < 0 Then
        var3 = var3 * (-1)
        var3ba = ToByteArray(var3)
        var3ba(3) = var3ba(3) + 128
    Else
        var3ba = ToByteArray(var3)
    End If
    footer(24) = var3ba(3)
    footer(25) = var3ba(2)
    footer(26) = var3ba(1)
    footer(27) = var3ba(0)
    If var4 < 0 Then
        var4 = var4 * (-1)
        var4ba = ToByteArray(var4)
        var4ba(3) = var4ba(3) + 128
    Else
        var4ba = ToByteArray(var4)
    End If
    footer(28) = var4ba(3)
    footer(29) = var4ba(2)
    footer(30) = var4ba(1)
    footer(31) = var4ba(0)
    If var5 < 0 Then
        var5 = var5 * (-1)
        var5ba = ToByteArray(var5)
        var5ba(3) = var5ba(3) + 128
    Else
        var5ba = ToByteArray(var5)
    End If
    footer(32) = var5ba(3)
    footer(33) = var5ba(2)
    footer(34) = var5ba(1)
    footer(35) = var5ba(0)
    If var6 < 0 Then
        var6 = var6 * (-1)
        var6ba = ToByteArray(var6)
        var6ba(3) = var6ba(3) + 128
    Else
        var6ba = ToByteArray(var6)
    End If
    footer(36) = var6ba(3)
    footer(37) = var6ba(2)
    footer(38) = var6ba(1)
    footer(39) = var6ba(0)
    If var7 < 0 Then
        var7 = var7 * (-1)
        var7ba = ToByteArray(var7)
        var7ba(3) = var7ba(3) + 128
    Else
        var7ba = ToByteArray(var7)
    End If
    footer(40) = var7ba(3)
    footer(41) = var7ba(2)
    footer(42) = var7ba(1)
    footer(43) = var7ba(0)
    If var8 < 0 Then
        var8 = var8 * (-1)
        var8ba = ToByteArray(var8)
        var8ba(3) = var8ba(3) + 128
    Else
        var8ba = ToByteArray(var8)
    End If
    footer(44) = var8ba(3)
    footer(45) = var8ba(2)
    footer(46) = var8ba(1)
    footer(47) = var8ba(0)
    Rem Footer
    footer(51) = footersize
    footer(55) = major
    footer(59) = minor
    footer(60) = magic(0)
    footer(61) = magic(1)
    footer(62) = magic(2)
    footer(63) = magic(3)
    Close
    Open CommonDialog1.filename For Binary Access Read Write As #10
    Dim pos As Long
    If isgbx = False Then pos = LOF(10) + 1 Else pos = LOF(10) - 63
    Seek #10, pos
    Put #10, , footer
    MsgBox "File saved successfully.", vbOKOnly, "GBXBuilder"
    Close
    Exit Sub
ErrHandler:
    MsgBox "Error saving file!", vbOKOnly, "GBXBuilder Error"
    Err.Clear
    Close
    Exit Sub
End Sub

Private Sub menustrip_Click()
On Error GoTo ErrHandler
    With CommonDialog1
        .CancelError = True
        .DialogTitle = "Save raw dump"
        .filename = suggestion
        .Filter = "Raw Game Boy ROMs (*.gb; *.gbc; *.bin)|*.gb;*.gbc;*.bin|"
        .ShowSave
    End With
    FileCopy original, CommonDialog1.filename
    Open original For Binary As #1
    truncatefile CommonDialog1.filename, LOF(1) - 64
    MsgBox "File saved successfully.", vbOKOnly, "GBXBuilder"
ErrHandler:
    Err.Clear
    Close
    Exit Sub
End Sub

Private Sub menuexit_Click()
    Close
    Unload config
    Set config = Nothing
End Sub

Private Sub menuopen_Click()
On Error GoTo ErrHandler
    loadfile
    If gbx1 = False Then
        MsgBox "This GBX file is in a newer format than this version of GBXBuilder can handle.", vbOKOnly, "GBXBuilder Error"
        gbx1 = True
        isgbx = previsgbx
        original = prevoriginal
        suggestion = prevsuggestion
    End If
ErrHandler:
    Err.Clear
    Close
    Exit Sub
End Sub

Private Sub menureset_Click()
    clearmsvs
    menureset.Checked = True
    menureset.Enabled = False
End Sub

Private Sub menuabout_Click()
    about.Show
End Sub

Public Sub loadfile()
    previsgbx = isgbx
    isgbx = False
    If Not arg = "" Then
        If Left$(arg, 1) = """" Then
            arg = Mid$(arg, 2, Len(arg) - 2)
        End If
        If Dir$(arg) = "" Then GoTo InvalidCommand
        Open arg For Binary As #1
        original = arg
        arg = ""
    Else
InvalidCommand:
        With CommonDialog1
            .CancelError = True
            .DialogTitle = "Open Game Boy (Color) ROM"
            .Filter = "Game Boy ROMs (*.gb; *.gbc; *.gbx; *.bin)|*.gb;*.gbc;*.gbx;*.bin|"
            .ShowOpen
        End With
        Open CommonDialog1.filename For Binary As #1
        original = CommonDialog1.filename
    End If
    prevoriginal = original
    prevsuggestion = suggestion
    Dim tempst As String
    tempst = Mid$(original, InStrRev(original, "\") + 1)
    suggestion = Left$(tempst, InStrRev(tempst, ".") - 1)
    Dim buffer As Byte
    Dim counter As Byte
    counter = 3
    Do
        Get #1, LOF(1) - counter, buffer
        If buffer = magic(-1 * counter + 3) Then
            isgbx = True
        Else
            isgbx = False
            GoTo Raw
        End If
        If counter = 0 Then GoTo GbxStart
        counter = counter - 1
    Loop While True
GbxStart:
    Dim halt As Byte
    Dim test As Byte
    counter = 15
    GoTo GbxZeroes
GbxZeroes:
    halt = counter - 3
    Do
        Get #1, LOF(1) - counter, buffer
        If counter = halt Then GoTo GbxZeroesNext
        counter = counter - 1
        If buffer <> 0 Then GoTo WrongVer
    Loop While True
GbxZeroesNext:
    If halt = 12 Then test = footersize
    If halt = 8 Then test = major
    If halt = 4 Then test = minor
    If buffer <> test Then GoTo WrongVer
    If halt = 4 Then GoTo GbxSuccess
    counter = counter - 1
    GoTo GbxZeroes
GbxSuccess:
    romsize = LOF(1) - 64
    assertgbx
    Exit Sub
WrongVer:
    gbx1 = False
    Exit Sub
Raw:
    romsize = LOF(1)
    assertgbx
    Exit Sub
End Sub

Private Sub assertgbx()
On Error GoTo ErrHandler
    If isgbx = True Then
        menustrip.Enabled = True
        Get #1, LOF(1) - 63, mapper(0)
        Get #1, LOF(1) - 62, mapper(1)
        Get #1, LOF(1) - 61, mapper(2)
        Get #1, LOF(1) - 60, mapper(3)
        Dim index As Byte
        index = 255
        If mapper(0) = 82 Then If mapper(1) = 79 Then If mapper(2) = 77 Then If mapper(3) = 0 Then index = 0
        If mapper(0) = 77 Then If mapper(1) = 66 Then If mapper(2) = 67 Then If mapper(3) = 49 Then index = 1
        If mapper(0) = 77 Then If mapper(1) = 66 Then If mapper(2) = 67 Then If mapper(3) = 50 Then index = 2
        If mapper(0) = 77 Then If mapper(1) = 66 Then If mapper(2) = 67 Then If mapper(3) = 51 Then index = 3
        If mapper(0) = 77 Then If mapper(1) = 66 Then If mapper(2) = 67 Then If mapper(3) = 53 Then index = 4
        If mapper(0) = 77 Then If mapper(1) = 66 Then If mapper(2) = 67 Then If mapper(3) = 54 Then index = 5
        If mapper(0) = 77 Then If mapper(1) = 66 Then If mapper(2) = 67 Then If mapper(3) = 55 Then index = 6
        If mapper(0) = 77 Then If mapper(1) = 66 Then If mapper(2) = 49 Then If mapper(3) = 77 Then index = 7
        If mapper(0) = 77 Then If mapper(1) = 77 Then If mapper(2) = 77 Then If mapper(3) = 49 Then index = 8
        If mapper(0) = 67 Then If mapper(1) = 65 Then If mapper(2) = 77 Then If mapper(3) = 82 Then index = 9
        If mapper(0) = 72 Then If mapper(1) = 85 Then If mapper(2) = 67 Then If mapper(3) = 49 Then index = 10
        If mapper(0) = 72 Then If mapper(1) = 85 Then If mapper(2) = 67 Then If mapper(3) = 51 Then index = 11
        If mapper(0) = 84 Then If mapper(1) = 65 Then If mapper(2) = 77 Then If mapper(3) = 53 Then index = 12
        If mapper(0) = 77 Then If mapper(1) = 49 Then If mapper(2) = 54 Then If mapper(3) = 49 Then index = 13
        If mapper(0) = 66 Then If mapper(1) = 66 Then If mapper(2) = 68 Then If mapper(3) = 0 Then index = 14
        If mapper(0) = 72 Then If mapper(1) = 73 Then If mapper(2) = 84 Then If mapper(3) = 75 Then index = 15
        If mapper(0) = 83 Then If mapper(1) = 78 Then If mapper(2) = 84 Then If mapper(3) = 88 Then index = 16
        If mapper(0) = 78 Then If mapper(1) = 84 Then If mapper(2) = 79 Then If mapper(3) = 49 Then index = 17
        If mapper(0) = 78 Then If mapper(1) = 84 Then If mapper(2) = 79 Then If mapper(3) = 50 Then index = 18
        If mapper(0) = 78 Then If mapper(1) = 84 Then If mapper(2) = 78 Then If mapper(3) = 0 Then index = 19
        If mapper(0) = 76 Then If mapper(1) = 73 Then If mapper(2) = 67 Then If mapper(3) = 72 Then index = 20
        If mapper(0) = 76 Then If mapper(1) = 66 Then If mapper(2) = 77 Then If mapper(3) = 67 Then index = 21
        If mapper(0) = 76 Then If mapper(1) = 73 Then If mapper(2) = 66 Then If mapper(3) = 65 Then index = 22
        If mapper(0) = 80 Then If mapper(1) = 75 Then If mapper(2) = 74 Then If mapper(3) = 68 Then index = 23
        If mapper(0) = 87 Then If mapper(1) = 73 Then If mapper(2) = 83 Then If mapper(3) = 68 Then index = 24
        If mapper(0) = 83 Then If mapper(1) = 65 Then If mapper(2) = 77 Then If mapper(3) = 49 Then index = 25
        If mapper(0) = 83 Then If mapper(1) = 65 Then If mapper(2) = 77 Then If mapper(3) = 50 Then index = 26
        If mapper(0) = 82 Then If mapper(1) = 79 Then If mapper(2) = 67 Then If mapper(3) = 75 Then index = 27
        If mapper(0) = 78 Then If mapper(1) = 71 Then If mapper(2) = 72 Then If mapper(3) = 75 Then index = 28
        If mapper(0) = 86 Then If mapper(1) = 70 Then If mapper(2) = 48 Then If mapper(3) = 49 Then index = 29
        If mapper(0) = 71 Then If mapper(1) = 66 Then If mapper(2) = 56 Then If mapper(3) = 49 Then index = 30
        If mapper(0) = 84 Then If mapper(1) = 80 Then If mapper(3) = 80 Then If mapper(3) = 49 Then index = 31
        If index = 255 Then
            mappertype.Text = ToAscii(mapper)
        Else
            mappertype.ListIndex = index
        End If
        Dim counter As Byte
        Dim buffer As Byte
        counter = 47
        index = 16
        Do
            Get #1, LOF(1) - counter, buffer
            footer(index) = buffer
            counter = counter - 1
            index = index + 1
        Loop While index < 48
        var1ba(3) = footer(16)
        var1ba(2) = footer(17)
        var1ba(1) = footer(18)
        var1ba(0) = footer(19)
        If var1ba(3) > 127 Then
            var1ba(3) = var1ba(3) - 128
            var1 = ToLong(var1ba)
            var1 = var1 * (-1)
            var1ba(3) = var1ba(3) + 128
        Else
            var1 = ToLong(var1ba)
        End If
        var2ba(3) = footer(20)
        var2ba(2) = footer(21)
        var2ba(1) = footer(22)
        var2ba(0) = footer(23)
        If var2ba(3) > 127 Then
            var2ba(3) = var2ba(3) - 128
            var2 = ToLong(var2ba)
            var2 = var2 * (-1)
            var2ba(3) = var2ba(3) + 128
        Else
            var2 = ToLong(var2ba)
        End If
        var3ba(3) = footer(24)
        var3ba(2) = footer(25)
        var3ba(1) = footer(26)
        var3ba(0) = footer(27)
        If var3ba(3) > 127 Then
            var3ba(3) = var3ba(3) - 128
            var3 = ToLong(var3ba)
            var3 = var3 * (-1)
            var3ba(3) = var3ba(3) + 128
        Else
            var3 = ToLong(var3ba)
        End If
        var4ba(3) = footer(28)
        var4ba(2) = footer(29)
        var4ba(1) = footer(30)
        var4ba(0) = footer(31)
        If var4ba(3) > 127 Then
            var4ba(3) = var4ba(3) - 128
            var4 = ToLong(var4ba)
            var4 = var4 * (-1)
            var4ba(3) = var4ba(3) + 128
        Else
            var4 = ToLong(var4ba)
        End If
        var5ba(3) = footer(32)
        var5ba(2) = footer(33)
        var5ba(1) = footer(34)
        var5ba(0) = footer(35)
        If var5ba(3) > 127 Then
            var5ba(3) = var5ba(3) - 128
            var5 = ToLong(var5ba)
            var5 = var5 * (-1)
            var5ba(3) = var5ba(3) + 128
        Else
            var5 = ToLong(var5ba)
        End If
        var6ba(3) = footer(36)
        var6ba(2) = footer(37)
        var6ba(1) = footer(38)
        var6ba(0) = footer(39)
        If var6ba(3) > 127 Then
            var6ba(3) = var6ba(3) - 128
            var6 = ToLong(var6ba)
            var6 = var6 * (-1)
            var6ba(3) = var6ba(3) + 128
        Else
            var6 = ToLong(var6ba)
        End If
        var7ba(3) = footer(40)
        var7ba(2) = footer(41)
        var7ba(1) = footer(42)
        var7ba(0) = footer(43)
        If var7ba(3) > 127 Then
            var7ba(3) = var7ba(3) - 128
            var7 = ToLong(var7ba)
            var7 = var7 * (-1)
            var7ba(3) = var7ba(3) + 128
        Else
            var7 = ToLong(var7ba)
        End If
        var8ba(3) = footer(44)
        var8ba(2) = footer(45)
        var8ba(1) = footer(46)
        var8ba(0) = footer(47)
        If var8ba(3) > 127 Then
            var8ba(3) = var8ba(3) - 128
            var8 = ToLong(var8ba)
            var8 = var8 * (-1)
            var8ba(3) = var8ba(3) + 128
        Else
            var8 = ToLong(var8ba)
        End If
        refreshmsvs
        verifymsvs
        Get #1, LOF(1) - 59, battery
        isram.Value = battery
        Get #1, LOF(1) - 58, rumble
        isrumble.Value = rumble
        Get #1, LOF(1) - 57, timer
        istimer.Value = timer
        If battery <> 0 Then
            Get #1, LOF(1) - 51, ramsizeba(3)
            Get #1, LOF(1) - 50, ramsizeba(2)
            Get #1, LOF(1) - 49, ramsizeba(1)
            Get #1, LOF(1) - 48, ramsizeba(0)
            ramsize = ToLong(ramsizeba)
            ramtext.Text = ramsize
        Else
            ramsize = 0
            ramtext.Text = 0
        End If
    Else
        menustrip.Enabled = False
    End If
    Exit Sub
ErrHandler:
    MsgBox "Fatal error: The GBX file seems to be corrupted or badly formatted.", vbOKOnly, "GBXBuilder Fatal Error"
    Unload config
    Set config = Nothing
    Exit Sub
End Sub

Private Sub clearmsvs()
    var1 = 0
    var1ba(0) = 0
    var1ba(1) = 0
    var1ba(2) = 0
    var1ba(3) = 0
    var2 = 0
    var2ba(0) = 0
    var2ba(1) = 0
    var2ba(2) = 0
    var2ba(3) = 0
    var3 = 0
    var3ba(0) = 0
    var3ba(1) = 0
    var3ba(2) = 0
    var3ba(3) = 0
    var4 = 0
    var4ba(0) = 0
    var4ba(1) = 0
    var4ba(2) = 0
    var4ba(3) = 0
    var5 = 0
    var5ba(0) = 0
    var5ba(1) = 0
    var5ba(2) = 0
    var5ba(3) = 0
    var6 = 0
    var6ba(0) = 0
    var6ba(1) = 0
    var6ba(2) = 0
    var6ba(3) = 0
    var7 = 0
    var7ba(0) = 0
    var7ba(1) = 0
    var7ba(2) = 0
    var7ba(3) = 0
    var8 = 0
    var8ba(0) = 0
    var8ba(1) = 0
    var8ba(2) = 0
    var8ba(3) = 0
    refreshmsvs
End Sub

Private Sub refreshmsvs()
    cvar1 = var1
    cvar2 = var2
    cvar3 = var3
    cvar4 = var4
    cvar5 = var5
    cvar6 = var6
    cvar7 = var7
    cvar8 = var8
End Sub

Public Function ToByteArray(ByVal lng As Long) As Byte()
    Dim bytearray(3) As Byte
    CopyMemory bytearray(0), ByVal VarPtr(lng), Len(lng)
    ToByteArray = bytearray
End Function

Public Function ToLong(vIn() As Byte) As Long
    Dim i As Long
    Dim vOut As Long
    vOut = 0
    For i = 0 To 3
        vOut = vOut Or (CLng(vIn(i)) * 256 ^ i)
    Next i
    ToLong = vOut
End Function

Public Function FromAscii(ByVal str As String, ByVal counter As Byte) As Byte
On Error GoTo ErrHandler
    Dim result As Byte
    Dim char As String
    result = 0
    char = Mid(str, counter, 1)
    If char = "!" Then result = 33
    If char = "+" Then result = 43
    If char = "0" Then result = 48
    If char = "1" Then result = 49
    If char = "2" Then result = 50
    If char = "3" Then result = 51
    If char = "4" Then result = 52
    If char = "5" Then result = 53
    If char = "6" Then result = 54
    If char = "7" Then result = 55
    If char = "8" Then result = 56
    If char = "9" Then result = 57
    If char = "?" Then result = 63
    If char = "A" Then result = 65
    If char = "B" Then result = 66
    If char = "C" Then result = 67
    If char = "D" Then result = 68
    If char = "E" Then result = 69
    If char = "F" Then result = 70
    If char = "G" Then result = 71
    If char = "H" Then result = 72
    If char = "I" Then result = 73
    If char = "J" Then result = 74
    If char = "K" Then result = 75
    If char = "L" Then result = 76
    If char = "M" Then result = 77
    If char = "N" Then result = 78
    If char = "O" Then result = 79
    If char = "P" Then result = 80
    If char = "Q" Then result = 81
    If char = "R" Then result = 82
    If char = "S" Then result = 83
    If char = "T" Then result = 84
    If char = "U" Then result = 85
    If char = "V" Then result = 86
    If char = "W" Then result = 87
    If char = "X" Then result = 88
    If char = "Y" Then result = 89
    If char = "Z" Then result = 90
    If char = "a" Then result = 65
    If char = "b" Then result = 66
    If char = "c" Then result = 67
    If char = "d" Then result = 68
    If char = "e" Then result = 69
    If char = "f" Then result = 70
    If char = "g" Then result = 71
    If char = "h" Then result = 72
    If char = "i" Then result = 73
    If char = "j" Then result = 74
    If char = "k" Then result = 75
    If char = "l" Then result = 76
    If char = "m" Then result = 77
    If char = "n" Then result = 78
    If char = "o" Then result = 79
    If char = "p" Then result = 80
    If char = "q" Then result = 81
    If char = "r" Then result = 82
    If char = "s" Then result = 83
    If char = "t" Then result = 84
    If char = "u" Then result = 85
    If char = "v" Then result = 86
    If char = "w" Then result = 87
    If char = "x" Then result = 88
    If char = "y" Then result = 89
    If char = "z" Then result = 90
    If result = 0 Then If char <> "" Then Err.Raise -1
    FromAscii = result
    Exit Function
ErrHandler:
    MsgBox "Illegal character in mapper identifier.", vbOKOnly, "GBXBuilder Error"
    Err.Raise -1
    Exit Function
End Function

Public Function ToAscii(bytearray() As Byte) As String
    Dim counter As Byte
    Dim str As String
    counter = 0
    str = ""
    If bytearray(0) = 0 Then Err.Raise -1
    If bytearray(1) = 0 Then If bytearray(2) <> 0 Then Err.Raise -1
    Do
        If bytearray(counter) = 0 Then If bytearray(3) <> 0 Then Err.Raise -1
        If bytearray(counter) = 33 Then str = str & "!"
        If bytearray(counter) = 43 Then str = str & "+"
        If bytearray(counter) = 48 Then str = str & "0"
        If bytearray(counter) = 49 Then str = str & "1"
        If bytearray(counter) = 50 Then str = str & "2"
        If bytearray(counter) = 51 Then str = str & "3"
        If bytearray(counter) = 52 Then str = str & "4"
        If bytearray(counter) = 53 Then str = str & "5"
        If bytearray(counter) = 54 Then str = str & "6"
        If bytearray(counter) = 55 Then str = str & "7"
        If bytearray(counter) = 56 Then str = str & "8"
        If bytearray(counter) = 57 Then str = str & "9"
        If bytearray(counter) = 63 Then str = str & "?"
        If bytearray(counter) = 65 Then str = str & "A"
        If bytearray(counter) = 66 Then str = str & "B"
        If bytearray(counter) = 67 Then str = str & "C"
        If bytearray(counter) = 68 Then str = str & "D"
        If bytearray(counter) = 69 Then str = str & "E"
        If bytearray(counter) = 70 Then str = str & "F"
        If bytearray(counter) = 71 Then str = str & "G"
        If bytearray(counter) = 72 Then str = str & "H"
        If bytearray(counter) = 73 Then str = str & "I"
        If bytearray(counter) = 74 Then str = str & "J"
        If bytearray(counter) = 75 Then str = str & "K"
        If bytearray(counter) = 76 Then str = str & "L"
        If bytearray(counter) = 77 Then str = str & "M"
        If bytearray(counter) = 78 Then str = str & "N"
        If bytearray(counter) = 79 Then str = str & "O"
        If bytearray(counter) = 80 Then str = str & "P"
        If bytearray(counter) = 81 Then str = str & "Q"
        If bytearray(counter) = 82 Then str = str & "R"
        If bytearray(counter) = 83 Then str = str & "S"
        If bytearray(counter) = 84 Then str = str & "T"
        If bytearray(counter) = 85 Then str = str & "U"
        If bytearray(counter) = 86 Then str = str & "V"
        If bytearray(counter) = 87 Then str = str & "W"
        If bytearray(counter) = 88 Then str = str & "X"
        If bytearray(counter) = 89 Then str = str & "Y"
        If bytearray(counter) = 90 Then str = str & "Z"
        counter = counter + 1
    Loop While counter < 4
    ToAscii = str
    Exit Function
End Function


