VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ucBrowse - v1.0.42 "
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkHideSelection 
      Caption         =   "HideSelection"
      Height          =   375
      Left            =   4200
      TabIndex        =   25
      Top             =   2280
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkHasButtons 
      Caption         =   "HasButtons"
      Height          =   375
      Left            =   4200
      TabIndex        =   24
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Apply"
      Height          =   375
      Left            =   6000
      TabIndex        =   23
      Top             =   3720
      Width           =   825
   End
   Begin VB.CheckBox chkCheckBoxes 
      Caption         =   "CheckBoxes (Design Only)"
      Height          =   375
      Left            =   4200
      TabIndex        =   22
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CheckBox chkFullRowSelect 
      Caption         =   "FullRowSelect"
      Height          =   375
      Left            =   4200
      TabIndex        =   21
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CheckBox chkHasLines 
      Caption         =   "HasLines"
      Height          =   375
      Left            =   5640
      TabIndex        =   20
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkHotTracking 
      Caption         =   "HotTracking"
      Height          =   375
      Left            =   5640
      TabIndex        =   19
      Top             =   2280
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.ComboBox cmbRoot 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset"
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   5160
      Width           =   825
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   5100
      TabIndex        =   10
      Top             =   5160
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "More..."
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   5160
      Width           =   825
   End
   Begin Project1.ucPickBox pbPath 
      Height          =   315
      Left            =   4200
      TabIndex        =   7
      Top             =   3000
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      DialogType      =   1
   End
   Begin VB.ComboBox cmbTheme 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Frame fmAppearance 
      Caption         =   "Appearance:"
      Height          =   735
      Left            =   4200
      TabIndex        =   1
      Top             =   390
      Width           =   2655
      Begin VB.OptionButton Option1 
         Caption         =   "3D"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   3
         Top             =   320
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Flat"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   320
         Width           =   975
      End
   End
   Begin Project1.ucBrowse ucBrowse1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   8916
      Appearance      =   0
      CheckBoxes      =   -1  'True
      FolderFlags     =   16384
      HasLines        =   -1  'True
      HideSelection   =   0   'False
   End
   Begin VB.Label Label3 
      Caption         =   "Root:"
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "TreeView Status:"
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Header 
      BackStyle       =   0  'Transparent
      Caption         =   "ucBrowse"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   60
      Width           =   1575
   End
   Begin VB.Label Header 
      BackStyle       =   0  'Transparent
      Caption         =   "ucBrowse"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B0B0B0&
      Height          =   375
      Index           =   2
      Left            =   150
      TabIndex        =   14
      Top             =   90
      Width           =   1575
   End
   Begin VB.Label lblActivePath 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   5880
      Width           =   6735
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Returned Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Header 
      Caption         =   " - System Folder Browser UserControl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   11
      Top             =   180
      Width           =   5295
   End
   Begin VB.Label lblPath 
      Caption         =   "Path:"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblTheme 
      Caption         =   "Theme:"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   4200
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bLoading As Boolean

Private Sub chkCheckBoxes_Click()
    With Me
        .ucBrowse1.CheckBoxes = (.chkCheckBoxes.Value = vbChecked)
    End With
End Sub

Private Sub chkFullRowSelect_Click()
    With Me
        .ucBrowse1.FullRowSelect = (.chkFullRowSelect.Value = vbChecked)
    End With
End Sub

Private Sub chkHasButtons_Click()
    With Me
        .ucBrowse1.HasButtons = (.chkHasButtons.Value = vbChecked)
    End With
End Sub

Private Sub chkHasLines_Click()
    With Me
        .ucBrowse1.HasLines = (.chkHasLines.Value = vbChecked)
    End With
End Sub

Private Sub chkHideSelection_Click()
    With Me
        .ucBrowse1.HideSelection = (.chkHideSelection.Value = vbChecked)
    End With
End Sub

Private Sub chkHotTracking_Click()
    With Me
        .ucBrowse1.HotTracking = (.chkHotTracking.Value = vbChecked)
    End With
End Sub

Private Sub cmbTheme_Click()
    With Me
        .ucBrowse1.Theme = .cmbTheme.ListIndex
        .Option1(.ucBrowse1.Appearance).Value = True
        .pbPath.Theme = .cmbTheme.ListIndex
    End With
End Sub

Private Sub Command1_Click()
    frmMultiples.Show
End Sub

Private Sub Command2_Click()
    With Me
        .ucBrowse1.CloseUp
    End With
End Sub

Private Sub Command3_Click()
    With Me
        '.ucBrowse1.Reset
    End With
End Sub

Private Sub Command4_Click()
    If Not bLoading Then
        '   Note: Do not attempt to Set the Root of the ucBrowse mutiple
        '         times in succession without first calling the CloseUp method.
        '         This can cause the control to hang as it is being subclassed
        '         repeatedly without first destroying the old window.....if this
        '         happens the BFF window is not released and the control acts
        '         as through it is hung.....at which point you must kill the
        '         process from the ProgramManager ;-(
        '         You have been notified....proceed at your own risk!!
        
        '   First CloseUp the Dialog
        Command2_Click
        '   NOw Set the New Root.....
        Select Case Me.cmbRoot.ListIndex
            Case 0: Me.ucBrowse1.Root = AdminTools               '0
            Case 1: Me.ucBrowse1.Root = AltStartUp               '1
            Case 2: Me.ucBrowse1.Root = ApplicationData          '2
            Case 3: Me.ucBrowse1.Root = CDBurnArea               '3
            Case 4: Me.ucBrowse1.Root = CommonAdminTools         '4
            Case 5: Me.ucBrowse1.Root = CommonAltStartUp         '5
            Case 6: Me.ucBrowse1.Root = CommonAppData            '6
            Case 7: Me.ucBrowse1.Root = CommonDesktopDirectory   '7
            Case 8: Me.ucBrowse1.Root = CommonFavorites          '8
            Case 9: Me.ucBrowse1.Root = CommonMyDocuments        '9
            Case 10: Me.ucBrowse1.Root = CommonMyMusic            '10
            Case 11: Me.ucBrowse1.Root = CommonMyPictures         '11
            Case 12: Me.ucBrowse1.Root = CommonMyVideo            '12
            Case 13: Me.ucBrowse1.Root = CommonProgramFiles       '13
            Case 14: Me.ucBrowse1.Root = CommonPrograms           '14
            Case 15: Me.ucBrowse1.Root = CommonStartMenu          '15
            Case 16: Me.ucBrowse1.Root = CommonStartUp            '16
            Case 17: Me.ucBrowse1.Root = CommonTemplates          '17
            Case 18: Me.ucBrowse1.Root = ComputersNearMe          '18
            Case 19: Me.ucBrowse1.Root = Connections              '19
            Case 20: Me.ucBrowse1.Root = ControlPanel             '20
            Case 21: Me.ucBrowse1.Root = DeskTop                  '21
            Case 22: Me.ucBrowse1.Root = DesktopDirectory         '22
            Case 23: Me.ucBrowse1.Root = Favorites                '23
            Case 24: Me.ucBrowse1.Root = Fonts                    '24
            Case 25: Me.ucBrowse1.Root = Internet                 '25
            Case 26: Me.ucBrowse1.Root = InternetCache            '26
            Case 27: Me.ucBrowse1.Root = InternetCookies          '27
            Case 28: Me.ucBrowse1.Root = InternetHistory          '28
            Case 29: Me.ucBrowse1.Root = LocalApplicationData     '29
            Case 30: Me.ucBrowse1.Root = LocalizedResources       '30
            Case 31: Me.ucBrowse1.Root = MyComputer               '31
            Case 32: Me.ucBrowse1.Root = MyDocuments              '32
            Case 33: Me.ucBrowse1.Root = MyDocumentsFolder        '33
            Case 34: Me.ucBrowse1.Root = MyMusic                  '34
            Case 35: Me.ucBrowse1.Root = MyNetworkPlaces          '35
            Case 36: Me.ucBrowse1.Root = MyPictures               '36
            Case 37: Me.ucBrowse1.Root = MyVideo                  '37
            Case 38: Me.ucBrowse1.Root = NetworkNeighborhood      '38
            Case 39: Me.ucBrowse1.Root = Printers                 '39
            Case 40: Me.ucBrowse1.Root = PrintHood                '40
            Case 41: Me.ucBrowse1.Root = Profile                  '41
            Case 42: Me.ucBrowse1.Root = Programs                 '42
            Case 43: Me.ucBrowse1.Root = ProgramsFiles            '43
            Case 44: Me.ucBrowse1.Root = Recent                   '44
            Case 45: Me.ucBrowse1.Root = RecycleBin               '45
            Case 46: Me.ucBrowse1.Root = SendTo                   '46
            Case 47: Me.ucBrowse1.Root = StartMenu                '47
            Case 48: Me.ucBrowse1.Root = StartUp                  '48
            Case 49: Me.ucBrowse1.Root = System                   '49
            Case 50: Me.ucBrowse1.Root = SystemResources          '50
            Case 51: Me.ucBrowse1.Root = Templates                '51
            Case 52: Me.ucBrowse1.Root = Windows                  '52
        End Select
        '   Process the events
        DoEvents
    End If
End Sub

Private Sub Form_Load()
    With Me
        .Caption = "ucBrowse - v" & .ucBrowse1.Version
        bLoading = True
        With .cmbTheme
            .AddItem "Auto"
            .AddItem "Classic"
            .AddItem "Blue"
            .AddItem "HomeStead"
            .AddItem "Metallic"
            .AddItem "None"
            .ListIndex = 0
        End With
        With .cmbRoot
            .AddItem "AdminTools"               '0
            .AddItem "AltStartUp"               '1
            .AddItem "ApplicationData"          '2
            .AddItem "CDBurnArea"               '3
            .AddItem "CommonAdminTools"         '4
            .AddItem "CommonAltStartUp"         '5
            .AddItem "CommonAppData"            '6
            .AddItem "CommonDesktopDirectory"   '7
            .AddItem "CommonFavorites"          '8
            .AddItem "CommonMyDocuments"        '9
            .AddItem "CommonMyMusic"            '10
            .AddItem "CommonMyPictures"         '11
            .AddItem "CommonMyVideo"            '12
            .AddItem "CommonProgramFiles"       '13
            .AddItem "CommonPrograms"           '14
            .AddItem "CommonStartMenu"          '15
            .AddItem "CommonStartUp"            '16
            .AddItem "CommonTemplates"          '17
            .AddItem "ComputersNearMe"          '18
            .AddItem "Connections"              '19
            .AddItem "Controls"                 '20
            .AddItem "DeskTop"                  '21
            .AddItem "DesktopDirectory"         '22
            .AddItem "Favorites"                '23
            .AddItem "Fonts"                    '24
            .AddItem "Internet"                 '25
            .AddItem "InternetCache"            '26
            .AddItem "InternetCookies"          '27
            .AddItem "InternetHistory"          '28
            .AddItem "LocalApplicationData"     '29
            .AddItem "LocalizedResources"       '30
            .AddItem "MyComputer"               '31
            .AddItem "MyDocuments"              '32
            .AddItem "MyDocumentsFolder"        '33
            .AddItem "MyMusic"                  '34
            .AddItem "MyNetworkPlaces"          '35
            .AddItem "MyPictures"               '36
            .AddItem "MyVideo"                  '37
            .AddItem "NetworkNeighborhood"      '38
            .AddItem "Printers"                 '39
            .AddItem "PrintHood"                '40
            .AddItem "Profile"                  '41
            .AddItem "Programs"                 '42
            .AddItem "ProgramsFiles"            '43
            .AddItem "Recent"                   '44
            .AddItem "RecycleBin"               '45
            .AddItem "SendTo"                   '46
            .AddItem "StartMenu"                '47
            .AddItem "StartUp"                  '48
            .AddItem "System"                   '49
            .AddItem "SystemResources"          '50
            .AddItem "Templates"                '51
            .AddItem "Windows"                  '52
            .ListIndex = 21
        End With
        .Option1(.ucBrowse1.Appearance).Value = True
        .pbPath.DialogType = ucFolder
        .chkCheckBoxes.Enabled = False
        .ucBrowse1.CheckBoxes = (.chkCheckBoxes.Value = vbChecked)
        bLoading = False
    End With
End Sub

Private Sub Option1_Click(index As Integer)
    With Me
        .ucBrowse1.Appearance = index
    End With
End Sub

Private Sub pbPath_Click()
    With Me
        .ucBrowse1.Path = .pbPath.Path
    End With
End Sub

Private Sub ucBrowse1_MouseEnter()
    Debug.Print "MouseEnter:"
End Sub

Private Sub ucBrowse1_MouseLeave()
    Debug.Print "MouseLeave:"
End Sub

Private Sub ucBrowse1_PathChange(ByVal sPath As String)
    Me.lblActivePath.Caption = sPath
End Sub

Private Sub ucBrowse1_Status(ByVal sStatus As String)
    Debug.Print sStatus
End Sub

