VERSION 5.00
Begin VB.UserControl ucBrowse 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   Picture         =   "ucBrowse.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ucBrowse.ctx":1A632
End
Attribute VB_Name = "ucBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+  File Description:
'       ucBrowse - SelfSubclassed System Browse For Folder UserControl
'
'   Product Name:
'       ucBrowse.ctl
'
'   Compatability:
'       Windows: 98(?), ME(?), NT, 2000, XP
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'   Based on the following On-Line Articles
'       (Paul Caton - Self-Subclassser)
'           http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'       (MrBoBo - System Treeview Thievery)
'           http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=40007&lngWId=1
'       (Randy Birch - IsWinXP)
'           http://vbnet.mvps.org/code/system/getversionex.htm
'       (Dieter Otter - GetCurrentThemeName)
'           http://www.vbarchiv.net/archiv/tipp_805.html
'       (Randy Birch / Brad Martinez - TreeView CheckBoxes)
'           http://vbnet.mvps.org/index.html?code/comctl/tvcheckbox.htm
'       (Randy Birch - TreeView Special Effects)
'           http://vbnet.mvps.org/code/comctl/tveffects.htm
'       (Randy Birch - Special Folders)
'           http://vbnet.mvps.org/index.html?code/callback/browsecallbackcdrom.htm
'
'   Legal Copyright & Trademarks:
'       Copyright © 2006-2007, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2006-2007, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Paul R. Territo, Ph.D shall not be liable
'       for any incidental or consequential damages suffered by any use of
'       this  software. This software is owned by Paul R. Territo, Ph.D and is
'       free for use in accordance with the terms of the License Agreement
'       in the accompanying the documentation.
'
'   Contact Information:
'       For Technical Assistance:
'       Email: pwterrito@insightbb.com
'
'-  Modification(s) History:
'       14Jul06 - Initial TestHarness and UserControl finished
'       12Oct07 - Added additonal cleanup routines to allow for dynamic
'                 root changing
'               - Added Reset routine to reset the control
'               - Added additional error checking for shutdown to prevent reloads
'                 which cause the control to appear to hang (actually does not
'                 hang, but the focus has been set back to the hidden BFF window)
'               - Added CoTaskMemFree API to allow for freeing of BFF Pointer
'       16Oct07 - Added HasButtons property and associated window style properties
'               - Added HideSelection property to fix focus managment issues pointed out by Carles P.V.
'               - Added IsWin2K to determine if unicode is available
'               - Added SHBrowseForFolderW and SHGetPathFromIDListW for Unicode Support
'       28Oct07 - Added IsFolder method to prevent incorrect qualifying the files
'                 passed back by QualifyPath method....thanks to Ruturaj for catching this!!
'               - Added IsFile method (Logical opposite of IsFolder).
'
'
'   Force Declarations
Option Explicit

'   Build Date & Time: 10/28/2007 8:05:13 PM
Const Major As Long = 1
Const Minor As Long = 0
Const Revision As Long = 50
Const DateTime As String = "10/28/2007 8:05:13 PM"

Private Type OSVERSIONINFO
    OSVSize         As Long             'size, in bytes, of this data structure
    dwVerMajor      As Long             'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4.
    dwVerMinor      As Long             'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0.
    dwBuildNumber   As Long             'NT: build number of the OS
                                        'Win9x: build number of the OS in low-order word.
                                        '       High-order word contains major & minor ver nos.
    PlatformID      As Long             'Identifies the operating system platform.
    szCSDVersion    As String * 128     'NT: string, such as "Service Pack 3"
                                        'Win9x: string providing arbitrary additional information
End Type

Private Const VER_PLATFORM_WIN32_NT = 2

Private Const BIF_STATUSTEXT = &H4&
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSELECTION = (WM_USER + 102)
Private Const WM_MOVE = &H3
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As String, ByVal dwMaxNameChars As Integer, ByVal pszColorBuff As String, ByVal cchMaxColorChars As Integer, ByVal pszSizeBuff As String, ByVal cchMaxSizeChars As Integer) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function lStrCat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHBrowseForFolderW Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetPathFromIDListW Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Const GW_NEXT = 2
Private Const GW_CHILD = 5
Private Const WM_CLOSE = &H10

'   Window Constants
Private Const GWL_STYLE             As Long = (-16)
Private Const GWL_EXSTYLE           As Long = (-20)
Private Const WH_CALLWNDPROC        As Long = 4
Private Const WS_BORDER             As Long = &H800000
Private Const WS_EX_CLIENTEDGE      As Long = &H200
Private Const WS_EX_STATICEDGE      As Long = &H20000
Private Const SWP_NOMOVE            As Long = &H2
Private Const SWP_NOSIZE            As Long = &H1
Private Const SWP_FRAMECHANGED      As Long = &H20
Private Const SWP_NOACTIVATE        As Long = &H10
Private Const SWP_NOZORDER          As Long = &H4
Private Const SWP_DRAWFRAME         As Long = SWP_FRAMECHANGED
Private Const SWP_FLAGS             As Long = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

'   Standard TreeView Message Bits
'(http://windowssdk.msdn.microsoft.com/en-us/library/ms650019.aspx)
Private Const TV_FIRST              As Long = &H1100
Private Const TVM_GETNEXTITEM       As Long = (TV_FIRST + 10)
Private Const TVM_GETITEM           As Long = (TV_FIRST + 12)
Private Const TVM_SETITEM           As Long = (TV_FIRST + 13)
Private Const TVM_SETBKCOLOR        As Long = (TV_FIRST + 29)
Private Const TVM_SETTEXTCOLOR      As Long = (TV_FIRST + 30)
Private Const TVM_GETBKCOLOR        As Long = (TV_FIRST + 31)
Private Const TVM_GETTEXTCOLOR      As Long = (TV_FIRST + 32)

Private Const TVS_CHECKBOXES        As Long = &H100
Private Const TVS_DISABLEDRAGDROP   As Long = &H10
Private Const TVS_EDITLABELS        As Long = &H8
Private Const TVS_FULLROWSELECT     As Long = &H1000
Private Const TVS_HASBUTTONS        As Long = &H1
Private Const TVS_HASLINES          As Long = &H2
Private Const TVS_INFOTIP           As Long = &H800
Private Const TVS_LINESATROOT       As Long = &H4
Private Const TVS_NOHSCROLL         As Long = &H8000
Private Const TVS_NONEVENHEIGHT     As Long = &H4000
Private Const TVS_NOSCROLL          As Long = &H2000
Private Const TVS_NOTOOLTIPS        As Long = &H80
Private Const TVS_SHOWSELALWAYS     As Long = &H20
Private Const TVS_SINGLEEXPAND      As Long = &H400
Private Const TVS_TRACKSELECT       As Long = &H200


'   Special Folder Flags
Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_INTERNET = &H1
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_CONTROLS = &H3
Private Const CSIDL_PRINTERS = &H4
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_SENDTO = &H9
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_MYDOCUMENTS = &HC
Private Const CSIDL_MYMUSIC = &HD
Private Const CSIDL_MYVIDEO = &HE
Private Const CSIDL_DESKTOPDIRECTORY = &H10
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_NETHOOD = &H13
Private Const CSIDL_FONTS = &H14
Private Const CSIDL_TEMPLATES = &H15
Private Const CSIDL_COMMON_STARTMENU = &H16
Private Const CSIDL_COMMON_PROGRAMS = &H17
Private Const CSIDL_COMMON_STARTUP = &H18
Private Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Private Const CSIDL_APPDATA = &H1A
Private Const CSIDL_PRINTHOOD = &H1B
Private Const CSIDL_LOCAL_APPDATA = &H1C
Private Const CSIDL_ALTSTARTUP = &H1D
Private Const CSIDL_COMMON_ALTSTARTUP = &H1E
Private Const CSIDL_COMMON_FAVORITES = &H1F
Private Const CSIDL_INTERNET_CACHE = &H20
Private Const CSIDL_COOKIES = &H21
Private Const CSIDL_HISTORY = &H22
Private Const CSIDL_COMMON_APPDATA = &H23
Private Const CSIDL_WINDOWS = &H24
Private Const CSIDL_SYSTEM = &H25
Private Const CSIDL_PROGRAM_FILES = &H26
Private Const CSIDL_MYPICTURES = &H27
Private Const CSIDL_PROFILE = &H28
Private Const CSIDL_SYSTEMX86 = &H29
Private Const CSIDL_PROGRAM_FILESX86 = &H2A
Private Const CSIDL_PROGRAM_FILES_COMMON = &H2B
Private Const CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
Private Const CSIDL_COMMON_TEMPLATES = &H2D
Private Const CSIDL_COMMON_DOCUMENTS = &H2E
Private Const CSIDL_COMMON_ADMINTOOLS = &H2F
Private Const CSIDL_ADMINTOOLS = &H30
Private Const CSIDL_CONNECTIONS = &H31
Private Const CSIDL_COMMON_MUSIC = &H35
Private Const CSIDL_COMMON_PICTURES = &H36
Private Const CSIDL_COMMON_VIDEO = &H37
Private Const CSIDL_RESOURCES = &H38
Private Const CSIDL_RESOURCES_LOCALIZED = &H39
Private Const CSIDL_COMMON_OEM_LINKS = &H3A
Private Const CSIDL_CDBURN_AREA = &H3B
Private Const CSIDL_COMPUTERSNEARME = &H3D
Private Const CSIDL_FLAG_PER_USER_INIT = &H800
Private Const CSIDL_FLAG_NO_ALIAS = &H1000
Private Const CSIDL_FLAG_DONT_VERIFY = &H4000
Private Const CSIDL_FLAG_CREATE = &H8000
Private Const CSIDL_FLAG_MASK = &HFF00

Private Type BROWSEINFO
    hwndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As ubFolderDialogFlags
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Public Enum ubFolderDialogFlags
    ReturnOnlyFSDirs = &H1
    DontGoBelowDomain = &H2
    StatusText = &H4
    ReturnFSAncestors = &H8
    EditBox = &H10
    Validate = &H20
    NewDialogStyle = &H40
    UseNewUI = (NewDialogStyle Or EditBox)
    BrowseIncludeURLs = &H80
    UAHInt = &H100
    NoneWFolderButton = &H200
    NoTranslateTargets = &H400
    BrowseForComputer = &H1000
    BrowseForPrinter = &H2000
    BrowseIncludeFiles = &H4000
    Shareable = &H8000
    ShowFolder_Default = ReturnOnlyFSDirs Or StatusText Or BrowseForComputer
End Enum
#If False Then
    Const ReturnOnlyFSDirs = &H1
    Const DontGoBelowDomain = &H2
    Const StatusText = &H4
    Const ReturnFSAncestors = &H8
    Const EditBox = &H10
    Const Validate = &H20
    Const NewDialogStyle = &H40
    Const UseNewUI = (NewDialogStyle Or EditBox)
    Const BrowseIncludeURLs = &H80
    Const UAHInt = &H100
    Const NoneWFolderButton = &H200
    Const NoTranslateTargets = &H400
    Const BrowseForComputer = &H1000
    Const BrowseForPrinter = &H2000
    Const BrowseIncludeFiles = &H4000
    Const Shareable = &H8000
    Const ShowFolder_Default = ReturnOnlyFSDirs Or StatusText Or BrowseForComputer
#End If

Public Enum ubAppearanceEnum
    [ubFlat] = &H0
    [ub3D] = &H1
End Enum
#If False Then
    Const ubFlat = &H0
    Const ub3D = &H1
#End If

Public Enum ucSpecialFoldersEnum
    AdminTools = CSIDL_ADMINTOOLS
    AltStartUp = CSIDL_ALTSTARTUP
    ApplicationData = CSIDL_APPDATA
    CDBurnArea = CSIDL_CDBURN_AREA
    CommonAdminTools = CSIDL_COMMON_ADMINTOOLS
    CommonAltStartUp = CSIDL_COMMON_ALTSTARTUP
    CommonAppData = CSIDL_COMMON_APPDATA
    CommonDesktopDirectory = CSIDL_COMMON_DESKTOPDIRECTORY
    CommonFavorites = CSIDL_COMMON_FAVORITES
    CommonMyDocuments = CSIDL_COMMON_DOCUMENTS
    CommonMyMusic = CSIDL_COMMON_MUSIC
    CommonMyPictures = CSIDL_COMMON_PICTURES
    CommonMyVideo = CSIDL_COMMON_VIDEO
    CommonProgramFiles = CSIDL_PROGRAM_FILES_COMMON
    CommonPrograms = CSIDL_COMMON_PROGRAMS
    CommonStartMenu = CSIDL_COMMON_STARTMENU
    CommonStartUp = CSIDL_COMMON_STARTUP
    CommonTemplates = CSIDL_COMMON_TEMPLATES
    ComputersNearMe = CSIDL_COMPUTERSNEARME
    Connections = CSIDL_CONNECTIONS
    ControlPanel = CSIDL_CONTROLS
    DeskTop = CSIDL_DESKTOP
    DesktopDirectory = CSIDL_DESKTOPDIRECTORY
    Favorites = CSIDL_FAVORITES
    Fonts = CSIDL_FONTS
    Internet = CSIDL_INTERNET
    InternetCache = CSIDL_INTERNET_CACHE
    InternetCookies = CSIDL_COOKIES
    InternetHistory = CSIDL_HISTORY
    LocalApplicationData = CSIDL_LOCAL_APPDATA
    LocalizedResources = CSIDL_RESOURCES_LOCALIZED
    MyComputer = CSIDL_DRIVES
    MyDocuments = CSIDL_MYDOCUMENTS
    MyDocumentsFolder = CSIDL_PERSONAL
    MyMusic = CSIDL_MYMUSIC
    MyNetworkPlaces = CSIDL_NETHOOD
    MyPictures = CSIDL_MYPICTURES
    MyVideo = CSIDL_MYVIDEO
    NetworkNeighborhood = CSIDL_NETWORK
    Printers = CSIDL_PRINTERS
    PrintHood = CSIDL_PRINTHOOD
    Profile = CSIDL_PROFILE
    Programs = CSIDL_PROGRAMS
    ProgramsFiles = CSIDL_PROGRAM_FILES
    Recent = CSIDL_RECENT
    RecycleBin = CSIDL_BITBUCKET
    SendTo = CSIDL_SENDTO
    StartMenu = CSIDL_STARTMENU
    StartUp = CSIDL_STARTUP
    System = CSIDL_SYSTEM
    SystemResources = CSIDL_RESOURCES
    Templates = CSIDL_TEMPLATES
    Windows = CSIDL_WINDOWS
End Enum
#If False Then
    Const AdminTools = CSIDL_ADMINTOOLS
    Const AltStartUp = CSIDL_ALTSTARTUP
    Const ApplicationData = CSIDL_APPDATA
    Const CDBurnArea = CSIDL_CDBURN_AREA
    Const CommonAdminTools = CSIDL_COMMON_ADMINTOOLS
    Const CommonAltStartUp = CSIDL_COMMON_ALTSTARTUP
    Const CommonAppData = CSIDL_COMMON_APPDATA
    Const CommonDesktopDirectory = CSIDL_COMMON_DESKTOPDIRECTORY
    Const CommonFavorites = CSIDL_COMMON_FAVORITES
    Const CommonMyDocuments = CSIDL_COMMON_DOCUMENTS
    Const CommonMyMusic = CSIDL_COMMON_MUSIC
    Const CommonMyPictures = CSIDL_COMMON_PICTURES
    Const CommonMyVideo = CSIDL_COMMON_VIDEO
    Const CommonProgramFiles = CSIDL_PROGRAM_FILES_COMMON
    Const CommonPrograms = CSIDL_COMMON_PROGRAMS
    Const CommonStartMenu = CSIDL_COMMON_STARTMENU
    Const CommonStartUp = CSIDL_COMMON_STARTUP
    Const CommonTemplates = CSIDL_COMMON_TEMPLATES
    Const ComputersNearMe = CSIDL_COMPUTERSNEARME
    Const Connections = CSIDL_CONNECTIONS
    Const ControlPanel = CSIDL_CONTROLS
    Const DeskTop = CSIDL_DESKTOP
    Const DesktopDirectory = CSIDL_DESKTOPDIRECTORY
    Const Favorites = CSIDL_FAVORITES
    Const Fonts = CSIDL_FONTS
    Const Internet = CSIDL_INTERNET
    Const InternetCache = CSIDL_INTERNET_CACHE
    Const InternetCookies = CSIDL_COOKIES
    Const InternetHistory = CSIDL_HISTORY
    Const LocalApplicationData = CSIDL_LOCAL_APPDATA
    Const LocalizedResources = CSIDL_RESOURCES_LOCALIZED
    Const MyComputer = CSIDL_DRIVES
    Const MyDocuments = CSIDL_MYDOCUMENTS
    Const MyDocumentsFolder = CSIDL_PERSONAL
    Const MyMusic = CSIDL_MYMUSIC
    Const MyNetworkPlaces = CSIDL_NETHOOD
    Const MyPictures = CSIDL_MYPICTURES
    Const MyVideo = CSIDL_MYVIDEO
    Const NetworkNeighborhood = CSIDL_NETWORK
    Const Printers = CSIDL_PRINTERS
    Const PrintHood = CSIDL_PRINTHOOD
    Const Profile = CSIDL_PROFILE
    Const Programs = CSIDL_PROGRAMS
    Const ProgramsFiles = CSIDL_PROGRAM_FILES
    Const Recent = CSIDL_RECENT
    Const RecycleBin = CSIDL_BITBUCKET
    Const SendTo = CSIDL_SENDTO
    Const StartMenu = CSIDL_STARTMENU
    Const StartUp = CSIDL_STARTUP
    Const System = CSIDL_SYSTEM
    Const SystemResources = CSIDL_RESOURCES
    Const Templates = CSIDL_TEMPLATES
    Const Windows = CSIDL_WINDOWS
#End If

Public Enum ubThemeEnum
    [ubAuto] = &H0
    [ubClassic] = &H1
    [ubBlue] = &H2
    [ubHomeStead] = &H3
    [ubMetallic] = &H4
    [ubNone] = &H5
End Enum
#If False Then
    Const ubAuto = &H0
    Const ubClassic = &H1
    Const ubBlue = &H2
    Const ubHomeStead = &H3
    Const ubMetallic = &H4
    Const ubNone = &H4
#End If

'   Private variables
Private bInternal As Boolean
Private bPathChanged As Boolean
Private m_Appearance As ubAppearanceEnum
Private m_CancelButtonWindow As Long
Private m_CheckBoxes As Boolean
Private m_DialogWindow As Long
Private m_Enabled As Boolean
Private m_FolderFlags As ubFolderDialogFlags
Private m_FullRowSelect As Boolean
Private m_HasButtons As Boolean
Private m_HasLines As Boolean
Private m_HideSelection As Boolean
Private m_HotTracking As Boolean
Private m_Path As String
Private m_Root As ucSpecialFoldersEnum
Private m_SysTreeWindow As Long
Private m_Theme As ubThemeEnum

Private WithEvents SDIHost As Form
Attribute SDIHost.VB_VarHelpID = -1
Private WithEvents MDIHost As MDIForm
Attribute MDIHost.VB_VarHelpID = -1

'==================================================================================================
' ucSubclass - A template UserControl for control authors that require self-subclassing without ANY
'              external dependencies. IDE safe.
'
' Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
'
' v1.0.0000 20040525 First cut.....................................................................
' v1.1.0000 20040602 Multi-subclassing version.....................................................
' v1.1.0001 20040604 Optimized the subclass code...................................................
' v1.1.0002 20040607 Substituted byte arrays for strings for the code buffers......................
' v1.1.0003 20040618 Re-patch when adding extra hWnds..............................................
' v1.1.0004 20040619 Optimized to death version....................................................
' v1.1.0005 20040620 Use allocated memory for code buffers, no need to re-patch....................
' v1.1.0006 20040628 Better protection in zIdx, improved comments..................................
' v1.1.0007 20040629 Fixed InIDE patching oops.....................................................
' v1.1.0008 20040910 Fixed bug in UserControl_Terminate, zSubclass_Proc procedure hidden...........
'==================================================================================================
'Subclasser declarations

Public Event MouseEnter()
Public Event MouseLeave()
Public Event Status(ByVal sStatus As String)
Public Event PathChange(ByVal sPath As String)

Private Const WM_ENABLE                 As Long = &HA
Private Const WM_EXITSIZEMOVE           As Long = &H232
Private Const WM_LBUTTONDOWN            As Long = &H201
Private Const WM_LBUTTONUP              As Long = &H202
Private Const WM_MOUSELEAVE             As Long = &H2A3
Private Const WM_MOUSEMOVE              As Long = &H200
Private Const WM_MOVING                 As Long = &H216
Private Const WM_RBUTTONDBLCLK          As Long = &H206
Private Const WM_RBUTTONDOWN            As Long = &H204
Private Const WM_SIZING                 As Long = &H214
Private Const WM_SYSCOLORCHANGE         As Long = &H15
Private Const WM_THEMECHANGED           As Long = &H31A
'Private Const WM_USER                   As Long = &H400


Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                             As Long
    dwFlags                            As TRACKMOUSEEVENT_FLAGS
    hwndTrack                          As Long
    dwHoverTime                        As Long
End Type

Private bTrack                       As Boolean
Private bTrackUser32                 As Boolean
Private bInCtrl                      As Boolean
Private bSubClass                    As Boolean

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Enum eMsgWhen
    MSG_AFTER = 1                                                                   'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2                                                                  'Message calls back before the original (previous) WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                  'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES           As Long = -1                                   'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                    'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                   'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                   'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                   'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                  'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                  'Table A (after) entry count patch offset

Private Type tSubData                                                               'Subclass data type
    hWnd                               As Long                                      'Handle of the window being subclassed
    nAddrSub                           As Long                                      'The address of our new WndProc (allocated memory).
    nAddrOrig                          As Long                                      'The address of the pre-existing WndProc
    nMsgCntA                           As Long                                      'Msg after table entry count
    nMsgCntB                           As Long                                      'Msg before table entry count
    aMsgTblA()                         As Long                                      'Msg after table array
    aMsgTblB()                         As Long                                      'Msg Before table array
End Type

Private sc_aSubData()                As tSubData                                    'Subclass data array

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
    'Parameters:
        'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
        'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
        'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
        'hWnd     - The window handle
        'uMsg     - The message number
        'wParam   - Message related data
        'lParam   - Message related data
    'Notes:
        'If you really know what you're doing, it's possible to change the values of the
        'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
        'values get passed to the default handler.. and optionaly, the 'after' callback
    Dim lpIDList As Long
    Dim lRet As Long
    Dim sBuffer As String
    Dim hWndA As Long
    Dim ClassWindow As String * 14
    Dim ClassCaption As String * 100
    Dim lOffset As Long
    
    '   See if the Path has been set via property but it did not take effect because
    '   the DialogWindow was not created yet.....this can occure if the control is set
    '   at runtime, but the host object is created but not visible yet! If this is the
    '   case the m_DialogWindow = 0 and bPathChanged = False....if we are setting the
    '   path at runtime, but the control and host are visible then bPathChanged = True
    If (m_DialogWindow) And (bPathChanged = False) And (Len(m_Path) > 0) And (m_Path <> "\") Then
'        Call SendMessage(m_DialogWindow, BFFM_SETSELECTION, 1, m_Path)
'        bPathChanged = True
    End If
    
    Select Case uMsg
        Case BFFM_INITIALIZED
            m_DialogWindow = lng_hWnd 'Handle of BrowseForFolder dialog
            'Move the whole  BrowseForFolder dialog off screen
            Call MoveWindow(m_DialogWindow, -Screen.Width, 0, 480, 480, True)
            'Set it's initial path
            Call SendMessage(m_DialogWindow, BFFM_SETSELECTION, 1, m_Path)
            'Enumerate child windows
            hWndA = GetWindow(lng_hWnd, GW_CHILD)
            Do While hWndA <> 0
                GetClassName hWndA, ClassWindow, 14
                'Found a button
                If Left$(ClassWindow, 6) = "Button" Then
                    GetWindowText hWndA, ClassCaption, 100
                    '   If it's the Cancel button, remember it's
                    '   handle so we can press it later
                    If UCase(Left(ClassCaption, 6)) = "CANCEL" Then
                        m_CancelButtonWindow = hWndA
                    End If
                End If
                '   Here's what we're really after - it's Treeview!
                If Left(ClassWindow, 13) = "SysTreeView32" Then
                    m_SysTreeWindow = hWndA
                End If
                hWndA = GetWindow(hWndA, GW_NEXT)
            Loop
            If m_SysTreeWindow <> 0 Then
                '   Steal the Treeview for our own use
                Call GrabSysTreeView
                '   Make the Window Flat so we can handle the
                '   Window Style Locally
                SetWindowStyle m_SysTreeWindow, ubFlat
                
                '   Now Sublass the RightMouseClick to kill the
                '   context menus
                Call Subclass_Start(m_SysTreeWindow)
                Call Subclass_AddMsg(m_SysTreeWindow, WM_RBUTTONDOWN, MSG_BEFORE_AND_AFTER)
                '   Now Refresh things
                Refresh
            Else
                '   Close the Window to prevent hangs
                CloseUp
                '   Opps, we can not find the SystemTreeView
                Debug.Assert False
            End If
            RaiseEvent Status("SystemTreeView Initalized")
            
        Case BFFM_SELCHANGED
            'Path has changed - better tell our form
            sBuffer = Space$(MAX_PATH)
            If Not IsWin2K Then
                lRet = SHGetPathFromIDList(ByVal wParam, ByVal sBuffer)
            Else
                lRet = SHGetPathFromIDListW(ByVal wParam, ByVal sBuffer)
            End If
            If lRet Then
                 'Trim off the null chars ending the path
                 'and display the returned folder
                 lOffset = InStr(sBuffer, Chr$(0))
                 m_Path = QualifyPath(Left$(sBuffer, lOffset - 1))
            Else
                 m_Path = ""
            End If
            RaiseEvent PathChange(m_Path)
            
        Case WM_RBUTTONDOWN
            If bBefore Then
                '   Supress the Right MouseClick for the SysTreeWindow
                '   This will eat the uMsg and prevent the popup window
                '   so we can show our own....
                bHandled = True
                lReturn = 0
            Else
            
            End If
            
        Case WM_MOVE
            TaskbarHide
            
        Case WM_CLOSE
            CloseUp
            
    End Select
    
End Sub

'   Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
    Dim hmod        As Long
    Dim bLibLoaded  As Boolean

    hmod = GetModuleHandleA(sModule)

    If hmod = 0 Then
        hmod = LoadLibraryA(sModule)
        If hmod Then
            bLibLoaded = True
        End If
    End If

    If hmod Then
        If GetProcAddress(hmod, sFunction) Then
            IsFunctionExported = True
        End If
    End If
    If bLibLoaded Then
        Call FreeLibrary(hmod)
    End If
End Function

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

    'Add a message to the table of those that will invoke a callback. You should Subclass_Subclass first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    'Parameters:
        'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
        'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
        'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    'Parameters:
    'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
    'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
    'When      - Whether the msg is to be removed from the before, after or both callback tables
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
    'Parameters:
    'lng_hWnd  - The handle of the window to be subclassed
    'Returns;
    'The sc_aSubData() index
    Const CODE_LEN              As Long = 204                                       'Length of the machine code in bytes
    Const FUNC_CWP              As String = "CallWindowProcA"                       'We use CallWindowProc to call the original WndProc
    Const FUNC_EBM              As String = "EbMode"                                'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    Const FUNC_SWL              As String = "SetWindowLongA"                        'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
    Const MOD_USER              As String = "user32"                                'Location of the SetWindowLongA & CallWindowProc functions
    Const MOD_VBA5              As String = "vba5"                                  'Location of the EbMode function if running VB5
    Const MOD_VBA6              As String = "vba6"                                  'Location of the EbMode function if running VB6
    Const PATCH_01              As Long = 18                                        'Code buffer offset to the location of the relative address to EbMode
    Const PATCH_02              As Long = 68                                        'Address of the previous WndProc
    Const PATCH_03              As Long = 78                                        'Relative address of SetWindowsLong
    Const PATCH_06              As Long = 116                                       'Address of the previous WndProc
    Const PATCH_07              As Long = 121                                       'Relative address of CallWindowProc
    Const PATCH_0A              As Long = 186                                       'Address of the owner object
    Static aBuf(1 To CODE_LEN)  As Byte                                             'Static code buffer byte array
    Static pCWP                 As Long                                             'Address of the CallWindowsProc
    Static pEbMode              As Long                                             'Address of the EbMode IDE break/stop/running function
    Static pSWL                 As Long                                             'Address of the SetWindowsLong function
    Dim i                       As Long                                             'Loop index
    Dim j                       As Long                                             'Loop index
    Dim nSubIdx                 As Long                                             'Subclass data index
    Dim sHex                    As String                                           'Hex code string
    
    'If it's the first time through here..
    If aBuf(1) = 0 Then
        
        'The hex pair machine code representation.
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
            "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
            "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
            "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
        
        'Convert the string from hex pairs to bytes and store in the static machine code buffer
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                                  'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                                                        'Next pair of hex characters
        
        'Get API function addresses
        If Subclass_InIDE Then                                                      'If we're running in the VB IDE
            aBuf(16) = &H90                                                         'Patch the code buffer to enable the IDE state code
            aBuf(17) = &H90                                                         'Patch the code buffer to enable the IDE state code
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                 'Get the address of EbMode in vba6.dll
            If pEbMode = 0 Then                                                     'Found?
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                             'VB5 perhaps
            End If
        End If

        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                        'Get the address of the CallWindowsProc function
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                        'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData                                       'Create the first sc_aSubData element
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If nSubIdx = -1 Then                                                        'If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1                                     'Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                    'Create a new sc_aSubData element
        End If
        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        .hWnd = lng_hWnd                                                            'Store the hWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                               'Allocate memory for the machine code WndProc
        .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                  'Set our WndProc in place
        Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                      'Copy the machine code from the static byte array to the code array in sc_aSubData
        Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                             'Original WndProc address for CallWindowProc, call the original WndProc
        Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                   'Patch the relative address of the SetWindowLongA api function
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                             'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                   'Patch the relative address of the CallWindowProc api function
        Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                             'Patch the address of this object instance into the static machine code buffer
    End With
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
    Dim i As Long
    
    i = UBound(sc_aSubData())                                                       'Get the upper bound of the subclass data array
    Do While i >= 0                                                                 'Iterate through each element
        With sc_aSubData(i)
            If .hWnd <> 0 Then                                                      'If not previously Subclass_Stop'd
                Call Subclass_Stop(.hWnd)                                           'Subclass_Stop
            End If
        End With
        i = i - 1                                                                   'Next element
    Loop
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
    'Parameters:
    'lng_hWnd  - The handle of the window to stop being subclassed
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                         'Restore the original WndProc
        Call zPatchVal(.nAddrSub, PATCH_05, 0)                                      'Patch the Table B entry count to ensure no further 'before' callbacks
        Call zPatchVal(.nAddrSub, PATCH_09, 0)                                      'Patch the Table A entry count to ensure no further 'after' callbacks
        Call GlobalFree(.nAddrSub)                                                  'Release the machine code memory
        .hWnd = 0                                                                   'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                                               'Clear the before table
        .nMsgCntA = 0                                                               'Clear the after table
        Erase .aMsgTblB                                                             'Erase the before table
        Erase .aMsgTblA                                                             'Erase the after table
    End With
End Sub

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
    If bTrack Then
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With
    
        If bTrackUser32 Then
            Call TrackMouseEvent(tme)
        Else
            Call TrackMouseEventComCtl(tme)
        End If
    End If
End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for sc_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry  As Long                                                             'Message table entry index
    Dim nOff1   As Long                                                             'Machine code buffer offset 1
    Dim nOff2   As Long                                                             'Machine code buffer offset 2
    
    If uMsg = ALL_MESSAGES Then                                                     'If all messages
        nMsgCnt = ALL_MESSAGES                                                      'Indicates that all messages will callback
    Else                                                                            'Else a specific message number
        Do While nEntry < nMsgCnt                                                   'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1
            
            If aMsgTbl(nEntry) = 0 Then                                             'This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                                              'Re-use this entry
                Exit Sub                                                            'Bail
            ElseIf aMsgTbl(nEntry) = uMsg Then                                      'The msg is already in the table!
                Exit Sub                                                            'Bail
            End If
        Loop                                                                        'Next entry
        nMsgCnt = nMsgCnt + 1                                                       'New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                'Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg                                                     'Store the message number in the table
    End If

    If When = eMsgWhen.MSG_BEFORE Then                                              'If before
        nOff1 = PATCH_04                                                            'Offset to the Before table
        nOff2 = PATCH_05                                                            'Offset to the Before table entry count
    Else                                                                            'Else after
        nOff1 = PATCH_08                                                            'Offset to the After table
        nOff2 = PATCH_09                                                            'Offset to the After table entry count
    End If

    If uMsg <> ALL_MESSAGES Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                            'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)                                           'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc                                                          'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for sc_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry As Long
    
    If uMsg = ALL_MESSAGES Then                                                     'If deleting all messages
        nMsgCnt = 0                                                                 'Message count is now zero
        If When = eMsgWhen.MSG_BEFORE Then                                          'If before
            nEntry = PATCH_05                                                       'Patch the before table message count location
        Else                                                                        'Else after
            nEntry = PATCH_09                                                       'Patch the after table message count location
        End If
        Call zPatchVal(nAddr, nEntry, 0)                                            'Patch the table message count to zero
    Else                                                                            'Else deleteting a specific message
        Do While nEntry < nMsgCnt                                                   'For each table entry
            nEntry = nEntry + 1
            If aMsgTbl(nEntry) = uMsg Then                                          'If this entry is the message we wish to delete
                aMsgTbl(nEntry) = 0                                                 'Mark the table slot as available
                Exit Do                                                             'Bail
            End If
        Loop                                                                        'Next entry
    End If
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
    'Get the upper bound of sc_aSubData() - If you get an error here, you're probably sc_AddMsg-ing before Subclass_Start
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                                              'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If .hWnd = lng_hWnd Then                                                'If the hWnd of this element is the one we're looking for
                If Not bAdd Then                                                    'If we're searching not adding
                    Exit Function                                                   'Found
                End If
            ElseIf .hWnd = 0 Then                                                   'If this an element marked for reuse.
                If bAdd Then                                                        'If we're adding
                    Exit Function                                                   'Re-use it
                End If
            End If
        End With
    zIdx = zIdx - 1                                                                 'Decrement the index
    Loop
    
    If Not bAdd Then
        Debug.Assert False                                                          'hWnd not found, programmer error
    End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function
'======================================================================================================
'   End SubClass Sections
'======================================================================================================

Public Property Get Appearance() As ubAppearanceEnum
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Value As ubAppearanceEnum)
    UserControl.Appearance = New_Value
    m_Appearance = New_Value
    If Not bInternal Then
        Refresh
    End If
    PropertyChanged "Appearance"
End Property

Private Sub BrowseForFolder(ByVal StartDir As String)
    Dim lpIDList As Long
    Dim szTitle As String
    Dim sBuffer As String
    Dim tBrowseInfo As BROWSEINFO
    
    On Error Resume Next
    With tBrowseInfo
        '   Make the DeskTop Own this Dialog ;-)
        .hwndOwner = GetDesktopWindow()
        .lpszTitle = lStrCat(szTitle, "")
        '   Set the Starting Root Node for the TreeView...
        '   Default: DeskTop = &H0
        .pIDLRoot = m_Root
        '   Set the Dialog Flags that we are after....
        If m_FolderFlags = 0 Then
            .ulFlags = BIF_RETURNONLYFSDIRS + BIF_STATUSTEXT
        Else
            .ulFlags = m_FolderFlags
        End If
        If bSubClass Then
            '   We need to process messages
            .lpfnCallback = sc_aSubData(0).nAddrSub
        End If
    End With
    If Not IsWin2K Then
        lpIDList = SHBrowseForFolder(tBrowseInfo)
    Else
        lpIDList = SHBrowseForFolderW(tBrowseInfo)
    End If
    '   Free the Pointer....
    Call CoTaskMemFree(lpIDList)
End Sub

Public Property Get CheckBoxes() As Boolean
    CheckBoxes = m_CheckBoxes
End Property

Public Property Let CheckBoxes(ByVal New_Value As Boolean)
    m_CheckBoxes = New_Value
    Refresh
    PropertyChanged "CheckBoxes"
End Property

Public Sub CloseUp()
    If m_SysTreeWindow Then
        '   Send the Treeview back to the BrowseForFolder dialog
        SetParent m_SysTreeWindow, m_DialogWindow
        '   Close the dialog.....
        '   First Put the Focus on the Cancel Button
        Call PutFocus(m_CancelButtonWindow)
        '   Now press it via code
        Call PostMessage(m_CancelButtonWindow, WM_LBUTTONDOWN, 0, ByVal 0&)
        Call PostMessage(m_CancelButtonWindow, WM_LBUTTONUP, 0, ByVal 0&)
        '   Now close the dialog.....
        Call SendMessage(m_DialogWindow, WM_CLOSE, 1, 0)
        '   Just to be sure...
        Call DestroyWindow(m_DialogWindow)
        m_CancelButtonWindow = 0
        m_SysTreeWindow = 0
        m_DialogWindow = 0
        RaiseEvent Status("SystemTreeView ShutDown")
    End If
End Sub

Private Sub DrawRectangle(ByRef lpRect As RECT, ByVal lColor As Long)
    Dim hBrush As Long
    Dim lRet As Long
    hBrush = CreateSolidBrush(lColor)
    lRet = FrameRect(UserControl.hDC, lpRect, hBrush)
    DeleteObject hBrush
End Sub

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    m_Enabled = NewValue
    Refresh
    PropertyChanged "Enabled"
End Property

Public Property Get FolderFlags() As ubFolderDialogFlags
    FolderFlags = m_FolderFlags
End Property

Public Property Let FolderFlags(sDialogFlags As ubFolderDialogFlags)
    m_FolderFlags = sDialogFlags
    PropertyChanged "FolderFlags"
End Property

Public Property Get FullRowSelect() As Boolean
    FullRowSelect = m_FullRowSelect
End Property

Public Property Let FullRowSelect(ByVal NewValue As Boolean)
    m_FullRowSelect = NewValue
    Refresh
    PropertyChanged "FullRowSelect"
End Property

Private Function GetThemeColors()
    Dim AutoTheme As String

    Select Case m_Theme
        Case [ubAuto]
            AutoTheme = GetThemeInfo
            Select Case AutoTheme
                Case "None"
                    GoTo Classic

                Case "NormalColor"
                    GoTo Blue

                Case "HomeStead"
                    GoTo HomeStead

                Case "Metallic"
                    GoTo Metallic

            End Select
        Case [ubClassic]
Classic:
            GetThemeColors = &H0

        Case [ubBlue]
Blue:
            GetThemeColors = &HB99D7F

        Case [ubHomeStead]
HomeStead:
            GetThemeColors = &H69A18B

        Case [ubMetallic]
Metallic:
            GetThemeColors = &HB99D7F

    End Select
    If Enabled = False Then
        GetThemeColors = &HC0C0C0
    End If
End Function

Private Function GetThemeInfo() As String
    Dim lResult As Long
    Dim sFilename As String
    Dim sColor As String
    Dim lPos As Long

    If IsWinXP Then
        '   Allocate Space
        sFilename = Space(255)
        sColor = Space(255)
        '   Read the data
        If GetCurrentThemeName(sFilename, 255, sColor, 255, vbNullString, 0) <> &H0 Then
            GetThemeInfo = "UxTheme_Error"
            Exit Function
        End If
        '   Find our trailing null terminator
        lPos = InStrRev(sColor, vbNullChar)
        '   Parse it....
        sColor = Mid(sColor, 1, lPos)
        '   Now replace the nulls....
        sColor = Replace(sColor, vbNullChar, "")
        If Trim$(sColor) = vbNullString Then sColor = "None"
        GetThemeInfo = sColor
    Else
        sColor = "None"
    End If
End Function

Private Sub GrabSysTreeView()
    '   Thievery in progress
    '   It's mine now!
    SetParent m_SysTreeWindow, UserControl.hWnd
    '   Temporary SubClass the BFF Dialog to catch the move event
    SubClassDialog
End Sub

Public Property Get HasButtons() As Boolean
    HasButtons = m_HasButtons
End Property

Public Property Let HasButtons(ByVal New_Value As Boolean)
    m_HasButtons = New_Value
    Refresh
    PropertyChanged "HasButtons"
End Property

Public Property Get HasLines() As Boolean
    HasLines = m_HasLines
End Property

Public Property Let HasLines(ByVal New_Value As Boolean)
    m_HasLines = New_Value
    Refresh
    PropertyChanged "HasLines"
End Property

Public Property Get HideSelection() As Boolean
    HideSelection = m_HideSelection
End Property

Public Property Let HideSelection(ByVal NewValue As Boolean)
    m_HideSelection = NewValue
    Refresh
    PropertyChanged "HideSelection"
End Property

Public Property Get HotTracking() As Boolean
    HotTracking = m_HotTracking
End Property

Public Property Let HotTracking(ByVal New_Value As Boolean)
    m_HotTracking = New_Value
    Refresh
    PropertyChanged "HotTracking"
End Property

Public Function IsFile(ByVal sPath As String) As Boolean
    '   If it is not a Folder, it must be a file....so
    '   the reverse logic holds for determining the File vs. Folder
    IsFile = IsFolder(sPath) = False
End Function

Public Function IsFolder(ByVal sPath As String) As Boolean
    Dim Result As Long
    '   Call the API to see if this a folder path
    Result = PathIsDirectory(sPath)
    IsFolder = (Result = vbDirectory) Or (Result = 1)
End Function

Public Function IsWin2K() As Boolean
'returns True if running Windows XP
    Dim OSV As OSVERSIONINFO

    OSV.OSVSize = Len(OSV)
    If GetVersionEx(OSV) = 1 Then
        IsWin2K = (OSV.PlatformID = VER_PLATFORM_WIN32_NT) And _
        (OSV.dwVerMajor = 5 And OSV.dwVerMinor = 0)
    End If
End Function

Public Function IsWinXP() As Boolean
    'returns True if running Windows XP
    Dim OSV As OSVERSIONINFO

    OSV.OSVSize = Len(OSV)
    If GetVersionEx(OSV) = 1 Then
        IsWinXP = (OSV.PlatformID = VER_PLATFORM_WIN32_NT) And _
        (OSV.dwVerMajor = 5 And OSV.dwVerMinor = 1) And _
        (OSV.dwBuildNumber >= 2600)
    End If
End Function

Private Sub MDIHost_Unload(Cancel As Integer)
    '   Just a little insurance to make sure we clean up
    '   before the Host Object Unloads
    CloseUp
End Sub

Public Property Get Path() As String
    If m_Path = vbNullString Then
        Path = vbNullString
    Else
        Path = QualifyPath(m_Path)
    End If
End Property

Public Property Let Path(sNewPath As String)
    If sNewPath = vbNullString Then
        m_Path = vbNullString
    Else
        m_Path = QualifyPath(sNewPath)
    End If
    If m_DialogWindow Then
        Call SendMessage(m_DialogWindow, BFFM_SETSELECTION, 1, m_Path)
        bPathChanged = True
    End If
    PropertyChanged "Path"
End Property

Public Function QualifyPath(ByVal sPath As String) As String
    Dim lStrCnt As Long
    
    If IsFolder(sPath) Then
        '   Look for the PathSep
        lStrCnt = InStrRev(sPath, "\")
        If (lStrCnt <> Len(sPath)) Or Right$(sPath, 1) <> "\" Then
            '   None, so add it...
            QualifyPath = sPath & "\"
        Else
            '   We are good, so return the value unchanged
            QualifyPath = sPath
        End If
    Else
        '   We are good, this is a File,
        '   so return the value unchanged
        QualifyPath = sPath
    End If
End Function

Public Sub Refresh()
    Dim lpRect As RECT
    Dim AutoTheme As String
    
    With UserControl
        AutoTheme = GetThemeInfo
        If (AutoTheme = "None") Or (m_Theme = ubClassic) Then
            If Appearance <> ub3D Then
                bInternal = True
                Appearance = ub3D
                bInternal = False
            End If
        ElseIf m_Theme <> ubNone Then
            If Appearance <> ubFlat Then
                bInternal = True
                Appearance = ubFlat
                bInternal = False
            End If
        End If
        If m_Appearance = ubFlat Then
            m_Appearance = ubFlat
            .Appearance = 0
            .BorderStyle = 0
            .BackColor = &HFFFFFF
            'Call SetBackColor(&HFFFFFF)
            'Call SetForeColor(&HFF)
            SetRect lpRect, 0, 0, ScaleWidth, ScaleHeight
            Call DrawRectangle(lpRect, GetThemeColors)
        Else
            .Appearance = 1
            .BorderStyle = 1
            .BackColor = &HFFFFFF
        End If
        '   Set the TreeView Specific Properties via APIs
        If m_SysTreeWindow Then
            '   Move the TreeView to our location
            SizeSysTreeView 1, 1, ScaleWidth - 2, ScaleHeight - 2
            '   Set the Style Bits for the TreeView
            Call SetSysTreeStyle(m_CheckBoxes, m_FullRowSelect, m_HasButtons, _
                m_HasLines, m_HideSelection, m_HotTracking)
            '   Set the Window Enabled State
            Call EnableWindow(m_SysTreeWindow, m_Enabled)
            'Call SetBackColor(&H1)
        End If

    End With
    
End Sub

Public Sub Reset()
    CloseUp
    Refresh
    If bSubClass Then
        CloseUp
        BrowseForFolder QualifyPath(App.Path)
    End If
End Sub

Public Property Get Root() As ucSpecialFoldersEnum
    Root = m_Root
End Property

Public Property Let Root(ByVal NewValue As ucSpecialFoldersEnum)
    m_Root = NewValue
    CloseUp
    Reset
    Refresh
    PropertyChanged "Root"
End Property

Private Sub SDIHost_Unload(Cancel As Integer)
    '   Just a little insurance to make sure we clean up
    '   before the Host Object Unloads
    CloseUp
End Sub

Private Sub SetBackColor(ByVal lColor As Long)
    Dim Style As Long
    Dim lColorRef As Long
    
    If m_SysTreeWindow Then
        lColorRef = SendMessage(m_SysTreeWindow, TVM_GETBKCOLOR, 0, ByVal 0)
        '   See if this is a System Color...
        If lColorRef = -1 Then
            lColorRef = lColor Or &HFFFFFF
        End If
        'Change the background
        Call SendMessage(m_SysTreeWindow, TVM_SETBKCOLOR, 0&, ByVal &HFF)
        'reset the treeview style so the
        'tree lines appear properly
        Style = GetWindowLong(m_SysTreeWindow, GWL_STYLE)

        'if the treeview has lines, temporarily
        'remove them so the back repaints to the
        'selected colour, then restore
        If Style And TVS_HASLINES Then
            Call SetWindowLong(m_SysTreeWindow, GWL_STYLE, Style And Not TVS_HASLINES)
            Call SetWindowLong(m_SysTreeWindow, GWL_STYLE, Style Or TVS_HASLINES)
        End If

    End If
  
End Sub

Private Sub SetForeColor(ByVal lColor As Long)

    Dim hwndTV As Long
    Dim Style As Long
    
    If m_SysTreeWindow Then
        hwndTV = m_SysTreeWindow
        
        'Change the background
        Call SendMessage(hwndTV, TVM_SETTEXTCOLOR, 0, ByVal lColor)
        
        'reset the treeview style so the
        'tree lines appear properly
        Style = GetWindowLong(hwndTV, GWL_STYLE)
        
        'if the treeview has lines, temporarily
        'remove them so the back repaints to the
        'selected colour, then restore
        If Style And TVS_HASLINES Then
            Call SetWindowLong(hwndTV, GWL_STYLE, Style Xor TVS_HASLINES)
            Call SetWindowLong(hwndTV, GWL_STYLE, Style)
        End If
    End If
   
End Sub

Private Function SetSysTreeStyle(ByVal bCheckBox As Boolean, ByVal bFullRowSelect As Boolean, _
    ByVal bHasButtons As Boolean, ByVal bHasLines As Boolean, ByVal bHideSelection As Boolean, ByVal bHotTracking As Boolean) As Boolean
    
    Dim dwStyle As Long
          
    dwStyle = GetWindowLong(m_SysTreeWindow, GWL_STYLE)
    '   Set the TreeView window styles. Note that
    '   this style is applied across the entire
    '   TreeView - you can not have some items
    '   allowing proeprties while others don't.
    
    '   First undo everything....then add them all back one at a time ;-)
    '   trick obatined from vbAccelerator web site
    dwStyle = dwStyle And Not (TVS_CHECKBOXES Or TVS_DISABLEDRAGDROP Or _
         TVS_EDITLABELS Or TVS_FULLROWSELECT Or TVS_HASBUTTONS Or _
         TVS_HASLINES Or TVS_INFOTIP Or TVS_LINESATROOT Or TVS_NOSCROLL Or _
         TVS_NOTOOLTIPS Or TVS_SHOWSELALWAYS Or TVS_SINGLEEXPAND Or _
         TVS_TRACKSELECT)
    
    If dwStyle Then
        If bCheckBox Then
            dwStyle = dwStyle Or TVS_CHECKBOXES
        Else
            dwStyle = dwStyle And Not TVS_CHECKBOXES
        End If
        If bFullRowSelect Then
            dwStyle = dwStyle Or TVS_FULLROWSELECT
        Else
            dwStyle = dwStyle And Not TVS_FULLROWSELECT
        End If
        If bHasButtons Then
            dwStyle = dwStyle Or TVS_HASBUTTONS
        Else
            dwStyle = dwStyle And Not TVS_HASBUTTONS
        End If
        If bHasLines Then
            dwStyle = dwStyle Or TVS_HASLINES
        Else
            dwStyle = dwStyle And Not TVS_HASLINES
        End If
        '   This is backwards, so be careful with this...
        If bHideSelection Then
            dwStyle = dwStyle And Not TVS_SHOWSELALWAYS
        Else
            dwStyle = dwStyle Or TVS_SHOWSELALWAYS
        End If
        If bHotTracking Then
            dwStyle = dwStyle Or TVS_TRACKSELECT
        Else
            dwStyle = dwStyle And Not TVS_TRACKSELECT
        End If
       
        SetSysTreeStyle = CBool(SetWindowLong(m_SysTreeWindow, GWL_STYLE, dwStyle))
        SetWindowPos m_SysTreeWindow, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
    End If
    
End Function

Private Sub SetWindowStyle(ByVal lhWnd As Long, Style As ubAppearanceEnum)
    Dim lStyle As Long

    If Style = ubFlat Then
        lStyle = GetWindowLong(lhWnd, GWL_STYLE)
        lStyle = lStyle And Not WS_BORDER
        SetWindowLong lhWnd, GWL_STYLE, lStyle
    
        lStyle = GetWindowLong(lhWnd, GWL_EXSTYLE)
        lStyle = lStyle And Not WS_EX_CLIENTEDGE 'Or WS_EX_STATICEDGE
        SetWindowLong lhWnd, GWL_EXSTYLE, lStyle
        SetWindowPos lhWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
    Else
        lStyle = GetWindowLong(lhWnd, GWL_STYLE)
        lStyle = lStyle And WS_BORDER
        SetWindowLong lhWnd, GWL_STYLE, lStyle
    
        lStyle = GetWindowLong(lhWnd, GWL_EXSTYLE)
        lStyle = lStyle And WS_EX_CLIENTEDGE 'Or WS_EX_STATICEDGE
        SetWindowLong lhWnd, GWL_EXSTYLE, lStyle
        SetWindowPos lhWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
    End If
End Sub

Public Sub SizeSysTreeView(mLeft As Long, mTop As Long, mWidth As Long, mHeight As Long)
    '   Called by the resize event of the Container holding the Treeview
    Call MoveWindow(m_SysTreeWindow, mLeft, mTop, mWidth, mHeight, True)
End Sub

Private Sub SubClassDialog()
    If m_DialogWindow Then
        Call Subclass_Start(m_DialogWindow)
        '   Subclass the BrowseForFolder Message
        Call Subclass_AddMsg(m_DialogWindow, WM_MOVE, MSG_AFTER)
    End If
End Sub

Private Sub TaskbarHide()
    '   Hide the BrowseForFolder dialog from the Taskbar
    ShowWindow m_DialogWindow, 0
    '   Done with Subclassing the dialog window so cancel
    Subclass_Stop m_DialogWindow
End Sub

Public Property Get Theme() As ubThemeEnum
    Theme = m_Theme
End Property

Public Property Let Theme(ByVal New_Value As ubThemeEnum)
    m_Theme = New_Value
    Refresh
    PropertyChanged "Theme"
End Property

Public Function TranslateColor(ByVal lColor As Long) As Long
    On Error GoTo Func_ErrHandler
    
    '   System Color code to long RGB
    If OleTranslateColor(lColor, 0, TranslateColor) Then
        TranslateColor = -1
    End If
    Exit Function
Func_ErrHandler:
End Function

Private Sub UserControl_InitProperties()
    m_Appearance = [ub3D]
    m_Theme = ubAuto
    FolderFlags = ShowFolder_Default
'    Path = App.Path
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Appearance = .ReadProperty("Appearance", [ub3D])
        m_CheckBoxes = .ReadProperty("CheckBoxes", False)
        m_Enabled = .ReadProperty("Enabled", True)
        m_FolderFlags = .ReadProperty("FolderFlags", [ShowFolder_Default])
        m_FullRowSelect = .ReadProperty("FullRowSelect", False)
        m_HasLines = .ReadProperty("HasLines", False)
        m_HasButtons = .ReadProperty("HasButtons", True)
        m_HideSelection = .ReadProperty("HideSelection", True)
        m_HotTracking = .ReadProperty("HotTracking", True)
        m_Path = .ReadProperty("Path", "")
        m_Root = .ReadProperty("Root", [DeskTop])
        m_Theme = .ReadProperty("Theme", [ubAuto])
    End With
    'If we're not in design mode
    If (Ambient.UserMode) Then
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False
            End If
        End If
        If bTrack Then
            'Add the messages that we're interested in
            With UserControl
                '   Start Subclassing using our Handle
                Call Subclass_Start(.hWnd)
                '   Subclass the BrowseForFolder Message
                Call Subclass_AddMsg(.hWnd, BFFM_INITIALIZED, MSG_AFTER)
                Call Subclass_AddMsg(.hWnd, BFFM_SELCHANGED, MSG_AFTER)
                '   Subclass the Parent's QueryClose Event
                If TypeOf .Parent Is Form Then
                    Set SDIHost = .Parent
                ElseIf TypeOf .Parent Is MDIForm Then
                    Set MDIHost = .Parent
                End If
                bSubClass = True
            End With
        End If
    End If
    '   Set the focus on the caller
    Call SetFocus(UserControl.Parent.hWnd)
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub

Private Sub UserControl_Show()
    Refresh
    If bSubClass Then
        BrowseForFolder QualifyPath(App.Path)
    End If
End Sub

Private Sub UserControl_Terminate()
    On Error GoTo Catch
    If bSubClass Then
        '   Return the BFF to its owner
        CloseUp
        '   Stop all subclassing
        Call Subclass_StopAll
        '   Set our Flag that were done....
        bSubClass = False
    End If
Catch:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Appearance", m_Appearance, [ub3D])
        Call .WriteProperty("CheckBoxes", m_CheckBoxes, False)
        Call .WriteProperty("Enabled", m_Enabled, True)
        Call .WriteProperty("FolderFlags", m_FolderFlags, [ShowFolder_Default])
        Call .WriteProperty("FullRowSelect", m_FullRowSelect, False)
        Call .WriteProperty("HasLines", m_HasLines, False)
        Call .WriteProperty("HasButtons", m_HasButtons, True)
        Call .WriteProperty("HideSelection", m_HideSelection, True)
        Call .WriteProperty("HotTracking", m_HotTracking, True)
        Call .WriteProperty("Path", m_Path, "")
        Call .WriteProperty("Root", m_Root, [DeskTop])
        Call .WriteProperty("Theme", m_Theme, [ubAuto])
    End With
End Sub

Public Property Get Version(Optional ByVal bDateTime As Boolean = False) As String
    If bDateTime Then
        Version = Major & "." & Minor & "." & Revision & " (" & DateTime & ")"
    Else
        Version = Major & "." & Minor & "." & Revision
    End If
End Property
