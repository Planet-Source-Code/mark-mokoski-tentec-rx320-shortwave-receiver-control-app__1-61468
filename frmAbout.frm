VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About COM Detect and Test"
   ClientHeight    =   4485
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7485
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3095.626
   ScaleMode       =   0  'User
   ScaleWidth      =   7028.802
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox websiteLabel 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      MouseIcon       =   "frmAbout.frx":1272
      MousePointer    =   99  'Custom
      MultiLine       =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "frmAbout.frx":157C
      Top             =   4080
      Width           =   4815
   End
   Begin VB.TextBox EmailLabel 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      Locked          =   -1  'True
      MouseIcon       =   "frmAbout.frx":158B
      MousePointer    =   99  'Custom
      MultiLine       =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "frmAbout.frx":1895
      Top             =   3720
      Width           =   5055
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "frmAbout.frx":18B3
      Top             =   3120
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "frmAbout.frx":18E0
      Top             =   840
      Width           =   5055
   End
   Begin VB.PictureBox MarkPic 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   2685
      Left            =   120
      ScaleHeight     =   1843.625
      ScaleMode       =   0  'User
      ScaleWidth      =   1411.69
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   2070
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   5640
      MouseIcon       =   "frmAbout.frx":1929
      MousePointer    =   99  'Custom
      Picture         =   "frmAbout.frx":1C33
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   6873.858
      Y1              =   2070.653
      Y2              =   2070.653
   End
   Begin VB.Label lblTitle 
      Caption         =   "COM Detect and Test"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   4995
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   6873.858
      Y1              =   2070.653
      Y2              =   2070.653
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 0.1.X (Development)"
      Height          =   225
      Left            =   2280
      TabIndex        =   3
      Top             =   600
      Width           =   4995
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '**************************************************************
    '
    '   My standard About form
    '
    '   Mark Mokoski
    '   04-SEPT-2002
    '
    '   Based on standard MSDN About form, with custom addtions
    '
    '**************************************************************

    Option Explicit

    ' Reg Key Security Options...
    Const READ_CONTROL = &H20000
    Const KEY_QUERY_VALUE = &H1
    Const KEY_SET_VALUE = &H2
    Const KEY_CREATE_SUB_KEY = &H4
    Const KEY_ENUMERATE_SUB_KEYS = &H8
    Const KEY_NOTIFY = &H10
    Const KEY_CREATE_LINK = &H20
    Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
    KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
    KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
    ' Reg Key ROOT Types...
    Const HKEY_LOCAL_MACHINE = &H80000002
    Const ERROR_SUCCESS = 0
    Const REG_SZ = 1                         ' Unicode nul terminated string
    Const REG_DWORD = 4                      ' 32-bit number

    Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
    Const gREGVALSYSINFOLOC = "MSINFO"
    Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
    Const gREGVALSYSINFO = "PATH"

    Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
    Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
    Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
    
    'Shell out API for HTML files, Mail and Web Browser
    Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal HWND As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdOK_Click()

    Unload Me

End Sub

Private Sub emailLabel_Click()

    'Sample call:
    'ShellExecute hWnd, vbNullString, "mailto:user@domain.com?body=hello%0a%0world", vbNullString, vbNullString, vbNormalFocus
    ShellExecute HWND, vbNullString, "mailto:markm@cmtelephone.com?Subject=Questions or Comments on " & App.Title & ".", vbNullString, vbNullString, vbNormalFocus
  
    'In order to be able to put carriage returns or tabs in your text,
    'replace vbCrLf and vbTab with the following HEX codes:
    '%0a%0d = vbCrLf
    '%09 = vbTab
    'These codes also work when sending URLs to a browser (GET, POST, etc.)
   
    
End Sub


Private Sub Form_Load()

    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    MarkPic.Picture = LoadResPicture(101, 0)
    
End Sub

Public Sub StartSysInfo()

    On Error GoTo SysInfoErr
  
    Dim SysInfoPath            As String
    
    ' Try To Get System Info Program Path\Name From Registry...

        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
            ' Try To Get System Info Program Path Only From Registry...
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
            ' Validate Existance Of Known 32 Bit File Version

                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                    SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
                    ' Error - File Can Not Be Found...
                Else
                    GoTo SysInfoErr
                End If

            ' Error - Registry Entry Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly

End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean

    Dim i                      As Long                                           ' Loop Counter
    Dim rc                     As Long                                          ' Return Code
    Dim hKey                   As Long                                        ' Handle To An Open Registry Key
    Dim KeyValType             As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal                 As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize             As Long                                  ' Size Of Registry Key Variable

    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------

    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
    KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
        If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
            tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
        Else                                                    ' WinNT Does NOT Null Terminate String...
            tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
        End If

    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------

        Select Case KeyValType                                  ' Search Data Types...
            Case REG_SZ                                             ' String Registry Key Data Type
                KeyVal = tmpVal                                     ' Copy String Value
            Case REG_DWORD                                          ' Double Word Registry Key Data Type

                For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                    KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
                Next

            KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
        End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:                                          ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key

End Function

Private Sub MarkPic_Click()

    'Sample call:
    'ShellExecute hWnd, vbNullString, "mailto:user@domain.com?body=hello%0a%0world", vbNullString, vbNullString, vbNormalFocus
    ShellExecute HWND, vbNullString, "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&txtCriteria=mark+mokoski&lngWId=1&B1=Quick+Search", vbNullString, vbNullString, vbNormalFocus
  
    'In order to be able to put carriage returns or tabs in your text,
    'replace vbCrLf and vbTab with the following HEX codes:
    '%0a%0d = vbCrLf
    '%09 = vbTab
    'These codes also work when sending URLs to a browser (GET, POST, etc.)

End Sub

Private Sub websiteLabel_Click()

    'Sample call:
    'ShellExecute hWnd, vbNullString, "mailto:user@domain.com?body=hello%0a%0world", vbNullString, vbNullString, vbNormalFocus
    ShellExecute HWND, vbNullString, "http://www.rjillc.com", vbNullString, vbNullString, vbNormalFocus
  
    'In order to be able to put carriage returns or tabs in your text,
    'replace vbCrLf and vbTab with the following HEX codes:
    '%0a%0d = vbCrLf
    '%09 = vbTab
    'These codes also work when sending URLs to a browser (GET, POST, etc.)
    
    
End Sub
