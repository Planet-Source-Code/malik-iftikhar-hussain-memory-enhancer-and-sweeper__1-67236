Attribute VB_Name = "mApiDeclaration"
Option Explicit

'Variables and Structure declaration
Public sURLName() As String


''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'API Declaration start here                            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' For getting version information
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type




Public Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

'--------------------------------------------------
'for file handling
Const MAX_PATH = 1024
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

'---------------------------------------
'for url cache entries
Private Type INTERNET_CACHE_ENTRY_INFO
    dwStructSize As Long
    lpszSourceUrlName As Long
    lpszLocalFileName As Long
    CacheEntryType As Long
    dwUseCount As Long
    dwHitRate As Long
    dwSizeLow As Long
    dwSizeHigh As Long
    LastModifiedTime As FILETIME
    ExpireTime As FILETIME
    LastAccessTime As FILETIME
    LastSyncTime As FILETIME
    lpHeaderInfo As Long
    dwHeaderInfoSize As Long
    lpszFileExtension As Long
    dwReserved As Long
    dwExemptDelta As Long
    'szRestOfData() As Byte
End Type

Public Declare Function FindFirstUrlCacheEntry Lib "wininet.dll" Alias "FindFirstUrlCacheEntryA" (ByVal lpszUrlSearchPattern As String, ByVal lpFirstCacheEntryInfo As Long, ByRef lpdwFirstCacheEntryInfoBufferSize As Long) As Long
Public Declare Function FindNextUrlCacheEntry Lib "wininet.dll" Alias "FindNextUrlCacheEntryA" (ByVal hEnumHandle As Long, ByVal lpNextCacheEntryInfo As Long, ByRef lpdwNextCacheEntryInfoBufferSize As Long) As Long
Public Declare Sub FindCloseUrlCache Lib "wininet.dll" (ByVal hEnumHandle As Long)
Public Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long

'-------------------------------------------------------------------
'preserve the width and height of form
Public lWidth  As Long
Public lHeight As Long
Public bClick  As Boolean

'for sweep button clicked
Public bSweep As Boolean


'API Sub and Function Declaration start here
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpShortPath As String, ByVal lpLongPath As String, ByVal nFullPathSize As Long) As Long


'api for file handling
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

'free memory at a given address
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Long, ByVal ByteLen As Long)
Public Declare Sub CopyMemoryToStruct Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, ByVal pSrc As Long, ByVal ByteLen As Long)

Public Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Public Declare Sub LocalFree Lib "kernel32" (hPtr As Long)

'string concatination
'append two string and return memory address
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function lstrcatURL Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
'string length of a string
'present in buffer
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long


'api for folder handling
'for browsing folder following Functions r used
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

'api for Clipboard data
Public Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function SetClipboardViewer Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetClipboardViewer Lib "user32" () As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Sub OpenClipboard Lib "user32" (ByVal hwnd As Long)
Public Declare Function CloseClipboard Lib "user32" () As Long

'for user profile directory
Const TOKEN_QUERY = (&H8)
Public Declare Function GetUserProfileDirectory Lib "userenv.dll" Alias "GetUserProfileDirectoryA" (ByVal hToken As Long, ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long

'for .ini file  manipulations
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'registry manipulations
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'shut down windows
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
'change computer name
Public Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
'desktop related api
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

'internet and network information
Public Const NETWORK_ALIVE_AOL As Long = &H4
Public Const NETWORK_ALIVE_LAN As Long = &H1
Public Const NETWORK_ALIVE_WAN As Long = &H2

Public Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
Public Const INTERNET_CONNECTION_LAN As Long = &H2
Public Const INTERNET_CONNECTION_MODEM As Long = &H1
Public Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
Public Const INTERNET_CONNECTION_OFFLINE As Long = &H20
Public Const INTERNET_CONNECTION_PROXY As Long = &H4


Public Declare Function IsNetworkAlive Lib "SENSAPI.dll" (ByRef lpdwFlags As Long) As Long
Public Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long





''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Application Name  :SweepMemory                        '
'Created By        :Mayur.Kotlikar                     '
'Creation Date     :10 June 2003                       '
'Purpose           :Assignment(Internal)               '
'Dependencies      :MicrosoftScripting Runtime Lib     '
'                  :VBA and VB Runtime Libraries       '
'   Component      :MS Tabbed Dialog control 6.0 SP5   '
'                  :MS Windows Common controls 6.0 SP4 '
'OS Dependencies   :None                               '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''





'Application start up
'Also used to collect user information




Public Sub Main()

On Error GoTo err
    Dim sTemp      As String
    Dim sFullPath  As String
    Dim lLong      As Long, lResult As Long
    
    'used for checking ini files
    Dim fso        As Scripting.FileSystemObject
    
    
    sTemp = Space$(255)
    
    'get comp name
    lLong = GetComputerName(sTemp, Len(sTemp))
    sTemp = Mid(sTemp, 1, InStr(1, sTemp, Chr$(0)))
    
    frmSweepMemory.lblCompName = sTemp
      
    sTemp = Space$(255)
    
    'get the user name
    lLong = GetUserName(sTemp, Len(sTemp))
    sTemp = Mid(sTemp, 1, InStr(1, sTemp, Chr$(0)))
    
    frmSweepMemory.lblUserName = sTemp
    
    'get temp folder path
    sTemp = Space$(255)
    lLong = GetTempPath(255, sTemp)
    'sTemp = Mid(sTemp, 1, InStr(1, sTemp, Chr$(0)))
    sFullPath = Space$(1024)
    
    lResult = GetLongPathName(sTemp, sFullPath, 1024)
    
    frmSweepMemory.lblTempFolderPath = sFullPath
       
        
    sTemp = Space$(0)
    sFullPath = Space$(0)
    lWidth = frmSweepMemory.Width
    lHeight = frmSweepMemory.Height
    
    frmSweepMemory.Show
    
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(App.Path & "\SweepMemory.ini") Then
        If MsgBox(".ini File of the application doesnot exist.This application requires .ini file." & vbCrLf & "Do you want to create it?", vbYesNo) = vbYes Then
            'create ini file
            IniManipulations "Create"
            
            'now load the form with tab2 on display for user to
            'allow settings
           
            
        End If
    Else  'since file exists apply its setting to Settings tab
          'we have Application section in our ini
          'and we have 3 keys Startup,Temp,Recycle
          
          sTemp = Space$(255)
          'for startup
          Call GetPrivateProfileString("Application", "Startup", vbNullString, sTemp, 255, App.Path & "\SweepMemory.ini")
          sTemp = Mid(sTemp, 1, InStr(1, sTemp, Chr$(0)) - 1)
          
          If Not IsNull(sTemp) Then
            frmSweepMemory.chkSettings(0).Value = IIf(sTemp = "True", 1, 0)
          End If
          
          sTemp = Space$(255)
          'for Temp
          Call GetPrivateProfileString("Application", "Temp", vbNullString, sTemp, 255, App.Path & "\SweepMemory.ini")
          sTemp = Mid(sTemp, 1, InStr(1, sTemp, Chr$(0)) - 1)
          
          If Not IsNull(sTemp) Then
            frmSweepMemory.chkSettings(1).Value = IIf(sTemp = "True", 1, 0)
            bSweep = True
            FindFiles frmSweepMemory.lblTempFolderPath.Caption
          End If
          
          sTemp = Space$(255)
          'for Recycle
          Call GetPrivateProfileString("Application", "Recycle", vbNullString, sTemp, 255, App.Path & "\SweepMemory.ini")
          sTemp = Mid(sTemp, 1, InStr(1, sTemp, Chr$(0)) - 1)
          
          If Not IsNull(sTemp) Then
            frmSweepMemory.chkSettings(2).Value = IIf(sTemp = "True", 1, 0)
          End If
          
              
              
              
              
    End If
        
        
    bSweep = False
    
    'get windows version
    GetWindowsVersion
    
    Exit Sub
err:
    MsgBox "Error while loading program." & err.Description, vbInformation
    bSweep = False
    
End Sub

'This procedure is used to loop
'and find the files in a folder
'the argument of this sub is
'name of folder
'this sub will fill the listview lvwFiles
'with found files
Public Sub FindFiles(sFolderPath As String, Optional sSpecification As String)
On Error GoTo err

    If Trim(sFolderPath) = vbNullString Then
        MsgBox "Please make some selection.", vbInformation
        Exit Sub
    End If

frmSweepMemory.lvwFilesAndFolders.ListItems.Clear
    
    Dim FileName   As String        ' Walking filename variable...
    Dim DirName    As String        ' SubDirectory Name
    Dim dirNames() As String        ' Buffer for directory name entries
    Dim nDir       As Integer       ' Number of directories in this path
    Dim i          As Integer       ' For-loop counter...
    Dim hSearch    As Long          ' Search Handle
    Dim WFD        As WIN32_FIND_DATA
    Dim Cont       As Integer
    Dim oLst       As ListItem
    Dim iCount     As Integer
    Dim sDelimit   As String
    
    Cont = 1
    
    'it is quite possible that the
    'folder may have space in between them
    'hence find for continous 4 blank spaces
    
    sDelimit = Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
        
    If Not Mid(sFolderPath, Len(sFolderPath)) = "\" Then
        sFolderPath = sFolderPath & "\"
    End If
 
    If IsMissing(sSpecification) Then
        sSpecification = "\*.*"
    End If
 
    hSearch = FindFirstFile(sFolderPath & sSpecification, WFD)
    
    'if everything is ok then
    If hSearch <> INVALID_HANDLE_VALUE Then
    Do While Cont <> 0
        'first get all the sub directory in the highest level
        DirName = Mid(WFD.cFileName, 1, InStr(1, WFD.cFileName, sDelimit) - 1)
        If (DirName <> ".") And (DirName <> "..") Then
            If GetFileAttributes(sFolderPath & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
               If bSweep Then  'delete
                    'set readonly to normal
                    If GetFileAttributes(sFolderPath & DirName) And FILE_ATTRIBUTE_READONLY Then
                        SetFileAttributes sFolderPath & DirName, FILE_ATTRIBUTE_NORMAL
                    End If
                    DeleteFile sFolderPath & DirName
                    Debug.Print sFolderPath & DirName & "  Error  " & err.LastDllError
               Else
                    If GetFileAttributes(sFolderPath & DirName) And FILE_ATTRIBUTE_READONLY Then
                        SetFileAttributes sFolderPath & DirName, FILE_ATTRIBUTE_NORMAL
                    End If
                    Set oLst = frmSweepMemory.lvwFilesAndFolders.ListItems.Add(, , DirName)
                    oLst.SubItems(1) = "Folder"
                    Debug.Print sFolderPath & DirName & "  Error  " & err.LastDllError
               End If
             Else
                If bSweep Then
                    'set readonly to normal
                    If GetFileAttributes(sFolderPath & DirName) And FILE_ATTRIBUTE_READONLY Then
                        SetFileAttributes sFolderPath & DirName, FILE_ATTRIBUTE_NORMAL
                    End If
                    DeleteFile sFolderPath & DirName
                    Debug.Print sFolderPath & DirName & "  Error  " & err.LastDllError
                Else
                    Set oLst = frmSweepMemory.lvwFilesAndFolders.ListItems.Add(, , DirName)
                    oLst.SubItems(1) = "File"
                    Debug.Print sFolderPath & DirName & "  Error  " & err.LastDllError
                End If
            End If
        End If
        iCount = iCount + 1
        Cont = FindNextFile(hSearch, WFD)
        Debug.Print err.LastDllError
    Loop
        Cont = FindClose(hSearch)
    End If
    
    Exit Sub
err:
    bSweep = False
    MsgBox err.LastDllError
    
    
End Sub


'get Temporary Internet Folder
'this folder is located below Local Settings
Public Function GetIETempFolder() As String
    Dim sIETemp As String
    Dim hToken  As Long
    Dim sBuffer As String
    Dim sDelimit   As String
    
   
    'it is quite possible that the
    'folder may have space in between them
    'hence find for continous 4 blank spaces
    
    sDelimit = Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
    
    
    sBuffer = Space$(255)
    sIETemp = "\Local Settings\Temporary Internet Files"

    'get token for current user
    OpenProcessToken GetCurrentProcess, TOKEN_QUERY, hToken
    GetUserProfileDirectory hToken, sBuffer, 255
    
    sBuffer = Mid(sBuffer, 1, InStr(1, sBuffer, Chr$(0)) - 1)
    
    GetIETempFolder = sBuffer & sIETemp
    

End Function

'sub for getting the URLs in the Temp Internet Files folder
Public Sub FindUrls()
    Dim hHandle     As Long   'for storing handle of first found
    Dim bRet        As Boolean
    Dim dwBuffer    As Long
    Dim ICEI        As INTERNET_CACHE_ENTRY_INFO
    Dim ptrICEI     As Long
    Dim LMEM_FIXED  As Long
    Dim oLst        As ListItem
    Dim iFileCount  As Integer
    Dim i
    Dim sTemp       As String
    LMEM_FIXED = &H0
    
    'first just determine the
    'size of buffer requird for storage
    hHandle = FindFirstUrlCacheEntry(vbNullString, ByVal 0, dwBuffer)
      
    'if function fails
    If IsNull(hHandle) Then
        Debug.Print err.LastDllError
        Exit Sub
    End If
    
    ptrICEI = LocalAlloc(LMEM_FIXED, dwBuffer)
    
    If ptrICEI Then   'allocation successful
        CopyMemory ByVal ptrICEI, dwBuffer, 4
    End If
    
    'now i can call the first file again
    '
    hHandle = FindFirstUrlCacheEntry(vbNullString, ByVal ptrICEI, dwBuffer)
    
    Do
        'data copied to icei structure
        'first file copied
        CopyMemoryToStruct ICEI, ptrICEI, Len(ICEI) 'dwBuffer
        
        If (ICEI.CacheEntryType And &H1) Then
            iFileCount = iFileCount + 1
            ReDim sURLName(iFileCount)
            sTemp = Space$(lstrlen(ByVal ICEI.dwSizeLow))
            Call lstrcatURL(ByVal sTemp, ByVal ICEI.lpszSourceUrlName)
            sURLName(iFileCount - 1) = Trim(sTemp)
        End If
        
        'again deallocate and allocate
        'buffer for subsequent api calls
        dwBuffer = 0
        LocalFree ptrICEI
        Call FindNextUrlCacheEntry(hHandle, ByVal 0, dwBuffer)
        
        'now we have buffer size
        'allocate memory again
        ptrICEI = LocalAlloc(LMEM_FIXED, dwBuffer)
        
        If ptrICEI Then   'if allocation successful
            CopyMemory ByVal ptrICEI, dwBuffer, 4
        End If
    
    Loop While FindNextUrlCacheEntry(hHandle, ptrICEI, dwBuffer)
    
    
    'now fill the list view using surlfiles
    
    For i = LBound(sURLName) To UBound(sURLName)
        Set oLst = frmSweepMemory.lvwFilesAndFolders.ListItems.Add(, , sURLName(i))
    Next i
    
    

End Sub


'creating ini file
'this sub should be used to save and create ini file
'if the user has accidently deleted the
'ini file ini file shold be created with default settings
'if user add certain settings it should be
'saved from here
'parameter passed is action
'action has two values "Create"
'                      "Save"

Public Sub IniManipulations(sAction As String)
On Error GoTo err
    Select Case Trim(sAction)
        Case "Create"      'to create a new .ini file
            Dim fso As Scripting.FileSystemObject
            Set fso = New Scripting.FileSystemObject
            fso.CreateTextFile App.Path & "\SweepMemory.ini", True
            
            'add one main section Application
            'we have 3 sting values
            '1 StartUp    for adding application to start up
            '2 Recycle    for empting recycle bin at start up
            '3 Temp       for emptying Temp at start up
            
            Call WritePrivateProfileSection("Application", vbNullString, App.Path & "\SweepMemory.ini")
            
            'now add privateprofilestring
            
            Call WritePrivateProfileString("Application", "Startup", "False", App.Path & "\SweepMemory.ini")
            Call WritePrivateProfileString("Application", "Recycle", "False", App.Path & "\SweepMemory.ini")
            Call WritePrivateProfileString("Application", "Temp", "False", App.Path & "\SweepMemory.ini")
            
        Case "Save"
    End Select
    
    Exit Sub
err:
    MsgBox err.Description & ". In IniManipulation."
End Sub



'shut down and restart windows
'reboot only for administrator
Public Sub ShutDown()
On Error GoTo err
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4

    Dim MSG, ret
    
    If MsgBox("Restart your Machine now?", vbYesNo) = vbYes Then
    'reboot the computer
    ret = ExitWindowsEx(EWX_REBOOT Or EWX_FORCE, 0)
    Else
        Exit Sub
    End If
    Exit Sub
err:
    MsgBox "Error while restarting machine.", vbInformation
    MsgBox err.LastDllError
End Sub


'registry manipulations....
'put your application information in registry
'for example if run on startup is called
'put ur application pathin
' HKEY_LOCAL_MACHINE
' Sub Key SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN
' PUT A STRING SWEEPMEMORY AND PATH OF UR APP

'this sub takes action as input
'if action="Register" then the app
'will be registered... and vice versa

Public Sub RegisterForStartUp(sAction As String)
    On Error GoTo err
    
    Dim sSubKey        As String
    Dim sStringValue   As String
    Dim sAppPath       As String
    Dim hOpenKeyHandle As Long
    Dim lReturn        As Long
    Dim sReceived      As String
    Dim lDataType      As Long
    
    Const ERROR_SUCCESS As Long = 0&
    Const HKEY_LOCAL_MACHINE = &H80000002
    Const key_all_access = &HF003F
    Const REG_SZ = 1
    'initial values set
    
    sSubKey = "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN"
    sStringValue = "SWEEPMEMORY"
    sAppPath = App.Path & "\SweepMemory.exe"
    sReceived = Space$(255)
    
    'first open the key and return its handle
    lReturn = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sSubKey, 0, key_all_access, hOpenKeyHandle)
    
    If lReturn <> ERROR_SUCCESS Then
        Debug.Print err.LastDllError
        Exit Sub
    End If
    
    'now query the registry for existance of sStringValue
    'if it is present do nothing
    'if not then do according to select case
    
    
    lReturn = RegQueryValueEx(hOpenKeyHandle, sStringValue, 0, lDataType, ByVal sReceived, 255)
    
    
    'now check for return value and take action
    Select Case sAction
        Case "Register"
            
            'key not found so we need to create it
            'if key is found then do noting
            If lReturn <> ERROR_SUCCESS Then
                'apply settings
                Call RegSetValueEx(hOpenKeyHandle, sStringValue, 0, REG_SZ, sAppPath, Len(sAppPath))
            End If
                 
        Case Else
            'if registry key found then
            'we need to delete it
            If lReturn = ERROR_SUCCESS Then
                Call RegDeleteValue(hOpenKeyHandle, sStringValue)
            End If
    End Select
    
    'close open regsitry key in all cases
    Call RegCloseKey(hOpenKeyHandle)
    Exit Sub
err:
    MsgBox "Error in RegisterForStartup"
    Call RegCloseKey(hOpenKeyHandle)
End Sub



'get internet information...
'this sub sets the label captions
'in the tab3 of internet frame
'this sub is called when the get information button is clicked

Public Sub InterNetNetwork()
    
    Dim lNetworkAlive As Long
    Dim bState        As Boolean
    Dim sIConnName    As String
    Dim lIConnFlag    As Long
    
    'get the network status
    bState = IsNetworkAlive(lNetworkAlive)
    
    
    If bState Then    ''system connected to internet
        'set the caption depending upon the connection type
        Select Case lNetworkAlive
            Case NETWORK_ALIVE_AOL
                frmSweepMemory.lblNetworkConnection.Caption = "AOL Network."
            Case NETWORK_ALIVE_LAN
                frmSweepMemory.lblNetworkConnection.Caption = "LAN Network."
            Case NETWORK_ALIVE_WAN
                frmSweepMemory.lblNetworkConnection.Caption = "WAN Network."
            Case Else
                frmSweepMemory.lblNetworkConnection.Caption = "Unknown Network."
        End Select
        
        
        'get the internet connection state
        sIConnName = Space$(255)
        bState = InternetGetConnectedStateEx(lIConnFlag, sIConnName, 255, 0)
        
        If bState Then  'internet connection exists
            If (lIConnFlag And INTERNET_CONNECTION_CONFIGURED) = INTERNET_CONNECTION_CONFIGURED Then
                frmSweepMemory.lblInternetConnectionState.Caption = "Configured Connection."
            End If
            If (lIConnFlag And INTERNET_CONNECTION_MODEM) = INTERNET_CONNECTION_MODEM Then
                frmSweepMemory.lblInternetConnectionState.Caption = "Modem Connection."
            End If
            If (lIConnFlag And INTERNET_CONNECTION_PROXY) = INTERNET_CONNECTION_PROXY Then
                frmSweepMemory.lblInternetConnectionState.Caption = "Proxy Connection."
            End If
            If (lIConnFlag And INTERNET_CONNECTION_LAN) = INTERNET_CONNECTION_LAN Then
                frmSweepMemory.lblInternetConnectionState.Caption = "LAN Connection."
            End If
        Else
            frmSweepMemory.lblInternetConnectionState.Caption = "No Internet Conncetion."
        End If
        
        
    Else
            frmSweepMemory.lblNetworkConnection.Caption = "No Network."
        
    End If
    
    
    
    
    

End Sub


'get the windows version number and built
Public Sub GetWindowsVersion()
    
    
    'get the structre variable first
    
    Dim osvr As OSVERSIONINFO
    Dim ret
    osvr.dwOSVersionInfoSize = Len(osvr)
    
    ret = GetVersionEx(osvr)
    
    frmSweepMemory.lblOSVersion = "Windows Version " & osvr.dwMajorVersion & "." & osvr.dwMinorVersion & " Built " & osvr.dwBuildNumber
    
    
    
    

End Sub


