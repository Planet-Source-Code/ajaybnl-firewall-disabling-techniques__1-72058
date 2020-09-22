Attribute VB_Name = "K"
Option Explicit
Private Const MAX_PATH                     As Long = 260
Private Const TH32CS_SNAPHEAPLIST          As Long = &H1
Private Const TH32CS_SNAPPROCESS           As Long = &H2
Private Const TH32CS_SNAPTHREAD            As Long = &H4
Private Const TH32CS_SNAPMODULE            As Long = &H8
Private Const TH32CS_SNAPALL               As Double = (TH32CS_SNAPHEAPLIST + TH32CS_SNAPPROCESS + TH32CS_SNAPTHREAD + TH32CS_SNAPMODULE)
Public Type PROCESSENTRY32
    dwSize                                     As Long
    cntUsage                                   As Long
    th32ProcessID                              As Long
    th32DefaultHeapID                          As Long
    th32ModuleID                               As Long
    cntThreads                                 As Long
    th32ParentProcessID                        As Long
    pcPriClassBase                             As Long
    dwFlags                                    As Long
    szexeFile                                  As String * MAX_PATH
End Type
Public Type FILETIME
    dwLowDateTime                              As Long
    dwHighDateTime                             As Long
End Type
Public Type WIN32_FIND_DATA
    dwFileAttributes                           As Long
    ftCreationTime                             As FILETIME
    ftLastAccessTime                           As FILETIME
    ftLastWriteTime                            As FILETIME
    nFileSizeHigh                              As Long
    nFileSizeLow                               As Long
    dwReserved0                                As Long
    dwReserved1                                As Long
    cFileName                                  As String * MAX_PATH
    cAlternate                                 As String * 14
End Type
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, _
                                                        lppe As Any) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, _
                                                       lppe As Any) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, _
                                                          ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
                                                     ByVal bInheritHandle As Long, _
                                                     ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, _
                                                                  ByVal th32ProcessID As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long




Public Function KillProcess(ByVal Proc As String) As String


Dim hSnapshot As Long
Dim lret      As Long
Dim P         As PROCESSENTRY32
Dim Hand      As Long
Dim ExitCode As Long
    'Process List & Termination
    On Error GoTo KillProcess_Error
    P.dwSize = Len(P)
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, ByVal 0)
    If hSnapshot Then
        lret = Process32First(hSnapshot, P)
        Do While lret
                    If InStr(1, P.szexeFile, Proc, vbTextCompare) > 0 Then
                    Hand = OpenProcess(1, True, P.th32ProcessID)
                    
lret = GetExitCodeProcess(Hand, ExitCode)
                   
                    TerminateProcess Hand, ExitCode
                    End If
            lret = Process32Next(hSnapshot, P)
        Loop
        lret = CloseHandle(hSnapshot)
    End If
    

Exit Function

KillProcess_Error:
    err.Clear

End Function

Public Function isProcess(ByVal Proc As String) As Boolean


Dim hSnapshot As Long
Dim lret      As Long
Dim P         As PROCESSENTRY32

    
    On Error GoTo KillProcess_Error
    P.dwSize = Len(P)
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, ByVal 0)
    If hSnapshot Then
        lret = Process32First(hSnapshot, P)
        Do While lret
            If InStr(1, P.szexeFile, Proc, vbTextCompare) > 0 Then
                isProcess = True
                Exit Function
            End If
            lret = Process32Next(hSnapshot, P)
        Loop
        lret = CloseHandle(hSnapshot)
    End If
    Exit Function

KillProcess_Error:
    err.Clear

End Function


