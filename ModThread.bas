Attribute VB_Name = "ModThread"
'===================================================================================================
'| ģ �� �� | ModThread
'| ��    �� | Multi-Thread
'| ˵    �� | please keep the original author of this module Notes;'API:CreateIExprSrvObj,CoInitialize,CoUninitialize by download@vbgood
'| �� �� �� | amicy QQ:35723195 MSN:xb@live.it
'| ��    �� | 2009-04-26 23:23:21
'| ��    �� | 2010-08-22 19:56:45
'| ��    �� | 1.0.0
'===================================================================================================
'| ��    �� | 2011-10-24
'| �� �� �� | ����ѧ�� http://www.vbgood.com/space-uid-144949.html
'| ˵    �� | �����п�Դ���˱�ʾ��л
'===================================================================================================

Option Explicit
Private Type UUID
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(7) As Byte
End Type

Private Declare Function CreateThread Lib "kernel32" (ByVal lpThreadA As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long

'ǿ�ƽ����߳���TerminateThread
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long

Private Declare Sub UserDllMain Lib "msvbvm60.dll" (u1 As Long, u2 As Long, ByVal u3_h As Long, ByVal u4_1 As Long, ByVal u5_0 As Long)
Private Declare Function VBDllGetClassObject Lib "msvbvm60.dll" (gloaders As Long, gvb As Long, ByVal gvbtab As Long, rclsid As UUID, riid As UUID, ppv As Any) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function CoInitialize Lib "ole32.dll" (ByVal pvReserved As Long) As Long
Private Declare Sub CoUninitialize Lib "ole32.dll" ()
Private Declare Function CreateIExprSrvObj Lib "msvbvm60.dll" (ByVal p1_0 As Long, ByVal p2_4 As Long, ByVal p3_0 As Long) As Long

Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
 
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Long)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)

Private Declare Function GlobalFree Lib "kernel32" (hMem As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long

Private Declare Function VirtualQuery Lib "kernel32" (ByVal lpAddress As Long, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long

Private Type MEMORY_BASIC_INFORMATION
     BaseAddress As Long
     AllocationBase As Long
     AllocationProtect As Long
     RegionSize As Long
     State As Long
     Protect As Long
     lType As Long
End Type

Private Type THREAD_DATA
    lpStartAddress As Long
    lpParam As Long
End Type
Dim RunOneFlag As Long

'===================================================================================================
'| �� �� �� | VBCreateThread
'| ˵    �� | ����һ���߳�(Create a thread)
'| ��    �� | lpStartAddress:   �̺߳�����ַ(thread function address )
'| ��    �� | lpParam:          �̲߳���(thread param)
'| ��    �� | lpThreadId:       �߳�ID(tid)
'| �� �� ֵ | �����߳̾��(hThread)
'===================================================================================================
Public Function VBCreateThread(ByVal lpStartAddress As Long, ByVal lpParam As Long, Optional ByRef lpThreadId As Long = 0) As Long
        Dim lpData As Long
        Dim nRet As Long
        
        '������ֽ��û�װ����
        lpData = GlobalAlloc(0, 8)
        PutMem4 lpData, lpStartAddress
        PutMem4 lpData + 4, lpParam
        
        VBCreateThread = CreateThread(0, 0, AddressOf ThreadHead, ByVal lpData, 0, lpThreadId)
End Function

'===================================================================================================
'| �� �� �� | InitVB
'| ˵    �� | ��ʼ��VB���п�(Init vb runtime)
'| ��    �� | �� (void)
'| �� �� ֵ | �� (void)

'ע�⣡����������������
'Ϊ�˷�ֹ���� ���뽫�����������Ϊ sub main (�������� > �������� ����Ϊ sub main)
'��Ϊ���߳�ʱ����ظ����� sub main����fromload

'ע�⣡����������
'InitVB�������sub main������ InitVB��ɾ�� sub main �Ĵ���
'���� �����Ҫ�ظ���ε��� sub main ����д����
'===================================================================================================
Public Function InitVB() As Long
        Dim fake As Long
        Dim lvb As Long
        Dim riid As UUID
        Dim aiid As UUID
        Dim ofac As Object
        Dim nRet As Long
        
        '������ʼ��
        CreateIExprSrvObj 0, 4, 0
        Call CoInitialize(0)
        
        With riid
                .data1 = 1
                .data4(0) = &HC0
                .data4(7) = &H46
        End With
        
        
        '660F56F6 >/$  55                push    ebp
        '660F56F7  |.  8BEC              mov     ebp, esp
        '660F56F9  |.  83EC 20           sub     esp, 20
        '660F56FC  |.  56                push    esi
        '660F56FD  |.  57                push    edi
        '660F56FE  |.  6A 08             push    8
        '660F5700  |.  33C0              xor     eax, eax
        '660F5702  |.  3905 80EE1066     cmp     dword ptr [6610EE80], eax
        '6610EE80 ��0
        '6610EE84 ��0
        
        '����~~  ������Ū����Զ�̬����ж�� ����TMD���̽�����ʱ������������
        '        GetMem4 GetProcAddress(GetModuleHandle("msvbvm60.dll"), "SetMemEvent") + 14, nRet
        '        PutMem4 nRet, 0
        '        PutMem4 nRet + 4, 0

        If RunOneFlag = 0 Then
                GetMem4 AddressOf CallThreadFunc, nRet
                '��CallThreadFunc���� ��̬д����
                If nRet <> &HFF505B58 Then
                        '���� sub main���ж��
                        WriteProcessMemory -1, AddressOf Main, &HC3C3C3C3, 4, ByVal 0
                        '��CallThreadFunc���� ��̬д����
                        WriteProcessMemory -1, AddressOf CallThreadFunc, &HFF505B58, 4, ByVal 0
                        WriteProcessMemory -1, Getval(AddressOf CallThreadFunc) + 4, &H909090E3, 4, ByVal 0
                        RunOneFlag = 1
                End If
                
        End If
        
        fake = GetFakeH(GetCurrentModule())
        UserDllMain nRet, lvb, GetCurrentModule(), 1, 0
        Call VBDllGetClassObject(nRet, lvb, ByVal fake, aiid, riid, ofac)
        
        InitVB = 0
End Function


'�õ���ǰDLL���
Public Function GetCurrentModule() As Long
        Dim mbi As MEMORY_BASIC_INFORMATION
        
        '���ȡһ�������ĵ�ַ
        If (VirtualQuery(AddressOf InitVB, mbi, LenB(mbi)) <> 0) Then
                GetCurrentModule = mbi.AllocationBase
        Else
                GetCurrentModule = 0
        End If
End Function

'tgyԭ�� ���������޸�
'ȡVBͷ,ȫ�µ�ȡVBͷ����,�ٶȱ�OPEN�ļ���ö� �����ַ����ĳ���Ӧ�ÿ��Լ�һЩѹ����
Public Function GetFakeH(ByVal hin As Long) As Long
        Dim lPtr     As Long
        Dim isvb As String
        Dim mdat(4095) As Byte
        
        lPtr = hin + 4096 '�ļ�ͷ+4KB �ӿ��ٶ�
        isvb = StrConv("VB5!", vbFromUnicode)

        CopyMemory mdat(0), ByVal lPtr, 4096
        
        GetFakeH = InStrB(mdat, isvb)
        
        If GetFakeH = 0 Then
                Exit Function
        End If
       GetFakeH = GetFakeH + lPtr - 1
End Function




Private Function CallThreadFunc(ByVal UserFuncAddr As Long, ByVal lpParam As Long) As Long
'д��������ڱ����ǵ��������� ���ȱ����㹻��������Ļ��
MsgBox "YES"
MsgBox "YES"
MsgBox "YES"
MsgBox "YES"

'��̬д��Ļ�� �ֶ�������ջ �ֶ�call
'00402C10    58              pop     eax
'00402C11    5B              pop     ebx
'00402C12    50              push    eax
'00402C13    FFE3            jmp     ebx
'00402C15    90              nop
'00402C16    90              nop
'00402C17    90              nop
'585850 FF
'FF505858
'E3909090
'909090E3
End Function

'�̳߳�ʼ�� �ͷ���Դ�������� ThreadHead �㶨
'�����Ǵ���ָ�� �൱��VB�İ���ַ����
Private Function ThreadHead(ByRef lpParam As THREAD_DATA) As Long

        '�̳߳�ʼ��
        Call InitVB
        
        On Error Resume Next
        
        '�����û��̺߳���
        ThreadHead = CallThreadFunc(lpParam.lpStartAddress, lpParam.lpParam)
        
        'ע�� lpParam�Ƕ�̬����� ��Ҫ�ͷ�
        GlobalFree lpParam
        CoUninitialize
        
        '����������������߳̽����¼�����
End Function

Private Function Getval(ByVal n As Long) As Long
Getval = n
End Function





