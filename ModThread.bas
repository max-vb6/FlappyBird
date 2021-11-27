Attribute VB_Name = "ModThread"
'===================================================================================================
'| 模 块 名 | ModThread
'| 描    述 | Multi-Thread
'| 说    明 | please keep the original author of this module Notes;'API:CreateIExprSrvObj,CoInitialize,CoUninitialize by download@vbgood
'| 创 建 人 | amicy QQ:35723195 MSN:xb@live.it
'| 日    期 | 2009-04-26 23:23:21
'| 修    订 | 2010-08-22 19:56:45
'| 版    本 | 1.0.0
'===================================================================================================
'| 修    订 | 2011-10-24
'| 修 订 人 | 菜鸟学飞 http://www.vbgood.com/space-uid-144949.html
'| 说    明 | 对所有开源的人表示感谢
'===================================================================================================

Option Explicit
Private Type UUID
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(7) As Byte
End Type

Private Declare Function CreateThread Lib "kernel32" (ByVal lpThreadA As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long

'强制结束线程用TerminateThread
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
'| 函 数 名 | VBCreateThread
'| 说    明 | 创建一个线程(Create a thread)
'| 参    数 | lpStartAddress:   线程函数地址(thread function address )
'| 参    数 | lpParam:          线程参数(thread param)
'| 参    数 | lpThreadId:       线程ID(tid)
'| 返 回 值 | 返回线程句柄(hThread)
'===================================================================================================
Public Function VBCreateThread(ByVal lpStartAddress As Long, ByVal lpParam As Long, Optional ByRef lpThreadId As Long = 0) As Long
        Dim lpData As Long
        Dim nRet As Long
        
        '分配八字节用户装参数
        lpData = GlobalAlloc(0, 8)
        PutMem4 lpData, lpStartAddress
        PutMem4 lpData + 4, lpParam
        
        VBCreateThread = CreateThread(0, 0, AddressOf ThreadHead, ByVal lpData, 0, lpThreadId)
End Function

'===================================================================================================
'| 函 数 名 | InitVB
'| 说    明 | 初始化VB运行库(Init vb runtime)
'| 参    数 | 无 (void)
'| 返 回 值 | 无 (void)

'注意！！！！！！！！！
'为了防止错误 必须将程序入口设置为 sub main (工程属性 > 启动对象 设置为 sub main)
'因为多线程时候会重复调用 sub main或者fromload

'注意！！！！！！
'InitVB函数会对sub main做处理 InitVB会删除 sub main 的代码
'所以 如果你要重复多次调用 sub main 请另写代码
'===================================================================================================
Public Function InitVB() As Long
        Dim fake As Long
        Dim lvb As Long
        Dim riid As UUID
        Dim aiid As UUID
        Dim ofac As Object
        Dim nRet As Long
        
        '基本初始化
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
        '6610EE80 置0
        '6610EE84 置0
        
        '尼玛~~  这样子弄后可以动态加载卸载 但是TMD进程结束的时候又有问题了
        '        GetMem4 GetProcAddress(GetModuleHandle("msvbvm60.dll"), "SetMemEvent") + 14, nRet
        '        PutMem4 nRet, 0
        '        PutMem4 nRet + 4, 0

        If RunOneFlag = 0 Then
                GetMem4 AddressOf CallThreadFunc, nRet
                '给CallThreadFunc函数 动态写入汇编
                If nRet <> &HFF505B58 Then
                        '不让 sub main运行多次
                        WriteProcessMemory -1, AddressOf Main, &HC3C3C3C3, 4, ByVal 0
                        '给CallThreadFunc函数 动态写入汇编
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


'得到当前DLL句柄
Public Function GetCurrentModule() As Long
        Dim mbi As MEMORY_BASIC_INFORMATION
        
        '随便取一个函数的地址
        If (VirtualQuery(AddressOf InitVB, mbi, LenB(mbi)) <> 0) Then
                GetCurrentModule = mbi.AllocationBase
        Else
                GetCurrentModule = 0
        End If
End Function

'tgy原创 这里略作修改
'取VB头,全新的取VB头方法,速度比OPEN文件快得多 用这种方法的程序应该可以加一些压缩壳
Public Function GetFakeH(ByVal hin As Long) As Long
        Dim lPtr     As Long
        Dim isvb As String
        Dim mdat(4095) As Byte
        
        lPtr = hin + 4096 '文件头+4KB 加快速度
        isvb = StrConv("VB5!", vbFromUnicode)

        CopyMemory mdat(0), ByVal lPtr, 4096
        
        GetFakeH = InStrB(mdat, isvb)
        
        If GetFakeH = 0 Then
                Exit Function
        End If
       GetFakeH = GetFakeH + lPtr - 1
End Function




Private Function CallThreadFunc(ByVal UserFuncAddr As Long, ByVal lpParam As Long) As Long
'写几句句用于被覆盖的垃圾代码 长度必须足够容纳下面的汇编
MsgBox "YES"
MsgBox "YES"
MsgBox "YES"
MsgBox "YES"

'动态写入的汇编 手动调整堆栈 手动call
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

'线程初始化 释放资源的事情由 ThreadHead 搞定
'参数是传的指针 相当于VB的按地址传输
Private Function ThreadHead(ByRef lpParam As THREAD_DATA) As Long

        '线程初始化
        Call InitVB
        
        On Error Resume Next
        
        '调用用户线程函数
        ThreadHead = CallThreadFunc(lpParam.lpStartAddress, lpParam.lpParam)
        
        '注意 lpParam是动态分配的 需要释放
        GlobalFree lpParam
        CoUninitialize
        
        '可以在这这里添加线程结束事件处理
End Function

Private Function Getval(ByVal n As Long) As Long
Getval = n
End Function





