Attribute VB_Name = "LibMemory"
'@Folder "SQLiteDBdev.MemoryTools"
'@IgnoreModule: This is part of the https://github.com/cristianbuse/VBA-MemoryTools
'''=============================================================================
''' VBA MemoryTools
''' -----------------------------------------------
''' https://github.com/cristianbuse/VBA-MemoryTools
''' -----------------------------------------------
''' MIT License
'''
''' Copyright (c) 2020 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================

Option Explicit
Option Private Module

'*******************************************************************************
'' Methods in this library module allow direct native memory manipulation in VBA
'' regardless of:
''  - the host Application (Excel, Word, AutoCAD etc.)
''  - the operating system (Mac, Windows)
''  - application environment (x32, x64)
'' A single API call to RtlMoveMemory/MemMove is needed to start the remote
''  referencing mechanism. Inside a REMOTE_MEMORY type, a 'remoteVT' Variant is
''  used to manipulate the VarType of the 'memValue' that we want to read/write.
''  The remote manipulation of the VarType is done by setting the VT_BYREF flag
''  on the 'remoteVT' Variant. This is done by using a CopyMemory API but only
''  once per initialization of the REMOTE_MEMORY variable (see MemIntAPI). Note
''  that besides the Static REMOTE_MEMORY used in MemIntAPI all the other memory
''  structs are initialized through InitRemoteMemory with no API calls. Once the
''  flag is set the 'remoteVT' is used to change the VarType of the first
''  Variant just by using a native VBA assignment operation (needs a utility
''  method for correct redirection). In order for the 'memValue' Variant to
''  point to aspecific memory address, 2 steps are needed:
''   1) the required address is assigned to the 'memValue' Variant
''   2) the VarType of the 'memValue' Variant is remotely changed via the
''      'remoteVT' Variant (through LetByRefVT) while making sure the VT_BYREF
''      flag is also set
''
'' Note that the 'LinkMem' method could have been a function returning a
''  REMOTE_MEMORY type and thus greatly improving readability but this approach
''  proved to be 2x slower in testing
'*******************************************************************************

'Used for raising errors
Private Const MODULE_NAME As String = "LibMemory"

#If Mac Then
    #If VBA7 Then
        Public Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr
    #Else
        Public Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As Long) As Long
    #End If
#Else 'Windows
    'https://msdn.microsoft.com/en-us/library/mt723419(v=vs.85).aspx
    #If VBA7 Then
        Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    #Else
        Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    #End If
#End If

'The size in bytes of a memory address
#If Win64 Then
    Public Const PTR_SIZE As Long = 8
#Else
    Public Const PTR_SIZE As Long = 4
#End If

#If Win64 Then
    #If Mac Then
        Public Const vbLongLong As Long = 20 'Apparently missing for x64 on Mac
    #End If
    Public Const vbLongPtr As Long = vbLongLong
#Else
    Public Const vbLongPtr As Long = vbLong
#End If

Public Type REMOTE_MEMORY
    memValue As Variant
    remoteVT As Variant
    isInitialized As Boolean 'In case state is lost
End Type

'https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/3fe7db9f-5803-4dc4-9d14-5425d3f5461f
'https://docs.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-variant?redirectedfrom=MSDN
'Flag used to simulate ByRef Variants
Public Const VT_BYREF As Long = &H4000

Private m_remoteMemory As REMOTE_MEMORY

'*******************************************************************************
'Returns an initialized (linked) REMOTE_MEMORY struct
'Links .remoteVt to the first 2 bytes of .memValue
'*******************************************************************************
Public Sub InitRemoteMemory(ByRef rm As REMOTE_MEMORY)
    rm.remoteVT = VarPtr(rm.memValue)
    MemIntAPI(VarPtr(rm.remoteVT)) = vbInteger + VT_BYREF
    rm.isInitialized = True
End Sub

'*******************************************************************************
'The only method in this module that uses CopyMemory!
'Assures that InitRemoteMemory can link the Var Type for new structs
'*******************************************************************************
#If Win64 Then
Private Property Let MemIntAPI(ByVal memAddress As LongLong, ByVal newValue As Integer)
#Else
Private Property Let MemIntAPI(ByVal memAddress As Long, ByVal newValue As Integer)
#End If
    Static rm As REMOTE_MEMORY
    If Not rm.isInitialized Then 'Link .remoteVt to .memValue's first 2 bytes
        rm.remoteVT = VarPtr(rm.memValue)
        CopyMemory rm.remoteVT, vbInteger + VT_BYREF, 2
        rm.isInitialized = True
    End If
    rm.memValue = memAddress
    LetByRefVT(rm.remoteVT) = vbInteger + VT_BYREF
    '
    LetByRefInt(rm.memValue) = newValue
    rm.memValue = Empty
End Property

'*******************************************************************************
'Read/Write a Byte from/to memory
'*******************************************************************************
#If Win64 Then
Public Property Get MemByte(ByVal memAddress As LongLong) As Byte
#Else
Public Property Get MemByte(ByVal memAddress As Long) As Byte
#End If
    #If Mac Then
        CopyMemory MemByte, ByVal memAddress, 1
    #Else
        LinkMem m_remoteMemory, memAddress, vbByte + VT_BYREF
        MemByte = m_remoteMemory.memValue
        m_remoteMemory.memValue = Empty
    #End If
End Property
#If Win64 Then
Public Property Let MemByte(ByVal memAddress As LongLong, ByVal newValue As Byte)
#Else
Public Property Let MemByte(ByVal memAddress As Long, ByVal newValue As Byte)
#End If
    #If Mac Then
        CopyMemory ByVal memAddress, newValue, 1
    #Else
        LinkMem m_remoteMemory, memAddress, vbByte + VT_BYREF
        LetByRefByte(m_remoteMemory.memValue) = newValue
        m_remoteMemory.memValue = Empty
    #End If
End Property

'*******************************************************************************
'Read/Write 2 Bytes (Integer) from/to memory
'*******************************************************************************
#If Win64 Then
Public Property Get MemInt(ByVal memAddress As LongLong) As Integer
#Else
Public Property Get MemInt(ByVal memAddress As Long) As Integer
#End If
    #If Mac Then
        CopyMemory MemInt, ByVal memAddress, 2
    #Else
        LinkMem m_remoteMemory, memAddress, vbInteger + VT_BYREF
        MemInt = m_remoteMemory.memValue
        m_remoteMemory.memValue = Empty
    #End If
End Property

#If Win64 Then
Public Property Let MemInt(ByVal memAddress As LongLong, ByVal newValue As Integer)
#Else
Public Property Let MemInt(ByVal memAddress As Long, ByVal newValue As Integer)
#End If
    #If Mac Then
        CopyMemory ByVal memAddress, newValue, 2
    #Else
        LinkMem m_remoteMemory, memAddress, vbInteger + VT_BYREF
        LetByRefInt(m_remoteMemory.memValue) = newValue
        m_remoteMemory.memValue = Empty
    #End If
End Property

'*******************************************************************************
'Read/Write 4 Bytes (Long) from/to memory
'*******************************************************************************
#If Win64 Then
Public Property Get MemLong(ByVal memAddress As LongLong) As Long
#Else
Public Property Get MemLong(ByVal memAddress As Long) As Long
#End If
    #If Mac Then
        CopyMemory MemLong, ByVal memAddress, 4
    #Else
        LinkMem m_remoteMemory, memAddress, vbLong + VT_BYREF
        MemLong = m_remoteMemory.memValue
        m_remoteMemory.memValue = Empty
    #End If
End Property
#If Win64 Then
Public Property Let MemLong(ByVal memAddress As LongLong, ByVal newValue As Long)
#Else
Public Property Let MemLong(ByVal memAddress As Long, ByVal newValue As Long)
#End If
    #If Mac Then
        CopyMemory ByVal memAddress, newValue, 4
    #Else
        LinkMem m_remoteMemory, memAddress, vbLong + VT_BYREF
        LetByRefLong(m_remoteMemory.memValue) = newValue
        m_remoteMemory.memValue = Empty
    #End If
End Property

'*******************************************************************************
'Read/Write 8 Bytes (LongLong) from/to memory
'*******************************************************************************
#If Win64 Then
Public Property Get MemLongLong(ByVal memAddress As LongLong) As LongLong
    #If Mac Then
        CopyMemory MemLongLong, ByVal memAddress, 8
    #Else
        LinkMem m_remoteMemory, memAddress, vbLongLong + VT_BYREF
        MemLongLong = m_remoteMemory.memValue
        m_remoteMemory.memValue = Empty
    #End If
End Property
Public Property Let MemLongLong(ByVal memAddress As LongLong, ByVal newValue As LongLong)
    #If Mac Then
        CopyMemory ByVal memAddress, newValue, 8
    #Else
        With m_remoteMemory
            If Not .isInitialized Then InitRemoteMemory m_remoteMemory
            '
            'Cannot set Variant/LongLong ByRef so we use Currency instead
            'Read LongLong as Currency
            Dim tempCurrency As Currency
            .memValue = newValue
            LetByRefVT(.remoteVT) = vbCurrency
            tempCurrency = .memValue
            '
            'Link to memory
            .memValue = memAddress
            LetByRefVT(.remoteVT) = vbCurrency + VT_BYREF
            '
            'Set new value
            LetByRefCurr(.memValue) = tempCurrency
            .memValue = Empty
        End With
    #End If
End Property
#End If

'*******************************************************************************
'Read/Write 4 Bytes (Long on x32) or 8 Bytes (LongLong on x64) from/to memory
'*******************************************************************************
#If Win64 Then
Public Property Get MemLongPtr(ByVal memAddress As LongLong) As LongLong
    MemLongPtr = MemLongLong(memAddress)
End Property
Public Property Let MemLongPtr(ByVal memAddress As LongLong, ByVal newValue As LongLong)
    MemLongLong(memAddress) = newValue
End Property
#Else
Public Property Get MemLongPtr(ByVal memAddress As Long) As Long
    MemLongPtr = MemLong(memAddress)
End Property
Public Property Let MemLongPtr(ByVal memAddress As Long, ByVal newValue As Long)
    MemLong(memAddress) = newValue
End Property
#End If

'*******************************************************************************
'Dereference an object by it's pointer
'*******************************************************************************
#If Win64 Then
Public Function MemObject(ByVal memAddress As LongLong) As Object
#Else
Public Function MemObject(ByVal memAddress As Long) As Object
#End If
    If memAddress = 0 Then Exit Function
    '
    #If Mac Then
        Dim Obj As Object
        CopyMemory Obj, memAddress, PTR_SIZE
        Set MemObject = Obj
        memAddress = 0 'We don't just use 0 (below) because we need 0& or 0^
        CopyMemory Obj, memAddress, PTR_SIZE
    #Else
        If Not m_remoteMemory.isInitialized Then InitRemoteMemory m_remoteMemory
        m_remoteMemory.memValue = memAddress
        Set MemObject = RemObject(m_remoteMemory.remoteVT, m_remoteMemory)
    #End If
End Function
'Utility - avoids extra stack frames by not calling LetByRefVT twice
Private Function RemObject(ByRef vt As Variant, ByRef rm As REMOTE_MEMORY) As Object
    vt = vbObject
    Set RemObject = rm.memValue
    vt = vbLongPtr
End Function

'*******************************************************************************
'Redirects the REMOTE_MEMORY.memValue Variant to the new memory address
'*******************************************************************************
#If Win64 Then
Private Sub LinkMem(ByRef rm As REMOTE_MEMORY, ByRef memAddr As LongLong, ByRef vt As VbVarType)
#Else
Private Sub LinkMem(ByRef rm As REMOTE_MEMORY, ByRef memAddr As Long, ByRef vt As VbVarType)
#End If
    If Not rm.isInitialized Then InitRemoteMemory rm
    rm.memValue = memAddr
    LetByRefVT(rm.remoteVT) = vt
End Sub

'*******************************************************************************
'Utilities for updating remote values that have the VT_BYREF flag set
'*******************************************************************************
Private Property Let LetByRefVT(ByRef V As Variant, ByRef vt As VbVarType)
    V = vt
End Property
Private Property Let LetByRefByte(ByRef V As Variant, ByRef newValue As Byte)
    V = newValue
End Property
Private Property Let LetByRefInt(ByRef V As Variant, ByRef newValue As Integer)
    V = newValue
End Property
Private Property Let LetByRefLong(ByRef V As Variant, ByRef newValue As Long)
    V = newValue
End Property
Private Property Let LetByRefCurr(ByRef V As Variant, ByRef newValue As Currency)
    V = newValue
End Property

'Method purpose explanation at:
'https://gist.github.com/cristianbuse/b9cc79164c1d31fdb30465f503ac36a9
'
'Practical note Jan-2021 from Vladimir Vissoultchev (https://github.com/wqweto):
'This is mostly not needed in client application code even for LARGEADDRESSAWARE
'   32-bit processes nowadays as a reliable technique to prevent pointer
'   arithmetic overflows is to VirtualAlloc a 64KB sentinel chunk around 2GB
'   boundary at application start up so that the boundary is never (rarely)
'   crossed in normal pointer operations.
'This same sentinel chunk fixes native PropertyBag as well which has troubles
'   when internal storage crosses 2GB boundary.
#If Win64 Then
Public Function UnsignedAdd(ByVal unsignedPtr As LongLong, ByVal signedOffset As LongLong) As LongLong
    UnsignedAdd = ((unsignedPtr Xor &H8000000000000000^) + signedOffset) Xor &H8000000000000000^
End Function
#Else
Public Function UnsignedAdd(ByVal unsignedPtr As Long, ByVal signedOffset As Long) As Long
    UnsignedAdd = ((unsignedPtr Xor &H80000000) + signedOffset) Xor &H80000000
End Function
#End If

'*******************************************************************************
'Redirects the instance of a class to another instance of the same class within
'   the scope of a private class Function (not  Sub) where the call happens.
'
'Warning! ONLY call this method from a Private Function of a class!
'
'vbArray + vbString Function return type is not supported. It would be
'   possible to find the correct address by reading memory in a loop but there
'   would be no checking available
'*******************************************************************************
#If Win64 Then
Public Sub RedirectInstance(ByVal funcReturnPtr As LongLong _
                          , ByVal currentInstance As Object _
                          , ByVal targetInstance As Object)
#Else
Public Sub RedirectInstance(ByVal funcReturnPtr As Long _
                          , ByVal currentInstance As Object _
                          , ByVal targetInstance As Object)
#End If
    Const methodName As String = "RedirectInstance"
    #If Win64 Then
        Dim originalPtr As LongLong
        Dim newPtr As LongLong
        Dim swapAddress As LongLong
    #Else
        Dim originalPtr As Long
        Dim newPtr As Long
        Dim swapAddress As Long
    #End If
    '
    originalPtr = ObjPtr(GetDefaultInterface(currentInstance))
    newPtr = ObjPtr(GetDefaultInterface(targetInstance))
    '
    'Validate Input
    If currentInstance Is Nothing Or targetInstance Is Nothing Then
        Err.Raise 91, methodName, "Object not set"
    ElseIf MemLongPtr(originalPtr) <> MemLongPtr(newPtr) Then 'Faster to compare vTables than to compare TypeName(s)
        Err.Raise 5, methodName, "Expected same VB class"
    ElseIf funcReturnPtr = 0 Then
        Err.Raise 5, methodName, "Missing Function Return Pointer"
    End If
    '
    'On x64 the shadow stack space is allocated next to the Function Return
    'On x32 the stack space has a fixed offset (found through testing)
    #If Win64 Then
        Const memOffsetNonVariant As LongLong = PTR_SIZE
        Const memOffsetVariant As LongLong = PTR_SIZE * 3
    #Else
        Const memOffsetNonVariant As Long = PTR_SIZE * 28
        Const memOffsetVariant As Long = PTR_SIZE * 31
    #End If
    '
    swapAddress = FindSwapAddress(funcReturnPtr, memOffsetNonVariant, originalPtr)
    If swapAddress = 0 Then
        swapAddress = FindSwapAddress(funcReturnPtr, memOffsetVariant, originalPtr)
        If swapAddress = 0 Then
            Err.Raise 5, methodName, "Invalid input or not called " _
            & "from class Function or vbArray + vbString function return type"
        End If
    End If
    '
    'Redirect Instance
    MemLongPtr(swapAddress) = newPtr
End Sub

'*******************************************************************************
'Finds the swap address (address of the instance pointer on the stack)
'*******************************************************************************
#If Win64 Then
Private Function FindSwapAddress(ByVal funcReturnPtr As LongLong _
                               , ByVal memOffset As LongLong _
                               , ByVal originalPtr As LongLong) As LongLong
    Dim swapAddr As LongLong: swapAddr = funcReturnPtr + memOffset
    '
    'Adjust alignment for Boolean/Byte/Integer/Long function return type
    'Needed on #Mac but not on #Win (at least not found in testing)
    'Safer to have for #Win as well
    swapAddr = swapAddr - (swapAddr Mod PTR_SIZE)
    '
    If MemLongLong(swapAddr) = originalPtr Then
        FindSwapAddress = swapAddr
    End If
End Function
#Else
Private Function FindSwapAddress(ByVal funcReturnPtr As Long _
                               , ByVal memOffset As Long _
                               , ByVal originalPtr As Long) As Long
    Dim startAddr As Long: startAddr = funcReturnPtr + memOffset
    Dim swapAddr As Long
    '
    'Adjust memory alignment for Boolean/Byte/Integer function return type
    startAddr = startAddr - (startAddr Mod PTR_SIZE)
    '
    swapAddr = GetSwapIndirectAddress(startAddr)
    If swapAddr = 0 Then
        'Adjust mem alignment for Currency/Date/Double function return type
        swapAddr = GetSwapIndirectAddress(startAddr + PTR_SIZE)
        If swapAddr = 0 Then Exit Function
    End If
    If MemLongPtr(swapAddr) = originalPtr Then
        FindSwapAddress = swapAddr
    End If
End Function
Private Function GetSwapIndirectAddress(ByVal startAddr As Long) As Long
    Const maxOffset As Long = PTR_SIZE * 100
    Dim swapAddr As Long: swapAddr = MemLong(startAddr) + PTR_SIZE * 2
    '
    'Check if the address is within acceptable limits. The address
    '   of the instance pointer (within the stack frame) cannot be too far
    '   from the function return address (first offsetted to startAddr)
    If startAddr < swapAddr And swapAddr - startAddr < maxOffset * PTR_SIZE Then
        GetSwapIndirectAddress = swapAddr
    End If
End Function
#End If

'*******************************************************************************
'Returns the default interface for an object
'Casting from IUnknown to IDispatch (Object) forces a call to QueryInterface for
'   the IDispatch interface (which knows about the default interface)
'*******************************************************************************
Public Function GetDefaultInterface(ByVal Obj As IUnknown) As Object
    Set GetDefaultInterface = Obj
End Function

'*******************************************************************************
'Returns the memory address of a variable of array type
'Returns error 5 for a non-array or an array wrapped in a Variant
'*******************************************************************************
#If Win64 Then
Public Function VarPtrArray(ByRef Arr As Variant) As LongLong
#Else
Public Function VarPtrArray(ByRef Arr As Variant) As Long
#End If
    Dim vt As VbVarType: vt = MemInt(VarPtr(Arr)) 'VarType(arr) ignores VT_BYREF
    If vt And (vbArray Or VT_BYREF) Then
        Const pArrayOffset As Long = 8
        VarPtrArray = MemLongPtr(VarPtr(Arr) + pArrayOffset)
    Else
        Err.Raise 5, "VarPtrArray", "Array required"
    End If
End Function

'*******************************************************************************
'Returns the pointer to the underlying SAFEARRAY structure of a VB array
'Returns error 5 for a non-array
'*******************************************************************************
#If Win64 Then
Public Function ArrPtr(ByRef Arr As Variant) As LongLong
#Else
Public Function ArrPtr(ByRef Arr As Variant) As Long
#End If
    Dim vt As VbVarType: vt = MemInt(VarPtr(Arr)) 'VarType(arr) ignores VT_BYREF
    If vt And vbArray Then
        Const pArrayOffset As Long = 8
        ArrPtr = MemLongPtr(VarPtr(Arr) + pArrayOffset)
        If vt And VT_BYREF Then ArrPtr = MemLongPtr(ArrPtr)
    Else
        Err.Raise 5, "ArrPtr", "Array required"
    End If
End Function


