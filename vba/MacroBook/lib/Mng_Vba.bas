Attribute VB_Name = "Mng_Vba"
Option Explicit
 
' vba manage library v1.0
 
Private Type T_INPUT_TYPE
    lType As Long
    bytXi(0 To 23) As Byte
End Type
 
Private Type T_KEY_BD_INPUT
    iVk As Integer
    iScan As Integer
    lFlags As Long
    lTime As Long
    lExtraInfo As Long
End Type
 
Declare Function SendInput Lib "user32" (ByVal nInputs As Long, pInputs As T_INPUT_TYPE, ByVal cbSize As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
 
Private Const INPUT_KEYBOARD = 1
Private Const KEYEVENTF_KEYDOWN = 0
Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_CONTROL = &H11
 
Public Function ClearImmidiateWindow()
    Dim atInputEvents(0 To 7) As T_INPUT_TYPE
    Dim tKeyEvent As T_KEY_BD_INPUT
    Dim vKeyNameArray As Variant
    Dim iKeyNameIdx As Integer
 
    'Set key name to send
    vKeyNameArray = Array(VK_CONTROL, vbKeyG, vbKeyA, vbKeyDelete)
 
    For iKeyNameIdx = 0 To UBound(vKeyNameArray)
        tKeyEvent.iVk = vKeyNameArray(iKeyNameIdx) 'key name
        tKeyEvent.iScan = 0
        tKeyEvent.lFlags = KEYEVENTF_KEYDOWN
        tKeyEvent.lTime = 0
        tKeyEvent.lExtraInfo = 0
        atInputEvents(iKeyNameIdx).lType = INPUT_KEYBOARD
        CopyMemory atInputEvents(iKeyNameIdx).bytXi(0), tKeyEvent, Len(tKeyEvent)
    Next
 
    For iKeyNameIdx = 0 To UBound(vKeyNameArray)
        tKeyEvent.iVk = vKeyNameArray(iKeyNameIdx)
        tKeyEvent.iScan = 0
        tKeyEvent.lFlags = KEYEVENTF_KEYUP 'release the key
        tKeyEvent.lTime = 0
        tKeyEvent.lExtraInfo = 0
        atInputEvents(iKeyNameIdx + UBound(vKeyNameArray) + 1).lType = INPUT_KEYBOARD
        CopyMemory atInputEvents(iKeyNameIdx + UBound(vKeyNameArray) + 1).bytXi(0), tKeyEvent, Len(tKeyEvent)
    Next
 
    'place the events into the stream
    SendInput iKeyNameIdx + UBound(vKeyNameArray) + 1, atInputEvents(0), Len(atInputEvents(0))
 
End Function
