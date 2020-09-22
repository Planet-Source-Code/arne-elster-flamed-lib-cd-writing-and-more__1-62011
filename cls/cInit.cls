VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Acronym: UEH = Unhandled Exception Handler

' Separate class to act as a Lib_Initialize & Lib_Terminate event
' when the DLL is first started/referenced. And IF app closes gracefully
' then this class will terminate & basically clean up for the entire DLL

' Easiest way to really know if your dll terminated cleanly is to
' set up some system that fires a terminate event. Therefore,
' this class is used mainly for that purpose.

'This API function installs your custom exception handler.
Private Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long

' API to determine if the client is compiled or not
Private Declare Function GetModuleFileName Lib "kernel32.dll" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

' All VB projects (compiled or not) have a hidden owner window. This can be proven
' by using Spy++. This is a great place to cache a property or two that you want
' available throughout your application & in the case of this DLL, the property
' remains available in IDE even if the End command/button executed. That is because
' the owner window doesn't close in IDE until IDE closes.

Private Declare Function GetWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER As Long = 4

' We aren't required to reset the original UEH, we can simply reset
' it to 0 (default handler) when we no longer need it. However, good coding is to
' reset what you set. So we will take the extra effort and cache, then reset the
' previous UEH. This property will be used on hidden owner window to help us
Private Const lvp_GPFhandle As String = "lvpGPF"

' class properties
Private isIDE_Client As Boolean     ' is the client compiled
Private classRef() As Long          ' array of lvSubclass pointers managed by DLL
Private classCount As Long          ' number of array pointers
Private m_OwnerHwnd As Long         ' hidden owner window's handle
Private m_AwareLevel As Long        ' GPF monitoring (0=none;1=IDE;2=Compiled;3=Both)
Private m_CallBackIde As Long       ' this is the callback supplied by user while in IDE
Private m_CallBackCmp As Long       ' this is the callback supplied by user when compiled
Private isEHptrSet As Boolean       ' primarily for compiled apps; allows resetting UEH to zero
Private bHideMessages As Boolean    ' allows/prevents messageboxes when GPF detected

Private Function AmInIDE() As Boolean
' prefer this method of determining if this project is in IDE
' Other methods throw an error, then trap it.

' How it works? It is called by a Debug.Assert command & debug statements
' are stripped from a compiled app. Therefore if compiled, this routine is never called
isIDE_Client = True
AmInIDE = True
End Function

Private Sub Class_Initialize()
' since the GPF protection reacts differently when project is compiled
' or not (in Debug mode as far as they are concerned), or whether or not
' the GPF callback is in an outside DLL, we need to check if DLL & Client are in IDE

' see if our dll is in IDE or compiled
Debug.Assert AmInIDE()

If Not isIDE_Client Then
    
    ' see if our client is in IDE or compiled
    ' This was stripped from Luke's (aka SelfTaught) implementation of Paul Caton's
    ' ASM subclassing thunks. Modified to check for different version of VB & to
    ' prevent false positives. Luke's routine looked for filename Like *vb6.exe
    ' which would return a compiled app titled My1stApp_vb6.exe as in IDE
    
    Dim ls As String
    Dim liLen As Long
    Dim vbVer As String
    
    ' get the executable that loaded this DLL
    ls = Space$(255)
    liLen = GetModuleFileName(0&, ls, Len(ls))
    
    ' check for 3 versions of VB; last one is a hope ;)
    If liLen > 7 Then
                
        ' extract the last 8 characters (i.e., vb6.exe, vb5.exe)
        vbVer = Right$(LCase$(Left$(ls, liLen)), 8)
        
        ' now see if executable is one of the versions
        If InStr("*\vb6.exe*/vb6.exe*:vb6.exe", vbVer) Then
            isIDE_Client = True
        ElseIf InStr("*\vb5.exe*/vb5.exe*:vb5.exe", vbVer) Then
            isIDE_Client = True
        Else
            isIDE_Client = (InStr("*\vb7.exe*/vb7.exe*:vb7.exe", vbVer) > 0)
        End If
        
    End If
    
End If
End Sub

Private Sub Class_Terminate()
' clean up if not already done
ResetGPFhandle          ' reset back to previous UEH
ReleaseGPFReferences    ' unsubclass any remaining windows
End Sub

Friend Sub UpdateClassRef(bRemove As Boolean, classPtr As Long)

' whenever a class first subclasses a window or whenever a class terminates
' this routine is called so we can keep track of which classes are
' active. This info needed to auto-unsubclass all windows this dll is managing
' and used in the event of encountering a GPF or user hits the End button

' this portion of the class is only applicable when client is in IDE
If isIDE_Client = False Then Exit Sub

Dim classNr As Long
' loop through our small array
For classNr = 0 To classCount - 1
    
    ' see if we logged that class already
    If classPtr = classRef(classNr) Then
        
        ' yep, is it being destroyed/terminated?
        If bRemove Then
            
            ' yep, update our array & count, then exit routine
            If classCount = 1 Then
                Erase classRef()
            Else
                classRef(classNr) = classRef(classCount - 1)
                ReDim Preserve classRef(0 To classCount - 2)
            End If
            classCount = classCount - 1
        End If
        ' we are already managing it & if not being removed, exit anyway
        ' no need to double log the class
        Exit Sub
    End If
    
Next

' if we got here then we have not logged the class
' Let's log it now if it is being initialized
If Not bRemove Then
    ' add it to end of array & update ref count
    ReDim Preserve classRef(0 To classCount)
    classRef(classCount) = classPtr
    classCount = classCount + 1
End If
End Sub

Friend Property Get AwarenessLevel() As Long
' used to supply lvSubclasser.GPF_Protect property with value
Dim lRtn As Long
AwarenessLevel = m_AwareLevel
If (m_AwareLevel Or 2) = m_AwareLevel Then
    ' Special case
    ' user wants GPF protection in compiled app, but not applicable
    ' if user didn't supply a callback address
    If m_CallBackIde = 0 Then
        AwarenessLevel = m_AwareLevel And Not 2
        Exit Property
    End If
End If
End Property

Friend Sub SetPrjOwner(hWnd As Long, ByVal GPFawareness As Long)
' This is called each time a window is first subclassed or
' anytime lvSubclasser.GPF_Protection property is set

' What we are really trying to do is ensure we have the hidden owner window
' And that we cache the value of the current UEH pointer on that window

Dim prevOwner As Long, gpfHandle As Long
If m_OwnerHwnd = 0 And hWnd <> 0 Then
    ' first time DLL is subclassing
    
    ' let's find the hidden owner window
    prevOwner = GetWindow(hWnd, GW_OWNER)
    Do While prevOwner
        m_OwnerHwnd = prevOwner
        prevOwner = GetWindow(m_OwnerHwnd, GW_OWNER)
    Loop
End If
If m_OwnerHwnd = 0 Then Exit Sub
' sanity check
If GPFawareness < 0 Then GPFawareness = 0
If GPFawareness > 3 Then GPFawareness = 3

' got owner, now let's see if we already cached the previous UEH
' This could happen if user hit the End button or executed End statement
' which prevented the DLL from properly cleaning up (can only happen in IDE)
gpfHandle = GetProp(m_OwnerHwnd, lvp_GPFhandle)

Dim bSet As Boolean
' determine if we need to set a new UEH trap

If GPFawareness = gpfNoErrorTraps Or (m_AwareLevel > 1 And m_CallBackCmp = 0 And isIDE_Client = False) Then
    ' remove error traps - brave aren't you
    ResetGPFhandle
Else
    ' see if gpf IDE enabled
    If (GPFawareness Or gpfIDE) = GPFawareness Then bSet = isIDE_Client
    ' now see if gpf compiled enabled, but may be ignored.
        ' See AwarenessLevel property above as to why
    If Not bSet Then bSet = (isIDE_Client = False And GPFawareness > 1 And m_CallBackCmp <> 0)

End If

If bSet Then
    
    ' we are going to establish a GPF error trap
    If gpfHandle = 0 And isEHptrSet = False Then ' not previously cached, so lets's set
        ' the trap & cache the previous UEH pointer
        gpfHandle = SetUnhandledExceptionFilter(AddressOf lvErrorChecker)
        ' for compiled exe this will be zero unless a previous UEH was in use
        SetProp m_OwnerHwnd, lvp_GPFhandle, gpfHandle
    Else    ' previously cached on owner, no caching - but set again just in
            ' case some other UEH was put into play from other user code
        Call SetUnhandledExceptionFilter(AddressOf lvErrorChecker)
    End If
    isEHptrSet = True

End If

' update the awareness level
m_AwareLevel = GPFawareness
End Sub

' class properties
Friend Property Get ClientIde() As Boolean
ClientIde = isIDE_Client    ' is our client compiled?
End Property

Friend Property Get Callback(forIDE As Boolean) As Long
If forIDE Then
    Callback = m_CallBackIde
Else
    Callback = m_CallBackCmp       ' do we have a callback to use when client is compiled
End If
End Property
Friend Property Let Callback(forIDE As Boolean, cAddress As Long)
' set the callback address for IDE and/or compiled
If forIDE Then
    m_CallBackIde = cAddress
Else
    m_CallBackCmp = cAddress       ' set the callback
    If cAddress = 0 And isIDE_Client = False Then
        ResetGPFhandle ' don't allow GPF trapping if no callback address provided
    Else
        ' activate if needed
        If m_AwareLevel > 1 And isIDE_Client = False Then SetPrjOwner 0, m_AwareLevel
    End If
End If
End Property

Friend Property Let HideMessages(bHide As Boolean)

bHideMessages = bHide
' This property can be used if displaying a message box inadvertently
' causes the GPF handling and your On Error statements to enter an infinite loop.

' Example: Let's say in your subclassed event, you trigger a GPF when your
' window gets focus. Well when a GPF message box appears & you subsequently close
' it, your window gets focus back & then triggers another GPF (infinite loop)

End Property
Friend Property Get HideMessages() As Boolean
HideMessages = bHideMessages
End Property


Friend Sub ResetGPFhandle()

' replace original UEH pointer if available
If isEHptrSet = True Then
    ' we cached the previous UEH
    
    ' Interesting enough, when in IDE your project has a UEH provided by the IDE
    ' But when compiled, none... That's why a Division By Zero causes you a GPF
    ' when your app is compiled -- hence just more reasons for using On Error statements
    
    SetUnhandledExceptionFilter GetProp(m_OwnerHwnd, lvp_GPFhandle)
    m_AwareLevel = 0
    m_CallBackIde = 0
    m_CallBackCmp = 0
    
End If
End Sub

Friend Sub ReleaseGPFReferences()

' applies only to IDE; never auto-unsubclass compiled apps
If isIDE_Client = False Then Exit Sub

' We are about to crash unless the DLL can prevent it.
' Unsubclass to help prevent that crash. Even though the error trapping
' has potential to allow code to continue; don't take chances

' This routine will actually unsubclass every window that was
' subclassed by this entire DLL (this thread only)

Dim classNr As Long
Dim sClass As lvSubclasser
Dim hWnd As Long

On Error Resume Next

For classNr = 0 To classCount - 1
    ' for each class we are managing....
    hWnd = 0
    CopyMemory sClass, classRef(classNr), &H4
    
    ' remove any system tray icons
    ' now unsubclass all windows subclassed by this lvSubclasser instance
    hWnd = sClass.LastSubclassed
    
    ' finish
    CopyMemory sClass, 0&, &H4
    RemoveClassSubClassing hWnd
Next

ExitTermination:

' clean up
Erase classRef
classCount = 0
End Sub
