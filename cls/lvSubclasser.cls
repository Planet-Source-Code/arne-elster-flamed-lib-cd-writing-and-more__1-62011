VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "lvSubclasser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Want to check & see if any updates posted?
' http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=59434&lngWId=1

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Feel free to use this for your own purposes. I have only 2 requests in return

' 1) If you make any changes to any executable statements or remove or append
'   executable statements, please specifically annotate what modifications you
'   have made and where those modifications can be found.  In addition should
'   you have made those modifications, please remove any references to LaVolpe
'   from the DLL properties window, to include the project description, company
'   name, file description, comments and legal trademarks. Also request you
'   remove all statements like the remarks at the very top of this module that
'   indicate where updates can be found. I do not want to be associated with
'   this DLL in any way, if it has been changed from what I posted

' 2) If you use the code as is, without changing any executable statements,
'   request you leave all the information in the DLL properties window as is.
'   Feel free to compile the DLL to any filename you desire
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Your sole interface to this DLL. This class must be called and initialized
' in order to subclass any window

' ////////// USAGE \\\\\\\\\\

' At the top of your form or class set a WithEvents variable to this class....
' ex:   Private WithEvents mySuclasser as lvEZscls3.lvSubclasser
' You will now have an object mySubclasser with two procedures
'    ProcessMessage & ProcessTrayIcon
' -- click on the top left dropdown box in your form's code window to see mySubclasser

' IN CASE YOU DIDN'T KNOW....
' You cannot subclass a window outside of your thread. Therefore, using this dll
' to try and subclass, say, the system tray won't work.

' The class will not let you subclass a window more than once, nor can you
' cause an error by trying to subclass,unsubclass the wrong window.

' Version 1: Posted on PSC mid Feb 05
' Version 2: Added Window MinMax Sizing class, the system tray icon class,
'           and modifying system menus. Moved common functions to a new
'           cFunctions class. Posted 08 Mar 05
' Version 3: Added GPF protection, rearranged DLL a bit & added more remarks
'           See dllErrorTrap for details on GPF protection. Posted 11 Mar 05
' Version 3.1: reworked how GPFs are trapped & handled. GPF crashes caused by
'           CopyMemory API could not be averted. They can be now

Private Const uMinMax As Long = 2       ' cMinMax
Private Const uTrayIcon As Long = 4     ' cTrayIcons (no longer used)
' ^^ more to come later, I'm sure

'V2. API & constants for the above added classes
' API to assist in restoring tray icons after a crash
Private Declare Function RegisterWindowMessage Lib "user32.dll" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Const WM_STYLECHANGED As Long = &H7D
Private Const WM_GETMINMAXINFO As Long = &H24
Private Const WM_DESTROY As Long = &H2
' Custom, used to restore sys tray icons after Explorer crash & used w/above API
Private WM_IECrashNotify As Long


' this is one of the procedures you will see in your form/class.
' For details see the DoWindowMessage function further down in this class module
Public Event ProcessMessage(ByVal hWnd As Long, ByVal wMsg As Long, _
    wParam As Long, lParam As Long, _
    bOverriden As Boolean, lOverrideReturn As Long)

' separate event for tray icons help keep user projects cleaner. In addition,
' this prevents user from having to know what the custom message value is in order to process it
' For details see the DoWindowMessage function further down in this class module
Public Event ProcessTrayIcon(ByVal hWnd As Long, ByVal wTrayIconID As Long, _
     wMsg As Long, bOverriden As Boolean, lOverrideReturn As Long)
' Note to self: if v3 done, don't forward the wMsg parameter &
'               change wParam to read wTrayIconID
'               change lParam to read wMsg
' * users don't need the custom wMsg value anyway & wParam/lParam are less descriptive
' ^^ Done

' API to help prevent bogus user-supplied data
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
' API to prevent user from trying to subclass a non-thread window
' VB dlls can't do that anyway, so prevent it now
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long

'V2 collections/classes
'Private cFnc As cFunctions ' optional functions like LoWord,HiWord. Created only if needed
'Private cTray As cTrays    ' collection of tray icons. Created only if needed
'Private cSizes As cSizing  ' collection of cMinMax classes. Created only if needed
'Private cSysMenu As cSysMenuMod ' optional functions to tweak system menu. Created only if needed

'V3
Public Enum GPFprotectConstants
    gpfNoErrorTraps = 0
    gpfIDE = 1
    gpfCompiled = 2
End Enum
Public Enum GPFmodeConstants
    gpf_ModeIDE = 1
    gpf_ModeCompiled = 2
End Enum

' linked list variable. The last window subclassed by this class
Private LastLink As Long

Public Function SubclassMe(ByVal hWnd As Long, Optional GPFawareness As GPFprotectConstants = gpfIDE Or gpfCompiled) As Long
' Function subclasses any window in the current thread
' hWnd : a valid window handle to a control or object
' Return Values:
'   -1 indicates passed hWnd already subclassed by this class
'    0 indicates failure, probably attempt to subclass outside of your thread
'    1 indicates success

' 1st precaution: prevent invalid window handle & subclassing outside of thread

' variable to hold copy of class you passed
Dim lClassPtr As Long   ' object pointer to your class
Dim lRtn As Long        ' generic

If IsValidWindow(hWnd, lClassPtr) = False Then Exit Function

' ensure GPF AwarenessLevel is set appropriately
' Affects entire thread. Note that the default is gpfIDE + gpfCompiled
' For compiled apps: this won't be in effect unless you also provide a callback address
GPFclass.SetPrjOwner hWnd, GPFawareness

' call the module's subclassing routine
' Other precautions handled there
lRtn = SubclassWindow(hWnd, lClassPtr, LastLink)

If lRtn = 1 Then ' success
    ' set the previously subclassed window's last linked list to this window
    If LastLink Then SetProp LastLink, lvp_LinkZ, hWnd
    ' update; must point to the last window this class subclasses
    LastLink = hWnd
Else
    ' failed. probably trying to subclass window outside your thread, or worse case,
    ' window prevented caching properties via the SetProperty API
End If

' return result
SubclassMe = lRtn
End Function

Public Function UnSubclassMe(ByVal hWnd As Long, Optional bEntireClass As Boolean) As Long
' Function unsubclasses a single window or all windows subclassed by this class instance
' hWnd : a valid window handle to a control or object unless bEntireClass is True
' bEntireClass : if True all windows subclassed by this class instance will be unsubclassed
' Return Values:
'    0 indicates failure, probably hWnd was never subclassed by this class
'    1 indicates success

Dim lRtn As Long ' generic
' when bEntireClass = True, no hWnd needed. We will use our linked list
If bEntireClass Then
    ' V2. Remove any tray icons related to this class
    'Set cTray = Nothing
    ' call module's routine to subclass all our windows
    ' Precautions taken in that module
    lRtn = RemoveClassSubClassing(LastLink)
    ' reset this flag
    LastLink = 0
    'Set cSizes = Nothing
Else
    ' 1st precaution: ensure valid window handle
    If IsValidWindow(hWnd, 0&) = False Then Exit Function
    
    ' since using a linked list, if we remove a window (link) from the chain
    ' we need to update both sides of that link
    Dim LinkA As Long, LinkZ As Long
    ' get both sides of the link
    LinkA = GetProp(hWnd, lvp_LinkA)
    LinkZ = GetProp(hWnd, lvp_LinkZ)
    
    ' V2. Remove any tray icons related to this hWnd only
    'ApplyTrayNotify hWnd, 0, , True
    ' try to unsubclass the window
    lRtn = UnSubclassWindow(hWnd)
    If lRtn Then
        ' success, window was unsubclassed, now we need to
        ' repair the broken link...
        
        ' connect both sides (links) of unsubclassed window
        If LinkA Then SetProp LinkA, lvp_LinkZ, LinkZ
        If LinkZ Then SetProp LinkZ, lvp_LinkA, LinkA
        
        ' if we unsubclassed our most recent window, we need to
        ' keep the the flag pointing to the now most-recent window
        If LastLink = hWnd Then LastLink = LinkA
        
        'V2. Remove the cMinMax class for this hWnd if applicable
        'ApplyMinMaxSize hWnd, , True
        'ApplySystemMenu hwnd, True
    End If
End If

UnSubclassMe = lRtn
End Function

Friend Sub DoWindowMessage(ByVal hWnd As Long, ByVal wMsg As Long, _
     wParam As Long, lParam As Long, _
    bOverriden As Boolean, lOverrideReturn As Long)
    
On Error GoTo ExitSubclassing

' STRONG SUGGESTION: ALWAYS USE AN ON ERROR GOTO [end of your routine]
' within your subclass procedure. There are many scenarios where an
' error can occur which could crash your application.
' See wm_destroy notes below for an example.

' called from our dll module only
' Handles both class events: ProcessMessage and ProcessTrayIcon
' Pass the message to our client with the following variables


' Parameters in relation to the PROCESSMESSAGE event
' hWnd : the hWnd the message is intended for. This class can subclass multiple windows
' wMsg : the window message. See www.msdn.com for detailed descriptions
' wParam : extra window message info. See www.msdn.com for detailed descriptions
' lParam : extra window message info. See www.msdn.com for detailed descriptions
' bOverriden : if True, then the value returned in lOverridenReturn will be passed
'               as the window message's result. If false, the window message is
'               sent along as is (including any changes you may have made
'               to wParam & lParam)
' lOverriden : only used if you want to return a specific result for the window
'               message. See www.msdn.com for detailed descriptions. Depending on
'               the message you may need to return a 1 or a 0 if you are going to
'               override the message & you don't want it being passed any further.
    
    
' VERSION 2 additions. Since we enable other classes (MinMax, Tray Icons), we will
' also take the initiative & reply to those classes without bothering the user, if applicable
' Version 3. The typical wMsg, wParam & lParam were modified in the Public Event as follows
'       wMsg not passed. This is a custom message user doesn't need to deal with
'       wParam changed to wTrayIconID for better clarification parameter meaning
'       lParam changed to wMsg for better clarification of parameter meaning

' Parameters in relation to the PROCESSTRAYEVENT event
' hWnd : the hWnd the message is intended for. This class can subclass multiple windows
' wTrayIconID : the TrayIconId (see TrayNotifier routine below)
' wMsg : the tray message. See www.msdn.com for detailed descriptions
' bOverriden : if True, then the value returned in lOverridenReturn will be passed
'               as the window message's result. If false, the window message is
'               sent along as is (including any changes you may have made
'               to wParam & lParam)
' lOverriden : only used if you want to return a specific result for the window
'               message. See www.msdn.com for detailed descriptions. Depending on
'               the message you may need to return a 1 or a 0 if you are going to
'               override the message & you don't want it being passed any further.

' **** WARNING ***** WARNING -- prevent sure-fire crash
' In your ProcessTrayIcon event, never allow your project to terminate directly.
' So should you be showing a popup menu with an Exit type menu item, DO NOT allow
' the menu item to call Unload or End. It must post a message to shutdown so that the
' tray icon subclassing can exit before your project terminates. The simple statement
' needed to end your app as a result of processing a tray icon event message is:
' PostMessage Me.hwnd, WM_CLOSE, 0, 0   ' < wm_close = &H10
' the above api command can be executed from anwyhere in your form/class/whatever


'V2. determine if this window has any other attached classes
Dim trackState As Long
Select Case wMsg
    Case WM_DESTROY ' window is being closed
        'If Not cTray Is Nothing Then cTray.RemoveHwnd hWnd   ' remove tray icons now
    Case WM_TrayNotify  ' custom tray icon notification
        ' raise a separate event for the user
        'RaiseEvent ProcessTrayIcon(hWnd, wParam, lParam, bOverriden, lOverrideReturn)
        Exit Sub
    Case WM_IECrashNotify   ' custom message
        ' Explorer crash & restored; reinitialize tray icons
        'If Not cTray Is Nothing Then cTray.RestoreSysTray True
        Exit Sub
'    Case WM_GETMINMAXINFO, WM_STYLECHANGED
'        trackState = GetProp(hWnd, lvp_Tracker)
'        ' got this message. see if the cMinMax is attached
'        If (trackState Or uMinMax) = trackState Then
'            ' yepper, let's autoreply for user
'            Dim tMinMax As cMinMax
'            Set tMinMax = cSizes.WindowMinMax(hWnd)
'            If wMsg = WM_GETMINMAXINFO Then
'                ' call class to update lParam & then boogy out
'                tMinMax.SetMinMaxInfo lParam
'                bOverriden = True
'                lOverrideReturn = 0
'            Else ' wm_stylechanged
'                tMinMax.SetBorderWidth ' update if borderstyle affected
'                ' let message go through
'            End If
'            Set tMinMax = Nothing
'        End If
    Case Else
End Select
    
If Not bOverriden Then
    
    RaiseEvent ProcessMessage(hWnd, wMsg, wParam, lParam, bOverriden, lOverrideReturn)
    
End If

' Notes about trapping the WM_DESTROY message (how-to)
' When a form closes, all of its child controls are destroyed first. Therefore if you
' are using an If ElseIf or Select Case structure similar to one below, an error will
' happen when form closes & your "Select Case" tests for Case Text1.hWnd because
' Text1 was already destroyed and your form can no longer reference it

'On Error Goto ExitSubclassing
'Select Case hWnd
'Case Text1.hWnd        ' Text1 will always get a wm_destroy message
'Case Picture1.hWnd     ' may get a wm_destroy message, unless Text1 was already destroyed: error occurs
'Case Me.hWnd           ' will never get a wm_destroy message since above controls unloaded & error occurred
'End Select

' Now simply moving things around a bit will help:

'On Error Goto ExitSubclassing
'Select Case hWnd
'Case Me.hWnd           ' will always get the wm_destroy message
'Case Text1.hWnd        ' will always get wm_destroy message
'Case Picture1.hWnd     ' may get a wm_destroy message, unless Text1 was already destroyed: error occurs
'End Select

' So how do you guarantee each control gets a wm_destroy message? If you absolutely
' need to know when a control gets that message, ensure it is ahead of any other
' controls in your Select Case structure. If you must get the message for multiple
' controls, do not subclass them using the same instance of lvSubclasser. Create
' another instance as needed. See USAGE remarks at top of this module
Exit Sub

ExitSubclassing:
Debug.Print "unexpected error in lvSubclasser.DoWindowMessage: "; Err.Number, Err.Description
Debug.Print vbTab; hWnd; wMsg; wParam; lParam
End Sub

Public Function ForwardMessage(ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long)

On Error Resume Next    ' never know what will be passed back to us?

' Should you want to forward the message and also modify and/or resend it,
' then you would call this function before you finished responding
' to the ProcessMessage procedure in your module(s).

' This function serves the identical purpose of the CallWindowProc API & simply
' provided so you don't need to declare and use that API. In addition, you would
' have no way of knowing the original window procedure to use. This routine
' will recall that for you & simply pass the message along.

Dim lOldProc As Long
lOldProc = GetProp(hWnd, lvp_WndProc)
If lOldProc Then ForwardMessage = CallWindowProc(lOldProc, hWnd, wMsg, wParam, lParam)

End Function

Public Function IsWindowSubclassed(ByVal hWnd As Long) As Long
' Function is provided in case you get lost somewhere in your routines or
' you want to test to see if any window is subclassed or not

' hwnd : any valid hWnd
' Return Values:
'   -1 indicates function failed to determine (bad hWnd for example)
'    0 indicates not subclassed by this dll
'    1 indicates subclassed by this class instance
'    2 indicates subclassed by another class instance (not this one)

' This is similar to the SubclassMe function at top of module. That function
' will tell you if subclassing was successful (not previously subclassed),
' already subclassed, or unable to subclass-failure. The difference is that
' this routine does not subclass if the passed hWnd is not already subclassed,
' and the hWnd to be tested doesn't have to have been subclassed by this instance
' in order to return a result

Dim lClassPtr As Long
If IsValidWindow(hWnd, lClassPtr) = False Then
    IsWindowSubclassed = -1
    Exit Function
End If

Dim lOldProc As Long
' get ref to hWnd's previous window procedure
lOldProc = GetProp(hWnd, lvp_WndProc)
' if zero, then this dll is not subclassing
If lOldProc = 0 Then Exit Function

' see if the class that subclassed is same as this class
If GetProp(hWnd, lvp_ClassRef) = lClassPtr Then
    IsWindowSubclassed = 1
Else
    IsWindowSubclassed = 2
End If
End Function

Private Sub Class_Terminate()
' class terminating or unloaded. Unsubclass any subclassed windows
' V2. Remove version 2 classes/collections

 ' remove all tray icons & set class to nothing
'Set cTray = Nothing
' unsubclass all windows that were subclassed by this instance
If LastLink Then RemoveClassSubClassing LastLink
LastLink = 0
' remove any sizing restriction classes
'Set cSizes = Nothing
' remove the common functions class
'Set cFnc = Nothing
' remove the system menu editor class
'Set cSysMenu = Nothing
' inform DLL we are terminating
GPFclass.UpdateClassRef True, GetMyPointer()
End Sub

'Public Property Get SysMenuEditor() As cSysMenuMod
''V2. returns a class instance for the cSysMenuMod functions
'' Created on demand & destroyed when lvSubclasser destroyed
'
'' Unlike MinMaxSize & TrayNotifier, this property does not need nor have an
'' associated Apply... function. Changes are immediate.
'
'' in fact, it is possible tomodify any window's system menu.
'' Should I restrict it to only our thread?
'
'If cSysMenu Is Nothing Then Set cSysMenu = New cSysMenuMod
'Set SysMenuEditor = cSysMenu
'End Property

'Public Property Get MinMaxSize(ByVal hWnd As Long) As cMinMax
''V2. Returns a cMinMax class so user can specify windows size parameters
''    Error checking allows a new class to be created on demand if non-existant
'
'' hWnd :: the handle to the window that will have sizing restrictions applied
'
'' In order for the settings to take effect, user must also call ApplyMinMaxSize
'
'If IsWindow(hWnd) = 0 Then Exit Property
'If cSizes Is Nothing Then Set cSizes = New cSizing
'
'On Error Resume Next
'cSizes.AutoAdd = True
'Set MinMaxSize = cSizes.WindowMinMax(hWnd)
'cSizes.AutoAdd = False
'End Property

'Public Function ApplyMinMaxSize(ByVal hWnd As Long, _
'        Optional Deactivate As Boolean, Optional Remove As Boolean) As Boolean
'' Function primarily activates a cMinMax class
'
'' PARAMETERS
'' hWnd :: window that the size restrictions will be applied to
''       this must be the same hWnd you used in your previous call to MinMaxSize
'' Deactivate :: Stops monitoring the WM_GETMINMAXINFO message & does not destroy class.
'' Remove :: If True, same as above except the class is also destroyed
''   If Deactivate is False and Remove is False then the DLL will
''   begin monitoring WM_GETMINMAXINFO messages & restricting window sizing as needed
'
'If cSizes Is Nothing Then Exit Function
'If IsWindow(hWnd) = 0 Then Exit Function
'
'Dim trackState As Long, MinMaxStructure As cMinMax
'On Error Resume Next
'
'trackState = GetProp(hWnd, lvp_Tracker)
'
'If Deactivate = False And Remove = False Then
'    ' activating
'    If GetProp(hWnd, lvp_WndProc) = 0 Then Exit Function
'    Set MinMaxStructure = cSizes.WindowMinMax(hWnd)
'    If Err Then Exit Function
'    ' get the windows borderwidth used for calculating maximized window position
'    MinMaxStructure.SetBorderWidth
'    trackState = trackState Or uMinMax
'    SetProp hWnd, lvp_Tracker, trackState
'Else
'    ' deactivated/removed, update tracking state & destroy class as needed
'    trackState = trackState And Not uMinMax
'    SetProp hWnd, lvp_Tracker, trackState
'    If Remove Then cSizes.RemoveHwnd hWnd
'End If
'
'If Err Then Err.Clear
'ApplyMinMaxSize = True
'
'End Function

'Public Property Get TrayNotifier(ByVal hWnd As Long, ByVal TrayIconID As Long) As cTrayIcons
'' Tray icon as referenced throughout is not the image, it is the placeholder
'' that the image rests on; it is the tray area that recieves mouse actions.
'
'' TrayIconId :: a unique identifier you assign for the tray icon. Used for multiple
''   icons from a single hWnd. Yes you can put more than 1 icon on the
''   system tray for the same hWnd
'' IMPORTANT. This TrayIconID is used whenever you want to modify the system tray
''   like changing tips, showing balloons or changing icons, activating or deactivating
''   Additionally, this TrayIconID is the wParam value within your ProcessTrayIcon event
'
'' Suggestion. If you are only going to have a single tray icon for your hWnd, then
'' use the hWnd as the TrayIconID which prevents you from having to cache another
'' variable within your project.
'
'' In order for the settings to take effect, user must also call ApplyTrayNotify
'
'If IsWindow(hWnd) = 0 Then Exit Property
'If cTray Is Nothing Then Set cTray = New cTrays
'
'On Error Resume Next
'cTray.AutoAdd = True    ' allow new cTrayIcons class to be created if needed
'Set TrayNotifier = cTray.TrayIcon(hWnd, TrayIconID)
'cTray.AutoAdd = False   ' toggle flag
'End Property
'
'Public Function ApplyTrayNotify(ByVal hWnd As Long, ByVal TrayIconID As Long, _
'    Optional Deactivate As Boolean, Optional RemoveAll As Boolean) As Boolean
'
'' PARAMETERS
'' hWnd :: window that tray notifications will be sent to or currently being sent to
''       this must be the same hWnd you used in your previous call to TrayNotifier
'' TrayIconID :: required if Deactive = False or RemoveAll = False
''   Specifies which tray icon then action is intended for.
''   See above property for a description.
'' Deactivate :: only applies to the passed TrayIconID
''   If True, then same as calling the cTrayIcons.RemoveTrayIcon function, except class not destroyed
'' RemoveAll :: If True, then TrayIconID not used. All tray icons for the passed hWnd
''   will be deactivated and their related cTrayIcons class destroyed.
'' If Deactivate = False and RemoveAll = False then TrayIconID will be activated & shown
'
'If cTray Is Nothing Then Exit Function  ' user didn't initialize any tray icons
'If IsWindow(hWnd) = 0 Then Exit Function ' invalid hWnd passed
'
'On Error Resume Next
'Dim TrayNotifier As cTrayIcons
'
'If Deactivate = False And RemoveAll = False Then
'    ' activating an individual tray icon; ensure we are subclassing
'    If GetProp(hWnd, lvp_WndProc) = 0 Then Exit Function
'
'    ' ensure the TrayNotifier was called prior to now or abort
'    Set TrayNotifier = cTray.TrayIcon(hWnd, TrayIconID)
'    If TrayNotifier Is Nothing Then Exit Function
'
'    ' activate the trayicon
'    TrayNotifier.BeginTrayNotifications hWnd, TrayIconID
'
'    ' set up feedback if explorer crashes so we can restore active tray icons
'    If WM_IECrashNotify = 0 Then
'        ' note no harm in repeatedly calling next function; it message
'        ' was already registered, we will simply return the registered value
'        WM_IECrashNotify = RegisterWindowMessage("TaskbarCreated")
'    End If
'
'Else
'    If RemoveAll Then
'        ' no need to pass a valid hWnd; as this option pertains to entire lvSubclasser
'        cTray.RemoveHwnd hWnd ' will also destroy any tray icons
'
'    Else    ' deactivate
'        ' ensure the hWnd & TrayIconID are valid
'        Set TrayNotifier = cTray.TrayIcon(hWnd, TrayIconID)
'        If TrayNotifier Is Nothing Then Exit Function
'        TrayNotifier.RemoveTrayIcon
'
'    End If
'
'End If
'Set TrayNotifier = Nothing
'If Err Then Err.Clear
'ApplyTrayNotify = True
'End Function

Public Property Let GPF_Protection(hWnd As Long, Activation As GPFprotectConstants)
' allow GPF error trapping at the level passed

' note that if you included gpfCompiled, it will not take effect until you also
' set the GPF_Callback property below
GPFclass.SetPrjOwner hWnd, Activation
End Property
Public Property Get GPF_Protection(hWnd As Long) As GPFprotectConstants
' returns GPF_Protection setting.
' Exception: If user included gpfCompiled & did not supply the Callback address
' the this property will NOT include gpfCompiled in its return value
GPF_Protection = GPFclass.AwarenessLevel
' possible returns are
' 0 :: no GPF trapping activated
' 1 :: trapping only in IDE
' 2 :: trapping only when compiled
' 3 :: both of the above
End Property

Public Property Let GPF_Callback(mode As GPFmodeConstants, fAddress As Long)
' When user opts to allow GPF error trapping when the app is compiled, the
' user must also supply a callback routine.

' That callback routine is provided for you to cut & paste: modSampleGPFcallback
' The callback return value is the action the error trapper will perform next.

' The DLL will not make the decision for you when the application is compiled.
' This property is not applied if the GPF_Protection property does not include
' the gpfCompiled flag and your application is compiled.

' Updated. The IDE GPF protection can also have a callback & it can be the same
' as one used for compiled apps or it can be different. It is not required for
' GPF protection while in IDE

If mode > -1 And mode < 4 Then
    If (mode Or gpf_ModeIDE) = mode Then GPFclass.Callback(True) = fAddress
    If (mode Or gpf_ModeCompiled) = mode Then GPFclass.Callback(False) = fAddress
End If
End Property

Public Property Let GPF_IsSilent(HideMessageBoxes As Boolean)

' This property can be used if displaying a message box inadvertently
' causes the GPF handling and your On Error statements to enter an infinite loop.
' N/A for compiled apps. Message boxes are never shown if your project is compiled.

' Example: Let's say in your subclassed event, you trigger a GPF when your
' window gets focus. Well when a GPF message box appears & you subsequently close
' it, your window gets focus back & then triggers another GPF (infinite loop)

GPFclass.HideMessages = HideMessageBoxes
End Property
Public Property Get GPF_IsSilent() As Boolean
GPF_IsSilent = GPFclass.HideMessages
End Property

'Public Property Get Functions() As cFunctions
'' returns a class instance for the generic functions
'' Created on demand & destroyed when lvSubclasser destroyed
'
'If cFnc Is Nothing Then Set cFnc = New cFunctions
'Set Functions = cFnc
'
'End Property

Friend Property Get LastSubclassed() As Long
' Only called from cInit & used during imminent GPF event (IDE only).
' Class passes last hWnd it subclassed, the GPF routines use it
' to unsubclass all windows
' While we're at it, all tray icons need to be removed also
'If Not cTray Is Nothing Then cTray.RestoreSysTray False
LastSubclassed = LastLink
End Property

Private Function IsValidWindow(hWnd As Long, mePtr As Long) As Boolean
' internal helper function.
' Ensure window is part of our thread & get this class' pointer

If IsWindow(hWnd) = 0 Then Exit Function

If GetWindowThreadProcessId(hWnd, 0) <> GetCurrentThreadId() Then Exit Function
mePtr = GetMyPointer
IsValidWindow = True
End Function

Private Function GetMyPointer() As Long
' slight hack for ObjPtr(Me) not working well in dlls
' Not positive above is a true statement; however; read it somewhere
Dim dllImplementation As lvSubclasser
Set dllImplementation = Me
GetMyPointer = ObjPtr(dllImplementation)
Set dllImplementation = Nothing
End Function
