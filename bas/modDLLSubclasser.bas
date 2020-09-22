Attribute VB_Name = "modDLLSubclasser"
Option Explicit
' Want to check & see if any updates posted?
' http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=59434&lngWId=1

' SUBCLASSING MADE EASIER - BASICALLY A DROP IN DLL

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

' First, subclassing by its nature is dangerous in VB. You can easily crash the IDE
' and lose any changes you made.  Therefore, strongly suggest compiling this DLL
' as is, then simply referencing it in your project.  You have the source code here,
' there is no reason to include the DLL modules in your project.

' If the DLL is compiled & not inside your application, I have not found a way of
' crashing the IDE unless the user executes and END statement in their code. The
' crash in this case is extremely rare & unpredicatable, but it could happen.
' So do not use END statements in your code. Terminate your application
' gracefully by using the Unload Me commands as needed.

' IMO, use of the End statement reflects poor coding, negates clean-up of other
' objects, and always produces memory leaks. I cannot think of any valid
' reason for using an End statement in your code. However, there are many reasons
' why you might want to or need to hit the End button on your VB toolbar while in
' IDE -- that's different, it is debugging where End statements get compiled.

' If using the DLL (compiled or not), hitting the Halt/End button on the VB toolbar
' has no ill effects short of memory leaks from other controls. This is a benefit.

' To use this simple subclasser, follow these easy steps.

'1. Preferably, compile the DLL & call it lvEZscls3 or whatever you want
'2. Open your project in design view & click the menu Project | References
'3. Find the dll you compiled in the list or use the browse button & go get it
'4. In each class or form you want to use subclassing, declare an instance of the class
'       Private WithEvents mySubclasser As lvEZscls3.lvSubclasser
'5. Before you use it the first time, you need to initialize it somewhere
'       If mySubclasser Is Nothing Then Set mySubclasser = New lvEZscls3.lvSubclasser
'6. Now simply call the SubclassMe or UnsubclassMe functions as needed. See the
'   lvSubclasser class for more information if needed.
'7. Good practice would be to unload the class sometime before you exit. However,
'   this module was written to self-unload when needed. To manually unload...
'       Set mySubclasser = Nothing
'8. Finally, within your new procedure in your form or class, you need to
'   add your message processing code. And one word of caution comes along...
'   If your code causes an error in your ProcessMessage event, all bets are
'   off and your project very well may crash. If you unload your project from
'   the ProcessTrayIcon event, your project will crash if DLL is uncompiled.
'   See lvSubclasser's DoWindowMessage for details & bunch of other helpful info
' New with v3 is auto-GPF (General Protection Fault) protection. Read dllErrorTrap.

'   ALWAYS start your project with Ctrl+F5


' MORE DETAILS TO HELP FOLLOWING ALONG...
' Q: Why a linked list is used instead of caching hWnds within the class file itself?
' A: If the class is destroyed unexpectedly (End for example), the class is destroyed
'       and the Class_Terminate event will not be fired, all info stored in it is gone.
'       Therefore I wouldn't be able to unsubclass all the windows that class has
'       subclassed (crashing as a result). However, by using a linked list, I can
'       navigate the links to find every window subclassed so it can then be unsubclassed.

' Q: Why not use Implements vs a public event?
' A: Good question. Implements is much faster than a public event. However, VB treats
'       the 2 differently. The public event somehow postpones the destruction caused
'       by an END statement and my routines here get a whiff of it which allows me
'       to unsubclass & prevent a crash. But when using Implementation, I lose that
'       capability. I don't know why. To use Implements would require the assembly
'       thunking that Paul Caton published & modifications of his code as done
'       very well by Luke (a.k.a SelfTaught).
'       For advanced subclassing, recommend those methods.

' Q: So, you can get a whiff of when END was executed? How?
' A: Ah, that was so easy & surprised other non-thunking coders out there didn't think
'       of it too. Most know how to create a copy of a class by using ObjPtr and
'       CopyMemory (see NewSubWndProc function). I decided to track just a
'       little bit more info: the value that pointer is assigned. Then I can check
'       that value to see if it changed & if it did change, then something else
'       was written at that memory location. Therefore the class doesn't exist
'       any longer and a rude termination of the application occurred. A graceful
'       termination would have had the class unload normally which would have fired
'       the Class_Terminate event and all windows would have been unsubclassed
'       normally which then would have prevented that check from ever firing.

' Q: So why use this if the thunking Implements versions are faster?
' A: It is pretty safe if the DLL is compiled. Beginners to intermediate coders can
'       safely experiment with subclassing & learning through a simpler method of
'       subclassing. Advanced users may wish to use the DLL for projects that do not
'       require a ton of subclassed windows since this code is small in comparison
'       to the large & complicated thunking code of more advanced methods.
'       Besides if not familiar with subclassing, challenge you to find another
'       project with so many comments & hints ;)
' New with v3 is auto-GPF (General Protection Fault) protection. Read dllErrorTrap.

' Now to the code.


' These values will be cached on the window itself as properties
Public Const lvp_ClassRef As String = "lvCref"      ' the class pointer
Private Const lvp_ClassVal As String = "lvCval"     ' value of the class pointer
Public Const lvp_WndProc As String = "lvWndProc"    ' the previous window procedure
Public Const lvp_LinkA As String = "lvLinkA"        ' near side linked list item
Public Const lvp_LinkZ As String = "lvLinkZ"        ' far side linked list item
Public Const lvp_Tracker As String = "lvState"

Private Const GWL_WNDPROC = (-4)     ' used for subclassing
Private Const WM_DESTROY = &H2       ' used for unsubclassing if needed
Private Const WM_USER = &H400        ' used to get tray icon notifications
Public Const WM_TrayNotify As Long = WM_USER + &H1962

' Subclassing APIs. See www.msdn.com for detailed info
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' Window property APIs
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

' Our favorite app-crashing API
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' API used to help prevent crash (in theory, unable to prove or disprove it)
Private Declare Function IsBadReadPtr Lib "kernel32.dll" (ByVal lp As Long, ByVal ucb As Long) As Long

' used to set some DLL initial variables & also as a clean up tool
Private Initializer As cInit

Private Sub Main()
    ' class mimics a C dll's Lib_Initialize & Lib_Terminate events
    ' in the class, GPF error trapping is triggered & clean up occurs
    Set Initializer = New cInit
End Sub

Public Function SubclassWindow(ByVal hWnd As Long, ByVal classPtr As Long, LinkA As Long) As Long
' Function subclasses passed window and only called by lvSubclasser class
' hWnd : a valid window handle to a control or object
' Return Values:
'   -1 indicates passed hWnd already subclassed
'    0 indicates failure, probably attempt to subclass outside of your thread
'    1 indicates success

If GetProp(hWnd, lvp_WndProc) <> 0 Then 'Already subclassed
    SubclassWindow = -1
    Exit Function
End If

Dim lRtn As Long
' set properties to be recalled when window gets a message.
' test each action to ensure subclassing will be successful

' this is the newest window subclassed, set link to previous window
If SetProp(hWnd, lvp_LinkA, LinkA) Then
    ' although not used immediately, set property to ensure it will take when subsuquent windows are subclassed
    If SetProp(hWnd, lvp_LinkZ, 0) Then
        ' cache reference to previous window procedure
        If SetProp(hWnd, lvp_WndProc, GetWindowLong(hWnd, GWL_WNDPROC)) Then
            ' V2. cache tracking property
            If SetProp(hWnd, lvp_Tracker, 0) Then
                ' cache reference to the class pointer passed
                If SetProp(hWnd, lvp_ClassRef, classPtr) Then
                    ' get value of that pointer & reference
                    CopyMemory lRtn, ByVal classPtr, &H4
                    lRtn = SetProp(hWnd, lvp_ClassVal, lRtn)
                End If
            End If
        End If
    End If
End If

If lRtn Then
    'Now subclass the window
    lRtn = Abs((SetWindowLong(hWnd, GWL_WNDPROC, AddressOf NewSubWndProc)) <> 0)
End If

If lRtn Then
    Initializer.UpdateClassRef False, classPtr
Else    ' failed. Why? Probably trying to subclass a window outside our thread
    ' remove those properties
    RemoveProp hWnd, lvp_WndProc
    RemoveProp hWnd, lvp_ClassVal
    RemoveProp hWnd, lvp_ClassRef
    RemoveProp hWnd, lvp_LinkA
    RemoveProp hWnd, lvp_Tracker
End If

SubclassWindow = lRtn
End Function

Public Property Get GPFclass() As cInit
' expose this class to our entire DLL
Set GPFclass = Initializer
End Property


Public Function NewSubWndProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' Function receives subclassed messages & forwards them along
' hWnd : a valid window handle to a control or object
' wMsg : the window message sent by Windows
' wParam : added message information
' lParam : added message information
' Return Value is determined by your class when you complete the ProcessMessage event

Dim lOldProc As Long    ' ref to old window procedure
Dim lClassPtr As Long   ' ref to class pointer
Dim lClassVal As Long   ' ref to pointer's value
Dim bDiscard As Boolean ' indication to override the window message return value
Dim lRtn As Long        ' generic
Dim dllImplementation As lvSubclasser   ' instance of your class

'Retrieve the needed property values
lOldProc = GetProp(hWnd, lvp_WndProc)
lClassPtr = GetProp(hWnd, lvp_ClassRef)
lClassVal = GetProp(hWnd, lvp_ClassVal)

' IsBadReadPtr is an added safety net & haven't been able to prove its usefulness
' It would work something like this... We have a class, it was destroyed by
'   and End statement and then immediately another process got that memory
'   allocated. This API will tell me if I don't have access to even read
'   the 4 bytes I want to read. If I have access then I'll read those 4 bytes,
'   if not, I avoid a crash from trying to access memory I don't have access to.
If IsBadReadPtr(lClassPtr, &H4) = 0 Then CopyMemory lRtn, ByVal lClassPtr, &H4

' Test for bad pointer
If lRtn <> lClassVal Then
    ' if the pointer is bad, don't allow the message to get through else CRASH
    Set Initializer = Nothing ' unsubclass all windows & replace original GPF handle
    ' although crash prevented, memory leaks likely occurred by your application
    ' Bottom line: don't use the End statement & avoid using the vb toolbar End button
    Exit Function                   ' abort
End If

On Error GoTo ExitSubclassing

' get a copy of your class
CopyMemory dllImplementation, lClassPtr, &H4

' Test for auto-unsubclassing
If wMsg = WM_DESTROY Then
    UnSubclassWindow hWnd ' destroying window, unsubclass now
    ' prevent user from modifying message. If wanting to abort closing app, user
    ' should have trapped the wm_syscommand msg & the sc_close wParam value.
    ' System tray icons will be removed within the DoWindowMessage
    Call dllImplementation.DoWindowMessage(hWnd, wMsg + 0&, wParam, lParam, False, 0)
Else
    ' now call your class' ProcessMessage event
    lRtn = 0
    Call dllImplementation.DoWindowMessage(hWnd, wMsg, wParam, lParam, bDiscard, lRtn)
End If
' remove reference to your class
CopyMemory dllImplementation, 0&, &H4

If bDiscard Then ' overriden by user, pass that overriden value
    NewSubWndProc = lRtn
Else
    ' forward the message with original or modified wParam & lParam values
    NewSubWndProc = CallWindowProc(lOldProc, hWnd, wMsg, wParam, lParam)
End If
Exit Function

ExitSubclassing:
End Function

Public Function UnSubclassWindow(ByVal hWnd As Long) As Long
' Function unsubclasses passed window
' hWnd : a valid window handle to a control or object
' Return Values:
'    0 indicates failure, probably attempt to unsubclass window not subclassed
'    1 indicates success

' sanity checks first
If hWnd = 0 Then Exit Function
Dim lOldProc As Long
' see if we subclassed this one or not
lOldProc = GetProp(hWnd, lvp_WndProc)
If lOldProc = 0 Then Exit Function

'Unsubclass the window
If SetWindowLong(hWnd, GWL_WNDPROC, lOldProc) Then
    'Success, clean up properties
    RemoveProp hWnd, lvp_WndProc
    RemoveProp hWnd, lvp_ClassRef
    RemoveProp hWnd, lvp_ClassVal
    RemoveProp hWnd, lvp_LinkA
    RemoveProp hWnd, lvp_LinkZ
    RemoveProp hWnd, lvp_Tracker    'V2
    UnSubclassWindow = 1
End If
End Function

Public Function RemoveClassSubClassing(ByVal hWnd As Long) As Long
' Function unsubclasses all windows subclassed by a specific class
' Can be called when lvSubclasser is set to Nothing. But also can be
' called as a result of End statement or End button clicked.

' hWnd : a valid window handle to a control or object
' Return Values:
'    0 indicates failure, probably attempt to unsubclass window not subclassed
'    1 indicates success
    
' sanity check
If hWnd = 0 Then Exit Function

' Linked list references
Dim LinkA As Long, LinkZ As Long
Dim lRtn As Long

' get the near & far link list items
LinkA = GetProp(hWnd, lvp_LinkA)
LinkZ = GetProp(hWnd, lvp_LinkZ)

' unsubclass the passed window first
If UnSubclassWindow(hWnd) Then lRtn = 1

' now we follow the far links & unsubclass those windows
Do While LinkZ
    hWnd = LinkZ
    LinkZ = GetProp(hWnd, lvp_LinkZ)
    If UnSubclassWindow(hWnd) Then lRtn = 1
Loop

' now we follow the near links & unsubclass those windows
hWnd = LinkA
Do While hWnd
    LinkA = GetProp(hWnd, lvp_LinkA)
    If UnSubclassWindow(hWnd) Then lRtn = 1
    hWnd = LinkA
Loop

RemoveClassSubClassing = lRtn
End Function
