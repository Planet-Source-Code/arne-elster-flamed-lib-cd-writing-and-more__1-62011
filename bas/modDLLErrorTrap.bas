Attribute VB_Name = "modDLLErrorTrap"
Option Explicit
' Want to check & see if any updates posted?
' http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=59434&lngWId=1

' Excellent piece of coding found in numerous places & variations to include the
' API Guide, PSC & probably the orignal source:
' Article: http://www.ftponline.com/Archives/premier/mgznarch/vbpj/1999/05may99/bb0599.pdf
' Source code:  http://www.fawcette.com/
' then enter VBPJ0599BB in the code locator & click GO (the article is interesting)

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

' Acronym: UEH = Unhandled Exception Handler

' A major added benefit of using the subclasser routines

' Throughout the remarks in this entire project, one common theme is always
' repeated over and over: don't use the End statement!
' The end statement will negate this added protection if DLL is not compiled

' =========================================================================================
' Notes from MSDN (the API talked about is in the cInit class)

'Issuing SetUnhandledExceptionFilter replaces the existing top-level exception filter
'for all existing and all future threads in the calling process.

'The exception handler specified by lpTopLevelExceptionFilter is executed in the context
'of the thread that caused the fault. This can affect the exception handler's ability
'to recover from certain exceptions, such as an invalid stack.
' =========================================================================================
' SUMMARY....
' Only run one GPF protection routine. In otherwords, if you have your own
' module or 3rd party software that incorporates the SetUnhandledExceptionFilter
' API, only run one because the most recent one activated will override the latest.
' Note that each time you subclass a window, the GPF Protection level is automatically
' set to gpfIDE + gpfCompiled. If you want any other setting than this default,
' you must provide that optional parameter every time you use
' lvSubclasser.SubclassMe hWnd, [optional parameter]

' =========================================================================================
'       CRASH SURVIVABILITY
' =========================================================================================
' DLL uncompiled - 99% crash survival if GPF Protection enabled & following 2 don't apply
'   Exception 1: If GPF caused in subclassed event & no On Error
'                   statements exist there -- 0% survivability
'   Exception 2: Executing an End statement negates GPF trapping

' DLL compiled - 99% crash survival if GPF Protection enabled
'   Note: You will not be able to step thru the error when DLL is compiled. That is
'   cause the DLL consumes the GPF & your project and/or callback will never be
'   notified except thru the message box warnings from the DLL. Though this isn't
'   ideal, it does keep your IDE alive.
'   -- If you absolutely cannot determine where the GPF happened, you can consider
'       including the uncompiled DLL in your project. When the uncompiled DLL hits
'       a GPF, it can allow stepping thru the error as long as Exception 1 above
'       doesn't apply. Stepping thru a GPF may cause a crash in itself

' Project compiled - up to your callback return value & whether or not
'   On Error statements exist. High probability of survival if On Error present

' The modSampleGPFcallback file is never used in this DLL. It is only part of
' this project so you will always have a copy of a generic callback function
'   - Required in your project if GPF protection used for compiled apps
'   - Optional if project is in IDE
' See lvSubclasser.GPF_Callback, GPF_Protection, & .GPF_IsSilent properties
' =========================================================================================


' =========================================================================================
' COMMON GPF CAUSING CODE

' To help prevent a crash within your application, ensure your
' subclassing events and any other events have ON ERROR statements
' where the event contains code functionality like the following:

'   - processing subclassed messages
'   - utilizes the CopyMemory API
'   - references objects like classes. Ref an Object set to Nothing can be a crash
'   - references VB collections for same reason as objects
'   - does any math (division by zero, exceeding integer,byte min/max values)
'   - in any form's Resize event if you have additional any code in that event

' bottom line: if you can get a VB Debug/End popup, your code will crash
' when compiled and no On Error statements exist
' =========================================================================================

'Possible return values for the Unhandled Exception Filter.
Private Const EXCEPTION_CONTINUE_EXECUTION = -1
Private Const EXCEPTION_CONTINUE_SEARCH = 0
Private Const EXCEPTION_EXECUTE_HANDLER = 1

'Maximum number of parameters an Exception_Record can have
Private Const EXCEPTION_MAXIMUM_PARAMETERS = 15

'Structure that contains processor-specific register data
Private Type CONTEXT    ' << the longest one I'm familiar with; it's huge!
  FltF0 As Double
  FltF1 As Double
  FltF2 As Double
  FltF3 As Double
  FltF4 As Double
  FltF5 As Double
  FltF6 As Double
  FltF7 As Double
  FltF8 As Double
  FltF9 As Double
  FltF10 As Double
  FltF11 As Double
  FltF12 As Double
  FltF13 As Double
  FltF14 As Double
  FltF15 As Double
  FltF16 As Double
  FltF17 As Double
  FltF18 As Double
  FltF19 As Double
  FltF20 As Double
  FltF21 As Double
  FltF22 As Double
  FltF23 As Double
  FltF24 As Double
  FltF25 As Double
  FltF26 As Double
  FltF27 As Double
  FltF28 As Double
  FltF29 As Double
  FltF30 As Double
  FltF31 As Double
  IntV0 As Double
  IntT0 As Double
  IntT1 As Double
  IntT2 As Double
  IntT3 As Double
  IntT4 As Double
  IntT5 As Double
  IntT6 As Double
  IntT7 As Double
  IntS0 As Double
  IntS1 As Double
  IntS2 As Double
  IntS3 As Double
  IntS4 As Double
  IntS5 As Double
  IntFp As Double
  IntA0 As Double
  IntA1 As Double
  IntA2 As Double
  IntA3 As Double
  IntA4 As Double
  IntA5 As Double
  IntT8 As Double
  IntT9 As Double
  IntT10 As Double
  IntT11 As Double
  IntRa As Double
  IntT12 As Double
  IntAt As Double
  IntGp As Double
  IntSp As Double
  IntZero As Double
  Fpcr As Double
  SoftFpcr As Double
  Fir As Double
  Psr As Long
  ContextFlags As Long
  Fill(4) As Long
End Type

'Structure that describes an exception.
Private Type EXCEPTION_RECORD
    ExceptionCode As Long
    ExceptionFlags As Long
    pExceptionRecord As Long  ' Pointer to an EXCEPTION_RECORD structure
    ExceptionAddress As Long
    NumberParameters As Long
    ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS) As Long
End Type

'Structure that contains exception information that can be used by a debugger.
Private Type EXCEPTION_DEBUG_INFO
        pExceptionRecord As EXCEPTION_RECORD
        dwFirstChance As Long
End Type

'The EXCEPTION_POINTERS structure contains an exception record with a
'machine-independent description of an exception and a context record
'with a machine-dependent description of the processor context at the
'time of the exception.
Private Type EXCEPTION_POINTERS
    pExceptionRecord As EXCEPTION_RECORD
    ContextRecord As CONTEXT
End Type

'Standard Exception Codes
Private Const EXCEPTION_ACCESS_VIOLATION = &HC0000005
Private Const EXCEPTION_DATATYPE_MISALIGNMENT = &H80000002
Private Const EXCEPTION_BREAKPOINT = &H80000003
Private Const EXCEPTION_SINGLE_STEP = &H80000004
Private Const EXCEPTION_ARRAY_BOUNDS_EXCEEDED = &HC000008C
Private Const EXCEPTION_FLT_DENORMAL_OPERAND = &HC000008D
Private Const EXCEPTION_FLT_DIVIDE_BY_ZERO = &HC000008E
Private Const EXCEPTION_FLT_INEXACT_RESULT = &HC000008F
Private Const EXCEPTION_FLT_INVALID_OPERATION = &HC0000090
Private Const EXCEPTION_FLT_OVERFLOW = &HC0000091
Private Const EXCEPTION_FLT_STACK_CHECK = &HC0000092
Private Const EXCEPTION_FLT_UNDERFLOW = &HC0000093
Private Const EXCEPTION_INT_DIVIDE_BY_ZERO = &HC0000094
Private Const EXCEPTION_INT_OVERFLOW = &HC0000095
Private Const EXCEPTION_PRIVILEGED_INSTRUCTION = &HC0000096
Private Const EXCEPTION_IN_PAGE_ERROR = &HC0000006
Private Const EXCEPTION_ILLEGAL_INSTRUCTION = &HC000001D
Private Const EXCEPTION_NONCONTINUABLE_EXCEPTION = &HC0000025
Private Const EXCEPTION_STACK_OVERFLOW = &HC00000FD
Private Const EXCEPTION_INVALID_DISPOSITION = &HC0000026
Private Const EXCEPTION_GUARD_PAGE_VIOLATION = &H80000001
Private Const EXCEPTION_INVALID_HANDLE = &HC0000008
Private Const EXCEPTION_CONTROL_C_EXIT = &HC000013A

Private Function GetExceptionText(ByVal ExceptionCode As Long) As String
'******************************
'  GetExceptionText
'******************************
'  This function receives an exception code value and returns the
'  text description of the exception.
'
  Dim strExceptionString As String
  
  Select Case ExceptionCode
    Case EXCEPTION_ACCESS_VIOLATION
      strExceptionString = "Access Violation"
    Case EXCEPTION_DATATYPE_MISALIGNMENT
      strExceptionString = "Data Type Misalignment"
    Case EXCEPTION_BREAKPOINT
      strExceptionString = "Breakpoint"
    Case EXCEPTION_SINGLE_STEP
      strExceptionString = "Single Step"
    Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED
      strExceptionString = "Array Bounds Exceeded"
    Case EXCEPTION_FLT_DENORMAL_OPERAND
      strExceptionString = "Float Denormal Operand"
    Case EXCEPTION_FLT_DIVIDE_BY_ZERO
      strExceptionString = "Divide By Zero"
    Case EXCEPTION_FLT_INEXACT_RESULT
      strExceptionString = "Floating Point Inexact Result"
    Case EXCEPTION_FLT_INVALID_OPERATION
      strExceptionString = "Invalid Operation"
    Case EXCEPTION_FLT_OVERFLOW
      strExceptionString = "Float Overflow"
    Case EXCEPTION_FLT_STACK_CHECK
      strExceptionString = "Float Stack Check"
    Case EXCEPTION_FLT_UNDERFLOW
      strExceptionString = "Float Underflow"
    Case EXCEPTION_INT_DIVIDE_BY_ZERO
      strExceptionString = "Integer Divide By Zero"
    Case EXCEPTION_INT_OVERFLOW
      strExceptionString = "Integer Overflow"
    Case EXCEPTION_PRIVILEGED_INSTRUCTION
      strExceptionString = "Privileged Instruction"
    Case EXCEPTION_IN_PAGE_ERROR
      strExceptionString = "In Page Error"
    Case EXCEPTION_ILLEGAL_INSTRUCTION
      strExceptionString = "Illegal Instruction"
    Case EXCEPTION_NONCONTINUABLE_EXCEPTION
      strExceptionString = "Non Continuable Exception"
    Case EXCEPTION_STACK_OVERFLOW
      strExceptionString = "Stack Overflow"
    Case EXCEPTION_INVALID_DISPOSITION
      strExceptionString = "Invalid Disposition"
    Case EXCEPTION_GUARD_PAGE_VIOLATION
      strExceptionString = "Guard Page Violation"
    Case EXCEPTION_INVALID_HANDLE
      strExceptionString = "Invalid Handle"
    Case EXCEPTION_CONTROL_C_EXIT
      strExceptionString = "Control-C Exit"
    Case Else
      strExceptionString = "Unknown (&H" & Right("00000000" & Hex(ExceptionCode), 8) & ")"
  End Select
  GetExceptionText = strExceptionString
End Function


Public Function lvErrorChecker(ByRef ExceptionPtrs As EXCEPTION_POINTERS) As Long
  
Static s_Count As Integer ' - recursion counter

'  This function will be called when an unhandled exception occurs.

' If your app is in IDE, all subclassing will be terminated & you may be able to step through
' the error to try to debug why the GPF occurred.

' If your app is compiled, this routine will notify you through the callback address
' you passed to the DLL, if applicable
  
Dim Rec As EXCEPTION_RECORD
Dim strException As String

'Get the current exception record.
Rec = ExceptionPtrs.pExceptionRecord

'If Rec.pExceptionRecord is not zero, then it is a nested exception and
'Rec.pExceptionRecord points to another EXCEPTION_RECORD structure.  Follow
'the pointers back to the original exception.
Do Until Rec.pExceptionRecord = 0
  CopyMemory Rec, ByVal Rec.pExceptionRecord, Len(Rec)
Loop

'Translate the exception code into a user-friendly string.
strException = GetExceptionText(Rec.ExceptionCode)
' Append either Read or Write if Access Violation (Page Fault)
If Rec.ExceptionCode = EXCEPTION_ACCESS_VIOLATION Then
    If Rec.ExceptionInformation(0) Then
        strException = "Write " & strException
    Else
        strException = "Read " & strException
    End If
End If
  
  

If GPFclass.ClientIde Then
    '  client is uncompiled
    
    If s_Count < 5 Then
    
        GPFclass.ReleaseGPFReferences       ' in IDE - unsubclass everything
        
        If Not GPFclass.HideMessages Then
        
            ' provide a popup & then fire a message box
            MsgBox "General Protection Fault Detected - Crash May Be Avoided" & vbCrLf & _
                    "Error: " & strException & vbCrLf & vbCrLf & _
                    "ALL SUBCLASSING HAS BEEN TERMINATED" & vbCrLf & _
                    "If you survive the crash, ensure you close you project gracefully & restart.", _
                    vbExclamation + vbOKOnly, "DLL Generated Message - Potential Crash Preparations"
        End If
        
        If GPFclass.Callback(True) Then
            
            ' user provided an IDE callback, send notification there
            lvErrorChecker = CallWindowProc(GPFclass.Callback(True), Rec.ExceptionCode, Rec.ExceptionFlags, Rec.ExceptionInformation(0), Rec.ExceptionAddress)
            If lvErrorChecker = EXCEPTION_CONTINUE_EXECUTION Then
                s_Count = s_Count + 1
            Else
                s_Count = 0
            End If
                
        Else
            s_Count = 0
            Err.Raise Rec.ExceptionCode, App.Title, "GPF::" & strException
            ' The above action in-effect will cancel the GPF if possible
        End If
        
    Else
        ' exceeded recursion count
        s_Count = 0
        lvErrorChecker = EXCEPTION_EXECUTE_HANDLER
    End If
    
Else    ' both DLL & client are  compiled
    
    ' The DLL will not make the decision to simply resume next, since a GPF can
    ' possibly crash the entire OS or cause unsaved work being lost. Therefore,
    ' since user wants notification, we will send it to him/her & wait for their
    ' answer. Your callback is good place to log needed info, close objects, etc
                
    If s_Count < 5 Then
    
        If (GPFclass.Callback(False) <> 0 And GPFclass.AwarenessLevel > 1) Then
            
            lvErrorChecker = CallWindowProc(GPFclass.Callback(False), Rec.ExceptionCode, Rec.ExceptionFlags, Rec.ExceptionInformation(0), Rec.ExceptionAddress)
            ' hope for the least destruction
            
            If lvErrorChecker = EXCEPTION_CONTINUE_EXECUTION Then
                s_Count = s_Count + 1
            Else
                s_Count = 0
            End If
            
            
        Else
        
            s_Count = 0
            lvErrorChecker = EXCEPTION_CONTINUE_SEARCH
            ' most definitely gonna crash
            
        End If
        
    Else
        ' exceeded recursion count, gonna crash
        s_Count = 0
        lvErrorChecker = EXCEPTION_EXECUTE_HANDLER
    End If
End If

End Function
