Attribute VB_Name = "MMouseCursor"
Option Explicit
'VBRUN.MousePointerConstants Or MSFORMS.fmMousePointer
Public Enum EMouseCursor
    MCDefault = 0         ' = VBRUN.MousePointerConstants.vbDefault        'MSFORMS.fmMousePointer.fmMousePointerDefault        = 0
    MCArrow = 1           ' = VBRUN.MousePointerConstants.vbArrow          'MSFORMS.fmMousePointer.fmMousePointerArrow          = 1
    MCCrosshair = 2       ' = VBRUN.MousePointerConstants.vbCrosshair      'MSFORMS.fmMousePointer.fmMousePointerCrosshair      = 2
    MCIbeam = 3           ' = VBRUN.MousePointerConstants.vbIbeam          'MSFORMS.fmMousePointer.fmMousePointerIbeam          = 3
    MCIconPointer = 4     ' = VBRUN.MousePointerConstants.vbIconPointer    'MSFORMS.fmMousePointer.fmMousePointerIconPointer    = 4
    MCSizePointer = 5     ' = VBRUN.MousePointerConstants.vbSizePointer    'MSFORMS.fmMousePointer.fmMousePointerSizePointer    = 5
    MCSizeNESW = 6        ' = VBRUN.MousePointerConstants.vbSizeNESW       'MSFORMS.fmMousePointer.fmMousePointerSizeNESW       = 6
    MCSizeNS = 7          ' = VBRUN.MousePointerConstants.vbSizeNS         'MSFORMS.fmMousePointer.fmMousePointerSizeNS         = 7
    MCSizeNWSE = 8        ' = VBRUN.MousePointerConstants.vbSizeNWSE       'MSFORMS.fmMousePointer.fmMousePointerSizeNWSE       = 8
    MCSizeWE = 9          ' = VBRUN.MousePointerConstants.vbSizeWE         'MSFORMS.fmMousePointer.fmMousePointerSizeWE         = 9
    MCUpArrow = 10        ' = VBRUN.MousePointerConstants.vbUpArrow        'MSFORMS.fmMousePointer.fmMousePointerUpArrow        = 10
    MCHourglass = 11      ' = VBRUN.MousePointerConstants.vbHourglass      'MSFORMS.fmMousePointer.fmMousePointerHourglass      = 11
    MCNoDrop = 12         ' = VBRUN.MousePointerConstants.vbNoDrop         'MSFORMS.fmMousePointer.fmMousePointerNoDrop         = 12
    MCArrowHourglass = 13 ' = VBRUN.MousePointerConstants.vbArrowHourglass 'MSFORMS.fmMousePointer.fmMousePointerArrowHourglass = 13
    MCArrowQuestion = 14  ' = VBRUN.MousePointerConstants.vbArrowQuestion  'MSFORMS.fmMousePointer.fmMousePointerArrowQuestion  = 14
    MCSizeAll = 15        ' = VBRUN.MousePointerConstants.vbSizeAll        'MSFORMS.fmMousePointer.fmMousePointerSizeAll        = 15
    MCCustom = 99 '(&H63) ' = VBRUN.MousePointerConstants.vbCustom         'MSFORMS.fmMousePointer.fmMousePointerCustom         = 99
End Enum

Private mIUVTable As TIUnknownVTable
Private mpVTable  As Long '= VarPtr(mIUVTable)

Public Type TMouseCursor
    pVTable     As LongPtr
    pThisObject As IUnknown
    PrevCursor  As EMouseCursor
    MyForm      As Object
End Type

'the prefix "New_" reminds us here we create an object !-)
Public Property Let New_MouseCursor(this As TMouseCursor, aForm As Object, NewCursor As EMouseCursor)
    
    If mpVTable = 0 Then
        
        'now obtain the function pointers
        'this will be done only once in the lifetime of the program!
        'every lightweight class has to use it's own set of IUnknown-functions
        With mIUVTable
            .PQueryInterface = FncPtr(AddressOf QueryInterface)
            .PAddRef = FncPtr(AddressOf AddRef)
            .PRelease = FncPtr(AddressOf Release)
        End With
        mpVTable = VarPtr(mIUVTable)
        
    End If
    With this
        Set .MyForm = aForm
        'save the previous MouseCursor in the variable PrevCursor
        .PrevCursor = Screen.MousePointer
        .pVTable = mpVTable
        RtlMoveMemory .pThisObject, VarPtr(.pVTable), SizeOf_LongPtr
        
        'now we set the new MouseCursor:
        .MyForm.MousePointer = NewCursor
        DoEvents
        
    End With
    
End Property

Private Function QueryInterface(this As TMouseCursor, riid As LongPtr, pvObj As LongPtr) As Long
    
    Debug.Print "QI" 'will not happen here
    pvObj = 0
    'bei Objekten die kein Interface haben:
    QueryInterface = E_NOINTERFACE
    
End Function

Private Function AddRef(this As TMouseCursor) As Long
    
    Debug.Print "AR" 'will not happen here
    'add a reference here if you use reference-counting
    'in the MouseCursor-object we don't nedd it
    
End Function

Private Function Release(this As TMouseCursor) As Long
    Debug.Print "RL"
    
    'subtract a reference here, if you use ref-counting
    'now restore the old mousecursor
    With this
        .MyForm.MousePointer = .PrevCursor
        DoEvents
        Set .MyForm = Nothing
        .pVTable = 0
        RtlZeroMemory .pThisObject, SizeOf_LongPtr
    End With
End Function

