Attribute VB_Name = "MIUnknown"
Option Explicit
'This module is just for reference.
'Adopt the function declarations by copying to your own lightweight-object-modules

'will be returned by QueryInterface, if the objekt does not implement a Interface:
Public Const E_NOINTERFACE As Long = &H80004002

'This is the typical VTable of the IUnknown interface:
Public Type TIUnknownVTable
    PQueryInterface As LongPtr
    PAddRef         As LongPtr
    PRelease        As LongPtr
End Type

'the following 3 functions, must be used in every class, to implement IUnkown
Private Function QueryInterface(this As TIUnknownVTable, riid As LongPtr, pvObj As LongPtr) As Long
    'pvObj = 0
    'for objects with no Interface return E_NOINTERFACE:
    'QueryInterface = E_NOINTERFACE
End Function
Private Function AddRef(this As TIUnknownVTable) As Long
    'if ref counting is needed, here we add a reference
End Function
Private Function Release(this As TIUnknownVTable) As Long
    'here we subtract a reference
End Function


