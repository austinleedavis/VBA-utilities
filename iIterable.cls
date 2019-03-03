VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iIterable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Interface for custom data structures that implies all instances can be iterated over using the For-Each construct
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    ''specified in iIterable
End Property
