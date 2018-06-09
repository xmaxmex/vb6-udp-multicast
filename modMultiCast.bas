Attribute VB_Name = "modMultiCast"

'**************************************
'Windows API/Global Declarations for :IP Multicasting with Winsock control
'**************************************
Public Type ipm_req
ipm_multiaddr As Long
ipm_interface As Long
End Type

Public Declare Function setsockopt Lib "wsock32" _
(ByVal s As Integer, ByVal level As Integer, _
ByVal optname As Integer, ByRef optval As Any, ByVal optlen As Integer) As Integer

Public Declare Function inet_addr Lib "wsock32" _
    (ByVal cp As String) As Long

