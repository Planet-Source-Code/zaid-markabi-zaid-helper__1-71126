Attribute VB_Name = "IsConnected32"
Public Declare Function RasEnumConnections Lib "RasApi32.dll" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnections As Long) As Long
Public Declare Function RasGetConnectStatus Lib "RasApi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
Public Const RAS95_MaxEntryName = 256
Public Const RAS95_MaxDeviceType = 16
Public Const RAS95_MaxDeviceName = 32

Public Type RASCONN95
   dwSize As Long
   hRasCon As Long
   szEntryName(RAS95_MaxEntryName) As Byte
   szDeviceType(RAS95_MaxDeviceType) As Byte
   szDeviceName(RAS95_MaxDeviceName) As Byte
End Type

Public Type RASCONNSTATUS95
   dwSize As Long
   RasConnState As Long
   dwError As Long
   szDeviceType(RAS95_MaxDeviceType) As Byte
   szDeviceName(RAS95_MaxDeviceName) As Byte
End Type
