<div align="center">

## ICQMAPI\.dll DECLARES


</div>

### Description

Establishes Basic communication with Mirabilis ICQ through the Official ICQ API's.
 
### More Info
 
In the function "ICQ_SET_LICENSE" you must set the variables "strName$, strPassword$, strLicense$" with the respective data Mirabilis sends to you when you license "ICQMAPI.dll" (This is Free)

You must obtain a License to use "ICQMAPI.DLL" from Mirabilis. It is free.

Some of the API's may not function correctly because in the documentation I was sent, some of the API's return two values of two different data types. If anyone knows how to fix this, your support would be appreciated.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[MicroVB INC](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/microvb-inc.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/microvb-inc-icqmapi-dll-declares__1-12359/archive/master.zip)

### API Declarations

```
Public Enum ONLINE_STATUS
  BICQAPI_USER_STATE_ONLINE = 0
  BICQAPI_USER_STATE_CHAT = 1
  BICQAPI_USER_STATE_AWAY = 2
  BICQAPI_USER_STATE_NA = 3
  BICQAPI_USER_STATE_OCCUPIED = 4
  BICQAPI_USER_STATE_DND = 5
  BICQAPI_USER_STATE_INVISIBLE = 6
  BICQAPI_USER_STATE_OFFLINE = 7
End Enum
Public Enum DOCKING_STATE
  DCK_FLOATING = 0
  DCK_DOCKED_RIGHT = 1
  DCK_DOCKED_LEFT = 2
  DCK_DOCKED_TOP = 3
  DCK_DOCKED_BOTTOM = 4
End Enum
Public Enum LIST_CHANGE
  USER_GONE_ON_OR_OFF = 1
  USER_FLOAT_WINDOW_ON_OR_OFF = 2
  USER_CHANGED_POSITION_IN_THE_LIST = 3
End Enum
Public Enum GENDER
  NOT_SPECIFIED = 0
  FEMALE = 1
  MALE = 2
End Enum
Public Enum LIST_TYPE
  REGULAR_MODE = 0
  GROUP_MODE = 1
End Enum
Public Type BSICQAPI_FireWallData
  m_bEnabled As Byte
  m_bSocksEnabled As Byte
  m_sSocksVersion As Integer
  m_szSocksHost As String   '512  chr
  m_iSocksPort As Integer
  m_bSocksAuthenticationMethod As Byte
End Type
Public Type BSICQAPI_User
  m_iUIN As Integer
  m_hFloatWindow As hWnd
  m_iIP As Integer
  m_szNickname As String   '20   chr
  m_szFirstName As String   '20   chr
  m_szLastName As String   '20   chr
  m_szEmail As String     '100  chr
  m_szCity As String     '100  chr
  m_szState As String     '100  chr
  m_iCountry As Integer
  m_szCountryName As String  '100  chr
  m_szHomePage As String   '100  chr
  m_iAge As Integer
  m_szPhone As String     '20   chr
  m_bGender As GENDER
  m_iHomeZip As Integer
  'VERSION 1.0001
  m_iStateFlags As ONLINE_STATUS
End Type
'CALLS
Declare Function ICQAPICall_GetDockingState Lib "ICQMAPI" () As DOCKING_STATE
Declare Function ICQAPICall_GetFirewallSettings Lib "ICQMAPI" () As BSICQAPI_FireWallData
Declare Function ICQAPICall_GetFullOwnerData Lib "ICQMAPI" (ppUser As BSICQAPI_User, iVersion As Integer) As BSICQAPI_User
Declare Function ICQAPICall_GetFullUserData Lib "ICQMAPI" (ppUser As BSICQAPI_User, iVersion As Integer) As BSICQAPI_User
Declare Function ICQAPICall_GetOnlineListDetails Lib "ICQMAPI" (iCount As Integer, ppUsers() As BSICQAPI_User)
Declare Function ICQAPICall_GetOnlineListPlacement Lib "ICQMAPI" (iVersion As Integer, iCount As Integer, piEvents() As Byte) As Integer
Declare Function ICQAPICall_RegisterNotify Lib "ICQMAPI" (iVersion As Integer, iCount As Integer, piEvents() As Byte)
'YOU MUST OBTAIN A LICENSE KEY FROM MIRABILIS
'=============================================
Declare Function ICQAPICall_SetLicenseKey Lib "ICQMAPI" (pszName As String, pszPassword As String, pszLicense As String)
'=============================================
Declare Function ICQAPICall_SendFile Lib "ICQMAPI" (iUIN As Integer, pszFileNames As String)
Declare Function ICQAPICall_UnRegisterNotify Lib "ICQMAPI" ()
Declare Function ICQAPICall_GetVersion Lib "ICQMAPI" () As Integer
Declare Function ICQAPICall_GetWindowHandle Lib "ICQMAPI" () As hWnd
'CALLS v1.0.0.1
Declare Function ICQAPICall_GetOnlineListType Lib "ICQMAPI" () As LIST_TYPE
Declare Function ICQAPICall_GetGroupOnlineListDetails Lib "ICQMAPI" (iGroupCount As Integer, ppGroups() As BPSICQAPI_Group)
Declare Function ICQAPICall_SetOwnerState Lib "ICQMAPI" (iState As ONLINE_STATUS)
Declare Function ICQAPICall_SetOwnerPhoneState Lib "ICQMAPI" (iPhoneState As Integer)
Declare Function ICQAPICall_SendMessage Lib "ICQMAPI" (iUIN As Integer, pszMessage As String)
Declare Function ICQAPICall_SendURL Lib "ICQMAPI" (iUIN As Integer, pszMessage As String)
Declare Function ICQAPICall_SendExternal Lib "ICQMAPI" (iUIN As Integer, pszExternal As String, pszMessage As String, bAutoSend As Byte)
'NOTIFICATIONS
Declare Function ICQAPINotify_OnlineListChange Lib "ICQMAPI" (iType As LIST_CHANGE)
Declare Function ICQAPINotify_FullUserDataChange Lib "ICQMAPI" (iUIN As Integer)
Declare Function ICQAPINotify_AppBarStateChange Lib "ICQMAPI" (iDockingState As DOCKING_STATE)
Declare Function ICQAPINotify_OnlinePlacementChange Lib "ICQMAPI" ()
Declare Function ICQAPINotify_OwnerChange Lib "ICQMAPI" (iUIN As Integer)
Declare Function ICQAPINotify_OwnerFullUserDataChange Lib "ICQMAPI" (iUIN As Integer)
Declare Function ICQAPINotify_OnlineListHandleChange Lib "ICQMAPI" (hWindow As hWnd)
Declare Function ICQAPINotify_FileReceived Lib "ICQMAPI" (pszFileNames As String)
'UTILITY FUNCTIONS
Declare Function ICQAPIUtil_FreeUser Lib "ICQMAPI" (pUser As BSICQAPI_User)
Declare Function ICQAPIUtil_FreeUsers Lib "ICQMAPI" (iCount As Integer, ppUsers() As BSICQAPI_User)
Declare Function ICQAPIUtil_SetUserNotificationFunc Lib "ICQMAPI" (uNotificationCode As UIN_TYPE, pUserFunc As void)
```


### Source Code

```
Public Function ICQ_SET_LICENSE()
  Dim strName$, strPassword$, strLicense$
  strName$ = ""
  strPassword$ = ""
  strLicense$ = ""
  ICQAPICall_SetLicenseKey strName, strPassword, strLicense
End Function
Public Function ICQ_GET_VERSION() As Integer
  ICQ_GET_VERSION = ICQAPICall_GetVersion
End Function
Public Function ICQ_GET_DOCKINGSTATE() As DOCKING_STATE
  ICQ_GET_DOCKINGSTATE = ICQAPICall_GetDockingState
End Function
Public Function ICQ_GET_FIREWALLSETTINGS() As BSICQAPI_FireWallData
  ICQ_GET_FIREWALLSETTINGS = ICQAPICall_GetFirewallSettings
End Function
Public Function ICQ_GET_FULL_OWNER_DATA(pUser As BSICQAPI_User) As BSICQAPI_User
  ICQ_GET_FULL_OWNER_DATA = ICQAPICall_GetFullOwnerData
End Function
```

