VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DIFxAPI_Wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'==========================================================================================
'
'                        <<<< Module Information >>>>
'
'   ---------------------------------------------------------------------------------------
'
'               Module Name................ <:> DIFxAPI_Wrapper
'               Module File................ <:> Cls_Mod_DIFxAPIWrapper.cls
'               Module Type................ <:> Class Module
'               Module Date................ <:> 2008-04-06
'               Module Time................ <:> 22:14
'               Module Version............. <:> v1.0
'               Module Copyright........... <:> Ranuka © 2008-04-06
'               Module Description......... <:> N/A
'               Module Dependencies........ <:> N/A
'
'   ---------------------------------------------------------------------------------------
'
'    Copyright (C) 2002 - 2008 Nymph Technologies
'
'  *******************
'  * Module Revision *
'  *******************
'
'------------------------------------------------------------------------------------------
'   | - Procedure - |    | - Comment - |    | - Version - |    | - Date - |              -
'------------------------------------------------------------------------------------------
'
'          N/A                 N/A                N/A               N/A
'
'------------------------------------------------------------------------------------------
'
'==========================================================================================

Option Explicit

'// DIFxAPI specific return codes

Private Const CERT_E_EXPIRED = -2146762495
Private Const APPLICATION_ERROR_MASK = &H20000000
Private Const ERROR_SEVERITY_ERROR = &HC0000000
Private Const ERROR_NO_SUCH_DEVINST = APPLICATION_ERROR_MASK Or ERROR_SEVERITY_ERROR Or &H20B = &HE000020B

'// API return codes

Private Const ERROR_CANNOT_MAKE = 82
Private Const ERROR_UNSUPPORTED_TYPE = 1630
Private Const ERROR_INSTALL_FAILURE = 1603
Private Const ERROR_CANT_ACCESS_FILE = 1920
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const ERROR_ALREADY_EXISTS = 183&
Private Const ERROR_SHARING_VIOLATION = 32&
Private Const ERROR_ACCESS_DENIED = 5&
Private Const ERROR_BAD_ENVIRONMENT = 10&
Private Const ERROR_INVALID_PARAMETER = 87
Private Const ERROR_INVALID_NAME = 123&
Private Const ERROR_FILENAME_EXCED_RANGE = 206&
Private Const ERROR_OUTOFMEMORY = 14&
Private Const ERROR_INVALID_FUNCTION = 1

'// DIFxAPI specific public types

Public Type INSTALLERINFO

pApplicationId          As String
pDisplayName            As String
pProductName            As String
pMfgName                As String

End Type

'// function parameter public enums

Public Enum UninstallFlags

Uninstall_Normal = 0
Uninstall_Force = 4
Uninstall_Silent = 2
Uninstall_DeleteFiles = 20

End Enum

Public Enum InstallOrPreinstallFlags

Normal = 0
Repair = 1
Force = 4
Silent = 2
LegacyMode = 10
OnlyIfDeveicePresent = 8

End Enum

'// private variables

Private Flag            As Long

Private Declare Function DriverPackagePreinstall Lib "Difxapi" Alias "DriverPackagePreinstallA" (ByVal lpctstr As String, ByVal Flags As Long) As Long
Private Declare Function DriverPackageInstall Lib "Difxapi" Alias "DriverPackageInstallA" (ByVal lpctstr As String, ByVal Flags As Long, ByRef pInstallerInfo As INSTALLERINFO, ByRef pNeedReboot As Boolean) As Long
Private Declare Function DriverPackageUninstall Lib "Difxapi" Alias "DriverPackageUninstallA" (ByVal lpctstr As String, ByVal Flags As Long, ByRef pInstallerInfo As INSTALLERINFO, ByRef pNeedReboot As Boolean) As Long

Public Function Install(InfPath As String, Flag As InstallOrPreinstallFlags, InstallerInfo_ApplicationId As String, InstallerInfo_DisplayName As String, InstallerInfo_ProductName As String, InstallerInfo_MfgName As String, NeedReboot As Boolean) As String

'==========================================================================================
'
' Procedure Author             : Ranuka
' Procedure Purpose            : N/A
' Procedure Date/Time          : 2008-04-06 @ 22:14
' Procedure Name               : Install
' Procedure Type               : Function
' Procedure Input              : N/A
' Procedure Output             : N/A
' Procedure Revision           : v1.0
' Procedure Description        : N/A
' Procedure Dependencies       : N/A
' Procedure Return Type Name   : String
'
'==========================================================================================

Dim RetVal      As Long
Dim AppInfo     As INSTALLERINFO

'// fill INSTALLERINFO structure to pass it to the DIFxAPI

AppInfo.pApplicationId = InstallerInfo_ApplicationId
AppInfo.pDisplayName = InstallerInfo_DisplayName
AppInfo.pProductName = InstallerInfo_ProductName
AppInfo.pMfgName = InstallerInfo_MfgName

'// call the DIFxAPI dll DriverPackageInstall method

RetVal = DriverPackageInstall(InfPath, Flag, AppInfo, NeedReboot)

If Not RetVal = ERROR_SUCCESS Then

'// driver installation failed, so get more detailed error information

Install = GetErrorMsg(RetVal)

Else    '// driver installation was a success, so return success

  Install = "Driver Installation Successful"
  
End If

End Function

Public Function UnInstall(InfPath As String, Flag As UninstallFlags, InstallerInfo_ApplicationId As String, InstallerInfo_DisplayName As String, InstallerInfo_ProductName As String, InstallerInfo_MfgName As String, NeedReboot As Boolean) As String

'==========================================================================================
'
' Procedure Author             : Ranuka
' Procedure Purpose            : N/A
' Procedure Date/Time          : 2008-04-06 @ 22:13
' Procedure Name               : UnInstall
' Procedure Type               : Function
' Procedure Input              : N/A
' Procedure Output             : N/A
' Procedure Revision           : v1.0
' Procedure Description        : N/A
' Procedure Dependencies       : N/A
' Procedure Return Type Name   : String
'
'==========================================================================================

Dim RetVal      As Long
Dim AppInfo     As INSTALLERINFO

'// fill INSTALLERINFO structure to pass it to the DIFxAPI

AppInfo.pApplicationId = InstallerInfo_ApplicationId
AppInfo.pDisplayName = InstallerInfo_DisplayName
AppInfo.pProductName = InstallerInfo_ProductName
AppInfo.pMfgName = InstallerInfo_MfgName

'// call the DIFxAPI dll DriverPackageUninstall method

RetVal = DriverPackageUninstall(InfPath, Flag, AppInfo, NeedReboot)

If Not RetVal = ERROR_SUCCESS Then

'//driver installation failed, so get more detailed error information

UnInstall = GetErrorMsg(RetVal)

Else    '// driver installation was a success, so return success

  UnInstall = "Driver UnInstall Successful"
  
End If

End Function

Public Function Preinstall(InfPath As String, Flag As InstallOrPreinstallFlags) As String

'==========================================================================================
'
' Procedure Author             : Ranuka
' Procedure Purpose            : N/A
' Procedure Date/Time          : 2008-04-06 @ 22:13
' Procedure Name               : Preinstall
' Procedure Type               : Function
' Procedure Input              : N/A
' Procedure Output             : N/A
' Procedure Revision           : v1.0
' Procedure Description        : N/A
' Procedure Dependencies       : N/A
' Procedure Return Type Name   : String
'
'==========================================================================================

Dim RetVal      As Long

'// call the DIFxAPI dll DriverPackagePreinstall method

RetVal = DriverPackagePreinstall(InfPath, Flag)

If Not RetVal = ERROR_SUCCESS Then

'//driver installation failed, so get more detailed error information

Preinstall = GetErrorMsg(RetVal)

Else    '// driver installation was a success, so return success

  Preinstall = "Driver Preinstall Successful"
  
End If

End Function

Private Function GetErrorMsg(ErrorCode As Long) As String

'==========================================================================================
'
' Procedure Author             : Ranuka
' Procedure Purpose            : N/A
' Procedure Date/Time          : 2008-04-06 @ 22:13
' Procedure Name               : GetErrorMsg
' Procedure Type               : Function
' Procedure Input              : N/A
' Procedure Output             : N/A
' Procedure Revision           : v1.0
' Procedure Description        : N/A
' Procedure Dependencies       : N/A
' Procedure Return Type Name   : String
'
'==========================================================================================

Select Case ErrorCode

Case CERT_E_EXPIRED

    GetErrorMsg = "Signing Certificate Is Expired"
    
    Exit Function

Case ERROR_NO_SUCH_DEVINST

    GetErrorMsg = "No Matching Device Found"
    
    Exit Function
    
Case ERROR_CANNOT_MAKE

    GetErrorMsg = "Driver Preinstall Failed For a Unspecified Reason"
    
    Exit Function
    
Case ERROR_UNSUPPORTED_TYPE

    GetErrorMsg = "Driver Package Type Is Not Supported"
    
    Exit Function
    
Case ERROR_INSTALL_FAILURE

    GetErrorMsg = "Installation Failed"
    
    Exit Function
    
Case ERROR_CANT_ACCESS_FILE

    GetErrorMsg = "Driver Package Files Could Not Be Accessed"
    
    Exit Function
    
Case ERROR_FILE_NOT_FOUND

    GetErrorMsg = "INF File Not Found"
    
    Exit Function
    
Case ERROR_NO_MORE_ITEMS

    GetErrorMsg = "Better Matching Driver Already Installed For The Device"
    
    Exit Function
    
Case ERROR_ALREADY_EXISTS

    GetErrorMsg = "Driver Package Is Already Preinstalled"
    
    Exit Function
    
Case ERROR_SHARING_VIOLATION

    GetErrorMsg = "A Component Of Drive Package Is Locked By A Thread Or Process"
    
    Exit Function
    
Case ERROR_ACCESS_DENIED

    GetErrorMsg = "Only Members Of The Administrators Group Can Access This Functionality"
    
    Exit Function
    
Case ERROR_BAD_ENVIRONMENT

    GetErrorMsg = "Driver Package Cannot Run On The Current Windows Version"
    
    Exit Function
    
Case ERROR_INVALID_PARAMETER

    GetErrorMsg = "A Supplied Parameter Is Not Valid"
    
    Exit Function
    
Case ERROR_INVALID_NAME

    GetErrorMsg = "Specified INF File Path Is Not Valid"
    
    Exit Function
    
Case ERROR_FILENAME_EXCED_RANGE

    GetErrorMsg = "Length, In Characters, Of The Specified INF File Path Is Greater Than The Maximum Supported Length"
    
    Exit Function
    
Case ERROR_OUTOFMEMORY

    GetErrorMsg = "Available System Memory Was Insufficient To Perform The Operation"
    
    Exit Function
    
Case ERROR_INVALID_FUNCTION

    GetErrorMsg = "Driver Package Is Not For A PnP Function Driver"
    
    Exit Function
        
Case Else

    GetErrorMsg = "Unknown Error, Installation Failed"
    
    Exit Function
    
End Select

End Function
