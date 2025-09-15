'Copyright (c) Microsoft Corporation. All rights reserved.
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
CONST wshOK                             =0
CONST VALUE_ICON_WARNING                =16
CONST wshYesNoDialog                    =4
CONST VALUE_ICON_QUESTIONMARK           =32
CONST VALUE_ICON_INFORMATION            =64
CONST HKEY_LOCAL_MACHINE                =&H80000002
CONST KEY_SET_VALUE                     =&H0002
CONST KEY_QUERY_VALUE                   =&H0001
CONST REG_SZ                            =1
CONST OfficeAppId                       = "0ff1ce15-a989-479d-af46-f275c6370663"
CONST STR_SYS32PATH                     = ":\Windows\System32\"
CONST STR_OSPPREARMPATH                 = "\Microsoft Office\Office16\OSPPREARM.EXE"
CONST STR_OSPPREARMPATH_DEBUG           = "\Microsoft Office Debug\Office16\OSPPREARM.EXE"
CONST REG_OSPP                          = "SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
CONST STR_HEARTBEATPATH                 = "\Microsoft\Office\Heartbeat\HeartbeatCache.xml"
CONST REG_SPP                           = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
CONST REG_SCAEXP_BASE                   = "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentEcs\"
CONST REG_SCAEXP_END                    =  "\Flights\ofc5juzsw45rcw0"
CONST REG_SCAEXP_OVERRIDE               = "HKLM\SOFTWARE\Microsoft\Office\Common\Experiment\ofc5juzsw45rcw0"
CONST REG_SCAO15                        = "HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration\SharedComputerLicensing"
CONST REG_SCAO16                        = "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration\SharedComputerLicensing"
CONST VALUE_SCAEXP_FLAVOR               = "exp"
CONST VALUE_SCAEXP_OVERRIDE_FLAVOR      = "expoverride"
CONST VALUE_SCAO15_FLAVOR               = "o15"
CONST VALUE_SCA016_FLAVOR               = "o16"
CONST VALUE_SCAREGKEY_SINGLE_TYPE       = "single"
CONST VALUE_SCAREGKEY_PER_APP_TYPE      = "perapp"
CONST VER_INFO                          = "Version Info: 2016 1.0"
'////////////////////////////////////////////////////////////////////////////////////////
CONST MSG_NOREGRIGHTS                   = "Insufficient rights to perform operation."
CONST MSG_ISCMD_ELEVATED                = "Ensure cmd.exe is elevated (right click > run as administrator)."
CONST MSG_CREDENTIALFAILURE             = "Connection failed with passed credentials."
CONST MSG_FILENOTFOUND                  = "File not found: "                
CONST MSG_SEPERATESMALL                 = "---------------"
CONST MSG_SEPERATE                      = "---------------------------------------"
CONST MSG_PROCESSING                    = "---Processing--------------------------"
CONST MSG_EXIT                          = "---Exiting-----------------------------"
CONST MSG_UNSUPPORTED                   = "Unsupported command passed."
CONST MSG_UNSUPPORTEDOPEROS7            = "The following command is supported on Windows 7 only: "
CONST MSG_UNSUPPORTEDOPEROS8            = "The following command is supported on Windows 8 and above only: "
CONST MSG_UNSUPPORTEDLOCAL              = "The following command is supported on local machine only: "
CONST MSG_CREDENTIALERR                 = "Passing credentials is not supported for this command."
CONST MSG_SUCCESS                       = "Successfully applied setting."
CONST MSG_NOKMSLICS                     = "No Office KMS licenses were found on the system."
CONST MSG_ACTATTEMPT                    = "Installed product key detected - attempting to activate the following product:"
CONST MSG_TOKACTATTEMPT                 = "Installed product key detected - attempting to token activate the following product:"
CONST MSG_NOKEYSINSTALLED               = "<No installed product keys detected>"
CONST MSG_UNINSTALLKEYSUCCESS           = "<Product key uninstall successful>"
CONST MSG_ACTSUCCESS                    = "<Product activation successful>"
CONST MSG_OFFLINEACTSUCCESS             = "<Offline product activation successful>"
CONST MSG_KEYINSTALLSUCCESS             = "<Product key installation successful>"
CONST MSG_PARTIALKEY                    = "Last 5 characters of installed product key: "
CONST MSG_UNINSTALLKEY                  = "Uninstalling product key for: "
CONST MSG_UNRECOGFILE                   = "Unrecognized file. Office licenses have an .xrm-ms file extension."
CONST MSG_INSTALLLICENSE                = "Installing Office license: "
CONST MSG_INSTALLLICSUCCESS             = "Office license installed successfully."
CONST MSG_SEARCHEVENTSKMS               = "Searching for KMS activation events on machine: "
CONST MSG_SEARCHEVENTSRET               = "Searching for Internet activation failure events on machine: "
CONST MSG_NOEVENTSSKMS                  = "No KMS activation events found on machine: "
CONST MSG_NOEVENTSRET                   = "No failure events found on machine: "
CONST MSG_OSPPSVC_NOINSTALL             = "Error: The Software Protection Platform service is not installed."
CONST MSG_OSPPSVC_NORUN                 = "Error: The Software Protection Platform service is not running."
CONST MSG_ERRPARTIALKEY                 = "The last 5 characters of an installed product key are required to run this option. Run the /dstatus option to display the partial product key."
CONST MSG_KEYNOTFOUND                   = "<Product key not found>"
CONST MSG_CMID                          = "Client Machine ID (CMID): "
CONST MSG_CMID_NOTFOUND                 = "Not found."
CONST MSG_NOLICENSEFOUND                = "<No licenses found>"
CONST MSG_AUTHERR                       = "Authorization Error: 0x"
CONST MSG_REMILID                       = "Removed Token-based Activation License with License ID (ILID): "
CONST MSG_NOTFOUNDILID                  = "License not found with License ID (ILID): "
CONST MSG_KMSLOOKUP                     = "KMS Lookup Domain: "
CONST MSG_INFO_ONLY                     = " (for information purposes only as the status is licensed)"
CONST MSG_ACT_ERROR_FOUND_KB            = "NOTICE: A KB article has been detected for activation failure: "
CONST MSG_ACT_ERROR_KB_LINK             = "FOR MORE INFORMATION PLEASE VISIT: https://support.microsoft.com/kb/2870357#Error0x"
CONST MSG_SCAEXP                        = "SCA Mode Experiment key "
CONST MSG_SCAEXP_OVERRIDE               = "SCA Mode Experiment Override key "
CONST MSG_SCAO15                        = "SCA Mode for O15 key "
CONST MSG_SCAO16                        = "SCA Mode for O16 key "
CONST MSG_SCAREGKEY_ACTIVE              = "found ENABLED."
CONST MSG_SCAREGKEY_INACTIVE            = "found DISABLED."
'////////////////////////////////////////////////////////////////////////////////////////
CONST MSG_VLActivationType              = "Activation Type Configuration: "
'////////////////////////////////////////////////////////////////////////////////////////
CONST MSG_Act_Recent                    = "Most recent successful activation client information: "
CONST MSG_KMS_DNS                       = "KMS machine name from DNS: "
CONST MSG_KMS_DNS_ERR                   = "DNS auto-discovery: KMS name not available"
CONST MSG_ADInfoAOName                  = "Activation Object name: "
CONST MSG_ADInfoAODN                    = "AO DN: "
CONST MSG_ADInfoExtendedPid             = "AO extended PID: "
CONST MSG_ADInfoActID                   = "AO activation ID: "
CONST MSG_ACTIVATION_INTERVAL           = "Activation Interval: "
CONST MSG_RENEWAL_INTERVAL              = "Renewal Interval: "
CONST MSG_HOST_CACHING                  = "KMS host caching: "
CONST MSG_HOST_REG_OVERRIDE             = "KMS machine registry override defined: "
CONST MSG_DEFAULT_PORT                  = "1688"
'////////////////////////////////////////////////////////////////////////////////////////
CONST MSG_PID                           = "PRODUCT ID: "
CONST MSG_SKUID                         = "SKU ID: "
CONST MSG_LICENSENAME                   = "LICENSE NAME: "
CONST MSG_DESCRIPTION                   = "LICENSE DESCRIPTION: "
CONST MSG_LICSTATUS                     = "LICENSE STATUS: "
CONST MSG_LICENSED                      = " ---LICENSED--- "
CONST MSG_UNLICENSED                    = " ---UNLICENSED--- "
CONST MSG_OOBGRACE                      = " ---OOB_GRACE--- "
CONST MSG_OOTGRACE                      = " ---OOT_GRACE--- "
CONST MSG_NONGENGRACE                   = " ---NON_GENUINE_GRACE--- "
CONST MSG_NOTIFICATION                  = " ---NOTIFICATIONS--- "
CONST MSG_EXTENDEDGRACE                 = " ---EXTENDED GRACE--- "
CONST MSG_LICUNKNOWN                    = " ---UNKNOWN--- "
CONST MSG_REMAINGRACE                   = "REMAINING GRACE: "
CONST MSG_LICEXPIRY                     = "BETA EXPIRATION: "
CONST MSG_ERRCODE                       = "ERROR CODE: "
CONST MSG_ERRDESC                       = "ERROR DESCRIPTION: "
CONST MSG_ERRUNKNOWN                    = "An unknown error occurred."
CONST MSG_ERRCODEVALUE                  = "An error code must start with '0x'. Example: 0xC004F009"
'////////////////////////////////////////////////////////////////////////////////////////
CONST MSG_NOEVENTS                      = "No events found."
CONST MSG_HEARTBEAT_LOCALONLY           = "HEARTBEATCACHE.XML SUPPORTED FOR LOCAL MACHINE ONLY."
CONST MSG_SEARCHFOR                     = "SEARCHING FOR INSTANCES OF "
CONST MSG_SEARCH_FORKEY                 = "SEARCHING FOR OFFICE PRODUCT KEYS ON MACHINE: "
CONST MSG_EVENT_1013                    = "ACTIVATION SUCCESS EVENT: "
CONST MSG_EVENT_8200                    = "ACTIVATION FAILURE EVENT: "
CONST MSG_EVENT_1016                    = "INSTALL PRODUCT KEY SUCCESS EVENT: "
CONST MSG_EVENT_1017                    = "INSTALL PRODUCT KEY FAILURE EVENT: "
CONST MSG_EVENT_2011                    = "OLC-OLS EVENT: "
CONST MSG_NOKEYSINSTALLED_SUB           = "<<<<< NO INSTALLED -SUBSCRIPTION- PRODUCT KEYS DETECTED >>>>>"
CONST MSG_NOKEYSINSTALLED_PERP          = "<<<<< NO INSTALLED -PERPETUAL- PRODUCT KEYS DETECTED >>>>"
CONST MSG_KEYSINSTALLED_SUB             = "<<<<< -SUBSCRIPTION- PRODUCT KEY DETECTED. REPORTING ON ADDITIONAL PROPERTIES >>>>>"
CONST MSG_KEYSINSTALLED_PERP            = "<<<<< -PERPETUAL- PRODUCT KEY DETECTED. REPORTING ON ADDITIONAL PROPERTIES >>>>>"
CONST MSG_SCAREGKEYS                    = "SCA REGISTRY KEYS"
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
On Error Resume Next

Set WshShell = WSCript.CreateObject("WSCript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objNetwork = WSCript.CreateObject("WSCript.Network")

Dim globalResource, globalErr, foundSlUi, strSluiPath, strLocal, objWMI, objWMI1, wmiErr, productinstances, strValue, Win7, productClass, tokenClass, intIsKms, kmsCounter, isAdActivated, dlicense, foundSubKey, foundPerpKey, Win8, strSkuId, errorKBs, getHeartBeat
globalResource = ""
globalErr = ""
foundSlUi = False
Win7 = False
Win8 = False
kmsCounter = 0
isAdActivated = False
dlicense = False
strSkuId = Null

currentDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))

' Activation error codes for which a KB is available
errorKBs = "80070422|80070426|C004F074|80070001|80070005|8007000D|8007232B|8007251D|C004F014|C004F038|C004F039|C004F041|C004F042|C004C003|4004F040"

Select Case WSCript.Arguments.Count
    Case 0
        verifyFileExists currentDir & "ospp.htm"
        showIePopUp currentDir & "ospp.htm"
        WScript.Quit
    Case 1
        var1 = WSCript.Arguments(0)
    Case 2
        var1 = WSCript.Arguments(0)
        var2 = WSCript.Arguments(1)
    Case 3
        var1 = WSCript.Arguments(0)
        var2 = WSCript.Arguments(1)
        var3 = WSCript.Arguments(2)
    Case 4
        var1 = WSCript.Arguments(0)
        var2 = WSCript.Arguments(1)
        var3 = WSCript.Arguments(2)
        var4 = WSCript.Arguments(3)
    Case Else
End Select
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Call Main(var1,var2,var3,var4)
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Sub Main(strCommand,strMachine,strUser,strPassword)

On Error Resume Next

getEngine()
pProcessing()
getSlui()
strLocal = objNetwork.ComputerName 
strCommand = LCase(strCommand)

Select Case strCommand
    Case "/act", "/dstatus", "/dstatusall", "/dinstid", "/dtokils", _
        "/remhst", "/stokflag", "/ctokflag", "/dcmid", "/dtokcerts", "/ckms-domain", "/dpid", "/actsub", "/dstatussub"
        connectWMI strMachine,strUser,strPassword,""
        performLicAction strCommand,"",""

        If strCommand = "/dstatussub" Then
            getHeartbeatXml strMachine
            quitExit()
        End If
    Case "/dlicenseinfo"
        If strMachine <> "" Then
            WScript.Echo MSG_SEARCH_FORKEY & strMachine
        Else
            WScript.Echo MSG_SEARCH_FORKEY & strLocal 
        End If
        
        WScript.Echo MSG_SEPERATE
        
        'Connect WMI
        dlicense = True
        connectWMI strMachine,strUser,strPassword,""
        
        'Display licensestatus for all product keys
        performLicAction "/dstatusall","",""
        
        'Display PID/MachineID for sub keys
        performLicAction "/dpid","",strMachine
        
        WScript.Echo MSG_SEPERATE
        
        'Display MachineId
        getMachineId strMachine
        
        'Display HeartBeatXml
        getHeartbeatXml strMachine
        
        'Display OLC-OLS events
        getEvents MSG_EVENT_2011,"Application","2011","Microsoft Office 15",strMachine
        
        'Display SPP related events: InstallProof & SLActivate      
        If Win7 = True Then
            getEvents MSG_EVENT_1016,"Application","1016","Office Software Protection Platform Service",strMachine
            getEvents MSG_EVENT_1017,"Application","1017","Office Software Protection Platform Service",strMachine
            getEvents MSG_EVENT_1013,"Application","1013","Office Software Protection Platform Service",strMachine
            getEvents MSG_EVENT_8200,"Application","8200","Office Software Protection Platform Service",strMachine
        Else
            getEvents MSG_EVENT_1016,"Application","1016","Microsoft-Windows-Security-SPP",strMachine
            getEvents MSG_EVENT_1017,"Application","1017","Microsoft-Windows-Security-SPP",strMachine
            getEvents MSG_EVENT_1013,"Application","1013","Microsoft-Windows-Security-SPP",strMachine
            getEvents MSG_EVENT_8200,"Application","1014","Microsoft-Windows-Security-SPP",strMachine
        End If

        WScript.Echo MSG_SEPERATE

        'Display SCA mode status
        checkScaStatus

        quitExit()
    Case "/dhistoryacterr", "/dhistorykms"
        connectWMI strMachine,strUser,strPassword,""
        performLicAction strCommand,"",strMachine
    Case "/puserops", "/duserops"
        connectWMI strMachine,strUser,strPassword,"reg"
        performRegAction strCommand
    Case "/osppsvcrestart", "/osppsvcauto"
        connectWMI strMachine,strUser,strPassword,""
        performServiceAction strCommand
    Case "/help", "help", "?", "/?", "/?"
        verifyFileExists currentDir & "ospp.htm"
        showIePopUp currentDir & "ospp.htm"
        quitExit()
    Case "/regmof"
        registerMof "osppwmi.mof"
    Case "/rearm"
        If strMachine = "" Then
            reARM ""
        Else
            globalPopFailure MSG_UNSUPPORTEDLOCAL & vbCr & strCommand,True
        End If
        quitExit()
    Case "/version"
        globalPopSuccess VER_INFO,True
    Case Else
        pos = InStr(strCommand,":")
        
        Select Case pos
            Case 6
                getCommand = Left(strCommand,5)
            Case 7
                getCommand = Left(strCommand,6)
            Case 8
                getCommand = Left(strCommand,7)
            Case 9
                getCommand = Left(strCommand,8)
            Case 13
                getCommand = Left(strCommand,12)
            Case Else
                globalPopFailure MSG_UNSUPPORTED,True
        End Select
        
        Select Case getCommand
            Case "/skms-domain", "/actype", "/inpkey", "/unpkey", "/inslic", "/actcid", "/sethst", "/setprt", "/ddescr", "/rtokil", "/tokact", "/cachst", "/rearm", "/pstatus", "/actp"
                strValue = Replace(strCommand,getCommand & ":","")
                If strValue = "" Then
                    globalPopFailure MSG_UNSUPPORTED & " A value is required for: " & strCommand,True
                End If
                
                If getCommand = "/ddescr" Then
                    If Left(strValue,2) = "0x" Then
                        getDescription strValue,""
                    Else
                        WScript.Echo MSG_ERRCODEVALUE
                        quitExit()
                    End If
                ElseIf getCommand = "/rearm" Then
                    If strMachine = "" Then
                        reARM strValue
                Else
                    globalPopFailure MSG_UNSUPPORTEDLOCAL & vbCr & strCommand,True
                End If
                quitExit()
                Else
                    connectWMI strMachine,strUser,strPassword,""
                    performLicAction getCommand,strValue,""
                    If getCommand = "/pstatus" And getHeartBeat = True Then
                        getHeartbeatXml strMachine
                    End If
                    quitExit()
                End If
            Case Else
                globalPopFailure MSG_UNSUPPORTED,True
        End Select
End Select

End Sub
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function showIePopUp(strPath)

On Error Resume Next

Set objExplorer = CreateObject("InternetExplorer.Application")
    With objExplorer
            .Navigate strPath
            .ToolBar = 0
            .StatusBar = 0
            .Width = 1000
            .Height = 593 
            .Left = 1
            .Top = 1
            .Visible = 1
    End With
        
End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function getEngine()

strEngine = LCase(Right(WScript.FullName,12))
If strEngine <> "\cscript.exe" Then
    WshShell.Popup "Unable to perform operation. " & WSCript.ScriptName & " requires the cscript engine." & _
     vbCr & "Command line example: cscript ospp.vbs ?", _
    ,WSCript.ScriptName, VALUE_ICON_WARNING
    WScript.Quit
End If

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function WMIDateStringToDate(dtmEventDate)

WMIDateStringToDate = CDate(Mid(dtmEventDate, 5, 2) & "/" & _
Mid(dtmEventDate, 7, 2) & "/" & Left(dtmEventDate, 4) _
& " " & Mid (dtmEventDate, 9, 2) & ":" & _
Mid(dtmEventDate, 11, 2) & ":" & Mid(dtmEventDate, _
13, 2))

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function getDescription(strSearch,cType)

If foundSlUi <> True Then
    If cType <> "wmi" Then
        globalPopFailure "slui.exe not found.",True
        quitExit()
    End If
Else
    Set objScriptExec = WshShell.Exec (strSluiPath & " 0x2a " & strSearch)
    readOut = objScriptExec.StdOut.ReadAll
    quitExit()
End If

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function checkRegRights(wmiObject,strKeyPath)

On Error Resume next

wmiObject.CheckAccess HKEY_LOCAL_MACHINE, strKeyPath, KEY_SET_VALUE, _
    bHasAccessRight

If bHasAccessRight = True Then
    'Success
Else
    globalPopFailure MSG_NOREGRIGHTS & vbCr & MSG_ISCMD_ELEVATED,True
End If   

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function quitExit()

Set WshShell = Nothing
Set objFSO = Nothing
Set objNetwork = Nothing
Set objWMI = Nothing

WScript.Echo MSG_SEPERATE
WScript.Echo MSG_EXIT
WSCript.Quit

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function verifyFileExists(file)

If Not objFSO.FileExists(file) Then
    If file = currentDir & "slerror.xml" Then
        WScript.Echo "[" & MSG_FILENOTFOUND & file &  "  Unable to display error description.]"
    ElseIf file = currentDir & "ospp.htm" Then
        globalPopFailure MSG_FILENOTFOUND & vbCr & file,False
        quitExit()
    Else
        globalPopFailure MSG_FILENOTFOUND & vbCr & file,True
    End If
End If

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function registerMof(strFile)

For Each Drv In objFSO.Drives
    If Drv.DriveType=2 Then
        If objFSO.FileExists(Drv.DriveLetter & STR_SYS32PATH & "wbem\mofcomp.exe") Then
            foundComp = True
            strMofExePath = Drv.DriveLetter & STR_SYS32PATH & "wbem\mofcomp.exe"
            If objFSO.FileExists(Drv.DriveLetter & STR_SYS32PATH & "wbem\" & strFile) Then
                foundMof = True
                strOWmi = Drv.DriveLetter & STR_SYS32PATH & "wbem\" & strFile
                Set objScriptExec = WshShell.Exec (strMofExePath & " " & strOWmi)
                readOut = objScriptExec.StdOut.ReadAll
                WScript.Echo readOut
                quitExit()
            End If
        End If
    End If
Next

If foundComp <> True Then
    globalPopFailure MSG_FILENOTFOUND & Replace(STR_SYS32PATH,":","") & "wbem\mofcomp.exe",True
Else
    If foundMof <> True Then
        globalPopFailure MSG_FILENOTFOUND & Replace(STR_SYS32PATH,":","") & "wbem\osppwmi.mof",True
    End If
End If

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function pProcessing()

WScript.Echo MSG_PROCESSING
WScript.Echo MSG_SEPERATE
    
End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function getSlui()

For Each Drv In objFSO.Drives
    If Drv.DriveType=2 Then
        If objFSO.FileExists(Drv.DriveLetter & STR_SYS32PATH & "slui.exe") Then
            strSluiPath = Drv.DriveLetter & STR_SYS32PATH & "slui.exe"
            foundSlUi = True
            Exit For
        End If
    End If
Next

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
' Returns the encoding for a givven file.
' Possible return values: ascii, unicode, unicodeFFFE (big-endian), utf-8
Function GetFileEncoding(strFileName)
    Dim strData
    Dim strEncoding

    Set oStream = CreateObject("ADODB.Stream")

    oStream.Type = 1 'adTypeBinary
    oStream.Open
    oStream.LoadFromFile(strFileName)

    ' Default encoding is ascii
    strEncoding =  "ascii"

    strData = BinaryToString(oStream.Read(2))

    ' Check for little endian (x86) unicode preamble
    If (Len(strData) = 2) and strData = (Chr(255) + Chr(254)) Then
        strEncoding = "unicode"
    Else
        oStream.Position = 0
        strData = BinaryToString(oStream.Read(3))

        ' Check for utf-8 preamble
        If (Len(strData) >= 3) and strData = (Chr(239) + Chr(187) + Chr(191)) Then
            strEncoding = "utf-8"
        End If
    End If

    oStream.Close

    GetFileEncoding = strEncoding
    
End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
' Converts binary data (VT_UI1 | VT_ARRAY) to a string (BSTR)
Function BinaryToString(dataBinary)  
    Dim i
    Dim str

    For i = 1 To LenB(dataBinary)
        str = str & Chr(AscB(MidB(dataBinary, i, 1)))
    Next

    BinaryToString = str
    
End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
' Returns string containing the whole text file data. 
' Supports ascii, unicode (little-endian) and utf-8 encoding.
Function ReadAllTextFile(strFileName)
    Dim strData
    Set oStream = CreateObject("ADODB.Stream")

    oStream.Type = 2 'adTypeText
    oStream.Open
    oStream.Charset = GetFileEncoding(strFileName)
    oStream.LoadFromFile(strFileName)

    strData = oStream.ReadText(-1) 'adReadAll

    oStream.Close

    ReadAllTextFile = strData
    
End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function sppErrHandle(strCommand)

globalErr = Hex(Err.Number)

Select Case Err.Number
    Case 0
        'Success
        Select Case strCommand
            Case "/act","/tokact","/actsub","/actp"
                WScript.Echo MSG_ACTSUCCESS
            Case "/inpkey"
                WScript.Echo MSG_KEYINSTALLSUCCESS
                quitExit()
            Case "/inslic"
                WScript.Echo MSG_INSTALLLICSUCCESS
                quitExit()
            Case "/ckms-domain","/skms-domain","/actype","/sethst","/setprt","/remhst","/stokflag","/ctokflag","/cachst"
                WScript.Echo MSG_SUCCESS
                quitExit()
            Case "/rtokil"
                WScript.Echo MSG_REMILID & UCase(strValue)
                quitExit()
            Case "/unpkey"
                WScript.Echo MSG_UNINSTALLKEYSUCCESS
                quitExit()
            Case Else
        End Select
    Case Else
        verifyFileExists currentDir & "slerror.xml"
        getResource("err" & "0x" & globalErr)
        If globalResource = "" Then
            If Len(globalErr) <> "8" Then
                WScript.Echo MSG_ERRDESC & MSG_ERRUNKNOWN & " (0x" & globalErr & ")"
            Else
                If foundSlUi = True Then
                    WScript.Echo MSG_ERRCODE & "0x" & globalErr
                    WScript.Echo MSG_ERRDESC & "Run the following: cscript ospp.vbs /ddescr:0x" & globalErr
                Else
                    WScript.Echo MSG_ERRCODE & "0x" & globalErr 
                End If
            End If
            If strCommand <> "/act" And strCommand <> "/actsub" And strCommand <> "/actp" Then
                quitExit()
            End If
        Else
            WScript.Echo MSG_ERRCODE & "0x" & globalErr 
            Wscript.Echo MSG_ERRDESC & globalResource
        End If
        
        If strCommand = "/dtokcerts" Or strCommand = "/ignore" Then
            quitExit()
        End If
End Select

If globalErr = "C004F074" Then
    WScript.Echo "To view the activation event history run: cscript " & WScript.ScriptName & " /dhistorykms"
End If

If strCommand = "/act" And globalErr <> "0" Then
    ' If a KB article is found, show the KB link
    lookupKBArticle(globalErr)
End If

globalResource = ""
globalErr = ""
Err.Clear

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function wmiErrHandle()

Select Case Err.Number
    Case 0
        'Successs
    Case 424
        globalPopFailure MSG_ERRCODE & Err.Number & vbCr & MSG_ERRDESC & MSG_CREDENTIALFAILURE,True            
    Case Else
        If Err.Description <> "" Then
            globalPopFailure MSG_ERRCODE & Err.Number & vbCr & MSG_ERRDESC & Err.Description,True
        Else
            globalPopFailure "An error occurred while making the connection." & vbCr & MSG_ERRCODE & Err.Number,True
        End If
End Select

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function setRegValue(wmiObject,opsValue,strValueName)

On Error Resume Next

Err.Clear()
If Win7 = True Then
    strKeyPath = REG_OSPP
Else
    strKeyPath = REG_SPP
End If

Select Case strValueName
    Case "UserOperations"
        wmiObject.CreateKey HKEY_LOCAL_MACHINE,strKeyPath
        wmiObject.SetDWORDValue HKEY_LOCAL_MACHINE,_
            strKeyPath,strValueName,opsValue
    Case Else
End Select

wmiErrHandle()
WScript.Echo MSG_SUCCESS
quitExit()

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function getResource(resource)

On Error Resume Next
Set xmlDoc = CreateObject("Msxml2.DOMDocument.6.0") 
xmlDoc.load(currentDir & "slerror.xml") 
Set ElemList = xmlDoc.getElementsByTagName(resource) 
resValue = ElemList.item(0).text
globalResource = resValue 

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function globalPopSuccess(strSuccess,boolQuit)

If boolQuit = True Then
    WshShell.Popup strSuccess,,WScript.ScriptName, wshOK + VALUE_ICON_INFORMATION
    quitExit()
Else
    WshShell.Popup strSuccess,,WScript.ScriptName, wshOK + VALUE_ICON_INFORMATION
End If

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function globalPopFailure(strFailure,boolQuit)

If boolQuit = True Then
    WshShell.Popup strFailure,,WScript.ScriptName, wshOK + VALUE_ICON_WARNING
    quitExit()
Else
    WshShell.Popup strFailure,,WScript.ScriptName, wshOK + VALUE_ICON_WARNING
End If

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function connectWMI(strMachine,strUser,strPassword,ctype)

On Error Resume Next

If ctype = "" Then
    If strMachine = "" Or LCase(strMachine) = LCase(strLocal) Then
        Set objWMI = GetObject("winmgmts:" _
                & "{impersonationLevel=impersonate}!\\" & "." & "\root\cimv2")
    Else
        If strUser = "" And strPassword = "" Then
            Set objWMI = GetObject("winmgmts:" _
                & "{impersonationLevel=impersonate}!\\" & strMachine & "\root\cimv2")
        Else
            Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
            Set objWMI = objSWbemLocator.ConnectServer _
                (strMachine, "\root\cimv2", strUser, strPassword)
            wmiErr = CStr(Hex(Err.Number))
            If Len(wmiErr) = "8" Then
                getDescription "0x" & wmiErr,"wmi"
            End If
            objWMI.Security_.ImpersonationLevel = 3
        End If
    End If
Else
    If strUser <> "" Then
        globalPopFailure MSG_CREDENTIALERR,True
    End If

    If strMachine = "" Or LCase(strMachine) = LCase(strLocal) Then
        Set objWMI1 = GetObject("winmgmts:" _
            & "{impersonationLevel=impersonate}!\\" & "." & "\root\default:StdRegProv")
            
        Set objWMI = GetObject("winmgmts:" _
                & "{impersonationLevel=impersonate}!\\" & "." & "\root\cimv2")
    Else
        Set objWMI1 = GetObject("winmgmts:" _
            & "{impersonationLevel=impersonate}!\\" & strMachine & "\root\default:StdRegProv")
            
        Set objWMI = GetObject("winmgmts:" _
                & "{impersonationLevel=impersonate}!\\" & strMachine & "\root\cimv2")
    End If
End If

wmiErrHandle()
setWinOS()

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Private Function TkaGetSigner()

On Error Resume Next

    If Win7 = True Then 
        Set TkaGetSigner = WScript.CreateObject("OSPPWMI.OSppWmiTokenActivationSigner")
    Else
        Set TkaGetSigner = WScript.CreateObject("SPPWMI.SppWmiTokenActivationSigner")
    End If
    
    If Hex(Err.Number) = "80020009" Then
        globalPopFailure MSG_ERRCODE & "0x" & Hex(Err.Number) & vbCr & MSG_ERRDESC & Err.Description,True
    End If

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function TkaPrintCertificate(strThumbprint)

    arrParams = Split(strThumbprint, "|")
    WScript.Echo "Thumbprint: " & arrParams(0)
    WScript.Echo "Subject: " & arrParams(1)
    WScript.Echo "Issuer: " & arrParams(2)
    vf = FormatDateTime(CDate(arrParams(3)), vbShortDate)
    WScript.Echo "Valid From: " & vf
    vt = FormatDateTime(CDate(arrParams(4)), vbShortDate)
    WScript.Echo "Valid To: " & vt
    WScript.Echo MSG_SEPERATE
    
End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function ExecuteQuery(strSelect,strWhere,strClass)
    
Err.Clear
    
If strWhere = "" Then
    Set productinstances = objWMI.ExecQuery("SELECT " & strSelect & " FROM " & strClass)
Else
    Set productinstances = objWMI.ExecQuery("SELECT " & strSelect & " FROM " & strClass & " WHERE " & strWhere)
End If
    
sppErrHandle ""

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function performLicAction(strCommand,strValue,strMachine)

On Error Resume Next

If strCommand = "/dhistorykms" Or strCommand = "/dhistoryacterr" Then
    verifyFileExists currentDir & "slerror.xml"
    If strCommand = "/dhistorykms" Then
        '12288 = KMS Activation event id
        eventCode = "12288"
        strSrcEvents = MSG_SEARCHEVENTSKMS
        strNoEvents = MSG_NOEVENTSSKMS
    Else
        '8200 = Internet Activation event id
        eventCode = "8200"
        strSrcEvents = MSG_SEARCHEVENTSRET
        strNoEvents = MSG_NOEVENTSRET
    End If
    
    If strMachine <> "" Then
        WScript.Echo strSrcEvents & strMachine
    Else
        WScript.Echo strSrcEvents & strLocal
    End If
    
    WScript.Echo "Event ID: " & eventCode
    WScript.Echo vbCr
    Set objEvents = objWMI.ExecQuery _
        ("Select * from Win32_NTLogEvent Where Logfile = 'Application' and " _
        & "EventCode = '" & eventCode & "'")
        If objEvents.Count > 0 Then
            For each objEvent in objEvents
                If strCommand = "/dhistoryacterr" Then
                    i = i + 1
                    dtmEventDate = objEvent.TimeWritten
                    strTimeWritten = WMIDateStringToDate(dtmEventDate)
                    WScript.Echo "Coordinated Universal Time Written: " & strTimeWritten
                    strReplCrs = Replace(objEvent.Message,vbCrLf,"")
                    WScript.Echo "MESSAGE: " & strReplCrs
                    strhr10 = Right(strReplCrs,10)
                    getResource("err" & strhr10)
                    If globalResource = "" Then
                        If foundSlUi = True Then
                            WScript.Echo MSG_ERRDESC & "Run the following: cscript ospp.vbs /ddescr:" & strhr10
                        Else
                            WScript.Echo MSG_ERRDESC & "Not available."
                        End If
                    Else
                        Wscript.Echo MSG_ERRDESC & globalResource
                    End If
                    WScript.Echo MSG_SEPERATE        
                Else
                    strhr10 = Mid(objEvent.Message,90,10)
                    strReplCrs = Replace(objEvent.Message,vbCrLf,"")
                    If Right(strReplCrs,2) = " 5" Then
                        strReplStrs = Replace(strReplCrs,"The client has sent an activation request to the key management service machine.Info:","")
                        dtmEventDate = objEvent.TimeWritten
                        strTimeWritten = WMIDateStringToDate(dtmEventDate)
                        WScript.Echo "Coordinated Universal Time Written: " & strTimeWritten
                        intColon = InStr(strReplStrs,":")
                        strErrHost = Left(strReplStrs,intColon)
                        strErrHost = Trim(strErrHost)
                        strErrHost = Replace(strErrHost,":","")
                        WScript.Echo "ERROR/HOST: " & strErrHost
                        Select Case strhr10
                            Case "0x00000000"
                                WScript.Echo MSG_ERRDESC & "N/A"
                            Case Else
                                getResource("err" & strhr10)
                                If globalResource = "" Then
                                    If foundSlUi = True Then
                                        WScript.Echo MSG_ERRDESC & "Run the following: cscript ospp.vbs /ddescr:" & strhr10
                                        ' If a KB article is found, show the KB link
                                        lookupKBArticle(Right(strhr10, 8))
                                    Else
                                        WScript.Echo MSG_ERRDESC & "Not available."
                                    End If
                                Else
                                    Wscript.Echo MSG_ERRDESC & globalResource
                                    ' If a KB article is found, show the KB link
                                    lookupKBArticle(Right(strhr10, 8))
                                End If
                        End Select
                        WScript.Echo MSG_SEPERATE
                    End If
                End If
            Next
        Else
            WScript.Echo MSG_SEPERATE
            If strMachine <> "" Then
                WScript.Echo strNoEvents & strMachine
            Else
                WScript.Echo strNoEvents & strLocal
            End If
            WScript.Echo MSG_SEPERATE
        End If
        quitExit()
End If

'Verify osppsvc service is installed for win7 case
If Win7 = True Then
    Set colListOfServices = objWMI.ExecQuery _
        ("Select * from Win32_Service ")
    For Each objService in colListOfServices
        If objService.Name = "osppsvc" Then
            installed = True
            Exit For
        End If
    Next
        
    If installed <> True Then
        globalPopFailure MSG_OSPPSVC_NOINSTALL,True
    End If
End If
        
Select Case strCommand
    'The following operations are performed @ a service level
    Case "/inpkey", "/dcmid", "/inslic", "/cachst", "/stokflag", "/ctokflag", "/dstatus", "/dstatusall", "/dpid", "/dstatussub", "/pstatus"
        If Win7 = True Then
            For Each objService in objWMI.InstancesOf("OfficeSoftwareProtectionService")
                Set objOspp = objService
                Exit For
            Next
        Else
            'Win8 and beyond
            For Each objService in objWMI.InstancesOf("SoftwareLicensingService")
                Set objOspp = objService
                Exit For
            Next
        End If
    Case Else
End Select

sppErrHandle ""

If strCommand = "/inpkey" Then
    i = i + 1
    Err.Clear
    objOspp.InstallProductKey(strValue)
    sppErrHandle(strCommand)
ElseIf strCommand = "/cachst" Then
    i = i + 1
    If strValue = "true" Then
        objOspp.DisableKeyManagementServiceHostCaching(False)
        sppErrHandle(strCommand)
    ElseIf strValue = "false" Then
        objOspp.DisableKeyManagementServiceHostCaching(True)
        sppErrHandle(strCommand) 
    Else
        globalPopFailure MSG_UNSUPPORTED & " A TRUE or FALSE value is required for: " & strCommand,True
    End If
ElseIf strCommand = "/dcmid" Then
    If Win8 = True Then
        'Check if KMS key installed. If yes retrieve the SKUID value.
        ExecuteQuery "ID, ApplicationId, Description, Name","ApplicationId = '" & OfficeAppId & "' " & "AND PartialProductKey <> NULL ",productClass
        For each instance in productinstances
            sppErrHandle ""
            intIsKms = InStr(UCase(instance.Description),"KMS")
            If intIsKms <> 0 Then
                strSkuId = instance.ID
                Exit For
            End If
        Next
        
        If strSkuId = Null Then 
            WScript.Echo MSG_CMID & MSG_CMID_NOTFOUND
            quitExit()
        End If
                
        'Return the last Office KMS event containing containing the SKUID value for the installed KMS key.
        '& retrieve the CMID value from the event.
        oKmsEventCounter = 0
        eventCode = "12288"
        Set objEvents = objWMI.ExecQuery _
        ("Select * from Win32_NTLogEvent Where Logfile = 'Application' and " _
        & "EventCode = '" & eventCode & "'")
        If objEvents.Count > 0 Then
            For each objEvent in objEvents
                oKmsEvent = InStr(objEvent.Message,strSkuId)
                If oKmsEvent <> 0 Then
                    oKmsEventCounter = oKmsEventCounter + 1
                    parseEvntMsg1 = InStr(objEvent.Message, ", ")
                    parseEvntMsg2 = InStr(parseEvntMsg1 + 2,objEvent.Message, ", ")
                    parseEvntMsg3 = InStr(parseEvntMsg2 + 2,objEvent.Message, ", ")
                    parseEvntMsg4 = InStr(parseEvntMsg3 + 2,objEvent.Message, ", ") 
                    WScript.Echo MSG_CMID & Mid(objEvent.Message,parseEvntMsg3 + 2,parseEvntMsg4 - (parseEvntMsg3 + 2))
                    Exit For
                End If
            Next
        End If
        
        If oKmsEventCounter = 0 Then
            WScript.Echo MSG_CMID & MSG_CMID_NOTFOUND
        End If
    Else
        If objOspp.ClientMachineID <> "" Or objOspp.ClientMachineID <> Null Then
            WScript.Echo MSG_CMID & objOspp.ClientMachineID
        Else
            WScript.Echo MSG_CMID & MSG_CMID_NOTFOUND
        End If
    End If
    quitExit()

ElseIf strCommand = "/inslic" Then
    i = i + 1
    If Right(strValue,7) = ".xrm-ms" Then
        verifyFileExists strValue
        WScript.Echo MSG_INSTALLLICENSE & strValue
    Else
        globalPopFailure MSG_UNRECOGFILE,True
    End If
    LicenseData = ReadAllTextFile(strValue)
    objOSpp.InstallLicense(LicenseData)
    SppErrHandle(strCommand)
ElseIf strCommand = "/stokflag" Then
    i = i + 1
    If Win7 = True Then
        objOspp.DisableKeyManagementServiceActivation(True)
        sppErrHandle(strCommand)
    Else
        'Unsupported - osppsvc only supports this.
        globalPopFailure MSG_UNSUPPORTEDOPEROS7 & vbCr & strCommand,True
    End If
ElseIf strCommand = "/ctokflag" Then
    i = i + 1
    If Win7 = True Then
        objOspp.DisableKeyManagementServiceActivation(False)
        SppErrHandle(strCommand)
    Else
        'Unsupported - osppsvc only supports this.
        globalPopFailure MSG_UNSUPPORTEDOPEROS7 & vbCr & strCommand,True
    End If
ElseIf strCommand = "/dtokils" Then
    Err.Clear
    Set objWmiDate = CreateObject("WBemScripting.SWbemDateTime")
    ExecuteQuery "ILID, ILVID, AuthorizationStatus, ExpirationDate, Description, AdditionalInfo","",tokenClass
    
    For Each instance in productinstances
        sppErrHandle ""
        i = i + 1
        WScript.Echo "License ID (ILID): " & instance.ILID
        WScript.Echo "Version ID (ILvID): " & instance.ILVID
        If Not IsNull(instance.ExpirationDate) Then
            objWmiDate.Value = instance.ExpirationDate
            If (objWmiDate.GetFileTime(false) <> 0) Then
                WScript.Echo "Expiry Date: " & objWmiDate.GetVarDate
            End If
        End If
        If Not IsNull(instance.AdditionalInfo) Then
            WScript.Echo "Additional Info: " & instance.AdditionalInfo
        End If
        If Not IsNull(instance.AuthorizationStatus) And instance.AuthorizationStatus <> 0 Then
            globalErr = CStr(Hex(instance.AuthorizationStatus))
            WScript.Echo MSG_AUTHERR & globalErr
            quitExit()
        Else            
            WScript.Echo "Description: " & instance.Description
        End If
        WScript.Echo MSG_SEPERATE
    Next
    If i = 0 Then
        WScript.Echo MSG_NOLICENSEFOUND
    End If
    quitExit()
ElseIf strCommand = "/rtokil" Then
    Err.Clear    
    ExecuteQuery "ILID, ID","",tokenClass
    
    For Each instance in productinstances
        sppErrHandle ""
        i = i + 1
        If LCase(strValue) = LCase(instance.ILID) Then
            instance.Uninstall
            SppErrHandle(strCommand)
        Else
            WScript.Echo MSG_NOTFOUNDILID & strValue & " Run /dtokils to display the ILID for installed licenses."
        End If
    Next
    If i = 0 Then
        WScript.Echo MSG_NOLICENSEFOUND
    End If
    quitExit()
ElseIf strCommand = "/dtokcerts" Then
    Err.Clear
    Set objSigner = TkaGetSigner()
    sppErrHandle(strCommand)
    ExecuteQuery "ID, Name, ApplicationId, PartialProductKey, Description, LicenseIsAddon ","ApplicationId = '" & OfficeAppId & "' " & "AND PartialProductKey <> NULL " & "AND LicenseIsAddon = FALSE",productClass
    
    For each instance in productinstances
        i = i + 1
        sppErrHandle ""
        iRet = instance.GetTokenActivationGrants(arrGrants)
        If Err.Number = 0 Then
            arrThumbprints = objSigner.GetCertificateThumbprints(arrGrants)
            If Err.Number = 0 Then
                For Each strThumbprint in arrThumbprints
                    TkaPrintCertificate strThumbprint
                Next
            Else
                sppErrHandle ""
            End If
        Else
            sppErrHandle ""
        End If
        WScript.Echo MSG_SEPERATE
        Err.Clear
    Next
ElseIf strCommand = "/tokact" Then
    Err.Clear
    Set objSigner = TkaGetSigner()
    sppErrHandle "/ignore"
    pos1 = InStr(strValue,":")
    If pos1 = 0 Then
        'PIN not passed
        strThumbprint = strValue
    Else
        'PIN passed
        strThumbprint = Left(strValue,pos1 - 1)
        strPin = Replace(strValue,strThumbprint & ":","")
    End If
    
    ExecuteQuery "ID, Name, ApplicationId, PartialProductKey, Description, LicenseIsAddon ","ApplicationId = '" & OfficeAppId & "' " & "AND PartialProductKey <> NULL " & "AND LicenseIsAddon = FALSE",productClass
    
    For each instance in productinstances
        i = i + 1
        sppErrHandle ""        
        WScript.Echo MSG_TOKACTATTEMPT 
        WScript.Echo MSG_SKUID & instance.ID
        WScript.Echo MSG_LICENSENAME & instance.Name
        WScript.Echo MSG_DESCRIPTION & instance.Description
        WScript.Echo MSG_PARTIALKEY & instance.PartialProductKey
        iRet = instance.GenerateTokenActivationChallenge(strChallenge)
        If Err.Number = 0 Then
            strAuthInfo1 = objSigner.Sign(strChallenge, strThumbprint, strPin, strAuthInfo2)
            If Err.Number = 0 Then
                iRet = instance.DepositTokenActivationResponse(strChallenge, strAuthInfo1, strAuthInfo2)
                SppErrHandle(strCommand)
            Else
                sppErrHandle ""
            End If
        Else
            sppErrHandle ""
        End If
        WScript.Echo MSG_SEPERATE
    Next
Else
    Err.Clear
    If strCommand = "/dstatus" Or strCommand = "/dstatusall" Or strCommand = "/dpid" Or strCommand = "/dstatussub" Or strCommand = "/pstatus" Then
        If Win7 = True Then
            ExecuteQuery "ID, ApplicationId, EvaluationEndDate, PartialProductKey, Description, Name, LicenseStatus, LicenseStatusReason, ProductKeyID, GracePeriodRemaining, DiscoveredKeyManagementServiceMachineName, DiscoveredKeyManagementServiceMachinePort, VLActivationInterval, VLRenewalInterval, KeyManagementServiceMachine, KeyManagementServicePort, ProductKeyID2","ApplicationId = '" & OfficeAppId & "' ",productClass
        Else
            ExecuteQuery "ID, ApplicationId, EvaluationEndDate, PartialProductKey, Description, Name, LicenseStatus, LicenseStatusReason, ProductKeyID, GracePeriodRemaining, KeyManagementServiceLookupDomain, VLActivationType, ADActivationObjectName, ADActivationObjectDN, ADActivationCsvlkPid, ADActivationCsvlkSkuId, VLActivationTypeEnabled, DiscoveredKeyManagementServiceMachineName, DiscoveredKeyManagementServiceMachinePort, VLActivationInterval, VLRenewalInterval, KeyManagementServiceMachine, KeyManagementServicePort, ProductKeyID2","ApplicationId = '" & OfficeAppId & "' ",productClass    
        End If
    ElseIf strCommand = "/act" Or strCommand = "/actsub" Or strCommand = "/actp" Then
        ExecuteQuery "ID, ApplicationId, PartialProductKey, Description, Name","ApplicationId = '" & OfficeAppId & "' " & "AND PartialProductKey <> NULL ",productClass
    ElseIf strCommand = "/unpkey" Then
        ExecuteQuery "ID, ApplicationId, Description, PartialProductKey, Name, ProductKeyID","ApplicationId = '" & OfficeAppId & "' " & "AND PartialProductKey <> NULL ",productClass
        
    ElseIf strCommand = "/dinstid" Or strCommand = "/actcid" Then
        ExecuteQuery "ID, ApplicationId, Description, PartialProductKey, Name, OfflineInstallationId","ApplicationId = '" & OfficeAppId & "' " & "AND PartialProductKey <> NULL ",productClass
    ElseIf strCommand = "/actype" Or strCommand = "/skms-domain" Or strCommand = "/ckms-domain" Then
        If Win7 = True Then
             'Unsupported - sppsvc only supports this.
            globalPopFailure MSG_UNSUPPORTEDOPEROS8 & vbCr & strCommand,True
        Else
            ExecuteQuery "ID, Description, PartialProductKey, ApplicationId ","ApplicationId = '" & OfficeAppId & "' ",productClass
        End If
    ElseIf strCommand = "/sethst" Or strCommand = "/setprt" Or strCommand = "/remhst" Then
        ExecuteQuery "ID, Description, PartialProductKey, ApplicationId ","ApplicationId = '" & OfficeAppId & "' ",productClass
    End If
            
    For Each instance in productinstances
        sppErrHandle ""
        If (LCase(instance.ApplicationId) = OfficeAppId) Then
            If instance.PartialProductKey <> "" Then
                i = i + 1
            End If
            intIsKms = InStr(UCase(instance.Description),"KMS")
            If intIsKms <> 0 Then
                kmsCounter = kmsCounter + 1
            End If
            Select Case strCommand
                Case "/actype"
                    Select Case strValue
                        Case "0","1","2","3"
                        Case Else
                            globalPopFailure MSG_UNSUPPORTED & " A value of" & vbCr &  _
                            "0  (for all)" & vbCr & "1  (for AD)" & vbCr & "2  (for KMS" & vbCr & _
                            "3  (for Token)" & vbCr & "Is required for: " & strCommand,True
                    End Select
                    If intIsKms <> 0 Then
                        If strValue <> 0 Then                    
                            instance.SetVLActivationTypeEnabled(strValue)
                        Else
                            instance.ClearVLActivationTypeEnabled()
                        End If
                    End If
                    sppErrHandle ""
                Case "/skms-domain"
                    If intIsKms <> 0 Then
                        instance.SetKeyManagementServiceLookupDomain(strValue)
                    End If
                    sppErrHandle ""
                Case "/ckms-domain"
                    If intIsKms <> 0 Then
                        instance.ClearKeyManagementServiceLookupDomain()
                    End If
                    sppErrHandle ""
                Case "/sethst"
                    If intIsKms <> 0 Then
                        instance.SetKeyManagementServiceMachine(strValue)
                    End If
                    sppErrHandle ""
                Case "/setprt"
                    If intIsKms <> 0 Then
                        instance.SetKeyManagementServicePort(strValue)
                    End If
                    sppErrHandle ""
                Case "/remhst"
                    If intIsKms <> 0 Then
                        instance.ClearKeyManagementServiceMachine()
                        sppErrHandle ""
                        instance.ClearKeyManagementServicePort()
                        sppErrHandle ""
                    End If
                Case "/act"
                    WScript.Echo MSG_ACTATTEMPT 
                    WScript.Echo MSG_SKUID & instance.ID
                    WScript.Echo MSG_LICENSENAME & instance.Name
                    WScript.Echo MSG_DESCRIPTION & instance.Description
                    WScript.Echo MSG_PARTIALKEY & instance.PartialProductKey            
                    instance.Activate
                    SppErrHandle(strCommand)
                    WScript.Echo MSG_SEPERATE
                Case "/unpkey"
                    If Len(strValue) <> "5" Then
                        globalPopFailure MSG_ERRPARTIALKEY,True
                    End If
                    If UCase(strValue) = instance.PartialProductKey Then
                        y = y + 1
                        WScript.Echo MSG_UNINSTALLKEY & instance.Name
                        instance.UninstallProductKey(instance.ProductKeyID)                            
                        SppErrHandle(strCommand)
                    End If
                Case "/dinstid"
                    WScript.Echo "Installation ID for: " & instance.Name & ": " & instance.OfflineInstallationId
                    WScript.Echo MSG_SEPERATE
                Case "/actcid"
                    instance.DepositOfflineConfirmationId instance.OfflineInstallationId, strValue
                    If Err.Number = 0 Then
                        If telsuccess <> True Then
                            WScript.Echo MSG_LICENSENAME & instance.Name
                            WScript.Echo MSG_OFFLINEACTSUCCESS
                            telsuccess = True
                        End If
                    Else
                        WScript.Echo MSG_LICENSENAME & instance.Name
                        sppErrHandle ""
                    End If
                    WScript.Echo MSG_SEPERATE
                Case "/dpid"
                    If instance.ProductKeyID <> "" Then
                        intOccurSub = InStr(LCase(instance.Name),"_sub")
                        intOccurPerp = InStr(LCase(instance.Name),"_retail") Or InStr(LCase(instance.Name),"_perp")
                        If intOccurSub <> 0 Or intOccurPerp <> 0 Then
                            WScript.Echo MSG_SEPERATE
                            If intOccurSub <> 0 Then 
                                foundSubKey = True
                                WScript.Echo MSG_KEYSINSTALLED_SUB
                            End If
                            If intOccurPerp <> 0 Then 
                                foundPerpKey = True
                                WScript.Echo MSG_KEYSINSTALLED_PERP
                            End If
                            WScript.Echo MSG_SEPERATE
                            WScript.Echo MSG_SKUID & UCase(instance.ID)
                            WScript.Echo MSG_LICENSENAME & instance.Name
                            WScript.Echo MSG_PARTIALKEY & instance.PartialProductKey 
                            WScript.Echo MSG_PID & instance.ProductKeyID2
                            'Determine machinekey from PID
                            strPid2 = Replace(instance.ProductKeyID2,"-","")
                            strPid2 = Right(strPid2,19)
                            strGid = Mid(strPid2,1,5)
                            strSerial1 = Mid(strPid2,6,3)
                            strSerial2 = Mid(strPid2,9,6)
                            WScript.Echo "Machine Key: " & strGid & "-" & strSerial1 & "-" & strSerial2
                        End If
                    End If
                Case "/dstatus", "/dstatusall"
                    getInstalled = False
                    verifyFileExists currentDir & "slerror.xml"
                    licSr = Hex(instance.LicenseStatusReason)
                    If strCommand = "/dstatusall" Then
                        getInstalled = True
                        WScript.Echo MSG_SKUID & instance.ID
                        WScript.Echo MSG_LICENSENAME & instance.Name
                        WScript.Echo MSG_DESCRIPTION & instance.Description            
                    Else
                        If instance.ProductKeyID <> "" Then
                            getInstalled = True
                            WScript.Echo MSG_PID & instance.ProductKeyID2                                                        
                            WScript.Echo MSG_SKUID & instance.ID
                            WScript.Echo MSG_LICENSENAME & instance.Name
                            WScript.Echo MSG_DESCRIPTION & instance.Description
                            'EvaluationEndDate always returns a value. When an expiry date is not defined "1/1/1601" is returned.
                            'So avoid displaying this against RTM licenses as that date = no expiry defined.
                            If instance.EvaluationEndDate <> "" Then
                                Set objDate = CreateObject("WBemScripting.SWbemDateTime")
                                objDate.Value = instance.EvaluationEndDate
                                If objDate.GetVarDate() <> "1/1/1601" Then
                                    WScript.Echo MSG_LICEXPIRY & objDate.GetVarDate()
                                End If
                                Set objDate = Nothing
                            End If
                         End If
                    End If
                    
                    If getInstalled = True Then
                        Select Case instance.LicenseStatus
                            Case 0
                                WScript.Echo MSG_LICSTATUS & MSG_UNLICENSED
                            Case 1
                                WScript.Echo MSG_LICSTATUS & MSG_LICENSED
                            Case 2
                                WScript.Echo MSG_LICSTATUS & MSG_OOBGRACE        
                            Case 3
                                WScript.Echo MSG_LICSTATUS & MSG_OOTGRACE
                            Case 4
                                WScript.Echo MSG_LICSTATUS & MSG_NONGENGRACE
                            Case 5
                                WScript.Echo MSG_LICSTATUS & MSG_NOTIFICATION
                            Case 6
                                WScript.Echo MSG_LICSTATUS & MSG_EXTENDEDGRACE    
                            Case Else
                                WScript.Echo MSG_LICSTATUS & MSG_LICUNKNOWN
                        End Select
                            
                        If licSr <> "0" Then
                            If instance.LicenseStatus <> 1 Then
                                WScript.Echo MSG_ERRCODE & "0x" & licSr
                            Else
                                WScript.Echo MSG_ERRCODE & "0x" & licSr & MSG_INFO_ONLY
                            End If
                            getResource("err" & "0x" & licSr)
                            If globalResource = "" Then
                                If foundSlUi <> True Then
                                    WScript.Echo MSG_ERRDESC & "Not available."
                                Else
                                    WScript.Echo MSG_ERRDESC & "Run the following: cscript ospp.vbs /ddescr:0x" & licSr
                                End if
                            Else
                                WScript.Echo MSG_ERRDESC & globalResource
                            End If
                        End If
                        
                        If instance.GracePeriodRemaining <> 0 Then
                            dGrace = instance.GracePeriodRemaining / 60 / 24
                            rndDown = Int(dGrace)
                            WScript.Echo MSG_REMAINGRACE & rndDown & " days " & " (" & instance.GracePeriodRemaining & " minute(s) before expiring" & ")"
                        End If
                            
                        If instance.PartialProductKey <> "" Then
                            WScript.Echo MSG_PARTIALKEY & instance.PartialProductKey
                            'Display additional volume info for KMS licenses
                            If intIsKms <> 0 Then
                                'Display activation type set (Win8+).
                                If Win7 <> True Then
                                    Select Case instance.VLActivationTypeEnabled
                                        Case 1
                                            WScript.Echo MSG_VLActivationType & "AD"
                                        Case 2
                                            WScript.Echo MSG_VLActivationType & "KMS"
                                        Case 3
                                            WScript.Echo MSG_VLActivationType & "Token"
                                        Case Else
                                            WScript.Echo MSG_VLActivationType & "ALL"
                                    End Select
                                    
                                    'Check to see if last activated via AD- display object info (Win8+).
                                    If instance.VLActivationType = 1 Then
                                        isAdActivated = True
                                        WScript.Echo MSG_Act_Recent + "AD"
                                        WScript.Echo vbTab & MSG_ADInfoAOName & instance.ADActivationObjectName
                                        WScript.Echo vbTab & MSG_ADInfoAODN & instance.ADActivationObjectDN
                                        WScript.Echo vbTab & MSG_ADInfoExtendedPid & instance.ADActivationCsvlkPid
                                        WScript.Echo vbTab & MSG_ADInfoActID & instance.ADActivationCsvlkSkuId
                                    End If
                                End If
                                
                                If isAdActivated = False Then
                                    strKms = instance.DiscoveredKeyManagementServiceMachineName
                                    strPort = instance.DiscoveredKeyManagementServiceMachinePort
                                        
                                    If IsNull(strKms) Or (strKms = "") Or IsNull(strPort) Or (strPort = 0) Then
                                        WScript.Echo vbTab & MSG_KMS_DNS_ERR
                                    Else
                                        WScript.Echo vbTab & MSG_KMS_DNS & strKMS & ":" & strPort
                                    End If
                                    
                                    'Check to see if registry override is defined
                                    strKms = instance.KeyManagementServiceMachine
                                    If strKms <> "" And Not IsNull(strKms) Then
                                         strPort = instance.KeyManagementServicePort
                                        If (strPort = 0) Then
                                            strPort = MSG_DEFAULT_PORT
                                        End If
                                        WScript.Echo vbTab & MSG_HOST_REG_OVERRIDE & strKms & ":" & strPort
                                    End If
                                        
                                    WScript.Echo vbTab & MSG_ACTIVATION_INTERVAL & instance.VLActivationInterval & " minutes"
                                    WScript.Echo vbTab & MSG_RENEWAL_INTERVAL & instance.VLRenewalInterval & " minutes"
                                        
                                     If (objOspp.KeyManagementServiceHostCaching = True) Then
                                        WScript.Echo vbTab & MSG_HOST_CACHING & "Enabled"
                                    Else
                                        WScript.Echo vbTab & MSG_HOST_CACHING & "Disabled"
                                    End If
                                    
                                    If Win7 <> True Then     
                                        If instance.KeyManagementServiceLookupDomain <> "" Then
                                            WScript.Echo vbTab & MSG_KMSLOOKUP & instance.KeyManagementServiceLookupDomain
                                        End If
                                    End If
                                End If                               
                            End If
                        End If
                        WScript.Echo MSG_SEPERATE
                    End If
                Case "/actsub","/actp"
                	intOccurSub=0
                    intOccurPerp=0
                    intOccurProduct=0
                    
                    If strCommand = "/actsub" Then
                        intOccurSub = InStr(LCase(instance.Name),"_sub")
                    Else
                        intOccurSub = InStr(LCase(instance.Name),"_sub")
                        intOccurPerp = InStr(LCase(instance.Name),"_retail") Or InStr(LCase(instance.Name),"_perp")
                        If intOccurSub <> 0 Or intOccurPerp <> 0 Then
                            intOccurProduct = InStr(LCase(instance.Name),strValue)
                        End If
                    End If

                    If strCommand = "/actsub" And intOccurSub <> 0 Or intOccurProduct <> 0 Then
                        foundCssKey = True
                        WScript.Echo MSG_ACTATTEMPT 
                        WScript.Echo MSG_SKUID & instance.ID
                        WScript.Echo MSG_LICENSENAME & instance.Name
                        WScript.Echo MSG_DESCRIPTION & instance.Description
                        WScript.Echo MSG_PARTIALKEY & instance.PartialProductKey
                        instance.Activate
                        sppErrHandle(strCommand)
                        WScript.Echo MSG_SEPERATE
                    End If
                Case "/dstatussub","/pstatus"
                	intOccurSub=0
                    intOccurPerp=0
                    intOccurProduct=0
                    dlicense = True

                    verifyFileExists currentDir & "slerror.xml"
                    licSr = Hex(instance.LicenseStatusReason)
                    If instance.ProductKeyID <> "" Then
                        If strCommand = "/dstatussub" Then
                            intOccurSub = InStr(LCase(instance.Name),"_sub")
                        Else
                            intOccurSub = InStr(LCase(instance.Name),"_sub")
                            intOccurPerp = InStr(LCase(instance.Name),"_retail") Or InStr(LCase(instance.Name),"_perp")
                            If intOccurSub <> 0 Or intOccurPerp <> 0 Then
                                intOccurProduct = InStr(LCase(instance.Name),strValue)
                                If intOccurSub <> 0 And intOccurProduct <> 0 Then
                                	getHeartBeat = True
                                End If
                            End If
                        End If

                        If strCommand = "/dstatussub" And intOccurSub <> 0 Or intOccurProduct <> 0 Then
                            foundCssKey = True
                            WScript.Echo MSG_PID & instance.ProductKeyID2
                            WScript.Echo MSG_SKUID & instance.ID
                            WScript.Echo MSG_LICENSENAME & instance.Name
                            WScript.Echo MSG_DESCRIPTION & instance.Description
                            'EvaluationEndDate always returns a value. When an expiry date is not defined "1/1/1601" is returned.
                            'So avoid displaying this against RTM licenses as that date = no expiry defined.
                            If instance.EvaluationEndDate <> "" Then
                                Set objDate = CreateObject("WBemScripting.SWbemDateTime")
                                objDate.Value = instance.EvaluationEndDate
                                If objDate.GetVarDate() <> "1/1/1601" Then
                                    WScript.Echo MSG_LICEXPIRY & objDate.GetVarDate()
                                End If
                                Set objDate = Nothing
                            End If

                            Select Case instance.LicenseStatus
                                Case 0
                                    WScript.Echo MSG_LICSTATUS & MSG_UNLICENSED
                                Case 1
                                    WScript.Echo MSG_LICSTATUS & MSG_LICENSED
                                Case 2
                                    WScript.Echo MSG_LICSTATUS & MSG_OOBGRACE
                                Case 3
                                    WScript.Echo MSG_LICSTATUS & MSG_OOTGRACE
                                Case 4
                                    WScript.Echo MSG_LICSTATUS & MSG_NONGENGRACE
                                Case 5
                                    WScript.Echo MSG_LICSTATUS & MSG_NOTIFICATION
                                Case 6
                                    WScript.Echo MSG_LICSTATUS & MSG_EXTENDEDGRACE
                                Case Else
                                    WScript.Echo MSG_LICSTATUS & MSG_LICUNKNOWN
                            End Select

                            If licSr <> "0" Then
                                If instance.LicenseStatus <> 1 Then
                                    WScript.Echo MSG_ERRCODE & "0x" & licSr
                                Else
                                    WScript.Echo MSG_ERRCODE & "0x" & licSr & MSG_INFO_ONLY
                                End If

                                getResource("err" & "0x" & licSr)
                                If globalResource = "" Then
                                    If foundSlUi <> True Then
                                       WScript.Echo MSG_ERRDESC & "Not available."
                                    Else
                                        WScript.Echo MSG_ERRDESC & "Run the following: cscript ospp.vbs /ddescr:0x" & licSr
                                    End if
                                Else
                                    WScript.Echo MSG_ERRDESC & globalResource
                                End If
                            End If

                            If instance.GracePeriodRemaining <> 0 Then
                                dGrace = instance.GracePeriodRemaining / 60 / 24
                                rndDown = Int(dGrace)
                                WScript.Echo MSG_REMAINGRACE & rndDown & " days " & " (" & instance.GracePeriodRemaining & " minute(s) before expiring" & ")"
                            End If

                            If instance.PartialProductKey <> "" Then
                                WScript.Echo MSG_PARTIALKEY & instance.PartialProductKey
                            End If

                            'Determine machinekey from PID
                            strPid2 = Replace(instance.ProductKeyID2,"-","")
                            strPid2 = Right(strPid2,19)
                            strGid = Mid(strPid2,1,5)
                            strSerial1 = Mid(strPid2,6,3)
                            strSerial2 = Mid(strPid2,9,6)
                            WScript.Echo "MACHINE KEY: " & strGid & "-" & strSerial1 & "-" & strSerial2
                            WScript.Echo MSG_SEPERATE
                         End If
                    End If
                Case Else
            End Select
        End If
    Next
End If

Select Case strCommand
    Case "/unpkey"
        If y = 0 Then
            WScript.Echo MSG_KEYNOTFOUND
            quitExit()
        End If
    Case "/ckms-domain","/skms-domain","/actype","/sethst","/setprt","/remhst"
        If kmsCounter = 0 Then
            WScript.Echo MSG_NOKMSLICS
            quitExit()
        Else
            sppErrHandle(strCommand)
        End If
    Case Else
End Select

If strCommand <> "/dpid" Then
    If strCommand = "/actsub" Or strCommand = "/dstatussub" Or strCommand = "/actp" Or strCommand = "/pstatus" Then
        If foundCssKey <> True Then
            WScript.Echo MSG_NOKEYSINSTALLED
            WScript.Echo MSG_SEPERATE
        End If
    Else
        If i = 0 Then
            WScript.Echo MSG_NOKEYSINSTALLED
            WScript.Echo MSG_SEPERATE
        End If
    End If
Else
    If foundSubKey <> True Then
        WScript.Echo MSG_SEPERATE
        WScript.Echo MSG_NOKEYSINSTALLED_SUB
        WScript.Echo MSG_SEPERATE
    End If
    If foundPerpKey <> True Then
        WScript.Echo MSG_SEPERATE
        WScript.Echo MSG_NOKEYSINSTALLED_PERP
        WScript.Echo MSG_SEPERATE
    End If
End If

If dlicense = False Then quitExit()

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function performRegAction(strCommand)

On Error Resume Next

If Win7 = True Then
    Set colListOfServices = objWMI.ExecQuery _
        ("Select * from Win32_Service ")
    For Each objService in colListOfServices
        If objService.Name = "osppsvc" Then
            installed = True
            Exit For
        End If
    Next
        
    If installed <> True Then
        globalPopFailure MSG_OSPPSVC_NOINSTALL,True
    End If
    checkRegRights objWMI1,REG_OSPP
Else
    checkRegRights objWMI1,REG_SPP
End If

Select Case strCommand
    Case "/puserops"
        setRegValue objWMI1,"1","UserOperations"
    Case "/duserops"
        setRegValue objWMI1,"0","UserOperations"
End Select

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function performServiceAction(strCommand)

On Error Resume Next

Set colListOfServices = objWMI.ExecQuery _
    ("Select * from Win32_Service ")
For Each objService in colListOfServices
    If objService.Name = "osppsvc" Then
        installed = True
        Exit For
    End If
Next
    
If installed <> True Then
    globalPopFailure MSG_OSPPSVC_NOINSTALL,True
End If

Set objService = Nothing
Set colListOfServices = Nothing

If strCommand = "/osppsvcauto" Then
    Set colListOfServices = objWMI.ExecQuery _
        ("Select * from Win32_Service where StartMode = 'Manual' or StartMode = 'Disabled'")
        For Each objService in colListOfServices
            If LCase(objService.Name) = "osppsvc" Then
                foundOsppNonAuto = True
                objService.Change , , , , "Automatic"
                WScript.Sleep(15000)
                Exit For
            End If
        Next
        If foundOsppNonAuto <> True Then
            WScript.Echo "Service startup type already set to automatic: Office Software Protection Platform"
            quitExit()
        End If
        
        Set objService = Nothing
        Set colListOfServices = Nothing
        Set colListOfServices = objWMI.ExecQuery _
        ("Select * from Win32_Service where StartMode = 'Auto'")
        For Each objService in colListOfServices
            If LCase(objService.Name) = "osppsvc" Then
                foundOsppAuto = True
                WScript.Echo "Successfully set service startup to automatic:" & objService.DisplayName
                quitExit()
            End If
        Next
        
        If foundOsppAuto <> True Then
            WScript.Echo "Unsuccessful setting service startup to automatic. " & MSG_ISCMD_ELEVATED
            quitExit()
        End If
Else
    Set colListOfServices = objWMI.ExecQuery _
        ("Select * from Win32_Service ")
    For Each objService in colListOfServices
        If LCase(objService.Name) = "osppsvc" Then
            Select Case LCase(objService.State)
                Case "running"
                    objService.StopService()
                    WScript.Sleep(15000)
                    objService.StartService()
                    WScript.Sleep(15000)
                Case Else
                    objService.StartService()
                    WScript.Sleep(15000)
            End Select
            Exit For
        End If
    Next
    
    Set objService = Nothing
    Set colListOfServices = Nothing
    Set colListOfServices = objWMI.ExecQuery _
        ("Select * from Win32_Service ")
    For Each objService in colListOfServices
        If LCase(objService.Name) = "osppsvc" Then
            If LCase(objService.State) = "running" Then
                WScript.Echo "Successfully restarted: " & objService.DisplayName
                quitExit()
            Else
                WScript.Echo "Unsuccessful restart: " & objService.DisplayName & ". Status: " _
                    & objService.State & ". " & MSG_ISCMD_ELEVATED
                quitExit()
            End If
            Exit For
        End If
    Next
End If

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function reARM(skuid)

progFiles = WshShell.ExpandEnvironmentStrings("%ProgramFiles%")

If objFSO.FileExists(progFiles & STR_OSPPREARMPATH) Then
    rearmPath = progFiles & STR_OSPPREARMPATH
ElseIf objFSO.FileExists(progFiles & STR_OSPPREARMPATH_DEBUG) Then
    rearmPath = progFiles & STR_OSPPREARMPATH_DEBUG
Else
    progFilesX86 = WshShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
    If objFSO.FileExists(progFilesX86 & STR_OSPPREARMPATH) Then
        rearmPath = progFilesX86 & STR_OSPPREARMPATH
    ElseIf objFSO.FileExists(progFilesX86 & STR_OSPPREARMPATH_DEBUG) Then
        rearmPath = progFilesX86 & STR_OSPPREARMPATH_DEBUG
    Else
        WScript.Echo MSG_FILENOTFOUND & "OSPPREARM.EXE"
        quitExit()
    End If
End If

If skuid = "" Then   
    Set objScriptExec = WshShell.Exec (rearmPath)
Else
    Set objScriptExec = WshShell.Exec (rearmPath & " " & skuid)
End If

readOut = objScriptExec.StdOut.ReadAll
WScript.Echo readOut
    
End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function setWinOS()

Set colOperatingSystems = objWMI.ExecQuery _
        ("Select * from Win32_OperatingSystem")
    For Each objOperatingSystem in colOperatingSystems
        Ver = Split(objOperatingSystem.Version, ".", -1, 1)
        ' Win7/Server2008R2
         If (Ver(0) = "6" And Ver(1) = "1") Then
            Win7 = True
            Exit For
         End If
            
         ' Win8/Server2012
        If (Ver(0) = "6" And Ver(1) = "2") Then
            Win8 = True
            Exit For
        End If
    Next

setWmiClasses()

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function setWmiClasses()

If Win7 = True Then
    productClass = "OfficeSoftwareProtectionProduct"
    tokenClass = "OfficeSoftwareProtectionTokenActivationLicense"
Else
    productClass = "SoftwareLicensingProduct"
    tokenClass = "SoftwareLicensingTokenActivationLicense"
End If

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function getEvents(strMessage,strLog,strEventCode,strSource,strMachine)
       
WScript.Echo MSG_SEARCHFOR & strMessage & "<<<<< Log: " & strLog & "  Source: " & strSource & "  Event: " & strEventCode & " >>>>>" 
WScript.Echo MSG_SEPERATE

Set objEvents = objWMI.ExecQuery _
    ("Select * from Win32_NTLogEvent Where Logfile = '" & strLog &  "'" & " and "_
    & "EventCode = '" & strEventCode & "'" & " and SourceName = '" & strSource & "'")
        
If objEvents.Count > 0 Then
    WScript.Echo MSG_SEPERATESMALL
    For each objEvent in objEvents
        dtmEventDate = objEvent.TimeWritten
        strTimeWritten = WMIDateStringToDate(dtmEventDate)
        WScript.Echo "Coordinated Universal Time Written: " & strTimeWritten
        WScript.Echo objEvent.Message
        WScript.Echo MSG_SEPERATESMALL
    Next
    WScript.Echo MSG_SEPERATE
Else
    WScript.Echo MSG_NOEVENTS
    WScript.Echo MSG_SEPERATE
End If 

Set objEvents = Nothing
        
End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function getMachineId(strMachine)

Set objInstances = objWMI.InstancesOf("Win32_ComputerSystemProduct")

For Each objInstance in objInstances
    WScript.Echo "MachineID: " & objInstance.uuid
    WScript.Echo MSG_SEPERATE
Next

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function getHeartbeatXml(strMachine)

On Error Resume Next

If strMachine <> "" Then
    WScript.Echo MSG_HEARTBEAT_LOCALONLY
Else
    If objFSO.FileExists(CreateObject("Shell.Application").Namespace(&H23&).Self.Path & STR_HEARTBEATPATH) Then
        fHeartBeat = True
        Set MyXmlDoc = CreateObject("Msxml2.DOMDocument.6.0") 
        MyXmlDoc.load(CreateObject("Shell.Application").Namespace(&H23&).Self.Path & STR_HEARTBEATPATH)
        WScript.Echo "HeartbeatCache.xml found: " & CreateObject("Shell.Application").Namespace(&H23&).Self.Path & STR_HEARTBEATPATH
        WScript.Echo "Contents:"
        WScript.Echo MSG_SEPERATESMALL
        Set nodes = MyXmlDoc.selectNodes("//*")
        For i = 0 to nodes.length - 1
            If LCase(nodes(i).nodeName) <> "heartbeatcache" Then
                'lastchk & skuid are attribs of License node
                If LCase(nodes(i).nodeName) = "license" Then
                    hLastChk = nodes(i).attributes.getNamedItem("lastCheckTime").text
                    WScript.Echo "lastCheckTime: " & hLastChk
                    hSkuId = nodes(i).attributes.getNamedItem("skuId").text
                    WScript.Echo "skuId: " & hSkuId
                Else
                    WScript.Echo nodes(i).nodeName & ": " & nodes(i).text
                End If
            End If
        Next
    End If
    
    If fHeartBeat <> True Then
        WScript.Echo MSG_FILENOTFOUND & "HeartbeatCache.xml"
    End If
    
End If
    
WScript.Echo MSG_SEPERATE

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
' Checks if there is a KB article for the specified error
Function lookupKBArticle(errorCode)
    If InStr(errorKBs, errorCode) > 0 Then
        WScript.Echo MSG_ACT_ERROR_FOUND_KB & "0x" & errorCode
        WScript.Echo MSG_ACT_ERROR_KB_LINK & errorCode
    End If
End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'Checks the registry for any key that enables SCA and reports whether SCA
'is enabled or not by any of those keys.
Function checkScaStatus()
    WScript.Echo MSG_SEARCHFOR & MSG_SCAREGKEYS
    WScript.Echo MSG_SEPERATE

    Dim flavorsList
    Set flavorsList = CreateObject("System.Collections.ArrayList")

    flavorsList.Add(VALUE_SCAEXP_FLAVOR)
    flavorsList.Add(VALUE_SCAEXP_OVERRIDE_FLAVOR)
    flavorsList.Add(VALUE_SCAO15_FLAVOR)
    flavorsList.Add(VALUE_SCA016_FLAVOR)

    Dim result, flavor
    result = False
    For Each flavor in flavorsList
        result = result Or checkScaStatusForFlavor(flavor)
    Next

    If result = True Then
        WScript.Echo "Machine in SCA mode."
    Else
        WScript.Echo "Machine NOT in SCA mode."
    End If

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'Checks the registry for a particular key that enables SCA in one specific
'flavor and returns whether SCA is enabled or disabled. Flavors are SCA Experiment,
'SCA Experiment override, SCA for O15 and SCA for O16
Function checkScaStatusForFlavor(flavor)

    On Error Resume Next

    dim key, message, keyType, result

    Select Case flavor
        Case VALUE_SCAEXP_FLAVOR
            key =  REG_SCAEXP_BASE
            message = MSG_SCAEXP
            keyType = VALUE_SCAREGKEY_PER_APP_TYPE
        Case VALUE_SCAEXP_OVERRIDE_FLAVOR
            key = REG_SCAEXP_OVERRIDE
            message = MSG_SCAEXP_OVERRIDE
            keyType = VALUE_SCAREGKEY_SINGLE_TYPE
        Case VALUE_SCAO15_FLAVOR
            key = REG_SCAO15
            message = MSG_SCAO15
            keyType = VALUE_SCAREGKEY_SINGLE_TYPE
        Case VALUE_SCA016_FLAVOR
            key = REG_SCAO16
            message = MSG_SCAO16
            keyType = VALUE_SCAREGKEY_SINGLE_TYPE
        Case Else
    End Select

    result = False
    If (keyType = VALUE_SCAREGKEY_SINGLE_TYPE) Then
        result = readAndReportScaKey(key, message)
    ElseIf (keyType = VALUE_SCAREGKEY_PER_APP_TYPE) Then
        Dim appsList
        Set appsList = CreateObject("System.Collections.ArrayList")

        appsList.Add("word")
        appsList.Add("excel")
        appsList.Add("powerpoint")
        appsList.Add("visio")
        appsList.Add("project")
        appsList.Add("onenote")
        appsList.Add("outlook")
        appsList.Add("access")
        appsList.Add("publisher")
        appsList.Add("lync")

        Dim app, keyPerApp
        For Each app In appsList
            keyPerApp = key & app & REG_SCAEXP_END
            result = result Or readAndReportScaKey(keyPerApp, message)
        Next
    End If

    checkScaStatusForFlavor = result
End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'Reads the registry looking for a SCA key and if found reports whether or not
'that key enables SCA.
Function readAndReportScaKey(key, message)

    On Error Resume Next

    Dim result, keyValue
    keyValue = WshShell.RegRead(key)

    If Err.Number = 0 Then
        If keyValue = 1 Then
            WScript.Echo message & MSG_SCAREGKEY_ACTIVE
            result = True
        Else
            WScript.Echo message & MSG_SCAREGKEY_INACTIVE
            result = False
        End If

        WScript.Echo key
        WScript.Echo MSG_SEPERATESMALL
    Else
        result = False
    End If
    Err.Clear

    readAndReportScaKey = result
End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////

'' SIG '' Begin signature block
'' SIG '' MIIpawYJKoZIhvcNAQcCoIIpXDCCKVgCAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' eRmuJTRxBA78Fz3WCj2JjTRm6cUMUZvuKsvOrf3TrTmg
'' SIG '' gg3lMIIGvTCCBKWgAwIBAgITMwAAABxIn4HfobC3dwAA
'' SIG '' AAAAHDANBgkqhkiG9w0BAQwFADCBiDELMAkGA1UEBhMC
'' SIG '' VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcT
'' SIG '' B1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
'' SIG '' b3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3Qg
'' SIG '' Q2VydGlmaWNhdGUgQXV0aG9yaXR5IDIwMTAwHhcNMjQw
'' SIG '' ODA4MjEzNjIzWhcNMzUwNjIzMjIwNDAxWjBfMQswCQYD
'' SIG '' VQQGEwJVUzEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
'' SIG '' cmF0aW9uMTAwLgYDVQQDEydNaWNyb3NvZnQgV2luZG93
'' SIG '' cyBDb2RlIFNpZ25pbmcgUENBIDIwMjQwggIiMA0GCSqG
'' SIG '' SIb3DQEBAQUAA4ICDwAwggIKAoICAQCafWt9J8F2Ki6u
'' SIG '' 49U0/8wrbe78VPggo/uwZIn0vwdoFyhlOzlfUl0SRj9c
'' SIG '' hbOaeo6bGIuHGMxeegFdABJphI1fME9pbz1OQYTd8Fd9
'' SIG '' B6mDyGBI+T91l39JFw/X741H9RgLVxK4ifMOwCzWlRJv
'' SIG '' UbOHjwNGbGB2gm1OZAVCUA17++oWnznEIHRQgNyN82LX
'' SIG '' 819rzsMfO7gzmgrsijkWYofXN803/kywuUGC8oVTAZw1
'' SIG '' xBwzq72sPdg0siKqXYEVqbn86gxctXoFY5KF2YW/vaWf
'' SIG '' YXlMzV014TqF83sYemMwC+H5QVpvgXNYUMhEnpxLwSc5
'' SIG '' 1ftubt4e+444DFGOOPll0OLvanXQ3v1OUngGikb74m5o
'' SIG '' uM+0EaS72bJWtAj4jlBs9NA6ObH5AtBMJbEs3zN/vAPa
'' SIG '' 7MhVToFg1T87ffDiT9hKGhDqvBhPRgqDdou/+AthQsH3
'' SIG '' 9QUgkyVmTtVnK9jLXiROlMRlfooQPJzedWDyg9nWBqHs
'' SIG '' K170cwv9R6FHkr5WX9Jn/RhxLb75GyVUUaOjwX9Jnebf
'' SIG '' O1W9ZjP3yKdXsqcmsZl5IKXAcLspbDqtpElTiecAT6Gh
'' SIG '' LLCZHjHCpxLrrvvlCnQx5UtA7bGIzdEJzrnL03UrHb4c
'' SIG '' yjkoyRd11aq/X9gveOS10+a8SiB1CBAwXDWFOgSgwx+q
'' SIG '' 36SjjgkopQIDAQABo4IBRjCCAUIwDgYDVR0PAQH/BAQD
'' SIG '' AgGGMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBQe
'' SIG '' gt8O14yz1wI0gw7aq61lua+47DAZBgkrBgEEAYI3FAIE
'' SIG '' DB4KAFMAdQBiAEMAQTAPBgNVHRMBAf8EBTADAQH/MB8G
'' SIG '' A1UdIwQYMBaAFNX2VsuP6KJcYmjRPZSQW9fOmhjEMFYG
'' SIG '' A1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9jcmwubWljcm9z
'' SIG '' b2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0Nl
'' SIG '' ckF1dF8yMDEwLTA2LTIzLmNybDBaBggrBgEFBQcBAQRO
'' SIG '' MEwwSgYIKwYBBQUHMAKGPmh0dHA6Ly93d3cubWljcm9z
'' SIG '' b2Z0LmNvbS9wa2kvY2VydHMvTWljUm9vQ2VyQXV0XzIw
'' SIG '' MTAtMDYtMjMuY3J0MA0GCSqGSIb3DQEBDAUAA4ICAQBD
'' SIG '' X/jfP7vplIw7XPW7aAOdkQXNF1Q0gTEATKsbueoVxwcL
'' SIG '' nLVFrNVwagwzCBQh7vXOmP1BfkzfBCII57owKSmJhz+H
'' SIG '' +BDNwEUppc66ReaMzicdAQORVL9Y5qXX/9mW6qbwsZcb
'' SIG '' /xtUeCo60ppqjx87OooMN2+0U24+wcSEvHziJMGFkIQd
'' SIG '' ny45YPtx0qwxjxSIaSCVlWpjCEe2u9jhqJ43X+Oa7KcK
'' SIG '' iB7sp2VOGr8va7gf0YYW8JvnzG/ATHnCGk5pKIcfxGWe
'' SIG '' RjVnDeqE2FtxtgTNwd2M51pJfbeLIT+tHzLnvtpLHRxl
'' SIG '' khPBFU3UphlHY9I61HOOpRlRSSEhd/zMXMZ5TXj9Socq
'' SIG '' /mc0+BLbPyO5rn6Wi5y2pczEdsyLoRjgFlrMHrG47Rc5
'' SIG '' FVBYA0dklvdNyNFypWzxAOqvHqRxifa6MYfOZ7BCnATV
'' SIG '' MOEnKevCgqkqRQWiosldbJHfpfFOdFjXjzG/Qc89DnwE
'' SIG '' mpfL+bEBvg1tNZDfiPkSlCGzOSOdMCY4h8pkBTQ7G6Gx
'' SIG '' cfSPeZghBD1O31Gd1U/xzlFW5Jl+5bSAv3kALuRjvH7v
'' SIG '' nHhEzMm726MVDOHWDQvj86KFMX5gtA7ikcAdtW1/fmnL
'' SIG '' iAZMSJuBHdztfcNVS6AO1DTlLie8+jUNlv/qu3J3zj5d
'' SIG '' kFS+KpYAm5VE9r5kKZZVdzCCByAwggUIoAMCAQICEzMA
'' SIG '' AACG53GUuU3/CfsAAAAAAIYwDQYJKoZIhvcNAQEMBQAw
'' SIG '' XzELMAkGA1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jvc29m
'' SIG '' dCBDb3Jwb3JhdGlvbjEwMC4GA1UEAxMnTWljcm9zb2Z0
'' SIG '' IFdpbmRvd3MgQ29kZSBTaWduaW5nIFBDQSAyMDI0MB4X
'' SIG '' DTI1MDUwODE4MjQ1M1oXDTI2MDUwNjE4MjQ1M1owdDEL
'' SIG '' MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
'' SIG '' EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
'' SIG '' c29mdCBDb3Jwb3JhdGlvbjEeMBwGA1UEAxMVTWljcm9z
'' SIG '' b2Z0IENvcnBvcmF0aW9uMIICIjANBgkqhkiG9w0BAQEF
'' SIG '' AAOCAg8AMIICCgKCAgEApH31lbdTfYCtnKM6hm0dcBgl
'' SIG '' L0fUQXvAsjycUQQnTbg9O0ONzUD4l9K0pCaMmLq88Z3P
'' SIG '' NgSZpNFCnR+AfdvW8SZCtaYwDflp/SXnWyEzAGiD9Vps
'' SIG '' OtEgtB7OsFynkTPZ6jk25AYadsClcZA0KaZ36mUTjJ+h
'' SIG '' LS3IBH6U1EJPyB8WiG8amw50S2N9S1Gq3qEZVH9LbizX
'' SIG '' +NmQIFA8hLVTo64uaXuAFcj7Kz7iIRBy18yq6Z7qqyup
'' SIG '' 0n04LK+2+38C+Mandd64RtQ0t9CyuiXwgFcOxz2w/+5v
'' SIG '' 0UDoafmTD4ZWWXy1r56rve3Jqd49/jSkRc2AhelN1Tbr
'' SIG '' XTjfwIS5+DTdPks+jJunosQXEgEbmMFdSutr1M8k4itu
'' SIG '' di0IO10qsFA3HQk167vtm7fWh/JKLEFoA6/Itl3PY8mV
'' SIG '' I4xMDJVCs8bp6s+7UPNDs00gIM9spFtKNegk0xFePz2/
'' SIG '' 4S1hZT4V6YfoycfjJIU9/Mq939rj8FQrj/XjGUrhGnym
'' SIG '' XPkzl9RrK9gRSMMXT/CAiUx+zKb8AL+Bte7KPOMTRhr0
'' SIG '' au06UvS8urm1CeNPe9PSZ6LON3Q+315xOV5kaCC7xQ1A
'' SIG '' qN7Z7fWbIZ8e8HwLIPvKYVxznemImIExu+97NRxj7zJc
'' SIG '' GQk/hzxV+1uc4iHTUAttop9+/6q0O8HSoOFDJmSGOMMC
'' SIG '' AwEAAaOCAb4wggG6MA4GA1UdDwEB/wQEAwIHgDAfBgNV
'' SIG '' HSUEGDAWBgorBgEEAYI3PQYBBggrBgEFBQcDAzAMBgNV
'' SIG '' HRMBAf8EAjAAMB0GA1UdDgQWBBRQoxo91TqmkNu2C06l
'' SIG '' TspQcpQogDBUBgNVHREETTBLpEkwRzEtMCsGA1UECxMk
'' SIG '' TWljcm9zb2Z0IElyZWxhbmQgT3BlcmF0aW9ucyBMaW1p
'' SIG '' dGVkMRYwFAYDVQQFEw0yMzA4NjUrNTA0NTkxMB8GA1Ud
'' SIG '' IwQYMBaAFB6C3w7XjLPXAjSDDtqrrWW5r7jsMGoGA1Ud
'' SIG '' HwRjMGEwX6BdoFuGWWh0dHA6Ly93d3cubWljcm9zb2Z0
'' SIG '' LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29mdCUyMFdpbmRv
'' SIG '' d3MlMjBDb2RlJTIwU2lnbmluZyUyMFBDQSUyMDIwMjQu
'' SIG '' Y3JsMHcGCCsGAQUFBwEBBGswaTBnBggrBgEFBQcwAoZb
'' SIG '' aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
'' SIG '' ZXJ0cy9NaWNyb3NvZnQlMjBXaW5kb3dzJTIwQ29kZSUy
'' SIG '' MFNpZ25pbmclMjBQQ0ElMjAyMDI0LmNydDANBgkqhkiG
'' SIG '' 9w0BAQwFAAOCAgEADDlkX+NgsTjRleN8tgJD0V5k5QWc
'' SIG '' rczHYC/AGulvcPtVDcT/LErOmJQvQA6k5XoaoQ1SmY08
'' SIG '' 0Y/+xDcfi877vrnOLFb2H0W+IjWMLI+/fM6KBdNPEQDl
'' SIG '' bydWc+IZy9d5pMlvT3tzgNFpvNuBZZ4YK1dTdzh4o9PU
'' SIG '' tMeMV7+ARlSy+TlcZIh/PpCqrigKbzw6HXPmQ386HBMC
'' SIG '' tDkQAbNdkVw7qHjxEvFwUWf1lAjfwT5Xu2RjIGHHz1kP
'' SIG '' qkRJ+uBnCx77n3SuZmJkJSp2swNgiM78HU8j8H6k0wXl
'' SIG '' JNDJB7YmX3BPfMV1WJTvSdwO6uiu8MUrg6xBoXlcEaBG
'' SIG '' 4WDny6ecTfXBNdSy12ywxdo/DAwU42P7IoB/ZREiCPPb
'' SIG '' wEJoTcds9/+rb6LtNe3sd6RVW9qhImtzjL//uj291F2x
'' SIG '' X9BGLIT405GCAoaWEoGenLSmWcewujxXM7UaJ/iSrMg2
'' SIG '' wI45ozVdFQfeMF8+YXtFVeE6uzVELOUoHwqmasaFxi7k
'' SIG '' 2F/YH6dbTErbCam/8mD6ZijOO+iqSlbLcKM/r8GPpcHz
'' SIG '' 4XEYacSb411dLbeMee04Tv03LU3VeUs7mhPLbTn1Fsyw
'' SIG '' m3L3iidxf+nAXTupyDI+FvH6eEtsbuOQ06Fni29+futP
'' SIG '' H6f8q1SycPjF8LaKe6g8Jvuqm/zGEKMvE5PF76Ixghre
'' SIG '' MIIa2gIBATB2MF8xCzAJBgNVBAYTAlVTMR4wHAYDVQQK
'' SIG '' ExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xMDAuBgNVBAMT
'' SIG '' J01pY3Jvc29mdCBXaW5kb3dzIENvZGUgU2lnbmluZyBQ
'' SIG '' Q0EgMjAyNAITMwAAAIbncZS5Tf8J+wAAAAAAhjANBglg
'' SIG '' hkgBZQMEAgEFAKCBojAZBgkqhkiG9w0BCQMxDAYKKwYB
'' SIG '' BAGCNwIBBDAvBgkqhkiG9w0BCQQxIgQgk3KaJh7gT82p
'' SIG '' eQ/9CvClBlohZ/vjZHs5MlydYp2XKpIwVAYKKwYBBAGC
'' SIG '' NwIBDDFGMESgIoAgAE0AaQBjAHIAbwBzAG8AZgB0ACAA
'' SIG '' TwBmAGYAaQBjAGWhHoAcaHR0cDovL29mZmljZS5taWNy
'' SIG '' b3NvZnQuY29tIDANBgkqhkiG9w0BAQEFAASCAgBFu7Eh
'' SIG '' N1fEQdRRvNzF0HJwWml1RBAIpiB2VA0dFgwtVzWg9qow
'' SIG '' Hf9tD7z2O6Q3/LRA8rjWihhZEX2ol8hR4LFjw70WdsyM
'' SIG '' +vBXuryZkZdMFJhJWZB3ulIq5IUeDqM+xHKv9XSXQDar
'' SIG '' NbXC61p6ltU9X3cFLLMRlIYAUHVqVNf59N+wTxl5Fa+P
'' SIG '' T60d9OvaBK6QYHzaXVcgUhRkMQxB/IaKM3yqv+lLPeRE
'' SIG '' GBIEmb84OL7SihDtt+gtD0KMvA9TAZjHH/cbvINz/+c3
'' SIG '' ugd+IPsNfOAn4zASZW6vvIvTMt9t7i0lBig+6l7odg0x
'' SIG '' 52qujcCAAh9ldZrqy5gxs5mK2p4srjFnzoqRFGZZ5yq6
'' SIG '' kcMgJW25XHjImsH/N5cLMeA6XXhKEP83JUgbt0eAW7In
'' SIG '' 5VCGS8x7WfLonbjt5praij0DpswH1GfS0bKOL4/R47pj
'' SIG '' e4IeVdF3HXSw9kFADYqrNU8I4nNvKUCpYO84idj3HJeT
'' SIG '' LxQcnugnUrvJjuVA8757maGb78xgFUauEj6/vYHC5RMG
'' SIG '' IW1vUbPLdUxTmQiRN2mbgsF+1Ua8mFXRlQMH0Sah4RDS
'' SIG '' uhzxb4Gn2CtxLI7Be6qD5xO+M4M7ImLWLSpMWfwC3vB1
'' SIG '' Kn21WcMXm2HVh7NbikaQMOLIJVDvIdEGzeiN0qvQ/Arh
'' SIG '' s3Av7VZlUD1UQd3FuqGCF5QwgheQBgorBgEEAYI3AwMB
'' SIG '' MYIXgDCCF3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8w
'' SIG '' DQYJYIZIAWUDBAIBBQAwggFSBgsqhkiG9w0BCRABBKCC
'' SIG '' AUEEggE9MIIBOQIBAQYKKwYBBAGEWQoDATAxMA0GCWCG
'' SIG '' SAFlAwQCAQUABCBAMLs3QgIyRV3BYcx1iPOKs18M72mb
'' SIG '' mTbC3llZ+IKEXgIGaHBsUInAGBMyMDI1MDgwODE4MDMy
'' SIG '' NS4zNDlaMASAAgH0oIHRpIHOMIHLMQswCQYDVQQGEwJV
'' SIG '' UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
'' SIG '' UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
'' SIG '' cmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmlj
'' SIG '' YSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRT
'' SIG '' UyBFU046OTIwMC0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1p
'' SIG '' Y3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WgghHqMIIH
'' SIG '' IDCCBQigAwIBAgITMwAAAgkIB+D5XIzmVQABAAACCTAN
'' SIG '' BgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEG
'' SIG '' A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
'' SIG '' ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
'' SIG '' MSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQ
'' SIG '' Q0EgMjAxMDAeFw0yNTAxMzAxOTQyNTVaFw0yNjA0MjIx
'' SIG '' OTQyNTVaMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
'' SIG '' V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
'' SIG '' A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYD
'' SIG '' VQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25z
'' SIG '' MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046OTIwMC0w
'' SIG '' NUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1l
'' SIG '' LVN0YW1wIFNlcnZpY2UwggIiMA0GCSqGSIb3DQEBAQUA
'' SIG '' A4ICDwAwggIKAoICAQDClEow9y4M3f1S9z1xtNEETwWL
'' SIG '' 1vEiiw0oD7SXEdv4sdP0xsVyidv6I2rmEl8PYs9LcZjz
'' SIG '' sWOHI7dQkRL28GP3CXcvY0Zq6nWsHY2QamCZFLF2IlRH
'' SIG '' 6BHx2RkN7ZRDKms7BOo4IGBRlCMkUv9N9/twOzAkpWNs
'' SIG '' M3b/BQxcwhVgsQqtQ8NEPUuiR+GV5rdQHUT4pjihZTkJ
'' SIG '' wraliz0ZbYpUTH5Oki3d3Bpx9qiPriB6hhNfGPjl0PIp
'' SIG '' 23D579rpW6ZmPqPT8j12KX7ySZwNuxs3PYvF/w13GsRX
'' SIG '' kzIbIyLKEPzj9lzmmrF2wjvvUrx9AZw7GLSXk28Dn1XS
'' SIG '' f62hbkFuUGwPFLp3EbRqIVmBZ42wcz5mSIICy3Qs/hwh
'' SIG '' EYhUndnABgNpD5avALOV7sUfJrHDZXX6f9ggbjIA6j2n
'' SIG '' hSASIql8F5LsKBw0RPtDuy3j2CPxtTmZozbLK8TMtxDi
'' SIG '' MCgxTpfg5iYUvyhV4aqaDLwRBsoBRhO/+hwybKnYwXxK
'' SIG '' eeOrsOwQLnaOE5BmFJYWBOFz3d88LBK9QRBgdEH5CLVh
'' SIG '' 7wkgMIeh96cH5+H0xEvmg6t7uztlXX2SV7xdUYPxA3vj
'' SIG '' jV3EkV7abSHD5HHQZTrd3FqsD/VOYACUVBPrxF+kUrZG
'' SIG '' XxYInZTprYMYEq6UIG1DT4pCVP9DcaCLGIOYEJ1g0wID
'' SIG '' AQABo4IBSTCCAUUwHQYDVR0OBBYEFEmL6NHEXTjlvfAv
'' SIG '' QM21dzMWk8rSMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl
'' SIG '' 0mWnG1M1GelyMF8GA1UdHwRYMFYwVKBSoFCGTmh0dHA6
'' SIG '' Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01p
'' SIG '' Y3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEw
'' SIG '' KDEpLmNybDBsBggrBgEFBQcBAQRgMF4wXAYIKwYBBQUH
'' SIG '' MAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lv
'' SIG '' cHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUy
'' SIG '' MFBDQSUyMDIwMTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAw
'' SIG '' FgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/
'' SIG '' BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQBcXnxvODwk
'' SIG '' 4h/jbUBsnFlFtrSuBBZb7wSZfa5lKRMTNfNlmaAC4bd7
'' SIG '' Wo0I5hMxsEJUyupHwh4kD5qkRZczIc0jIABQQ1xDUBa+
'' SIG '' WTxrp/UAqC17ijFCePZKYVjNrHf/Bmjz7FaOI41kxueR
'' SIG '' hwLNIcQ2gmBqDR5W4TS2htRJYyZAs7jfJmbDtTcUOMhE
'' SIG '' l1OWlx/FnvcQbot5VPzaUwiT6Nie8l6PZjoQsuxiasuS
'' SIG '' AmxKIQdsHnJ5QokqwdyqXi1FZDtETVvbXfDsofzTta4e
'' SIG '' n2qf48hzEZwUvbkz5smt890nVAK7kz2crrzN3hpnfFuf
'' SIG '' tp/rXLWTvxPQcfWXiEuIUd2Gg7eR8QtyKtJDU8+PDwEC
'' SIG '' kzoaJjbGCKqx9ESgFJzzrXNwhhX6Rc8g2EU/+63mmqWe
'' SIG '' CF/kJOFg2eJw7au/abESgq3EazyD1VlL+HaX+MBHGzQm
'' SIG '' HtvOm3Ql4wVTN3Wq8X8bCR68qiF5rFasm4RxF6zajZeS
'' SIG '' HC/qS5336/4aMDqsV6O86RlPPCYGJOPtf2MbKO7XJJeL
'' SIG '' /UQN0c3uix5RMTo66dbATxPUFEG5Ph4PHzGjUbEO7D35
'' SIG '' LuEBiiG8YrlMROkGl3fBQl9bWbgw9CIUQbwq5cTaExlf
'' SIG '' EpMdSoydJolUTQD5ELKGz1TJahTidd20wlwi5Bk36XIm
'' SIG '' zsH4Ys15iXRfAjCCB3EwggVZoAMCAQICEzMAAAAVxedr
'' SIG '' ngKbSZkAAAAAABUwDQYJKoZIhvcNAQELBQAwgYgxCzAJ
'' SIG '' BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
'' SIG '' DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
'' SIG '' ZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29m
'' SIG '' dCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eSAyMDEw
'' SIG '' MB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIyNVow
'' SIG '' fDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWlj
'' SIG '' cm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0G
'' SIG '' CSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDk4aZM57Ry
'' SIG '' IQt5osvXJHm9DtWC0/3unAcH0qlsTnXIyjVX9gF/bErg
'' SIG '' 4r25PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPP
'' SIG '' dzLAEBjoYH1qUoNEt6aORmsHFPPFdvWGUNzBRMhxXFEx
'' SIG '' N6AKOG6N7dcP2CZTfDlhAnrEqv1yaa8dq6z2Nr41JmTa
'' SIG '' mDu6GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyF
'' SIG '' Vk3v3byNpOORj7I5LFGc6XBpDco2LXCOMcg1KL3jtIck
'' SIG '' w+DJj361VI/c+gVVmG1oO5pGve2krnopN6zL64NF50Zu
'' SIG '' yjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viS
'' SIG '' kR4dPf0gz3N9QZpGdc3EXzTdEonW/aUgfX782Z5F37Zy
'' SIG '' L9t9X4C626p+Nuw2TPYrbqgSUei/BQOj0XOmTTd0lBw0
'' SIG '' gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlMjgK8
'' SIG '' QmguEOqEUUbi0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDw
'' SIG '' wvoSCtdjbwzJNmSLW6CmgyFdXzB0kZSU2LlQ+QuJYfM2
'' SIG '' BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AFemzFER1y
'' SIG '' 7435UsSFF5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJp
'' SIG '' xq57t7c+auIurQIDAQABo4IB3TCCAdkwEgYJKwYBBAGC
'' SIG '' NxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQUKqdS/mTE
'' SIG '' mr6CkTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJd
'' SIG '' g/Tl0mWnG1M1GelyMFwGA1UdIARVMFMwUQYMKwYBBAGC
'' SIG '' N0yDfQEBMEEwPwYIKwYBBQUHAgEWM2h0dHA6Ly93d3cu
'' SIG '' bWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0
'' SIG '' b3J5Lmh0bTATBgNVHSUEDDAKBggrBgEFBQcDCDAZBgkr
'' SIG '' BgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMC
'' SIG '' AYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV
'' SIG '' 9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEug
'' SIG '' SaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtp
'' SIG '' L2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0w
'' SIG '' Ni0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUF
'' SIG '' BzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
'' SIG '' L2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNy
'' SIG '' dDANBgkqhkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwEx
'' SIG '' JFvhnnJL/Klv6lwUtj5OR2R4sQaTlz0xM7U518JxNj/a
'' SIG '' ZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZnOlNN3Zi6th54
'' SIG '' 2DYunKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZ
'' SIG '' GN5tggz1bSNU5HhTdSRXud2f8449xvNo32X2pFaq95W2
'' SIG '' KFUn0CS9QKC/GbYSEhFdPSfgQJY4rPf5KYnDvBewVIVC
'' SIG '' s/wMnosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8s
'' SIG '' CXgU6ZGyqVvfSaN0DLzskYDSPeZKPmY7T7uG+jIa2Zb0
'' SIG '' j/aRAfbOxnT99kxybxCrdTDFNLB62FD+CljdQDzHVG2d
'' SIG '' Y3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/
'' SIG '' HltEAY5aGZFrDZ+kKNxnGSgkujhLmm77IVRrakURR6nx
'' SIG '' t67I6IleT53S0Ex2tVdUCbFpAUR+fKFhbHP+CrvsQWY9
'' SIG '' af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKiexcd
'' SIG '' FYmNcP7ntdAoGokLjzbaukz5m/8K6TT4JDVnK+ANuOaM
'' SIG '' mdbhIurwJ0I9JZTmdHRbatGePu1+oDEzfbzL6Xu/OHBE
'' SIG '' 0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZqELQdVTNY
'' SIG '' s6FwZvKhggNNMIICNQIBATCB+aGB0aSBzjCByzELMAkG
'' SIG '' A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAO
'' SIG '' BgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
'' SIG '' dCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0
'' SIG '' IEFtZXJpY2EgT3BlcmF0aW9uczEnMCUGA1UECxMeblNo
'' SIG '' aWVsZCBUU1MgRVNOOjkyMDAtMDVFMC1EOTQ3MSUwIwYD
'' SIG '' VQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNl
'' SIG '' oiMKAQEwBwYFKw4DAhoDFQB8762rPTQd7InDCQdb1kgF
'' SIG '' KQkCRKCBgzCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYD
'' SIG '' VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
'' SIG '' MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
'' SIG '' JjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
'' SIG '' QSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA7EB7qTAiGA8y
'' SIG '' MDI1MDgwODEzMzUzN1oYDzIwMjUwODA5MTMzNTM3WjB0
'' SIG '' MDoGCisGAQQBhFkKBAExLDAqMAoCBQDsQHupAgEAMAcC
'' SIG '' AQACAgZNMAcCAQACAhJEMAoCBQDsQc0pAgEAMDYGCisG
'' SIG '' AQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEA
'' SIG '' AgMHoSChCjAIAgEAAgMBhqAwDQYJKoZIhvcNAQELBQAD
'' SIG '' ggEBAFpsAZ5Z5yqNiq2VsfTl0FpVKjpciCmFgXjllL1p
'' SIG '' 5gPZNz9CLquQubxNWc3fzIjCsAaLuYH+EpoP75x0eFbS
'' SIG '' kZMFTNkPNgWL+llqdnV1uT025OWv67SmFzQs9QXNF7Mg
'' SIG '' JVZPuZrWAG0W8MMCXk1TuN62jCWtDLiHZr0NiTRUv09x
'' SIG '' diYfds5Pzh1bKDVj62JleXPz8Ks7i9Cm0o+ruvpOa26h
'' SIG '' 0p3H4QHvtOW+ZU3N/6+A1tuAbQ0w7EMtBvnit8aAYKKR
'' SIG '' DYM0PZaelfZPN2nu4v2jNHN6bxePDEdH0GXki9hN5eOt
'' SIG '' dtr0u3E1yVVrwXLjzvxewiTyobEnVhBYQXrukOcxggQN
'' SIG '' MIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UE
'' SIG '' CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEe
'' SIG '' MBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYw
'' SIG '' JAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0Eg
'' SIG '' MjAxMAITMwAAAgkIB+D5XIzmVQABAAACCTANBglghkgB
'' SIG '' ZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3
'' SIG '' DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCDZfN2RSrabmf1G
'' SIG '' n1/10PlJftayiudvNTgv6DqQkK80BTCB+gYLKoZIhvcN
'' SIG '' AQkQAi8xgeowgecwgeQwgb0EIGgbLB7IvfQCLmUOUZhj
'' SIG '' dUqK8bikfB6ZVVdoTjNwhRM+MIGYMIGApH4wfDELMAkG
'' SIG '' A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAO
'' SIG '' BgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
'' SIG '' dCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0
'' SIG '' IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAIJCAfg+VyM
'' SIG '' 5lUAAQAAAgkwIgQgqJlLEzDVZlk7lHsj1ZLov+XQbJiM
'' SIG '' 3NbX9X1xDUl97oUwDQYJKoZIhvcNAQELBQAEggIAQQrb
'' SIG '' FK6IDmBCzpxgWmogK4kZWsSBzJ/efbfXWy2lKW8WV5BO
'' SIG '' utWq50pIZVFXKD3bOcrx68hi+k0QcTbaj+B0UCGGP6aS
'' SIG '' r5Gulh+OUSqeoqiQKoOhT1UbZ4JceREJTOPQWkA3vRJn
'' SIG '' dpdptIILzjDn7amU96jcfPvRGcAVqU5yPylKWF9qk7ip
'' SIG '' FZ5/9/Ff/6q+wOLxfYfniOb9y8oAYw996LoClgF3Le/a
'' SIG '' GBYp0fN0njdve13mWpXjdRb+j6VraS4xGyr+kB9Eb8HP
'' SIG '' Nnvvewcfmo2Z0sIDZ8GlezH2mLhhuqGGxJo3+mVyfbnY
'' SIG '' FRKoGsFqoLeJoPmt84hmMgz1ZWAQzK2GW4ipqlkqf3+J
'' SIG '' TJpZxjq/l607hTskhTt0zsTGdhhOYiH44mcfMm0l9+6Q
'' SIG '' inIwm35/Dx0PAvKRG8hSOOljXogJxx4lcI7HPymo3z5g
'' SIG '' 5Pnl07+u7hBJSgwwnO5IL9uP428oCbx12K1rP8ZUtHv0
'' SIG '' caZ8hK1ZaEaqyeicwqmms9kE13sObboBqfqOa+zYNX/f
'' SIG '' zj8u2zPgQGFMayfSefIXoYiC34NFOAVvxt6jfk0LcE2z
'' SIG '' hmTMyuxORAiYfkTrEUykTNlRVvPFMCD7XMBDqmzfv7XX
'' SIG '' mvJzrLFaed64u5B7s7GwaYR5i1FFyNu2kuAive/YqSQR
'' SIG '' TyjaxyNJOvEfICdg+po=
'' SIG '' End signature block
