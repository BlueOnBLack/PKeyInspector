using namespace System
using namespace System.IO
using namespace System.Net
using namespace System.Web
using namespace System.Numerics
using namespace System.Security.Cryptography
using namespace System.Collections.Generic
using namespace System.Drawing
using namespace System.IO.Compression
using namespace System.Management.Automation
using namespace System.Net
using namespace System.Diagnostics
using namespace System.Reflection
using namespace System.Reflection.Emit
using namespace System.Runtime.InteropServices
using namespace System.Security.AccessControl
using namespace System.Security.Principal
using namespace System.ServiceProcess
using namespace System.Text
using namespace System.Text.RegularExpressions
using namespace System.Threading
using namespace System.Windows.Forms

param (
    [switch]$AutoMode,
    [switch]$RunHWID,
    [switch]$RunoHook,
    [switch]$RunVolume,
    [switch]$RunTsforge,
    [switch]$RunUpgrade,
    [switch]$RunCheckActivation,
    [switch]$RunWmiRepair,
    [switch]$RunTokenStoreReset,
    [switch]$RunUninstallLicenses,
    [switch]$RunScrubOfficeC2R,
    [switch]$RunOfficeLicenseInstaller,
    [switch]$RunOfficeOnlineInstallation,
    
    # Office pattern for Auto Mode
    [string]$LicensePattern = $null
)

Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

$ProgressPreference = 'SilentlyContinue'

<#
# Auto Mode example !
# Remove Store, Install Windows & office *365app* licenses, 
# Also, Activate with hwid/kms38, also install oHook bypass

* Option A
powershell -ep bypass -nop -f Activate_.ps1 -AutoMode -RunHWID

* Option B
Set-Location 'C:\Users\Administrator\Desktop'
.\Activate.ps1 -AutoMode -RunTokenStoreReset -RunOfficeLicenseInstaller -RunHWID -RunoHook -LicensePattern "365App"

* Option C
-- icm -scr ([scriptblock]::Create((irm officertool.org/Download/Activate.php))) -arg $true,$true
-- powershell -ep bypass -nop -c icm -scr ([scriptblock]::Create((irm officertool.org/Download/Activate.php))) -arg $true,$true
#>

<#
.SYNOPSIS
This script automates the Windows activation process. 
It checks system requirements, verifies the PowerShell version, 
and handles manual activation through HWID, KMS38, and OHooks methods.

.DESCRIPTION
This script requires PowerShell 3.0 or higher and administrator privileges to run. 
It uses methods for activating Windows, including HWID and KMS38 activations. 
Links for manual activation guides are provided.

.REMOTE EXECUTION:
- To remotely execute the script, use the following command:
  irm tinyurl.com/tshook | iex

.MANUAL ACTIVATION GUIDES:
- HWID Activation: https://massgrave.dev/manual_hwid_activation
- KMS38 Activation: https://massgrave.dev/manual_kms38_activation
- OHooks Activation: https://massgrave.dev/manual_ohook_activation

.VERSION:
- This is the PowerShell 1.0 version of HWID_Activation.cmd
- Credits: WindowsAddict, Mass Project

.CREDITS:
- Code logic borrowed from abbodi1406 KMS_VL_ALL & R`Tool Projects
#>

# Kernel-Mode Windows Versions
# https://www.geoffchappell.com/studies/windows/km/versions.htm

# ZwQuerySystemInformation
# https://www.geoffchappell.com/studies/windows/km/ntoskrnl/api/ex/sysinfo/query.htm

# RtlGetNtVersionNumbers Function
# The RtlGetNtVersionNumbers function gets Windows version numbers directly from NTDLL.
# https://www.geoffchappell.com/studies/windows/win32/ntdll/api/ldrinit/getntversionnumbers.htm

# RtlGetVersion function (wdm.h)
# https://learn.microsoft.com/en-us/windows/win32/devnotes/rtlgetversion
# https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/nf-wdm-rtlgetversion

# Process Environment Block (PEB)
# https://www.geoffchappell.com/studies/windows/km/ntoskrnl/inc/api/pebteb/peb/index.htm

# KUSER_SHARED_DATA
# https://www.geoffchappell.com/studies/windows/km/ntoskrnl/inc/api/ntexapi_x/kuser_shared_data/index.htm

# KUSER_SHARED_DATA structure (ntddk.h)
# https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/ntddk/ns-ntddk-kuser_shared_data

# The read-only user-mode address for the shared data is 0x7FFE0000, both in 32-bit and 64-bit Windows.
# The only formal definition among headers in the Windows Driver Kit (WDK) or the Software Development Kit (SDK) is in assembly language headers: KS386.
# INC from the WDK and KSAMD64.
# INC from the SDK both define MM_SHARED_USER_DATA_VA for the user-mode address.

# That they also define USER_SHARED_DATA for the kernel-mode address suggests that they too are intended for kernel-mode programming,
# albeit of a sort that is at least aware of what address works for user-mode access.

# Among relatively large structures,
# the KUSER_SHARED_DATA is highly unusual for having exactly the same layout in 32-bit and 64-bit Windows.
# This is because the one instance must be simultaneously accessible by both 32-bit and 64-bit code on 64-bit Windows,
# and it's desired that 32-bit user-mode code can run unchanged on both 32-bit and 64-bit Windows.

# 2.2.9.6 OSEdition Enumeration
# https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-mde2/d92ead8f-faf3-47a8-a341-1921dc2c463b

<#
Example's from MAS' AIO

$editionIDPtr = [IntPtr]::Zero
$hresults = $Global:PKHElper::GetEditionNameFromId(126,[ref]$editionIDPtr)
if ($hresults -eq 0) {
    $editionID = [Marshal]::PtrToStringUni($editionIDPtr)
    Write-Host "BrandingInfo: 126 > Edition: $editionID"
}

[int]$brandingInfo = 0
$hresults = $Global:PKHElper::GetEditionIdFromName("enterprisesn", [ref]$brandingInfo)
if ($brandingInfo -ne 0) {
    Write-Host "Edition: enterprisesn > BrandingInfo: $brandingInfo"
}

Full Table, include Name, Sku, etc etc provide by [abbodi1406]
#>
$Global:productTypeTable = @'
ProductID,OSEdition,DWORD
ultimate,PRODUCT_ULTIMATE,0x00000001
homebasic,PRODUCT_HOME_BASIC,0x00000002
homepremium,PRODUCT_HOME_PREMIUM,0x00000003
enterprise,PRODUCT_ENTERPRISE,0x00000004
homebasicn,PRODUCT_HOME_BASIC_N,0x00000005
business,PRODUCT_BUSINESS,0x00000006
serverstandard,PRODUCT_STANDARD_SERVER,0x00000007
serverdatacenter,PRODUCT_DATACENTER_SERVER,0x00000008
serversbsstandard,PRODUCT_SMALLBUSINESS_SERVER,0x00000009
serverenterprise,PRODUCT_ENTERPRISE_SERVER,0x0000000A
starter,PRODUCT_STARTER,0x0000000B
serverdatacentercore,PRODUCT_DATACENTER_SERVER_CORE,0x0000000C
serverstandardcore,PRODUCT_STANDARD_SERVER_CORE,0x0000000D
serverenterprisecore,PRODUCT_ENTERPRISE_SERVER_CORE,0x0000000E
serverenterpriseia64,PRODUCT_ENTERPRISE_SERVER_IA64,0x0000000F
businessn,PRODUCT_BUSINESS_N,0x00000010
serverweb,PRODUCT_WEB_SERVER,0x00000011
serverhpc,PRODUCT_CLUSTER_SERVER,0x00000012
serverhomestandard,PRODUCT_HOME_SERVER,0x00000013
serverstorageexpress,PRODUCT_STORAGE_EXPRESS_SERVER,0x00000014
serverstoragestandard,PRODUCT_STORAGE_STANDARD_SERVER,0x00000015
serverstorageworkgroup,PRODUCT_STORAGE_WORKGROUP_SERVER,0x00000016
serverstorageenterprise,PRODUCT_STORAGE_ENTERPRISE_SERVER,0x00000017
serverwinsb,PRODUCT_SERVER_FOR_SMALLBUSINESS,0x00000018
serversbspremium,PRODUCT_SMALLBUSINESS_SERVER_PREMIUM,0x00000019
homepremiumn,PRODUCT_HOME_PREMIUM_N,0x0000001A
enterprisen,PRODUCT_ENTERPRISE_N,0x0000001B
ultimaten,PRODUCT_ULTIMATE_N,0x0000001C
serverwebcore,PRODUCT_WEB_SERVER_CORE,0x0000001D
servermediumbusinessmanagement,PRODUCT_MEDIUMBUSINESS_SERVER_MANAGEMENT,0x0000001E
servermediumbusinesssecurity,PRODUCT_MEDIUMBUSINESS_SERVER_SECURITY,0x0000001F
servermediumbusinessmessaging,PRODUCT_MEDIUMBUSINESS_SERVER_MESSAGING,0x00000020
serverwinfoundation,PRODUCT_SERVER_FOUNDATION,0x00000021
serverhomepremium,PRODUCT_HOME_PREMIUM_SERVER,0x00000022
serverwinsbv,PRODUCT_SERVER_FOR_SMALLBUSINESS_V,0x00000023
serverstandardv,PRODUCT_STANDARD_SERVER_V,0x00000024
serverdatacenterv,PRODUCT_DATACENTER_SERVER_V,0x00000025
serverenterprisev,PRODUCT_ENTERPRISE_SERVER_V,0x00000026
serverdatacentervcore,PRODUCT_DATACENTER_SERVER_CORE_V,0x00000027
serverstandardvcore,PRODUCT_STANDARD_SERVER_CORE_V,0x00000028
serverenterprisevcore,PRODUCT_ENTERPRISE_SERVER_CORE_V,0x00000029
serverhypercore,PRODUCT_HYPERV,0x0000002A
serverstorageexpresscore,PRODUCT_STORAGE_EXPRESS_SERVER_CORE,0x0000002B
serverstoragestandardcore,PRODUCT_STORAGE_STANDARD_SERVER_CORE,0x0000002C
serverstorageworkgroupcore,PRODUCT_STORAGE_WORKGROUP_SERVER_CORE,0x0000002D
serverstorageenterprisecore,PRODUCT_STORAGE_ENTERPRISE_SERVER_CORE,0x0000002E
startern,PRODUCT_STARTER_N,0x0000002F
professional,PRODUCT_PROFESSIONAL,0x00000030
professionaln,PRODUCT_PROFESSIONAL_N,0x00000031
serversolution,PRODUCT_SB_SOLUTION_SERVER,0x00000032
serverforsbsolutions,PRODUCT_SERVER_FOR_SB_SOLUTIONS,0x00000033
serversolutionspremium,PRODUCT_STANDARD_SERVER_SOLUTIONS,0x00000034
serversolutionspremiumcore,PRODUCT_STANDARD_SERVER_SOLUTIONS_CORE,0x00000035
serversolutionem,PRODUCT_SB_SOLUTION_SERVER_EM,0x00000036
serverforsbsolutionsem,PRODUCT_SERVER_FOR_SB_SOLUTIONS_EM,0x00000037
serverembeddedsolution,PRODUCT_SOLUTION_EMBEDDEDSERVER,0x00000038
serverembeddedsolutioncore,PRODUCT_SOLUTION_EMBEDDEDSERVER_CORE,0x00000039
professionalembedded,PRODUCT_PROFESSIONAL_EMBEDDED,0x0000003A
serveressentialmanagement,PRODUCT_ESSENTIALBUSINESS_SERVER_MGMT,0x0000003B
serveressentialadditional,PRODUCT_ESSENTIALBUSINESS_SERVER_ADDL,0x0000003C
serveressentialmanagementsvc,PRODUCT_ESSENTIALBUSINESS_SERVER_MGMTSVC,0x0000003D
serveressentialadditionalsvc,PRODUCT_ESSENTIALBUSINESS_SERVER_ADDLSVC,0x0000003E
serversbspremiumcore,PRODUCT_SMALLBUSINESS_SERVER_PREMIUM_CORE,0x0000003F
serverhpcv,PRODUCT_CLUSTER_SERVER_V,0x00000040
embedded,PRODUCT_EMBEDDED,0x00000041
startere,PRODUCT_STARTER_E,0x00000042
homebasice,PRODUCT_HOME_BASIC_E,0x00000043
homepremiume,PRODUCT_HOME_PREMIUM_E,0x00000044
professionale,PRODUCT_PROFESSIONAL_E,0x00000045
enterprisee,PRODUCT_ENTERPRISE_E,0x00000046
ultimatee,PRODUCT_ULTIMATE_E,0x00000047
enterpriseeval,PRODUCT_ENTERPRISE_EVALUATION,0x00000048
prerelease,PRODUCT_PRERELEASE,0x0000004A
servermultipointstandard,PRODUCT_MULTIPOINT_STANDARD_SERVER,0x0000004C
servermultipointpremium,PRODUCT_MULTIPOINT_PREMIUM_SERVER,0x0000004D
serverstandardeval,PRODUCT_STANDARD_EVALUATION_SERVER,0x0000004F
serverdatacentereval,PRODUCT_DATACENTER_EVALUATION_SERVER,0x00000050
prereleasearm,PRODUCT_PRODUCT_PRERELEASE_ARM,0x00000051
prereleasen,PRODUCT_PRODUCT_PRERELEASE_N,0x00000052
enterpriseneval,PRODUCT_ENTERPRISE_N_EVALUATION,0x00000054
embeddedautomotive,PRODUCT_EMBEDDED_AUTOMOTIVE,0x00000055
embeddedindustrya,PRODUCT_EMBEDDED_INDUSTRY_A,0x00000056
thinpc,PRODUCT_THINPC,0x00000057
embeddeda,PRODUCT_EMBEDDED_A,0x00000058
embeddedindustry,PRODUCT_EMBEDDED_INDUSTRY,0x00000059
embeddede,PRODUCT_EMBEDDED_E,0x0000005A
embeddedindustrye,PRODUCT_EMBEDDED_INDUSTRY_E,0x0000005B
embeddedindustryae,PRODUCT_EMBEDDED_INDUSTRY_A_E,0x0000005C
professionalplus,PRODUCT_PRODUCT_PROFESSIONAL_PLUS,0x0000005D
serverstorageworkgroupeval,PRODUCT_STORAGE_WORKGROUP_EVALUATION_SERVER,0x0000005F
serverstoragestandardeval,PRODUCT_STORAGE_STANDARD_EVALUATION_SERVER,0x00000060
corearm,PRODUCT_CORE_ARM,0x00000061
coren,PRODUCT_CORE_N,0x00000062
corecountryspecific,PRODUCT_CORE_COUNTRYSPECIFIC,0x00000063
coresinglelanguage,PRODUCT_CORE_SINGLELANGUAGE,0x00000064
core,PRODUCT_CORE,0x00000065
professionalwmc,PRODUCT_PROFESSIONAL_WMC,0x00000067
mobilecore,PRODUCT_MOBILE_CORE,0x00000068
embeddedindustryeval,PRODUCT_EMBEDDED_INDUSTRY_EVAL,0x00000069
embeddedindustryeeval,PRODUCT_EMBEDDED_INDUSTRY_E_EVAL,0x0000006A
embeddedeval,PRODUCT_EMBEDDED_EVAL,0x0000006B
embeddedeeval,PRODUCT_EMBEDDED_E_EVAL,0x0000006C
coresystemserver,PRODUCT_NANO_SERVER,0x0000006D
servercloudstorage,PRODUCT_CLOUD_STORAGE_SERVER,0x0000006E
coreconnected,PRODUCT_CORE_CONNECTED,0x0000006F
professionalstudent,PRODUCT_PROFESSIONAL_STUDENT,0x00000070
coreconnectedn,PRODUCT_CORE_CONNECTED_N,0x00000071
professionalstudentn,PRODUCT_PROFESSIONAL_STUDENT_N,0x00000072
coreconnectedsinglelanguage,PRODUCT_CORE_CONNECTED_SINGLELANGUAGE,0x00000073
coreconnectedcountryspecific,PRODUCT_CORE_CONNECTED_COUNTRYSPECIFIC,0x00000074
connectedcar,PRODUCT_CONNECTED_CAR,0x00000075
industryhandheld,PRODUCT_INDUSTRY_HANDHELD,0x00000076
ppipro,PRODUCT_PPI_PRO,0x00000077
serverarm64,PRODUCT_ARM64_SERVER,0x00000078
education,PRODUCT_EDUCATION,0x00000079
educationn,PRODUCT_EDUCATION_N,0x0000007A
iotuap,PRODUCT_IOTUAP,0x0000007B
serverhi,PRODUCT_CLOUD_HOST_INFRASTRUCTURE_SERVER,0x0000007C
enterprises,PRODUCT_ENTERPRISE_S,0x0000007D
enterprisesn,PRODUCT_ENTERPRISE_S_N,0x0000007E
professionals,PRODUCT_PROFESSIONAL_S,0x0000007F
professionalsn,PRODUCT_PROFESSIONAL_S_N,0x00000080
enterpriseseval,PRODUCT_ENTERPRISE_S_EVALUATION,0x00000081
enterprisesneval,PRODUCT_ENTERPRISE_S_N_EVALUATION,0x00000082
iotuapcommercial,PRODUCT_IOTUAPCOMMERCIAL,0x00000083
mobileenterprise,PRODUCT_MOBILE_ENTERPRISE,0x00000085
analogonecore,PRODUCT_HOLOGRAPHIC,0x00000087
holographic,PRODUCT_HOLOGRAPHIC_BUSINESS,0x00000088
professionalsinglelanguage,PRODUCT_PRO_SINGLE_LANGUAGE,0x0000008A
professionalcountryspecific,PRODUCT_PRO_CHINA,0x0000008B
enterprisesubscription,PRODUCT_ENTERPRISE_SUBSCRIPTION,0x0000008C
enterprisesubscriptionn,PRODUCT_ENTERPRISE_SUBSCRIPTION_N,0x0000008D
serverdatacenternano,PRODUCT_DATACENTER_NANO_SERVER,0x0000008F
serverstandardnano,PRODUCT_STANDARD_NANO_SERVER,0x00000090
serverdatacenteracor,PRODUCT_DATACENTER_A_SERVER_CORE,0x00000091
serverstandardacor,PRODUCT_STANDARD_A_SERVER_CORE,0x00000092
serverdatacentercor,PRODUCT_DATACENTER_WS_SERVER_CORE,0x00000093
serverstandardcor,PRODUCT_STANDARD_WS_SERVER_CORE,0x00000094
utilityvm,PRODUCT_UTILITY_VM,0x00000095
serverdatacenterevalcor,PRODUCT_DATACENTER_EVALUATION_SERVER_CORE,0x0000009F
serverstandardevalcor,PRODUCT_STANDARD_EVALUATION_SERVER_CORE,0x000000A0
professionalworkstation,PRODUCT_PRO_WORKSTATION,0x000000A1
professionalworkstationn,PRODUCT_PRO_WORKSTATION_N,0x000000A2
serverazure,PRODUCT_AZURE_SERVER,0x000000A3
professionaleducation,PRODUCT_PRO_FOR_EDUCATION,0x000000A4
professionaleducationn,PRODUCT_PRO_FOR_EDUCATION_N,0x000000A5
serverazurecor,PRODUCT_AZURE_SERVER_CORE,0x000000A8
serverazurenano,PRODUCT_AZURE_NANO_SERVER,0x000000A9
enterpriseg,PRODUCT_ENTERPRISEG,0x000000AB
enterprisegn,PRODUCT_ENTERPRISEGN,0x000000AC
businesssubscription,PRODUCT_BUSINESS,0x000000AD
businesssubscriptionn,PRODUCT_BUSINESS_N,0x000000AE
serverrdsh,PRODUCT_SERVERRDSH,0x000000AF
cloud,PRODUCT_CLOUD,0x000000B2
cloudn,PRODUCT_CLOUDN,0x000000B3
hubos,PRODUCT_HUBOS,0x000000B4
onecoreupdateos,PRODUCT_ONECOREUPDATEOS,0x000000B6
cloude,PRODUCT_CLOUDE,0x000000B7
andromeda,PRODUCT_ANDROMEDA,0x000000B8
iotos,PRODUCT_IOTOS,0x000000B9
clouden,PRODUCT_CLOUDEN,0x000000BA
iotedgeos,PRODUCT_IOTEDGEOS,0x000000BB
iotenterprise,PRODUCT_IOTENTERPRISE,0x000000BC
modernpc,PRODUCT_LITE,0x000000BD
iotenterprises,PRODUCT_IOTENTERPRISES,0x000000BF
systemos,PRODUCT_XBOX_SYSTEMOS,0x000000C0
nativeos,PRODUCT_XBOX_NATIVEOS,0x000000C1
gamecorexbox,PRODUCT_XBOX_GAMEOS,0x000000C2
gameos,PRODUCT_XBOX_ERAOS,0x000000C3
durangohostos,PRODUCT_XBOX_DURANGOHOSTOS,0x000000C4
scarletthostos,PRODUCT_XBOX_SCARLETTHOSTOS,0x000000C5
keystone,PRODUCT_XBOX_KEYSTONE,0x000000C6
cloudhost,PRODUCT_AZURE_SERVER_CLOUDHOST,0x000000C7
cloudmos,PRODUCT_AZURE_SERVER_CLOUDMOS,0x000000C8
cloudcore,PRODUCT_AZURE_SERVER_CLOUDCORE,0x000000C9
cloudeditionn,PRODUCT_CLOUDEDITIONN,0x000000CA
cloudedition,PRODUCT_CLOUDEDITION,0x000000CB
winvos,PRODUCT_VALIDATION,0x000000CC
iotenterprisesk,PRODUCT_IOTENTERPRISESK,0x000000CD
iotenterprisek,PRODUCT_IOTENTERPRISEK,0x000000CE
iotenterpriseseval,PRODUCT_IOTENTERPRISESEVAL,0x000000CF
agentbridge,PRODUCT_AZURE_SERVER_AGENTBRIDGE,0x000000D0
nanohost,PRODUCT_AZURE_SERVER_NANOHOST,0x000000D1
wnc,PRODUCT_WNC,0x000000D2
serverazurestackhcicor,PRODUCT_AZURESTACKHCI_SERVER_CORE,0x00000196
serverturbine,PRODUCT_DATACENTER_SERVER_AZURE_EDITION,0x00000197
serverturbinecor,PRODUCT_DATACENTER_SERVER_CORE_AZURE_EDITION,0x00000198
'@ | ConvertFrom-Csv

<#
wuerror.h
https://github.com/larsch/wunow/blob/master/wunow/WUError.cs
https://github.com/microsoft/IIS.Setup/blob/main/iisca/lib/wuerror.h
https://learn.microsoft.com/en-us/troubleshoot/windows-client/installing-updates-features-roles/common-windows-update-errors
#>
$Global:WU_ERR_TABLE = @'
ERROR, MESSEGE
0x80D02002, "The operation timed out"
0x8024A10A, "Indicates that the Windows Update Service is shutting down."
0x00240001, "Windows Update Agent was stopped successfully."
0x00240002, "Windows Update Agent updated itself."
0x00240003, "Operation completed successfully but there were errors applying the updates."
0x00240004, "A callback was marked to be disconnected later because the request to disconnect the operation came while a callback was executing."
0x00240005, "The system must be restarted to complete installation of the update."
0x00240006, "The update to be installed is already installed on the system."
0x00240007, "The update to be removed is not installed on the system."
0x00240008, "The update to be downloaded has already been downloaded."
0x00242015, "The installation operation for the update is still in progress."
0x80240001, "Windows Update Agent was unable to provide the service."
0x80240002, "The maximum capacity of the service was exceeded."
0x80240003, "An ID cannot be found."
0x80240004, "The object could not be initialized."
0x80240005, "The update handler requested a byte range overlapping a previously requested range."
0x80240006, "The requested number of byte ranges exceeds the maximum number 2^31 - 1)."
0x80240007, "The index to a collection was invalid."
0x80240008, "The key for the item queried could not be found."
0x80240009, "Another conflicting operation was in progress. Some operations such as installation cannot be performed twice simultaneously."
0x8024000A, "Cancellation of the operation was not allowed."
0x8024000B, "Operation was cancelled."
0x8024000C, "No operation was required."
0x8024000D, "Windows Update Agent could not find required information in the update's XML data."
0x8024000E, "Windows Update Agent found invalid information in the update's XML data."
0x8024000F, "Circular update relationships were detected in the metadata."
0x80240010, "Update relationships too deep to evaluate were evaluated."
0x80240011, "An invalid update relationship was detected."
0x80240012, "An invalid registry value was read."
0x80240013, "Operation tried to add a duplicate item to a list."
0x80240014, "Updates requested for install are not installable by caller."
0x80240016, "Operation tried to install while another installation was in progress or the system was pending a mandatory restart."
0x80240017, "Operation was not performed because there are no applicable updates."
0x80240018, "Operation failed because a required user token is missing."
0x80240019, "An exclusive update cannot be installed with other updates at the same time."
0x8024001A, "A policy value was not set."
0x8024001B, "The operation could not be performed because the Windows Update Agent is self-updating."
0x8024001D, "An update contains invalid metadata."
0x8024001E, "Operation did not complete because the service or system was being shut down."
0x8024001F, "Operation did not complete because the network connection was unavailable."
0x80240020, "Operation did not complete because there is no logged-on interactive user."
0x80240021, "Operation did not complete because it timed out."
0x80240022, "Operation failed for all the updates."
0x80240023, "The license terms for all updates were declined."
0x80240024, "There are no updates."
0x80240025, "Group Policy settings prevented access to Windows Update."
0x80240026, "The type of update is invalid."
0x80240027, "The URL exceeded the maximum length."
0x80240028, "The update could not be uninstalled because the request did not originate from a WSUS server."
0x80240029, "Search may have missed some updates before there is an unlicensed application on the system."
0x8024002A, "A component required to detect applicable updates was missing."
0x8024002B, "An operation did not complete because it requires a newer version of server."
0x8024002C, "A delta-compressed update could not be installed because it required the source."
0x8024002D, "A full-file update could not be installed because it required the source."
0x8024002E, "Access to an unmanaged server is not allowed."
0x8024002F, "Operation did not complete because the DisableWindowsUpdateAccess policy was set."
0x80240030, "The format of the proxy list was invalid."
0x80240031, "The file is in the wrong format."
0x80240032, "The search criteria string was invalid."
0x80240033, "License terms could not be downloaded."
0x80240034, "Update failed to download."
0x80240035, "The update was not processed."
0x80240036, "The object's current state did not allow the operation."
0x80240037, "The functionality for the operation is not supported."
0x80240038, "The downloaded file has an unexpected content type."
0x80240039, "Agent is asked by server to resync too many times."
0x80240040, "WUA API method does not run on Server Core installation."
0x80240041, "Service is not available while sysprep is running."
0x80240042, "The update service is no longer registered with AU."
0x80240043, "There is no support for WUA UI."
0x80240044, "Only administrators can perform this operation on per-machine updates."
0x80240045, "A search was attempted with a scope that is not currently supported for this type of search."
0x80240046, "The URL does not point to a file."
0x80240047, "The operation requested is not supported."
0x80240048, "The featured update notification info returned by the server is invalid."
0x80240FFF, "An operation failed due to reasons not covered by another error code."
0x80241001, "Search may have missed some updates because the Windows Installer is less than version 3.1."
0x80241002, "Search may have missed some updates because the Windows Installer is not configured."
0x80241003, "Search may have missed some updates because policy has disabled Windows Installer patching."
0x80241004, "An update could not be applied because the application is installed per-user."
0x80241FFF, "Search may have missed some updates because there was a failure of the Windows Installer."
0x80244000, "WU_E_PT_SOAPCLIENT_* error codes map to the SOAPCLIENT_ERROR enum of the ATL Server Library."
0x80244001, "Same as SOAPCLIENT_INITIALIZE_ERROR - initialization of the SOAP client failed, possibly because of an MSXML installation failure."
0x80244002, "Same as SOAPCLIENT_OUTOFMEMORY - SOAP client failed because it ran out of memory."
0x80244003, "Same as SOAPCLIENT_GENERATE_ERROR - SOAP client failed to generate the request."
0x80244004, "Same as SOAPCLIENT_CONNECT_ERROR - SOAP client failed to connect to the server."
0x80244005, "Same as SOAPCLIENT_SEND_ERROR - SOAP client failed to send a message for reasons of WU_E_WINHTTP_* error codes."
0x80244006, "Same as SOAPCLIENT_SERVER_ERROR - SOAP client failed because there was a server error."
0x80244007, "Same as SOAPCLIENT_SOAPFAULT - SOAP client failed because there was a SOAP fault for reasons of WU_E_PT_SOAP_* error codes."
0x80244008, "Same as SOAPCLIENT_PARSEFAULT_ERROR - SOAP client failed to parse a SOAP fault."
0x80244009, "Same as SOAPCLIENT_READ_ERROR - SOAP client failed while reading the response from the server."
0x8024400A, "Same as SOAPCLIENT_PARSE_ERROR - SOAP client failed to parse the response from the server."
0x8024400B, "Same as SOAP_E_VERSION_MISMATCH - SOAP client found an unrecognizable namespace for the SOAP envelope."
0x8024400C, "Same as SOAP_E_MUST_UNDERSTAND - SOAP client was unable to understand a header."
0x8024400D, "Same as SOAP_E_CLIENT - SOAP client found the message was malformed; fix before resending."
0x8024400E, "Same as SOAP_E_SERVER - The SOAP message could not be processed due to a server error; resend later."
0x8024400F, "There was an unspecified Windows Management Instrumentation WMI) error."
0x80244010, "The number of round trips to the server exceeded the maximum limit."
0x80244011, "WUServer policy value is missing in the registry."
0x80244012, "Initialization failed because the object was already initialized."
0x80244013, "The computer name could not be determined."
0x80244015, "The reply from the server indicates that the server was changed or the cookie was invalid; refresh the state of the internal cache and retry."
0x80244016, "Same as HTTP status 400 - the server could not process the request due to invalid syntax."
0x80244017, "Same as HTTP status 401 - the requested resource requires user authentication."
0x80244018, "Same as HTTP status 403 - server understood the request, but declined to fulfill it."
0x8024401A, "Same as HTTP status 405 - the HTTP method is not allowed."
0x8024401B, "Same as HTTP status 407 - proxy authentication is required."
0x8024401C, "Same as HTTP status 408 - the server timed out waiting for the request."
0x8024401D, "Same as HTTP status 409 - the request was not completed due to a conflict with the current state of the resource."
0x8024401E, "Same as HTTP status 410 - requested resource is no longer available at the server."
0x8024401F, "Same as HTTP status 500 - an error internal to the server prevented fulfilling the request."
0x80244020, "Same as HTTP status 500 - server does not support the functionality required to fulfill the request."
0x80244021, "Same as HTTP status 502 - the server, while acting as a gateway or proxy, received an invalid response from the upstream server it accessed in attempting to fulfill the request."
0x80244022, "Same as HTTP status 503 - the service is temporarily overloaded."
0x80244023, "Same as HTTP status 503 - the request was timed out waiting for a gateway."
0x80244024, "Same as HTTP status 505 - the server does not support the HTTP protocol version used for the request."
0x80244025, "Operation failed due to a changed file location; refresh internal state and resend."
0x80244026, "Operation failed because Windows Update Agent does not support registration with a non-WSUS server."
0x80244027, "The server returned an empty authentication information list."
0x80244028, "Windows Update Agent was unable to create any valid authentication cookies."
0x80244029, "A configuration property value was wrong."
0x8024402A, "A configuration property value was missing."
0x8024402B, "The HTTP request could not be completed and the reason did not correspond to any of the WU_E_PT_HTTP_* error codes."
0x8024402C, "Same as ERROR_WINHTTP_NAME_NOT_RESOLVED - the proxy server or target server name cannot be resolved."
0x8024502D, "Windows Update Agent failed to download a redirector cabinet file with a new redirectorId value from the server during the recovery."
0x8024502E, "A redirector recovery action did not complete because the server is managed."
0x8024402F, "External cab file processing completed with some errors."
0x80244030, "The external cab processor initialization did not complete."
0x80244031, "The format of a metadata file was invalid."
0x80244032, "External cab processor found invalid metadata."
0x80244033, "The file digest could not be extracted from an external cab file."
0x80244034, "An external cab file could not be decompressed."
0x80244035, "External cab processor was unable to get file locations."
0x80240436, "The server does not support category-specific search; Full catalog search has to be issued instead."
0x80244FFF, "A communication error not covered by another WU_E_PT_* error code."
0x80245001, "The redirector XML document could not be loaded into the DOM class."
0x80245002, "The redirector XML document is missing some required information."
0x80245003, "The redirectorId in the downloaded redirector cab is less than in the cached cab."
0x80245FFF, "The redirector failed for reasons not covered by another WU_E_REDIRECTOR_* error code."
0x8024C001, "A driver was skipped."
0x8024C002, "A property for the driver could not be found. It may not conform with required specifications."
0x8024C003, "The registry type read for the driver does not match the expected type."
0x8024C004, "The driver update is missing metadata."
0x8024C005, "The driver update is missing a required attribute."
0x8024C006, "Driver synchronization failed."
0x8024C007, "Information required for the synchronization of applicable printers is missing."
0x8024CFFF, "A driver error not covered by another WU_E_DRV_* code."
0x80248000, "An operation failed because Windows Update Agent is shutting down."
0x80248001, "An operation failed because the data store was in use."
0x80248002, "The current and expected states of the data store do not match."
0x80248003, "The data store is missing a table."
0x80248004, "The data store contains a table with unexpected columns."
0x80248005, "A table could not be opened because the table is not in the data store."
0x80248006, "The current and expected versions of the data store do not match."
0x80248007, "The information requested is not in the data store."
0x80248008, "The data store is missing required information or has a NULL in a table column that requires a non-null value."
0x80248009, "The data store is missing required information or has a reference to missing license terms, file, localized property or linked row."
0x8024800A, "The update was not processed because its update handler could not be recognized."
0x8024800B, "The update was not deleted because it is still referenced by one or more services."
0x8024800C, "The data store section could not be locked within the allotted time."
0x8024800D, "The category was not added because it contains no parent categories and is not a top-level category itself."
0x8024800E, "The row was not added because an existing row has the same primary key."
0x8024800F, "The data store could not be initialized because it was locked by another process."
0x80248010, "The data store is not allowed to be registered with COM in the current process."
0x80248011, "Could not create a data store object in another process."
0x80248013, "The server sent the same update to the client with two different revision IDs."
0x80248014, "An operation did not complete because the service is not in the data store."
0x80248015, "An operation did not complete because the registration of the service has expired."
0x80248016, "A request to hide an update was declined because it is a mandatory update or because it was deployed with a deadline."
0x80248017, "A table was not closed because it is not associated with the session."
0x80248018, "A table was not closed because it is not associated with the session."
0x80248019, "A request to remove the Windows Update service or to unregister it with Automatic Updates was declined because it is a built-in service and/or Automatic Updates cannot fall back to another service."
0x8024801A, "A request was declined because the operation is not allowed."
0x8024801B, "The schema of the current data store and the schema of a table in a backup XML document do not match."
0x8024801C, "The data store requires a session reset; release the session and retry with a new session."
0x8024801D, "A data store operation did not complete because it was requested with an impersonated identity."
0x80248FFF, "A data store error not covered by another WU_E_DS_* code."
0x80249001, "Parsing of the rule file failed."
0x80249002, "Failed to get the requested inventory type from the server."
0x80249003, "Failed to upload inventory result to the server."
0x80249004, "There was an inventory error not covered by another error code."
0x80249005, "A WMI error occurred when enumerating the instances for a particular class."
0x8024A000, "Automatic Updates was unable to service incoming requests."
0x8024A002, "The old version of the Automatic Updates client has stopped because the WSUS server has been upgraded."
0x8024A003, "The old version of the Automatic Updates client was disabled."
0x8024A004, "Automatic Updates was unable to process incoming requests because it was paused."
0x8024A005, "No unmanaged service is registered with AU."
0x8024A006, "The default service registered with AU changed during the search."
0x8024AFFF, "An Automatic Updates error not covered by another WU_E_AU * code."
0x80242000, "A request for a remote update handler could not be completed because no remote process is available."
0x80242001, "A request for a remote update handler could not be completed because the handler is local only."
0x80242002, "A request for an update handler could not be completed because the handler could not be recognized."
0x80242003, "A remote update handler could not be created because one already exists."
0x80242004, "A request for the handler to install uninstall) an update could not be completed because the update does not support install uninstall)."
0x80242005, "An operation did not complete because the wrong handler was specified."
0x80242006, "A handler operation could not be completed because the update contains invalid metadata."
0x80242007, "An operation could not be completed because the installer exceeded the time limit."
0x80242008, "An operation being done by the update handler was cancelled."
0x80242009, "An operation could not be completed because the handler-specific metadata is invalid."
0x8024200A, "A request to the handler to install an update could not be completed because the update requires user input."
0x8024200B, "The installer failed to install uninstall) one or more updates."
0x8024200C, "The update handler should download self-contained content rather than delta-compressed content for the update."
0x8024200D, "The update handler did not install the update because it needs to be downloaded again."
0x8024200E, "The update handler failed to send notification of the status of the install uninstall) operation."
0x8024200F, "The file names contained in the update metadata and in the update package are inconsistent."
0x80242010, "The update handler failed to fall back to the self-contained content."
0x80242011, "The update handler has exceeded the maximum number of download requests."
0x80242012, "The update handler has received an unexpected response from CBS."
0x80242013, "The update metadata contains an invalid CBS package identifier."
0x80242014, "The post-reboot operation for the update is still in progress."
0x80242015, "The result of the post-reboot operation for the update could not be determined."
0x80242016, "The state of the update after its post-reboot operation has completed is unexpected."
0x80242017, "The OS servicing stack must be updated before this update is downloaded or installed."
0x80242018, "A callback installer called back with an error."
0x80242019, "The custom installer signature did not match the signature required by the update."
0x8024201A, "The installer does not support the installation configuration."
0x8024201B, "The targeted session for isntall is invalid."
0x80242FFF, "An update handler error not covered by another WU_E_UH_* code."
0x80246001, "A download manager operation could not be completed because the requested file does not have a URL."
0x80246002, "A download manager operation could not be completed because the file digest was not recognized."
0x80246003, "A download manager operation could not be completed because the file metadata requested an unrecognized hash algorithm."
0x80246004, "An operation could not be completed because a download request is required from the download handler."
0x80246005, "A download manager operation could not be completed because the network connection was unavailable."
0x80246006, "A download manager operation could not be completed because the version of Background Intelligent Transfer Service BITS) is incompatible."
0x80246007, "The update has not been downloaded."
0x80246008, "A download manager operation failed because the download manager was unable to connect the Background Intelligent Transfer Service BITS)."
0x80246009, "A download manager operation failed because there was an unspecified Background Intelligent Transfer Service BITS) transfer error."
0x8024600A, "A download must be restarted because the location of the source of the download has changed."
0x8024600B, "A download must be restarted because the update content changed in a new revision."
0x80246FFF, "There was a download manager error not covered by another WU_E_DM_* error code."
0x8024D001, "Windows Update Agent could not be updated because an INF file contains invalid information."
0x8024D002, "Windows Update Agent could not be updated because the wuident.cab file contains invalid information."
0x8024D003, "Windows Update Agent could not be updated because of an internal error that caused setup initialization to be performed twice."
0x8024D004, "Windows Update Agent could not be updated because setup initialization never completed successfully."
0x8024D005, "Windows Update Agent could not be updated because the versions specified in the INF do not match the actual source file versions."
0x8024D006, "Windows Update Agent could not be updated because a WUA file on the target system is newer than the corresponding source file."
0x8024D007, "Windows Update Agent could not be updated because regsvr32.exe returned an error."
0x8024D008, "An update to the Windows Update Agent was skipped because previous attempts to update have failed."
0x8024D009, "An update to the Windows Update Agent was skipped due to a directive in the wuident.cab file."
0x8024D00A, "Windows Update Agent could not be updated because the current system configuration is not supported."
0x8024D00B, "Windows Update Agent could not be updated because the system is configured to block the update."
0x8024D00C, "Windows Update Agent could not be updated because a restart of the system is required."
0x8024D00D, "Windows Update Agent setup is already running."
0x8024D00E, "Windows Update Agent setup package requires a reboot to complete installation."
0x8024D00F, "Windows Update Agent could not be updated because the setup handler failed during execution."
0x8024D010, "Windows Update Agent could not be updated because the registry contains invalid information."
0x8024D011, "Windows Update Agent must be updated before search can continue."
0x8024D012, "Windows Update Agent must be updated before search can continue.  An administrator is required to perform the operation."
0x8024D013, "Windows Update Agent could not be updated because the server does not contain update information for this version."
0x8024DFFF, "Windows Update Agent could not be updated because of an error not covered by another WU_E_SETUP_* error code."
0x8024E001, "An expression evaluator operation could not be completed because an expression was unrecognized."
0x8024E002, "An expression evaluator operation could not be completed because an expression was invalid."
0x8024E003, "An expression evaluator operation could not be completed because an expression contains an incorrect number of metadata nodes."
0x8024E004, "An expression evaluator operation could not be completed because the version of the serialized expression data is invalid."
0x8024E005, "The expression evaluator could not be initialized."
0x8024E006, "An expression evaluator operation could not be completed because there was an invalid attribute."
0x8024E007, "An expression evaluator operation could not be completed because the cluster state of the computer could not be determined."
0x8024EFFF, "There was an expression evaluator error not covered by another WU_E_EE_* error code."
0x80243001, "The results of download and installation could not be read from the registry due to an unrecognized data format version."
0x80243002, "The results of download and installation could not be read from the registry due to an invalid data format."
0x80243003, "The results of download and installation are not available; the operation may have failed to start."
0x80243004, "A failure occurred when trying to create an icon in the taskbar notification area."
0x80243FFD, "Unable to show UI when in non-UI mode; WU client UI modules may not be installed."
0x80243FFE, "Unsupported version of WU client UI exported functions."
0x80243FFF, "There was a user interface error not covered by another WU_E_AUCLIENT_* error code."
0x8024F001, "The event cache file was defective."
0x8024F002, "The XML in the event namespace descriptor could not be parsed."
0x8024F003, "The XML in the event namespace descriptor could not be parsed."
0x8024F004, "The server rejected an event because the server was too busy."
0x8024F005, "The specified callback cookie is not found."
0x8024FFFF, "There was a reporter error not covered by another error code."
0x80247001, "An operation could not be completed because the scan package was invalid."
0x80247002, "An operation could not be completed because the scan package requires a greater version of the Windows Update Agent."
0x80247FFF, "Search using the scan package failed."
'@ | ConvertFrom-Csv

<#
WUSA.exe / CbsCore.dll
https://github.com/seven-mile/CallCbsCore/blob/master/CbsUtil.cpp
https://github.com/insystemsco/scripts/blob/master/run-patch-scan.vbs

Ghidra -> CbsCore.dll
char * FUN_180030fb0(int param_1)

{
  if (param_1 < -0x7ff0f7f0) {
    if (param_1 == -0x7ff0f7f1) {
      return "CBS_E_MANIFEST_VALIDATION_DUPLICATE_ELEMENT";
    }
    if (param_1 < -0x7ff0fdd5) {
      if (param_1 == -0x7ff0fdd6) {
        return "SPAPI_E_INVALID_INF_LOGCONFIG";
      }
      if (param_1 < -0x7ff0fdee) {
        if (param_1 == -0x7ff0fdef) {
          return "SPAPI_E_NO_DEVICE_SELECTED";
        }
        if (param_1 < -0x7ff0fdfb) {
          if (param_1 == -0x7ff0fdfc) {
            return "SPAPI_E_KEY_DOES_NOT_EXIST";
          }
          if (param_1 < -0x7ff0fefd) {
            if (param_1 == -0x7ff0fefe) {
              return "SPAPI_E_LINE_NOT_FOUND";
            }
            if (param_1 == -0x7ff10000) {
              return "SPAPI_E_EXPECTED_SECTION_NAME";
            }
            if (param_1 == -0x7ff0ffff) {
              return "SPAPI_E_BAD_SECTION_NAME_LINE";
            }
            if (param_1 == -0x7ff0fffe) {
              return "SPAPI_E_SECTION_NAME_TOO_LONG";
            }
            if (param_1 == -0x7ff0fffd) {
              return "SPAPI_E_GENERAL_SYNTAX";
            }
            if (param_1 == -0x7ff0ff00) {
              return "SPAPI_E_WRONG_INF_STYLE";
            }
            if (param_1 == -0x7ff0feff) {
              return "SPAPI_E_SECTION_NOT_FOUND";
            }
          }
          else {
            if (param_1 == -0x7ff0fefd) {
              return "SPAPI_E_NO_BACKUP";
            }
            if (param_1 == -0x7ff0fe00) {
              return "SPAPI_E_NO_ASSOCIATED_CLASS";
            }
            if (param_1 == -0x7ff0fdff) {
              return "SPAPI_E_CLASS_MISMATCH";
            }
            if (param_1 == -0x7ff0fdfe) {
              return "SPAPI_E_DUPLICATE_FOUND";
            }
            if (param_1 == -0x7ff0fdfd) {
              return "SPAPI_E_NO_DRIVER_SELECTED";
            }
          }
        }
        else {
          switch(param_1) {
          case -0x7ff0fdfb:
            return "SPAPI_E_INVALID_DEVINST_NAME";
          case -0x7ff0fdfa:
            return "SPAPI_E_INVALID_CLASS";
          case -0x7ff0fdf9:
            return "SPAPI_E_DEVINST_ALREADY_EXISTS";
          case -0x7ff0fdf8:
            return "SPAPI_E_DEVINFO_NOT_REGISTERED";
          case -0x7ff0fdf7:
            return "SPAPI_E_INVALID_REG_PROPERTY";
          case -0x7ff0fdf6:
            return "SPAPI_E_NO_INF";
          case -0x7ff0fdf5:
            return "SPAPI_E_NO_SUCH_DEVINST";
          case -0x7ff0fdf4:
            return "SPAPI_E_CANT_LOAD_CLASS_ICON";
          case -0x7ff0fdf3:
            return "SPAPI_E_INVALID_CLASS_INSTALLER";
          case -0x7ff0fdf2:
            return "SPAPI_E_DI_DO_DEFAULT";
          case -0x7ff0fdf1:
            return "SPAPI_E_DI_NOFILECOPY";
          case -0x7ff0fdf0:
            return "SPAPI_E_INVALID_HWPROFILE";
          }
        }
      }
      else {
        switch(param_1) {
        case -0x7ff0fdee:
          return "SPAPI_E_DEVINFO_LIST_LOCKED";
        case -0x7ff0fded:
          return "SPAPI_E_DEVINFO_DATA_LOCKED";
        case -0x7ff0fdec:
          return "SPAPI_E_DI_BAD_PATH";
        case -0x7ff0fdeb:
          return "SPAPI_E_NO_CLASSINSTALL_PARAMS";
        case -0x7ff0fdea:
          return "SPAPI_E_FILEQUEUE_LOCKED";
        case -0x7ff0fde9:
          return "SPAPI_E_BAD_SERVICE_INSTALLSECT";
        case -0x7ff0fde8:
          return "SPAPI_E_NO_CLASS_DRIVER_LIST";
        case -0x7ff0fde7:
          return "SPAPI_E_NO_ASSOCIATED_SERVICE";
        case -0x7ff0fde6:
          return "SPAPI_E_NO_DEFAULT_DEVICE_INTERFACE";
        case -0x7ff0fde5:
          return "SPAPI_E_DEVICE_INTERFACE_ACTIVE";
        case -0x7ff0fde4:
          return "SPAPI_E_DEVICE_INTERFACE_REMOVED";
        case -0x7ff0fde3:
          return "SPAPI_E_BAD_INTERFACE_INSTALLSECT";
        case -0x7ff0fde2:
          return "SPAPI_E_NO_SUCH_INTERFACE_CLASS";
        case -0x7ff0fde1:
          return "SPAPI_E_INVALID_REFERENCE_STRING";
        case -0x7ff0fde0:
          return "SPAPI_E_INVALID_MACHINENAME";
        case -0x7ff0fddf:
          return "SPAPI_E_REMOTE_COMM_FAILURE";
        case -0x7ff0fdde:
          return "SPAPI_E_MACHINE_UNAVAILABLE";
        case -0x7ff0fddd:
          return "SPAPI_E_NO_CONFIGMGR_SERVICES";
        case -0x7ff0fddc:
          return "SPAPI_E_INVALID_PROPPAGE_PROVIDER";
        case -0x7ff0fddb:
          return "SPAPI_E_NO_SUCH_DEVICE_INTERFACE";
        case -0x7ff0fdda:
          return "SPAPI_E_DI_POSTPROCESSING_REQUIRED";
        case -0x7ff0fdd9:
          return "SPAPI_E_INVALID_COINSTALLER";
        case -0x7ff0fdd8:
          return "SPAPI_E_NO_COMPAT_DRIVERS";
        case -0x7ff0fdd7:
          return "SPAPI_E_NO_DEVICE_ICON";
        }
      }
    }
    else if (param_1 < -0x7ff0fcff) {
      if (param_1 == -0x7ff0fd00) {
        return "SPAPI_E_UNRECOVERABLE_STACK_OVERFLOW";
      }
      switch(param_1) {
      case -0x7ff0fdd5:
        return "SPAPI_E_DI_DONT_INSTALL";
      case -0x7ff0fdd4:
        return "SPAPI_E_INVALID_FILTER_DRIVER";
      case -0x7ff0fdd3:
        return "SPAPI_E_NON_WINDOWS_NT_DRIVER";
      case -0x7ff0fdd2:
        return "SPAPI_E_NON_WINDOWS_DRIVER";
      case -0x7ff0fdd1:
        return "SPAPI_E_NO_CATALOG_FOR_OEM_INF";
      case -0x7ff0fdd0:
        return "SPAPI_E_DEVINSTALL_QUEUE_NONNATIVE";
      case -0x7ff0fdcf:
        return "SPAPI_E_NOT_DISABLEABLE";
      case -0x7ff0fdce:
        return "SPAPI_E_CANT_REMOVE_DEVINST";
      case -0x7ff0fdcd:
        return "SPAPI_E_INVALID_TARGET";
      case -0x7ff0fdcc:
        return "SPAPI_E_DRIVER_NONNATIVE";
      case -0x7ff0fdcb:
        return "SPAPI_E_IN_WOW64";
      case -0x7ff0fdca:
        return "SPAPI_E_SET_SYSTEM_RESTORE_POINT";
      case -0x7ff0fdc9:
        return "SPAPI_E_INCORRECTLY_COPIED_INF";
      case -0x7ff0fdc8:
        return "SPAPI_E_SCE_DISABLED";
      case -0x7ff0fdc7:
        return "SPAPI_E_UNKNOWN_EXCEPTION";
      case -0x7ff0fdc6:
        return "SPAPI_E_PNP_REGISTRY_ERROR";
      case -0x7ff0fdc5:
        return "SPAPI_E_REMOTE_REQUEST_UNSUPPORTED";
      case -0x7ff0fdc4:
        return "SPAPI_E_NOT_AN_INSTALLED_OEM_INF";
      case -0x7ff0fdc3:
        return "SPAPI_E_INF_IN_USE_BY_DEVICES";
      case -0x7ff0fdc2:
        return "SPAPI_E_DI_FUNCTION_OBSOLETE";
      case -0x7ff0fdc1:
        return "SPAPI_E_NO_AUTHENTICODE_CATALOG";
      case -0x7ff0fdc0:
        return "SPAPI_E_AUTHENTICODE_DISALLOWED";
      case -0x7ff0fdbf:
        return "SPAPI_E_AUTHENTICODE_TRUSTED_PUBLISHER";
      case -0x7ff0fdbe:
        return "SPAPI_E_AUTHENTICODE_TRUST_NOT_ESTABLISHED";
      case -0x7ff0fdbd:
        return "SPAPI_E_AUTHENTICODE_PUBLISHER_NOT_TRUSTED";
      case -0x7ff0fdbc:
        return "SPAPI_E_SIGNATURE_OSATTRIBUTE_MISMATCH";
      case -0x7ff0fdbb:
        return "SPAPI_E_ONLY_VALIDATE_VIA_AUTHENTICODE";
      case -0x7ff0fdba:
        return "SPAPI_E_DEVICE_INSTALLER_NOT_READY";
      case -0x7ff0fdb9:
        return "SPAPI_E_DRIVER_STORE_ADD_FAILED";
      case -0x7ff0fdb8:
        return "SPAPI_E_DEVICE_INSTALL_BLOCKED";
      case -0x7ff0fdb7:
        return "SPAPI_E_DRIVER_INSTALL_BLOCKED";
      case -0x7ff0fdb6:
        return "SPAPI_E_WRONG_INF_TYPE";
      case -0x7ff0fdb5:
        return "SPAPI_E_FILE_HASH_NOT_IN_CATALOG";
      case -0x7ff0fdb4:
        return "SPAPI_E_DRIVER_STORE_DELETE_FAILED";
      }
    }
    else {
      switch(param_1) {
      case -0x7ff0f800:
        return "CBS_E_INTERNAL_ERROR";
      case -0x7ff0f7ff:
        return "CBS_E_NOT_INITIALIZED";
      case -0x7ff0f7fe:
        return "CBS_E_ALREADY_INITIALIZED";
      case -0x7ff0f7fd:
        return "CBS_E_INVALID_PARAMETER";
      case -0x7ff0f7fc:
        return "CBS_E_OPEN_FAILED";
      case -0x7ff0f7fb:
        return "CBS_E_INVALID_PACKAGE";
      case -0x7ff0f7fa:
        return "CBS_E_PENDING";
      case -0x7ff0f7f9:
        return "CBS_E_NOT_INSTALLABLE";
      case -0x7ff0f7f8:
        return "CBS_E_IMAGE_NOT_ACCESSIBLE";
      case -0x7ff0f7f7:
        return "CBS_E_ARRAY_ELEMENT_MISSING";
      case -0x7ff0f7f6:
        return "CBS_E_REESTABLISH_SESSION";
      case -0x7ff0f7f5:
        return "CBS_E_PROPERTY_NOT_AVAILABLE";
      case -0x7ff0f7f4:
        return "CBS_E_UNKNOWN_UPDATE";
      case -0x7ff0f7f3:
        return "CBS_E_MANIFEST_INVALID_ITEM";
      case -0x7ff0f7f2:
        return "CBS_E_MANIFEST_VALIDATION_DUPLICATE_ATTRIBUTES";
      }
    }
  }
  else if (param_1 < -0x7ff0f7b0) {
    if (param_1 == -0x7ff0f7b1) {
      return "CBS_E_RESOLVE_FAILED";
    }
    switch(param_1) {
    case -0x7ff0f7f0:
      return "CBS_E_MANIFEST_VALIDATION_MISSING_REQUIRED_ATTRIBUTES";
    case -0x7ff0f7ef:
      return "CBS_E_MANIFEST_VALIDATION_MISSING_REQUIRED_ELEMENTS";
    case -0x7ff0f7ee:
      return "CBS_E_MANIFEST_VALIDATION_UPDATES_PARENT_MISSING";
    case -0x7ff0f7ed:
      return "CBS_E_INVALID_INSTALL_STATE";
    case -0x7ff0f7ec:
      return "CBS_E_INVALID_CONFIG_VALUE";
    case -0x7ff0f7eb:
      return "CBS_E_INVALID_CARDINALITY";
    case -0x7ff0f7ea:
      return "CBS_E_DPX_JOB_STATE_SAVED";
    case -0x7ff0f7e9:
      return "CBS_E_PACKAGE_DELETED";
    case -0x7ff0f7e8:
      return "CBS_E_IDENTITY_MISMATCH";
    case -0x7ff0f7e7:
      return "CBS_E_DUPLICATE_UPDATENAME";
    case -0x7ff0f7e6:
      return "CBS_E_INVALID_DRIVER_OPERATION_KEY";
    case -0x7ff0f7e5:
      return "CBS_E_UNEXPECTED_PROCESSOR_ARCHITECTURE";
    case -0x7ff0f7e4:
      return "CBS_E_EXCESSIVE_EVALUATION";
    case -0x7ff0f7e3:
      return "CBS_E_CYCLE_EVALUATION";
    case -0x7ff0f7e2:
      return "CBS_E_NOT_APPLICABLE ";
    case -0x7ff0f7e1:
      return "CBS_E_SOURCE_MISSING";
    case -0x7ff0f7e0:
      return "CBS_E_CANCEL";
    case -0x7ff0f7df:
      return "CBS_E_ABORT";
    case -0x7ff0f7de:
      return "CBS_E_ILLEGAL_COMPONENT_UPDATE";
    case -0x7ff0f7dd:
      return "CBS_E_NEW_SERVICING_STACK_REQUIRED";
    case -0x7ff0f7dc:
      return "CBS_E_SOURCE_NOT_IN_LIST";
    case -0x7ff0f7db:
      return "CBS_E_CANNOT_UNINSTALL";
    case -0x7ff0f7da:
      return "CBS_E_PENDING_VICTIM";
    case -0x7ff0f7d9:
      return "CBS_E_STACK_SHUTDOWN_REQUIRED";
    case -0x7ff0f7d8:
      return "CBS_E_INSUFFICIENT_DISK_SPACE";
    case -0x7ff0f7d7:
      return "CBS_E_AC_POWER_REQUIRED";
    case -0x7ff0f7d6:
      return "CBS_E_STACK_UPDATE_FAILED_REBOOT_REQUIRED";
    case -0x7ff0f7d5:
      return "CBS_E_SQM_REPORT_IGNORED_AI_FAILURES_ON_TRANSACTION_RESOLVE";
    case -0x7ff0f7d4:
      return "CBS_E_DEPENDENT_FAILURE";
    case -0x7ff0f7d3:
      return "CBS_E_PAC_INITIAL_FAILURE";
    case -0x7ff0f7d2:
      return "CBS_E_NOT_ALLOWED_OFFLINE";
    case -0x7ff0f7d1:
      return "CBS_E_EXCLUSIVE_WOULD_MERGE";
    case -0x7ff0f7d0:
      return "CBS_E_IMAGE_UNSERVICEABLE";
    case -0x7ff0f7cf:
      return "CBS_E_STORE_CORRUPTION";
    case -0x7ff0f7ce:
      return "CBS_E_STORE_TOO_MUCH_CORRUPTION";
    case -0x7ff0f7cd:
      return "CBS_S_STACK_RESTART_REQUIRED";
    case -0x7ff0f7c0:
      return "CBS_E_SESSION_CORRUPT";
    case -0x7ff0f7bf:
      return "CBS_E_SESSION_INTERRUPTED";
    case -0x7ff0f7be:
      return "CBS_E_SESSION_FINALIZED";
    case -0x7ff0f7bd:
      return "CBS_E_SESSION_READONLY";
    }
  }
  else if (param_1 < -0x7ff0f66f) {
    if (param_1 == -0x7ff0f670) {
      return "PSFX_E_UNSUPPORTED_COMPRESSION_SWITCH";
    }
    switch(param_1) {
    case -0x7ff0f700:
      return "CBS_E_XML_PARSER_FAILURE";
    case -0x7ff0f6ff:
      return "CBS_E_MANIFEST_VALIDATION_MULTIPLE_UPDATE_COMPONENT_ON_SAME_FAMILY_NOT_ALLOWED";
    case -0x7ff0f6fe:
      return "CBS_E_BUSY";
    case -0x7ff0f6fd:
      return "CBS_E_INVALID_RECALL";
    case -0x7ff0f6fc:
      return "CBS_E_MORE_THAN_ONE_ACTIVE_EDITION";
    case -0x7ff0f6fb:
      return "CBS_E_NO_ACTIVE_EDITION";
    case -0x7ff0f6fa:
      return "CBS_E_DOWNLOAD_FAILURE";
    case -0x7ff0f6f9:
      return "CBS_E_GROUPPOLICY_DISALLOWED";
    case -0x7ff0f6f8:
      return "CBS_E_METERED_NETWORK";
    case -0x7ff0f6f7:
      return "CBS_E_PUBLIC_OBJECT_LEAK";
    case -0x7ff0f6f5:
      return "CBS_E_REPAIR_PACKAGE_CORRUPT";
    case -0x7ff0f6f4:
      return "CBS_E_COMPONENT_NOT_INSTALLED_BY_CBS";
    case -0x7ff0f6f3:
      return "CBS_E_MISSING_PACKAGE_MAPPING_INDEX";
    case -0x7ff0f6f2:
      return "CBS_E_EMPTY_PACKAGE_MAPPING_INDEX";
    case -0x7ff0f6f1:
      return "CBS_E_WINDOWS_UPDATE_SEARCH_FAILURE";
    case -0x7ff0f6f0:
      return "CBS_E_WINDOWS_AUTOMATIC_UPDATE_SETTING_DISALLOWED";
    case -0x7ff0f6e0:
      return "CBS_E_HANG_DETECTED";
    case -0x7ff0f6df:
      return "CBS_E_PRIMITIVES_FAILED";
    case -0x7ff0f6de:
      return "CBS_E_INSTALLERS_FAILED";
    case -0x7ff0f6dd:
      return "CBS_E_SAFEMODE_ENTERED";
    case -0x7ff0f6dc:
      return "CBS_E_SESSIONS_LEAKED";
    case -0x7ff0f6db:
      return "CBS_E_INVALID_EXECUTESTATE";
    case -0x7ff0f6c0:
      return "CBS_E_WUSUS_MAPPING_UNAVAILABLE";
    case -0x7ff0f6bf:
      return "CBS_E_WU_MAPPING_UNAVAILABLE";
    case -0x7ff0f6be:
      return "CBS_E_WUSUS_BYPASS_MAPPING_UNAVAILABLE";
    case -0x7ff0f6bd:
      return "CBS_E_WUSUS_MISSING_PACKAGE_MAPPING_INDEX";
    case -0x7ff0f6bc:
      return "CBS_E_WU_MISSING_PACKAGE_MAPPING_INDEX";
    case -0x7ff0f6bb:
      return "CBS_E_WUSUS_BYPASS_MISSING_PACKAGE_MAPPING_INDEX";
    case -0x7ff0f6ba:
      return "CBS_E_SOURCE_MISSING_FROM_WUSUS_CAB";
    case -0x7ff0f6b9:
      return "CBS_E_SOURCE_MISSING_FROM_WUSUS_EXPRESS";
    case -0x7ff0f6b8:
      return "CBS_E_SOURCE_MISSING_FROM_WU_CAB";
    case -0x7ff0f6b7:
      return "CBS_E_SOURCE_MISSING_FROM_WU_EXPRESS";
    case -0x7ff0f6b6:
      return "CBS_E_SOURCE_MISSING_FROM_WUSUS_BYPASS_CAB";
    case -0x7ff0f6b5:
      return "CBS_E_SOURCE_MISSING_FROM_WUSUS_BYPASS_EXPRESS";
    case -0x7ff0f6b4:
      return "CBS_E_3RD_PARTY_MAPPING_UNAVAILABLE";
    case -0x7ff0f6b3:
      return "CBS_E_3RD_PARTY_MISSING_PACKAGE_MAPPING_INDEX";
    case -0x7ff0f6b2:
      return "CBS_E_SOURCE_MISSING_FROM_3RD_PARTY_CAB";
    case -0x7ff0f6b1:
      return "CBS_E_SOURCE_MISSING_FROM_3RD_PARTY_EXPRESS";
    case -0x7ff0f6b0:
      return "CBS_E_INVALID_WINDOWS_UPDATE_COUNT";
    case -0x7ff0f6af:
      return "CBS_E_INVALID_NO_PRODUCT_REGISTERED";
    case -0x7ff0f6ae:
      return "CBS_E_INVALID_ACTION_LIST_PACKAGE_COUNT";
    case -0x7ff0f6ad:
      return "CBS_E_INVALID_ACTION_LIST_INSTALL_REASON";
    case -0x7ff0f6ac:
      return "CBS_E_INVALID_WINDOWS_UPDATE_COUNT_WSUS";
    case -0x7ff0f6ab:
      return "CBS_E_INVALID_PACKAGE_REQUEST_ON_MULTILINGUAL_FOD";
    case -0x7ff0f680:
      return "PSFX_E_DELTA_NOT_SUPPORTED_FOR_COMPONENT";
    case -0x7ff0f67f:
      return "PSFX_E_REVERSE_AND_FORWARD_DELTAS_MISSING";
    case -0x7ff0f67e:
      return "PSFX_E_MATCHING_COMPONENT_NOT_FOUND";
    case -0x7ff0f67d:
      return "PSFX_E_MATCHING_COMPONENT_DIRECTORY_MISSING";
    case -0x7ff0f67c:
      return "PSFX_E_MATCHING_BINARY_MISSING";
    case -0x7ff0f67b:
      return "PSFX_E_APPLY_REVERSE_DELTA_FAILED";
    case -0x7ff0f67a:
      return "PSFX_E_APPLY_FORWARD_DELTA_FAILED";
    case -0x7ff0f679:
      return "PSFX_E_NULL_DELTA_HYDRATION_FAILED";
    case -0x7ff0f678:
      return "PSFX_E_INVALID_DELTA_COMBINATION";
    case -0x7ff0f677:
      return "PSFX_E_REVERSE_DELTA_MISSING";
    }
  }
  else {
    if (param_1 == -0x7ff0f000) {
      return "SPAPI_E_ERROR_NOT_INSTALLED";
    }
    if (param_1 == 0xf0801) {
      return "CBS_S_BUSY";
    }
    if (param_1 == 0xf0802) {
      return "CBS_S_ALREADY_EXISTS";
    }
    if (param_1 == 0xf0803) {
      return "CBS_S_STACK_SHUTDOWN_REQUIRED";
    }
  }
  return "Unknown Error";
}
#>
$Global:CBS_ERR_TABLE = @'
ERROR, MESSEGE
0xf0801,    The Component-Based Servicing system is currently busy and cannot process the request right now.
0xf0802,    The item or component you are trying to create or add already exists in the system.
0xf0803,    The servicing stack needs to be shut down and restarted to complete the operation.
0xF0804,    The servicing stack restart is required to complete the operation.
0x800F0991, The requested operation could not be completed due to a component store corruption or a missing manifest file.
-0x7ff0f7f1,Manifest validation failed: a duplicate element was found.
-0x7ff0fdd6,Invalid INF log configuration.
-0x7ff0fdef,No device was selected for this operation.
-0x7ff0fdfc,The specified key does not exist.
-0x7ff0fefe,The line requested was not found.
-0x7ff10000,An expected section name is missing or invalid.
-0x7ff0ffff,The section name line is malformed or incorrect.
-0x7ff0fffe,The section name provided is too long.
-0x7ff0fffd,A general syntax error was detected.
-0x7ff0ff00,The INF file has an incorrect style or format.
-0x7ff0feff,The specified section was not found.
-0x7ff0fefd,No backup copy is available.
-0x7ff0fe00,No associated class was found for this operation.
-0x7ff0fdff,There is a mismatch in the specified class.
-0x7ff0fdfe,A duplicate item was found.
-0x7ff0fdfd,No driver was selected.
-0x7ff0fdfb,The device instance name is invalid.
-0x7ff0fdfa,The specified class is invalid.
-0x7ff0fdf9,A device instance with this name already exists.
-0x7ff0fdf8,The device information set is not registered.
-0x7ff0fdf7,The registry property is invalid.
-0x7ff0fdf6,No INF file was found.
-0x7ff0fdf5,The specified device instance does not exist.
-0x7ff0fdf4,Cannot load the class icon.
-0x7ff0fdf3,The class installer is invalid.
-0x7ff0fdf2,Proceed with the default action.
-0x7ff0fdf1,No file copy operation was performed.
-0x7ff0fdf0,The hardware profile is invalid.
-0x7ff0fdee,The device information list is locked.
-0x7ff0fded,The device information data is locked.
-0x7ff0fdec,The specified path is invalid.
-0x7ff0fdeb,No class installation parameters are available.
-0x7ff0fdea,The file queue is locked.
-0x7ff0fde9,The service installation section is malformed.
-0x7ff0fde8,No class driver list is available.
-0x7ff0fde7,No associated service was found.
-0x7ff0fde6,No default device interface exists.
-0x7ff0fde5,The device interface is currently active.
-0x7ff0fde4,The device interface has been removed.
-0x7ff0fde3,The interface installation section is malformed.
-0x7ff0fde2,The specified interface class does not exist.
-0x7ff0fde1,The reference string is invalid.
-0x7ff0fde0,The machine name is invalid.
-0x7ff0fddf,Communication with the remote machine failed.
-0x7ff0fdde,The machine is unavailable for remote operations.
-0x7ff0fddd,Configuration Manager services are not available.
-0x7ff0fddc,The property page provider is invalid.
-0x7ff0fddb,The specified device interface does not exist.
-0x7ff0fdda,Post-processing is required to complete the operation.
-0x7ff0fdd9,The co-installer is invalid.
-0x7ff0fdd8,No compatible drivers were found.
-0x7ff0fdd7,No device icon is available.
-0x7ff0fd00,An unrecoverable stack overflow occurred.
-0x7ff0fdd5,Do not install the device.
-0x7ff0fdd4,The filter driver is invalid.
-0x7ff0fdd3,This is not a Windows NT driver.
-0x7ff0fdd2,This is not a Windows driver.
-0x7ff0fdd1,No catalog file was found for the OEM INF.
-0x7ff0fdd0,The device installation queue contains non-native items.
-0x7ff0fdcf,The component cannot be disabled.
-0x7ff0fdce,Cannot remove the device instance.
-0x7ff0fdcd,The target specified is invalid.
-0x7ff0fdcc,The driver is not native to this system.
-0x7ff0fdcb,Operation is running in WOW64 (32-bit on 64-bit).
-0x7ff0fdca,A system restore point needs to be set.
-0x7ff0fdc9,The INF file was incorrectly copied.
-0x7ff0fdc8,Security Configuration Engine (SCE) is disabled.
-0x7ff0fdc7,An unknown exception occurred.
-0x7ff0fdc6,A Plug and Play (PNP) registry error occurred.
-0x7ff0fdc5,The remote request is not supported.
-0x7ff0fdc4,The specified OEM INF is not installed.
-0x7ff0fdc3,The INF file is currently in use by other devices.
-0x7ff0fdc2,This device installation function is obsolete.
-0x7ff0fdc1,No Authenticode catalog was found.
-0x7ff0fdc0,Authenticode signature is disallowed.
-0x7ff0fdbf,Authenticode signature from a trusted publisher.
-0x7ff0fdbe,Authenticode trust could not be established.
-0x7ff0fdbd,The Authenticode publisher is not trusted.
-0x7ff0fdbc,The signature's OS attribute does not match.
-0x7ff0fdbb,Validation must be performed via Authenticode only.
-0x7ff0fdba,The device installer is not ready.
-0x7ff0fdb9,Failed to add to the driver store.
-0x7ff0fdb8,Device installation is blocked.
-0x7ff0fdb7,Driver installation is blocked.
-0x7ff0fdb6,The INF file type is incorrect.
-0x7ff0fdb5,The file hash is not found in the catalog.
-0x7ff0fdb4,Failed to delete from the driver store.
-0x7ff0f800,An internal Component-Based Servicing (CBS) error occurred.
-0x7ff0f7ff,The Component-Based Servicing (CBS) system is not initialized.
-0x7ff0f7fe,The Component-Based Servicing (CBS) system is already initialized.
-0x7ff0f7fd,An invalid parameter was provided to CBS.
-0x7ff0f7fc,CBS failed to open a required resource.
-0x7ff0f7fb,The package is invalid or corrupt.
-0x7ff0f7fa,The CBS operation is pending.
-0x7ff0f7f9,The component or package cannot be installed.
-0x7ff0f7f8,The image cannot be accessed.
-0x7ff0f7f7,A required element in the array is missing.
-0x7ff0f7f6,The session needs to be reestablished.
-0x7ff0f7f5,The requested property is not available.
-0x7ff0f7f4,An unknown update was encountered.
-0x7ff0f7f3,The manifest contains an invalid item.
-0x7ff0f7f2,Manifest validation failed: duplicate attributes were found.
-0x7ff0f7b1,The Component-Based Servicing system failed to resolve the requested operation or component.
-0x7ff0f7f0,Manifest validation failed: required attributes are missing.
-0x7ff0f7ef,Manifest validation failed: required elements are missing.
-0x7ff0f7ee,Manifest validation failed: the update's parent is missing.
-0x7ff0f7ed,The installation state is invalid.
-0x7ff0f7ec,The configuration value is invalid.
-0x7ff0f7eb,The cardinality value is invalid.
-0x7ff0f7ea,The DPX job state has been saved.
-0x7ff0f7e9,The package has been deleted.
-0x7ff0f7e8,An identity mismatch was detected.
-0x7ff0f7e7,A duplicate update name was found.
-0x7ff0f7e6,The driver operation key is invalid.
-0x7ff0f7e5,An unexpected processor architecture was encountered.
-0x7ff0f7e4,Excessive evaluation was detected.
-0x7ff0f7e3,A cycle was detected during evaluation.
-0x7ff0f7e2,The operation is not applicable.
-0x7ff0f7e1,A required source is missing.
-0x7ff0f7e0,The operation was cancelled.
-0x7ff0f7df,The operation was aborted.
-0x7ff0f7de,An illegal component update was attempted.
-0x7ff0f7dd,A new servicing stack is required.
-0x7ff0f7dc,The source was not found in the list.
-0x7ff0f7db,The component cannot be uninstalled.
-0x7ff0f7da,A pending victim state was detected.
-0x7ff0f7d9,The servicing stack needs to be shut down.
-0x7ff0f7d8,There is insufficient disk space available.
-0x7ff0f7d7,AC power is required for this operation.
-0x7ff0f7d6,The servicing stack update failed; a reboot is required.
-0x7ff0f7d5,SQM report ignored AI failures on transaction resolve.
-0x7ff0f7d4,A dependent failure occurred.
-0x7ff0f7d3,PAC initialization failed.
-0x7ff0f7d2,The operation is not allowed offline.
-0x7ff0f7d1,An exclusive operation would cause a merge conflict.
-0x7ff0f7d0,The image is unserviceable.
-0x7ff0f7cf,Store corruption was detected.
-0x7ff0f7ce,Too much corruption was found in the store.
-0x7ff0f7cd,A servicing stack restart is required (status).
-0x7ff0f7c0,The session is corrupt.
-0x7ff0f7bf,The session was interrupted.
-0x7ff0f7be,The session has been finalized.
-0x7ff0f7bd,The session is read-only.
-0x7ff0f670,Unsupported compression switch in PSFX.
-0x7ff0f700,The XML parser encountered a failure.
-0x7ff0f6ff,Manifest validation failed: multiple update components on the same family are not allowed.
-0x7ff0f6fe,The Component-Based Servicing system is currently busy.
-0x7ff0f6fd,The recall operation attempted is invalid.
-0x7ff0f6fc,More than one active edition exists, which is not allowed.
-0x7ff0f6fb,No active edition is available.
-0x7ff0f6fa,Failure occurred while downloading the package or component.
-0x7ff0f6f9,This operation is disallowed by Group Policy.
-0x7ff0f6f8,Operation failed because the network connection is metered, restricting data usage.
-0x7ff0f6f7,A public object leak was detected, indicating a potential resource management issue.
-0x7ff0f6f5,The repair package is corrupt and cannot be used.
-0x7ff0f6f4,The component was not installed by CBS and cannot be serviced by it.
-0x7ff0f6f3,Missing package mapping index; the system cannot locate the package mapping.
-0x7ff0f6f2,The package mapping index is empty, causing lookup failures.
-0x7ff0f6f1,Windows Update search failed to find the required updates.
-0x7ff0f6f0,The automatic Windows Update setting is disallowed by policy or configuration.
-0x7ff0f6e0,A failure to respond was detected while processing the operation.
-0x7ff0f6df,Primitive operations failed during servicing.
-0x7ff0f6de,Installer operations failed to complete successfully.
-0x7ff0f6dd,The system has entered safe mode, restricting certain operations.
-0x7ff0f6dc,Sessions have leaked, indicating resource management issues.
-0x7ff0f6db,An invalid execution state was encountered.
-0x7ff0f6c0,WSUS (Windows Server Update Services) mapping is unavailable.
-0x7ff0f6bf,Windows Update mapping is unavailable.
-0x7ff0f6be,WSUS bypass mapping is unavailable.
-0x7ff0f6bd,Missing package mapping index in WSUS.
-0x7ff0f6bc,Missing package mapping index in Windows Update.
-0x7ff0f6bb,Missing package mapping index in WSUS bypass.
-0x7ff0f6ba,Source is missing from the WSUS CAB file.
-0x7ff0f6b9,Source is missing from the WSUS Express package.
-0x7ff0f6b8,Source is missing from the Windows Update CAB file.
-0x7ff0f6b7,Source is missing from the Windows Update Express package.
-0x7ff0f6b6,Source is missing from the WSUS bypass CAB file.
-0x7ff0f6b5,Source is missing from the WSUS bypass Express package.
-0x7ff0f6b4,Third-party mapping is unavailable.
-0x7ff0f6b3,Missing package mapping index for third-party components.
-0x7ff0f6b2,Source is missing from the third-party CAB file.
-0x7ff0f6b1,Source is missing from the third-party Express package.
-0x7ff0f6b0,An invalid count of Windows updates was specified.
-0x7ff0f6af,No registered product found; invalid state.
-0x7ff0f6ae,Invalid count in the action list package.
-0x7ff0f6ad,An invalid reason was specified for action list installation.
-0x7ff0f6ac,Invalid Windows Update count for WSUS.
-0x7ff0f6ab,Invalid package request on multilingual Features on Demand (FOD).
-0x7ff0f680,Delta updates are not supported for this component.
-0x7ff0f67f,Reverse and forward delta files are missing.
-0x7ff0f67e,The matching component was not found.
-0x7ff0f67d,The matching component directory is missing.
-0x7ff0f67c,The matching binary file is missing.
-0x7ff0f67b,Failed to apply the reverse delta update.
-0x7ff0f67a,Failed to apply the forward delta update.
-0x7ff0f679,Null delta hydration process failed.
-0x7ff0f678,An invalid combination of delta updates was specified.
-0x7ff0f677,The reverse delta update is missing.
-0x7ff0f000,The error indicates that the component is not installed.
'@ | ConvertFrom-Csv

<#
So technically, error messege, stored in couple location's.

winhttp.dll > Windows Update common errors and mitigation
* https://learn.microsoft.com/en-us/troubleshoot/windows-client/installing-updates-features-roles/common-windows-update-errors

netmsg.dll > Network Management Error Codes
* https://learn.microsoft.com/en-us/windows/win32/netmgmt/network-management-error-codes

Kernel32.dll ,api-ms-win-core-synch-l1-2-0.dll > Win32 Error Codes & HRESULT Values
* https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/18d8fbe8-a967-4f1c-ae50-99ca8e491d2d
* https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/705fb797-2175-4a90-b5a3-3918024b10b8

NTDLL.dll > NTSTATUS Values
* https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/87fba13e-bf06-450e-83b1-9241dc81e781

SLC.dll > Windows Activation Error
* https://howtoedge.com/windows-activation-error-codes-and-solutions/

qmgr.dll > BITS Return Values
* https://learn.microsoft.com/en-us/windows/win32/bits/bits-return-values
* https://gitlab.winehq.org/wine/wine/-/blob/master/include/bitsmsg.h?ref_type=heads

Also, it include in header files too, 
Check Microsoft Error Lookup Tool as example

there is also other place it could save, possibly,
C:\Windows\Logs\CBS\CBS.log, Also include CBS ERROR
could not find DLL Source, or Header file ? .. well

# --------------------------------------------------------------

Microsoft Error Lookup Tool
https://www.microsoft.com/en-us/download/details.aspx?id=100432

SLMGR.vbs, Source code ->
"On a computer running Microsoft Windows non-core edition, run 'slui.exe 0x2a 0x%ERRCODE%' to display the error text."

RtlInitUnicodeStringEx
https://www.geoffchappell.com/studies/windows/km/ntoskrnl/api/rtl/string/initunicodestringex.htm

# --------------------------------------------------------------

slui 0x2a 0xC004F014
using API-SPY -> Debug info
#29122    9:06:51.521 AM    2    KERNELBASE.dll    RtlInitUnicodeStringEx ( 0x00000016fcd7f810, "SLC.dll" )    STATUS_SUCCESS        0.0000001
#29123    9:06:51.521 AM    2    KERNELBASE.dll    RtlDosApplyFileIsolationRedirection_Ustr ( TRUE, 0x00000016fcd7f810, 0x00007ffa653e2138, 0x00000016fcd7f638, 0x00000016fcd7f620, 0x00000016fcd7f5f8, NULL, NULL, NULL )    STATUS_SXS_KEY_NOT_FOUND    0xc0150008 = The requested lookup key was not found in any active activation context.     0.0000008
#29124    9:06:51.521 AM    2    KERNELBASE.dll    RtlFindMessage ( 0x00007ffa639e0000, 11, 1024, 3221549076, 0x00000016fcd7f6f8 )    STATUS_SUCCESS        0.0000484

17-win32 error {api-ms-win-core-synch-l1-2-0}

349	10:57:41.933 PM	1	Kernel32.dll	LoadLibraryEx ( "api-ms-win-core-synch-l1-2-0", NULL, 2048 )	0x00007ffd5dea0000		0.0000012
28480	10:57:45.636 PM	2	KERNELBASE.dll	RtlFindMessage ( 0x00007ffd5dea0000, 11, 1024, 17, 0x00000062b88ff458 )	STATUS_SUCCESS		0.0001046

0x...-SL ERROR ---> ntdll.dll { not from slc.dll, from }

329	10:52:10.339 PM	1	ntdll.dll	DllMain ( 0x00007ffd5c610000, DLL_PROCESS_ATTACH, 0x0000008aba5af380 )	TRUE		0.0000185
28461	10:52:14.324 PM	2	KERNELBASE.dll	RtlFindMessage ( 0x00007ffd5c610000, 11, 1024, 3221549076, 0x0000008aba9ff4d8 )	STATUS_SUCCESS		0.0000233

# --------------------------------------------------------------

2.1.1 HRESULT Values
https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/705fb797-2175-4a90-b5a3-3918024b10b8

2.2 Win32 Error Codes
https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/18d8fbe8-a967-4f1c-ae50-99ca8e491d2d

2.3.1 NTSTATUS Values
https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/596a1078-e883-4972-9bbc-49e60bebca55

Network Management Error Codes
https://learn.microsoft.com/en-us/windows/win32/netmgmt/network-management-error-codes

https://github.com/SystemRage/py-kms/blob/master/py-kms/pykms_Misc.py
http://joshpoley.blogspot.com/2011/09/hresults-user-0x004.html  (slerror.h)

Troubleshoot Windows activation error codes
https://learn.microsoft.com/en-us/troubleshoot/windows-server/licensing-and-activation/troubleshoot-activation-error-codes

https://github.com/SystemRage/py-kms/blob/master/py-kms/pykms_Misc.py
http://joshpoley.blogspot.com/2011/09/hresults-user-0x004.html  (slerror.h)

Windows Activation Error Codes and Solutions on Windows 11/10
https://howtoedge.com/windows-activation-error-codes-and-solutions/

Additional Resources for Windows Server Update Services
https://learn.microsoft.com/de-de/security-updates/windowsupdateservices/18127498

Windows Update error codes by component
https://learn.microsoft.com/en-us/windows/deployment/update/windows-update-error-reference?source=recommendations

Windows Update common errors and mitigation
https://learn.microsoft.com/en-us/troubleshoot/windows-client/installing-updates-features-roles/common-windows-update-errors

BITS Return Values
https://learn.microsoft.com/en-us/windows/win32/bits/bits-return-values

bitsmsg.h
https://gitlab.winehq.org/wine/wine/-/blob/master/include/bitsmsg.h?ref_type=heads

# --------------------------------------------------------------
              Alternative Source
# --------------------------------------------------------------

RCodes.ini
https://forums.mydigitallife.net/threads/multi-oem-retail-project-mrp-mk3.71555/

MicrosoftOfficeAssistant
https://github.com/audricd/MicrosoftOfficeAssistant/blob/master/scripts/roiscan.vbs

--> Alternative CBS / WU, hard codec error database, No dll found, yet! <--

https://github.com/larsch/wunow/blob/master/wunow/WUError.cs
https://github.com/microsoft/IIS.Setup/blob/main/iisca/lib/wuerror.h

#>
<#
Clear-Host
write-host

write-host ------------------------------------------------------------
Write-Host "             NUMBER FORMAT TEST                           " -ForegroundColor Red
write-host ------------------------------------------------------------

Write-Host
Write-Host 'Test win32 error' -ForegroundColor Red
$unsignedNumber = 17
$hexRepresentation = "0x{0:X}" -f $unsignedNumber
$unsignedLong = [long]$unsignedNumber
$overflowedNumber = $unsignedLong - 0x100000000

# Construct the UInt32 HRESULT, And, 
# Format it as a hexadecimal string
$hResultUInt32 = 0x80000000 -bor 0x00070000 -bor $unsignedNumber
$hexNegativeString = "0x{0:X}" -f $hResultUInt32

Write-Host
Write-Warning 'unsignedNumber'
Parse-ErrorMessage -log -MessageId $unsignedNumber
Write-Warning 'overflowedNumber'
Parse-ErrorMessage -log -MessageId $overflowedNumber
Write-Warning 'hexRepresentation'
Parse-ErrorMessage -log -MessageId $hexRepresentation
Write-Warning 'hexNegativeString'
Parse-ErrorMessage -log -MessageId $hexNegativeString

write-host
write-host ------------------------------------------------------------
write-host

Write-Host
Write-Host 'Test SL error' -ForegroundColor Red
$unsignedNumber = 3221549172
$hexRepresentation = "0x{0:X}" -f $unsignedNumber
$unsignedLong = [long]$unsignedNumber
$overflowedNumber = $unsignedLong - 0x100000000

Write-Host
Write-Warning 'unsignedNumber'
Parse-ErrorMessage -log -MessageId $unsignedNumber
Write-Warning 'overflowedNumber'
Parse-ErrorMessage -log -MessageId $overflowedNumber
Write-Warning 'hexRepresentation'
Parse-ErrorMessage -log -MessageId $hexRepresentation

write-host
write-host ------------------------------------------------------------
Write-Host "             DIFFRENT ERROR TEST                          " -ForegroundColor Red
write-host ------------------------------------------------------------
write-host
write-host

Write-Host "Locate -> Activation error" -ForegroundColor Green
Parse-ErrorMessage -log -MessageId 0xC004B007

Write-Host "Locate -> NT STATUS error" -ForegroundColor Green
Parse-ErrorMessage -log -MessageId 0x40000016

Write-Host "Locate -> WIN32 error" -ForegroundColor Green
write-host "0x Positive"
Parse-ErrorMessage -log -MessageId 0x0000215B
write-host "0x Negative"
Parse-ErrorMessage -log -MessageId 0x8007232B
write-host "[Negative] 0x"
Parse-ErrorMessage -log -MessageId -0x7FF8DCD5

Write-Host "Locate -> HRESULT error" -ForegroundColor Green
Parse-ErrorMessage -log -MessageId 0x00030203

Write-Host "Locate -> WU error" -ForegroundColor Green
Parse-ErrorMessage -log -MessageId 0x8024000E

Write-Host "Locate -> network error" -ForegroundColor Green
Parse-ErrorMessage -log -MessageId 0x853 

Write-Host
Write-Host "Locate -> Bits error" -ForegroundColor Green
Parse-ErrorMessage -log -MessageId 0x80200010

Write-Host
Write-Host "Locate -> CBS error" -ForegroundColor Green
Parse-ErrorMessage -log -MessageId 0x800f0831 
Write-Host
write-host "0x Negative"
Parse-ErrorMessage -log -MessageId 0x800f081e 
Write-Host
write-host "[Negative] 0x"
Parse-ErrorMessage -MessageId -0x7ff10000L -Log

write-host
write-host
write-host ------------------------------------------------------------
Write-Host "             OCTA + LEADING TEST                          " -ForegroundColor Red
write-host ------------------------------------------------------------
write-host
write-host

Write-Host "** Octa Test" -ForegroundColor Green
Parse-ErrorMessage -log -MessageId 0225D

Write-Host "** Leading Test" -ForegroundColor Green
Parse-ErrorMessage -log -MessageId 225AASASS

write-host
write-host
write-host ------------------------------------------------------------
Write-Host "             FLAG TEST                                     " -ForegroundColor Red
write-host ------------------------------------------------------------
write-host
write-host

Write-Warning "Testing HRESULT"
Parse-ErrorMessage -log -MessageId 0x00030203 -Flags HRESULT

Write-Host
Write-Warning "Testing WIN32"
Parse-ErrorMessage -log -MessageId '0x80070005' -Flags WIN32

Write-Host
Write-Warning "Testing NTSTATUS"
Parse-ErrorMessage -log -MessageId '0xC0000005' -Flags NTSTATUS

Write-Host
Write-Warning "Testing ACTIVATION"
Parse-ErrorMessage -log -MessageId '0xC004F074' -Flags ACTIVATION

Write-Host
Write-Warning "Testing NETWORK"
Parse-ErrorMessage -log -MessageId '0x853' -Flags NETWORK

Write-Host
Write-Warning "Testing NETWORK -> BITS"
Parse-ErrorMessage -log -MessageId '0x8019019B' -Flags BITS

Write-Host
Write-Warning "Testing NETWORK -> Windows HTTP Services"
Parse-ErrorMessage -log -MessageId '0x80072EE7' -Flags HTTP

Write-Host
Write-Warning "Testing CBS"
Parse-ErrorMessage -log -MessageId '0x800F081F' -Flags CBS

Write-Host
Write-Warning "Testing WINDOWS UPDATE"
Parse-ErrorMessage -log -MessageId '0x00240007' -Flags UPDATE

Write-Host
Write-Host
Write-Host

Write-Host "Testing *** ALL CASE" -ForegroundColor Green
Write-Host "Mode: No flags" -ForegroundColor Green
Parse-ErrorMessage -log -MessageId 0x80072EE7
write-Host "Mode: -Flags ALL" -ForegroundColor Green
Parse-ErrorMessage -log -MessageId 0x80072EE7 -Flags ALL
write-Host "Mode: -Flags ([ErrorMessageType]::ALL)" -ForegroundColor Green
Parse-ErrorMessage -log -MessageId 0x80072EE7 -Flags ([ErrorMessageType]::ALL)

Write-Host "Testing *** BOR CASE" -ForegroundColor Green
Write-Host "WIN32 -bor HRESULT -bor NTSTATUS" -ForegroundColor Green
Parse-ErrorMessage -log -MessageId 0x00030206 -Flags ([ErrorMessageType]::WIN32 -bor [ErrorMessageType]::NTSTATUS -bor [ErrorMessageType]::HRESULT)
#>
enum ErrorMessageType {
    WIN32      = 1
    NTSTATUS   = 2
    ACTIVATION = 4
    NETWORK    = 8
    CBS        = 16
    BITS       = 32
    HTTP       = 64
    UPDATE     = 128
    HRESULT    = 256
    ALL        = 511
}
function Parse-MessageId {
    param (
        [string] $MessageId
    )
    if ($MessageId -match '^(-?0x[0-9a-fA-F]+).*$') { 
        $MessageId = $matches[1]
        $isNegative = $MessageId.StartsWith('-')
        if ($isNegative) {
            $MessageId = $MessageId.TrimStart('-')
        }
    
        try {

            $hexVal = [Convert]::ToUInt32($MessageId, 16)
            if ($isNegative) { 
                $hexVal = Parse-MessageId -MessageId (-1 * $hexVal)
            }
            $isWin32Err = ($hexVal -band 0x80000000) -ne 0 -and (($hexVal -shr 16) -band 0x0FFF) -eq 7
            if ($isWin32Err){
                return ($hexVal -band 0x0000FFFF)
            }
            else {
                return $hexVal
            }
        }
        catch {
            Write-Warning "Invalid hex value: '$MessageId'. Error: $($_.Exception.Message)"
            return $null
        }
    }
    elseif ($MessageId.StartsWith("0")) {

        if ($MessageId -eq "0") {
            return 0
        }

        $numericPart = ""
        $foundOctalDigits = $false

        for ($i = 1; $i -lt $MessageId.Length; $i++) {
            $char = $MessageId[$i]
            if ($char -ge '0' -and $char -le '7') {
                $numericPart += $char
                $foundOctalDigits = $true
            } else {
                break
            }
        }
        if ($foundOctalDigits) {
            try {
                $decimalValue = [Convert]::ToInt32($numericPart, 8)
                return $decimalValue
            } catch {
                return $null
            }
        }
        elseif ($MessageId.Length -gt 1) {
            return 0
        }

    }
    else {
        $MessageId = $MessageId -replace '^(?<decimal>-?\d+).*$', '${decimal}'
       
        try {
            $uintVal = [uint32]::Parse($MessageId)
            $isWin32Err = ($uintVal -band 0x80000000) -ne 0 -and (($uintVal -shr 16) -band 0x0FFF) -eq 7

            if ($isWin32Err){
                return ($uintVal -band 0x0000FFFF)
            }

            return $uintVal
        }
        catch {
            try {
                $longVal = [long]::Parse($MessageId)
                if ($longVal -lt 0) {
                    $wrappedVal = $longVal + 0x100000000L
                    if ($wrappedVal -ge 0 -and $wrappedVal -le [uint32]::MaxValue) {
                        $unsignedVal = [uint32]$wrappedVal
                        return $unsignedVal
                    } else {
                        return $null
                    }
                }
                elseif ($longVal -gt [uint32]::MaxValue) {
                    return $null
                }
                else {
                    return [uint32]$longVal
                }
            }
            catch {
                if ($MessageId -match '^\d+') {
                    return [long]$matches[0]
                }
                return $null
            }
        }
    }
}
function Parse-ErrorMessage {
    param (
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string] $MessageId,

        [Parameter(Mandatory = $false)]
        [ErrorMessageType]$Flags = [ErrorMessageType]::ALL,

        [Parameter(Mandatory=$false)]
        [switch]$Log,

        [Parameter(Mandatory=$false)]
        [switch]$LastWin32Error,

        [Parameter(Mandatory=$false)]
        [switch]$LastNTStatus
    )

    if ($MessageId -and ($LastWin32Error -or $LastNTStatus)) {
        throw "Choice MessageId -or Win32Error\LastNTStatus Only.!"
    }
    
    if ($LastWin32Error -or $LastNTStatus) {
        if ($LastWin32Error -and $LastNTStatus) {
            throw "Choice Win32Error -or LastNTStatus Only.!"
        }

        if($LastWin32Error) {
            # Last win32 error
            $MessageId = [marshal]::ReadInt32((NtCurrentTeb), 0x68)
            $Flags = [ErrorMessageType]::WIN32
        } elseif ($LastNTStatus) {
            # Last NTSTATUS error
            $MessageId = [marshal]::ReadInt32((NtCurrentTeb), 0x1250)
            $Flags = [ErrorMessageType]::NTSTATUS
        }

        
    }
    
    if ($MessageId -eq "0" -or $MessageId -eq "0x0") {
        return "Status OK"
    }

    $MessegeValue = Parse-MessageId -MessageId $MessageId
    if ($null -eq $MessegeValue) {
        Write-Warning "Invalid message ID: $MessageId"
        continue
    }

    $apiList = @()
    if (($Flags -eq $null) -or -not ($Flags -is [ErrorMessageType])) {
        $Flags = [ErrorMessageType]::ALL
    }

    # If ALL is set, expand it to all meaningful flags
    if (($Flags -band [ErrorMessageType]::ALL) -eq [ErrorMessageType]::ALL) {
        $Flags = [ErrorMessageType]::WIN32      -bor `
                 [ErrorMessageType]::NTSTATUS   -bor `
                 [ErrorMessageType]::ACTIVATION -bor `
                 [ErrorMessageType]::NETWORK    -bor `
                 [ErrorMessageType]::CBS        -bor `
                 [ErrorMessageType]::BITS       -bor `
                 [ErrorMessageType]::HTTP       -bor `
                 [ErrorMessageType]::UPDATE     -bor `
                 [ErrorMessageType]::HRESULT
    }
    foreach ($Flag in [Enum]::GetValues([ErrorMessageType]) | Where-Object { $_ -ne [ErrorMessageType]::ALL }) {
            $isValueExist = ($Flags -band $flag) -eq $flag
            if ($isValueExist) {
                switch ($flag) {
                    ([ErrorMessageType]::HTTP)         { $apiList += "winhttp.dll" }
                    ([ErrorMessageType]::BITS)         { $apiList += "qmgr.dll" }
                    ([ErrorMessageType]::NETWORK)      { $apiList += "netmsg.dll" }
                    ([ErrorMessageType]::WIN32)        { $apiList += "KernelBase.dll","Kernel32.dll"}  #,"api-ms-win-core-synch-l1-2-0.dll" }
                    ([ErrorMessageType]::HRESULT)      { $apiList += "KernelBase.dll","Kernel32.dll"}  #,"api-ms-win-core-synch-l1-2-0.dll" }
                    ([ErrorMessageType]::NTSTATUS)     { $apiList += "ntdll.dll" }
                    ([ErrorMessageType]::ACTIVATION)   { $apiList += "slc.dll", "sppc.dll"}
                }
            }
    }
    $apiList = $apiList | Sort-Object -Unique

    # Define booleans for the flags of interest
    $IsAll    = (($Flags -band [ErrorMessageType]::ALL)    -eq [ErrorMessageType]::ALL)
    $IsCBS    = (($Flags -band [ErrorMessageType]::CBS)    -eq [ErrorMessageType]::CBS)
    $IsUpdate = (($Flags -band [ErrorMessageType]::UPDATE) -eq [ErrorMessageType]::UPDATE)

    if ($IsAll -or $IsUpdate) {
        if ($Log) {
            Write-Warning "Trying Look In WU ERROR_TABLE"
        }
        $messege = $Global:WU_ERR_TABLE | Where-Object { @(Parse-MessageId $_.ERROR) -eq $MessegeValue } | Select-Object -ExpandProperty MESSEGE
        if ($messege) {
            return $messege
        }
        if ($IsUpdate -and ($Flags -eq [ErrorMessageType]::UPDATE)) {
            return
        }
    }

    if ($IsAll -or $IsCBS) {
        if ($Log) {
            Write-Warning "Trying Look In CBS ERROR_TABLE"
        }
        $messege = $Global:CBS_ERR_TABLE | Where-Object { @(Parse-MessageId $_.ERROR) -eq $MessegeValue } | Select-Object -ExpandProperty MESSEGE
        if ($messege) {
            return $messege
        }
        if ($IsCBS -and ($Flags -eq [ErrorMessageType]::CBS)) {
            return
        }
    }
    foreach ($dll in $apiList) {
        
        if (-not $baseMap.ContainsKey($dll)) {
            if ($Log) {
                Write-Warning "$dll failed to load"
            }
            continue
        }

        $hModule = $baseMap[$dll]
        if ($Log) {
            Write-Warning "$dll loaded at base address: $hModule"
        }

        # Find message resource
        $messageEntryPtr = [IntPtr]::Zero
        $result = $Global:ntdll::RtlFindMessage(
            $hModule, 11, 1024, $MessegeValue, [ref]$messageEntryPtr)
        if ($result -ne 0) {
            # Free Handle returned from LoadLibraryExA
            # $null = $Global:kernel32::FreeLibrary($hModule)
            continue
        }

        # Extract MESSAGE_RESOURCE_ENTRY fields
        $length = [Marshal]::ReadInt16($messageEntryPtr, 0)
        $flags  = [Marshal]::ReadInt16($messageEntryPtr, 2)
        $textPtr = [IntPtr]::Add($messageEntryPtr, 4)

        try {
            # Decode string (Unicode or ANSI)
            if (($flags -band 0x0001) -ne 0) {
                $charCount = ($length - 4) / 2
                return [Marshal]::PtrToStringUni($textPtr, $charCount)
            } else {
                $charCount = $length - 4
                return [Marshal]::PtrToStringAnsi($textPtr, $charCount)
            }
        }
        catch {
        }
        finally {
            # Free Handle returned from LoadLibraryExA
            # $null = $Global:kernel32::FreeLibrary($hModule)
        }
    }
}

<#
ntstatus.h
https://www.cnblogs.com/george-cw/p/12613148.html
https://codemachine.com/downloads/win71/ntstatus.h
https://github.com/danmar/clang-headers/blob/master/ntstatus.h
https://home.cs.colorado.edu/~main/cs1300-old/include/ddk/ntstatus.h
https://searchfox.org/mozilla-central/source/third_party/rust/winapi/src/shared/ntstatus.rs
2.3 NTSTATUS
https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/87fba13e-bf06-450e-83b1-9241dc81e781

//
//  Values are 32 bit values layed out as follows:
//
//   3 3 2 2 2 2 2 2 2 2 2 2 1 1 1 1 1 1 1 1 1 1
//   1 0 9 8 7 6 5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0
//  +---+-+-+-----------------------+-------------------------------+
//  |Sev|C|R|     Facility          |               Code            |
//  +---+-+-+-----------------------+-------------------------------+
//
//  where
//
//      Sev - is the severity code
//
//          00 - Success
//          01 - Informational
//          10 - Warning
//          11 - Error
//
//      C - is the Customer code flag
//
//      R - is a reserved bit
//
//      Facility - is the facility code
//
//      Code - is the facility's status code
//

winerror.h
https://doxygen.reactos.org/d4/ded/winerror_8h_source.html

//
//  HRESULTs are 32 bit values layed out as follows:
//
//   3 3 2 2 2 2 2 2 2 2 2 2 1 1 1 1 1 1 1 1 1 1
//   1 0 9 8 7 6 5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0
//  +-+-+-+-+-+---------------------+-------------------------------+
//  |S|R|C|N|r|    Facility         |               Code            |
//  +-+-+-+-+-+---------------------+-------------------------------+
//
//  where
//
//      S - Severity - indicates success/fail
//
//          0 - Success
//          1 - Fail (COERROR)
//
//      R - reserved portion of the facility code, corresponds to NT's
//              second severity bit.
//
//      C - reserved portion of the facility code, corresponds to NT's
//              C field.
//
//      N - reserved portion of the facility code. Used to indicate a
//              mapped NT status value.
//
//      r - reserved portion of the facility code. Reserved for internal
//              use. Used to indicate HRESULT values that are not status
//              values, but are instead message ids for display strings.
//
//      Facility - is the facility code
//
//      Code - is the facility's status code
//

Facility Codes
5 Appendix A: Product Behavior
https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/1714a7aa-8e53-4076-8f8d-75073b780a41
2.1 HRESULT
https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/0642cb2f-2075-4469-918c-4441e69c548a

Error Codes: Win32 vs. HRESULT vs. NTSTATUS
https://jpassing.com/2007/08/20/error-codes-win32-vs-hresult-vs-ntstatus/
HRESULT_FACILITY macro (winerror.h)
https://learn.microsoft.com/en-us/windows/win32/api/winerror/nf-winerror-hresult_facility
HRESULT_FROM_NT macro (winerror.h)
https://learn.microsoft.com/en-us/windows/win32/api/winerror/nf-winerror-hresult_from_nt
HRESULT_FROM_WIN32 macro (winerror.h)
https://learn.microsoft.com/en-us/windows/win32/api/winerror/nf-winerror-hresult_from_win32
2.1.2 HRESULT From WIN32 Error Code Macro
https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/0c0bcf55-277e-4120-b5dc-f6115fc8dc38

-------------------------------------------------

Clear-Host
Write-Host

Write-Warning "Check ERROR_NOT_SAME_DEVICE WIN32 -> 0x00000011L"
Parse-ErrorFacility -Log $true -HResult 0x00000011L

Write-Warning "Check ERROR_HANDLE_DISK_FULL WIN32 -> 0x00000027L"
Parse-ErrorFacility -Log $true -HResult 0x00000027L

Write-Warning "Check CONVERT10_S_NO_PRESENTATION HRESULTS -> 0x000401C0L"
Parse-ErrorFacility -Log $true -HResult 0x000401C0L

Write-Warning "Check MK_S_ME HRESULTS -> 0x000401E4L"
Parse-ErrorFacility -Log $true -HResult 0x000401E4L

Write-Warning "Check STATUS_SERVICE_NOTIFICATION NTSTATUS -> 0x40000018L"
Parse-ErrorFacility -Log $true -HResult 0x40000018L

Write-Warning "Check STATUS_BAD_STACK NTSTATUS -> 0xC0000028L"
Parse-ErrorFacility -Log $true -HResult 0xC0000028L

Write-Warning "Check STATUS_NDIS_INDICATION_REQUIRED NTSTATUS -> 0x40230001L"
Parse-ErrorFacility -Log $true -HResult 0x40230001L

Write-Warning "Check WU -> 0x00242015"
Parse-ErrorFacility -Log $true -HResult 0x00242015

Write-Warning "Check CBS -> 2148469005 "
Parse-ErrorFacility -Log $true -HResult 2148469005
#>
enum HRESULT_Facility {
    FACILITY_NULL                             = 0x0         # General (no specific source)
    FACILITY_RPC                              = 0x1         # Remote Procedure Call
    FACILITY_DISPATCH                         = 0x2         # COM Dispatch
    FACILITY_STORAGE                          = 0x3         # Storage
    FACILITY_ITF                              = 0x4         # Interface-specific
    FACILITY_WIN32                            = 0x7         # Standard Win32 errors
    FACILITY_WINDOWS                          = 0x8         # Windows system component
    FACILITY_SECURITY                         = 0x9         # Security subsystem
    FACILITY_SSPI                             = 0x9         # Security Support Provider Interface
    FACILITY_CONTROL                          = 0xA         # Control
    FACILITY_CERT                             = 0xB         # Certificate services
    FACILITY_INTERNET                         = 0xC         # Internet-related
    FACILITY_MEDIASERVER                      = 0xD         # Media server
    FACILITY_MSMQ                             = 0xE         # Microsoft Message Queue
    FACILITY_SETUPAPI                         = 0xF         # Setup API
    FACILITY_SCARD                            = 0x10        # Smart card subsystem
    FACILITY_COMPLUS                          = 0x11        # COM+ services
    FACILITY_AAF                              = 0x12        # Advanced Authoring Format
    FACILITY_URT                              = 0x13        # .NET runtime
    FACILITY_ACS                              = 0x14        # Access Control Services
    FACILITY_DPLAY                            = 0x15        # DirectPlay
    FACILITY_UMI                              = 0x16        # UMI (Universal Management Infrastructure)
    FACILITY_SXS                              = 0x17        # Side-by-Side (Assembly)
    FACILITY_WINDOWS_CE                       = 0x18        # Windows CE
    FACILITY_HTTP                             = 0x19        # HTTP services
    FACILITY_USERMODE_COMMONLOG               = 0x1A        # Common Logging
    FACILITY_WER                              = 0x1B        # Windows Error Reporting
    FACILITY_USERMODE_FILTER_MANAGER          = 0x1F        # File system filter manager
    FACILITY_BACKGROUNDCOPY                   = 0x20        # Background Intelligent Transfer Service (BITS)
    FACILITY_CONFIGURATION                    = 0x21        # Configuration
    FACILITY_WIA                              = 0x21        # Windows Image Acquisition
    FACILITY_STATE_MANAGEMENT                 = 0x22        # State management services
    FACILITY_METADIRECTORY                    = 0x23        # Meta-directory services
    FACILITY_WINDOWSUPDATE                    = 0x24        # Windows Update
    FACILITY_DIRECTORYSERVICE                 = 0x25        # Directory services (e.g., Active Directory)
    FACILITY_GRAPHICS                         = 0x26        # Graphics subsystem
    FACILITY_NAP                              = 0x27        # Network Access Protection
    FACILITY_SHELL                            = 0x27        # Windows Shell
    FACILITY_TPM_SERVICES                     = 0x28        # Trusted Platform Module services
    FACILITY_TPM_SOFTWARE                     = 0x29        # TPM software stack
    FACILITY_UI                               = 0x2A        # User Interface
    FACILITY_XAML                             = 0x2B        # XAML parser
    FACILITY_ACTION_QUEUE                     = 0x2C        # Action queue
    FACILITY_PLA                              = 0x30        # Performance Logs and Alerts
    FACILITY_WINDOWS_SETUP                    = 0x30        # Windows Setup
    FACILITY_FVE                              = 0x31        # Full Volume Encryption (BitLocker)
    FACILITY_FWP                              = 0x32        # Windows Filtering Platform
    FACILITY_WINRM                            = 0x33        # Windows Remote Management
    FACILITY_NDIS                             = 0x34        # Network Driver Interface Specification
    FACILITY_USERMODE_HYPERVISOR              = 0x35        # User-mode Hypervisor
    FACILITY_CMI                              = 0x36        # Configuration Management Infrastructure
    FACILITY_USERMODE_VIRTUALIZATION          = 0x37        # User-mode virtualization
    FACILITY_USERMODE_VOLMGR                  = 0x38        # Volume Manager
    FACILITY_BCD                              = 0x39        # Boot Configuration Data
    FACILITY_USERMODE_VHD                     = 0x3A        # Virtual Hard Disk
    FACILITY_SDIAG                            = 0x3C        # System Diagnostics
    FACILITY_WEBSERVICES                      = 0x3D        # Web Services
    FACILITY_WINPE                            = 0x3D        # Windows Preinstallation Environment
    FACILITY_WPN                              = 0x3E        # Windows Push Notification
    FACILITY_WINDOWS_STORE                    = 0x3F        # Windows Store
    FACILITY_INPUT                            = 0x40        # Input subsystem
    FACILITY_EAP                              = 0x42        # Extensible Authentication Protocol
    FACILITY_WINDOWS_DEFENDER                 = 0x50        # Windows Defender
    FACILITY_OPC                              = 0x51        # OPC (Open Packaging Conventions)
    FACILITY_XPS                              = 0x52        # XML Paper Specification
    FACILITY_RAS                              = 0x53        # Remote Access Service
    FACILITY_MBN                              = 0x54        # Mobile Broadband
    FACILITY_POWERSHELL                       = 0x54        # PowerShell
    FACILITY_EAS                              = 0x55        # Exchange ActiveSync
    FACILITY_P2P_INT                          = 0x62        # Peer-to-Peer internal
    FACILITY_P2P                              = 0x63        # Peer-to-Peer
    FACILITY_DAF                              = 0x64        # Device Association Framework
    FACILITY_BLUETOOTH_ATT                    = 0x65        # Bluetooth Attribute Protocol
    FACILITY_AUDIO                            = 0x66        # Audio subsystem
    FACILITY_VISUALCPP                        = 0x6D        # Visual C++ runtime
    FACILITY_SCRIPT                           = 0x70        # Scripting engine
    FACILITY_PARSE                            = 0x71        # Parsing engine
    FACILITY_BLB                              = 0x78        # Backup/Restore infrastructure
    FACILITY_BLB_CLI                          = 0x79        # Backup/Restore client
    FACILITY_WSBAPP                           = 0x7A        # Windows Server Backup Application
    FACILITY_BLBUI                            = 0x80        # Backup UI
    FACILITY_USN                              = 0x81        # Update Sequence Number Journal
    FACILITY_USERMODE_VOLSNAP                 = 0x82        # Volume Snapshot Service
    FACILITY_TIERING                          = 0x83        # Storage Tiering
    FACILITY_WSB_ONLINE                       = 0x85        # Windows Server Backup Online
    FACILITY_ONLINE_ID                        = 0x86        # Windows Live ID
    FACILITY_DLS                              = 0x99        # Downloadable Sound (DLS)
    FACILITY_SOS                              = 0xA0        # SOS debugging
    FACILITY_DEBUGGERS                        = 0xB0        # Debuggers
    FACILITY_USERMODE_SPACES                  = 0xE7        # Storage Spaces (user-mode)
    FACILITY_DMSERVER                         = 0x100       # Digital Media Server
    FACILITY_RESTORE                          = 0x100       # System Restore
    FACILITY_SPP                              = 0x100       # Software Protection Platform
    FACILITY_DEPLOYMENT_SERVICES_SERVER       = 0x101       # Windows Deployment Server
    FACILITY_DEPLOYMENT_SERVICES_IMAGING      = 0x102       # Imaging services
    FACILITY_DEPLOYMENT_SERVICES_MANAGEMENT   = 0x103       # Deployment management
    FACILITY_DEPLOYMENT_SERVICES_UTIL         = 0x104       # Deployment utilities
    FACILITY_DEPLOYMENT_SERVICES_BINLSVC      = 0x105       # BINL service
    FACILITY_DEPLOYMENT_SERVICES_PXE          = 0x107       # PXE boot service
    FACILITY_DEPLOYMENT_SERVICES_TFTP         = 0x108       # Trivial File Transfer Protocol
    FACILITY_DEPLOYMENT_SERVICES_TRANSPORT_MANAGEMENT = 0x110 # Transport management
    FACILITY_DEPLOYMENT_SERVICES_DRIVER_PROVISIONING = 0x116 # Driver provisioning
    FACILITY_DEPLOYMENT_SERVICES_MULTICAST_SERVER = 0x121     # Multicast server
    FACILITY_DEPLOYMENT_SERVICES_MULTICAST_CLIENT = 0x122     # Multicast client
    FACILITY_DEPLOYMENT_SERVICES_CONTENT_PROVIDER = 0x125     # Content provider
    FACILITY_LINGUISTIC_SERVICES              = 0x131       # Linguistic analysis services
    FACILITY_WEB                              = 0x375       # Web Platform
    FACILITY_WEB_SOCKET                       = 0x376       # WebSockets
    FACILITY_AUDIOSTREAMING                   = 0x446       # Audio streaming
    FACILITY_ACCELERATOR                      = 0x600       # Hardware acceleration
    FACILITY_MOBILE                           = 0x701       # Windows Mobile
    FACILITY_WMAAECMA                         = 0x7CC       # Audio echo cancellation
    FACILITY_WEP                              = 0x801       # Windows Enforcement Platform
    FACILITY_SYNCENGINE                       = 0x802       # Sync engine
    FACILITY_DIRECTMUSIC                      = 0x878       # DirectMusic
    FACILITY_DIRECT3D10                       = 0x879       # Direct3D 10
    FACILITY_DXGI                             = 0x87A       # DirectX Graphics Infrastructure
    FACILITY_DXGI_DDI                         = 0x87B       # DXGI Device Driver Interface
    FACILITY_DIRECT3D11                       = 0x87C       # Direct3D 11
    FACILITY_LEAP                             = 0x888       # Leap Motion (or similar input)
    FACILITY_AUDCLNT                          = 0x889       # Audio client
    FACILITY_WINCODEC_DWRITE_DWM              = 0x898       # Imaging, DirectWrite, DWM
    FACILITY_DIRECT2D                         = 0x899       # Direct2D graphics
    FACILITY_DEFRAG                           = 0x900       # Defragmentation
    FACILITY_USERMODE_SDBUS                   = 0x901       # Secure Digital bus (user-mode)
    FACILITY_JSCRIPT                          = 0x902       # JScript engine
    FACILITY_PIDGENX                          = 0xA01       # Product ID Generator (extended)
    FACILITY_UNKNOWN                          = 0xFFF       # Unknown facility
}
enum NTSTATUS_FACILITY {
    FACILITY_DEBUGGER             = 0x1
    FACILITY_RPC_RUNTIME          = 0x2
    FACILITY_RPC_STUBS            = 0x3
    FACILITY_IO_ERROR_CODE        = 0x4
    FACILITY_CODCLASS_ERROR_CODE  = 0x6
    FACILITY_NTWIN32              = 0x7
    FACILITY_NTCERT               = 0x8
    FACILITY_NTSSPI               = 0x9
    FACILITY_TERMINAL_SERVER      = 0xA
    FACILITY_MUI_ERROR_CODE       = 0xB
    FACILITY_USB_ERROR_CODE       = 0x10
    FACILITY_HID_ERROR_CODE       = 0x11
    FACILITY_FIREWIRE_ERROR_CODE  = 0x12
    FACILITY_CLUSTER_ERROR_CODE   = 0x13
    FACILITY_ACPI_ERROR_CODE      = 0x14
    FACILITY_SXS_ERROR_CODE       = 0x15
    FACILITY_TRANSACTION          = 0x19
    FACILITY_COMMONLOG            = 0x1A
    FACILITY_VIDEO                = 0x1B
    FACILITY_FILTER_MANAGER       = 0x1C
    FACILITY_MONITOR              = 0x1D
    FACILITY_GRAPHICS_KERNEL      = 0x1E
    FACILITY_DRIVER_FRAMEWORK     = 0x20
    FACILITY_FVE_ERROR_CODE       = 0x21
    FACILITY_FWP_ERROR_CODE       = 0x22
    FACILITY_NDIS_ERROR_CODE      = 0x23
    FACILITY_TPM                  = 0x29
    FACILITY_RTPM                 = 0x2A
    FACILITY_HYPERVISOR           = 0x35
    FACILITY_IPSEC                = 0x36
    FACILITY_VIRTUALIZATION       = 0x37
    FACILITY_VOLMGR               = 0x38
    FACILITY_BCD_ERROR_CODE       = 0x39
    FACILITY_WIN32K_NTUSER        = 0x3E
    FACILITY_WIN32K_NTGDI         = 0x3F
    FACILITY_RESUME_KEY_FILTER    = 0x40
    FACILITY_RDBSS                = 0x41
    FACILITY_BTH_ATT              = 0x42
    FACILITY_SECUREBOOT           = 0x43
    FACILITY_AUDIO_KERNEL         = 0x44
    FACILITY_VSM                  = 0x45
    FACILITY_VOLSNAP              = 0x50
    FACILITY_SDBUS                = 0x51
    FACILITY_SHARED_VHDX          = 0x5C
    FACILITY_SMB                  = 0x5D
    FACILITY_INTERIX              = 0x99
    FACILITY_SPACES               = 0xE7
    FACILITY_SECURITY_CORE        = 0xE8
    FACILITY_SYSTEM_INTEGRITY     = 0xE9
    FACILITY_LICENSING            = 0xEA
    FACILITY_PLATFORM_MANIFEST    = 0xEB
    FACILITY_APP_EXEC             = 0xEC
    FACILITY_MAXIMUM_VALUE        = 0xED
    FACILITY_UNKNOWN              = 0xFFFF
    FACILITY_NT_BIT               = 0x10000000
}
function Parse-ErrorFacility {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$HResult,

        [Parameter(Mandatory = $false)]
        [bool]$Log = $false
    )

    # Define a helper to check if $ntFacility is valid enum member
    function Is-ValidNTFacility {
        param($facility)
        try {
            [NTSTATUS_FACILITY]$facility | Out-Null
            return $true
        } catch {
            return $false
        }
    }
    # Convert input string to integer HRESULT (hex or decimal)
    $HResultDecimal = [uint32](Parse-MessageId $HResult)
    if ($log) {
        Write-Warning "HResultDecimal is $HResultDecimal"
        Write-Warning ("HResultDecimal (hex): 0x{0:X8}" -f $HResultDecimal)
        Write-Warning ("HResultDecimal (type): {0}" -f $HResultDecimal.GetType().Name)            
    }
    if ($null -eq $HResultDecimal -or $HResultDecimal -eq '') {
        if ($Log) { Write-Warning "Failed to parse HResult input." }
        return [HRESULT_Facility]::UNKNOWN_FACILITY
    }

    # 2. If less than 0x10000, treat as Win32 error and convert to HRESULT
    if ($HResultDecimal -lt 0x10000) {
        if ($Log) { Write-Warning "Input is a Win32 error code. Converting to HRESULT." }
        $HResultDecimal = ($HResultDecimal -band 0xFFFF) -bor 0x80070000
        if ($Log) { Write-Warning ("Converted HRESULT: 0x{0:X8}" -f $HResultDecimal) }
    }

    # 3. Extract facility using official HRESULT_FACILITY macro (bits 16-28, 13 bits)
    $facility13 = ($HResultDecimal -shr 16) -band 0x1FFF
    if ($Log) { Write-Warning ("[13-bit mask] FacilityValue = $facility13") }
    try {
        if ($facility13 -ne 0) { return [HRESULT_Facility]$facility13 }
    } catch {}

    # Fallbacks to smaller masks for legacy compatibility
    $facility12 = ($HResultDecimal -shr 16) -band 0xFFF
    if ($Log) { Write-Warning ("[12-bit mask] FacilityValue = $facility12") }
    try {
        if ($facility12 -ne 0) { return [HRESULT_Facility]$facility12 }
    } catch {}

    $facility11 = ($HResultDecimal -shr 16) -band 0x7FF
    if ($Log) { Write-Warning ("[11-bit mask] FacilityValue = $facility11") }
    try {
        if ($facility11 -ne 0) { return [HRESULT_Facility]$facility11 }
    } catch {}

    # N (1 bit): If set, indicates that the error code is an NTSTATUS value
    # R (1 bit): Reserved. If the N bit is set, this bit is defined by the NTSTATUS numbering space
    # #define HRESULT_FROM_NT(x)      ((HRESULT) ((x) | FACILITY_NT_BIT))
    $Is_N_BIT = ($HResultDecimal -band 0x10000000) -ne 0  # bit 28
    $Is_R_BIT = ($HResultDecimal -band 0x40000000) -ne 0  # bit 30

    if ($log) {
        Write-Warning "Is_N_BIT = $Is_N_BIT"
        Write-Warning "Is_R_BIT = $Is_R_BIT"
    }
    if ($Is_N_BIT -or $Is_R_BIT) {
        
        $Severity = ($HResultDecimal -shr 30) -band 0x3
        $ntFacility = ($HResultDecimal -shr 16) -band 0xFFF  # 12-bit mask
        $SeverityLabel = @('SUCCESS', 'INFORMATIONAL', 'WARNING', 'ERROR')[$Severity]
        if ($Log) { Write-Warning "[NTSTATUS detected with $SeverityLabel severity] FacilityValue = $ntFacility" }

        # Special case: NTSTATUS_FROM_WIN32 (0xC007xxxx)
        if (($HResultDecimal -band 0xFFFF0000) -eq 0xC0070000) {
            $win32Code = $HResultDecimal -band 0xFFFF
            if ($Log) {
                Write-Warning "NTSTATUS_FROM_WIN32 detected. Original Win32 error code: 0x{0:X}" -f $win32Code
                Write-Warning "[Facility = NTWIN32 (7)]"
            }
            return [NTSTATUS_FACILITY]::FACILITY_NTWIN32
        }

        $win32ERR = 0
        $win32ERR = $Global:ntdll::RtlNtStatusToDosError($HResultDecimal)
        if ($log) {
            Write-Warning "RtlNtStatusToDosError return $win32ERR"
        }

        if (($win32ERR -notin (0, 317)) -and -not (Is-ValidNTFacility $ntFacility)) {
            # NTSTATUS facility invalid or unknown, try Win32 facility fallback
            try {
                if ($log) {
                    Write-Warning "Redirct error with win32ERR value"
                }
                return Parse-ErrorFacility -HResult $win32ERR
            }
            catch {
                if ($log) {
                    Write-Warning "Return FACILITY_UNKNOWN"
                }
                return [NTSTATUS_FACILITY]::FACILITY_UNKNOWN
            }
        }
        else {
            # Return the NTSTATUS facility (including facility 0)
            try {
                if ($log) {
                    Write-Warning "Parse ntFacility"
                }
                return [NTSTATUS_FACILITY]$ntFacility
            }
            catch {
                if ($log) {
                    Write-Warning "Return FACILITY_UNKNOWN"
                }
                return [NTSTATUS_FACILITY]::FACILITY_UNKNOWN
            }
        }
    }

    if ($facility13 -eq 0 -and $facility12 -eq 0 -and $facility11 -eq 0) {
        return [HRESULT_Facility]::FACILITY_NULL
    }

    return [HRESULT_Facility]::FACILITY_UNKNOWN
}

# WIN32 API Parts
function Dump-MemoryAddress {
    param (
        [Parameter(Mandatory)]
        [IntPtr] $Pointer, 

        [Parameter(Mandatory)]
        [UInt32] $Length,

        [string] $FileName = "memdump.bin"
    )

    $desktop = [Environment]::GetFolderPath('Desktop')
    $outputPath = Join-Path $desktop $FileName

    try {
        # Allocate managed buffer
        $buffer = New-Object byte[] $Length

        # Perform memory copy
        [Marshal]::Copy(
            $Pointer, # Source pointer
            $buffer,                 # Destination array
            0,                       # Start index
            $Length                  # Number of bytes
        )

        # Write to file
        [System.IO.File]::WriteAllBytes($outputPath, $buffer)

        Write-Host "Memory dumped to: $outputPath"
    } catch {
        Write-Error "Failed to dump memory: $_"
    }
}
function New-IntPtr {
    param(
        [Parameter(Mandatory=$false)]
        [int]$Size,

        [Parameter(Mandatory=$false)]
        [int]$InitialValue = 0,

        [Parameter(Mandatory=$false)]
        [IntPtr]$hHandle,

        [Parameter(Mandatory=$false)]
        [byte[]]$Data,

        [Parameter(Mandatory=$false)]
        [switch]$UsePointerSize,

        [switch]$MakeRefType,
        [switch]$WriteSizeAtZero,
        [switch]$Release
    )

    if ($hHandle -or $Release -or $MakeRefType) {
        if ($Size) {
            throw [System.ArgumentException] "Size option can't go with hHandle, Release or MakeRefType"
        }
        if ($MakeRefType -and $Release) {
            throw [System.ArgumentException] "Cannot specify both MakeRefType and Release"
        }
        if (!$hHandle -or (!$Release -and !$MakeRefType)) {
            throw [System.ArgumentException] "hHandle must be provided with either Release or MakeRefType"
        }
    }

    if ($MakeRefType) {
        $handlePtr = [Marshal]::AllocHGlobal([IntPtr]::Size)
        [Marshal]::WriteIntPtr($handlePtr, $hHandle)
        return $handlePtr
    }

    if ($Release) {
        if ($hHandle -and $hHandle -ne [IntPtr]::Zero) {
            [Marshal]::FreeHGlobal($hHandle)
        }
        return
    }

    if ($Data) {
        $Size = $Data.Length
        $ptr = [Marshal]::AllocHGlobal($Size)
        [Marshal]::Copy($Data, 0, $ptr, $Size)
        return $ptr
    }

    if ($Size -le 0) {
        throw [ArgumentException]::new("Size must be a positive non-zero integer.")
    }
    $ptr = [Marshal]::AllocHGlobal($Size)
    $Global:ntdll::RtlZeroMemory($ptr, [UIntPtr]::new($Size))
    if ($WriteSizeAtZero) {
        if ($UsePointerSize) {
            [Marshal]::WriteIntPtr($ptr, 0, [IntPtr]::new($Size))
        }
        else {
            [Marshal]::WriteInt32($ptr, 0, $Size)
        }
    }
    elseif (($Size -ge 4) -and ($InitialValue -ne 0)) {
        if ($UsePointerSize) {
            [Marshal]::WriteIntPtr($ptr, 0, [IntPtr]::new($InitialValue))
        }
        else {
            [Marshal]::WriteInt32($ptr, 0, $InitialValue)
        }
    }

    return $ptr
}
function IsValid-IntPtr {
    param (
        [Parameter(Mandatory = $false)]
        [Object]$handle
    )

    if ($null -eq $handle) {
        return $false
    }

    if ($handle -is [IntPtr]) {
        return ($handle -ne [IntPtr]::Zero)
    }

    if ($handle -is [UIntPtr]) {
        return ($handle -ne [UIntPtr]::Zero)
    }

    if ($handle -is [ValueType]) {
        $tname = $handle.GetType().Name
        if ($tname -in @('SByte','Byte','Int16','UInt16','Int32','UInt32','Int64','UInt64')) {
            if ([IntPtr]::Size -eq 4) {
                # x86: cast to Int32 first
                $val = [int32]$handle
                return ([IntPtr]$val -ne [IntPtr]::Zero)
            }
            else {
                # x64: cast to Int64
                $val = [int64]$handle
                return ([IntPtr]$val -ne [IntPtr]::Zero)
            }
        }
    }

    return $false
}

function Free-IntPtr {
    param (
        [Parameter(Mandatory=$false)]
        [Object]$handle,

        [ValidateSet(
            "HGlobal", "Handle", "NtHandle",
            "ServiceHandle", "Heap", "STRING",
            "UNICODE_STRING", "BSTR", "VARIANT",
            "Local", "Auto", "Desktop", "WindowStation",
            "License")]
        [string]$Method = "HGlobal"
    )
    $IsValidPointer = IsValid-IntPtr $handle
    if (!$IsValidPointer) {
        return
    }

    try {
        $Module = [AppDomain]::CurrentDomain.GetAssemblies()| ? { $_.ManifestModule.ScopeName -eq "WIN32U" } | select -Last 1
        $WIN32U = $Module.GetTypes()[0]
    }
    catch {
        $Module = [AppDomain]::CurrentDomain.DefineDynamicAssembly("null", 1).DefineDynamicModule("WIN32U", $False).DefineType("null")
        @(
            @('null', 'null', [int], @()), # place holder
            @('NtUserCloseDesktop',       'win32U.dll', [Int], @([IntPtr])),
            @('NtUserCloseWindowStation', 'win32U.dll', [Int], @([IntPtr]))
        ) | % {
            $Module.DefinePInvokeMethod(($_[0]), ($_[1]), 22, 1, [Type]($_[2]), [Type[]]($_[3]), 1, 3).SetImplementationFlags(128) # Def` 128, fail-safe 0 
        }
        $WIN32U = $Module.CreateType()
    }

    [IntPtr]$ptrToFree = $handle
    #Write-Warning "Free $handle -> $Method"

    switch ($Method) {
        "HGlobal" {
            [Marshal]::FreeHGlobal($ptrToFree)
        }
        "Handle" {
            $null = $Global:kernel32::CloseHandle($ptrToFree)
        }
        "NtHandle" {
            $null = $Global:ntdll::NtClose($ptrToFree)
        }
        "ServiceHandle" {
            $null = $Global:advapi32::CloseServiceHandle($ptrToFree)
        }
        "BSTR" {
            $null = [Marshal]::FreeBSTR($ptrToFree)
        }
        "Heap" {
            $null = $Global:ntdll::RtlFreeHeap(
                ((NtCurrentTeb -ProcessHeap)), 0, $ptrToFree)
        }
        "Local" {
            $null = $Global:kernel32::LocalFree($ptrToFree)
        }
        "STRING" {
            $null = Free-NativeString -StringPtr $ptrToFree
        }
        "UNICODE_STRING" {
            $null = Free-NativeString -StringPtr $ptrToFree
        }
        "VARIANT" {
            $null = Free-Variant -variantPtr $ptrToFree
        }
        "Desktop" {
            $null = $WIN32U::NtUserCloseDesktop($ptrToFree)
        }
        "WindowStation" {
            $null = $WIN32U::NtUserCloseWindowStation($ptrToFree)
        }
        "License" {
            $null = $Global:SLC::SLClose($ptrToFree)
        }

        <#
        ## Disabled, use heap instead
        "Process_Parameter" {
           #$global:ntdll::RtlDestroyEnvironment($ptrToFree)
            $global:ntdll::RtlDestroyProcessParameters($ptrToFree)
        }
        #>

        "Auto" {
            # Best effort guess based on pointer value (basic heuristics)
            # Could be expanded if needed
            try {
                [Marshal]::FreeHGlobal($ptrToFree)
            } catch {
                $null = $Global:kernel32::CloseHandle($ptrToFree)
            }
        }
        default {
                Write-Warning "Unknown freeing method specified: $Method. Attempting HGlobal."
                [Marshal]::FreeHGlobal($ptrToFree)
        }
    }
    if ($handle.Value) {
        $handle.Value = 0
    }
    $handle = $null
    $ptrToFree = 0
}

# DLL Loader
function Register-NativeMethods {
    param (
        [Parameter(Mandatory)]
        [Array]$FunctionList,

        # Global defaults
        $NativeCallConv      = [CallingConvention]::Winapi,
        $NativeCharSet       = [CharSet]::Unicode,
        $ImplAttributes      = [MethodImplAttributes]::PreserveSig,
        $TypeAttributes      = [TypeAttributes]::Public -bor [TypeAttributes]::Abstract -bor [TypeAttributes]::Sealed,
        $Attributes          = [MethodAttributes]::Public -bor [MethodAttributes]::Static -bor [MethodAttributes]::PinvokeImpl,
        $CallingConventions  = [CallingConventions]::Standard
    )

    # Dynamic assembly + module
    $asmName = New-Object System.Reflection.AssemblyName "DynamicDllHelperAssembly"
    $asm     = [AppDomain]::CurrentDomain.DefineDynamicAssembly($asmName, [AssemblyBuilderAccess]::Run)
    $mod     = $asm.DefineDynamicModule("DynamicDllHelperModule")
    $tb      = $mod.DefineType("NativeMethods", $TypeAttributes)

    foreach ($func in $FunctionList) {
        # Per-function overrides
        $funcCharSet = if ($func.ContainsKey("CharSet")) { 
            [System.Runtime.InteropServices.CharSet]::$($func.CharSet) 
        } else { 
            $NativeCharSet 
        }

        $funcCallConv = if ($func.ContainsKey("CallConv")) { 
            $func.CallConv 
        } else { 
            $NativeCallConv 
        }

        $tb.DefinePInvokeMethod(
            $func.Name,
            $func.Dll,
            $Attributes,
            $CallingConventions,
            $func.ReturnType,
            $func.Parameters,
            $funcCallConv,
            $funcCharSet
        ).SetImplementationFlags($ImplAttributes)
    }

    return $tb.CreateType()
}
Function Init-CLIPC {

    $functions = @(
        @{ Name = "ClipGetSubscriptionStatus";  Dll = "clipc.dll"; ReturnType = [uint32]; Parameters = [Type[]]@([IntPtr].MakeByRefType(),[IntPtr],[IntPtr],[IntPtr]) }
    )
    return Register-NativeMethods $functions
}
Function Init-SLC {
    
    <#
    .SYNOPSIS

    Should be called from -> Slc.dll
    but work from sppc.dll, osppc.dll
    maybe Slc.dll call sppc.dll -or osppc.dll

    Windows 10 DLL File Information - sppc.dll
    https://windows10dll.nirsoft.net/sppc_dll.html

    List of files that are statically linked to sppc.dll, 
    slc.dll, etc, etc, 
    This means that when one of the above files is loaded, 
    sppc.dll will be loaded too.
    (The opposite of the previous 'Static Linking' section)

    "OSPPC.dll" is a dynamic link library (DLL) file,
    that is a core component of Microsoft Office's Software Protection Platform.
    Essentially, it's involved in the licensing and activation of your Microsoft Office products.
    can be found under windows 7 ~ Vista, For older MSI version

    SLIsGenuineLocal, SLGetLicensingStatusInformation, SLGetWindowsInformation, SLGetWindowsInformationDWORD -> 
    is likely --> ZwQueryLicenseValue with: (Security-SPP-Action-StateData, Security-SPP-LastWindowsActivationHResult, etc)
    So, instead, use Get-ProductPolicy instead, to enum all value's

    >>> SLIsGenuineLocal function (slpublic.h)
    This function checks the **Tampered flag** of the license associated with the specified application. If the license is not valid, 
    or if the Tampered flag of the license is set, the installation is not considered valid. 

    >>> https://www.geoffchappell.com/studies/windows/km/ntoskrnl/api/ex/slmem/queryvalue.htm
    If the license has been **tampered with**, the function fails (returning STATUS_INTERNAL_ERROR). 
    If the licensing cache is corrupt, the function fails (returning STATUS_DATA_ERROR). 
    If there are no licensing descriptors but the kernel thinks it has the licensing descriptors sorted, 
    the function fails (returning STATUS_OJBECT_NAME_NOT_FOUND). 
    If the licensing descriptors are not sorted, they have to be.

    #>
    $functions = @(
        @{ Name = "SLOpen";                       Dll = "sppc.dll"; ReturnType = [int]; Parameters = [Type[]]@([IntPtr].MakeByRefType()) },
        @{ Name = "SLGetLicense";                 Dll = "sppc.dll"; ReturnType = [int]; Parameters = [Type[]]@(
            [IntPtr],                        # hSLC
            [Guid].MakeByRefType(),          # pSkuId
            [UInt32].MakeByRefType(),        # pBufferSize (pointer to UInt32)
            [IntPtr].MakeByRefType()         # pBuffer (pointer to BYTE*)
        )},
        @{ name = 'SLGetProductSkuInformation'; Dll = "sppc.dll"; returnType = [Int32]; parameters = @([IntPtr], [Guid].MakeByRefType(), [String], [UInt32].MakeByRefType(), [UInt32].MakeByRefType(), [IntPtr].MakeByRefType()) },
        @{ name = 'SLGetServiceInformation';    Dll = "sppc.dll"; returnType = [Int32]; parameters = @([IntPtr], [String], [UInt32].MakeByRefType(), [UInt32].MakeByRefType(), [IntPtr].MakeByRefType()) },
        
        @{ Name = "SLClose";                      Dll = "sppc.dll"; ReturnType = [int]; Parameters = [Type[]]@([IntPtr]) },
        @{ Name = "SLGetLicenseInformation";      Dll = "sppc.dll"; ReturnType = [int]; Parameters = [Type[]]@( 
           [IntPtr],                 # hSLC
           [Guid].MakeByRefType(),   # pSLLicenseId
           [string],                 # pwszValueName
           [IntPtr].MakeByRefType(), # peDataType (optional)
           [UInt32].MakeByRefType(), # pcbValue
           [IntPtr].MakeByRefType()  # ppbValue
           )},
        @{ Name = "SLGetPKeyInformation"; Dll = "sppc.dll"; ReturnType = [int]; Parameters = [Type[]]@(
            [IntPtr],                         # hSLC
            [Guid].MakeByRefType(),           # pPKeyId
            [string],                         # pwszValueName
            [IntPtr].MakeByRefType(),         # peDataType
            [UInt32].MakeByRefType(),         # pcbValue
            [IntPtr].MakeByRefType()          # ppbValue
        )},
        @{  Name       = 'SLGetInstalledProductKeyIds'
            Dll        = "sppc.dll"
            ReturnType = [UInt32]
            Parameters = @(
                [IntPtr],                         # HSLC
                [Guid].MakeByRefType(),           # pProductSkuId (nullable)
                [UInt32].MakeByRefType(),         # *pnProductKeyIds
                [IntPtr].MakeByRefType()          # **ppProductKeyIds
            )
        },
        @{  Name       = 'SLGetApplicationInformation'
            Dll        = "sppc.dll"
            ReturnType = [Int32]
            Parameters = @(
                [IntPtr],                  # HSLC hSLC
		        [Guid].MakeByRefType(),    # const SLID* pApplicationId
		        [string],                  # PCWSTR pwszValueName
		        [IntPtr],                  # SLDATATYPE* peDataType (optional)
		        [IntPtr],                  # UINT* pcbValue (output)
		        [IntPtr]                   # PBYTE* ppbValue (output pointer-to-pointer)
            )
        },
        @{
            Name       = 'SLGetGenuineInformation'
            Dll        = "sppc.dll"
            ReturnType = [Int32]  # HRESULT (return type of the function)
            Parameters = @(
                [Guid].MakeByRefType(),         # const SLID* pQueryId
                [string],                       # PCWSTR pwszValueName
                [int].MakeByRefType(),          # SLDATATYPE* peDataType (optional)
                [int].MakeByRefType(),          # UINT* pcbValue (out)
                [IntPtr].MakeByRefType()        # BYTE** ppbValue (out)
            )
            },
            @{
                Name       = 'SLGetSLIDList'
                Dll        = "sppc.dll"
                ReturnType = [Int32]  # HRESULT (return type of the function)
                Parameters = @(
                    [IntPtr],             # hSLC (HSLC handle)
                    [Int32],              # eQueryIdType (SLIDTYPE)
                    [IntPtr],             # null (no query ID passed)
                    [Int32],              # eReturnIdType (SLIDTYPE)
                    [int].MakeByRefType(), 
                    [IntPtr].MakeByRefType()
                )
            },
            @{
                Name       = 'SLUninstallLicense'
                Dll        = "sppc.dll"
                ReturnType = [Int32]  # HRESULT
                Parameters = @(
                    [IntPtr],              # hSLC
                    [Guid].MakeByRefType() # const SLID* pLicenseFileId
                )
            },
            @{
                Name       = 'SLInstallLicense'
                Dll        = "sppc.dll"
                ReturnType = [Int32]  # HRESULT
                Parameters = @(
                    [IntPtr],                # HSLC hSLC
                    [UInt32],                # UINT cbLicenseBlob
                    [IntPtr],                # const BYTE* pbLicenseBlob
                    [Guid].MakeByRefType()   # SLID* pLicenseFileId (output GUID)
                )
            },
            @{
                Name       = 'SLInstallProofOfPurchase'
                Dll        = "sppc.dll"
                ReturnType = [Int32]  # HRESULT
                Parameters = @(
                    [IntPtr],                         # HSLC hSLC
                    [string],                         # pwszPKeyAlgorithm (e.g., "msft:rm/algorithm/pkey/2005")
                    [string],                         # pwszPKeyString (the product key)
                    [IntPtr],                         # cbPKeySpecificData (size of specific data, could be 0)
                    [IntPtr],                         # pbPKeySpecificData (optional additional data, can be NULL)
                    [Guid].MakeByRefType()            # SLID* pPkeyId (output GUID)
                )
            },
            @{
                Name       = 'SLUninstallProofOfPurchase'
                Dll        = "sppc.dll"
                ReturnType = [Int32]  # HRESULT
                Parameters = @(
                    [IntPtr],                         # HSLC hSLC
                    [Guid]                            # pPKeyId (the GUID returned from SLInstallProofOfPurchase)
                )
            },
            @{
                Name       = 'SLFireEvent'
                Dll        = "sppc.dll"
                ReturnType = [Int32]  # HRESULT (return type of the function)
                Parameters = @(
                    [IntPtr],              # hSLC
                    [String],              # pwszEventId (PCWSTR)
                    [Guid].MakeByRefType() # pApplicationId (SLID*)
                )
            },
            @{
                Name       = 'SLReArm'
                Dll        = 'sppc.dll'
                ReturnType = [Int32] # HRESULT
                Parameters = @(
                    [IntPtr],               # hSLC (HSLC handle)
                    [Guid].MakeByRefType(), # pApplicationId (const SLID* - pointer to GUID)
                    [Guid].MakeByRefType(), # pProductSkuId (const SLID* - pointer to GUID, optional)
                    [UInt32]                # dwFlags (DWORD)
                )
            },
            @{
                Name       = 'SLReArmWindows'
                Dll        = 'slc.dll'
                ReturnType = [Int32] # HRESULT
                Parameters = @()
            },
            @{
                Name       = 'SLActivateProduct'
                Dll        = 'sppcext.dll'
                ReturnType = [Int32] # HRESULT
                Parameters = @(
                    [IntPtr],           # hSLC (HSLC handle)
                    [Guid].MakeByRefType(), # pProductSkuId (const SLID* - pointer to GUID)
                    [UInt32],           # cbAppSpecificData (UINT)
                    [IntPtr],           # pvAppSpecificData (const PVOID - pointer to arbitrary data, typically IntPtr.Zero if not used)
                    [IntPtr],           # pActivationInfo (const SL_ACTIVATION_INFO_HEADER* - pointer to structure, typically IntPtr.Zero if not used)
                    [string],           # pwszProxyServer (PCWSTR - string for proxy server, can be $null)
                    [UInt16]            # wProxyPort (WORD - unsigned 16-bit integer for proxy port)
                )
            },
            @{
                # Probably internet activation API
                Name       = 'SLpIAActivateProduct'
                Dll        = 'sppc.dll'
                ReturnType = [uint32] # HRESULT
                Parameters = @(
                    [IntPtr],           # hSLC (HSLC handle)
                    [Guid].MakeByRefType() # pProductSkuId (const SLID* - pointer to GUID)
                )
            },
            @{
                # Probably Volume activation API
                Name       = 'SLpVLActivateProduct'
                Dll        = 'sppc.dll'
                ReturnType = [uint32] # HRESULT
                Parameters = @(
                    [IntPtr],           # hSLC (HSLC handle)
                    [Guid].MakeByRefType() # pProductSkuId (const SLID* - pointer to GUID)
                )
            },
            @{
                Name       = 'SLGetLicensingStatusInformation'
                Dll        = 'sppc.dll'
                ReturnType = [Int32] # HRESULT
                Parameters = @(
                    [IntPtr],                     # hSLC (HSLC handle)
                    [GUID].MakeByRefType(),       # pAppID (const SLID * - pass [IntPtr]::Zero or allocated GUID)
                    [GUID].MakeByRefType(),       # pProductSkuId (const SLID * - pass [IntPtr]::Zero or allocated GUID)
                    [IntPtr],                     # pwszRightName (PCWSTR - pass [IntPtr]::Zero for NULL)
                    [uint32].MakeByRefType(),     # pnStatusCount (UINT *)
                    [IntPtr].MakeByRefType()      # ppLicensingStatus (SL_LICENSING_STATUS **)
            )
        },
        @{
                Name       = 'SLConsumeWindowsRight'
                Dll        = 'slc.dll'
                ReturnType = [Int32] # HRESULT
                Parameters = @(
                    [IntPtr]                     # hSLC (HSLC handle)
            )
        },
        @{
                Name       = 'SLConsumeRight'
                Dll        = 'sppc.dll'
                ReturnType = [Int32] # HRESULT
                Parameters = @(
                    [IntPtr],                     # hSLC (HSLC handle)
                    [GUID].MakeByRefType(),       # pAppID (const SLID * - pass [IntPtr]::Zero or allocated GUID)
                    [IntPtr],                     # pProductSkuId (const SLID * - pass [IntPtr]::Zero or allocated GUID)
                    [IntPtr],                     # pwszRightName (PCWSTR - pass [IntPtr]::Zero for NULL)
                    [IntPtr]                      # pvReserved    -> Null
            )
        },
        @{
                Name       = 'SLGetPKeyId'
                Dll        = 'sppc.dll'
                ReturnType = [Int32] # HRESULT
                Parameters = @(
                    [IntPtr],                     # hSLC (HSLC handle)
                    [string],                     # pwszPKeyAlgorithm
                    [string],                     # pwszPKeyString
                    [IntPtr],                     # cbPKeySpecificData -> NULL
                    [IntPtr],                     # pbPKeySpecificData -> Null
                    [GUID].MakeByRefType()        # pPKeyId (const SLID * - pass [IntPtr]::Zero or allocated GUID)
            )
        },
        @{
                Name       = 'SLGenerateOfflineInstallationIdEx'
                Dll        = 'sppc.dll'
                ReturnType = [Int32] # HRESULT
                Parameters = @(
                    [IntPtr],                     # hSLC (HSLC handle)
                    [GUID].MakeByRefType(),       # pProductSkuId (const SLID * - pass [IntPtr]::Zero or allocated GUID)
                    [IntPtr],                     # const SL_ACTIVATION_INFO_HEADER *pActivationInfo // Zero
                    [IntPtr].MakeByRefType()      # [out] ppwszInstallationId
            )
        },
        @{
                Name       = 'SLGetActiveLicenseInfo'
                Dll        = 'sppc.dll'
                ReturnType = [Int32] # HRESULT
                Parameters = @(
                    [IntPtr],     # hSLC (HSLC handle)
                    [IntPtr],     # Reserved
                    [uint32].MakeByRefType(),
                    [IntPtr].MakeByRefType()
            )
        },
        @{
                Name       = 'SLGetTokenActivationGrants'
                Dll        = 'sppcext.dll'
                ReturnType = [Int32] # HRESULT
                Parameters = @(
                    [IntPtr],
                    [Guid].MakeByRefType(),
                    [IntPtr].MakeByRefType()
            )
        },
        @{
                Name       = 'SLFreeTokenActivationGrants'
                Dll        = 'sppcext.dll'
                ReturnType = [Int32] # HRESULT
                Parameters = @(
                    [IntPtr]
            )
        }
    )
    return Register-NativeMethods $functions
}
Function Init-NTDLL {
$functions = @(
    @{ Name = "NtDuplicateToken";          Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = @([IntPtr], [Int], [IntPtr], [Int], [Int], [IntPtr].MakeByRefType())},
    @{ Name = "NtQuerySystemInformation";  Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = @([Int32],[IntPtr],[Int32],[Int32].MakeByRefType())},
    @{ Name = "CsrClientCallServer";       Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = @([IntPtr],[IntPtr],[Int32],[Int32])},
    @{ Name = "NtResumeThread";            Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = @([IntPtr],[Int32])},
    @{ Name = "RtlMoveMemory";             Dll = "ntdll.dll"; ReturnType = [Void];   Parameters = @([IntPtr],[IntPtr],[UintPtr])},
    @{ Name = "RtlGetVersion";             Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = [Type[]]@([IntPtr]) },
    @{ Name = "RtlGetCurrentPeb";          Dll = "ntdll.dll"; ReturnType = [IntPtr]; Parameters = [Type[]]@() },
    @{ Name = "RtlGetProductInfo";         Dll = "ntdll.dll"; ReturnType = [Boolean];  Parameters = [Type[]]@([UInt32], [UInt32], [UInt32], [UInt32], [Uint32].MakeByRefType()) },
    @{ Name = "RtlGetNtVersionNumbers";    Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = [Type[]]@([Uint32].MakeByRefType(), [Uint32].MakeByRefType(), [Uint32].MakeByRefType()) },
    @{ Name = "RtlZeroMemory";             Dll = "ntdll.dll"; ReturnType = [Void];   Parameters = [Type[]]@([IntPtr], [UIntPtr]) },
    @{ Name = "RtlFreeHeap";               Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = [Type[]]@([IntPtr], [uint32], [IntPtr]) },
    @{ Name = "RtlGetProcessHeaps";        Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = [Type[]]@([Int32], [IntPtr]) },
    @{ Name = "NtGetNextProcess";          Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = [Type[]]@([IntPtr], [UInt32], [UInt32], [UInt32], [IntPtr].MakeByRefType()) },
    @{ Name = "NtQueryInformationProcess"; Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = [Type[]]@([IntPtr], [UInt32], [IntPtr], [UInt32], [UInt32].MakeByRefType()) },
    @{ Name = "ZwQueryLicenseValue";       Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = [Type[]]@([IntPtr], [UInt32].MakeByRefType(), [IntPtr], [UInt32], [UInt32].MakeByRefType()) },
    @{ Name = "RtlCreateUnicodeString";    Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = [Type[]]@([IntPtr], [string]) },
    @{ Name = "RtlFreeUnicodeString";      Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = [Type[]]@([IntPtr]) },
    @{ Name = "LdrGetDllHandleEx";         Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = [Type[]]@([Int32], [IntPtr], [IntPtr], [IntPtr], [IntPtr].MakeByRefType()) },
    @{ Name = "ZwQuerySystemInformation";  Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = [Type[]]@([int32], [IntPtr], [uint32], [uint32].MakeByRefType() ) },
    @{ Name = "RtlFindMessage";            Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = [Type[]]@(
        [IntPtr],                 # DllHandle
        [Uint32],                 # MessageTableId
        [Uint32],                 # MessageLanguageId
        [Uint32],                 # MessageId // ULONG
        [IntPtr].MakeByRefType()  # ref IntPtr for output MESSAGE_RESOURCE_ENTRY*
    ) },
    @{ Name = "RtlNtStatusToDosError"; Dll = "ntdll.dll"; ReturnType = [Int32]; Parameters = [Type[]]@([Int32]) },
        
    <#
        [In]String\Flags, [In][REF]Flags, [In][REF]UNICODE_STRING, [Out]Handle
        void LdrLoadDll(ulonglong param_1,uint *param_2,uint *param_3,undefined8 *param_4)

        https://rextester.com/KCUV42565
        RtlInitUnicodeStringStruct (&unicodestring, L"USER32.dll");
        LdrLoadDllStruct (NULL, 0, &unicodestring, &hModule);

        https://doxygen.reactos.org/d7/d55/ldrapi_8c_source.html
        NTSTATUS
        NTAPI
        DECLSPEC_HOTPATCH
        LdrLoadDll(
            _In_opt_ PWSTR SearchPath,
            _In_opt_ PULONG DllCharacteristics,
            _In_ PUNICODE_STRING DllName,
            _Out_ PVOID *BaseAddress)
        {
    #>
    @{ Name = "LdrLoadDll";    Dll = "ntdll.dll";     ReturnType = [Int32];      Parameters = [Type[]]@(
            
        # [IntPtr]::Zero // [STRING] -> NULL -> C Behavior
        [IntPtr],
            
        # [IntPtr]::Zero // Uint.makeByRef[]
        # Legit, no flags, can be 0x0 -> if (param_2 == (uint *)0x0) {uVar4 = 0;}
        [IntPtr],
            
        [IntPtr],                      # ModuleFileName Pointer (from RtlCreateUnicodeString)
        [IntPtr].MakeByRefType()       # out ModuleHandle
    ) },
    @{ Name = "LdrUnLoadDll";  Dll = "ntdll.dll";     ReturnType = [Int32];      Parameters = [Type[]]@(
        [IntPtr]                       # ModuleHandle (PVOID*)
    )},
    @{
        Name       = "LdrGetProcedureAddressForCaller"
        Dll        = "ntdll.dll"
        ReturnType = [Int32]
        Parameters = [Type[]]@(
            [IntPtr],                  # [HMODULE] Module handle pointer
            [IntPtr],                  # [PSTRING] Pointer to STRING struct (pass IntPtr directly, NOT [ref])
            [Int32],                   # [ULONG]   Ordinal / Flags (usually 0)
            [IntPtr].MakeByRefType(),  # [PVOID*]  Out pointer to procedure address (pass [ref])
            [byte],                    # [Flags]   0 or 1 (usually 0)
            [IntPtr]                   # [Caller]  Nullable caller address, pass [IntPtr]::Zero if none
        )
    },
    @{ Name = "NtOpenProcess";             Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = [Type[]]@([IntPtr].MakeByRefType(),[Int32], [IntPtr], [IntPtr]) },
    @{ Name = "NtClose";                   Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = [Type[]]@([IntPtr]) },
    @{ Name = "NtOpenProcessToken";        Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = [Type[]]@([IntPtr], [UInt32], [IntPtr].MakeByRefType()) },
    @{ Name = "NtAdjustPrivilegesToken";   Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = [Type[]]@([IntPtr], [bool] , [IntPtr], [UInt32], [IntPtr], [IntPtr]) },
    @{
        Name       = "NtCreateUserProcess";
        Dll        = "ntdll.dll";
        ReturnType = [Int32];
        Parameters = [Type[]]@(
            [IntPtr].MakeByRefType(),  # out PHANDLE ProcessHandle
            [IntPtr].MakeByRefType(),  # out PHANDLE ThreadHandle
            [Int32],                   # ACCESS_MASK ProcessDesiredAccess
            [Int32],                   # ACCESS_MASK ThreadDesiredAccess
            [IntPtr],                  # POBJECT_ATTRIBUTES ProcessObjectAttributes (nullable)
            [IntPtr],                  # POBJECT_ATTRIBUTES ThreadObjectAttributes (nullable)
            [UInt32],                  # ULONG ProcessFlags
            [UInt32],                  # ULONG ThreadFlags
            [IntPtr],                  # PRTL_USER_PROCESS_PARAMETERS (nullable)
            [IntPtr],                  # PPS_CREATE_INFO
            [IntPtr]                   # PPS_ATTRIBUTE_LIST (nullable)
        )
    },
    @{
        Name       = "RtlCreateProcessParametersEx";
        Dll        = "ntdll.dll";
        ReturnType = [Int32];
        Parameters = [Type[]]@(
            [IntPtr].MakeByRefType(),  # OUT PRTL_USER_PROCESS_PARAMETERS*
            [IntPtr],                  # PUNICODE_STRING ImagePathName
            [IntPtr],                  # PUNICODE_STRING DllPath
            [IntPtr],                  # PUNICODE_STRING CurrentDirectory
            [IntPtr],                  # PUNICODE_STRING CommandLine
            [IntPtr],                  # PVOID Environment
            [IntPtr],                  # PUNICODE_STRING WindowTitle
            [IntPtr],                  # PUNICODE_STRING DesktopInfo
            [IntPtr],                  # PUNICODE_STRING ShellInfo
            [IntPtr],                  # PUNICODE_STRING RuntimeData
            [Int32]                    # ULONG Flags
        )
    }
    @{ Name = "CsrCaptureMessageMultiUnicodeStringsInPlace";  Dll = "ntdll.dll"; ReturnType = [Int32];  Parameters = @(
        [IntPtr].MakeByRefType(),
        [Int32],[IntPtr])
    }
)
return Register-NativeMethods $functions
}
function Init-DismApi {

    <#
        Managed DismApi Wrapper
        https://github.com/jeffkl/ManagedDism/tree/main

        Windows 10 DLL File Information - DismApi.dll
        https://windows10dll.nirsoft.net/dismapi_dll.html

        DISMSDK
        https://github.com/Chuyu-Team/DISMSDK/blob/main/dismapi.h

        DISM API Functions
        https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/dism/dism-api-functions
    #>
    $functions = @(
        @{ 
            Name = "DismInitialize"; 
            Dll  = "DismApi.dll";
            ReturnType = [Int32]; 
            # logLevel, logFilePath, scratchDirectory
            Parameters = [Type[]]@([Int32], [IntPtr], [IntPtr])
        },
        @{ 
            Name = "DismOpenSession"; 
            Dll  = "DismApi.dll";
            ReturnType = [Int32]; 
            # imagePath, windowsDirectory, systemDrive, out session
            Parameters = [Type[]]@([string], [IntPtr], [IntPtr], [IntPtr].MakeByRefType()) 
        },
        @{ 
            Name = "DismCloseSession"; 
            Dll  = "DismApi.dll";
            ReturnType = [Int32]; 
            # session handle
            Parameters = [Type[]]@([IntPtr])
        },
        @{ 
            Name = "_DismGetTargetEditions"; 
            Dll  = "DismApi.dll";
            ReturnType = [Int32]; 
            # session, out editionIds, out count
            Parameters = [Type[]]@([IntPtr], [IntPtr].MakeByRefType(), [UInt32].MakeByRefType())
        },
        @{ 
            Name = "DismShutdown"; 
            Dll  = "DismApi.dll";
            ReturnType = [Int32]; 
            # no parameters
            Parameters = [Type[]]@()
        },
        @{
            Name = "DismDelete";
            Dll  = "DismApi.dll";
            ReturnType = [Int32];
            # parameter is a void* pointer to the structure to free
            Parameters = [Type[]]@([IntPtr])
        }
    )
    return Register-NativeMethods $functions
}
Function Init-advapi32 {

    $functions = @(
        @{ Name = "OpenProcessToken";        Dll = "advapi32.dll"; ReturnType = [UInt32]; Parameters = [Type[]]@([IntPtr], [UInt32], [IntPtr].MakeByRefType()) },
        @{ Name = "LookupPrivilegeValue";    Dll = "advapi32.dll"; ReturnType = [UInt32]; Parameters = [Type[]]@([IntPtr], [string], [Int64].MakeByRefType()) },
        @{ Name = "AdjustTokenPrivileges";   Dll = "advapi32.dll"; ReturnType = [UInt32]; Parameters = [Type[]]@([IntPtr], [bool] , [IntPtr], [Int32], [IntPtr], [IntPtr]) },
        @{ Name = "GetTokenInformation";     Dll = "advapi32.dll"; ReturnType = [UInt32]; Parameters = [Type[]]@([IntPtr], [Int32] , [IntPtr], [Int32], [Int32].MakeByRefType()) },
        @{ Name = "LookupPrivilegeNameW";    Dll = "advapi32.dll"; ReturnType = [UInt32]; Parameters = [Type[]]@([IntPtr], [Int32].MakeByRefType() , [IntPtr], [Int32].MakeByRefType()) },
        @{ Name = "LsaNtStatusToWinError";   Dll = "advapi32.dll"; ReturnType = [UInt32]; Parameters = [Type[]]@([UInt32]) },
        @{ Name = "LsaOpenPolicy";           Dll = "advapi32.dll"; ReturnType = [UInt32]; Parameters = [Type[]]@([IntPtr], [IntPtr], [UInt32], [IntPtr].MakeByRefType()) },
        @{ Name = "LsaLookupPrivilegeValue"; Dll = "advapi32.dll"; ReturnType = [UInt32]; Parameters = [Type[]]@([IntPtr], [IntPtr], [Int64].MakeByRefType()) },
        @{ Name = "LsaClose";                Dll = "advapi32.dll"; ReturnType = [UInt32]; Parameters = [Type[]]@([IntPtr]) },
        @{ Name = "OpenServiceW";            Dll = "advapi32.dll"; ReturnType = [IntPtr]; Parameters = [Type[]]@([IntPtr],[IntPtr],[Int32]) },
        @{ Name = "OpenSCManagerW";          Dll = "advapi32.dll"; ReturnType = [IntPtr]; Parameters = [Type[]]@([Int32],[IntPtr],[Int32]) },
        @{ Name = "CloseServiceHandle";      Dll = "advapi32.dll"; ReturnType = [BOOL];   Parameters = [Type[]]@([IntPtr]) },
        @{ Name = "StartServiceW";           Dll = "advapi32.dll"; ReturnType = [BOOL];   Parameters = [Type[]]@([IntPtr],[Int32],[IntPtr]) },
        @{ Name = "QueryServiceStatusEx";    Dll = "advapi32.dll"; ReturnType = [BOOL];   Parameters = [Type[]]@([IntPtr],[Int32],[IntPtr],[Int32],[UInt32].MakeByRefType()) },
        @{ Name = "CreateProcessWithTokenW"; Dll = "advapi32.dll"; ReturnType = [BOOL];   Parameters = [Type[]]@([IntPtr], [Int32], [IntPtr], [IntPtr], [Int32], [IntPtr],[IntPtr],[IntPtr],[IntPtr]) }
    )
    return Register-NativeMethods $functions
}
Function Init-KERNEL32 {

    $functions = @(
        @{ Name = "RevertToSelf";            Dll = "KernelBase.dll"; ReturnType = [bool]; Parameters = [Type[]]@() },
        @{ Name = "ImpersonateLoggedOnUser"; Dll = "KernelBase.dll"; ReturnType = [bool]; Parameters = [Type[]]@([IntPtr]) },
        @{ Name = "FindFirstFileW"; Dll = "KernelBase.dll"; ReturnType = [IntPtr]; Parameters = [Type[]]@([string], [IntPtr]) },
        @{ Name = "FindNextFileW";  Dll = "KernelBase.dll"; ReturnType = [bool];   Parameters = [Type[]]@([IntPtr], [IntPtr]) },
        @{ Name = "FindClose";      Dll = "KernelBase.dll"; ReturnType = [bool];   Parameters = [Type[]]@([IntPtr]) },
        @{ Name = "LocalFree" ;     Dll = "KernelBase.dll"; ReturnType = [IntPtr]; Parameters = [Type[]]@([IntPtr]) },
        @{ Name = "LoadLibraryExW"; Dll = "KernelBase.dll"; ReturnType = [IntPtr]; Parameters = [Type[]]@([string], [IntPtr], [UInt32]) },
        @{ Name = "FreeLibrary";    Dll = "KernelBase.dll"; ReturnType = [BOOL];   Parameters = [Type[]]@([IntPtr]) },
        @{ Name = "HeapFree";       Dll = "KernelBase.dll"; ReturnType = [bool]  ; Parameters = [Type[]]@([IntPtr], [uint32], [IntPtr]) },
        @{ Name = "ResumeThread";   Dll = "KernelBase.dll"; ReturnType = [int32];  Parameters = [Type[]]@([IntPtr]) },
        @{ Name = "GetProcAddress"; Dll = "KernelBase.dll"; ReturnType = [IntPtr]; Parameters = [Type[]]@([IntPtr], [string]) },
        @{ Name = "CloseHandle";    Dll = "KernelBase.dll"; ReturnType = [bool];   Parameters = [Type[]]@([IntPtr]) },
        @{ Name = "LocalFree";      Dll = "KernelBase.dll"; ReturnType = [bool];   Parameters = [Type[]]@([IntPtr]) },
        @{ Name = "CreateProcessW"; Dll = "KernelBase.dll"; ReturnType = [bool];   Parameters = [Type[]]@([IntPtr],[IntPtr],[IntPtr],[IntPtr],[bool],[Int32],[IntPtr],[IntPtr],[IntPtr],[IntPtr]) },
        @{ Name = "WaitForSingleObject";   Dll = "KernelBase.dll"; ReturnType = [int32];  Parameters = [Type[]]@([IntPtr],[int32]) },
        @{ Name = "EnumSystemFirmwareTables"; Dll = "KernelBase.dll"; ReturnType = [UInt32]; Parameters = [Type[]]@([UInt32], [IntPtr], [UInt32]) },
        @{ Name = "GetSystemFirmwareTable";   Dll = "KernelBase.dll"; ReturnType = [UInt32]; Parameters = [Type[]]@([UInt32], [UInt32], [IntPtr], [UInt32]) },
        @{ Name = "UpdateProcThreadAttribute";  Dll = "KernelBase.dll"; ReturnType = [bool];   Parameters = [Type[]]@([IntPtr],[uint32],[uint32],[IntPtr],[Int32],[IntPtr],[IntPtr]) },
        @{ Name = "InitializeProcThreadAttributeList";    Dll = "KernelBase.dll"; ReturnType = [bool];   Parameters = [Type[]]@([IntPtr],[uint32],[uint32],[IntPtr]) },
        @{ Name = "DeleteProcThreadAttributeList";    Dll = "KernelBase.dll"; ReturnType = [void];   Parameters = [Type[]]@([IntPtr]) }
    )
    return Register-NativeMethods $functions
}
Function Init-PKHELPER {

    $functions = @(
        @{
            Name = "GetEditionIdFromName"
            Dll = "pkeyhelper.dll"
            ReturnType = [int]
            Parameters = [Type[]]@(
                [string],                     # edition Name
                [int].MakeByRefType()         # out Branding Value
            )
        },
        @{
            Name = "GetEditionNameFromId"
            Dll = "pkeyhelper.dll"
            ReturnType = [int]
            Parameters = [Type[]]@(
                [int],                     # Branding Value
                [intptr].MakeByRefType()   # out edition Name
            )
        },
        @{
            Name = "SkuGetProductKeyForEdition"
            Dll = "pkeyhelper.dll"
            ReturnType = [int]
            Parameters = [Type[]]@(
                [int],                    # editionId
                [IntPtr],                 # sku
                [IntPtr].MakeByRefType()  # ref productKey
                [IntPtr].MakeByRefType()  # ref keyType
            )
        },
        @{
            Name = "IsDefaultPKey"
            Dll = "pkeyhelper.dll"
            ReturnType = [uint32]
            Parameters = [Type[]]@(
                [string],               # 29 digits cd-key
                [bool].MakeByRefType()  # Default bool Propertie = 0, [ref]$Propertie
            )

        <#
            [bool]$results = 0
            $hr = $Global:PKHElper::IsDefaultPKey(
                "89DNY-M3VP8-CB7JK-3QGBC-Q3WV6", [ref]$results)
            if ($hr -eq 0) {
	            $results
            }            
        #>
        },
        @{
            Name = "GetDefaultProductKeyForPfn"
            Dll = "pkeyhelper.dll"
            ReturnType = [uint32]
            Parameters = [Type[]]@(
                [string],                  # "Microsoft.Windows.100.res-v3274_8wekyb3d8bbwe"
                [IntPtr].MakeByRefType(),  # Handle to result
                [uint32]                   # Flags
            )

        <#
            $output = [IntPtr]::Zero
            $hr = $Global:PKHElper::GetDefaultProductKeyForPfn(
                "Microsoft.Windows.100.res-v3274_8wekyb3d8bbwe", [ref]$output, 0)
            if ($hr -eq 0) {
	            [marshal]::PtrToStringUni($outPut)
                # free pointer later
            }            
        #>
        }
    )
    return Register-NativeMethods $functions
}
Function Init-PIDGENX {
     
    <#
    https://github.com/chughes-3
    https://github.com/chughes-3/UpdateProductKey/blob/master/UpdateProductKeys/PidChecker.cs

    [DllImport("pidgenx.dll", EntryPoint = "PidGenX", CharSet = CharSet.Auto)]
    static extern int PidGenX(string ProductKey, string PkeyPath, string MSPID, int UnknownUsage, IntPtr ProductID, IntPtr DigitalProductID, IntPtr DigitalProductID4);

    * sppcomapi.dll
    * __int64 __fastcall GetWindowsPKeyInfo(_WORD *a1, __int64 a2, __int64 a3, __int64 a4)
    __int128 v46[3]; // __m128 v46[3], 48 bytes total
    int v47[44];
    int v48[320];
    memset(v46, 0, sizeof(v46)); // size of structure 2
    memset_0(v47, 0, 0xA4ui64);
    memset_0(v48, 0, 0x4F8ui64);
    v47[0] = 164;   // size of structure 3
    v48[0] = 1272;  // size of structure 4

    $PIDPtr   = New-IntPtr -Size 0x30  -WriteSizeAtZero
    $DPIDPtr  = New-IntPtr -Size 0xB0  -InitialValue 0xA4
    $DPID4Ptr = New-IntPtr -Size 0x500 -InitialValue 0x4F8

    $result = $Global:PIDGENX::PidGenX(
        # Most important Roles
        $key, $configPath,
        # Default value for MSPID, 03612 ?? 00000 ?
        # PIDGENX2 -> v26 = L"00000" // SPPCOMAPI, GetWindowsPKeyInfo -> L"03612"
        "00000",
        # Unknown1
        0,
        # Structs
        $PIDPtr, $DPIDPtr, $DPID4Ptr
    )

    $result = $Global:PIDGENX::PidGenX2(
        # Most important Roles
        $key, $configPath,
        # Default value for MSPID, 03612 ?? 00000 ?
        # PIDGENX2 -> v26 = L"00000" // SPPCOMAPI, GetWindowsPKeyInfo -> L"03612"
        "00000",
        # Unknown1 / [Unknown2, Added in PidGenX2!]
        0,0,
        # Structs
        $PIDPtr, $DPIDPtr, $DPID4Ptr
    )
    #>

    $functions = @(
        @{
            Name       = "PidGenX"
            Dll        = "pidgenx.dll"
            ReturnType = [int]
            Parameters = [Type[]]@([string], [string], [string], [int], [IntPtr], [IntPtr], [IntPtr])
        },
        @{
            Name       = "PidGenX2"
            Dll        = "pidgenx.dll"
            ReturnType = [int]
            Parameters = [Type[]]@([string], [string], [string], [int], [int], [IntPtr], [IntPtr], [IntPtr])
        }
    )
    return Register-NativeMethods $functions -ImplAttributes ([MethodImplAttributes]::IL)
}

<#

     *********************

      !Managed Api & 
             Com Warper.!
        -  Helper's -

     *********************

    Get-SysCallData <> based on PowerSploit 3.0.0.0
    https://www.powershellgallery.com/packages/PowerSploit/3.0.0.0
    https://www.powershellgallery.com/packages/PowerSploit/1.0.0.0/Content/PETools%5CGet-PEHeader.ps1

#>
function Get-Base26Name {
    param (
        [int]$idx
    )

    $result = [System.Text.StringBuilder]::new()
    while ($idx -ge 0) {
        $remainder = $idx % 26
        [void]$result.Insert(0, [char](65 + $remainder))
        $idx = [math]::Floor($idx / 26) - 1
    }

    return $result.ToString()
}
function Process-Parameters {
    param (
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$InterfaceSpec,

        [switch]$Ignore
    )

    # Initialize the parameters list with the base parameter (thisPtr)
    $allParams = New-Object System.Collections.Generic.List[string]

    if (-not $Ignore) {
       $BaseParams = "IntPtr thisPtr"
       $allParams.Add($BaseParams) # Add the base parameter (thisPtr) first
    }

    # Process user-provided parameters if they exist
    if (-not [STRING]::IsNullOrEmpty($InterfaceSpec.Params)) {
        # Split the user-provided parameters by comma and trim whitespace
        $userParams = $InterfaceSpec.Params.Split(',') | ForEach-Object { $_.Trim() }
        
        foreach ($param in $userParams) {
            $modifier = ""
            $typeAndName = $param

            # Check for 'ref' or 'out' keywords, optionally wrapped in brackets, and separate them
            if ($param -match "^\s*\[?(ref|out)\]?\s+(.+)$") {
                $modifier = $Matches[1]                 # This will capture "ref" or "out" (e.g., if input was "[REF]", $Matches[1] will be "REF")
                $modifier = $modifier.ToLowerInvariant() # Convert modifier to lowercase ("REF" -> "ref")
                $typeAndName = $Matches[2]             # Extract the actual type and name
            }

            # Split the type and name (e.g., "uint Flags" -> "uint", "Flags")
            $parts = $typeAndName.Split(' ', 2) 
            if ($parts.Length -eq 2) {
                $type = $parts[0]
                $name = $parts[1]
                $fixedType = $type # Default to original type if no match

                switch ($type.ToLowerInvariant()) {
                    # Fully qualified .NET types
                    "system.boolean" { $fixedType = "bool" }
                    "system.byte"    { $fixedType = "byte" }
                    "system.char"    { $fixedType = "char" }
                    "system.decimal" { $fixedType = "decimal" }
                    "system.double"  { $fixedType = "double" }
                    "system.int16"   { $fixedType = "short" }
                    "system.int32"   { $fixedType = "int" }
                    "system.int64"   { $fixedType = "long" }
                    "system.intptr"  { $fixedType = "IntPtr" }
                    "system.object"  { $fixedType = "object" }
                    "system.sbyte"   { $fixedType = "sbyte" }
                    "system.single"  { $fixedType = "float" }
                    "system.string"  { $fixedType = "string" }
                    "system.uint16"  { $fixedType = "ushort" }
                    "system.uint32"  { $fixedType = "uint" }
                    "system.uint64"  { $fixedType = "ulong" }
                    "system.uintptr" { $fixedType = "UIntPtr" }

                    # Alternate type spellings and aliases
                    "boolean"        { $fixedType = "bool" }
                    "dword32"        { $fixedType = "uint" }
                    "dword64"        { $fixedType = "ulong" }
                    "int16"          { $fixedType = "short" }
                    "int32"          { $fixedType = "int" }
                    "int64"          { $fixedType = "long" }
                    "single"         { $fixedType = "float" }
                    "uint16"         { $fixedType = "ushort" }
                    "uint32"         { $fixedType = "uint" }
                    "uint64"         { $fixedType = "ulong" }

                    # --- Additional C/C++ & WinAPI aliases ---
                    "double"         { $fixedType = "double" }
                    "float"          { $fixedType = "float" }
                    "long"           { $fixedType = "long" }
                    "longlong"       { $fixedType = "long" }
                    "tchar"          { $fixedType = "char" }
                    "uchar"          { $fixedType = "byte" }
                    "ulong"          { $fixedType = "ulong" }
                    "ulonglong"      { $fixedType = "ulong" }
                    "short"          { $fixedType = "short" }
                    "ushort"         { $fixedType = "ushort" }

                    # --- Additional typedefs ---
                    "atom"           { $fixedType = "ushort" }
                    "dword_ptr"      { $fixedType = "UIntPtr" }
                    "dwordlong"      { $fixedType = "ulong" }
                    "farproc"        { $fixedType = "IntPtr" }
                    "hhook"          { $fixedType = "IntPtr" }
                    "hresult"        { $fixedType = "int" }
                    "NTSTATUS"       { $fixedType = "Int32" }
                    "int_ptr"        { $fixedType = "IntPtr" }
                    "intptr_t"       { $fixedType = "IntPtr" }
                    "long_ptr"       { $fixedType = "IntPtr" }
                    "lpbyte"         { $fixedType = "IntPtr" }
                    "lpdword"        { $fixedType = "IntPtr" }
                    "lparam"         { $fixedType = "IntPtr" }
                    "pcstr"          { $fixedType = "IntPtr" }
                    "pcwstr"         { $fixedType = "IntPtr" }
                    "pstr"           { $fixedType = "IntPtr" }
                    "pwstr"          { $fixedType = "IntPtr" }
                    "uint_ptr"       { $fixedType = "UIntPtr" }
                    "uintptr_t"      { $fixedType = "UIntPtr" }
                    "wparam"         { $fixedType = "UIntPtr" }

                    # C# built-in types
                    "bool"           { $fixedType = "bool" }
                    "byte"           { $fixedType = "byte" }
                    "char"           { $fixedType = "char" }
                    "decimal"        { $fixedType = "decimal" }
                    "int"            { $fixedType = "int" }
                    "intptr"         { $fixedType = "IntPtr" }
                    "nint"           { $fixedType = "nint" }
                    "nuint"          { $fixedType = "nuint" }
                    "object"         { $fixedType = "object" }
                    "sbyte"          { $fixedType = "sbyte" }
                    "string"         { $fixedType = "string" }
                    "uint"           { $fixedType = "uint" }
                    "uintptr"        { $fixedType = "UIntPtr" }

                    # Common WinAPI handle types
                    "hbitmap"        { $fixedType = "IntPtr" }
                    "hbrush"         { $fixedType = "IntPtr" }
                    "hcurs"          { $fixedType = "IntPtr" }
                    "hdc"            { $fixedType = "IntPtr" }
                    "hfont"          { $fixedType = "IntPtr" }
                    "hicon"          { $fixedType = "IntPtr" }
                    "hmenu"          { $fixedType = "IntPtr" }
                    "hpen"           { $fixedType = "IntPtr" }
                    "hrgn"           { $fixedType = "IntPtr" }

                    # Pointer-based aliases
                    "pbyte"          { $fixedType = "IntPtr" }
                    "pchar"          { $fixedType = "IntPtr" }
                    "pdword"         { $fixedType = "IntPtr" }
                    "pint"           { $fixedType = "IntPtr" }
                    "plong"          { $fixedType = "IntPtr" }
                    "puint"          { $fixedType = "IntPtr" }
                    "pulong"         { $fixedType = "IntPtr" }
                    "pvoid"          { $fixedType = "IntPtr" }
                    "lpvoid"         { $fixedType = "IntPtr" }

                    # Special types
                    "guid"           { $fixedType = "Guid" }

                    # Windows/WinAPI types (common aliases)
                    "dword"          { $fixedType = "uint" }
                    "handle"         { $fixedType = "IntPtr" }
                    "hinstance"      { $fixedType = "IntPtr" }
                    "hmodule"        { $fixedType = "IntPtr" }
                    "hwnd"           { $fixedType = "IntPtr" }
                    "ptr"            { $fixedType = "IntPtr" }
                    "size_t"         { $fixedType = "UIntPtr" }
                    "ssize_t"        { $fixedType = "IntPtr" }
                    "void*"          { $fixedType = "IntPtr" }
                    "word"           { $fixedType = "ushort" }
                    "phandle"        { $fixedType = "IntPtr" }
                    "lresult"        { $fixedType = "IntPtr" }

                    # STRSAFE typedefs
                    "strsafe_lpstr"       { $fixedType = "string" }       # ANSI
                    "strsafe_lpcstr"      { $fixedType = "string" }       # ANSI
                    "strsafe_lpwstr"      { $fixedType = "string" }       # Unicode
                    "strsafe_lpcwstr"     { $fixedType = "string" }       # Unicode
                    "strsafe_lpcuwstr"    { $fixedType = "string" }       # Unicode
                    "strsafe_pcnzch"      { $fixedType = "string" }       # ANSI char
                    "strsafe_pcnzwch"     { $fixedType = "string" }       # Unicode wchar
                    "strsafe_pcunzwch"    { $fixedType = "string" }       # Unicode wchar

                    # Wide-character (Unicode) types
                    "lpcstr"        { $fixedType = "string" }             # ANSI string
                    "lpcwstr"       { $fixedType = "string" }             # Unicode string
                    "lpstr"         { $fixedType = "string" }             # ANSI string
                    "lpwstr"        { $fixedType = "string" }             # Unicode string
                    "pstring"       { $fixedType = "string" }             # ANSI string (likely)
                    "pwchar"        { $fixedType = "string" }             # Unicode char*
                    "lpwchar"       { $fixedType = "string" }             # Unicode char*
                    "pczpwstr"      { $fixedType = "string" }             # Unicode string
                    "pzpwstr"       { $fixedType = "string" }
                    "pzwstr"        { $fixedType = "string" }
                    "pzzwstr"       { $fixedType = "string" }
                    "pczzwstr"      { $fixedType = "string" }
                    "puczzwstr"     { $fixedType = "string" }
                    "pcuczzwstr"    { $fixedType = "string" }
                    "pnzwch"        { $fixedType = "string" }
                    "pcnzwch"       { $fixedType = "string" }
                    "punzwch"       { $fixedType = "string" }
                    "pcunzwch"      { $fixedType = "string" }

                    # ANSI string types
                    "npstr"         { $fixedType = "string" }             # ANSI string
                    "pzpcstr"       { $fixedType = "string" }
                    "pczpcstr"      { $fixedType = "string" }
                    "pzzstr"        { $fixedType = "string" }
                    "pczzstr"       { $fixedType = "string" }
                    "pnzch"         { $fixedType = "string" }
                    "pcnzch"        { $fixedType = "string" }

                    # UCS types
                    "ucschar"       { $fixedType = "uint" }               # leave as uint
                    "pucschar"      { $fixedType = "IntPtr" }
                    "pcucschar"     { $fixedType = "IntPtr" }
                    "puucschar"     { $fixedType = "IntPtr" }
                    "pcuucschar"    { $fixedType = "IntPtr" }
                    "pucsstr"       { $fixedType = "IntPtr" }
                    "pcucsstr"      { $fixedType = "IntPtr" }
                    "puucsstr"      { $fixedType = "IntPtr" }
                    "pcuucsstr"     { $fixedType = "IntPtr" }

                    # Neutral ANSI/Unicode (TCHAR-based) Types
                    "ptchar"        { $fixedType = "IntPtr" }              # keep IntPtr due to TCHAR ambiguity
                    "tbyte"         { $fixedType = "byte" }
                    "ptbyte"        { $fixedType = "IntPtr" }
                    "ptstr"         { $fixedType = "IntPtr" }
                    "lptstr"        { $fixedType = "IntPtr" }
                    "pctstr"        { $fixedType = "IntPtr" }
                    "lpctstr"       { $fixedType = "IntPtr" }
                    "putstr"        { $fixedType = "IntPtr" }
                    "lputstr"       { $fixedType = "IntPtr" }
                    "pcutstr"       { $fixedType = "IntPtr" }
                    "lpcutstr"      { $fixedType = "IntPtr" }
                    "pzptstr"       { $fixedType = "IntPtr" }
                    "pzzstr"        { $fixedType = "IntPtr" }
                    "pczztstr"      { $fixedType = "IntPtr" }
                    "pzzwstr"       { $fixedType = "string" }             # Unicode string
                    "pczzwstr"      { $fixedType = "string" }
                }
                # Reconstruct the parameter string with the fixed type and optional modifier
                $formattedParam = "$fixedType $name"
                if (-not [STRING]::IsNullOrEmpty($modifier)) {
                    $formattedParam = "$modifier $formattedParam"
                }
                $allParams.Add($formattedParam)
            } else {
                # If the parameter couldn't be parsed, add it as is
                $allParams.Add($param)
            }
        }
    }
    
    # Join all processed parameters with a comma and add indentation for readability
    $Params = $allParams -join ("," + "`n" + " " * 10)

    return $Params
}
function Process-ReturnType {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$ReturnType
    )

    $fixedReturnType = $ReturnType

    switch ($ReturnType.ToLowerInvariant()) {
        
        # Void
        "void"           { $fixedReturnType = "void" }

        # Fully qualified .NET types
        "system.boolean" { $fixedReturnType = "bool" }
        "system.byte"    { $fixedReturnType = "byte" }
        "system.char"    { $fixedReturnType = "char" }
        "system.decimal" { $fixedReturnType = "decimal" }
        "system.double"  { $fixedReturnType = "double" }
        "system.int16"   { $fixedReturnType = "short" }
        "system.int32"   { $fixedReturnType = "int" }
        "system.int64"   { $fixedReturnType = "long" }
        "system.intptr"  { $fixedReturnType = "IntPtr" }
        "system.object"  { $fixedReturnType = "object" }
        "system.sbyte"   { $fixedReturnType = "sbyte" }
        "system.single"  { $fixedReturnType = "float" }
        "system.string"  { $fixedReturnType = "string" }
        "system.uint16"  { $fixedReturnType = "ushort" }
        "system.uint32"  { $fixedReturnType = "uint" }
        "system.uint64"  { $fixedReturnType = "ulong" }
        "system.uintptr" { $fixedReturnType = "UIntPtr" }

        # Alternate type spellings and aliases
        "boolean"        { $fixedReturnType = "bool" }
        "dword32"        { $fixedReturnType = "uint" }
        "dword64"        { $fixedReturnType = "ulong" }
        "int16"          { $fixedReturnType = "short" }
        "int32"          { $fixedReturnType = "int" }
        "int64"          { $fixedReturnType = "long" }
        "single"         { $fixedReturnType = "float" }
        "uint16"         { $fixedReturnType = "ushort" }
        "uint32"         { $fixedReturnType = "uint" }
        "uint64"         { $fixedReturnType = "ulong" }

        # --- Additional C/C++ & WinAPI aliases ---
        "double"         { $fixedReturnType = "double" }
        "float"          { $fixedReturnType = "float" }
        "long"           { $fixedReturnType = "long" }
        "longlong"       { $fixedReturnType = "long" }
        "tchar"          { $fixedReturnType = "char" }
        "uchar"          { $fixedReturnType = "byte" }
        "ulong"          { $fixedReturnType = "ulong" }
        "ulonglong"      { $fixedReturnType = "ulong" }
        "short"          { $fixedReturnType = "short" }
        "ushort"         { $fixedReturnType = "ushort" }

        # --- Additional typedefs ---
        "atom"           { $fixedReturnType = "ushort" }
        "dword_ptr"      { $fixedReturnType = "UIntPtr" }
        "dwordlong"      { $fixedReturnType = "ulong" }
        "farproc"        { $fixedReturnType = "IntPtr" }
        "hhook"          { $fixedReturnType = "IntPtr" }
        "hresult"        { $fixedReturnType = "int" }
        "NTSTATUS"       { $fixedReturnType = "Int32" }
        "int_ptr"        { $fixedReturnType = "IntPtr" }
        "intptr_t"       { $fixedReturnType = "IntPtr" }
        "long_ptr"       { $fixedReturnType = "IntPtr" }
        "lpbyte"         { $fixedReturnType = "IntPtr" }
        "lpdword"        { $fixedReturnType = "IntPtr" }
        "lparam"         { $fixedReturnType = "IntPtr" }
        "pcstr"          { $fixedReturnType = "IntPtr" }
        "pcwstr"         { $fixedReturnType = "IntPtr" }
        "pstr"           { $fixedReturnType = "IntPtr" }
        "pwstr"          { $fixedReturnType = "IntPtr" }
        "uint_ptr"       { $fixedReturnType = "UIntPtr" }
        "uintptr_t"      { $fixedReturnType = "UIntPtr" }
        "wparam"         { $fixedReturnType = "UIntPtr" }

        # C# built-in types
        "bool"           { $fixedReturnType = "bool" }
        "byte"           { $fixedReturnType = "byte" }
        "char"           { $fixedReturnType = "char" }
        "decimal"        { $fixedReturnType = "decimal" }
        "int"            { $fixedReturnType = "int" }
        "intptr"         { $fixedReturnType = "IntPtr" }
        "nint"           { $fixedReturnType = "nint" }
        "nuint"          { $fixedReturnType = "nuint" }
        "object"         { $fixedReturnType = "object" }
        "sbyte"          { $fixedReturnType = "sbyte" }
        "string"         { $fixedReturnType = "string" }
        "uint"           { $fixedReturnType = "uint" }
        "uintptr"        { $fixedReturnType = "UIntPtr" }

        # Common WinAPI handle types
        "hbitmap"        { $fixedReturnType = "IntPtr" }
        "hbrush"         { $fixedReturnType = "IntPtr" }
        "hcurs"          { $fixedReturnType = "IntPtr" }
        "hdc"            { $fixedReturnType = "IntPtr" }
        "hfont"          { $fixedReturnType = "IntPtr" }
        "hicon"          { $fixedReturnType = "IntPtr" }
        "hmenu"          { $fixedReturnType = "IntPtr" }
        "hpen"           { $fixedReturnType = "IntPtr" }
        "hrgn"           { $fixedReturnType = "IntPtr" }

        # Pointer-based aliases
        "pbyte"          { $fixedReturnType = "IntPtr" }
        "pchar"          { $fixedReturnType = "IntPtr" }
        "pdword"         { $fixedReturnType = "IntPtr" }
        "pint"           { $fixedReturnType = "IntPtr" }
        "plong"          { $fixedReturnType = "IntPtr" }
        "puint"          { $fixedReturnType = "IntPtr" }
        "pulong"         { $fixedReturnType = "IntPtr" }
        "pvoid"          { $fixedReturnType = "IntPtr" }
        "lpvoid"         { $fixedReturnType = "IntPtr" }

        # Special types
        "guid"           { $fixedReturnType = "Guid" }

        # Windows/WinAPI types (common aliases)
        "dword"          { $fixedReturnType = "uint" }
        "handle"         { $fixedReturnType = "IntPtr" }
        "hinstance"      { $fixedReturnType = "IntPtr" }
        "hmodule"        { $fixedReturnType = "IntPtr" }
        "hwnd"           { $fixedReturnType = "IntPtr" }
        "ptr"            { $fixedReturnType = "IntPtr" }
        "size_t"         { $fixedReturnType = "UIntPtr" }
        "ssize_t"        { $fixedReturnType = "IntPtr" }
        "void*"          { $fixedReturnType = "IntPtr" }
        "word"           { $fixedReturnType = "ushort" }
        "phandle"        { $fixedReturnType = "IntPtr" }
        "lresult"        { $fixedReturnType = "IntPtr" }                  

        # STRSAFE typedefs
        "strsafe_lpstr"       { $fixedReturnType = "string" }       # ANSI
        "strsafe_lpcstr"      { $fixedReturnType = "string" }       # ANSI
        "strsafe_lpwstr"      { $fixedReturnType = "string" }       # Unicode
        "strsafe_lpcwstr"     { $fixedReturnType = "string" }       # Unicode
        "strsafe_lpcuwstr"    { $fixedReturnType = "string" }       # Unicode
        "strsafe_pcnzch"      { $fixedReturnType = "string" }       # ANSI char
        "strsafe_pcnzwch"     { $fixedReturnType = "string" }       # Unicode wchar
        "strsafe_pcunzwch"    { $fixedReturnType = "string" }       # Unicode wchar

        # Wide-character (Unicode) types
        "lpcstr"        { $fixedReturnType = "string" }             # ANSI string
        "lpcwstr"       { $fixedReturnType = "string" }             # Unicode string
        "lpstr"         { $fixedReturnType = "string" }             # ANSI string
        "lpwstr"        { $fixedReturnType = "string" }             # Unicode string
        "pstring"       { $fixedReturnType = "string" }             # ANSI string (likely)
        "pwchar"        { $fixedReturnType = "string" }             # Unicode char*
        "lpwchar"       { $fixedReturnType = "string" }             # Unicode char*
        "pczpwstr"      { $fixedReturnType = "string" }             # Unicode string
        "pzpwstr"       { $fixedReturnType = "string" }
        "pzwstr"        { $fixedReturnType = "string" }
        "pzzwstr"       { $fixedReturnType = "string" }
        "pczzwstr"      { $fixedReturnType = "string" }
        "puczzwstr"     { $fixedReturnType = "string" }
        "pcuczzwstr"    { $fixedReturnType = "string" }
        "pnzwch"        { $fixedReturnType = "string" }
        "pcnzwch"       { $fixedReturnType = "string" }
        "punzwch"       { $fixedReturnType = "string" }
        "pcunzwch"      { $fixedReturnType = "string" }

        # ANSI string types
        "npstr"         { $fixedReturnType = "string" }             # ANSI string
        "pzpcstr"       { $fixedReturnType = "string" }
        "pczpcstr"      { $fixedReturnType = "string" }
        "pzzstr"        { $fixedReturnType = "string" }
        "pczzstr"       { $fixedReturnType = "string" }
        "pnzch"         { $fixedReturnType = "string" }
        "pcnzch"        { $fixedReturnType = "string" }

        # UCS types
        "ucschar"       { $fixedReturnType = "uint" }               # leave as uint
        "pucschar"      { $fixedReturnType = "IntPtr" }
        "pcucschar"     { $fixedReturnType = "IntPtr" }
        "puucschar"     { $fixedReturnType = "IntPtr" }
        "pcuucschar"    { $fixedReturnType = "IntPtr" }
        "pucsstr"       { $fixedReturnType = "IntPtr" }
        "pcucsstr"      { $fixedReturnType = "IntPtr" }
        "puucsstr"      { $fixedReturnType = "IntPtr" }
        "pcuucsstr"     { $fixedReturnType = "IntPtr" }

        # Neutral ANSI/Unicode (TCHAR-based) Types
        "ptchar"        { $fixedReturnType = "IntPtr" }              # keep IntPtr due to TCHAR ambiguity
        "tbyte"         { $fixedReturnType = "byte" }
        "ptbyte"        { $fixedReturnType = "IntPtr" }
        "ptstr"         { $fixedReturnType = "IntPtr" }
        "lptstr"        { $fixedReturnType = "IntPtr" }
        "pctstr"        { $fixedReturnType = "IntPtr" }
        "lpctstr"       { $fixedReturnType = "IntPtr" }
        "putstr"        { $fixedReturnType = "IntPtr" }
        "lputstr"       { $fixedReturnType = "IntPtr" }
        "pcutstr"       { $fixedReturnType = "IntPtr" }
        "lpcutstr"      { $fixedReturnType = "IntPtr" }
        "pzptstr"       { $fixedReturnType = "IntPtr" }
        "pzzstr"        { $fixedReturnType = "IntPtr" }
        "pczztstr"      { $fixedReturnType = "IntPtr" }
        "pzzwstr"       { $fixedReturnType = "string" }             # Unicode string
        "pczzwstr"      { $fixedReturnType = "string" }
    }

    return $fixedReturnType
}
function Invoke-Object {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        $Interface,

        [Parameter(ValueFromRemainingArguments = $true)]
        [object[]]$Params,

        [Parameter(Mandatory)]
        [ValidateSet("API", "COM")]
        [string]$type
    )
    [int]$count = 0
    [void][Int]::TryParse($Params.Count, [ref]$count)
    
    $sb = New-Object System.Text.StringBuilder
    if ($type -eq 'COM') {
        if ($count -gt 0) {
            [void]$sb.Append('$Interface.IUnknownPtr,')
        } else {
            [void]$sb.Append('$Interface.IUnknownPtr')
        }
    }
    if ($count -gt 0) {
        for ($i = 0; $i -lt $count; $i++) {
            if ($i -gt 0) {
                [void]$sb.Append(',')
            }
            [void]$sb.Append("`$Params[$i]")
        }
    }

    $argsString = $sb.ToString()
    return & (
        [scriptblock]::Create("`$Interface.DelegateInstance.Invoke($argsString)")
    )
}
function Get-SysCallData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$DllName,

        [Parameter(Mandatory = $true)]
        [string]$FunctionName,

        [Parameter(Mandatory = $true)]
        [int]$BytesToRead
    )

if (!([PSTypeName]'PE').Type) {
$code = @"
using System;
using System.Runtime.InteropServices;

public class PE
{
    [Flags]
    public enum IMAGE_DOS_SIGNATURE : ushort
    {
        DOS_SIGNATURE = 0x5A4D, // MZ
        OS2_SIGNATURE = 0x454E, // NE
        OS2_SIGNATURE_LE = 0x454C, // LE
        VXD_SIGNATURE = 0x454C, // LE
    }
        
    [Flags]
    public enum IMAGE_NT_SIGNATURE : uint
    {
        VALID_PE_SIGNATURE = 0x00004550 // PE00
    }
        
    [Flags]
    public enum IMAGE_FILE_MACHINE : ushort
    {
        UNKNOWN = 0,
        I386 = 0x014c, // Intel 386.
        R3000 = 0x0162, // MIPS little-endian =0x160 big-endian
        R4000 = 0x0166, // MIPS little-endian
        R10000 = 0x0168, // MIPS little-endian
        WCEMIPSV2 = 0x0169, // MIPS little-endian WCE v2
        ALPHA = 0x0184, // Alpha_AXP
        SH3 = 0x01a2, // SH3 little-endian
        SH3DSP = 0x01a3,
        SH3E = 0x01a4, // SH3E little-endian
        SH4 = 0x01a6, // SH4 little-endian
        SH5 = 0x01a8, // SH5
        ARM = 0x01c0, // ARM Little-Endian
        THUMB = 0x01c2,
        ARMNT = 0x01c4, // ARM Thumb-2 Little-Endian
        AM33 = 0x01d3,
        POWERPC = 0x01F0, // IBM PowerPC Little-Endian
        POWERPCFP = 0x01f1,
        IA64 = 0x0200, // Intel 64
        MIPS16 = 0x0266, // MIPS
        ALPHA64 = 0x0284, // ALPHA64
        MIPSFPU = 0x0366, // MIPS
        MIPSFPU16 = 0x0466, // MIPS
        AXP64 = ALPHA64,
        TRICORE = 0x0520, // Infineon
        CEF = 0x0CEF,
        EBC = 0x0EBC, // EFI public byte Code
        AMD64 = 0x8664, // AMD64 (K8)
        M32R = 0x9041, // M32R little-endian
        CEE = 0xC0EE
    }
        
    [Flags]
    public enum IMAGE_FILE_CHARACTERISTICS : ushort
    {
        IMAGE_RELOCS_STRIPPED = 0x0001, // Relocation info stripped from file.
        IMAGE_EXECUTABLE_IMAGE = 0x0002, // File is executable (i.e. no unresolved external references).
        IMAGE_LINE_NUMS_STRIPPED = 0x0004, // Line nunbers stripped from file.
        IMAGE_LOCAL_SYMS_STRIPPED = 0x0008, // Local symbols stripped from file.
        IMAGE_AGGRESIVE_WS_TRIM = 0x0010, // Agressively trim working set
        IMAGE_LARGE_ADDRESS_AWARE = 0x0020, // App can handle >2gb addresses
        IMAGE_REVERSED_LO = 0x0080, // public bytes of machine public ushort are reversed.
        IMAGE_32BIT_MACHINE = 0x0100, // 32 bit public ushort machine.
        IMAGE_DEBUG_STRIPPED = 0x0200, // Debugging info stripped from file in .DBG file
        IMAGE_REMOVABLE_RUN_FROM_SWAP = 0x0400, // If Image is on removable media =copy and run from the swap file.
        IMAGE_NET_RUN_FROM_SWAP = 0x0800, // If Image is on Net =copy and run from the swap file.
        IMAGE_SYSTEM = 0x1000, // System File.
        IMAGE_DLL = 0x2000, // File is a DLL.
        IMAGE_UP_SYSTEM_ONLY = 0x4000, // File should only be run on a UP machine
        IMAGE_REVERSED_HI = 0x8000 // public bytes of machine public ushort are reversed.
    }
        
    [Flags]
    public enum IMAGE_NT_OPTIONAL_HDR_MAGIC : ushort
    {
        PE32 = 0x10b,
        PE64 = 0x20b
    }
        
    [Flags]
    public enum IMAGE_SUBSYSTEM : ushort
    {
        UNKNOWN = 0, // Unknown subsystem.
        NATIVE = 1, // Image doesn't require a subsystem.
        WINDOWS_GUI = 2, // Image runs in the Windows GUI subsystem.
        WINDOWS_CUI = 3, // Image runs in the Windows character subsystem.
        OS2_CUI = 5, // image runs in the OS/2 character subsystem.
        POSIX_CUI = 7, // image runs in the Posix character subsystem.
        NATIVE_WINDOWS = 8, // image is a native Win9x driver.
        WINDOWS_CE_GUI = 9, // Image runs in the Windows CE subsystem.
        EFI_APPLICATION = 10,
        EFI_BOOT_SERVICE_DRIVER = 11,
        EFI_RUNTIME_DRIVER = 12,
        EFI_ROM = 13,
        XBOX = 14,
        WINDOWS_BOOT_APPLICATION = 16
    }
        
    [Flags]
    public enum IMAGE_DLLCHARACTERISTICS : ushort
    {
        DYNAMIC_BASE = 0x0040, // DLL can move.
        FORCE_INTEGRITY = 0x0080, // Code Integrity Image
        NX_COMPAT = 0x0100, // Image is NX compatible
        NO_ISOLATION = 0x0200, // Image understands isolation and doesn't want it
        NO_SEH = 0x0400, // Image does not use SEH. No SE handler may reside in this image
        NO_BIND = 0x0800, // Do not bind this image.
        WDM_DRIVER = 0x2000, // Driver uses WDM model
        TERMINAL_SERVER_AWARE = 0x8000
    }
        
    [Flags]
    public enum IMAGE_SCN : uint
    {
        TYPE_NO_PAD = 0x00000008, // Reserved.
        CNT_CODE = 0x00000020, // Section contains code.
        CNT_INITIALIZED_DATA = 0x00000040, // Section contains initialized data.
        CNT_UNINITIALIZED_DATA = 0x00000080, // Section contains uninitialized data.
        LNK_INFO = 0x00000200, // Section contains comments or some other type of information.
        LNK_REMOVE = 0x00000800, // Section contents will not become part of image.
        LNK_COMDAT = 0x00001000, // Section contents comdat.
        NO_DEFER_SPEC_EXC = 0x00004000, // Reset speculative exceptions handling bits in the TLB entries for this section.
        GPREL = 0x00008000, // Section content can be accessed relative to GP
        MEM_FARDATA = 0x00008000,
        MEM_PURGEABLE = 0x00020000,
        MEM_16BIT = 0x00020000,
        MEM_LOCKED = 0x00040000,
        MEM_PRELOAD = 0x00080000,
        ALIGN_1BYTES = 0x00100000,
        ALIGN_2BYTES = 0x00200000,
        ALIGN_4BYTES = 0x00300000,
        ALIGN_8BYTES = 0x00400000,
        ALIGN_16BYTES = 0x00500000, // Default alignment if no others are specified.
        ALIGN_32BYTES = 0x00600000,
        ALIGN_64BYTES = 0x00700000,
        ALIGN_128BYTES = 0x00800000,
        ALIGN_256BYTES = 0x00900000,
        ALIGN_512BYTES = 0x00A00000,
        ALIGN_1024BYTES = 0x00B00000,
        ALIGN_2048BYTES = 0x00C00000,
        ALIGN_4096BYTES = 0x00D00000,
        ALIGN_8192BYTES = 0x00E00000,
        ALIGN_MASK = 0x00F00000,
        LNK_NRELOC_OVFL = 0x01000000, // Section contains extended relocations.
        MEM_DISCARDABLE = 0x02000000, // Section can be discarded.
        MEM_NOT_CACHED = 0x04000000, // Section is not cachable.
        MEM_NOT_PAGED = 0x08000000, // Section is not pageable.
        MEM_SHARED = 0x10000000, // Section is shareable.
        MEM_EXECUTE = 0x20000000, // Section is executable.
        MEM_READ = 0x40000000, // Section is readable.
        MEM_WRITE = 0x80000000 // Section is writeable.
    }
    
    [StructLayout(LayoutKind.Sequential, Pack=1)]
    public struct _IMAGE_DOS_HEADER
    {
        public IMAGE_DOS_SIGNATURE e_magic; // Magic number
        public ushort e_cblp; // public bytes on last page of file
        public ushort e_cp; // Pages in file
        public ushort e_crlc; // Relocations
        public ushort e_cparhdr; // Size of header in paragraphs
        public ushort e_minalloc; // Minimum extra paragraphs needed
        public ushort e_maxalloc; // Maximum extra paragraphs needed
        public ushort e_ss; // Initial (relative) SS value
        public ushort e_sp; // Initial SP value
        public ushort e_csum; // Checksum
        public ushort e_ip; // Initial IP value
        public ushort e_cs; // Initial (relative) CS value
        public ushort e_lfarlc; // File address of relocation table
        public ushort e_ovno; // Overlay number
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 8)]
        public string e_res; // This will contain 'Detours!' if patched in memory
        public ushort e_oemid; // OEM identifier (for e_oeminfo)
        public ushort e_oeminfo; // OEM information; e_oemid specific
        [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst=10)] // , ArraySubType=UnmanagedType.U4
        public ushort[] e_res2; // Reserved public ushorts
        public int e_lfanew; // File address of new exe header
    }
        
    [StructLayout(LayoutKind.Sequential, Pack=1)]
    public struct _IMAGE_FILE_HEADER
    {
        public IMAGE_FILE_MACHINE Machine;
        public ushort NumberOfSections;
        public uint TimeDateStamp;
        public uint PointerToSymbolTable;
        public uint NumberOfSymbols;
        public ushort SizeOfOptionalHeader;
        public IMAGE_FILE_CHARACTERISTICS Characteristics;
    }
        
    [StructLayout(LayoutKind.Sequential, Pack=1)]
    public struct _IMAGE_NT_HEADERS32
    {
        public IMAGE_NT_SIGNATURE Signature;
        public _IMAGE_FILE_HEADER FileHeader;
        public _IMAGE_OPTIONAL_HEADER32 OptionalHeader;
    }
        
    [StructLayout(LayoutKind.Sequential, Pack=1)]
    public struct _IMAGE_NT_HEADERS64
    {
        public IMAGE_NT_SIGNATURE Signature;
        public _IMAGE_FILE_HEADER FileHeader;
        public _IMAGE_OPTIONAL_HEADER64 OptionalHeader;
    }
        
    [StructLayout(LayoutKind.Sequential, Pack=1)]
    public struct _IMAGE_OPTIONAL_HEADER32
    {
        public IMAGE_NT_OPTIONAL_HDR_MAGIC Magic;
        public byte MajorLinkerVersion;
        public byte MinorLinkerVersion;
        public uint SizeOfCode;
        public uint SizeOfInitializedData;
        public uint SizeOfUninitializedData;
        public uint AddressOfEntryPoint;
        public uint BaseOfCode;
        public uint BaseOfData;
        public uint ImageBase;
        public uint SectionAlignment;
        public uint FileAlignment;
        public ushort MajorOperatingSystemVersion;
        public ushort MinorOperatingSystemVersion;
        public ushort MajorImageVersion;
        public ushort MinorImageVersion;
        public ushort MajorSubsystemVersion;
        public ushort MinorSubsystemVersion;
        public uint Win32VersionValue;
        public uint SizeOfImage;
        public uint SizeOfHeaders;
        public uint CheckSum;
        public IMAGE_SUBSYSTEM Subsystem;
        public IMAGE_DLLCHARACTERISTICS DllCharacteristics;
        public uint SizeOfStackReserve;
        public uint SizeOfStackCommit;
        public uint SizeOfHeapReserve;
        public uint SizeOfHeapCommit;
        public uint LoaderFlags;
        public uint NumberOfRvaAndSizes;
        [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst=16)]
        public _IMAGE_DATA_DIRECTORY[] DataDirectory;
    }
        
    [StructLayout(LayoutKind.Sequential, Pack=1)]
    public struct _IMAGE_OPTIONAL_HEADER64
    {
        public IMAGE_NT_OPTIONAL_HDR_MAGIC Magic;
        public byte MajorLinkerVersion;
        public byte MinorLinkerVersion;
        public uint SizeOfCode;
        public uint SizeOfInitializedData;
        public uint SizeOfUninitializedData;
        public uint AddressOfEntryPoint;
        public uint BaseOfCode;
        public ulong ImageBase;
        public uint SectionAlignment;
        public uint FileAlignment;
        public ushort MajorOperatingSystemVersion;
        public ushort MinorOperatingSystemVersion;
        public ushort MajorImageVersion;
        public ushort MinorImageVersion;
        public ushort MajorSubsystemVersion;
        public ushort MinorSubsystemVersion;
        public uint Win32VersionValue;
        public uint SizeOfImage;
        public uint SizeOfHeaders;
        public uint CheckSum;
        public IMAGE_SUBSYSTEM Subsystem;
        public IMAGE_DLLCHARACTERISTICS DllCharacteristics;
        public ulong SizeOfStackReserve;
        public ulong SizeOfStackCommit;
        public ulong SizeOfHeapReserve;
        public ulong SizeOfHeapCommit;
        public uint LoaderFlags;
        public uint NumberOfRvaAndSizes;
        [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst=16)]
        public _IMAGE_DATA_DIRECTORY[] DataDirectory;
    }
        
    [StructLayout(LayoutKind.Sequential, Pack=1)]
    public struct _IMAGE_DATA_DIRECTORY
    {
        public uint VirtualAddress;
        public uint Size;
    }
        
    [StructLayout(LayoutKind.Sequential, Pack=1)]
    public struct _IMAGE_EXPORT_DIRECTORY
    {
        public uint Characteristics;
        public uint TimeDateStamp;
        public ushort MajorVersion;
        public ushort MinorVersion;
        public uint Name;
        public uint Base;
        public uint NumberOfFunctions;
        public uint NumberOfNames;
        public uint AddressOfFunctions; // RVA from base of image
        public uint AddressOfNames; // RVA from base of image
        public uint AddressOfNameOrdinals; // RVA from base of image
    }
       
    [StructLayout(LayoutKind.Sequential, Pack=1)]
    public struct _IMAGE_SECTION_HEADER
    {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 8)]
        public string Name;
        public uint VirtualSize;
        public uint VirtualAddress;
        public uint SizeOfRawData;
        public uint PointerToRawData;
        public uint PointerToRelocations;
        public uint PointerToLinenumbers;
        public ushort NumberOfRelocations;
        public ushort NumberOfLinenumbers;
        public IMAGE_SCN Characteristics;
    }
        
    [StructLayout(LayoutKind.Sequential, Pack=1)]
    public struct _IMAGE_IMPORT_DESCRIPTOR
    {
        public uint OriginalFirstThunk; // RVA to original unbound IAT (PIMAGE_THUNK_DATA)
        public uint TimeDateStamp; // 0 if not bound,
                                            // -1 if bound, and real date/time stamp
                                            // in IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT (new BIND)
                                            // O.W. date/time stamp of DLL bound to (Old BIND)
        public uint ForwarderChain; // -1 if no forwarders
        public uint Name;
        public uint FirstThunk; // RVA to IAT (if bound this IAT has actual addresses)
    }

    [StructLayout(LayoutKind.Sequential, Pack=1)]
    public struct _IMAGE_THUNK_DATA32
    {
        public Int32 AddressOfData; // PIMAGE_IMPORT_BY_NAME
    }

    [StructLayout(LayoutKind.Sequential, Pack=1)]
    public struct _IMAGE_THUNK_DATA64
    {
        public Int64 AddressOfData; // PIMAGE_IMPORT_BY_NAME
    }
        
    [StructLayout(LayoutKind.Sequential, Pack=1)]
    public struct _IMAGE_IMPORT_BY_NAME
    {
        public ushort Hint;
        public char Name;
    }
}
"@

$compileParams = New-Object System.CodeDom.Compiler.CompilerParameters
$compileParams.ReferencedAssemblies.AddRange(@('System.dll', 'mscorlib.dll'))
$compileParams.GenerateInMemory = $True
Add-Type -TypeDefinition $code -CompilerParameters $compileParams -PassThru -WarningAction SilentlyContinue | Out-Null
}
function Convert-RVAToFileOffset([int]$Rva, [PSObject[]]$SectionHeaders) {
    foreach ($Section in $SectionHeaders) {
        if ($Rva -ge $Section.VirtualAddress -and $Rva -lt ($Section.VirtualAddress + $Section.VirtualSize)) {
            return $Rva - $Section.VirtualAddress + $Section.PointerToRawData
        }
    }
    return $Rva
}

    $DllPath = Join-Path -Path $env:windir -ChildPath "System32\$DllName"

    if (-not (Test-Path $DllPath)) {
        Write-Error "DLL file not found at: $DllPath"
        return $null
    }

    $FileByteArray = [System.IO.File]::ReadAllBytes($DllPath)
    $Handle = [GCHandle]::Alloc($FileByteArray, 'Pinned')
    $PEBaseAddr = $Handle.AddrOfPinnedObject()

    try {
        # Parse DOS header
        $DosHeader = [Marshal]::PtrToStructure($PEBaseAddr, [Type] [PE+_IMAGE_DOS_HEADER])
        $PointerNtHeader = [IntPtr]($PEBaseAddr.ToInt64() + $DosHeader.e_lfanew)

        # Detect architecture
        $NtHeader32 = [Marshal]::PtrToStructure($PointerNtHeader, [Type] [PE+_IMAGE_NT_HEADERS32])
        $Architecture = $NtHeader32.FileHeader.Machine.ToString()
        $PEStruct = @{}

        if ($Architecture -eq 'AMD64') {
            $PEStruct['NT_HEADER'] = [PE+_IMAGE_NT_HEADERS64]
        } elseif ($Architecture -eq 'I386') {
            $PEStruct['NT_HEADER'] = [PE+_IMAGE_NT_HEADERS32]
        } else {
            Write-Error "Unsupported architecture: $Architecture"
            return $null
        }

        # Parse correct NT header
        $NtHeader = [Marshal]::PtrToStructure($PointerNtHeader, [Type] $PEStruct['NT_HEADER'])
        $NumSections = $NtHeader.FileHeader.NumberOfSections

        # Parse section headers
        $PointerSectionHeader = [IntPtr] ($PointerNtHeader.ToInt64() + [Marshal]::SizeOf([Type] $PEStruct['NT_HEADER']))
        $SectionHeaders = New-Object PSObject[]($NumSections)
        for ($i = 0; $i -lt $NumSections; $i++) {
            $SectionHeaders[$i] = [Marshal]::PtrToStructure(
                [IntPtr]($PointerSectionHeader.ToInt64() + ($i * [Marshal]::SizeOf([Type] [PE+_IMAGE_SECTION_HEADER]))),
                [Type] [PE+_IMAGE_SECTION_HEADER]
            )
        }

        # Check for exports
        if ($NtHeader.OptionalHeader.DataDirectory[0].VirtualAddress -eq 0) {
            Write-Error "Module does not contain any exports."
            return $null
        }

        # Get Export Directory
        $ExportDirRVA = $NtHeader.OptionalHeader.DataDirectory[0].VirtualAddress
        $ExportDirOffset = Convert-RVAToFileOffset -Rva $ExportDirRVA -SectionHeaders $SectionHeaders
        $ExportDirectory = [Marshal]::PtrToStructure(
            [IntPtr]($PEBaseAddr.ToInt64() + $ExportDirOffset),
            [Type] [PE+_IMAGE_EXPORT_DIRECTORY]
        )

        # Export table pointers
        $AddressOfNamesOffset        = Convert-RVAToFileOffset -Rva $ExportDirectory.AddressOfNames        -SectionHeaders $SectionHeaders
        $AddressOfNameOrdinalsOffset = Convert-RVAToFileOffset -Rva $ExportDirectory.AddressOfNameOrdinals -SectionHeaders $SectionHeaders
        $AddressOfFunctionsOffset    = Convert-RVAToFileOffset -Rva $ExportDirectory.AddressOfFunctions    -SectionHeaders $SectionHeaders

        # Loop through exported names to find the function
        for ($i = 0; $i -lt $ExportDirectory.NumberOfNames; $i++) {
            $nameRVA = [Marshal]::ReadInt32([IntPtr]($PEBaseAddr.ToInt64() + $AddressOfNamesOffset + ($i * 4)))
            $funcNameOffset = Convert-RVAToFileOffset -Rva $nameRVA -SectionHeaders $SectionHeaders
            $funcName = [Marshal]::PtrToStringAnsi([IntPtr]($PEBaseAddr.ToInt64() + $funcNameOffset))

            if ($funcName -eq $FunctionName) {
                $ordinal = [Marshal]::ReadInt16([IntPtr]($PEBaseAddr.ToInt64() + $AddressOfNameOrdinalsOffset + ($i * 2)))
                $funcRVA = [Marshal]::ReadInt32([IntPtr]($PEBaseAddr.ToInt64() + $AddressOfFunctionsOffset + ($ordinal * 4)))

                # Skip forwarded exports
                if ($funcRVA -ge $ExportDirRVA -and $funcRVA -lt ($ExportDirRVA + $NtHeader.OptionalHeader.DataDirectory[0].Size)) {
                    Write-Error "Function '$FunctionName' is a forwarded export and cannot be read."
                    return $null
                }

                # Get file offset and extract bytes
                $funcFileOffset = Convert-RVAToFileOffset -Rva $funcRVA -SectionHeaders $SectionHeaders
                if ($funcFileOffset -ge $FileByteArray.Length) {
                    Write-Error "Function RVA points outside the file. Cannot read bytes."
                    return $null
                }

                $bytesAvailable = $FileByteArray.Length - $funcFileOffset
                if ($BytesToRead -gt $bytesAvailable) {
                    $BytesToRead = $bytesAvailable
                    Write-Warning "Read would go beyond file size. Reading to end of file ($BytesToRead bytes)."
                }

                $funcBytes = $FileByteArray[$funcFileOffset..($funcFileOffset + $BytesToRead - 1)]
                return $funcBytes
            }
        }

        Write-Error "Function '$FunctionName' not found in DLL."
        return $null
    } finally {
        $Handle.Free()
    }
}

<#

     *********************

     !Managed Com Warper.!
      -  Example code. -

     *********************

# netlistmgr.h
# https://raw.githubusercontent.com/nihon-tc/Rtest/refs/heads/master/header/Microsoft%20SDKs/Windows/v7.0A/Include/netlistmgr.h

# get_IsConnectedToInternet 
# https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/nf-netlistmgr-inetworklistmanager-get_isconnectedtointernet

-------------------------------

Clear-host
write-host "`n`nCLSID & Propertie's [Test]`nDCB00C01-570F-4A9B-8D69-199FDBA5723B->Default->IsConnected,IsConnectedToInternet`n"
$NetObj = "DCB00C01-570F-4A9B-8D69-199FDBA5723B" | Initialize-ComObject
write-host "IsConnected: $($NetObj.IsConnected)"
write-host "IsConnectedToInternet: $($NetObj.IsConnectedToInternet)"
$NetObj | Release-ComObject

-------------------------------

Clear-host
write-host "`n`nIEnumerator & Params\values [Test]`nDCB00C01-570F-4A9B-8D69-199FDBA5723B->DCB00000-570F-4A9B-8D69-199FDBA5723B->GetNetwork`n"
[intPtr]$ppEnumNetwork = [intPtr]::Zero
Use-ComInterface `
    -CLSID "DCB00C01-570F-4A9B-8D69-199FDBA5723B" `
    -IID "DCB00000-570F-4A9B-8D69-199FDBA5723B" `
    -Index 1 `
    -Name "GetNetwork" `
    -Return "uint" `
    -Params 'system.UINT32 Flags, out INTPTR ppEnumNetwork' `
    -Values @(1, [ref]$ppEnumNetwork)

if ($ppEnumNetwork -ne [IntPtr]::Zero) {
    $networkList = $ppEnumNetwork | Receive-ComObject
    foreach ($network in $networkList) {
        "Name: $($network.GetName()), IsConnected: $($network.IsConnected())"
    }
    $networkList | Release-ComObject
}

-------------------------------

Clear-host
write-host "`n`nVoid & No Return [Test]`n17CCA47D-DAE5-4E4A-AC42-CC54E28F334A->f2dcb80d-0670-44bc-9002-cd18688730af->ShowProductKeyUI`n"
Use-ComInterface `
    -CLSID "17CCA47D-DAE5-4E4A-AC42-CC54E28F334A" `
    -IID "f2dcb80d-0670-44bc-9002-cd18688730af" `
    -Index 3 `
    -Name "ShowProductKeyUI" `
    -Return "void"

-------------------------------

Clear-host
"ApiMajorVersion", "ApiMinorVersion", "ProductVersionString" | ForEach-Object {
    $name = $_
    $outVarPtr = New-Variant -Type VT_EMPTY
    $inVarPtr  = New-Variant -Type VT_BSTR -Value $name
    try {
        $ret = Use-ComInterface `
            -CLSID "C2E88C2F-6F5B-4AAA-894B-55C847AD3A2D" `
            -IID "85713fa1-7796-4fa2-be3b-e2d6124dd373" `
            -Index 1 -Name "GetInfo" `
            -Values @($inVarPtr, $outVarPtr) `
            -Type IDispatch

        if ($ret -eq 0) {
            $value = Parse-Variant -variantPtr $outVarPtr
            Write-Host "$name -> $value"
        }

    } finally {
        Free-IntPtr -handle $inVarPtr  -Method VARIANT
        Free-IntPtr -handle $outVarPtr -Method VARIANT
    }
}

#>
function Build-ComInterfaceSpec {
    param (
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [ValidatePattern('^[A-F0-9]{8}-([A-F0-9]{4}-){3}[A-F0-9]{12}$')]
        [string]$CLSID,

        [Parameter(Position = 2)]
        [string]$IID,

        [Parameter(Mandatory = $true, Position = 3)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$Index,

        [Parameter(Mandatory = $true, Position = 4)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory = $true, Position = 5)]
        [ValidateSet(
            # Void
            "void",

            # Fully qualified .NET types
            "system.boolean", "system.byte", "system.char", "system.decimal", "system.double",
            "system.int16", "system.int32", "system.int64", "system.intptr", "system.object",
            "system.sbyte", "system.single", "system.string", "system.uint16", "system.uint32",
            "system.uint64", "system.uintptr",

            # Alternate type spellings and aliases
            "boolean", "dword32", "dword64", "int16", "int32", "int64", "single", "uint16",
            "uint32", "uint64",

            # Additional C/C++ & WinAPI aliases
            "double", "float", "long", "longlong", "tchar", "uchar", "ulong", "ulonglong",
            "short", "ushort",

            # Additional typedefs
            "atom", "dword_ptr", "dwordlong", "farproc", "hhook", "hresult", "NTSTATUS",
            "int_ptr", "intptr_t", "long_ptr", "lpbyte", "lpdword", "lparam", "pcstr",
            "pcwstr", "pstr", "pwstr", "uint_ptr", "uintptr_t", "wparam",

            # C# built-in types
            "bool", "byte", "char", "decimal", "int", "intptr", "nint", "nuint", "object",
            "sbyte", "string", "uint", "uintptr",

            # Common WinAPI handle types
            "hbitmap", "hbrush", "hcurs", "hdc", "hfont", "hicon", "hmenu", "hpen", "hrgn",

            # Pointer-based aliases
            "pbyte", "pchar", "pdword", "pint", "plong", "puint", "pulong", "pvoid", "lpvoid",

            # Special types
            "guid",

            # Windows/WinAPI types (common aliases)
            "dword", "handle", "hinstance", "hmodule", "hwnd", "lpcstr", "lpcwstr", "lpstr",
            "lpwstr", "ptr", "size_t", "ssize_t", "void*", "word", "phandle", "lresult",

            # STRSAFE typedefs
            "strsafe_lpstr", "strsafe_lpcstr", "strsafe_lpwstr", "strsafe_lpcwstr",
            "strsafe_lpcuwstr", "strsafe_pcnzch", "strsafe_pcnzwch", "strsafe_pcunzwch",

            # Wide-character (Unicode) types
            "pstring", "pwchar", "lpwchar", "pczpwstr", "pzpwstr", "pzwstr", "pzzwstr",
            "pczzwstr", "puczzwstr", "pcuczzwstr", "pnzwch", "pcnzwch", "punzwch", "pcunzwch",

            # ANSI string types
            "npstr", "pzpcstr", "pczpcstr", "pzzstr", "pczzstr", "pnzch", "pcnzch",

            # UCS types
            "ucschar", "pucschar", "pcucschar", "puucschar", "pcuucschar", "pucsstr",
            "pcucsstr", "puucsstr", "pcuucsstr",

            # Neutral ANSI/Unicode (TCHAR-based) Types
            "ptchar", "tbyte", "ptbyte", "ptstr", "lptstr", "pctstr", "lpctstr", "putstr",
            "lputstr", "pcutstr", "lpcutstr", "pzptstr", "pzzstr", "pczztstr", "pzzwstr", "pczzwstr"
        )]
        [string]$Return,
        
        [Parameter(Position = 6)]
        [string]$Params,

        [Parameter(Position = 7)]
        [string]$InterFaceType,

        [Parameter(Position = 8)]
        [string]$CharSet
    )

    if (-not [string]::IsNullOrEmpty($IID)) {
        if (-not [regex]::Match($IID,'^[A-F0-9]{8}-([A-F0-9]{4}-){3}[A-F0-9]{12}$')){
            throw "ERROR: $IID not match ^[A-F0-9]{8}-([A-F0-9]{4}-){3}[A-F0-9]{12}$"
        }
    }

    # Create and return the interface specification object
    $interfaceSpec = [PSCustomObject]@{
        Index   = $Index
        Return  = $Return
        Name    = $Name
        Params  = if ($Params) { $Params } else { "" }
        CLSID   = $CLSID
        IID     = if ($IID) { $IID } else { "" }
        Type    = if ($InterFaceType) { $InterFaceType } else { "" }
        CharSet = $CharSet
    }

    return $interfaceSpec
}
function Build-ComDelegate {
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline)]
        [PSCustomObject]$InterfaceSpec,

        [Parameter(Mandatory=$true)]
        [string]$UNIQUE_ID
    )

    # External function calls for Params and ReturnType
    $Params = Process-Parameters -InterfaceSpec $InterfaceSpec
    $fixedReturnType = Process-ReturnType -ReturnType $InterfaceSpec.Return
    $charSet = if ($InterfaceSpec.CharSet) { "CharSet = CharSet.$($InterfaceSpec.CharSet)" } else { "CharSet = CharSet.Unicode" }

    # Construct the delegate code template
    $Return = @"
    [UnmanagedFunctionPointer(CallingConvention.StdCall, $charSet)]
    public delegate $($fixedReturnType) $($UNIQUE_ID)(
        $($Params)
    );
"@

    # Define the C# namespace and using statements
    $namespace = "namespace DynamicDelegates"
    $using = "`nusing System;`nusing System.Runtime.InteropServices;`n"
    
    # Combine all parts to form the final C# code
    return "$using`n$namespace`n{`n$Return`n}`n"
}
function Initialize-ComObject {
    param (
        [Parameter(ValueFromPipeline, Position = 0)]
        [PSCustomObject]$InterfaceSpec,

        [Parameter(ValueFromPipeline, Position = 1)]
        [ValidatePattern('^[A-F0-9]{8}-([A-F0-9]{4}-){3}[A-F0-9]{12}$')]
        [GUID]$CLSID,

        [switch]
        $CreateInstance
    )

    if ($CLSID -and $InterfaceSpec -and $CLSID.ToString() -eq $InterfaceSpec) {
        $InterfaceSpec = $null
    }

    # Oppsite XOR Case, Validate it Not both
    if (-not ([bool]$InterfaceSpec -xor [bool]$CLSID)) {
        throw "Select CLSID OR $InterfaceSpec"
    }

    # ------ BASIC SETUP -------

    if ($InterfaceSpec) {
        $CLSID = [guid]$InterfaceSpec.CLSID
    }

    $comObj = [Activator]::CreateInstance([type]::GetTypeFromCLSID($clsid))
    if (-not $comObj) {
        throw "Failed to create COM object for CLSID $clsid"
    }

    if (-not $InterfaceSpec -or $CreateInstance)  {
       return $comObj
    }

    $iid = if ($InterfaceSpec.IID) {
        [guid]$InterfaceSpec.IID
    } else {
        [guid]"00000000-0000-0000-C000-000000000046"
    }

    # ------ QueryInterface Delegate -------
    
    try {
     [QueryInterfaceDelegate] | Out-Null
    }
    catch {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

[UnmanagedFunctionPointer(CallingConvention.StdCall)]
public delegate int QueryInterfaceDelegate(IntPtr thisPtr, ref Guid riid, out IntPtr ppvObject);
"@ -Language CSharp -ErrorAction Stop
    }

    $iUnknownPtr = [Marshal]::GetIUnknownForObject($comObj)
    $queryInterfacePtr = [Marshal]::ReadIntPtr(
        [Marshal]::ReadIntPtr($iUnknownPtr))

    $queryInterface = [Marshal]::GetDelegateForFunctionPointer(
        $queryInterfacePtr, [QueryInterfaceDelegate])

    # ------ Continue with IID Setup -------

    $interfacePtr = [IntPtr]::Zero
    $hresult = $queryInterface.Invoke($iUnknownPtr, [ref]$iid, [ref]$interfacePtr)
    if ($hresult -ne 0 -or $interfacePtr -eq [IntPtr]::Zero) {
        throw "QueryInterface failed with HRESULT 0x{0:X8}" -f $hresult
    }
    $requestedVTablePtr = [Marshal]::ReadIntPtr($interfacePtr)

    # ------ Check if inherit *** [can fail, or misled, be aware!] -------

    $interfaces = @(
        @("00020400-0000-0000-C000-000000000046", 7,  "IDispatch"),                # 1
        @("00000003-0000-0000-C000-000000000046", 9,  "IMarshal"),                 # 2
        @("00000118-0000-0000-C000-000000000046", 8,  "IOleClientSite"),           # 3
        @("00000112-0000-0000-C000-000000000046", 24, "IOleObject"),               # 4
        @("0000010B-0000-0000-C000-000000000046", 8,  "IPersistFile"),             # 5
        @("0000010C-0000-0000-C000-000000000046", 4,  "IPersist"),                 # 6
        @("00000109-0000-0000-C000-000000000046", 7,  "IPersistStream"),           # 7
        @("0000010E-0000-0000-C000-000000000046", 12, "IDataObject"),              # 8
        @("0000000C-0000-0000-C000-000000000046", 13, "IStream"),                  # 9
        @("0000000B-0000-0000-C000-000000000046", 15, "IStorage"),                 # 10
        @("0000010A-0000-0000-C000-000000000046", 11, "IPersistStorage"),          # 11
        @("00000139-0000-0000-C000-000000000046", 7,  "IEnumSTATPROPSTG"),         # 12
        @("0000013A-0000-0000-C000-000000000046", 7,  "IEnumSTATPROPSETSTG"),      # 13
        @("0000000D-0000-0000-C000-000000000046", 7,  "IEnumSTATSTG"),             # 14
        @("00020404-0000-0000-C000-000000000046", 7,  "IEnumVARIANT"),             # 15
        @("00000102-0000-0000-C000-000000000046", 7,  "IEnumMoniker"),             # 16
        @("00000101-0000-0000-C000-000000000046", 7,  "IEnumString"),              # 17
        @("B196B286-BAB4-101A-B69C-00AA00341D07", 7,  "IConnectionPoint"),         # 18
        @("55272A00-42CB-11CE-8135-00AA004BB851", 5,  "IPropertyBag"),             # 19
        @("00000114-0000-0000-C000-000000000046", 5,  "IOleWindow"),               # 20
        @("B196B283-BAB4-101A-B69C-00AA00341D07", 4,  "IProvideClassInfo"),        # 21
        @("A6BC3AC0-DBAA-11CE-9DE3-00AA004BB851", 4,  "IProvideClassInfo2"),       # 22
        @("B196B28B-BAB4-101A-B69C-00AA00341D07", 4,  "ISpecifyPropertyPages"),    # 23
        @("EB5E0020-8F75-11D1-ACDD-00C04FC2B085", 4,  "IPersistPropertyBag"),      # 24
        @("B196B284-BAB4-101A-B69C-00AA00341D07", 4,  "IConnectionPointContainer") # 25
    )

    $baseMethodOffset = 0
    if ($InterfaceSpec.Type -and (-not [string]::IsNullOrEmpty($InterfaceSpec.Type))) {
        $interface = $interfaces | ? { $_[2] -eq $InterfaceSpec.Type }
        if ($interface) {
            $baseMethodOffset = $interface[1]
    }}

    if ($baseMethodOffset -eq 0) {
        $baseMethodOffset = 3

        foreach ($iface in $interfaces) {
            $iid = $iface[0]
            $totalMethods = $iface[1]
            $ptr = [IntPtr]::Zero

            $hr = $queryInterface.Invoke($interfacePtr, [ref]$iid, [ref]$ptr)
            if ($hr -eq 0 -and $ptr -ne [IntPtr]::Zero) {
                $baseMethodOffset = $totalMethods
                [Marshal]::Release($ptr) | Out-Null
                break
    }}}

    # ------ Continue with IID Setup -------

    $timestampSuffix = (Get-Date -Format "yyyyMMddHHmmssfff")
    $simpleUniqueDelegateName = "$($InterfaceSpec.Name)$timestampSuffix"
    $delegateCode = Build-ComDelegate -InterfaceSpec $InterfaceSpec -UNIQUE_ID $simpleUniqueDelegateName
    Add-Type -TypeDefinition $delegateCode -Language CSharp -ErrorAction Stop

    $delegateType = $null
    $fullDelegateTypeName = "DynamicDelegates.$simpleUniqueDelegateName"
    $delegateType = [AppDomain]::CurrentDomain.GetAssemblies() |
        ForEach-Object { $_.GetType($fullDelegateTypeName, $false, $true) } |
        Where-Object { $_ } |
        Select-Object -First 1

    if (-not $delegateType) {
        throw "Delegate type '$simpleUniqueDelegateName' not found."
    }

    $methodIndex = $baseMethodOffset + ([int]$InterfaceSpec.Index - 1)
    $funcPtr = [Marshal]::ReadIntPtr($requestedVTablePtr, $methodIndex * [IntPtr]::Size)
    $delegateInstance = [Marshal]::GetDelegateForFunctionPointer($funcPtr, $delegateType)

    return [PSCustomObject]@{
        ComObject        = $comObj
        IUnknownPtr      = $iUnknownPtr
        InterfacePtr     = $interfacePtr
        VTablePtr        = $requestedVTablePtr
        FunctionPtr      = $funcPtr
        DelegateType     = $delegateType
        DelegateInstance = $delegateInstance
        InterfaceSpec    = $InterfaceSpec
        MethodIndex      = $methodIndex
        DelegateCode     = $delegateCode
    }
}
function Receive-ComObject {
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline)]
        [object]$punk
    )

    try {
        return [Marshal]::GetObjectForIUnknown([intPtr]$punk)
    }
    catch {
        return $punk
    }
}
function Release-ComObject {
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline)]
        $comInterface
    )

    $ISComObject = $comInterface.GetType().Name -match '__ComObject'
    $IsPSCustomObject = $comInterface.GetType().Name -match 'PSCustomObject'

    if ($ISComObject) {
        [Marshal]::ReleaseComObject($comInterface) | Out-Null
    }
    if ($IsPSCustomObject) {
        try {
            if ($comInterface.ComObject) {
                [Marshal]::ReleaseComObject($comInterface.ComObject) | Out-Null
            }
        } catch {}

        try {
            if ($comInterface.IUnknownPtr -and $comInterface.IUnknownPtr -ne [IntPtr]::Zero) {
                [Marshal]::Release($comInterface.IUnknownPtr) | Out-Null
            }
        } catch {}

        try {
            if ($comInterface.InterfacePtr -and $comInterface.InterfacePtr -ne [IntPtr]::Zero) {
                [Marshal]::Release($comInterface.InterfacePtr) | Out-Null
            }
        } catch {}

        # Cleanup
        $comInterface.ComObject        = $null
        $comInterface.DelegateInstance = $null
        $comInterface.VTablePtr        = $null
        $comInterface.FunctionPtr      = $null
        $comInterface.DelegateType     = $null
        $comInterface.InterfaceSpec    = $null
        $comInterface.IUnknownPtr      = [IntPtr]::Zero
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
function Use-ComInterface {
    param (
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [ValidatePattern('^[A-F0-9]{8}-([A-F0-9]{4}-){3}[A-F0-9]{12}$')]
        [string]$CLSID,

        [Parameter(Position = 2)]
        [ValidatePattern('^[A-F0-9]{8}-([A-F0-9]{4}-){3}[A-F0-9]{12}$')]
        [string]$IID,

        [Parameter(Mandatory = $true, Position = 3)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$Index,

        [Parameter(Mandatory = $true, Position = 4)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory = $false, Position = 5)]
        [ValidateSet(
            # Void
            "void",

            # Fully qualified .NET types
            "system.boolean", "system.byte", "system.char", "system.decimal", "system.double",
            "system.int16", "system.int32", "system.int64", "system.intptr", "system.object",
            "system.sbyte", "system.single", "system.string", "system.uint16", "system.uint32",
            "system.uint64", "system.uintptr",

            # Alternate type spellings and aliases
            "boolean", "dword32", "dword64", "int16", "int32", "int64", "single", "uint16",
            "uint32", "uint64",

            # Additional C/C++ & WinAPI aliases
            "double", "float", "long", "longlong", "tchar", "uchar", "ulong", "ulonglong",
            "short", "ushort",

            # Additional typedefs
            "atom", "dword_ptr", "dwordlong", "farproc", "hhook", "hresult", "NTSTATUS",
            "int_ptr", "intptr_t", "long_ptr", "lpbyte", "lpdword", "lparam", "pcstr",
            "pcwstr", "pstr", "pwstr", "uint_ptr", "uintptr_t", "wparam",

            # C# built-in types
            "bool", "byte", "char", "decimal", "int", "intptr", "nint", "nuint", "object",
            "sbyte", "string", "uint", "uintptr",

            # Common WinAPI handle types
            "hbitmap", "hbrush", "hcurs", "hdc", "hfont", "hicon", "hmenu", "hpen", "hrgn",

            # Pointer-based aliases
            "pbyte", "pchar", "pdword", "pint", "plong", "puint", "pulong", "pvoid", "lpvoid",

            # Special types
            "guid",

            # Windows/WinAPI types (common aliases)
            "dword", "handle", "hinstance", "hmodule", "hwnd", "lpcstr", "lpcwstr", "lpstr",
            "lpwstr", "ptr", "size_t", "ssize_t", "void*", "word", "phandle", "lresult",

            # STRSAFE typedefs
            "strsafe_lpstr", "strsafe_lpcstr", "strsafe_lpwstr", "strsafe_lpcwstr",
            "strsafe_lpcuwstr", "strsafe_pcnzch", "strsafe_pcnzwch", "strsafe_pcunzwch",

            # Wide-character (Unicode) types
            "pstring", "pwchar", "lpwchar", "pczpwstr", "pzpwstr", "pzwstr", "pzzwstr",
            "pczzwstr", "puczzwstr", "pcuczzwstr", "pnzwch", "pcnzwch", "punzwch", "pcunzwch",

            # ANSI string types
            "npstr", "pzpcstr", "pczpcstr", "pzzstr", "pczzstr", "pnzch", "pcnzch",

            # UCS types
            "ucschar", "pucschar", "pcucschar", "puucschar", "pcuucschar", "pucsstr",
            "pcucsstr", "puucsstr", "pcuucsstr",

            # Neutral ANSI/Unicode (TCHAR-based) Types
            "ptchar", "tbyte", "ptbyte", "ptstr", "lptstr", "pctstr", "lpctstr", "putstr",
            "lputstr", "pcutstr", "lpcutstr", "pzptstr", "pzzstr", "pczztstr", "pzzwstr", "pczzwstr"
        )]
        [string]$Return,
        
        [Parameter(Position = 6)]
        [string]$Params,

        [Parameter(Position = 7)]
        [object[]]$Values,

        [Parameter(Position = 8)]
        [ValidateSet(
            "IOleObject", "IDataObject", "IStream", "IPersistStorage", 
            "IStorage", "IMarshal", "IPersistFile", "IOleClientSite", 
            "IDispatch", "IEnumSTATPROPSTG", "IEnumSTATPROPSETSTG", 
            "IEnumSTATSTG", "IPersistStream", "IEnumVARIANT", "IEnumMoniker",
            "IConnectionPoint", "IEnumString", "IOleWindow", "IPropertyBag",
            "IPersist", "IProvideClassInfo", "IProvideClassInfo2", "ISpecifyPropertyPages",
            "IPersistPropertyBag", "IConnectionPointContainer"
        )]
        [string]$Type,

        [Parameter(Mandatory = $false, Position = 9)]
        [ValidateSet("Unicode", "Ansi")]
        [string]$CharSet = "Unicode"
    )

    # Detect platform
    if (-not $CallingConvention) {
        if ([IntPtr]::Size -eq 8) {
            $CallingConvention = "StdCall" 
        }
        else {
            $CallingConvention = "StdCall"
        }
    }

    # Lazy Mode Detection
    $Count = 0
    [void][int]::TryParse($Values.Count,[ref]$count)
    $lazyMode = (-not $Params) -and ($Count -gt 0)
    $IsArrayObj = $Count -eq 1 -and $Values[0] -is [System.Array]

     if (-not $Return) {
        $Return = "Int32"
    }

    if ($IsArrayObj) {
        Write-error "Cast all Params with '-Values @()' Please"
        return
    }

    if ($lazyMode) {
        
        try {
            $idx = 0
            $Params = (
                $Values | % {
                    ++$idx
                    if ($_.Value -or ($_ -is [ref])) {
                        $byRef = 'ref '
                        $Name  = $_.Value.GetType().Name
                    }
                    else {
                        $byRef = ''
                        $Name  = $_.GetType().Name
                    }
                    "{0}{1} {2}" -f $byRef, $Name, (Get-Base26Name -idx $idx)
                }
            ) -join ", "
        }
        catch {
            throw "auto parse params fail"
        }

        $CharSet = if ($Function -like "*A") { "Ansi" } else { "Unicode" }
    }

    $interfaceSpec = Build-ComInterfaceSpec `
        -CLSID $CLSID  `
        -IID $IID  `
        -Index $Index  `
        -Name $Name  `
        -Return $Return  `
        -Params $Params `
        -InterFaceType $Type `
        -CharSet $CharSet

    $comObj = $interfaceSpec | Initialize-ComObject

    try {
        return $comObj | Invoke-Object -Params $Values -type COM
    }
    finally {
        $comObj | Release-ComObject
    }
}

<#

     *********************

     !Managed Api Warper.!
      -  Example code. -

     *********************

Clear-Host
Write-Host

Invoke-UnmanagedMethod `
    -Dll "kernel32.dll" `
    -Function "Beep" `
    -Return "bool" `
    -Params "uint dwFreq, uint dwDuration" `
    -Values @(750, 300)  # 750 Hz beep for 300 ms

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Invoke-UnmanagedMethod `
    -Dll "User32.dll" `
    -Function "MessageBoxA" `
    -Values @(
        [IntPtr]0,
        "Text Block",
        "Text title",
        20,
        [UIntPtr]::new(9),
        1,2,"Alpha",
        ([REF]1),
        ([REF]"1"),
        [Int16]1,
        ([REF][uInt16]2)
    )

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Test Charset <> Ansi
$Func = Register-NativeMethods @(
    @{ 
        Name       = "MessageBoxA"
        Dll        = "user32.dll"
        ReturnType = [int]
        CharSet    = 'Ansi'
        Parameters = [Type[]]@(
            [IntPtr],    # hWnd
            [string],    # lpText
            [string],    # lpCaption
            [uint32]     # uType
        )
    })
$Func::MessageBoxA(
    [IntPtr]::Zero, "Hello from ANSI!", "MessageBoxA", 0)

# Test Charset <> Ansi
Invoke-UnmanagedMethod `
    -Dll "user32.dll" `
    -Function "MessageBoxA" `
    -Return "int32" `
    -Params "HWND hWnd, LPCSTR lpText, LPCSTR lpCaption, UINT uType" `
    -Values @(0, "Hello from ANSI!", "MessageBoxA", 0) `
    -CharSet Ansi

# Test Charset <> Ansi
Invoke-UnmanagedMethod `
    -Dll "User32.dll" `
    -Function "MessageBoxA" `
    -Values @(
        [IntPtr]0,
        "Text Block",
        "Text title",
        20,
        ([ref]1),
        [UintPtr]::new(1),
        ([ref][IntPtr]2),
        ([ref][guid]::Empty)
    )

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Clear-Host
Write-Host
$buffer = New-IntPtr -Size 256
$result = Invoke-UnmanagedMethod `
  -Dll "kernel32.dll" `
  -Function "GetComputerNameA" `
  -Return "bool" `
  -Params "IntPtr lpBuffer, ref uint lpnSize" `
  -Values @($buffer, [ref]256)

if ($result) {
    $computerName = [Marshal]::PtrToStringAnsi($buffer)
    Write-Host "Computer Name: $computerName"
} else {
    Write-Host "Failed to get computer name"
}
New-IntPtr -hHandle $buffer -Release

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# --- For GetComputerNameA (ANSI Version) ---
$computerNameA = New-Object byte[] 250
$handle = [gchandle]::Alloc($computerNameA, 'Pinned')
Invoke-UnmanagedMethod "kernel32.dll" "GetComputerNameA" void -Values @($handle.AddrOfPinnedObject(), [ref]250)
Write-Host ("Computer Name (A): {0}" -f ([Encoding]::ASCII.GetString($computerNameA).Trim([char]0)))
$handle.Free()

# --- For GetComputerNameW (Unicode Version) ---
$computerNameW = New-Object byte[] (250*2)
$handle = [gchandle]::Alloc($computerNameW, 'Pinned')
Invoke-UnmanagedMethod "kernel32.dll" "GetComputerNameW" void -Values @($handle.AddrOfPinnedObject(), [ref]250)
Write-Host ("Computer Name (W): {0}" -f ([Encoding]::Unicode.GetString($computerNameW).Trim([char]0)))
$handle.Free()

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Clear-Host
Write-Host

# ZwQuerySystemInformation
# https://www.geoffchappell.com/studies/windows/km/ntoskrnl/api/ex/sysinfo/query.htm?tx=61&ts=0,1677

# SYSTEM_PROCESS_INFORMATION structure
# https://www.geoffchappell.com/studies/windows/km/ntoskrnl/api/ex/sysinfo/process.htm

# ZwQuerySystemInformation
# https://www.geoffchappell.com/studies/windows/km/ntoskrnl/api/ex/sysinfo/query.htm?tx=61&ts=0,1677

# SYSTEM_BASIC_INFORMATION structure
# https://www.geoffchappell.com/studies/windows/km/ntoskrnl/inc/api/ntexapi/system_basic_information.htm

# Step 1: Get required buffer size
$ReturnLength = 0
$dllResult = Invoke-UnmanagedMethod `
  -Dll "ntdll.dll" `
  -Function "ZwQuerySystemInformation" `
  -Return "uint32" `
  -Params "int SystemInformationClass, IntPtr SystemInformation, uint SystemInformationLength, ref uint ReturnLength" `
  -Values @(0, [IntPtr]::Zero, 0, [ref]$ReturnLength)

# Allocate buffer (add some extra room just in case)
$infoBuffer = New-IntPtr -Size $ReturnLength

# Step 2: Actual call with allocated buffer
$result = Invoke-UnmanagedMethod `
  -Dll "ntdll.dll" `
  -Function "ZwQuerySystemInformation" `
  -Return "uint32" `
  -Params "int SystemInformationClass, IntPtr SystemInformation, uint SystemInformationLength, ref uint ReturnLength" `
  -Values @(0, $infoBuffer, $ReturnLength, [ref]$ReturnLength)

if ($result -ne 0) {
    Write-Host "NtQuerySystemInformation failed: 0x$("{0:X}" -f $result)"
    Parse-ErrorMessage -MessageId $result
    New-IntPtr -hHandle $infoBuffer -Release
    return
}

# Parse values from the structure
$sysBasicInfo = [PSCustomObject]@{
    PageSize                     = [Marshal]::ReadInt32($infoBuffer,  0x08)
    NumberOfPhysicalPages        = [Marshal]::ReadInt32($infoBuffer,  0x0C)
    LowestPhysicalPageNumber     = [Marshal]::ReadInt32($infoBuffer,  0x10)
    HighestPhysicalPageNumber    = [Marshal]::ReadInt32($infoBuffer,  0x14)
    AllocationGranularity        = [Marshal]::ReadInt32($infoBuffer,  0x18)
    MinimumUserModeAddress       = [Marshal]::ReadIntPtr($infoBuffer, 0x20)
    MaximumUserModeAddress       = [Marshal]::ReadIntPtr($infoBuffer, 0x28)
    ActiveProcessorsAffinityMask = [Marshal]::ReadIntPtr($infoBuffer, 0x30)
    NumberOfProcessors           = [Marshal]::ReadByte($infoBuffer,   0x38)
}

# Step 1: Get required buffer size
$ReturnLength = 0
$dllResult = Invoke-UnmanagedMethod `
  -Dll "ntdll.dll" `
  -Function "ZwQuerySystemInformation" `
  -Return "uint32" `
  -Params "int SystemInformationClass, IntPtr SystemInformation, uint SystemInformationLength, ref uint ReturnLength" `
  -Values @(5, [IntPtr]::Zero, 0, [ref]$ReturnLength)

# Allocate buffer (add some extra room just in case)
$ReturnLength += 200
$procBuffer = New-IntPtr -Size $ReturnLength

# Step 2: Actual call with allocated buffer
$result = Invoke-UnmanagedMethod `
  -Dll "ntdll.dll" `
  -Function "ZwQuerySystemInformation" `
  -Return "uint32" `
  -Params "int SystemInformationClass, IntPtr SystemInformation, uint SystemInformationLength, ref uint ReturnLength" `
  -Values @(5, $procBuffer, $ReturnLength, [ref]$ReturnLength)

if ($result -ne 0) {
    Write-Host "NtQuerySystemInformation failed: 0x$("{0:X}" -f $result)"
    Parse-ErrorMessage -MessageId $result
    New-IntPtr -hHandle $procBuffer -Release
    return
}

$offset = 0
$processList = @()

while ($true) {
    try {
        $entryPtr = [IntPtr]::Add($procBuffer, $offset)
        $nextOffset = [Marshal]::ReadInt32($entryPtr, 0x00)

        $namePtr = [Marshal]::ReadIntPtr($entryPtr, 0x38 + [IntPtr]::Size)
        $processName = if ($namePtr -ne [IntPtr]::Zero) {
            [Marshal]::PtrToStringUni($namePtr)
        } else {
            "[System]"
        }

        $procObj = [PSCustomObject]@{
            ProcessId       = [Marshal]::ReadIntPtr($entryPtr, 0x50)
            ProcessName     = $processName
            NumberOfThreads = [Marshal]::ReadInt32($entryPtr, 0x04)
        }

        $processList += $procObj

        if ($nextOffset -eq 0) { break }
        $offset += $nextOffset
    } catch {
        Write-Host "Parsing error at offset $offset. Stopping."
        break
    }
}

New-IntPtr -hHandle $infoBuffer -Release
New-IntPtr -hHandle $procBuffer -Release

$sysBasicInfo | Format-List
$processList | Sort-Object ProcessName | Format-Table ProcessId, ProcessName, NumberOfThreads -AutoSize

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

GitHub - jhalon/SharpCall: Simple PoC demonstrating syscall execution in C#
https://github.com/jhalon/SharpCall

Red Team Tactics: Utilizing Syscalls in C# - Writing The Code - Jack Hacks
https://jhalon.github.io/utilizing-syscalls-in-csharp-2/

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Clear-Host

Write-Host
$hProc = [IntPtr]::Zero
$hProcNext = [IntPtr]::Zero
$ret = Invoke-UnmanagedMethod `
    -Dll NTDLL `
    -Function ZwGetNextProcess `
    -Values @($hProc, [UInt32]0x02000000, [UInt32]0x00, [UInt32]0x00, ([ref]$hProcNext)) `
    -Mode Allocate -SysCall
write-host "NtGetNextProcess Test: $ret"
write-host "hProcNext Value :$hProcNext"

Write-Host
$hThread = [IntPtr]::Zero
$hThreadNext = [IntPtr]::Zero
$ret = Invoke-UnmanagedMethod `
    -Dll NTDLL `
    -Function ZwGetNextThread `
    -Values @([IntPtr]::new(-1), $hThread, 0x0040, 0x00, 0x00, ([ref]$hThreadNext)) `
    -Mode AllocateEx -SysCall
write-host "NtGetNextThread Test: $ret"
write-host "hThreadNext Value :$hThreadNext"

Write-Host
$ret = Invoke-UnmanagedMethod `
    -Dll NTDLL `
    -Function NtClose `
    -Values @([IntPtr]$hProcNext) `
    -Mode Allocate -SysCall
write-host "NtClose Test: $ret"

Write-Host
$FileHandle = [IntPtr]::Zero
$IoStatusBlock    = New-IntPtr -Size 16
$ObjectAttributes = New-IntPtr -Size 48 -WriteSizeAtZero
$filePath = ("\??\{0}\test.txt" -f [Environment]::GetFolderPath('Desktop'))
$ObjectName = Init-NativeString -Encoding Unicode -Value $filePath
[Marshal]::WriteIntPtr($ObjectAttributes, 0x10, $ObjectName)
[Marshal]::WriteInt32($ObjectAttributes,  0x18, 0x40)
$ret = Invoke-UnmanagedMethod `
    -Dll NTDLL `
    -Function NtCreateFile `
    -Values @(
        ([ref]$FileHandle),   # OUT HANDLE
        0x40100080,           # DesiredAccess (GENERIC_WRITE | SYNCHRONIZE | FILE_WRITE_DATA)
        $ObjectAttributes,    # POBJECT_ATTRIBUTES
        $IoStatusBlock,       # PIO_STATUS_BLOCK
        [IntPtr]::Zero,       # AllocationSize
        0x80,                 # FileAttributes (FILE_ATTRIBUTE_NORMAL)
        0x07,                 # ShareAccess (read|write|delete)
        0x5,                  # CreateDisposition (FILE_OVERWRITE_IF)
        0x20,                 # CreateOptions (FILE_NON_DIRECTORY_FILE)
        [IntPtr]::Zero,       # EaBuffer
        0x00                  # EaLength
    ) `
    -Mode Protect -SysCall
Free-NativeString -StringPtr $ObjectName
write-host "NtCreateFile Test: $ret"
#>
function Build-ApiDelegate {
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline)]
        [PSCustomObject]$InterfaceSpec,

        [Parameter(Mandatory=$true)]
        [string]$UNIQUE_ID
    )

    $namespace = "namespace DynamicDelegates"
    $using = "`nusing System;`nusing System.Runtime.InteropServices;`n"
    $Params = Process-Parameters -InterfaceSpec $InterfaceSpec -Ignore
    $fixedReturnType = Process-ReturnType -ReturnType $InterfaceSpec.Return
    $charSet = if ($InterfaceSpec.CharSet) { "CharSet = CharSet.$($InterfaceSpec.CharSet)" } else { "CharSet = CharSet.Unicode" }
    $Return = @"
    [UnmanagedFunctionPointer(CallingConvention.$($InterfaceSpec.CallingType), $charSet)]
    public delegate $($fixedReturnType) $($UNIQUE_ID)(
        $($Params)
    );
"@

    return "$using`n$namespace`n{`n$Return`n}`n"
}
function Build-ApiInterfaceSpec {
    param (
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$Dll,

        [Parameter(Mandatory = $true, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [string]$Function,

        [Parameter(Mandatory = $true, Position = 3)]
        [ValidateSet("StdCall", "Cdecl")]
        [string]$CallingConvention = "StdCall",

        [Parameter(Mandatory = $true, Position = 4)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet(
            # Void
            "void",

            # Fully qualified .NET types
            "system.boolean", "system.byte", "system.char", "system.decimal", "system.double",
            "system.int16", "system.int32", "system.int64", "system.intptr", "system.object",
            "system.sbyte", "system.single", "system.string", "system.uint16", "system.uint32",
            "system.uint64", "system.uintptr",

            # Alternate type spellings and aliases
            "boolean", "dword32", "dword64", "int16", "int32", "int64", "single", "uint16",
            "uint32", "uint64",

            # Additional C/C++ & WinAPI aliases
            "double", "float", "long", "longlong", "tchar", "uchar", "ulong", "ulonglong",
            "short", "ushort",

            # Additional typedefs
            "atom", "dword_ptr", "dwordlong", "farproc", "hhook", "hresult", "NTSTATUS",
            "int_ptr", "intptr_t", "long_ptr", "lpbyte", "lpdword", "lparam", "pcstr",
            "pcwstr", "pstr", "pwstr", "uint_ptr", "uintptr_t", "wparam",

            # C# built-in types
            "bool", "byte", "char", "decimal", "int", "intptr", "nint", "nuint", "object",
            "sbyte", "string", "uint", "uintptr",

            # Common WinAPI handle types
            "hbitmap", "hbrush", "hcurs", "hdc", "hfont", "hicon", "hmenu", "hpen", "hrgn",

            # Pointer-based aliases
            "pbyte", "pchar", "pdword", "pint", "plong", "puint", "pulong", "pvoid", "lpvoid",

            # Special types
            "guid",

            # Windows/WinAPI types (common aliases)
            "dword", "handle", "hinstance", "hmodule", "hwnd", "lpcstr", "lpcwstr", "lpstr",
            "lpwstr", "ptr", "size_t", "ssize_t", "void*", "word", "phandle", "lresult",

            # STRSAFE typedefs
            "strsafe_lpstr", "strsafe_lpcstr", "strsafe_lpwstr", "strsafe_lpcwstr",
            "strsafe_lpcuwstr", "strsafe_pcnzch", "strsafe_pcnzwch", "strsafe_pcunzwch",

            # Wide-character (Unicode) types
            "pstring", "pwchar", "lpwchar", "pczpwstr", "pzpwstr", "pzwstr", "pzzwstr",
            "pczzwstr", "puczzwstr", "pcuczzwstr", "pnzwch", "pcnzwch", "punzwch", "pcunzwch",

            # ANSI string types
            "npstr", "pzpcstr", "pczpcstr", "pzzstr", "pczzstr", "pnzch", "pcnzch",

            # UCS types
            "ucschar", "pucschar", "pcucschar", "puucschar", "pcuucschar", "pucsstr",
            "pcucsstr", "puucsstr", "pcuucsstr",

            # Neutral ANSI/Unicode (TCHAR-based) Types
            "ptchar", "tbyte", "ptbyte", "ptstr", "lptstr", "pctstr", "lpctstr", "putstr",
            "lputstr", "pcutstr", "lpcutstr", "pzptstr", "pzzstr", "pczztstr", "pzzwstr", "pczzwstr"
        )]
        [string]$Return,

        [Parameter(Mandatory = $false, Position = 5)]
        [string]$Params,

        [Parameter(Mandatory = $false, Position = 6)]
        [ValidateSet("Unicode", "Ansi")]
        [string]$CharSet = "Unicode"
    )

    return [PSCustomObject]@{
        Dll     = $Dll
        Function= $Function
        Return  = $Return
        Params  = $Params
        CallingType = $CallingConvention
        CharSet     = $CharSet
    }
}
function Initialize-ApiObject {
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline)]
        [PSCustomObject]$ApiSpec,

        [Parameter(Mandatory = $false)]
        [string]$Mode = '',

        [Parameter(Mandatory=$false, ValueFromPipeline)]
        [switch]$SysCall
    )
    
    $hModule = [IntPtr]::Zero
    $BaseAddress = Ldr-LoadDll -dwFlags SEARCH_SYS32 -dll $ApiSpec.Dll
    if ($BaseAddress -ne $null -and $BaseAddress -ne [IntPtr]::Zero) {
        $hModule = [IntPtr]$BaseAddress
    }

    if ($hModule -eq [IntPtr]::Zero) {
        throw "Failed to load DLL: $($ApiSpec.Dll)"
    }

    $funcAddress = [IntPtr]::Zero
    $AnsiPtr = Init-NativeString -Value $ApiSpec.Function -Encoding Ansi
    $hresult = $Global:ntdll::LdrGetProcedureAddressForCaller(
        $hModule, $AnsiPtr, 0, [ref]$funcAddress, 0, 0)
    Free-NativeString -StringPtr $AnsiPtr
    if ($funcAddress -eq [IntPtr]::Zero -or $hresult -ne 0) {
        throw "Failed to find function: $($ApiSpec.Function)"
    }

    # Build delegate
    $baseAddress = [IntPtr]::Zero;
    $uniqueName = "$($ApiSpec.Function)Api$(Get-Random)"
    $delegateCode = Build-ApiDelegate -InterfaceSpec $ApiSpec -UNIQUE_ID $uniqueName

    Add-Type -TypeDefinition $delegateCode -Language CSharp -ErrorAction Stop
    $delegateType = [AppDomain]::CurrentDomain.GetAssemblies() |
        ForEach-Object { $_.GetType("DynamicDelegates.$uniqueName", $false, $true) } |
        Where-Object { $_ } |
        Select-Object -First 1

    if (-not $delegateType) {
        throw "Failed to get delegate type for $uniqueName"
    }

    if ($SysCall) {
        $SysCallPtr = [IntPtr]::Zero
        if ([IntPtr]::Size -gt 4) {
            if (![TEB]::IsRobustValidx64Stub($funcAddress)) {
                $dllName = if (($ApiSpec.Dll).EndsWith('.dll')) { $dllName } else { "$($ApiSpec.Dll).dll" }
                $SysCallPtr = New-IntPtr -Data (
                    Get-SysCallData -DllName $dllName -FunctionName $ApiSpec.Function -BytesToRead 25
                )
                if (![TEB]::IsRobustValidx64Stub($SysCallPtr)) {
                    Free-IntPtr $SysCallPtr
                    throw 'x64 stub not valid'
                }
            }
        }
        $sysCallID = if ($SysCallPtr -ne [IntPtr]::zero) {
            [Marshal]::ReadInt32(
                $SysCallPtr, 0x04)
        } elseif ([IntPtr]::Size -gt 4) {
            [Marshal]::ReadInt32(
                $funcAddress, 0x04)
        } else {
            0 # Place Holder
        }
        Free-IntPtr $SysCallPtr
        
        [byte[]]$shellcode = if ([IntPtr]::Size -gt 4) { 
            [byte[]]([TEB]::GenerateSyscallx64(
                ([BitConverter]::GetBytes($sysCallID))))
        } else {
            [byte[]]([TEB]::GenerateSyscallx86($funcAddress))
        }

        $lpflOldProtect = [Uint32]0;
        $baseAddress = [IntPtr]::Zero;
        $baseAddressPtr = [IntPtr]::Zero;
        $regionSize = [UIntPtr]::new($shellcode.Length);

        if ($Mode -eq 'Protect') {
            $baseAddressPtr = [gchandle]::Alloc($shellcode, 'pinned')
            $baseAddress = $baseAddressPtr.AddrOfPinnedObject()
            [IntPtr]$tempBase = $baseAddress

            if ([TEB]::NtProtectVirtualMemory(
                    [IntPtr]::new(-1),
                    [ref]$tempBase,
                    ([ref]$regionSize),
                    0x00000040,
                    [ref]$lpflOldProtect) -ne 0) {
                throw "Fail to Protect Memory for SysCall"
            }
        }
        elseif ($Mode -match "Allocate|AllocateEx") {
            $ret = if ($Mode -eq 'Allocate') {
                [TEB]::ZwAllocateVirtualMemory(
                    [IntPtr]::new(-1),
                    [ref] $baseAddress,
                    [UIntPtr]::new(0x00),
                    [ref] $regionSize,
                    0x3000, 
                    0x40 
                )
            } elseif ($Mode -eq 'AllocateEx') {
                [TEB]::ZwAllocateVirtualMemoryEx(
                   [IntPtr]::new(-1),
                   [ref]$baseAddress,
                   [ref]$regionSize,
                   0x3000, 0x40,
                   [IntPtr]0,0)
            }

            if ($ret -ne 0) {
                throw "Fail to allocate Memory for SysCall"
            }

            [Marshal]::Copy(
                $shellcode, 0, $baseAddress, $shellcode.Length)
        }

        $delegate = [Marshal]::GetDelegateForFunctionPointer(
            $baseAddress, $delegateType)
    }
    else {
        $delegate = [Marshal]::GetDelegateForFunctionPointer(
            $funcAddress, $delegateType)
    }

    return [PSCustomObject]@{
        Dll              = $ApiSpec.Dll
        Function         = $ApiSpec.Function
        FunctionPtr      = $funcAddress
        DelegateInstance = $delegate
        DelegateType     = $delegateType
        DelegateCode     = $delegateCode
        baseAddress      = $baseAddress
        baseAddressPtr   = $baseAddressPtr
        RegionSize       = $regionSize
    }
}
function Release-ApiObject {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline)]
        [PSCustomObject]$ApiObject
    )

    process {
        try {
            if ($ApiObject.baseAddressPtr -and $ApiObject.baseAddressPtr -ne [IntPtr]::Zero) {
                $null = $ApiObject.baseAddressPtr.Free()
            }
            elseif ($ApiObject.BaseAddress -and $ApiObject.BaseAddress -ne [IntPtr]::Zero) {
                [IntPtr]$baseAddrLocal = $ApiObject.BaseAddress
                [UIntPtr]$regionSizeLocal = $ApiObject.RegionSize

                $null = [TEB]::ZwFreeVirtualMemory(
                    [IntPtr]::new(-1),
                    [ref]$baseAddrLocal, 
                    [ref]$regionSizeLocal,
                    0x8000
                );
            }
            $ApiObject.Dll = $null
            $ApiObject.Function = $null
            $ApiObject.BaseAddress = $null
            $ApiObject.baseAddressPtr = $null
            $ApiObject.DelegateType = $null
            $ApiObject.DelegateCode = $null
            $ApiObject.DelegateInstance = $null
            $ApiObject.FunctionPtr = 0x0
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()

        } catch {
            Write-Warning "Failed to release ApiObject: $_"
        }
    }
}
function Invoke-UnmanagedMethod {
    param (
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$Dll,

        [Parameter(Mandatory = $true, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [string]$Function,

        [Parameter(Mandatory = $false, Position = 3)]
        [ValidateSet(
            # Void
            "void",

            # Fully qualified .NET types
            "system.boolean", "system.byte", "system.char", "system.decimal", "system.double",
            "system.int16", "system.int32", "system.int64", "system.intptr", "system.object",
            "system.sbyte", "system.single", "system.string", "system.uint16", "system.uint32",
            "system.uint64", "system.uintptr",

            # Alternate type spellings and aliases
            "boolean", "dword32", "dword64", "int16", "int32", "int64", "single", "uint16",
            "uint32", "uint64",

            # Additional C/C++ & WinAPI aliases
            "double", "float", "long", "longlong", "tchar", "uchar", "ulong", "ulonglong",
            "short", "ushort",

            # Additional typedefs
            "atom", "dword_ptr", "dwordlong", "farproc", "hhook", "hresult", "NTSTATUS",
            "int_ptr", "intptr_t", "long_ptr", "lpbyte", "lpdword", "lparam", "pcstr",
            "pcwstr", "pstr", "pwstr", "uint_ptr", "uintptr_t", "wparam",

            # C# built-in types
            "bool", "byte", "char", "decimal", "int", "intptr", "nint", "nuint", "object",
            "sbyte", "string", "uint", "uintptr",

            # Common WinAPI handle types
            "hbitmap", "hbrush", "hcurs", "hdc", "hfont", "hicon", "hmenu", "hpen", "hrgn",

            # Pointer-based aliases
            "pbyte", "pchar", "pdword", "pint", "plong", "puint", "pulong", "pvoid", "lpvoid",

            # Special types
            "guid",

            # Windows/WinAPI types (common aliases)
            "dword", "handle", "hinstance", "hmodule", "hwnd", "lpcstr", "lpcwstr", "lpstr",
            "lpwstr", "ptr", "size_t", "ssize_t", "void*", "word", "phandle", "lresult",

            # STRSAFE typedefs
            "strsafe_lpstr", "strsafe_lpcstr", "strsafe_lpwstr", "strsafe_lpcwstr",
            "strsafe_lpcuwstr", "strsafe_pcnzch", "strsafe_pcnzwch", "strsafe_pcunzwch",

            # Wide-character (Unicode) types
            "pstring", "pwchar", "lpwchar", "pczpwstr", "pzpwstr", "pzwstr", "pzzwstr",
            "pczzwstr", "puczzwstr", "pcuczzwstr", "pnzwch", "pcnzwch", "punzwch", "pcunzwch",

            # ANSI string types
            "npstr", "pzpcstr", "pczpcstr", "pzzstr", "pczzstr", "pnzch", "pcnzch",

            # UCS types
            "ucschar", "pucschar", "pcucschar", "puucschar", "pcuucschar", "pucsstr",
            "pcucsstr", "puucsstr", "pcuucsstr",

            # Neutral ANSI/Unicode (TCHAR-based) Types
            "ptchar", "tbyte", "ptbyte", "ptstr", "lptstr", "pctstr", "lpctstr", "putstr",
            "lputstr", "pcutstr", "lpcutstr", "pzptstr", "pzzstr", "pczztstr", "pzzwstr", "pczzwstr"
        )]
        [string]$Return,

        [Parameter(Mandatory = $false, Position = 4)]
        [string]$Params,

        [Parameter(Mandatory = $false, Position = 5)]
        [ValidateSet("StdCall", "Cdecl")]
        [string]$CallingConvention,

        [Parameter(Mandatory = $false, Position = 6)]
        [object[]]$Values,

        [Parameter(Mandatory = $false, Position = 7)]
        [ValidateSet("Unicode", "Ansi")]
        [string]$CharSet = "Unicode",

        [Parameter(Mandatory = $false, Position = 8)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Allocate', 'AllocateEx', 'Protect')]
        [string]$Mode = 'Allocate',

        [Parameter(Mandatory = $false, Position = 9)]
        [switch]$SysCall
    )

    # Detect platform
    if (-not $CallingConvention) {
        if ([IntPtr]::Size -eq 8) {
            $CallingConvention = "StdCall" 
        }
        else {
            $CallingConvention = "StdCall"
        }
    }

    # Lazy Mode Detection
    $Count = 0
    [void][int]::TryParse($Values.Count,[ref]$count)
    $lazyMode = (-not $Params) -and ($Count -gt 0)
    $IsArrayObj = $Count -eq 1 -and $Values[0] -is [System.Array]

    if (-not $Return) {
        $Return = "Int32"
    }

    if ($IsArrayObj) {
        Write-error "Cast all Params with '-Values @()' Please"
        return
    }

    if ($lazyMode) {
        
        try {
            $idx = 0
            $Params = (
                $Values | % {
                    ++$idx
                    if ($_.Value -or ($_ -is [ref])) {
                        $byRef = 'ref '
                        $Name  = $_.Value.GetType().Name
                    }
                    else {
                        $byRef = ''
                        $Name  = $_.GetType().Name
                    }
                    "{0}{1} {2}" -f $byRef, $Name, (Get-Base26Name -idx $idx)
                }
            ) -join ", "
        }
        catch {
            throw "auto parse params fail"
        }

        $CharSet = if ($Function -like "*A") { "Ansi" } else { "Unicode" }
    }

    $apiSpec = Build-ApiInterfaceSpec -Dll $Dll  `
        -Function $Function  `
        -Return $Return  `
        -CallingConvention $CallingConvention `
        -Params $Params `
        -CharSet $CharSet

    $apiObj = if ($SysCall) {
        Initialize-ApiObject -ApiSpec $apiSpec -Mode $Mode -SysCall
    }
    else {
        $apiSpec | Initialize-ApiObject
    }

    try {
        return $apiObj | Invoke-Object -Params $Values -type API
    }
    finally {
        $apiObj | Release-ApiObject
    }
}

<#
.HELPERS 

UnicodeString function helper, 
just for testing purpose

+++ Struct Info +++

typedef struct _UNICODE_STRING {
  USHORT Length; [ushort = 2]
  USHORT MaximumLength; [ushort = 2]
  ** in x64 enviroment Add 4 byte's padding **
  PWSTR  Buffer; [IntPtr].Size
} UNICODE_STRING, *PUNICODE_STRING;

Buffer Offset == [IntPtr].Size { x86=4, x64=8 }

+++ Test Code +++

Clear-Host
Write-Host

$unicodeStringPtr = Init-NativeString -Value 99 -Encoding Unicode
Parse-NativeString -StringPtr $unicodeStringPtr -Encoding Unicode
Free-NativeString -StringPtr $unicodeStringPtr

$ansiStringPtr = Init-NativeString -Value 99 -Encoding Ansi
Parse-NativeString -StringPtr $ansiStringPtr -Encoding Ansi
Free-NativeString -StringPtr $ansiStringPtr

$unicodeStringPtr = [IntPtr]::Zero
$unicodeStringPtr = Manage-UnicodeString -Value 'data123'
Parse-UnicodeString -unicodeStringPtr $unicodeStringPtr
Manage-UnicodeString -UnicodeStringPtr $unicodeStringPtr -Release
#>
function Init-NativeString {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Ansi', 'Unicode')]
        [string]$Encoding,

        [Int32]$Length = 0,
        [Int32]$MaxLength = 0,
        [Int32]$BufferSize = 0
    )

     # Determine required byte length of the string in the specified encoding
    if ($Encoding -eq 'Ansi') {
        $requiredSize = [System.Text.Encoding]::ASCII.GetByteCount($Value)
    } else {
        $requiredSize = [System.Text.Encoding]::Unicode.GetByteCount($Value)
    }

    if ($BufferSize -gt 0 -and $BufferSize -lt $requiredSize) {
        throw "BufferSize ($BufferSize) is smaller than the encoded string size ($requiredSize)."
    }

    $stringPtr = New-IntPtr -Size 16

    if ($Encoding -eq 'Ansi') {
        if ($Length -le 0) {
            $Length = [System.Text.Encoding]::ASCII.GetByteCount($Value)
            if ($Length -ge 0xFFFE) {
                $Length = 0xFFFC
            }
        }

        if ($BufferSize -gt 0) {
            $bufferPtr = New-IntPtr -Size $BufferSize
            $bytes = [System.Text.Encoding]::ASCII.GetBytes($Value)
            [Marshal]::Copy($bytes, 0, $bufferPtr, $bytes.Length)
        }
        else {
            $bufferPtr = [Marshal]::StringToHGlobalAnsi($Value)
        }
        if ($MaxLength -le 0) {
            $maxLength = $Length + 1
        }
    }
    else {
        if ($Length -le 0) {
            $Length = $Value.Length * 2
            if ($Length -ge 0xFFFE) {
                $Length = 0xFFFC
            }
        }
        if ($BufferSize -gt 0) {
            $bufferPtr = New-IntPtr -Size $BufferSize
            $bytes = [System.Text.Encoding]::Unicode.GetBytes($Value)
            [Marshal]::Copy($bytes, 0, $bufferPtr, $bytes.Length)
        }
        else {
            $bufferPtr = [Marshal]::StringToHGlobalUni($Value)
        }
        if ($MaxLength -le 0) {
            $maxLength = $Length + 2
        }
    }

    [Marshal]::WriteInt16($stringPtr, 0, $Length)
    [Marshal]::WriteInt16($stringPtr, 2, $maxLength)
    [Marshal]::WriteIntPtr($stringPtr, [IntPtr]::Size, $bufferPtr)

    return $stringPtr
}
function Parse-NativeString {
    param (
        [Parameter(Mandatory = $true)]
        [IntPtr]$StringPtr,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Ansi', 'Unicode')]
        [string]$Encoding = 'Ansi'
    )

    if ($StringPtr -eq [IntPtr]::Zero) {
        return
    }

    $Length = [Marshal]::ReadInt16($StringPtr, 0)
    $Size   = [Marshal]::ReadInt16($StringPtr, 2)
    $BufferPtr = [Marshal]::ReadIntPtr($StringPtr, [IntPtr]::Size)
    if ($Length -le 0) {
        return $null
    }

    if ($Encoding -eq 'Ansi') {
        # Length is number of bytes
        $Data = [Marshal]::PtrToStringAnsi($BufferPtr, $Length)
    } else {
        # Unicode, length is bytes, divide by 2 for chars
        $Data = [Marshal]::PtrToStringUni($BufferPtr, $Length / 2)
    }

    return [PSCustomObject]@{
        Length        = $Length
        MaximumLength = $Size
        StringData    = $Data
    }
}
function Free-NativeString {
    param (
        [Parameter(Mandatory = $true)]
        [IntPtr]$StringPtr
    )
    
    if ($StringPtr -eq [IntPtr]::Zero) {
        Write-Warning 'Failed to free pointer: The pointer is null'
        return
    }

    $ptr = [IntPtr]::Zero
    try {
        $bufferPtr = [Marshal]::ReadIntPtr($StringPtr, [IntPtr]::Size)
        if ($bufferPtr -ne [IntPtr]::Zero) {
            [Marshal]::FreeHGlobal($bufferPtr)
        }
        else {
            Write-Warning 'Failed to free buffer: The buffer pointer is null.'
        }
        [Marshal]::FreeHGlobal($StringPtr)
    }
    catch {
        Write-Warning 'An error occurred while attempting to free memory'
        return
    }
}

<#
.SYNOPSIS
Manages native UNICODE_STRING memory and content for P/Invoke.

.DESCRIPTION
This function allows for the creation of new UNICODE_STRING structures,
in-place updating of existing ones, and safe release of all associated
unmanaged memory (both the structure and its internal string buffer)
using low-level NTDLL APIs.

.USE
Manage-UnicodeString -Value '?'
Manage-UnicodeString -Value '?' -UnicodeStringPtr ?
Manage-UnicodeString -UnicodeStringPtr ? -Release
#>
function Manage-UnicodeString {
    param (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string] $Value,

        [Parameter(Mandatory = $false)]
        [IntPtr] $UnicodeStringPtr = [IntPtr]::Zero,

        [switch] $Release
    )

    # Check if the pointer is valid (non-zero)
    $isValidPtr = $UnicodeStringPtr -ne [IntPtr]::Zero

    # Case 1: Value only - allocate and create a new string (if pointer is zero)
    if ($Value -and -not $isValidPtr -and -not $Release) {
        $unicodeStringPtr = New-IntPtr -Size 16
        $returnValue = $Global:ntdll::RtlCreateUnicodeString($unicodeStringPtr, $Value)

        # Check if the lowest byte is 1 (indicating success as per the C code's CONCAT71)
        if (($returnValue -band 0xFF) -ne 1) {
                throw "Failed to create Unicode string for '$Value'. NTSTATUS return value: 0x$hexReturnValue"
        }

        return $unicodeStringPtr
    }

    # Case 2: Value + existing pointer - reuse the pointer (if pointer is valid)
    elseif ($Value -and $isValidPtr -and -not $Release) {
        $null = $Global:ntdll::RtlFreeUnicodeString($unicodeStringPtr)
        $Global:ntdll::RtlZeroMemory($unicodeStringPtr, [UIntPtr]::new(16))
        $returnValue = $Global:ntdll::RtlCreateUnicodeString($unicodeStringPtr, $Value)
        
        # Check if the lowest byte is 1 (indicating success as per the C code's CONCAT71)
        if (($returnValue -band 0xFF) -ne 1) {
                throw "Failed to create Unicode string for '$Value'. NTSTATUS return value: 0x$hexReturnValue"
        }
        return
    }

    # Case 3: Pointer + Release - cleanup the string (if pointer is valid)
    elseif (-not $Value -and $isValidPtr -and $Release) {
        $null = $Global:ntdll::RtlFreeUnicodeString($unicodeStringPtr)
        New-IntPtr -hHandle $unicodeStringPtr -Release
        return
    }

    # Invalid combinations (no valid operation matched)
    else {
        throw "Invalid parameter combination. You must provide one of the following:
        1) -Value to create a new string,
        2) -Value and -unicodeStringPtr to reuse the pointer,
        3) -unicodeStringPtr and -Release to free the string."
    }
}

<#
.HELPERS 

* VARIANT structure (oaidl.h)
* https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-variant

struct {
    VARTYPE vt;           0x0
    WORD    wReserved1;   0x2
    WORD    wReserved2;
    WORD    wReserved3;
    union 
    {
        LONGLONG llVal;
        LONG lVal;
        BYTE bVal;
        SHORT iVal;
        FLOAT fltVal;
        DOUBLE dblVal;
        VARIANT_BOOL boolVal;
        VARIANT_BOOL __OBSOLETE__VARIANT_BOOL;
        SCODE scode;
        CY cyVal;
        DATE date;
        BSTR bstrVal;
        IUnknown *punkVal;
        IDispatch *pdispVal;
        SAFEARRAY *parray;
        BYTE *pbVal;
        SHORT *piVal;
        LONG *plVal;
        LONGLONG *pllVal;
        FLOAT *pfltVal;
        DOUBLE *pdblVal;
        VARIANT_BOOL *pboolVal;
        VARIANT_BOOL *__OBSOLETE__VARIANT_PBOOL;
        SCODE *pscode;
        CY *pcyVal;
        DATE *pdate;
        BSTR *pbstrVal;
        IUnknown **ppunkVal;
        IDispatch **ppdispVal;
        SAFEARRAY **pparray;
        VARIANT *pvarVal;
        PVOID byref;
        CHAR cVal;
        USHORT uiVal;
        ULONG ulVal;
        ULONGLONG ullVal;
        INT intVal;
        UINT uintVal;
        DECIMAL *pdecVal;
        CHAR *pcVal;
        USHORT *puiVal;
        ULONG *pulVal;
        ULONGLONG *pullVal;
        INT *pintVal;
        UINT *puintVal;
    }
} VARIANT *LPVARIANT;

enum VARENUM
{
    VT_EMPTY  = 0,
    VT_NULL	= 1,
    VT_I2	= 2,
    VT_I4	= 3,
    VT_R4	= 4,
    VT_R8	= 5,
    VT_CY	= 6,
    VT_DATE	= 7,
    VT_BSTR	= 8,
    VT_DISPATCH	= 9,
    VT_ERROR	= 10,
    VT_BOOL	= 11,
    VT_VARIANT	= 12,
    VT_UNKNOWN	= 13,
    VT_DECIMAL	= 14,
    VT_I1	= 16,
    VT_UI1	= 17,
    VT_UI2	= 18,
    VT_UI4	= 19,
    VT_I8	= 20,
    VT_UI8	= 21,
    VT_INT	= 22,
    VT_UINT	= 23,
    VT_VOID	= 24,
    VT_HRESULT	= 25,
    VT_PTR	= 26,
    VT_SAFEARRAY	= 27,
    VT_CARRAY	= 28,
    VT_USERDEFINED	= 29,
    VT_LPSTR	= 30,
    VT_LPWSTR	= 31,
    VT_RECORD	= 36,
    VT_INT_PTR	= 37,
    VT_UINT_PTR	= 38,
    VT_FILETIME	= 64,
    VT_BLOB	= 65,
    VT_STREAM	= 66,
    VT_STORAGE	= 67,
    VT_STREAMED_OBJECT	= 68,
    VT_STORED_OBJECT	= 69,
    VT_BLOB_OBJECT	= 70,
    VT_CF	= 71,
    VT_CLSID	= 72,
    VT_VERSIONED_STREAM	= 73,
    VT_BSTR_BLOB	= 0xfff,
    VT_VECTOR	= 0x1000,
    VT_ARRAY	= 0x2000,
    VT_BYREF	= 0x4000,
    VT_RESERVED	= 0x8000,
    VT_ILLEGAL	= 0xffff,
    VT_ILLEGALMASKED	= 0xfff,
    VT_TYPEMASK	= 0xfff
} ;

~~~~~~~~~~~~~~~

"ApiMajorVersion", "ApiMinorVersion", "ProductVersionString" | ForEach-Object {
    $name = $_
    $outVarPtr = New-Variant -Type VT_EMPTY
    $inVarPtr  = New-Variant -Type VT_BSTR -Value $name
    try {
        $ret = Use-ComInterface `
            -CLSID "C2E88C2F-6F5B-4AAA-894B-55C847AD3A2D" `
            -IID "85713fa1-7796-4fa2-be3b-e2d6124dd373" `
            -Index 1 -Name "GetInfo" `
            -Values @($inVarPtr, $outVarPtr) `
            -Type IDispatch

        if ($ret -eq 0) {
            $value = Parse-Variant -variantPtr $outVarPtr
            Write-Host "$name -> $value"
        }

    } finally {
        Free-Variant $inVarPtr
        Free-Variant $outVarPtr
    }
}
#>
function New-Variant {
    param(
        [Parameter(Mandatory)]
        [ValidateSet(
            "VT_EMPTY","VT_NULL","VT_I2",
            "VT_I4","VT_R4","VT_R8",
            "VT_BOOL","VT_BSTR","VT_DATE"
        )] 
        [string]$Type,

        [object]$Value
    )

    # Allocate VARIANT struct (24 bytes)
    $variantPtr = New-IntPtr -Size 24

    # Map type string to VARENUM
    $vt = switch ($Type) {
        "VT_EMPTY" {0}
        "VT_NULL"  {1}
        "VT_I2"    {2}
        "VT_I4"    {3}
        "VT_R4"    {4}
        "VT_R8"    {5}
        "VT_DATE"  {7}
        "VT_BSTR"  {8}
        "VT_BOOL"  {11}
        default    { throw "Unsupported VARIANT type $Type" }
    }

    [Marshal]::WriteInt16($variantPtr, 0, $vt)

    # Write value
    switch ($vt) {
        0  { } # VT_EMPTY, do nothing
        2  { [Marshal]::WriteInt16($variantPtr, 8, [int16]$Value) }  # VT_I2
        3  { [Marshal]::WriteInt32($variantPtr, 8, [int32]$Value) }  # VT_I4
        4  { [Marshal]::WriteInt32($variantPtr, 8, [BitConverter]::ToInt32([BitConverter]::GetBytes([float]$Value),0)) } # VT_R4
        5  { [Marshal]::WriteInt64($variantPtr, 8, [BitConverter]::ToInt64([BitConverter]::GetBytes([double]$Value),0)) } # VT_R8
        7  { # VT_DATE = OLE Automation DATE
            $dateVal = [double]([datetime]$Value).ToOADate()
            [Marshal]::WriteInt64($variantPtr, 8, [BitConverter]::ToInt64([BitConverter]::GetBytes($dateVal),0))
        }
        8  { # VT_BSTR
            $bstr = [Marshal]::StringToBSTR($Value)
            [Marshal]::WriteIntPtr($variantPtr, 8, $bstr)
        }
        11 { # VT_BOOL
            $boolVal = if ($Value) { -1 } else { 0 } # VARIANT_TRUE/-FALSE
            [Marshal]::WriteInt16($variantPtr, 8, $boolVal)
        }
    }

    return $variantPtr
}
function Parse-Variant {
    param([IntPtr]$variantPtr)

    if ($variantPtr -eq [IntPtr]::Zero) { return $null }

    $vt = [Marshal]::ReadInt16($variantPtr, 0)

    switch ($vt) {
        0  { return $null } # VT_EMPTY
        1  { return $null } # VT_NULL
        2  { return [Marshal]::ReadInt16($variantPtr, 8) } # VT_I2
        3  { return [Marshal]::ReadInt32($variantPtr, 8) } # VT_I4
        4  { return [BitConverter]::ToSingle([BitConverter]::GetBytes([Marshal]::ReadInt32($variantPtr, 8)),0) } # VT_R4
        5  { return [BitConverter]::ToDouble([BitConverter]::GetBytes([Marshal]::ReadInt64($variantPtr, 8)),0) } # VT_R8
        7  { return [datetime]::FromOADate([BitConverter]::ToDouble([BitConverter]::GetBytes([Marshal]::ReadInt64($variantPtr, 8)),0)) } # VT_DATE
        8  { # VT_BSTR
            $bstrPtr = [Marshal]::ReadIntPtr($variantPtr, 8)
            if ($bstrPtr -ne [IntPtr]::Zero) {
                return [Marshal]::PtrToStringBSTR($bstrPtr)
            }
            return $null
        }
        11 { return ([Marshal]::ReadInt16($variantPtr, 8) -ne 0) } # VT_BOOL
        default { return "[Unsupported VARIANT type $vt]" }
    }
}
function Free-Variant {
    param([IntPtr]$variantPtr)
    if ($variantPtr -eq [IntPtr]::Zero) { return }

    $vt = [Marshal]::ReadInt16($variantPtr, 0)
    if ($vt -eq 8) { # VT_BSTR
        $bstrPtr = [Marshal]::ReadIntPtr($variantPtr, 8)
        if ($bstrPtr -ne [IntPtr]::Zero) { [Marshal]::FreeBSTR($bstrPtr) }
    }
    [Marshal]::FreeHGlobal($variantPtr)
}

<#
.SYNOPSIS
* keyhelper API
* Source, change edition` script by windows addict

* KMS Local Activation Tool
* https://github.com/laomms/KmsTool/blob/main/Form1.cs

* 'Retail', 'OEM', 'Volume', 'Volume:GVLK', 'Volume:MAK'
  Any other case, it use default key

_wcsnicmp(input, "Retail", 6)           ? _CHANNEL_ENUM = 1
_wcsnicmp(input, "OEM", 3)              ? _CHANNEL_ENUM = 2
_wcsnicmp(input, "Volume:MAK", 10)      ? _CHANNEL_ENUM = 4
_wcsnicmp(input, "Volume:GVLK", 11)     ? _CHANNEL_ENUM = 3
_wcsnicmp(input, "Volume", 6)           ? _CHANNEL_ENUM = 3
)
#>
function Get-ProductKeys {
    param (
        [Parameter(Mandatory = $true)]
        [string]$EditionID,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Default', 'Retail', 'OEM', 'Volume:GVLK', 'Volume:MAK')]
        [string]$ProductKeyType
    )

    $id = 0
    $result = @()
    $defaultKey = ''

    if ([int]::TryParse($EditionID, [ref]$null)) {
        $id = [int]$EditionID
    }
    else {
        $id = $Global:productTypeTable | Where-Object ProductID -eq $EditionID | Select-Object -ExpandProperty DWORD
        if ($id -eq 0) {
            $null = $Global:PKHElper::GetEditionIdFromName($EditionID, [ref]$id)
        }
    }

    if ($id -eq 0) {
        throw "Could not resolve edition ID from input '$EditionID'."
    }

    # Step 1 - Retrieve the 'Default' key for the edition upfront
    $keyOutPtr = $typeOutPtr = $ProductKeyTypePtr = [IntPtr]::Zero
    try {
        $hResults = $Global:PKHElper::SkuGetProductKeyForEdition($id, [IntPtr]::zero, [ref]$keyOutPtr, [ref]$typeOutPtr)
        if ($hResults -eq 0) {
            $defaultKey = [Marshal]::PtrToStringUni($keyOutPtr)
        }
    }
    catch { }
    finally {
        ($keyOutPtr, $typeOutPtr) | % { Free-IntPtr -handle $_ -Method Heap}
        ($ProductKeyTypePtr, $ProductKeyTypePtr) | % { Free-IntPtr -handle $_ -Method Auto}
    }

    # Step 2 - Case of specic group Key
    if ($ProductKeyType) {
        # Handle specific ProductKeyType request
        try {
            $keyOutPtr = $typeOutPtr = $ProductKeyTypePtr = [IntPtr]::Zero
            if ($ProductKeyType -eq 'Default') {
                $keyOut = $defaultKey
            }
            else {
                $ProductKeyTypePtr = [Marshal]::StringToHGlobalUni($ProductKeyType)
                $hResults = $Global:PKHElper::SkuGetProductKeyForEdition($id, $ProductKeyTypePtr, [ref]$keyOutPtr, [ref]$typeOutPtr)
                if ($hResults -eq 0) {
                    $keyOut = [Marshal]::PtrToStringUni($keyOutPtr)
                }
            }

            $isDefault = !($keyOut -eq $defaultKey)
            $IsValue   = !([String]::IsNullOrWhiteSpace($keyOut))

            if ($IsValue -and (($ProductKeyType -eq 'Default') -or ($ProductKeyType -ne 'Default' -and $isDefault))) {
                $result += [PSCustomObject]@{
                    ProductKeyType = $ProductKeyType
                    ProductKey     = $keyOut
                }
            }
        }
        catch {}
        finally {
            ($keyOutPtr, $typeOutPtr) | % { Free-IntPtr -handle $_ -Method Heap }
            ($ProductKeyTypePtr, $ProductKeyTypePtr) | % { Free-IntPtr -handle $_ -Method Auto }
        }
    }

    # Step 3 - Case of Whole option's
    if (-not $ProductKeyType) {
        # Loop through other key types (excluding 'Default' as it's handled above)
        foreach ($group in @('Retail', 'OEM', 'Volume:GVLK', 'Volume:MAK' )) {
            try {
                $keyOutPtr = $typeOutPtr = $ProductKeyTypePtr = [IntPtr]::Zero
                $ProductKeyTypePtr = [Marshal]::StringToHGlobalUni($group)
                $hResults = $Global:PKHElper::SkuGetProductKeyForEdition($id, $ProductKeyTypePtr, [ref]$keyOutPtr, [ref]$typeOutPtr)
                if ($hResults -eq 0) {
                    $keyOut = [Marshal]::PtrToStringUni($keyOutPtr)
                    if (-not [string]::IsNullOrWhiteSpace($keyOut)) {
                        $result += [PSCustomObject]@{
                            ProductKeyType = $group
                            ProductKey     = $keyOut
                        }
                    }
                }
            }
            catch {}
            finally {
                Free-IntPtr -handle $ProductKeyTypePtr -Method Auto
                ($keyOutPtr, $typeOutPtr) | % { Free-IntPtr -handle $_ -Method Heap }
            }
        }
            
        # Now, filter the collected results based on your specific rules
        $seenKeys = @{}
        $filterResults = @()

        # Add the 'Default' key to results if it's valid
        if (-not [string]::IsNullOrWhiteSpace($defaultKey)) {
            $seenKeys[$defaultKey] = $true
            $filterResults += [PSCustomObject]@{
                ProductKeyType = "Default"
                ProductKey     = $defaultKey
            }
        }

        # Add other entries only if their ProductKey hasn't been seen yet
        foreach ($item in $result) {
            if (-not [string]::IsNullOrWhiteSpace($item.ProductKey) -and -not $seenKeys.ContainsKey($item.ProductKey)) {
                $filterResults += $item
                $seenKeys[$item.ProductKey] = $true
            }
        }
        $result = $filterResults
    }

    return $result
}

<#
.SYNOPSIS
Read PkeyConfig data from System,
Include Windows & Office pKeyConfig license's
#>
function Init-XMLInfo {
    $paths = @(
        "C:\Windows\System32\spp\tokens\pkeyconfig\pkeyconfig.xrm-ms",
        "C:\Windows\System32\spp\tokens\pkeyconfig\pkeyconfig-csvlk.xrm-ms",
        "C:\Windows\System32\spp\tokens\pkeyconfig\pkeyconfig-downlevel.xrm-ms",
        "C:\Program Files\Microsoft Office\root\Licenses16\pkeyconfig-office.xrm-ms"
    )

    $entries = @()
    foreach ($path in $paths) {
        if (Test-Path -Path $path) {
            $extracted = Extract-Base64Xml -FilePath $path
            if ($extracted) {
                $entries += $extracted
            }
        }
    }

    return $entries
}
function Extract-Base64Xml {
    param (
        [string]$FilePath
    )

    # Check if the file exists
    if (-Not (Test-Path $FilePath)) {
        Write-Warning "File not found: $FilePath"
        return $null
    }

    # Read the content of the file
    $content = Get-Content -Path $FilePath -Raw

    # Use regex to find all Base64 encoded strings between <tm:infoBin> tags
    $matches = [regex]::Matches($content, '<tm:infoBin name="pkeyConfigData">(.*?)<\/tm:infoBin>', [RegexOptions]::Singleline)

    $configurationsList = @()

    foreach ($match in $matches) {
        # Extract the Base64 encoded string
        $base64String = $match.Groups[1].Value.Trim()

        # Decode the Base64 string
        try {
            $decodedBytes = [Convert]::FromBase64String($base64String)
            $decodedString = [Encoding]::UTF8.GetString($decodedBytes)
            [xml]$xmlData = $decodedString

            # Process ProductKeyConfiguration
            #$xmlData.OuterXml | Out-File 'C:\Users\Administrator\Desktop\License.txt'
            if ($xmlData.ProductKeyConfiguration.Configurations) {
                foreach ($config in $xmlData.ProductKeyConfiguration.Configurations.ChildNodes) {
                    # Create a PSCustomObject for each configuration
                    $configObj = [PSCustomObject]@{
                        ActConfigId       = $config.ActConfigId
                        RefGroupId        = $config.RefGroupId
                        EditionId         = $config.EditionId
                        ProductDescription = $config.ProductDescription
                        ProductKeyType    = $config.ProductKeyType
                        IsRandomized      = $config.IsRandomized
                    }
                    $configurationsList += $configObj
                }
            }
        } catch {
            Write-Warning "Failed to decode Base64 string: $_"
        }
    }

    # Return the list of configurations
    return $configurationsList
}

<#
.SYNOPSIS
Get System Build numbers using low level methods

#>
Function Init-osVersion {
    
    <#
        First try read from KUSER_SHARED_DATA 
        And, if fail, Read from PEB.!

        RtlGetNtVersionNumbers Read from PEB. [X64 offset]
        * v3 = NtCurrentPeb();
        * OSMajorVersion -> 0x118 (v3->OSMajorVersion)
        * OSMinorVersion -> 0x11C (v3->OSMinorVersion)
        * OSBuildNumber  -> 0x120 (v3->OSBuildNumber | 0xF0000000)

        RtlGetVersion, do the same, just read extra info from PEB
        * v2 = NtCurrentPeb();
        * a1[1] = v2->OSMajorVersion;
        * a1[2] = v2->OSMinorVersion;
        * a1[3] = v2->OSBuildNumber;
        * a1[4] = v2->OSPlatformId;
        * Buffer = v2->CSDVersion.Buffer;
    #>

    if (-not $Global:PebPtr -or $Global:PebPtr -eq [IntPtr]::Zero) {
        $Global:PebPtr = NtCurrentTeb -Peb
        #$Global:PebPtr = $Global:ntdll::RtlGetCurrentPeb()
    }

    try {
        # 0x026C, ULONG NtMajorVersion; NT 4.0 and higher
        $NtMajorVersion = [Marshal]::ReadInt32([IntPtr](0x7FFE0000 + 0x26C))

        # 0x0270, ULONG NtMinorVersion; NT 4.0 and higher
        $NtMinorVersion = [Marshal]::ReadInt32([IntPtr](0x7FFE0000 + 0x270))

        # 0x0260, ULONG NtBuildNumber; NT 10.0 & higher
        $NtBuildNumber  = [Marshal]::ReadInt32([IntPtr](0x7FFE0000 + 0x0260))

        if (($NtMajorVersion -lt 10) -or (
            $NtBuildNumber -lt 10240)) {
      
          # this offset for nt 10.0 & Up
          # NT 6.3 end in 9600,
          # nt 10.0 start with 10240 (RTM)

          # Before, we stop throw, 
          # Try read from PEB memory.

          $offset = if ([IntPtr]::Size -eq 8) { 0x120 } else { 0x0AC }
          $NtBuildNumber = [int][Marshal]::ReadInt16($Global:PebPtr, $offset)

          # 0xAC, 0x0120, USHORT OSBuildNumber; 4.0 and higher
          if ($NtBuildNumber -lt 1381) {
            throw }
        }

        # Extract Service Pack Major (high byte) and Minor (low byte)
        # *((_WORD *)a1 + 138) = HIBYTE(v2->OSCSDVersion);
        # *((_WORD *)a1 + 139) = (unsigned __int8)v2->OSCSDVersion;
        $offset = if ([IntPtr]::Size -eq 8) { 0x122 } else { 0xAE }
        $oscVersion = [Marshal]::ReadInt16($Global:PebPtr, $offset)
        $wServicePackMajor = ($oscVersion -shr 8) -band 0xFF
        $wServicePackMinor = $oscVersion -band 0xFF

        # Retrieve the OS version details
        return [PSCustomObject]@{
            Major   = $NtMajorVersion
            Minor   = $NtMinorVersion
            Build   = $NtBuildNumber
            UBR     = $Global:ubr
            Version = ($NtMajorVersion,$NtMinorVersion,$NtBuildNumber)
            ServicePackMajor = $wServicePackMajor
            ServicePackMinor = $wServicePackMinor
        }
    }
    catch {}
        
    # Fallback: REGISTRY
    try {
        $major = (Get-ItemProperty -Path $Global:CurrentVersion -Name CurrentMajorVersionNumber -ea 0).CurrentMajorVersionNumber
        $minor = (Get-ItemProperty -Path $Global:CurrentVersion -Name CurrentMinorVersionNumber -ea 0).CurrentMinorVersionNumber
        $build = (Get-ItemProperty -Path $Global:CurrentVersion -Name CurrentBuildNumber -ea 0).CurrentBuildNumber
        $osVersion = [PSCustomObject]@{
            Major   = [int]$major
            Minor   = [int]$minor
            Build   = [int]$build
            UBR     = $Global:ubr
            Version = @([int]$major, [int]$minor, [int]$build)
            ServicePackMajor = 0
            ServicePackMinor = 0
        }
        return $osVersion
    }
    catch {
    }

    Clear-host
    Write-Host
    write-host "Failed to retrieve OS version from all methods."
    Write-Host
    read-host
    exit 1
}

<#
.SYNOPSIS
Get Edition Name using low level methods

#>
function Get-ProductID {
    
    <# 
        Experiment way,
        who work only on online active system,
        that why i don't use it !

        $LicensingProducts = (
            Get-SLIDList -eQueryIdType SL_ID_APPLICATION -eReturnIdType SL_ID_PRODUCT_SKU -pQueryId $windowsAppID | ? { Retrieve-SKUInfo -SkuId $_ -eReturnIdType SL_ID_PKEY }
            ) | % {
            [PSCustomObject]@{
                ID            = $_
                Description   = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'Description'
                Name          = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'productName'
                LicenseFamily = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'Family'
            }
        }
        $ID_PKEY = $LicensingProducts | ? Name -NotMatch 'ESU' | ? Description -NotMatch 'ESU' | select -First 1
        [XML]$licenseData = Get-LicenseDetails $ID_PKEY.ID -ReturnRawData
        $Branding = $licenseData.licenseGroup.license[1].otherInfo.infoTables.infoList.infoStr | ? Name -EQ win:branding

        $ID_PKEY.LicenseFamily
        $Branding.'#text'
    #>

    # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ! ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    <#
        Retrieves Windows Product Policy values from the registry
        HKLM:\SYSTEM\CurrentControlSet\Control\ProductOptions -> ProductPolicy
    #>

    $KernelEdition = Get-ProductPolicy -Filter Kernel-EditionName -UseApi | select -ExpandProperty Value
    #$KernelEdition = Get-ProductPolicy -Filter Kernel-EditionName | ? Name -Match 'Kernel-EditionName' | select -ExpandProperty Value
    if ($KernelEdition -and (-not [string]::IsNullOrWhiteSpace($KernelEdition))) {
        return $KernelEdition
    }

    <#
        Extract Edition Info from Registry -> 
            DigitalProductId4
    #>

    # DigitalProductId4, WCHAR szEditionType[260];
    $DigitalProductId4 = Parse-DigitalProductId4
    if ($DigitalProductId4 -and $DigitalProductId4.EditionType -and 
        (-not [String]::IsNullOrWhiteSpace($DigitalProductId4.EditionType))) {
        return $DigitalProductId4.EditionType
    }

    <#
        Use RtlGetProductInfo to get brand info, And convert the value

        Alternative, 
        * HKLM:\SYSTEM\CurrentControlSet\Control\ProductOptions, ProductPolicy
        * Get-ProductPolicy -> Read 'Kernel-BrandingInfo' -or 'Kernel-ProductInfo' -> Value
          Get-ProductPolicy | ? name -Match "Kernel-BrandingInfo|Kernel-ProductInfo" | select -First 1 -ExpandProperty Value
          which i believe, the source data of the function
        * Win32_OperatingSystem Class -> OperatingSystemSKU
          which i believe, call -> RtlGetProductInfo
        * Also, this registry value --> 
          HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions->OSProductPfn
    #>
    try {
        <#
        -- It call ZwQueryLicenseValue -> Kernel-BrandingInfo \ Kernel-ProductInfo
        -- Replace with direct call ...

        [UInt32]$BrandingInfo = 0
        $status = $Global:ntdll::RtlGetProductInfo(
            $OperatingSystemInfo.dwOSMajorVersion,
            $OperatingSystemInfo.dwOSMinorVersion,
            $OperatingSystemInfo.dwSpMajorVersion,
            $OperatingSystemInfo.dwSpMinorVersion,
            [Ref]$BrandingInfo)

        if (!$status) {
            throw }

        # Get Branding info Number of current Build
        [INT]$BrandingInfo
        #>

        [INT]$BrandingInfo = Get-ProductPolicy |
            ? name -Match "Kernel-BrandingInfo|Kernel-ProductInfo" |
                ? Value | Select -First 1 -ExpandProperty Value

        # Get editionID Name using hard coded table,
        # provide by abbodi1406 :)
        $match = $Global:productTypeTable | Where-Object {
            [Convert]::ToInt32($_.DWORD, 16) -eq $BrandingInfo
        }
        if ($match) {
            return $match.ProductID
        }

        # using API to convert from BradingInfo to EditionName
        $editionIDPtr = [IntPtr]::Zero
        $hresults = $Global:PKHElper::GetEditionNameFromId(
            $BrandingInfo, [ref]$editionIDPtr)
        if ($hresults -eq 0) {
            $editionID = [Marshal]::PtrToStringUni($editionIDPtr)
            return $editionID
        }
    }
    catch { }
    Finally {
        New-IntPtr -hHandle $productTypePtr -Release
    }

    # Key: HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion Propery: EditionID
    $EditionID = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name EditionID -ea 0).EditionID
    if ($EditionID) {
        return $EditionID }

    Clear-host
    Write-Host
    write-host "Failed to Edition ID version from all methods."
    Write-Host
    read-host
    exit 1
}

<#
.SYNOPSIS
Retrieves Windows edition upgrade paths and facilitates interactive upgrades.
[Bool] USeApi Option, Only available for > *Current* edition

.EXAMPLE
# Target edition for current system:
Get-EditionTargetsFromMatrix
Get-EditionTargetsFromMatrix -UseApi

.EXAMPLE
# Target edition for a specific ID (e.g., 'EnterpriseSN'):
Get-EditionTargetsFromMatrix -EditionID 'EnterpriseSN'
Get-EditionTargetsFromMatrix -EditionID 'EnterpriseSN' -RawData

.EXAMPLE
# Upgrade from the current version (interactive selection):
Get-EditionTargetsFromMatrix -UpgradeFrom

.EXAMPLE
# Upgrade from a specific base version (e.g., 'EnterpriseSN' or 'CoreCountrySpecific'):
Get-EditionTargetsFromMatrix -UpgradeFrom -EditionID 'EnterpriseSN'
Get-EditionTargetsFromMatrix -UpgradeFrom -EditionID 'CoreCountrySpecific'

.EXAMPLE
# Upgrade to any chosen version (interactive product key selection):
Get-EditionTargetsFromMatrix -UpgradeTo

.EXAMPLE
# Upgrade to a specific edition (e.g., 'EnterpriseSEval' with product key selection):
Get-EditionTargetsFromMatrix -EditionID EnterpriseSEval -UpgradeTo

.EXAMPLE
# List all available editions:
Get-EditionTargetsFromMatrix -ReturnEditionList

--------------------------------------------------

PowerShell, Also support this function apparently
--> Get-WindowsEdition -Online -Target .!

#>
function Get-EditionTargetsFromMatrix {
    param (
        [Parameter(Mandatory = $false)]
        [ValidateSet(
            "ultimate","homebasic","homepremium","enterprise","homebasicn","business","serverstandard","serverdatacenter","serversbsstandard","serverenterprise","starter",
            "serverdatacentercore","serverstandardcore","serverenterprisecore","serverenterpriseia64","businessn","serverweb","serverhpc","serverhomestandard","serverstorageexpress",
            "serverstoragestandard","serverstorageworkgroup","serverstorageenterprise","serverwinsb","serversbspremium","homepremiumn","enterprisen","ultimaten","serverwebcore",
            "servermediumbusinessmanagement","servermediumbusinesssecurity","servermediumbusinessmessaging","serverwinfoundation","serverhomepremium","serverwinsbv","serverstandardv",
            "serverdatacenterv","serverenterprisev","serverdatacentervcore","serverstandardvcore","serverenterprisevcore","serverhypercore","serverstorageexpresscore","serverstoragestandardcore",
            "serverstorageworkgroupcore","serverstorageenterprisecore","startern","professional","professionaln","serversolution","serverforsbsolutions","serversolutionspremium",
            "serversolutionspremiumcore","serversolutionem","serverforsbsolutionsem","serverembeddedsolution","serverembeddedsolutioncore","professionalembedded","serveressentialmanagement",
            "serveressentialadditional","serveressentialmanagementsvc","serveressentialadditionalsvc","serversbspremiumcore","serverhpcv","embedded","startere","homebasice",
            "homepremiume","professionale","enterprisee","ultimatee","enterpriseeval","prerelease","servermultipointstandard","servermultipointpremium","serverstandardeval",
            "serverdatacentereval","prereleasearm","prereleasen","enterpriseneval","embeddedautomotive","embeddedindustrya","thinpc","embeddeda","embeddedindustry","embeddede",
            "embeddedindustrye","embeddedindustryae","professionalplus","serverstorageworkgroupeval","serverstoragestandardeval","corearm","coren","corecountryspecific","coresinglelanguage",
            "core","professionalwmc","mobilecore","embeddedindustryeval","embeddedindustryeeval","embeddedeval","embeddedeeval","coresystemserver","servercloudstorage","coreconnected",
            "professionalstudent","coreconnectedn","professionalstudentn","coreconnectedsinglelanguage","coreconnectedcountryspecific","connectedcar","industryhandheld",
            "ppipro","serverarm64","education","educationn","iotuap","serverhi","enterprises","enterprisesn","professionals","professionalsn","enterpriseseval",
            "enterprisesneval","iotuapcommercial","mobileenterprise","analogonecore","holographic","professionalsinglelanguage","professionalcountryspecific","enterprisesubscription",
            "enterprisesubscriptionn","serverdatacenternano","serverstandardnano","serverdatacenteracor","serverstandardacor","serverdatacentercor","serverstandardcor","utilityvm",
            "serverdatacenterevalcor","serverstandardevalcor","professionalworkstation","professionalworkstationn","serverazure","professionaleducation","professionaleducationn",
            "serverazurecor","serverazurenano","enterpriseg","enterprisegn","businesssubscription","businesssubscriptionn","serverrdsh","cloud","cloudn","hubos","onecoreupdateos",
            "cloude","andromeda","iotos","clouden","iotedgeos","iotenterprise","modernpc","iotenterprises","systemos","nativeos","gamecorexbox","gameos","durangohostos",
            "scarletthostos","keystone","cloudhost","cloudmos","cloudcore","cloudeditionn","cloudedition","winvos","iotenterprisesk","iotenterprisek","iotenterpriseseval",
            "agentbridge","nanohost","wnc","serverazurestackhcicor","serverturbine","serverturbinecor"
        )]
        [string]$EditionID = $null,

        [Parameter(Mandatory = $false)]
        [switch]$ReturnEditionList,

        [Parameter(Mandatory = $false)]
        [switch]$UpgradeFrom,

        [Parameter(Mandatory = $false)]
        [switch]$UpgradeTo,

        [Parameter(Mandatory = $false)]
        [switch]$UseApi,

        [Parameter(Mandatory = $false)]
        [switch]$RawData
    )

    $targets = @();

    [string]$xmlPath = "C:\Windows\servicing\Editions\EditionMappings.xml"
    [string]$MatrixPath = "C:\Windows\servicing\Editions\EditionMatrix.xml"
    if (-not $xmlPath -or -not $MatrixPath) {
        Write-Host
         Write-Warning "Required files not found: `n$xmlPath`n$MatrixPath"
        return
    }
    $CurrentEdition = Get-ProductID
    if ($UseApi -and (
            $EditionID -and ($EditionID -ne $CurrentEdition))) {
        Write-Warning "UseApi Only for Current edition."
        return @()
    }

    function Find-Upgrades {
        param (
            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$EditionID
        )

        if ($EditionID -and ($Global:productTypeTable.ProductID -notcontains $EditionID)) {
            Write-Warning "EditionID '$EditionID' is not found in the product type table."
            return
        }
    
        $parentEdition = $null
        $relatedEditions = @()

        [xml]$xml = Get-Content -Path $xmlPath
        $WindowsEditions = $xml.WindowsEditions.Edition

        $isVirtual = $WindowsEditions.Name -contains $EditionID
        $isParent = $WindowsEditions.ParentEdition -contains $EditionID

        # If the selected edition is a Virtual Edition, get the Parent Edition
        if ($isVirtual) {
            $selectedEditionNode = $WindowsEditions | Where-Object { $_.Name -eq $EditionID }
            $parentEdition = $selectedEditionNode.ParentEdition
        }

        # If the edition is a Parent Edition, find all related Virtual Editions
        if ($isParent) {
            try {
                $relatedEditions = $WindowsEditions | Where-Object { $_.ParentEdition -eq $EditionID -and $_.virtual -eq "true" }
            }
            catch {
            }
        }

        # If the edition is a Virtual Edition, find all other Virtual Editions linked to the same Parent Edition
        if ($isVirtual) {
            try {
                $relatedEditions += $WindowsEditions | Where-Object { $_.ParentEdition -eq $parentEdition -and $_.virtual -eq "true" }
            }
            catch {
            }
        }

        return [PSCustomObject]@{
            Editions = $relatedEditions | Select-Object -ExpandProperty Name
            Parent   = $parentEdition
        }
    }
    Function Dism-GetTargetEditions {
        try {
            $hr = $Global:DismAPI::DismInitialize(
                0, [IntPtr]::Zero, [IntPtr]::Zero)
            if ($hr -ne 0) {
                Write-Warning "DismInitialize failed: $hr"
                return @()
            }

            $session = [IntPtr]::Zero
            $hr = $Global:DismAPI::DismOpenSession(
                "DISM_{53BFAE52-B167-4E2F-A258-0A37B57FF845}", [IntPtr]::Zero, [IntPtr]::Zero, [ref]$session)
            if ($hr -ne 0) { 
                Write-Warning "DismOpenSession failed: $hr"
                return
            }

            $count = 0
            $editionIds = [IntPtr]::Zero
            $hr = $Global:DismAPI::_DismGetTargetEditions($session, [ref]$editionIds, [ref]$count)
            if ($hr -ne 0) { 
                Write-Warning "_DismGetTargetEditions failed: $hr"
            }

            if ($hr -eq 0 -and $count -gt 0) {
                try {
                    return Convert-PointerArrayToStrings -PointerToArray $editionIds -Count $count
                }
                catch {
                    Write-Warning "Failed to convert editions: $_"
                    return @()
                }
            }
        }
        catch {
        }
        finally {
            if ($editionIds -and (
                $editionIds -ne [IntPtr]::Zero)) {
                    $null = $Global:DismAPI::DismDelete($editionIds)
            }
            if ($session -and (
                $session -ne [IntPtr]::Zero)) {
                    $null = $Global:DismAPI::DismCloseSession($session)
            }
            $null = $Global:DismAPI::DismShutdown()
        }
    }
    function Convert-PointerArrayToStrings {
        param (
            [Parameter(Mandatory = $true)]
            [IntPtr] $PointerToArray,

            [Parameter(Mandatory = $true)]
            [UInt32] $Count
        )

        if ($PointerToArray -eq [IntPtr]::Zero -or $Count -eq 0) {
            return @()
        }

        $strings = @()
        for ($i = 0; $i -lt $Count; $i++) {
            # Calculate pointer to pointer at index $i
            $ptrToStringPtr = [IntPtr]::Add($PointerToArray, $i * [IntPtr]::Size)

            # Read the string pointer
            $stringPtr = [Marshal]::ReadIntPtr($ptrToStringPtr)
            if ($stringPtr -ne [IntPtr]::Zero) {
                # Read the Unicode string from the pointer
                $edition = [Marshal]::PtrToStringUni($stringPtr)
                $strings += $edition
            }
        }
        return $strings
    }

    if (-Not (Test-Path $MatrixPath)) {
        Write-Warning "EditionMatrix.xml not found at $MatrixPath"
        return
    }
    if ($EditionID -and ($Global:productTypeTable.ProductID -notcontains $EditionID)) {
        Write-Warning "EditionID '$EditionID' is not found in the product type table."
        return
    }
    [xml]$xml = Get-Content $MatrixPath
    $LicensingProducts = Get-SLIDList -eQueryIdType SL_ID_PRODUCT_SKU -eReturnIdType SL_ID_PRODUCT_SKU | % {
        [PSCustomObject]@{
            ID            = $_
            Description   = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'Description'
            Name          = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'productName'
            LicenseFamily = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'Family'
        }
    }
    $uniqueFamilies = $LicensingProducts.LicenseFamily | Select-Object -Unique
    
    if ($UpgradeFrom) {
        if (-not $EditionID) {
            $EditionID = $CurrentEdition
        }
        if (-not $EditionID) {
            Write-Host
            Write-Warning "EditionID is missing. Upgrade may not proceed correctly."
            return
        }

        # Find edition node in XML
        # Editions where this ID is the source (normal lookup)
        $sourceNode = $xml.TmiMatrix.Edition | Where-Object { $_.ID -eq $EditionID }

        # Editions where this ID is a target (reverse lookup)
        $targetNodes = $xml.TmiMatrix.Edition | Where-Object {
            $_.Target.ID -contains $EditionID
        }

        # Combine all
        $editionNode = @()
        if ($sourceNode) { $editionNode += $sourceNode }
        if ($targetNodes) { $editionNode += $targetNodes }
        $Upgrades = Find-Upgrades -EditionID $EditionID

        if ($UseApi -and (
            $EditionID -eq $CurrentEdition)) {
                $targetEditions = Dism-GetTargetEditions
        }
        else {
            if ($editionNode.Target.ID) {
                $targetEditions += $editionNode.Target.ID
            }

            if ($Upgrades.Editions) {
                $targetEditions += $Upgrades.Editions
            }

            if ($Upgrades.Parent) {
                $targetEditions += $Upgrades.Parent
            }

            $targetEditions = $targetEditions | ? { $_ -ne $CurrentEdition} | select -Unique
            if (-not $targetEditions) {
                Write-Host
                Write-Warning "No upgrade targets found for EditionID '$EditionID'."
                return
            }
            $targetEditions = $targetEditions | Where-Object { $uniqueFamilies  -contains $_ } | select -Unique
        }

        if ($targetEditions.Count -eq 0) {
            Write-Host
            Write-Warning "No targets license's found in Current system for '$EditionID' Edition."
            return
        }
        elseif ($targetEditions.Count -eq 1) {
            $chosenTarget = $targetEditions
            Write-Host
            Write-Warning "Only one upgrade target found: $chosenTarget. Selecting automatically."
        } else {
            # Multiple targets: let user choose
            $chosenTarget = $null
            while (-not $chosenTarget) {
                Clear-Host
                Write-Host
                Write-Host "[Available upgrade targets]"
                Write-Host
                for ($i = 0; $i -lt $targetEditions.Count; $i++) {
                    Write-Host "[$($i+1)] $($targetEditions[$i])"
                }
                $selection = Read-Host "Select upgrade target edition by number (or 'q' to quit)"
                if ($selection -eq 'q') { break }

                $parsedSelection = 0
                if ([int]::TryParse($selection, [ref]$parsedSelection) -and
                    $parsedSelection -ge 1 -and
                    $parsedSelection -le $targetEditions.Count) {
                    $chosenTarget = $targetEditions[$parsedSelection - 1]
                } else {
                    Write-Host "Invalid selection, please try again."
                }
            }
        }

        if (-not $chosenTarget) {
            Write-Host
            Write-Warning "No target edition selected. Cancelling."
            return
        }

        $UpgradeTo = $true
        $EditionID = $chosenTarget
    }
    if ($UpgradeTo) {
        $filteredKeys = $Global:PKeyDatabase | ? { $LicensingProducts.ID -contains $_.ActConfigId}
        if ($EditionID) {
            if ($EditionID -eq $CurrentEdition) {
                Write-Host
                Write-Warning "Attempting to upgrade to the same edition ($EditionID) already installed. No upgrade needed."
                return
            }
            $matchingKeys = @($filteredKeys | Where-Object { $_.EditionId -eq $EditionID })
            if (-not $matchingKeys -or $matchingKeys.Count -eq 0) {
                Write-Host
                Write-Warning "No matching keys found for EditionID '$EditionID'"
                return
            }
        } else {
            # No EditionID specified, use all keys
            $matchingKeys = @($filteredKeys)
        }

        if (-not $matchingKeys -or $matchingKeys.Count -eq 0) {
            Write-Host
            Write-Warning "No product keys available."
            return
        }

        if ($matchingKeys.Count -gt 1) {
            # Multiple keys: show Out-GridView for selection
            $selectedKey = $null
            while (-not $selectedKey) {
                Clear-Host
                Write-Host
                Write-Host "[Available product keys]"
                Write-Host

                for ($i = 0; $i -lt $matchingKeys.Count; $i++) {
                    $item = $matchingKeys[$i]
                    Write-Host ("{0,-4} {1,-30} | {2,-50} | {3,-15} | {4}" -f ("[$($i+1)]"), $item.EditionId, $item.ProductDescription, $item.ProductKeyType, $item.RefGroupId)
                }

                Write-Host
                $input = Read-Host "Select a product key by number (or 'q' to quit)"
                if ($input -eq 'q') { break }

                $parsed = 0
                if ([int]::TryParse($input, [ref]$parsed) -and $parsed -ge 1 -and $parsed -le $matchingKeys.Count) {
                    $selectedKey = $matchingKeys[$parsed - 1]
                } else {
                    Write-Host "Invalid selection. Please try again."
                }
            }            if (-not $selectedKey) {
                Write-Host
                Write-Warning "No selection made. Operation cancelled."
                return
            }
        }
        elseif ($matchingKeys.Count -eq 1) {
            # Only one key: select automatically
            $selectedKey = $matchingKeys
        }
        else {
            Write-Host
            Write-Warning "No product keys available."
            return
        }
        if (-not $selectedKey) {
            Write-Host
            Write-Warning "No selection made. Operation cancelled."
            return
        }

        # Simulated Key Installation
        Write-Host
        SL-InstallProductKey -Keys @(
            (Encode-Key $($selectedKey.RefGroupId) 0 0))

        return
    }
    if (-not $EditionID) {
        $EditionID = Get-ProductID
    }
    if ($ReturnEditionList) {
        return $xml.TmiMatrix.Edition | Select-Object -ExpandProperty ID
    }

    if ($UseApi -and (
        $EditionID -eq $CurrentEdition)) {
            $targets = Dism-GetTargetEditions
    }
    else {
        $Upgrades = Find-Upgrades -EditionID $EditionID
        $editionNode = $xml.TmiMatrix.Edition | Where-Object { $_.ID -eq $EditionID }
        
        if ($editionNode.Target.ID) {
            $targets += $editionNode.Target.ID
        }

        if ($Upgrades.Editions) {
            $targets += $Upgrades.Editions
        }

        if ($Upgrades.Parent) {
            $targets += $Upgrades.Parent
        }
        $FilterList = @($CurrentEdition, $EditionID)
        $targets = $targets | ? {$_ -notin $FilterList} | Sort-Object -Unique
    }
    if ($targets) {
        if ($RawData) {
            return $targets
        }
        Write-Host
        Write-Host "Edition '$EditionID' can be upgraded/downgraded to:" -ForegroundColor Green
        $targets | ForEach-Object { Write-Host "  - $_" }
    } else {
        if ($RawData) {
            return @()
        }
        Write-Host
        Write-Warning "Edition '$EditionID' has no defined upgrade targets."
    }
}

<#
Retrieves Windows Product Policy values from the registry.
Supports filtering by policy names or returns all by default.

Adapted from Windows Product Policy Editor by kost:
https://forums.mydigitallife.net/threads/windows-product-policy-editor.39411/

Software Licensing
https://www.geoffchappell.com/studies/windows/km/ntoskrnl/api/ex/slmem/index.htm?tx=57,58

Windows Vista introduces a formal scheme of named license values,
with API functions to manage them. The license values are stored together -
as binary data for a single registry value. The data format is presented separately.
Like registry values, each license value has its own data.

Windows Internals Book 7th Edition Tools
https://github.com/zodiacon/WindowsInternals/blob/master/SlPolicy/SlPolicy.cpp

windows Sdk
https://github.com/mic101/windows/blob/master/WRK-v1.2/public/internal/base/inc/zwapi.h

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct UNICODE_STRING {
    public ushort Length;
    public ushort MaximumLength;
    public IntPtr Buffer;
}

public static class Ntdll {
    [DllImport("ntdll.dll", CharSet = CharSet.Unicode)]
    public static extern int ZwQueryLicenseValue(
        ref UNICODE_STRING ValueName,
        out uint Type,
        IntPtr Data,
        uint DataSize,
        out uint ResultDataSize
    );
}
#>
function Get-ProductPolicy {
    param (
        [Parameter(Mandatory=$false)]
        [string[]]$Filter = @(),
        
        [Parameter(Mandatory=$false)]
        [switch]$OutList,

        [Parameter(Mandatory=$false)]
        [switch]$UseApi
    )

    if ($OutList -eq $true) {
        # Oppsite XOR Case, Validate it Not both
        if (-not ([bool]$filter -xor [bool]$OutList)) {
            Write-Warning "Can't use both -OutList and -Filter"
            return
        }

        # Oppsite XOR Case, Validate it Not both
        if (-not ($UseApi -xor $OutList)) {
            Write-Warning "Can't use both -OutList and -UseApi"
            return
        }
    }

    if ($UseApi -and (-not $Filter -or $Filter.Count -eq 0)) {
        Write-Warning "API mode requires at least one value name in -Filter."
        return $null
    }

    $results = @()

    if ($UseApi) {
        foreach ($valueName in $Filter) {
            try {
                [uint32]$type = 0
                [uint32]$resultSize = 0
                $unicodeStringPtr = Init-NativeString -Value $valueName -Encoding Unicode

                # Allocate a buffer to receive the value (arbitrary size like 3 KB)
                $dataSize = 3000
                $dataBuffer = [Marshal]::AllocHGlobal($dataSize)

                try {
                    $status = $Global:ntdll::ZwQueryLicenseValue(
                        $unicodeStringPtr,
                        [ref]$type,
                        $dataBuffer,
                        [uint32]$dataSize,
                        [ref]$resultSize
                    )

                    if ($status -eq 0) {
                        $result = [PSCustomObject]@{
                            Name  = $valueName
                            Type  = $type
                            Size  = $resultSize
                            Value = $null
                        }

                        # The optional Type argument provides the address of a variable that is to receive the type of data:
                        # REG_SZ (0x01) for a string; REG_BINARY (0x03) for binary data; REG_DWORD (0x04) for a dword.
                        # which is execly same 0000-------->>>>>>>>> SLDATATYPE enumeration (slpublic.h) 0000----->>> Info:
                        
                        <#
                        typedef enum _tagSLDATATYPE {
                            SL_DATA_NONE = REG_NONE,
                            SL_DATA_SZ = REG_SZ,
                            SL_DATA_DWORD = REG_DWORD,
                            SL_DATA_BINARY = REG_BINARY,
                            SL_DATA_MULTI_SZ,
                            SL_DATA_SUM = 100
                        } SLDATATYPE;
                        #>

                        $result.Value = Parse-RegistryData -dataType $type -ptr $dataBuffer -valueSize $resultSize -valueName $valueName
                        $results += $result
                    }
                    else {
                        $statusHex = "0x{0:X}" -f $status

                        switch ($statusHex) {
                            "0x00000000" {
                                # success - no warning needed
                            }
                            "0xC0000272" {
                                Write-Warning "Failed to query '$valueName' via UseApi: License quota exceeded or unsupported. Status: $statusHex"
                                break
                            }
                            "0xC0000023" {
                                Write-Warning "ZwQueryLicenseValue failed for '$valueName': Invalid handle. Status: $statusHex"
                                break
                            }
                            "0xC0000034" {
                                Write-Warning "ZwQueryLicenseValue failed for '$valueName': Value not found. Status: $statusHex"
                                break
                            }
                            "0xC000001D" {
                                Write-Warning "ZwQueryLicenseValue failed for '$valueName': Not implemented (API not supported). Status: $statusHex"
                                break
                            }
                            "0xC00000BB" {
                                Write-Warning "ZwQueryLicenseValue failed for '$valueName': Operation not supported. Status: $statusHex"
                                break
                            }
                            default {
                                Write-Warning "ZwQueryLicenseValue failed for '$valueName' with status: $statusHex"
                            }
                        }
                    }
                }
                finally {
                    if ($dataBuffer -ne [IntPtr]::Zero) { 
                        [Marshal]::FreeHGlobal($dataBuffer) }
                }
            }
            finally {
                if ($unicodeStringPtr -ne [IntPtr]::Zero) {
                    Free-NativeString -StringPtr $unicodeStringPtr
                }
            }
        }

        return $results
    }

    $policyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\ProductOptions"
    $blob = (Get-ItemProperty -Path $policyPath -Name ProductPolicy).ProductPolicy
    if (-not $blob) {
        Write-Warning "ProductPolicy blob not found in registry."
        return $null
    }

    function Read-UInt16($bytes, $offset) {
        if ($offset + 2 -gt $bytes.Length) { return 0 }
        return [BitConverter]::ToUInt16($bytes, $offset)
    }

    function Read-UnicodeString($bytes, $offset, $length) {
        $length = $length -band 0xFFFE  # ensure even length
        if ($offset + $length -gt $bytes.Length) { return "" }
        $str = [System.Text.Encoding]::Unicode.GetString($bytes, $offset, $length)
        $nullIndex = $str.IndexOf([char]0)
        if ($nullIndex -ge 0) {
            $str = $str.Substring(0, $nullIndex)
        }
        return $str.Trim()
    }

    $offset = 20  # skip header
    $entryHeaderSize = 16

    $results = @()

    while ($offset -lt $blob.Length) {
        if ($offset + $entryHeaderSize -gt $blob.Length) { break }

        $cbSize = Read-UInt16 $blob $offset
        $cbName = Read-UInt16 $blob ($offset + 2)
        $type   = Read-UInt16 $blob ($offset + 4)
        $cbData = Read-UInt16 $blob ($offset + 6)

        $nameOffset = $offset + $entryHeaderSize
        $dataOffset = $nameOffset + $cbName

        if ($dataOffset + $cbData -gt $blob.Length) {
            break
        }

        $name = Read-UnicodeString $blob $nameOffset $cbName

        # The data type follows the familiar scheme for registry data:
        # REG_SZ (0x01) for a string,
        # REG_BINARY (0x03) for binary data
        # and REG_DWORD (0x04) for a dword.

        # If Filter empty => get all, else filter by name
        if (($Filter.Count -eq 0) -or ($Filter -contains $name)) {

            # Use Parse-RegistryData to parse the value, passing blob and offset
            # $val = Parse-RegistryData -dataType $type -blob $blob -dataOffset $dataOffset -valueSize $cbData -valueName $name
            
            switch ($type) {
                0 {  
                    # SL_DATA_NONE
                    $val = $null
                }

                1 {  
                    # SL_DATA_SZ (Unicode string)
                    $val = [System.Text.Encoding]::Unicode.GetString($blob, $dataOffset, $cbData).TrimEnd([char]0)
                }

                4 {  
                    # SL_DATA_DWORD
                    if ($cbData -ne 4) {
                        $val = $null
                    } else {
                        $val = [BitConverter]::ToUInt32($blob, $dataOffset)
                    }
                }

                3 {  
                    # SL_DATA_BINARY
                    if ($name -match "Security-SPP-LastWindowsActivationTime" -and $cbData -eq 8) {
                        $fileTime = [BitConverter]::ToInt64($blob, $dataOffset)
                        $val = [DateTime]::FromFileTimeUtc($fileTime)
                    }
                    elseif ($name -eq "Security-SPP-LastWindowsActivationHResult" -and $cbData -eq 4) {
                        $val = [BitConverter]::ToUInt32($blob, $dataOffset)
                    }
                    else {
                        $bytes = New-Object byte[] $cbData
                        [System.Buffer]::BlockCopy($blob, $dataOffset, $bytes, 0, $cbData)
                        $val = [BitConverter]::ToString($bytes)
                    }
                }

                7 {  
                    # SL_DATA_MULTI_SZ
                    $raw = [System.Text.Encoding]::Unicode.GetString($blob, $dataOffset, $cbData)
                    $val = $raw -split "`0" | Where-Object { $_ -ne '' }
                }

                100 {  
                    # SL_DATA_SUM
                    $val = $null
                }

                default {
                    $val = $null
                }
            }

            $results += [PSCustomObject]@{
                Name  = $name
                Type  = $type
                Value = $val
            }
        }

        if ($cbSize -lt $entryHeaderSize) { break }

        $offset += $cbSize
    }

    if ($OutList) {
        return $results.Name
    }
    return $results
}

<#
Alternative call instead of, 
SoftwareLicensingService --> OA3xOriginalProductKey

~~~~~~~~~~~~~~~~~~~~

Evasions: Firmware tables
https://evasions.checkpoint.com/src/Evasions/techniques/firmware-tables.html

typedef struct _SYSTEM_FIRMWARE_TABLE_INFORMATION {
    ULONG ProviderSignature;
    SYSTEM_FIRMWARE_TABLE_ACTION Action;
    ULONG TableID;
    ULONG TableBufferLength;
    UCHAR TableBuffer[ANYSIZE_ARRAY];  // <- the result will reside in this field
} SYSTEM_FIRMWARE_TABLE_INFORMATION, *PSYSTEM_FIRMWARE_TABLE_INFORMATION;

// helper enum
typedef enum _SYSTEM_FIRMWARE_TABLE_ACTION
{
    SystemFirmwareTable_Enumerate,
    SystemFirmwareTable_Get
} SYSTEM_FIRMWARE_TABLE_ACTION, *PSYSTEM_FIRMWARE_TABLE_ACTION;

~~~~~~~~~~~~~~~~~~~~

UINT __stdcall GetSystemFirmwareTable(
        DWORD FirmwareTableProviderSignature,
        DWORD FirmwareTableID,
        PVOID pFirmwareTableBuffer,
        DWORD BufferSize)

Heap = RtlAllocateHeap(NtCurrentPeb()->ProcessHeap, KernelBaseGlobalData, BufferSize + 16);
Heap[0] = FirmwareTableProviderSignature; // FirmwareTableProviderSignature
Heap[1] = 1;                              // Action -- 1
Heap[2] = FirmwareTableID;                // FirmwareTableID
Heap[3] = BufferSize;                     // Payload Only

v8 = BufferSize + 16;                     // HeadSize (16) & Payload Size
v11 = NtQuerySystemInformation(0x4C, Heap, v8, &ReturnLength);

So, what happen here, 
* Header  = 16 Byte's
* Heap[3] = PayLoad size, Only!
Allocate --> Header Size & Payload Size. ( DWORD BufferSize & 16 bytes above )
Set Heap[3] --> Payload Size only. ( DWORD BufferSize )
NT! Api call, Total length. ( DWORD BufferSize + 16 )

Case Fail!
 --> Heap[3] = 0
 --> ReturnLength = 16,
 * Return --> Heap[3] --> `0 (Not ReturnLength!)

~~~~~~~~~~~~~~~~~~~~

__kernel_entry NTSTATUS NtQuerySystemInformation(
  [in]            SYSTEM_INFORMATION_CLASS SystemInformationClass,
  [in, out]       PVOID                    SystemInformation,
  [in]            ULONG                    SystemInformationLength,
  [out, optional] PULONG                   ReturnLength

~~~~~~~~~~~~~~~~~~~~

FirmwareTables by vurdalakov
https://github.com/vurdalakov/firmwaretables

get_win8key by Christian Korneck
https://github.com/christian-korneck/get_win8key

ctBiosKey.cpp
https://gist.github.com/hosct/456055c0eec4e71bb504489410ed7fb6#file-ctbioskey-cpp

[C++, C#, VB.NET, PowerShell] Read MSDM license information from BIOS ACPI tables | My Digital Life Forums
https://forums.mydigitallife.net/threads/c-c-vb-net-powershell-read-msdm-license-information-from-bios-acpi-tables.43788/

ACPI Tables
https://www.kernel.org/doc/html/next/arm64/acpi_object_usage.html

Microsoft Software Licensing Tables (SLIC and MSDM)
https://learn.microsoft.com/en-us/previous-versions/windows/hardware/design/dn653305(v=vs.85)?redirectedfrom=MSDN

ACPI Software Programming Model
https://uefi.org/htmlspecs/ACPI_Spec_6_4_html/05_ACPI_Software_Programming_Model/ACPI_Software_Programming_Model.html#system-description-table-header

var table = FirmwareTables.GetAcpiTable("MDSM");
var productKeyLength = (int)table.GetPayloadUInt32(16); // offset 52
var productKey = table.GetPayloadString(20, productKeyLength); // offset 56 > Till End
Console.WriteLine("OEM Windows product key: '{0}'", productKey);

Example Code:
~~~~~~~~~~~~

Clear-Host

Write-Host
Write-Host "Get-OA3xOriginalProductKey" -ForegroundColor Green
Get-OA3xOriginalProductKey

Write-Host
Write-Host "Get-ServiceInfo" -ForegroundColor Green
Get-ServiceInfo -loopAllValues | Format-Table -AutoSize

Write-Host
Write-Host "Get-ActiveLicenseInfo" -ForegroundColor Green
Get-ActiveLicenseInfo | Format-List
);

#>
function Get-OA3xOriginalProductKey {
    
    function String-ToHex {
        param (
            [char[]]$chars,
            [switch]$rev,
            [switch]$hex2int
        )

        if ([string]::IsNullOrEmpty($chars)) {
            throw "String is null -or empty"
        }

        if ($rev) {
            [Array]::Reverse($chars)
        }

        if ($hex2int) {
            return ($chars | ForEach-Object `
                -Begin { $hexValue = New-Object System.Text.StringBuilder(256) } `
                -Process { $hexValue.AppendFormat('{0:X2}', [byte]$_) | Out-Null } `
                -End { [Convert]::ToUInt32("0x$($hexValue.ToString())", 16) })
        } else {
            return ($chars | ForEach-Object `
                -Begin {$val = 0} `
                -Process {$val = (($val -shl 8) -bor [byte]$_)} `
                -End { $val })
        }
    }

    # Constants
    [UInt32]$Provider = String-ToHex ACPI #-hex2int
    [UInt32]$TableId  = String-ToHex MSDM -rev #-hex2int
    $PayLoad, $HeaderSize, $tableInfo, $table_Get = 0x00, 0x10, 0x4C, 0x01

    # First call: dummy buffer
    $buffer  = New-IntPtr -Size $HeaderSize
    @($Provider, $table_Get, $TableId, $PayLoad) | % `
        -Begin { $i=-1 } `
        -Process { [Marshal]::WriteInt32($buffer, (++$I*4), [int]$_)}

    [int]$returnLen = 0
    $status = $Ntdll::NtQuerySystemInformation(
        $tableInfo, $buffer, $HeaderSize, [ref]$returnLen
    )

    $PayLoad = [Marshal]::ReadInt32($buffer, 0xC)
    Free-IntPtr -handle $buffer
    
    # So, if you have OEM information in UEFI,
    # this check will succeed, and results >0, >16
    # $returnLen should be at least, 
    # 16 for Header & 56 for Base & 29 for CD KEy.!
    if ($PayLoad -le 0 -or (
        $returnLen -le $HeaderSize)) {
            return $null 
    }

    # Second call: real buffer
    $buffer = New-IntPtr -Size ($HeaderSize + $PayLoad)
    @($Provider, $table_Get, $TableId, $PayLoad) | % `
        -Begin { $i=-1 } `
        -Process { [Marshal]::WriteInt32($buffer, (++$I*4), [int]$_)}
    
    try {
        [int]$returnLen = 0
        if (0 -ne $Ntdll::NtQuerySystemInformation(
           $tableInfo, $buffer, ($HeaderSize + $PayLoad), [ref]$returnLen)) {
              return $null
        }

        # memcpy_0(buffer, v10+4, v10[3]);
        # v10[0-3] => 16 bytes, v10[4-?] => Rest bytes, v10[3] => Payload Size
        $pkey = $null
        $pkLen = [Marshal]::ReadInt32($buffer, ($HeaderSize + 0x34))
        if ($pkLen -gt 0) {
           $pkey = [Marshal]::PtrToStringAnsi(
              [IntPtr]::Add($buffer, ($HeaderSize + 0x38)), $pkLen)
        }
        return $pkey
    }
    finally {
        Free-IntPtr -handle $buffer
    }
}

<#
.SYNOPSIS
Get Ubr value.

#>
function Scan-FolderWithAPI {
    param(
        [string]$folder
    )
    
    $maxUBR = $null
    $bufferSize = 592
    $cFileNameOffset = 44
    $regex = [regex]'10\.0\.\d+\.(\d+)'
    $wildcard = "$folder\*-edition*10.*.*.*"

    $pBuffer = [Marshal]::AllocHGlobal($bufferSize)
    $Global:ntdll::RtlZeroMemory($pBuffer,[UIntPtr]::new($bufferSize))

    $handle = $Global:KERNEL32::FindFirstFileW($wildcard, $pBuffer)
    if ($handle -eq [IntPtr]::Zero) {
        [Marshal]::FreeHGlobal($pBuffer)
        return $null
    }

    do {
        $strPtr = [IntPtr]::Add($pBuffer, $cFileNameOffset)
        $filename = [Marshal]::PtrToStringUni($strPtr)
        #Write-Warning $filename

        if ($regex.IsMatch($filename)) {
            $ubr = [int]$regex.Match($filename).Groups[1].Value
            if ($maxUBR -eq $null -or $ubr -gt $maxUBR) {
                $maxUBR = $ubr
            }
        }
    } while ($Global:KERNEL32::FindNextFileW($handle, $pBuffer))

    $null = $Global:KERNEL32::FindClose($handle)
    $null = [Marshal]::FreeHGlobal($pBuffer)

    return $maxUBR
}
function Get-LatestUBR {
    param (
      [bool]$UsPs1 = $false
    )
    
    $UBR = $null
    $wildcardPattern = '*-edition*10.*.*.*'
    $regexVersion = [regex]'10\.0\.\d+\.(\d+)'
    $Manifestsfolder = 'C:\Windows\WinSxS\Manifests'
    $Packagessfolder = 'C:\Windows\servicing\Packages'
    $swTotal = [System.Diagnostics.Stopwatch]::StartNew()

    # Try Packages folder
    if (!$UsPs1) {
        $UBR = Scan-FolderWithAPI $Packagessfolder
    } else {
        $files = [Directory]::EnumerateFiles(
            $Packagessfolder, $wildcardPattern, [SearchOption]::TopDirectoryOnly)
        foreach ($file in $files) {
            #Write-Warning $file
            $match = $regexVersion.Match($file)
            if ($match.Success) {
                $candidateUBR = [int]$match.Groups[1].Value
                if ($UBR -eq $null -or $candidateUBR -gt $UBR) {
                    $UBR = $candidateUBR
                }
            }
        }
    }

    # If no result, try Manifests folder
    if ((!$UBR -or $UBR -eq 0) -and !$UsPs1) {
        $UBR = Scan-FolderWithAPI $Manifestsfolder
    }
    elseif ((!$UBR -or $UBR -eq 0) -and $UsPs1) {
        $files = [Directory]::EnumerateFiles(
            $Manifestsfolder, $wildcardPattern, [SearchOption]::TopDirectoryOnly)
        foreach ($file in $files) {
            #Write-Warning $file
            $match = $regexVersion.Match($file)
            if ($match.Success) {
                $candidateUBR = [int]$match.Groups[1].Value
                if ($candidateUBR -gt $UBR) {
                    $UBR = $candidateUBR
                }
            }
        }
    }

    # Fallback to registry if still nothing
    if (!$UBR -or $UBR -eq 0) {
        try {
            $regPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
            $UBR = Get-ItemPropertyValue -Path $regPath -Name UBR -ErrorAction Stop
        }
        catch {
            #Write-Warning "Failed to read UBR from registry: $_"
        }
    }
    
    $swTotal.Stop()
    #Write-Warning "Results: $UBR"
    #Write-Warning "Total Get-LatestUBR time: $($swTotal.ElapsedMilliseconds) ms"

    if (!$UBR) {
        return 0
    }

    return $UBR
}

<#
.SYNOPSIS
Extract DigitalProductId + DigitalProductId[4] Data,
using registry Key.

janek2012's magic decoding function
https://forums.mydigitallife.net/threads/the-ultimate-pid-checker.20816/
https://xdaforums.com/t/extract-windows-rt-product-key-without-jailbreak-or-pc.2442791/

PIDX Checker Class
https://github.com/IonBazan/pidgenx/blob/master/pidxcheckerclass.h

sppcomapi.dll
__int64 __fastcall CLicensingStateTools::get_DefaultKeyFromRegistry(CLicensingStateTools *this, unsigned __int16 **a2)
--> v6 = ReadProductKeyFromRegistry(0i64, &hMem);
--> Value = CRegUtilT<void *,CRegType,0,1>::GetValue(a1, v10, L"DigitalProductId", (BYTE **)&hMem, &v14);
--> v13 = CProductKeyUtilsT<CEmptyType>::BinaryDecode((char *)hMem + 52, v11, &v15);
a1: pointer to 16-byte product key data (from DigitalProductId4.m_abCdKey or registry).
a2: length of the data (unused much in the snippet).
a3: output pointer to store the decoded Unicode product key string.

__int64 __fastcall CProductKeyUtilsT(__m128i *a1)
{
  char Src[54];
  [__m128i] v21 = *a1;
  [__int16 *v20;] v20 = 0i64;
  v22 = *(_OWORD *)L"BCDFGHJKMPQRTVWXY2346789";
  if ( (_mm_srli_si128(v21, 8).m128i_u64[0] & 0xF0000000000000i64) != 0 )
    BREAK CODE
  [__int64] v6 = 24i64;
  [BOOL] v7 = (v21.m128i_i8[14] & 8) != 0;
  v21.m128i_i8[14] ^= (v21.m128i_i8[14] ^ (4 * ((v21.m128i_i8[14] & 8) != 0))) & 8;
  do
  {
    __int64 LODWORD(v8) = 0;
    for ( i = 14i64; i >= 0; --i )
    {
      v10 = v21.m128i_u8[i] + ((_DWORD)v8 << 8);
      v21.m128i_i8[i] = v10 / 0x18;
      v8 = v10 % 0x18;
    }
    *(_WORD *)&Src[2 * v6-- - 2] = *((_WORD *)v22 + v8);
  }
  while ( v6 >= 0 );
  
  if ( v21.m128i_i8[0] )
      BREAK CODE
  else
  {
    if ( v7 )
    {
      [__int64] v11 = 2 * v8;
      memmove_0(&v24, Src, 2 * v8);
      *(_WORD *)&Src[v11 - 2] = 78; ` Insert [N]
    }
    v12 = STRAPI_CreateCchBufferN(0x2Du, 0x1Eui64, &v20);
    if ( v12 >= 0 )
    {
      v13 = v20;
      v14 = &v24;
      for ( j = 0; j < 25; ++j )
      {
        v16 = *v14++;
        v17 = j + j / 5;
        v13[v17] = v16;
      }
      *a3 = v13;
    }
    else
       BREAK CODE
  }
}
#>
function Get-StringFromBytes {
    param(
        [byte[]]$array,
        [int]$start,
        [int]$length
    )
    if ($start + $length -le $array.Length) {
        return [Encoding]::Unicode.GetString($array, $start, $length).TrimEnd([char]0)
    }
    else {
        Write-Warning "Requested string range $start to $($start + $length) exceeds array length $($array.Length)"
        return ""
    }
}
function Get-AsciiString {
    param([byte[]]$array, [int]$start, [int]$length)
    if ($start + $length -le $array.Length) {
        return [Encoding]::ASCII.GetString($array, $start, $length).TrimEnd([char]0)
    }
    else {
        Write-Warning "Requested ASCII string range $start to $($start + $length) exceeds array length $($array.Length)"
        return ""
    }
}

<#
Source,
LicensingDiagSpp.dll, LicensingWinRT.dll, SppComApi.dll, SppWinOb.dll
__int64 __fastcall CProductKeyUtilsT<CEmptyType>::BinaryDecode(__m128i *a1, __int64 a2, unsigned __int16 **a3)

# DigitalProductId (normal key)
$pKeyBytes = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "DigitalProductId" -ErrorAction Stop | Select-Object -ExpandProperty DigitalProductId
$pKey = Get-DigitalProductKey -bCDKeyArray $pKeyBytes[52..66]
SL-InstallProductKey $pKey

# DigitalProductId4 (Windows 10/11 keys)
$pKeyBytes = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "DigitalProductId4" -ErrorAction Stop | Select-Object -ExpandProperty DigitalProductId4
$pKey = Get-DigitalProductKey -bCDKeyArray $pKeyBytes[808..822]
SL-InstallProductKey $pKey
#>
function Get-DigitalProductKey {
    param (
        [Parameter(Mandatory=$true)]
        [byte[]]$bCDKeyArray,

        [Parameter(Mandatory=$false)]
        [switch]$Log
    )

    # Clone input to v21 (like C++ __m128i copy)
    $keyData = $bCDKeyArray.Clone()

    # +2 for N` Logic Shift right [else fail]
    $Src = New-Object char[] 27

    # Character set for base-24 decoding
    $charset = "BCDFGHJKMPQRTVWXY2346789"

    # Validate input length
    if ($keyData.Length -lt 15 -or $keyData.Length -gt 16) {
        throw "Input data must be a 15 or 16 byte array."
    }

    # Win.8 key check
    if (($keyData[14] -band 0xF0) -ne 0) {
        throw "Failed to decode.!"
    }

    # N-flag
    $T = 0
    $BYTE14 = [byte]$keyData[14]
    $flag = (($BYTE14 -band 0x08) -ne 0)

    # BYTE14(v22) = (4 * (((BYTE14(v22) & 8) != 0) & 2)) | BYTE14(v22) & 0xF7;
    $keyData[14] = (4 * (([int](($BYTE14 -band 8) -ne 0)) -band 2)) -bor ($BYTE14 -band 0xF7)

    # BYTE14(v22) ^= (BYTE14(v22) ^ (4 * ((BYTE14(v22) & 8) != 0))) & 8;
    #$keyData[14] = $BYTE14 -bxor (($BYTE14 -bxor (4 * ([int](($BYTE14 -band 8) -ne 0)))) -band 8)

    # Base-24 decoding loop
    for ($idx = 24; $idx -ge 0; $idx--) {
        $last = 0
        for ($j = 14; $j -ge 0; $j--) {
            $val = $keyData[$j] + ($last -shl 8)
            $keyData[$j] = [math]::Floor($val / 0x18)
            $last = $val % 0x18
        }
        $Src[$idx] = $charset[$last]
    }

    if ($keyData[0] -ne 0) {
        throw "Invalid product key data"
    }

    # Handle N-flag
    $rev = $last -gt 13
    $pos = if ($rev) {25} else {-1}
    if ($Log) {
        $Output = (0..4 | % { -join $Src[(5*$_)..((5*$_)+4)] }) -join '-'
        Write-Warning "Before, $Output"
    }

    # Shift Left, Insert N, At position 0 >> $Src[0]=`N`
    if ($flag -and ($last -le 0)) {
        $Src[0] = [Char]78
    }
    # Shift right, Insert N, Count 1-25 [27 Base,0-24 & 2` Spacer's]
    elseif ($flag -and $rev) {
        while ($pos-- -gt $last){$Src[$pos + 1]=$Src[$pos]}
        $T, $Src[$last+1] = 1, [char]78
    }
    # Shift left, Insert N,
    elseif ($flag -and !$rev) {
        while (++$pos -lt $last){$Src[$pos] = $Src[$pos + 1]}
        $Src[$last] = [char]78
    }

    # Dynamically format 5x5 with dashes
    $Output = (0..4 | % { -join $Src[((5*$_)+$T)..((5*$_)+4+$T)] }) -join '-'
    if ($Log) {
        Write-Warning "After,  $Output"
    }
    return $Output
}
function Parse-DigitalProductId {
    param (
        [string]$RegistryPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    )

    try {
        $digitalProductId = (Get-ItemProperty -Path $RegistryPath -ErrorAction Stop).DigitalProductId
    }
    catch {
        Write-Warning "Failed to read DigitalProductId from registry path $RegistryPath"
        return $null
    }

    if (-not $digitalProductId) {
        Write-Warning "DigitalProductId property not found in registry."
        return $null
    }

    # Ensure byte array
    $byteArray = if ($digitalProductId -is [byte[]]) { $digitalProductId } else { [byte[]]$digitalProductId }

    # Define offsets and lengths for each field in one hashtable
    $offsets = @{
        uiSize        = @{ Offset = 0;  Length = 4  }
        MajorVersion  = @{ Offset = 4;  Length = 2  }
        MinorVersion  = @{ Offset = 6;  Length = 2  }
        ProductId     = @{ Offset = 8;  Length = 24 }
        EditionId     = @{ Offset = 36; Length = 16 }
        bCDKey        = @{ Offset = 52; Length = 16 }
    }

    # Extract components safely
    $uiSize = [BitConverter]::ToUInt32($byteArray, $offsets.UISize.Offset)
    $productId = Get-AsciiString -array $byteArray -start $offsets.ProductId.Offset -length $offsets.ProductId.Length
    $editionId = Get-AsciiString -array $byteArray -start $offsets.EditionId.Offset -length $offsets.EditionId.Length

    # Extract bCDKey array for product key decoding
    $bCDKeyArray = $byteArray[$offsets.bCDKey.Offset..($offsets.bCDKey.Offset + $offsets.bCDKey.Length - 1)]

    # Decode Digital Product Key (placeholder function - implement accordingly)
    $digitalProductKey = Get-DigitalProductKey -bCDKeyArray $bCDKeyArray

    # Extract MajorVersion and MinorVersion from byte array
    $majorVersion = [BitConverter]::ToUInt16($byteArray, $offsets.MajorVersion.Offset)
    $minorVersion = [BitConverter]::ToUInt16($byteArray, $offsets.MinorVersion.Offset)

    # Return structured object
    return [PSCustomObject]@{
        UISize       = $uiSize
        MajorVersion = $majorVersion
        MinorVersion = $minorVersion
        ProductId    = $productId
        EditionId    = $editionId
        DigitalKey   = $digitalProductKey
    }
}
function Parse-DigitalProductId4 {
    param(
        [string]$RegistryPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion",
        [IntPtr]$Pointer = [System.IntPtr]::Zero,
        [int]$Length = 0,

        [switch] $FromIntPtr,
        [switch] $FromRegistry
    )

    $ParamsCheck = -not ($FromIntPtr -xor $FromRegistry)
    if ($ParamsCheck) {
        $FromIntPtr = $null
        $FromRegistry = $true
        Write-Warning "use default values, read from registry"
    }

    # Retrieve DigitalProductId4
    if ($FromIntPtr) {
        if ($Pointer -ne [System.IntPtr]::Zero -and $Length -gt 0) {
            try {
                $byteArray = New-Object byte[] $Length
                [Marshal]::Copy($Pointer, $byteArray, 0, $Length)
            }
            catch {
                Write-Warning "Failed to copy memory from pointer."
                return $null
            }
        }
    }
    if ($FromRegistry) {
        try {
            $digitalProductId4 = (Get-ItemProperty -Path $RegistryPath -ErrorAction Stop).DigitalProductId4
        }
        catch {
            Write-Warning "Failed to read DigitalProductId4 from registry path $RegistryPath"
            return $null
        }

        if (-not $digitalProductId4) {
            Write-Warning "DigitalProductId4 property not found in registry."
            return $null
        }

        # Ensure we have a byte array
        $byteArray = if ($digitalProductId4 -is [byte[]]) { $digitalProductId4 } else { [byte[]]$digitalProductId4 }
    }

    # Offsets dictionary for structured fields with length included
    $offsets = @{
        uiSize        = @{ Offset = 0;    Length = 4  }
        MajorVersion  = @{ Offset = 4;    Length = 2  }
        MinorVersion  = @{ Offset = 6;    Length = 2  }
        AdvancedPid  = @{ Offset = 8;    Length = 128 }
        ActivationId = @{ Offset = 136;  Length = 128 }
        EditionType  = @{ Offset = 280;  Length = 520 }
        EditionId    = @{ Offset = 888;  Length = 128 }
        KeyType      = @{ Offset = 1016; Length = 128 }
        EULA         = @{ Offset = 1144; Length = 128 }
        bCDKey       = @{ Offset = 808;  Length = 16  }
    }

    # Extract values
    $uiSize = if ($byteArray.Length -ge 4) { [BitConverter]::ToUInt32($byteArray, 0) } else { 0 }

    $advancedPid = Get-StringFromBytes -array $byteArray -start $offsets.AdvancedPid.Offset -length $offsets.AdvancedPid.Length
    $activationId = Get-StringFromBytes -array $byteArray -start $offsets.ActivationId.Offset -length $offsets.ActivationId.Length
    $editionType = Get-StringFromBytes -array $byteArray -start $offsets.EditionType.Offset -length $offsets.EditionType.Length
    $editionId = Get-StringFromBytes -array $byteArray -start $offsets.EditionId.Offset -length $offsets.EditionId.Length
    $keyType = Get-StringFromBytes -array $byteArray -start $offsets.KeyType.Offset -length $offsets.KeyType.Length
    $eula = Get-StringFromBytes -array $byteArray -start $offsets.EULA.Offset -length $offsets.EULA.Length

    # Extract bCDKey array used for key retrieval
    $bCDKeyOffset = $offsets.bCDKey.Offset
    $bCDKeyLength = $offsets.bCDKey.Length
    $bCDKeyArray = $byteArray[$bCDKeyOffset..($bCDKeyOffset + $bCDKeyLength - 1)]

    # Extract MajorVersion and MinorVersion from byte array
    $majorVersion = [BitConverter]::ToUInt16($byteArray, $offsets.MajorVersion.Offset)
    $minorVersion = [BitConverter]::ToUInt16($byteArray, $offsets.MinorVersion.Offset)

    # Call to external helper to decode the Digital Product Key
    # You need to define this function based on your key decoding logic
    $digitalProductKey = Get-DigitalProductKey -bCDKeyArray $bCDKeyArray

    # Return a structured object
    return [PSCustomObject]@{
        UISize       = $uiSize
        MajorVersion = $majorVersion
        MinorVersion = $minorVersion
        AdvancedPID  = $advancedPid
        ActivationID = $activationId
        EditionType  = $editionType
        EditionID    = $editionId
        KeyType      = $keyType
        EULA         = $eula
        DigitalKey   = $digitalProductKey
    }
}

<#
Adjusting Token Privileges in PowerShell
https://www.leeholmes.com/adjusting-token-privileges-in-powershell/

typedef struct _TOKEN_PRIVILEGES {
  DWORD               PrivilegeCount;
  LUID_AND_ATTRIBUTES Privileges[ANYSIZE_ARRAY];
} TOKEN_PRIVILEGES, *PTOKEN_PRIVILEGES;

typedef struct _LUID_AND_ATTRIBUTES {
  LUID  Luid;
  DWORD Attributes;
} LUID_AND_ATTRIBUTES, *PLUID_AND_ATTRIBUTES;

typedef struct _LUID {
  DWORD LowPart;
  LONG  HighPart;
} LUID, *PLUID;

--------------------

Clear-Host
Write-Host

Write-Host
$length = [Uint32]1
$Ptr    = [IntPtr]::Zero
$lastErr = Invoke-UnmanagedMethod `
    -Dll "ntdll.dll" `
    -Function "NtEnumerateBootEntries" `
    -Return "Int64" `
    -Params "IntPtr Ptr, ref uint length" `
    -Values @($Ptr, [ref]$length)
Parse-ErrorMessage `
    -MessageId $lastErr

Write-Host
# Get Minimal Privileges To Load Some NtDll function
Adjust-TokenPrivileges `
    -Privilege @("SeDebugPrivilege","SeImpersonatePrivilege","SeIncreaseQuotaPrivilege","SeAssignPrimaryTokenPrivilege", "SeSystemEnvironmentPrivilege") `
    -Log -SysCall

Write-Host
$length = [Uint32]1
$Ptr    = [IntPtr]::Zero
$lastErr = Invoke-UnmanagedMethod `
    -Dll "ntdll.dll" `
    -Function "NtEnumerateBootEntries" `
    -Return "Int64" `
    -Params "IntPtr Ptr, ref uint length" `
    -Values @($Ptr, [ref]$length)
Parse-ErrorMessage `
    -MessageId $lastErr
#>
Function Adjust-TokenPrivileges {
    param(
        [Parameter(Mandatory=$false)]
        [Process]$Process,

        [Parameter(Mandatory=$false)]
        [IntPtr]$hProcess,

        [Parameter(Mandatory=$false)]
        [ValidateSet(
        "SeAssignPrimaryTokenPrivilege", "SeAuditPrivilege", "SeBackupPrivilege",
        "SeChangeNotifyPrivilege", "SeCreateGlobalPrivilege", "SeCreatePagefilePrivilege",
        "SeCreatePermanentPrivilege", "SeCreateSymbolicLinkPrivilege", "SeCreateTokenPrivilege",
        "SeDebugPrivilege", "SeEnableDelegationPrivilege", "SeImpersonatePrivilege",
        "SeIncreaseQuotaPrivilege", "SeIncreaseWorkingSetPrivilege", "SeLoadDriverPrivilege",
        "SeLockMemoryPrivilege", "SeMachineAccountPrivilege", "SeManageVolumePrivilege", 
        "SeProfileSingleProcessPrivilege", "SeRelabelPrivilege", "SeRemoteShutdownPrivilege",
        "SeRestorePrivilege", "SeSecurityPrivilege", "SeShutdownPrivilege", "SeSyncAgentPrivilege",
        "SeSystemEnvironmentPrivilege", "SeSystemProfilePrivilege", "SeSystemtimePrivilege",
        "SeTakeOwnershipPrivilege", "SeTcbPrivilege", "SeTimeZonePrivilege", "SeTrustedCredManAccessPrivilege",
        "SeUndockPrivilege", "SeDelegateSessionUserImpersonatePrivilege", "SeIncreaseBasePriorityPrivilege",
        "SeNetworkLogonRight", "SeInteractiveLogonRight", "SeRemoteInteractiveLogonRight", "SeDenyNetworkLogonRight",
        "SeDenyBatchLogonRight", "SeDenyServiceLogonRight", "SeDenyInteractiveLogonRight", "SeDenyRemoteInteractiveLogonRight",
        "SeBatchLogonRight", "SeServiceLogonRight"
        )]
        [string[]]$Privilege,

        [Parameter(Mandatory=$false)]
        [Switch] $AdjustAll,

        [Parameter(Mandatory=$false)]
        [switch] $Query,

        [Parameter(Mandatory=$false)]
        [Switch] $Disable,

        [Parameter(Mandatory=$false)]
        [Switch] $SysCall,

        [Parameter(Mandatory=$false)]
        [Switch] $Log
    )

    function Get-PrivilegeLuid {
        param (
            [ValidateNotNullOrEmpty()]
            [string]$PrivilegeName
        )

        $policyHandle = [IntPtr]::Zero
        $objAttr = New-IntPtr -Size 60 -WriteSizeAtZero

        $status = $Global:advapi32::LsaOpenPolicy(
            [IntPtr]::Zero, $objAttr, 0x800, [ref]$policyHandle)
    
        if ($status -ne 0) {
            $policyHandle = $null
            return $null
        }
    
        try {
            $luid = [Int64]0
            $privName = Init-NativeString -Value $PrivilegeName -Encoding Unicode
            $status = $Global:advapi32::LsaLookupPrivilegeValue(
                $policyHandle, $privName, [ref]$luid)
            $Global:advapi32::LsaClose($policyHandle) | Out-Null

            if ($status -ne 0) {
                return $null
            }
        }
        Finally {
            Free-NativeString -StringPtr $privName | Out-Null
            $privName = $null
            $policyHandle = $null
        }
        return $luid
    }

    $TOKEN_QUERY = 0x00000008;
    $TOKEN_ADJUST_PRIVILEGES = 0x00000020;
    $SE_PRIVILEGE_ENABLED = 0x00000002;
    $SE_PRIVILEGE_DISABLED = 0x00000000;
    
    if ($Process -and $hProcess) {
        throw "-Process or -hProcess Only."
    }

    if (-not $Process -and -not $hProcess) {
        $Process = [Process]::GetCurrentProcess()
    }

    if ((!$Privilege -or $Privilege.Count -eq 0) -and (!$AdjustAll) -and (!$Query)) {
        throw "use -Privilege or -AdjustAll -or -Query"
    }

    $count = [bool]($Privilege -and $Privilege.Count -gt 0) + [bool]$AdjustAll + [bool]$Query
    if ($count -gt 1) {
        throw "use -Privilege or -AdjustAll -or -Query"
    }

    if ($Privilege ) {
        if ($Privilege.Count -gt 0 -and $AdjustAll) {
            throw "use -Privilege or -AdjustAll"
        }
    }

    # Validate the handle is valid and non-zero
    $hproc = if ($Process) {$Process.Handle} else {$hProcess}
    if ($hproc -eq [IntPtr]::Zero -or $hproc -eq 0 -or $hproc -eq $Null) {
        throw "Invalid process handle."
    }
    
    $hToken = [IntPtr]::Zero
    $hproc = [IntPtr]$hproc

    if ($SysCall) {
        $retVal = $Global:ntdll::NtOpenProcessToken(
            $hproc, ($TOKEN_ADJUST_PRIVILEGES -bor $TOKEN_QUERY), [ref]$hToken)
    }
    else {
        $retVal = $Global:advapi32::OpenProcessToken(
            $hproc, ($TOKEN_ADJUST_PRIVILEGES -bor $TOKEN_QUERY), [ref]$hToken)
    }

    # if both return same result, which can be true if both *false
    # well, in that case -> throw, and return error
    if ((!$SysCall -and $retVal -ne 0 -and $hToken -ne [IntPtr]::Zero) -eq (
        $SysCall -and $retVal -eq 0 -and $hToken -ne [IntPtr]::Zero)) {
            throw "OpenProcessToken failed with -> $retVal"
    }

    if ($Query) {
        # Allocate memory for TOKEN_PRIVILEGES
        $tokenInfoPtr = [Marshal]::AllocHGlobal($tokenInfoLength)
        try {
            $tokenInfoLength = 0
            $Global:advapi32::GetTokenInformation($hToken, 3, [IntPtr]0, 0, [ref]$tokenInfoLength) | Out-Null
            if ($tokenInfoLength -le 0) {
                throw "GetTokenInformation failed .!"
            }
            $tokenInfoPtr = New-IntPtr -Size $tokenInfoLength
            if (0 -eq (
                $Global:advapi32::GetTokenInformation($hToken, 3, $tokenInfoPtr, $tokenInfoLength, [ref]$tokenInfoLength))) { 
                    throw "GetTokenInformation failed on second call" }

            $privileges = @()
            $Count = [Marshal]::ReadInt32($tokenInfoPtr)

            for ($i=0; $i -lt $Count; $i++) {
                $offset = 4 + ($i * 12)
                $luid = [Marshal]::ReadInt64($tokenInfoPtr, $offset)
                $attr = [Marshal]::ReadInt32($tokenInfoPtr, $offset+8)
                $enabled = ($attr -band 2) -ne 0

                $size = 0
                $namePtr = [IntPtr]::Zero
                $Global:advapi32::LookupPrivilegeNameW([IntPtr]::Zero, [ref]$luid, $namePtr, [ref]$size) | Out-Null
                $namePtr = [Marshal]::AllocHGlobal(($size + 1) * 2)
                try {
                    $Global:advapi32::LookupPrivilegeNameW([IntPtr]::Zero, [ref]$luid, $namePtr, [ref]$size) | Out-Null
                    $privName = [Marshal]::PtrToStringUni($namePtr)
                    $privileges += [PSCustomObject]@{
                        Name    = $privName
                        LUID    = $luid
                        Enabled = $enabled
                    }
                }
                finally {
                    [Marshal]::FreeHGlobal($namePtr)
                }
            }

            return $privileges
        }
        finally {
            Free-IntPtr -handle $tokenInfoPtr
            Free-IntPtr -handle $hProc -Method NtHandle
            Free-IntPtr -handle $hToken -Method NtHandle
        }
    }

    if ($AdjustAll) {
        $Privilege = (
            "SeAssignPrimaryTokenPrivilege", "SeAuditPrivilege", "SeBackupPrivilege",
            "SeChangeNotifyPrivilege", "SeCreateGlobalPrivilege", "SeCreatePagefilePrivilege",
            "SeCreatePermanentPrivilege", "SeCreateSymbolicLinkPrivilege", "SeCreateTokenPrivilege",
            "SeDebugPrivilege", "SeEnableDelegationPrivilege", "SeImpersonatePrivilege",
            "SeIncreaseQuotaPrivilege", "SeIncreaseWorkingSetPrivilege", "SeLoadDriverPrivilege",
            "SeLockMemoryPrivilege", "SeMachineAccountPrivilege", "SeManageVolumePrivilege", 
            "SeProfileSingleProcessPrivilege", "SeRelabelPrivilege", "SeRemoteShutdownPrivilege",
            "SeRestorePrivilege", "SeSecurityPrivilege", "SeShutdownPrivilege", "SeSyncAgentPrivilege",
            "SeSystemEnvironmentPrivilege", "SeSystemProfilePrivilege", "SeSystemtimePrivilege",
            "SeTakeOwnershipPrivilege", "SeTcbPrivilege", "SeTimeZonePrivilege", "SeTrustedCredManAccessPrivilege",
            "SeUndockPrivilege", "SeDelegateSessionUserImpersonatePrivilege", "SeIncreaseBasePriorityPrivilege",
            "SeNetworkLogonRight", "SeInteractiveLogonRight", "SeRemoteInteractiveLogonRight", "SeDenyNetworkLogonRight",
            "SeDenyBatchLogonRight", "SeDenyServiceLogonRight", "SeDenyInteractiveLogonRight", "SeDenyRemoteInteractiveLogonRight",
            "SeBatchLogonRight", "SeServiceLogonRight"
        )
    }

    # Bug fix ~~~~ !
    # Update case of 1 fail, and function break.
    
    # Prepare
    $validEntries = @()
    foreach ($priv in $Privilege) {
        try {
            [Int64]$luid = 0
            if ($SysCall) {
                $luid = Get-PrivilegeLuid -PrivilegeName $priv
                if ($luid -le 0) { throw "Get-PrivilegeLuid failed for '$priv'" }
            } else {
                $result = $Global:advapi32::LookupPrivilegeValue([IntPtr]::Zero, $priv, [ref]$luid)
                if ($result -eq 0) { throw "LookupPrivilegeValue failed for '$priv'" }
            }

            if ($luid -ne 0) {
                $validEntries += [PSCustomObject]@{
                    Name = $priv
                    LUID = $luid
                }
            }
        }
        catch {
            Write-Warning $_.Exception.Message
        }
    }

    if ($validEntries.Count -eq 0) {
        Write-Warning "No valid privileges could be resolved."
        return $false
    }

    # Allocate proper size
    $Count = $validEntries.Count
    $TokPriv1LuidSize = 4 + (12 * $Count)
    $TokPriv1LuidPtr = New-IntPtr -Size $TokPriv1LuidSize -InitialValue $Count

    # Write privileges into the structure
    for ($i = 0; $i -lt $Count; $i++) {
        $offset = 4 + (12 * $i)
        [Marshal]::WriteInt64($TokPriv1LuidPtr, $offset, $validEntries[$i].LUID)

        $attrValue = if ($Disable) { $SE_PRIVILEGE_DISABLED } else { $SE_PRIVILEGE_ENABLED }
        [Marshal]::WriteInt32($TokPriv1LuidPtr, $offset + 8, $attrValue)

        if ($Log) {
            Write-Host ">>> Privilege: $($validEntries[$i].Name)"
            Write-Host ("Offset $offset (LUID) : 0x{0:X}" -f $validEntries[$i].LUID)
            Write-Host ("Offset $($offset+8) (Attr): 0x{0:X} {1}" -f $attrValue,
                $(if ($attrValue -eq 2) { 'SE_PRIVILEGE_ENABLED' }
                  elseif ($attrValue -eq 0) { 'SE_PRIVILEGE_DISABLED' }
                  else { 'UNKNOWN' }))
        }
    }
    try {
        if ($FailToWriteBlock) {
            Write-Warning "Failed to build privilege block. Skipping AdjustTokenPrivileges."
        } 
        else {
            if ($SysCall) {
                $retVal = $Global:ntdll::NtAdjustPrivilegesToken(
                    $hToken, $false, $TokPriv1LuidPtr, $TokPriv1LuidSize, [IntPtr]::Zero, [IntPtr]::Zero)
                if ($retVal -eq 0) {
                    return $true
                } elseif ($retVal -eq 262) {
                    Write-Warning "AdjustTokenPrivileges succeeded but not all privileges assigned."
                    return $true
                } else {
                    $status = Parse-ErrorMessage -MessageId $retVal -Flags NTSTATUS
                    Write-Warning "NtAdjustPrivilegesToken failed: $status"
                    return $false
                }
            } else {
                $retVal = $Global:advapi32::AdjustTokenPrivileges(
                    $hToken, $false, $TokPriv1LuidPtr, $TokPriv1LuidSize, [IntPtr]::Zero, [IntPtr]::Zero)
                $lastErr = [Marshal]::GetLastWin32Error()
                if ($retVal -eq 0) {
                    $status = Parse-ErrorMessage -MessageId $lastErr -Flags WIN32
                    Write-Warning "AdjustTokenPrivileges failed: $status"
                    returh $false
                } elseif ($lastErr -eq 1300) {
                    Write-Warning "AdjustTokenPrivileges succeeded but not all privileges assigned."
                    return $true
                } else {
                    return $true
                }
            }
        }
    }
    Finally {
        Free-IntPtr -handle $TokPriv1LuidPtr
        Free-IntPtr -handle $hProc -Method NtHandle
        Free-IntPtr -handle $hToken -Method NtHandle
    }
}

<#
SID structure (winnt.h)
https://learn.microsoft.com/en-us/windows/win32/api/winnt/ns-winnt-sid

typedef struct _SID {
  BYTE                     Revision;
  BYTE                     SubAuthorityCount;
  SID_IDENTIFIER_AUTHORITY IdentifierAuthority;
#if ...
  DWORD                    *SubAuthority[];
#else
  DWORD                    SubAuthority[ANYSIZE_ARRAY];
#endif
} SID, *PISID;

~~~~~~~~~~~~~~~~~~~

Well-known SIDs
https://learn.microsoft.com/en-us/windows/win32/secauthz/well-known-sids

The SECURITY_NT_AUTHORITY (S-1-5) predefined identifier authority produces SIDs that are not universal but are meaningful only on Windows installations.
You can use the following RID values with SECURITY_NT_AUTHORITY to create well-known SIDs.

SECURITY_LOCAL_SYSTEM_RID
String value: S-1-5-18
A special account used by the operating system.

The following table has examples of domain-relative RIDs that you can use to form well-known SIDs for **local groups** (aliases).
For more information about local and global groups, see Local Group Functions and Group Functions.

DOMAIN_ALIAS_RID_ADMINS
Value: 0x00000220
String value: S-1-5-32-544
A local group used for administration of the domain.

~~~~~~~~~~~~~~~~~~~

TOKEN_INFORMATION_CLASS enumeration (winnt.h)
https://learn.microsoft.com/en-us/windows/win32/api/winnt/ne-winnt-token_information_class

typedef enum _TOKEN_INFORMATION_CLASS {
  TokenUser = 1,
  TokenGroups,
  TokenPrivileges,
  TokenOwner,
  TokenPrimaryGroup,
  TokenDefaultDacl,
  TokenSource,
  TokenType,
  TokenImpersonationLevel,
  TokenStatistics,
  TokenRestrictedSids,
  TokenSessionId,
  TokenGroupsAndPrivileges,
  TokenSessionReference,
  ...
  ...
  ...

~~~~~~~~~~~~~~~~~~~

$isSystem = Check-AccountType -AccType System
$isAdmin  = Check-AccountType -AccType Administrator
Write-Host "is Admin* Acc ? $isAdmin"
Write-Host "is System Acc ? $isSystem"
Write-Host
#>
function Check-AccountType {
    param (
       [Parameter(Mandatory)]
       [ValidateSet("System","Administrator")]
       [string]$AccType
    )

$isMember = $false

if (!([PSTypeName]'TOKEN').Type) {
Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Security.Principal;

public class TOKEN {
 
    [DllImport("kernelbase.dll")]
    public static extern IntPtr GetCurrentProcessId();

    [DllImport("ntdll.dll")]
    public static extern void RtlZeroMemory(
        IntPtr Destination,
        UIntPtr Length);

    [DllImport("ntdll.dll")]
    public static extern Int32 NtClose(
        IntPtr hObject);


    [DllImport("ntdll.dll")]
    public static extern Int32 RtlCheckTokenMembershipEx(
        IntPtr TokenHandle,
        IntPtr Sid,
        Int32 Flags,
        ref Boolean IsMember);

    [DllImport("ntdll.dll")]
    public static extern Int32 NtOpenProcess(
        ref IntPtr ProcessHandle,
        UInt32 DesiredAccess,
        IntPtr ObjectAttributes,
        IntPtr ClientId);

    [DllImport("ntdll.dll")]
    public static extern Int32 NtOpenProcessToken(
        IntPtr ProcessHandle,
        uint DesiredAccess,
        out IntPtr TokenHandle);

    [DllImport("ntdll.dll")]
    public static extern Int32 NtQueryInformationToken(
        IntPtr TokenHandle,
        int TokenInformationClass,
        IntPtr TokenInformation,
        UInt32 TokenInformationLength,
        out uint ReturnLength );
}
'@
  }

function Check {
    param (
        [Parameter(Mandatory)]
        [IntPtr]$pSid,

        [Parameter(Mandatory)]
        [int[]]$Subs,

        [Parameter(Mandatory)]
        [ValidateSet("Account", "Group")]
        [string]$Type
    )

    if ($null -eq $Subs -or $Subs.Length -le 0 -or $pSid -eq [IntPtr]::Zero) {
        throw "Invalid parameters: pSid or Subs is empty."
    }
    [Marshal]::WriteByte($pSid, 1, $Subs.Length)
    for ($i = 0; $i -lt $Subs.Length; $i++) {
        [Marshal]::WriteInt32($pSid, 8 + 4 * $i, $Subs[$i])
    }

    switch ($Type) {
        "Group" {
            $ret = [TOKEN]::RtlCheckTokenMembershipEx(
                0, $pSid, 0, [ref]$isMember)
        }
        "Account" {
            if ([IntPtr]::Size -eq 8) {
                # 64-bit sizes and layout
                $clientIdSize = 16
                $objectAttrSize = 48
            } else {
                # 32-bit sizes and layout (WOW64)
                $clientIdSize = 8
                $objectAttrSize = 24
            }
            $hproc  = [IntPtr]::Zero
            $procID = [TOKEN]::GetCurrentProcessId()
            $clientIdPtr   = [marshal]::AllocHGlobal($clientIdSize)
            $attributesPtr = [marshal]::AllocHGlobal($objectAttrSize)
            [TOKEN]::RtlZeroMemory($clientIdPtr, [Uintptr]::new($clientIdSize))
            [TOKEN]::RtlZeroMemory($attributesPtr, [Uintptr]::new($objectAttrSize))
            [marshal]::WriteInt32($attributesPtr, 0x0, $objectAttrSize)
            if ([IntPtr]::Size -eq 8) {
              [Marshal]::WriteInt64($clientIdPtr, 0, [Int64]$procID)
            }
            else {
              [Marshal]::WriteInt32($clientIdPtr, 0x0, $procID)
            }
            try {
                if (0 -ne [TOKEN]::NtOpenProcess(
                    [ref]$hproc, 0x0400, $attributesPtr, $clientIdPtr)) {
                        throw "NtOpenProcess fail."
                }
            }
            finally {
                @($clientIdPtr, $attributesPtr) | % {[Marshal]::FreeHGlobal($_)}
            }

            $hToken = [IntPtr]::Zero
            if (0 -ne [TOKEN]::NtOpenProcessToken(
                $hproc, 0x00000008, [ref]$hToken)) {
                    throw "NtOpenProcessToken fail."
            }
            try {
                [UInt32]$ReturnLength = 0
                $TokenInformation = [marshal]::AllocHGlobal(100)
                if (0 -ne [TOKEN]::NtQueryInformationToken(
                    $hToken,1,$TokenInformation, 100, [ref]$ReturnLength)) {
                        throw "NtQueryInformationToken fail."
                }

                $pUserSid = [Marshal]::ReadIntPtr($TokenInformation)
                $isMember = ($Subs.Length -eq [Marshal]::ReadByte($pUserSid,1)) -and
                    ([Marshal]::ReadByte($pUserSid,0) -eq [Marshal]::ReadByte($pSid,0)) -and
                    ([Marshal]::ReadByte($pUserSid,7) -eq [Marshal]::ReadByte($pSid,7))
                if ($isMember) {
                    for ($i=0; $i -lt $Subs.Length; $i++) {
                        if ([Marshal]::ReadInt32($pUserSid, 8 + 4*$i) -ne $Subs[$i]) {
                            $isMember = $false
                            break
                }}}
            }
            finally {            
                [marshal]::FreeHGlobal($TokenInformation)
                @($hproc, $hToken) | % { [TOKEN]::NtClose($_) | Out-Null }
            }
        }
    }

    return $isMember
}
  
  #SECURITY_NT_AUTHORITY (S-1-5)
  $isMember = $false
  $Rev, $Auth, $Count, $MaxCount = 1,5,0,10
  $pSid = [Marshal]::AllocHGlobal(8+(4*$MaxCount))
  @($Rev, $Count, 0,0,0,0,0, $Auth) | ForEach -Begin { $i = 0 } -Process { [Marshal]::WriteByte($pSid, $i++, $_)  }
  try {
    switch ($AccType) {
        "System" {
            # S-1-5-[18] // @([1],Count,0,0,0,0,0,[5] && 18)
            $isMember = Check -pSid $pSid -Subs @(18) -Type Account
        }
        "Administrator" {
            # S-1-5-[32]-[544] // @([1],Count,0,0,0,0,0,[5] && 32,544)
            $isMember = Check -pSid $pSid -Subs @(32, 544) -Type Group
        }
    }
  }
  catch {
    Write-warning "An error occurred: $_"
    if (-not [Environment]::Is64BitProcess -and [Environment]::Is64BitOperatingSystem) {
        Write-warning "This script could fail on x86 PowerShell in a 64-bit system."
    }
    $isMember = $null
  }
  [Marshal]::FreeHGlobal($pSid)
  return $isMember
}

<#
* Thread Environment Block (TEB)
* https://www.geoffchappell.com/studies/windows/km/ntoskrnl/inc/api/pebteb/teb/index.htm

* Process Environment Block (PEB)
* https://www.geoffchappell.com/studies/windows/km/ntoskrnl/inc/api/pebteb/peb/index.htm

[TEB]
--> NT_TIB NtTib; 0x00
---->
    Struct {
    ...
    PNT_TIB Self; <<<<< gs:[0x30] / fs:[0x18]
    } NT_TIB
#>
if (!([PSTypeName]'TEB').Type) {
$TEB = @"
using System;
using System.Runtime.InteropServices;

public static class TEB
{
    public delegate IntPtr GetAddress();
    public delegate void GetAddressByPointer(IntPtr Ret);
    public delegate void GetAddressByReference(ref IntPtr Ret);

    public static IntPtr CallbackResult;

    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    public delegate void CallbackDelegate(IntPtr callback, IntPtr TEB);

    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    public delegate void RemoteThreadDelgate(IntPtr callback);

    public static CallbackDelegate GetCallback()
    {
        return new CallbackDelegate((IntPtr del, IntPtr val) =>
        {
            CallbackResult = val;
        });
    }

    // Example in C#
    public static bool IsRobustValidx64Stub(IntPtr funcAddress)
    {
        byte[] buffer = new byte[30];
        System.Runtime.InteropServices.Marshal.Copy(funcAddress, buffer, 0, 30);

        // Look for the "mov r10, rcx" pattern
        int movR10RcxIndex = -1;
        for (int i = 0; i < buffer.Length - 2; i++) {
            if (buffer[i] == 0x4C && buffer[i+1] == 0x8B && buffer[i+2] == 0xD1) {
                movR10RcxIndex = i;
                break;
            }
        }
        if (movR10RcxIndex == -1) return false;

        // Look for the "mov eax, [syscall_id]" pattern
        int movEaxIndex = -1;
        for (int i = movR10RcxIndex; i < buffer.Length - 1; i++) {
            if (buffer[i] == 0xB8) {
                movEaxIndex = i;
                break;
            }
        }
        if (movEaxIndex == -1) return false;

        // Look for the "syscall" pattern
        int syscallIndex = -1;
        for (int i = movEaxIndex; i < buffer.Length - 1; i++) {
            if (buffer[i] == 0x0F && buffer[i+1] == 0x05) {
                syscallIndex = i;
                break;
            }
        }
        if (syscallIndex == -1) return false;

        // Look for the "ret" pattern
        for (int i = syscallIndex; i < buffer.Length; i++) {
            if (buffer[i] == 0xC3) {
                return true;
            }
        }

        return false;
    }

    public static byte[] GenerateSyscallx64 (byte[] syscall)
    {
        return  new byte[]
        {
            0x4C, 0x8B, 0xD1,                                       // mov r10, rcx
            0xB8, syscall[0], syscall[1], syscall[2], syscall[3],   // mov eax, syscall
            0x0F, 0x05,                                             // syscall
            0xC3                                                    // ret
        };
    }

    public static byte[] GenerateSyscallx86(IntPtr stubAddress)
    {
        int maxStubSize = 20;
        byte[] stubBytes = new byte[maxStubSize];
        Marshal.Copy(stubAddress, stubBytes, 0, maxStubSize);

        // Validate the start: mov eax, [syscall_id] (opcode B8)
        if (stubBytes[0] != 0xB8)
        {
            throw new Exception("Invalid x86 syscall stub: 'mov eax' instruction not found.");
        }

        // Find the 'mov edx, [Wow64SystemServiceCall]' instruction (opcode BA)
        int movEdxIndex = -1;
        for (int i = 5; i < maxStubSize; i++)
        {
            if (stubBytes[i] == 0xBA)
            {
                movEdxIndex = i;
                break;
            }
        }
        if (movEdxIndex == -1)
        {
            throw new Exception("Invalid x86 syscall stub: 'mov edx' not found.");
        }

        // Find the 'call edx' instruction (opcode FF D2)
        int callEdxIndex = -1;
        for (int i = movEdxIndex + 5; i < maxStubSize - 1; i++)
        {
            if (stubBytes[i] == 0xFF && stubBytes[i + 1] == 0xD2)
            {
                callEdxIndex = i;
                break;
            }
        }
        if (callEdxIndex == -1)
        {
            throw new Exception("Invalid x86 syscall stub: 'call edx' not found.");
        }

        // Find the end of the stub: 'retn [size]' (C2) or 'ret' (C3)
        int stubLength = -1;
        for (int i = callEdxIndex + 2; i < maxStubSize; i++)
        {
            if (stubBytes[i] == 0xC2) // retn with parameters
            {
                stubLength = i + 3;
                break;
            }
            else if (stubBytes[i] == 0xC3) // ret with no parameters
            {
                stubLength = i + 1;
                break;
            }
        }
        if (stubLength == -1)
        {
            throw new Exception("Could not find the 'ret' or 'retn' instruction.");
        }

        byte[] syscallShellcode = new byte[stubLength];
        Array.Copy(stubBytes, syscallShellcode, stubLength);
        return syscallShellcode;
    }

    [DllImport("kernel32.dll", CharSet=CharSet.Unicode)]
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    public static extern IntPtr LoadLibraryW(string lpLibFileName);
 
    [DllImport("kernel32.dll")]
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    public static extern IntPtr GetProcAddress(
        IntPtr hModule,
        string lpProcName);

    [DllImport("kernel32.dll")]
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    public static extern IntPtr GetProcessHeap();

    [DllImport("ntdll.dll", CallingConvention = CallingConvention.StdCall)]
    public static extern int ZwAllocateVirtualMemory(
        IntPtr ProcessHandle,
        ref IntPtr BaseAddress,
        UIntPtr ZeroBits,
        ref UIntPtr RegionSize,
        uint AllocationType,
        uint Protect
    );

    [DllImport("ntdll.dll", CallingConvention = CallingConvention.StdCall)]
    public static extern int ZwAllocateVirtualMemoryEx(
        IntPtr ProcessHandle,
        ref IntPtr BaseAddress,
        ref UIntPtr RegionSize,
        uint AllocationType,
        uint Protect,
        IntPtr ExtendedParameters,
        uint ParameterCount
    );

    [DllImport("ntdll.dll", CallingConvention = CallingConvention.StdCall)]
    public static extern int ZwFreeVirtualMemory(
        IntPtr ProcessHandle,
        ref IntPtr BaseAddress,
        ref UIntPtr RegionSize,
        uint FreeType
    );
 
    [DllImport("ntdll.dll", SetLastError = true)]
    public static extern int NtProtectVirtualMemory(
        IntPtr ProcessHandle,           // Handle to the process
        ref IntPtr BaseAddress,         // Base address of the memory region -> ByRef
        ref UIntPtr RegionSize,         // Size of the region to protect
        uint NewProtection,             // New protection (e.g., PAGE_EXECUTE_READWRITE)
        out uint OldProtection          // Old protection (output)
    );

    [DllImport("kernel32.dll", CallingConvention = CallingConvention.StdCall)]
    public static extern uint GetCurrentProcessId();

    [DllImport("kernel32.dll", CallingConvention = CallingConvention.StdCall)]
    public static extern uint GetCurrentThreadId();

    [DllImport("ntdll.dll", CallingConvention = CallingConvention.StdCall)]
    public static extern IntPtr RtlGetCurrentPeb();

    [DllImport("ntdll.dll", CallingConvention = CallingConvention.StdCall)]
    public static extern IntPtr RtlGetCurrentServiceSessionId();

    [DllImport("ntdll.dll", CallingConvention = CallingConvention.StdCall)]
    public static extern IntPtr RtlGetCurrentTransaction();
}
"@
Add-Type -TypeDefinition $TEB -ErrorAction Stop
}
Function NtCurrentTeb {
    
    <#
    Example Use
    NtCurrentTeb -Mode Buffer -Method Base
    NtCurrentTeb -Mode Buffer -Method Extend
    NtCurrentTeb -Mode Buffer -Method Protect
    NtCurrentTeb -Mode Remote -Method Base
    NtCurrentTeb -Mode Remote -Method Extend
    NtCurrentTeb -Mode Remote -Method Protect
    #>

    param (
        # Mode options for retrieving the TEB address:
        # Return   -> value returned directly in CPU register
        # Pinned   -> use a managed variable pinned in memory
        # Buffer   -> use an unmanaged temporary buffer
        # GCHandle -> use a GCHandle pinned buffer
        # Remote   -> using Callback, to receive to TEB pointer
        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateSet("Return" ,"Pinned", "Buffer", "GCHandle", "Remote")]
        [string]$Mode = "Return",

        # Allocation method for virtual memory
        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateSet("Base", "Extend", "Protect")]
        [string]$Method = "Base",

        # Optional flags to select which fields to read from TEB/PEB
        [switch]$ClientID,
        [switch]$Peb,
        [switch]$Ldr,
        [switch]$ProcessHeap,
        [switch]$Parameters,
    
        # Enable logging/debug output
        [Parameter(Mandatory = $false, Position = 7)]
        [switch]$Log = $false,

        # Self Check Function
        [Parameter(Mandatory = $false, Position = 8)]
        [switch]$SelfCheck
    )

    function Build-ASM-Shell {

        <#
        Online x86 / x64 Assembler and Disassembler
        https://defuse.ca/online-x86-assembler.htm

        add rax, 0x??           --> Add 0x?? Bytes, From Position
        mov Type, [Type + 0x??] --> Move to 0x?? Offset, read Value

        So,
        Example Read Pointer Value, 
        & Also, 
        Add 0x?? From Value

        // Return to gs:[0x00], store value at rax
        // (NtCurrentTeb) -eq ([Marshal]::ReadIntPtr((NtCurrentTeb), 0x30))
        // ([marshal]::ReadIntPtr((NtCurrentTeb),0x40)) -eq ([marshal]::ReadIntPtr((([Marshal]::ReadIntPtr((NtCurrentTeb), 0x30))),0x40))
        ** mov rax, gs:[0x30]

        // Move (de-ref`) -or Add\[+], and store value
        ** mov Type, [Type + 0x??]
        ** add rax,  0x??

        // Ret value
        ** Ret
        #>

        $shellcode        = [byte[]]@()
        $is64             = [IntPtr]::Size -gt 4
        $ret              = [byte[]]([byte]0xC3)

        if ($is64) {
            $addClient = [byte[]]@([byte]0x48,[byte]0x83,[byte]0xC0,[byte]0x40)  # add rax, 0x40          // gs:[0x40]
            $movPeb    = [byte[]]@([byte]0x48,[byte]0x8B,[byte]0x40,[byte]0x60)  # mov rax, [rax + 0x60]  // gs:[0x60] // RtlGetCurrentPeb
            $movLdr    = [byte[]]@([byte]0x48,[byte]0x8B,[byte]0x40,[byte]0x18)  # mov rax, [rax + 0x18]
            $movParams = [byte[]]@([byte]0x48,[byte]0x8B,[byte]0x40,[byte]0x20)  # mov rax, [rax + 0x20]
            $movHeap   = [byte[]]@([byte]0x48,[byte]0x8B,[byte]0x40,[byte]0x30)  # mov rax, [rax + 0x30]
            $basePtr   = [byte[]]@([byte]0x65,[byte]0x48,[byte]0x8B,[byte]0x04,  # mov rax, gs:[0x30]     #// Self dereference pointer at gs:[0x30],
                                   [byte]0x25,[byte]0x30,[byte]0x00,[byte]0x00,                           #// so, effectually, return to gs->0x0
                                   [byte]0x00)
            $InByRef  = [byte[]]@([byte]0x48,[byte]0x89,[byte]0x01)              # mov [rcx], rax         #// moves the 64-bit value from the RAX register
                                                                                                          #// into the memory location pointed to by the RCX register.
        }
        else {
            $addClient = [byte[]]@([byte]0x83,[byte]0xC0,[byte]0x20)             # add eax, 0x20          // fs:[0x20]
            $movPeb    = [byte[]]@([byte]0x8B,[byte]0x40,[byte]0x30)             # mov eax, [eax + 0x30]  // fs:[0x30] // RtlGetCurrentPeb
            $movLdr    = [byte[]]@([byte]0x8B,[byte]0x40,[byte]0x0C)             # mov eax, [eax + 0x0c]
            $movParams = [byte[]]@([byte]0x8B,[byte]0x40,[byte]0x10)             # mov eax, [eax + 0x10]
            $movHeap   = [byte[]]@([byte]0x8B,[byte]0x40,[byte]0x18)             # mov eax, [eax + 0x18]
            $basePtr   = [byte[]]@([byte]0x64,[byte]0xA1,[byte]0x18,             # mov eax, fs:[0x18]     #// Self dereference pointer at fs:[0x18], 
                                   [byte]0x00,[byte]0x00,[byte]0x00)                                      #// so, effectually, return  to fs->0x0
            $InByRef = [byte[]]@(
                [byte]0x8B, [byte]0x4C, [byte]0x24, [byte]0x04,                  # mov ecx, [esp + 4]     ; load first argument pointer from stack into ECX
                [byte]0x89, [byte]0x01                                           # mov [ecx], eax         ; store 32-bit value from EAX into memory pointed by ECX
            )
        }

        $shellcode = $basePtr
        if ($ClientID) { $shellcode += $addClient }
        if ($Peb) {
            $shellcode += $movPeb
            if ($Ldr) { $shellcode += $movLdr }
            if ($Parameters) { $shellcode += $movParams }
            if ($ProcessHeap) { $shellcode += $movHeap }
        }
        if ($Mode -ne "Return") {$shellcode += $InByRef}
        $shellcode += $ret
        $shellcode
    }

    if ($SelfCheck) {
        Clear-Host
        Write-Host
        $isX64 = [IntPtr]::Size -gt 4

        Write-Host "`nGetCurrentProcessId Test" -ForegroundColor Green
        $Offset = if ($isX64) {0x40} else {0x20}
        $procPtr = [Marshal]::ReadIntPtr((NtCurrentTeb), $Offset)
        Write-Host ("TEB offset 0x{0:X} value: {1}" -f $Offset, $procPtr)
        $clientIDProc = [Marshal]::ReadIntPtr((NtCurrentTeb -ClientID), 0x0)
        Write-Host ("ClientID Process Pointer: {0}" -f $clientIDProc)
        Write-Host ("GetCurrentProcessId(): {0}" -f [TEB]::GetCurrentProcessId())

        Write-Host "`nGetCurrentThreadId Test" -ForegroundColor Green
        $threadPtr = [Marshal]::ReadIntPtr((NtCurrentTeb), ($Offset + [IntPtr]::Size))
        Write-Host ("TEB offset 0x{0:X} value: {1}" -f ($Offset + [IntPtr]::Size), $threadPtr)
        $clientIDThread = [Marshal]::ReadIntPtr((NtCurrentTeb -ClientID), [IntPtr]::Size)
        Write-Host ("ClientID Thread Pointer: {0}" -f $clientIDThread)
        Write-Host ("GetCurrentThreadId(): {0}" -f [TEB]::GetCurrentThreadId())

        Write-Host "`nRtlGetCurrentPeb Test" -ForegroundColor Green
        $Offset = if ($isX64) {0x60} else {0x30}
        $pebPtr = [Marshal]::ReadIntPtr((NtCurrentTeb), $Offset)
        Write-Host ("TEB offset 0x{0:X} value: {1}" -f $Offset, $pebPtr)
        $pebViaFunction = NtCurrentTeb -Peb
        Write-Host ("NtCurrentTeb -Peb returned: {0}" -f $pebViaFunction)
        $pebViaTEB = [TEB]::RtlGetCurrentPeb()
        Write-Host ("RtlGetCurrentPeb(): {0}" -f $pebViaTEB)

        Write-Host "`nGetProcessHeap Test" -ForegroundColor Green
        $HeapViaFunction = NtCurrentTeb -ProcessHeap
        Write-Host ("NtCurrentTeb -ProcessHeap returned: {0}" -f $HeapViaFunction)
        $HeapViaTEB = [TEB]::GetProcessHeap()
        Write-Host ("GetProcessHeap(): {0}" -f $HeapViaTEB)
        
        Write-Host "`nRtlGetCurrentServiceSessionId Test" -ForegroundColor Green
        $serviceSessionId = [TEB]::RtlGetCurrentServiceSessionId()
        Write-Host ("Service Session Id: {0}" -f $serviceSessionId)
        $Offset = if ($isX64) {0x90} else {0x50}
        $sessionPtr = [Marshal]::ReadIntPtr((NtCurrentTeb -Peb), $Offset)
        Write-Host ("PEB offset 0x{0:X} value: {1}" -f $Offset, $sessionPtr)

        Write-Host "`nRtlGetCurrentTransaction Test" -ForegroundColor Green
        $transaction = [TEB]::RtlGetCurrentTransaction()
        Write-Host ("Current Transaction: {0}" -f $transaction)
        $Offset = if ($isX64) {0x17B8} else {0x0FAC}
        $txnPtr = [Marshal]::ReadIntPtr((NtCurrentTeb -Peb), $Offset)
        Write-Host ("PEB offset 0x{0:X} value: {1}" -f $Offset, $txnPtr)

        Write-Host "`nNtCurrentTeb Mode Test" -ForegroundColor Green
        $defaultPtr = [Marshal]::ReadIntPtr((NtCurrentTeb -Log), [IntPtr]::Size)
        Write-Host ("Default Mode Ptr: {0}" -f $defaultPtr)
        $returnPtr = [Marshal]::ReadIntPtr((NtCurrentTeb -Mode Return -Log), [IntPtr]::Size)
        Write-Host ("Return Mode Ptr: {0}" -f $returnPtr)

        $bufferPtr = [Marshal]::ReadIntPtr((NtCurrentTeb -Mode Buffer -Log), [IntPtr]::Size)
        Write-Host ("Buffer Mode Ptr: {0}" -f $bufferPtr)
        $pinnedPtr = [Marshal]::ReadIntPtr((NtCurrentTeb -Mode Pinned -Log), [IntPtr]::Size)
        Write-Host ("Pinned Mode Ptr: {0}" -f $pinnedPtr)
        $gcHandlePtr = [Marshal]::ReadIntPtr((NtCurrentTeb -Mode GCHandle -Log), [IntPtr]::Size)
        Write-Host ("GCHandle Mode Ptr: {0}" -f $gcHandlePtr)
        $callbackHandlePtr = [Marshal]::ReadIntPtr((NtCurrentTeb -Mode Remote -Log), [IntPtr]::Size)
        Write-Host ("Remote Mode Ptr: {0}" -f $callbackHandlePtr)

        Write-Host
        return
    }

    if ($Ldr -or $Parameters -or $ProcessHeap) {
      $Peb = $true
    }
    $Count = [bool]$Ldr -and [bool]$Parameters + [bool]$ProcessHeap
    if ($Count -ge 1) {
        throw "Cannot specify both -Ldr and -Parameters. Choose one."
    }
    if ($ClientID -and $Peb) {
        throw "Cannot specify both -ClientID and -Peb. Choose one."
    }

    if ($Mode -eq 'Remote') {
        [TEB]::CallbackResult = 0
        if (!$Global:MyCallbackDelegate) {
          $callbackDelegate = {
            param([IntPtr] $delPtr, [IntPtr] $valPtr)
            [TEB]::CallbackResult = $valPtr
          }
          $handle = [gchandle]::Alloc($callbackDelegate, [GCHandleType]::Normal)
          $Global:MyCallbackDelegate = $callbackDelegate
          #$Global:MyCallbackDelegate = [Teb]::GetCallback();
        }
        $callbackPtr = [Marshal]::GetFunctionPointerForDelegate(([TEB+CallbackDelegate]$Global:MyCallbackDelegate));

        [byte[]]$shellcode = $null
        if ([IntPtr]::Size -eq 8)
        {
            [byte[]]$shellcode = [byte[]]@(
                0x48, 0x89, 0xC8,                  ## mov rax, rcx      ## Move function address (first param) to rax.
                0x65, 0x48, 0x8B, 0x14, 0x25, 0x30, 0x00, 0x00, 0x00,   ## mov rdx, gs:[0x30]
                                                                        ## Set second param (rdx) from a known memory location.
                0x48, 0x83, 0xEC, 0x28,            ## sub rsp, 40       ## Allocate space on the stack for the call.
                0xFF, 0xD0,                        ## call rax          ## Call the function using the address from rax.
                0x48, 0x83, 0xC4, 0x28,            ## add rsp, 40       ## Clean up the stack.
                0xC3                               ## ret               ## Return to the caller.
            );
        }
        elseif ([IntPtr]::Size -eq 4)
        {
            [byte[]]$shellcode = [byte[]]@(
                0x64, 0xA1, 0x18, 0x00, 0x00, 0x00, ## mov eax, fs:[0x18] ## Get a specific address from the Thread Information Block.
                0x50,                               ## push eax           ## Push this address onto the stack to use later.
                0x8B, 0x44, 0x24, 0x08,             ## mov eax, [esp + 8] ## Get a second value (a function pointer or argument) from the stack.
                0x50,                               ## push eax           ## Push this value onto the stack as well.
                0xFF, 0x14, 0x24,                   ## call [esp]         ## Call the function whose address is now at the top of the stack.
                0x83, 0xC4, 0x08,                   ## add esp, 8         ## Clean up the two values we pushed on the stack.
                0xC3                                ## ret                ## Return to the calling code.
            )
        }

        $baseAddressPtr = $null
        $len = $shellcode.Length
        $lpflOldProtect = [UInt32]0
        $baseAddress = [IntPtr]::Zero
        $regionSize = [uintptr]::new($len)
    
        if ($Method -match "Base|Extend") {
        
            ## Allocate
            $ntStatus = if ($Method -eq "Base") {
                [TEB]::ZwAllocateVirtualMemory(
                   [IntPtr]::new(-1),
                   [ref]$baseAddress,
                   [UIntPtr]::new(0x00),
                   [ref]$regionSize,
                   0x3000, 0x40)
            } elseif ($Method -eq "Extend") {
                [TEB]::ZwAllocateVirtualMemoryEx(
                   [IntPtr]::new(-1),
                   [ref]$baseAddress,
                   [ref]$regionSize,
                   0x3000, 0x40,
                   [IntPtr]0,0)
            }

            if ($ntStatus -ne 0) {
                throw "ZwAllocateVirtualMemory failed with result: $ntStatus"
            }

            $Address = [IntPtr]::Zero
            [marshal]::Copy($shellcode, 0x00, $baseAddress, $len)
            ## Allocate

        } else {
        
            ## Protect
            $baseAddressPtr = [gchandle]::Alloc($shellcode, 'pinned')
            $baseAddress = $baseAddressPtr.AddrOfPinnedObject()
            [IntPtr]$tempBase = $baseAddress
            if ([TEB]::NtProtectVirtualMemory(
                    [IntPtr]::new(-1),
                    [ref]$tempBase,
                    ([ref]$regionSize),
                    0x00000040,
                    [ref]$lpflOldProtect) -ne 0) {
                throw "Fail to Protect Memory for SysCall"
            }
            ## Protect
        }

        $handle = [IntPtr]::Zero
        try
        {
            $Caller = [Marshal]::GetDelegateForFunctionPointer($baseAddress, [TEB+RemoteThreadDelgate]);
            $handle = [gchandle]::Alloc($Caller, [GCHandleType]::Normal)
            $Caller.Invoke($callbackPtr);
        }
        catch {}
        finally
        {
            Start-Sleep -Milliseconds 400
            if ($handle.IsAllocated) { $handle.Free() }

            if ($baseAddressPtr -ne $null) {
                $baseAddressPtr.Free()
            } else {
                [TEB]::ZwFreeVirtualMemory(
                    [IntPtr]::new(-1),
                    [ref]$baseAddress,
                    [ref]$regionSize,
                    0x4000) | Out-Null
            }
        }

        if ($Log) {
            Write-Warning "Mode: Remote. TypeOf: Callback Delegate"
        }
    
        if (-not [TEB]::CallbackResult -or [TEB]::CallbackResult -eq [IntPtr]::Zero) {
            throw "Failure to get results from callback!"
        }

        $isX64 = [IntPtr]::Size -eq 8

        if ($ClientID) {
            if ($isX64) {
                return [IntPtr]::Add([TEB]::CallbackResult, 0x40)
            } else {
                return [IntPtr]::Add([TEB]::CallbackResult, 0x20)
            }
        }

        if ($Peb) {
            if ($isX64) {
                $CallbackResult = [Marshal]::ReadIntPtr([TEB]::CallbackResult, 0x60)
                if ($Ldr) { $CallbackResult = [Marshal]::ReadIntPtr($CallbackResult, 0x18) }
                if ($Parameters) { $CallbackResult = [Marshal]::ReadIntPtr($CallbackResult, 0x20) }
                if ($ProcessHeap) { $CallbackResult = [Marshal]::ReadIntPtr($CallbackResult, 0x30) }
            } else {
                $CallbackResult = [Marshal]::ReadIntPtr([TEB]::CallbackResult, 0x30)
                if ($Ldr) { $CallbackResult = [Marshal]::ReadIntPtr($CallbackResult, 0x0c) }
                if ($Parameters) { $CallbackResult = [Marshal]::ReadIntPtr($CallbackResult, 0x10) }
                if ($ProcessHeap) { $CallbackResult = [Marshal]::ReadIntPtr($CallbackResult, 0x18) }
            }
            return $CallbackResult
        }

        return [TEB]::CallbackResult
    }
    
    [byte[]]$shellcode = [byte[]](Build-ASM-Shell)
    $baseAddressPtr = $null
    $len = $shellcode.Length
    $lpflOldProtect = [UInt32]0
    $baseAddress = [IntPtr]::Zero
    $regionSize = [uintptr]::new($len)
    
    if ($Method -match "Base|Extend") {
        
        ## Allocate
        $ntStatus = if ($Method -eq "Base") {
            [TEB]::ZwAllocateVirtualMemory(
               [IntPtr]::new(-1),
               [ref]$baseAddress,
               [UIntPtr]::new(0x00),
               [ref]$regionSize,
               0x3000, 0x40)
        } elseif ($Method -eq "Extend") {
            [TEB]::ZwAllocateVirtualMemoryEx(
               [IntPtr]::new(-1),
               [ref]$baseAddress,
               [ref]$regionSize,
               0x3000, 0x40,
               [IntPtr]0,0)
        }

        if ($ntStatus -ne 0) {
            throw "ZwAllocateVirtualMemory failed with result: $ntStatus"
        }

        $Address = [IntPtr]::Zero
        [marshal]::Copy($shellcode, 0x00, $baseAddress, $len)
        ## Allocate

    } else {
        
        ## Protect
        $baseAddressPtr = [gchandle]::Alloc($shellcode, 'pinned')
        $baseAddress = $baseAddressPtr.AddrOfPinnedObject()
        [IntPtr]$tempBase = $baseAddress
        if ([TEB]::NtProtectVirtualMemory(
                [IntPtr]::new(-1),
                [ref]$tempBase,
                ([ref]$regionSize),
                0x00000040,
                [ref]$lpflOldProtect) -ne 0) {
            throw "Fail to Protect Memory for SysCall"
        }
        ## Protect
    }

    try {
        switch ($Mode) {
          "Return" {
            if ($log) {
               Write-Warning "Mode: Return.   TypeOf:GetAddress"
            }
            $Address = [Marshal]::GetDelegateForFunctionPointer(
                $baseAddress,[TEB+GetAddress]).Invoke()
          }
          "Buffer" {
            if ($log) {
               Write-Warning "Mode: Buffer.   TypeOf:GetAddressByPointer"
            }
            $baseAdd = [marshal]::AllocHGlobal([IntPtr]::Size)
            [Marshal]::GetDelegateForFunctionPointer(
                $baseAddress,[TEB+GetAddressByPointer]).Invoke($baseAdd)
            $Address = [marshal]::ReadIntPtr($baseAdd)
            [marshal]::FreeHGlobal($baseAdd)
          }
          "GCHandle" {
            if ($log) {
               Write-Warning "Mode: GCHandle. TypeOf:GetAddressByPointer"
            }
            $gcHandle = [GCHandle]::Alloc($Address, [GCHandleType]::Pinned)
            $baseAdd = $gcHandle.AddrOfPinnedObject()
            [Marshal]::GetDelegateForFunctionPointer(
                $baseAddress,[TEB+GetAddressByPointer]).Invoke($baseAdd)
            $gcHandle.Free()
          }
          "Pinned" {
            if ($log) {
               Write-Warning "Mode: [REF].    TypeOf:GetAddressByReference"
            }
            [Marshal]::GetDelegateForFunctionPointer(
                $baseAddress,[TEB+GetAddressByReference]).Invoke([ref]$Address)
          }
        }
        return $Address
    }
    finally {
        if ($baseAddressPtr -ne $null) {
            $baseAddressPtr.Free()
        } else {
            [TEB]::ZwFreeVirtualMemory(
                [IntPtr]::new(-1),
                [ref]$baseAddress,
                [ref]$regionSize,
                0x4000) | Out-Null
        }
    }
}

<#
    LdrLoadDll Data Convert Helper
    ------------------------------
    
    >>>>>>>>>>>>>>>>>>>>>>>>>>>
    API-SPY --> SLUI 0x2a ERROR
    >>>>>>>>>>>>>>>>>>>>>>>>>>>

    0x00000010 - LOAD_IGNORE_CODE_AUTHZ_LEVEL
    LdrLoadDll(1,[Ref]0, 0x000000cae83fda50, 0x000000cae83fda98)
    
    0x00000008 - LOAD_WITH_ALTERED_SEARCH_PATH            
    LdrLoadDll(9,[Ref]0, 0x000000cae7fee930, 0x000000cae7fee978)
            
    0x00000800 - LOAD_LIBRARY_SEARCH_SYSTEM32
    LdrLoadDll(2049, [Ref]0, 0x000000cae83fed00, 0x000000cae83fed48 )
    
    0x00002000 -bor 0x00000008 - LOAD_LIBRARY_SAFE_CURRENT_DIRS & LOAD_WITH_ALTERED_SEARCH_PATH
    LdrLoadDll(8201, [Ref]0, 0x000000cae85fcbb0, 0x000000cae85fcbf8 )

    >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    HMODULE __stdcall LoadLibraryExW(LPCWSTR lpLibFileName,HANDLE hFile,DWORD dwFlags)
    >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    if ((dwFlags & 0x62) == 0) {
    local_res8[0] = 0;
    if ((dwFlags & 1) != 0) {
        local_res8[0] = 2;
        uVar3 = 2;
    }
    if ((char)dwFlags < '\0') {
        uVar3 = uVar3 | 0x800000;
        local_res8[0] = uVar3;
    }
    if ((dwFlags & 4) != 0) {
        uVar3 = uVar3 | 4;
        local_res8[0] = uVar3;
    }
    if ((dwFlags >> 0xf & 1) != 0) {
        local_res8[0] = uVar3 | 0x80000000;
    }
    iVar1 = LdrLoadDll(dwFlags & 0x7f08 | 1,local_res8,local_28,&local_res20);
    }

    >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    void LdrLoadDll(ulonglong param_1,uint *param_2,uint *param_3,undefined8 *param_4)
    >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

      if (param_2 == (uint *)0x0) {
        uVar4 = 0;
      }
      else {
        uVar4 = (*param_2 & 4) * 2;
        uVar3 = uVar4 | 0x40;
        if ((*param_2 & 2) == 0) {
          uVar3 = uVar4;
        }
        uVar4 = uVar3 | 0x80;
        if ((*param_2 & 0x800000) == 0) {
          uVar4 = uVar3;
        }
        uVar3 = uVar4 | 0x100;
        if ((*param_2 & 0x1000) == 0) {
          uVar3 = uVar4;
        }
        uVar4 = uVar3 | 0x400000;
        if (-1 < (int)*param_2) {
          uVar4 = uVar3;
        }
      }

    SearchPath
    ----------
    (0x00000001 -band 0x7f08) -bor 1 // DONT_RESOLVE_DLL_REFERENCES
    (0x00000010 -band 0x7f08) -bor 1 // LOAD_IGNORE_CODE_AUTHZ_LEVEL
    (0x00000200 -band 0x7f08) -bor 1 // LOAD_LIBRARY_SEARCH_APPLICATION_DIR
    (0x00001000 -band 0x7f08) -bor 1 // LOAD_LIBRARY_SEARCH_DEFAULT_DIRS
    (0x00000100 -band 0x7f08) -bor 1 // LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR
    (0x00000800 -band 0x7f08) -bor 1 // LOAD_LIBRARY_SEARCH_SYSTEM32
    (0x00000400 -band 0x7f08) -bor 1 // LOAD_LIBRARY_SEARCH_USER_DIRS
    (0x00000008 -band 0x7f08) -bor 1 // LOAD_WITH_ALTERED_SEARCH_PATH
    (0x00000080 -band 0x7f08) -bor 1 // LOAD_LIBRARY_REQUIRE_SIGNED_TARGET
    (0x00002000 -band 0x7f08) -bor 1 // LOAD_LIBRARY_SAFE_CURRENT_DIRS

    This --> will auto bypass to LoadLibraryEx?
    0x00000002, LOAD_LIBRARY_AS_DATAFILE
    0x00000040, LOAD_LIBRARY_AS_DATAFILE_EXCLUSIVE
    0x00000020, LOAD_LIBRARY_AS_IMAGE_RESOURCE

    DllCharacteristics
    ------------------
    Auto deteced by function.
    According to dwFlag value,
    who provide by user.
#>
enum LOAD_LIBRARY {
    NO_DLL_REF = 0x00000001
    IGNORE_AUTHZ = 0x00000010
    AS_DATAFILE = 0x00000002
    AS_DATAFILE_EXCL = 0x00000040
    AS_IMAGE_RES = 0x00000020
    SEARCH_APP = 0x00000200
    SEARCH_DEFAULT = 0x00001000
    SEARCH_DLL_LOAD = 0x00000100
    SEARCH_SYS32 = 0x00000800
    SEARCH_USER = 0x00000400
    ALTERED_SEARCH = 0x00000008
    REQ_SIGNED = 0x00000080
    SAFE_CURRENT = 0x00002000
}
function Ldr-LoadDll {
    param (
        [Parameter(Mandatory = $true)]
        [LOAD_LIBRARY]$dwFlags,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$dll,

        [Parameter(Mandatory = $false)]
        [switch]$Log,

        [Parameter(Mandatory = $false)]
        [switch]$ForceNew
    )
    $Zero = [IntPtr]::Zero
    $HResults, $FlagsPtr, $stringPtr = $Zero, $Zero, $Zero

    if (!$dwFlags -or [Int32]$dwFlags -eq $null) {
        throw "can't access dwFlags value"
    }

    if ([Int32]$dwFlags -lt 0) {
        throw "dwFlags Can't be less than 0"
    }

    if (-not $global:LoadedModules) {
        $global:LoadedModules = Get-LoadedModules -SortType Memory | 
            Select-Object BaseAddress, ModuleName, LoadAsData
    }

    # $ForceNew == $Log ==> $false
    $ReUseHandle = !$Log -and !$ForceNew

    # Equivalent to: if ((dwFlags & 0x62) == 0)
    # AS_DATAFILE, AS_DATAFILE_EXCLUSIVE, AS_IMAGE_RESOURCE
    $IsDataLoad = ([Int32]$dwFlags -band 0x62) -ne 0

    if ($ReUseHandle) {
        $dllObjList = $global:LoadedModules |
            Where-Object { $_.ModuleName -ieq $dll }

        if ($dllObjList) {
            if ($IsDataLoad) {
                $dllObj = $dllObjList | Where-Object { $_.LoadAsData } | Select-Object -Last 1 -ExpandProperty BaseAddress
            } else {
                $dllObj = $dllObjList | Where-Object { -not $_.LoadAsData } | Select-Object -Last 1 -ExpandProperty BaseAddress
            }

            if ($dllObj) {
                #Write-Warning "Returning reusable module object for $dll"
                return $dllObj
            }
        }
    }

    try {
        $FlagsPtr = New-IntPtr -Size 4
        if ($IsDataLoad) {
            
            # Data Load -> Begin
            if ($Log) {
                Write-host "Flags      = $([Int32]$dwFlags)"
                Write-host "SearchPath = NULL"
                Write-host "Function   = LoadLibraryExW"
                return
            }

            #Write-Warning 'Logging only --> LoadLibraryExW'
            $HResults = $Global:kernel32::LoadLibraryExW(
                $dll, [IntPtr]::Zero, [Int32]$dwFlags)
            
            if ($HResults -ne [IntPtr]::Zero) {
                $dllInfo = [PSCustomObject]@{
                    BaseAddress = $HResults
                    ModuleName  = $dll
                    LoadAsData  = $true
                }
                $global:LoadedModules += $dllInfo
            }
            return $HResults
            # Data Load -> End

        } else {
            
            # Normal Load -> Begin
            $LoadFlags = 0
            if (([Int32]$dwFlags -band 1) -ne 0) { $LoadFlags = 2 }
            if (([Int32]$dwFlags -band 0x80) -ne 0) { $LoadFlags = $LoadFlags -bor 0x800000 }
            if (([Int32]$dwFlags -band 4) -ne 0) { $LoadFlags = $LoadFlags -bor 4 }
            if ((([Int32]$dwFlags -shr 15) -band 1) -ne 0) { $LoadFlags = $LoadFlags -bor 0x80000000 }
            $Flags = ([Int32]$dwFlags -band 0x7f08) -bor 1
            $DllCharacteristics = $LoadFlags

            if ($Log) {
                Write-host "Flags      = $Flags"
                Write-host "SearchPath = $DllCharacteristics"
                Write-host "Function   = LdrLoadDll"
                return
            }

            # parameter [1] Filepath
            $FilePath = [IntPtr]::new($Flags)
            
            # parameter [2] DllCharacteristics
            [Marshal]::WriteInt32(
                $FlagsPtr, $DllCharacteristics)
            
            # parameter [3] UnicodeString
            $stringPtr = Init-NativeString -Value $dll -Encoding Unicode
            
            # Out Results
            #Write-Warning 'Logging only --> LdrLoadDll'
            $null = $Global:ntdll::LdrLoadDll(
                $FilePath,         # Flags
                $FlagsPtr,         # NULL / [REF]Long
                $stringPtr,        # [REF]UnicodeString
                [ref]$HResults     # [Out]Handle
            )
            if ($HResults -ne [IntPtr]::Zero) {
                $dllInfo = [PSCustomObject]@{
                    BaseAddress = $HResults
                    ModuleName  = $dll
                    LoadAsData  = $false
                }
                $global:LoadedModules += $dllInfo
            }
            return $HResults
            # Normal Load -> End
        }
    }
    catch {
    }
    finally {
        $FilePath = $null
        Free-IntPtr -handle $FlagsPtr  -Method Auto
        Free-IntPtr -handle $stringPtr -Method UNICODE_STRING
    }

    return $HResults
}

<#
PEB structure (winternl.h)
https://learn.microsoft.com/en-us/windows/win32/api/winternl/ns-winternl-peb

PEB_LDR_DATA structure (winternl.h)
https://learn.microsoft.com/en-us/windows/win32/api/winternl/ns-winternl-peb_ldr_data

PEB
https://www.geoffchappell.com/studies/windows/km/ntoskrnl/inc/api/pebteb/peb/index.htm

PEB_LDR_DATA
https://www.geoffchappell.com/studies/windows/km/ntoskrnl/inc/api/ntpsapi_x/peb_ldr_data.htm?tx=185

LDR_DATA_TABLE_ENTRY
https://www.geoffchappell.com/studies/windows/km/ntoskrnl/inc/api/ntldr/ldr_data_table_entry/index.htm?tx=179,185

.........................

typedef struct PEB {
  BYTE                          Reserved1[2];
  BYTE                          BeingDebugged;
  BYTE                          Reserved2[1];
  PVOID                         Reserved3[2];
  
  PPEB_LDR_DATA                 Ldr;
  ---> Pointer to PEB_LDR_DATA struct
}

typedef struct PEB_LDR_DATA {
 0x0C, 0x10, LIST_ENTRY InLoadOrderModuleList;
 0x14, 0x20, LIST_ENTRY InMemoryOrderModuleList;
 0x1C, 0x30, LIST_ENTRY InInitializationOrderModuleList;
  ---> Pointer to LIST_ENTRY struct

}

typedef struct LIST_ENTRY {
   struct LDR_DATA_TABLE_ENTRY *Flink;
   ---> Pointer to next _LDR_DATA_TABLE_ENTRY struct
}

typedef struct LDR_DATA_TABLE_ENTRY {
    0x00 0x00 LIST_ENTRY InLoadOrderLinks;
    0x08 0x10 LIST_ENTRY InMemoryOrderLinks;
    0x10 0x20 LIST_ENTRY InInitializationOrderLinks;
    ---> Actual LIST_ENTRY struct, Not Pointer

    ...
    PVOID DllBase;
    PVOID EntryPoint;
    ...
    UNICODE_STRING FullDllName;
}

.........................

** x64 system example **
You don't get Pointer to [LDR_DATA_TABLE_ENTRY] Offset 0x0, it depend
So, you need to consider, [LinkPtr] & [+Data Offset -0x00\0x10\0x20] -> Actual Offset of Data to read

[PEB_LDR_DATA] & 0x10 -> Read Pointer -> \ List Head [LIST_ENTRY]->Flink \ -> [LDR_DATA_TABLE_ENTRY]->[LIST_ENTRY]->0x00 [AKA] InLoadOrderLinks           [& Repeat]
[PEB_LDR_DATA] & 0x20 -> Read Pointer -> \ List Head [LIST_ENTRY]->Flink \ -> [LDR_DATA_TABLE_ENTRY]->[LIST_ENTRY]->0x10 [AKA] InMemoryOrderLinks         [& Repeat]
[PEB_LDR_DATA] & 0x30 -> Read Pointer -> \ List Head [LIST_ENTRY]->Flink \ -> [LDR_DATA_TABLE_ENTRY]->[LIST_ENTRY]->0x20 [AKA] InInitializationOrderLinks [& Repeat]


.........................

- (*PPEB_LDR_DATA)->InMemoryOrderModuleList -> [LIST_ENTRY] head
- each [LIST_ENTRY] contain [*flink], which point to next [LIST_ENTRY]
- [LDR_DATA_TABLE] is also [LIST_ENTRY], first offset 0x0 is [LIST_ENTRY],
  Like this -> (LDR_DATA_TABLE_ENTRY *) = (LIST_ENTRY *)

the result of this is!
[LIST_ENTRY] head, is actually [LIST_ENTRY] And not [LDR_DATA_TABLE]
only used to start the Loop chain, to Read the next [LDR_DATA_TABLE]
and than, read next [LDR_DATA_TABLE] item from [0x0  LIST_ENTRY] InLoadOrderLinks
which is actually [0x0] flink* -> pointer to another [LDR_DATA_TABLE]

C Code ->

LIST_ENTRY* head = &Peb->Ldr->InMemoryOrderModuleList;
LIST_ENTRY* current = head->Flink;

while (current != head) {
    LDR_DATA_TABLE_ENTRY* module = (LDR_DATA_TABLE_ENTRY*)current;
    wprintf(L"Loaded Module: %wZ\n", &module->FullDllName);
    current = current->Flink;
}

Diagram ->

[PEB_LDR_DATA]
 --- InMemoryOrderModuleList (LIST_ENTRY head)
        - Flink
[LDR_DATA_TABLE_ENTRY]
 --- LIST_ENTRY InLoadOrderLinks (offset 0x0)
 --- DllBase, EntryPoint, SizeOfImage, etc.
        - Flink
Another [LDR_DATA_TABLE_ENTRY]

.........................

Managed code ? sure.
[Process]::GetCurrentProcess().Modules

#>
function Read-MemoryValue {
    param (
        [Parameter(Mandatory)]
        [IntPtr]$LinkPtr,

        [Parameter(Mandatory)]
        [int]$Offset,

        [Parameter(Mandatory)]
        [ValidateSet("IntPtr","Int16", "UInt16", "Int32", "UInt32", "UnicodeString")]
        [string]$Type
    )

    # Calculate the actual address to read from:
    $Address = [IntPtr]::Add($LinkPtr, $Offset)

    try {
        switch ($Type) {
            "IntPtr" {
                return [Marshal]::ReadIntPtr($Address)
            }
            "Int16" {
                return [Marshal]::ReadInt16($Address)
            }
            "UInt16" {
                $rawValue = [Marshal]::ReadInt16($Address)
                return [UInt16]($rawValue -band 0xFFFF)
            }
            "Int32" {
                return [Marshal]::ReadInt32($Address)
            }
            "UInt32" {
                return [UInt32]([Marshal]::ReadInt32($Address))
            }
            "UnicodeString" {
                try {
                    $strData = Parse-NativeString -StringPtr $Address -Encoding Unicode | Select-Object -ExpandProperty StringData
                    return $strData
                }
                catch {}
                return $null
            }
        }
    }
    catch {
        Write-Warning "Failed to read memory value at offset 0x$([Convert]::ToString($Offset,16)) (Type: $Type). Error: $_"
        return $null
    }
}
function Get-LoadedModules {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("Load", "Memory", "Init")]
        [string]$SortType = "Memory",

        [Parameter(Mandatory=$false)]
        [IntPtr]$Peb = [IntPtr]::Zero
    )

    Enum PebOffset_x86 {
        ldrOffset = 0x0C
        InLoadOrderModuleList = 0x0C
        InMemoryOrderModuleList = 0x14
        InInitializationOrderModuleList = 0x1C
        InLoadOrderLinks = 0x00
        InMemoryOrderLinks = 0x08
        InInitializationOrderLinks = 0x10
        DllBase = 0x18
        EntryPoint = 0x1C
        SizeOfImage = 0x20
        FullDllName = 0x24
        BaseDllName = 0x2C
        Flags = 0x34
        
        # ObsoleteLoadCount
        LoadCount = 0x38

        LoadReason = 0x94
        ReferenceCount = 0x9C
    }
    Enum PebOffset_x64 {
        ldrOffset = 0x18
        InLoadOrderModuleList = 0x10
        InMemoryOrderModuleList = 0x20
        InInitializationOrderModuleList = 0x30
        InLoadOrderLinks = 0x00
        InMemoryOrderLinks = 0x10
        InInitializationOrderLinks = 0x20
        DllBase = 0x30
        EntryPoint = 0x38
        SizeOfImage = 0x40
        FullDllName = 0x48
        BaseDllName = 0x58
        Flags = 0x68

        # ObsoleteLoadCount
        LoadCount = 0x6C

        LoadReason = 0x010C
        ReferenceCount = 0x0114
    }
    if ([IntPtr]::Size -eq 8) {
        $PebOffset = [PebOffset_x64]
    } else {
        $PebOffset = [PebOffset_x86]
    }

    # Get Peb address pointr
    #$Peb = NtCurrentTeb -Peb
    
    # Get PEB->Ldr address pointr
    #$ldrPtr = [Marshal]::ReadIntPtr(
    #    [IntPtr]::Add($Peb, $PebOffset::ldrOffset.value__))

    $ldrPtr = NtCurrentTeb -Ldr
    if (-not $ldrPtr -or (
        $ldrPtr -eq [IntPtr]::Zero)) {
            throw "PEB->Ldr is null. Cannot continue."
    }

    try {
        
        # Storage to hold module list
        $modules = @()

        # Determine offsets based on sorting type
        switch ($SortType) {
            "Load" {
                $ModuleListOffset   = $PebOffset::InLoadOrderModuleList.value__
                $LinkOffsetInEntry  = $PebOffset::InLoadOrderLinks.value__
            }
            "Memory" {
                $ModuleListOffset   = $PebOffset::InMemoryOrderModuleList.value__
                $LinkOffsetInEntry  = $PebOffset::InMemoryOrderLinks.value__
            }
            "Init" {
                $ModuleListOffset   = $PebOffset::InInitializationOrderModuleList.value__
                $LinkOffsetInEntry  = $PebOffset::InInitializationOrderLinks.value__
            }
        }

        <#
            PEB_LDR_DATA->?*ModuleList --> [LIST_ENTRY] Head
            Results depend on List Type, by user choice.
        #>
        $ListHeadPtr = [IntPtr]::Add($ldrPtr, $ModuleListOffset)

        <#
            *Flink --> First [LDR_DATA_TABLE_ENTRY] -> Offset of:
            InLoadOrderLinks -or InMemoryOrderLinks -or InInitializationOrderLinks
            So, you dont get base address of [LDR_DATA_TABLE_ENTRY], it shifted, depend, result from:

            * InLoadOrderLinks = if ([IntPtr]::Size -eq 8) { 0x00 } else { 0x00 }
            * InMemoryOrderLinks = if ([IntPtr]::Size -eq 8) { 0x10 } else { 0x08 }
            * InInitializationOrderLinks = if ([IntPtr]::Size -eq 8) { 0x20 } else { 0x10 }
        #>
        $NextLinkPtr = [Marshal]::ReadIntPtr($ListHeadPtr)

        <#
           Shift offset, to Fix BaseAddress MAP,
           so, calculate -> NextLinkPtr & (+Offset -StructBaseOffset) == Data Object *Fixed* Offset
           will be used later when call Read-MemoryValue function.
        #>
        $PebOffsetMap = @{}
        foreach ($Name in ("DllBase", "EntryPoint", "SizeOfImage", "FullDllName", "BaseDllName", "Flags", "LoadCount" , "LoadReason", "ReferenceCount")) {
            $PebOffsetMap[$Name] = $PebOffset.GetField($Name).GetRawConstantValue() - $LinkOffsetInEntry
        }

        # Start parse Data

        enum LdrFlagsMap {
            PackagedBinary         = 0x00000001
            MarkedForRemoval       = 0x00000002
            ImageDll               = 0x00000004
            LoadNotificationsSent  = 0x00000008
            TelemetryEntryProcessed= 0x00000010
            ProcessStaticImport    = 0x00000020
            InLegacyLists          = 0x00000040
            InIndexes              = 0x00000080
            ShimDll                = 0x00000100
            InExceptionTable       = 0x00000200
            LoadInProgress         = 0x00001000
            LoadConfigProcessed    = 0x00002000
            EntryProcessed         = 0x00004000
            ProtectDelayLoad       = 0x00008000
            DontCallForThreads     = 0x00040000
            ProcessAttachCalled    = 0x00080000
            ProcessAttachFailed    = 0x00100000
            CorDeferredValidate    = 0x00200000
            CorImage               = 0x00400000
            DontRelocate           = 0x00800000
            CorILOnly              = 0x01000000
            ChpeImage              = 0x02000000
            Redirected             = 0x10000000
            CompatDatabaseProcessed= 0x80000000
        }

        enum LdrLoadReasonMap {
            StaticDependency = 0
            StaticForwarderDependency = 1
            DynamicForwarderDependency = 2
            DelayloadDependency = 3
            DynamicLoad = 4
            AsImageLoad = 5
            AsDataLoad = 6
            EnclavePrimary = 7
            EnclaveDependency = 8
            Unknown = -1
        }

        do {

            $flagsValue = Read-MemoryValue -LinkPtr $NextLinkPtr -Offset $PebOffsetMap['Flags'] -Type UInt32
            $allFlagValues = [Enum]::GetValues([LdrFlagsMap])
            $FlagNames = $allFlagValues | ? {  ($flagsValue -band [int]$_) -ne 0  } | ForEach-Object { $_.ToString() }
            $ReadableFlags = if ($FlagNames.Count -gt 0) {  $FlagNames -join ", "  } else {  "None"  }

            $LoadReasonValue = Read-MemoryValue -LinkPtr $NextLinkPtr -Offset $PebOffsetMap['LoadReason'] -Type UInt32
            try {
                $LoadReasonName = [LdrLoadReasonMap]$LoadReasonValue
            } catch {
                $LoadReasonName = "Unknown ($LoadReasonValue)"
            }

            $modules += [PSCustomObject]@{
                BaseAddress = Read-MemoryValue -LinkPtr $NextLinkPtr -Offset $PebOffsetMap['DllBase']     -Type IntPtr
                EntryPoint  = Read-MemoryValue -LinkPtr $NextLinkPtr -Offset $PebOffsetMap['EntryPoint']  -Type IntPtr
                SizeOfImage = Read-MemoryValue -LinkPtr $NextLinkPtr -Offset $PebOffsetMap['SizeOfImage'] -Type UInt32
                FullDllName = Read-MemoryValue -LinkPtr $NextLinkPtr -Offset $PebOffsetMap['FullDllName'] -Type UnicodeString
                ModuleName  = Read-MemoryValue -LinkPtr $NextLinkPtr -Offset $PebOffsetMap['BaseDllName'] -Type UnicodeString
                Flags       = $ReadableFlags
                LoadReason  = $LoadReasonName
                ReferenceCount = Read-MemoryValue -LinkPtr $NextLinkPtr -Offset $PebOffsetMap['ReferenceCount'] -Type UInt16
                LoadAsData  = $false
            }

            <#
                [LIST_ENTRY], 0x? -> [LIST_ENTRY] ???OrderLinks
                *Flink --> Next [LIST_ENTRY] -> [LDR_DATA_TABLE_ENTRY]
                So, we Read Item Pointer for next [LIST_ENTRY], AKA [LDR_DATA_TABLE_ENTRY]
                but, again, not BaseAddress of [LDR_DATA_TABLE_ENTRY], it depend on user Req.
                [LDR_DATA_TABLE_ENTRY] --> 0x0 -> [LIST_ENTRY], [LIST_ENTRY], [LIST_ENTRY], [Actuall Data]
            #>
			
            $NextLinkPtr = [Marshal]::ReadIntPtr($NextLinkPtr)

        } while ($NextLinkPtr -ne $ListHeadPtr)
    }
    catch {
        Write-Warning "Failed to enumerate modules. Error: $_"
    }

    return $modules
}
function Get-DllHandle {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$DllName,

        [Parameter(Mandatory=$false)]
        [ValidateSet("AddReference", "SkipReference", "PinReference", IgnoreCase=$true)]
        [string]$Flags = "SkipReference"
    )

    # Enum declaration
    enum LdrGetDllHandleFlags {
        AddReference   = 0
        SkipReference  = 1
        PinReference   = 2
    }

    $FlagValue = [enum]::Parse([LdrGetDllHandleFlags], $Flags) -as [Int32]
    $StringPtr = Init-NativeString -Value $DllName -Encoding Unicode

    try {
        $DllHandle = [IntPtr]::Zero
        $Ntstatus = $Global:ntdll::LdrGetDllHandleEx(
            $FlagValue, [IntPtr]::Zero, [IntPtr]::Zero, $StringPtr, [Ref]$DllHandle)

        if ($Ntstatus -eq 0) {
            return [IntPtr]$DllHandle
        } elseif ($Ntstatus -eq 3221225781) {
            Write-Warning "DLL module not found: $DllName"
        } else {
            Write-Warning "ERROR ($($Ntstatus)): $(Parse-ErrorMessage -MessageId $Ntstatus -Flags NTSTATUS)"
        }
    }
    finally {
        Free-NativeString -StringPtr $StringPtr | Out-Null
        $DllName = $Ntstatus = $StringPtr = $null
        if ($DllHandle -eq [IntPtr]::Zero) {
            $DllHandle = $null
        }
    }

    return [IntPtr]::Zero
}

ENUM SERVICE_STATUS {
    STOPPED = 0x00000001
    START_PENDING = 0x00000002
    STOP_PENDING = 0x00000003
    RUNNING = 0x00000004
    CONTINUE_PENDING = 0x00000005
    PAUSE_PENDING = 0x00000006
    PAUSED = 0x00000007
}
function Query-Process {
    param (
        [int]    $ProcessId,
        [string] $ProcessName,
        [switch] $First
    )

    if ($ProcessId -and $ProcessName) {
        throw "or use none, or 1 option, not both"
    }
        
    $results = @()
    $ProcessHandle = [IntPtr]::Zero
    $NewProcessHandle = [IntPtr]::Zero
    $FilterSearch = $PSBoundParameters.ContainsKey('ProcessName') -or $PSBoundParameters.ContainsKey('ProcessId')

    while (!$global:ntdll::NtGetNextProcess(
            $ProcessHandle, 0x02000000, 0, 0, [ref]$NewProcessHandle)) {
        
        $Procid = 0x0
        $Procname = $null
        $buffer = [IntPtr]::Zero
        $hProcess = $NewProcessHandle

        try {

            # Get Process Info using, NtQueryInformationProcess -> 0x0 -> PROCESS_BASIC_INFORMATION
            # https://crashpad.chromium.org/doxygen/structcrashpad_1_1process__types_1_1PROCESS__BASIC__INFORMATION.html
            # x64 padding, should be 8, x86 padding should be 4, So, 32 At x64, and 16 on x86

            $retLen = 0
            $size = if ([IntPtr]::Size -gt 4) {0x30} else {0x18}
            $pbiPtr = New-IntPtr -Size $size
            $status = $global:ntdll::NtQueryInformationProcess(
                $hProcess,0,$pbiPtr,[uint32]$size, [ref]$retLen)
            if ($status -eq 0) {
                # ~~~~~~~~
                $pebOffset = if ([IntPtr]::Size -eq 8) {8} else {4}
                $Peb = [Marshal]::ReadIntPtr($pbiPtr, $pebOffset)
                # ~~~~~~~~
                $pidoffset = if ([IntPtr]::Size -eq 8) {32} else {16}
                $pidPtr = [Marshal]::ReadIntPtr($pbiPtr, $pidoffset)
                $Procid = if ([IntPtr]::Size -eq 8) { $pidPtr.ToInt64() } else { $pidPtr.ToInt32() }
                # ~~~~~~~~
                $inheritOffset = if ([IntPtr]::Size -eq 8) {40} else {20}
                $inheritPtr = [Marshal]::ReadIntPtr($pbiPtr, $inheritOffset)
                $InheritedPid = if ([IntPtr]::Size -eq 8) { $inheritPtr.ToInt64() } else { $inheritPtr.ToInt32() }
                # ~~~~~~~~
            }

            # Get Process Name using, NtQueryInformationProcess -> 0x1b
            # .. should be large enough to hold a UNICODE_STRING structure as well as the string itself.

            $bufSize, $retLen = 1024, 0
            $buffer = New-IntPtr -Size $bufSize
            $status = $global:ntdll::NtQueryInformationProcess(
                $hProcess,27, $buffer, $bufSize,[ref]$retLen)
            if ($status -eq 0) {
                $Procname = Parse-NativeString -StringPtr $buffer -Encoding Unicode | select -ExpandProperty StringData
            }
        }
        finally {
            Free-IntPtr -handle $pbiPtr
            Free-IntPtr -handle $buffer
        }

        $ProcObj = [PSCustomObject]@{
                PebBaseAddress = $Peb
                UniqueProcessId   = $procId
                InheritedFromUniqueProcessId = $InheritedPid
                ImageFileName = $procName
            }

        if ($FilterSearch) {
            $match = $false

            if ($PSBoundParameters.ContainsKey('ProcessId') -and $procId -eq $ProcessId) {
                $match = $true
            }
            if ($PSBoundParameters.ContainsKey('ProcessName')) {
                $filterName = if ($ProcessName -like '*.exe') { $ProcessName } else { "$ProcessName.exe" }
                $fullname = $procName.ToLower()
                $LiteName = $filterName.ToLower()
                if ($fullname.EndsWith($LiteName)) {
                    $match = $true
                }
            }

            if ($match) {
                $results += $ProcObj
                if ($First) {
                    break
                }
            }
        } else {
            $results += $ProcObj
        }
        Free-IntPtr -handle $ProcessHandle -Method NtHandle
        $ProcessHandle = $NewProcessHandle
        $hProcess = $null
    }
    
    # free the last object
    Free-IntPtr -handle $ProcessHandle -Method NtHandle

    return $results
}
Function Obtain-UserToken {
    param (
        [ValidateNotNullOrEmpty()]
        [String] $UserName,
        [String] $Password,
        [String] $Domain,
        [switch] $loadProfile
    )

    try {
        $Module = [AppDomain]::CurrentDomain.GetAssemblies()| ? { $_.ManifestModule.ScopeName -eq "USER" } | select -Last 1
        $USER = $Module.GetTypes()[0]
    }
    catch {
        $Module = [AppDomain]::CurrentDomain.DefineDynamicAssembly("null", 1).DefineDynamicModule("USER", $False).DefineType("null")
        @(
            @('null', 'null', [int], @()), # place holder
            @('NtDuplicateToken', 'ntdll.dll',   [int32], @([IntPtr], [Int], [IntPtr], [Int], [Int], [IntPtr].MakeByRefType())),
            @('LoadUserProfileW', 'Userenv.dll', [bool],  @([IntPtr], [IntPtr])),
            @('LogonUserExExW',   'sspicli.dll', [bool],  @([IntPtr], [IntPtr], [IntPtr], [Int], [Int], [IntPtr], [IntPtr].MakeByRefType(), [IntPtr], [IntPtr], [IntPtr], [IntPtr]))
        ) | % {
            $Module.DefinePInvokeMethod(($_[0]), ($_[1]), 22, 1, [Type]($_[2]), [Type[]]($_[3]), 1, 3).SetImplementationFlags(128) # Def` 128, fail-safe 0 
        }
        $USER = $Module.CreateType()
    }


    $phToken = [IntPtr]::Zero
    $UserNamePtr = [Marshal]::StringToHGlobalUni($UserName)
    $PasswordPtr = if ([string]::IsNullOrEmpty($Password)) { [IntPtr]::Zero } else { [Marshal]::StringToHGlobalUni($Password) }
    $DomainPtr = if ([string]::IsNullOrEmpty($Domain)) { [IntPtr]::Zero } else { [Marshal]::StringToHGlobalUni($Domain) }

    try {

        <#
          LogonUser --> LogonUserExExW
          A handle to the primary token that represents a user
          The handle must have the TOKEN_QUERY, TOKEN_DUPLICATE, and TOKEN_ASSIGN_PRIMARY access rights
          For more information, see Access Rights for Access-Token Objects
          The user represented by the token must have read and execute access to the application
          specified by the lpApplicationName or the lpCommandLine parameter.
        #>

        <#
        # Work, but actualy fail, so, no thank you
        if (!($USER::LogonUserExExW(
            $UserNamePtr, $DomainPtr, $PasswordPtr,
            0x02, # 0x02, 0x03, 0x07, 0x08
            0x00, # LOGON32_PROVIDER_DEFAULT
            [IntPtr]0, ([ref]$phToken), [IntPtr]0,
            [IntPtr]0, [IntPtr]0, [IntPtr]0))) {
                throw "LogonUserExExW Failure .!"
            }
        #>

        #<#
        if (!(Invoke-UnmanagedMethod `
            -Dll sspicli `
            -Function LogonUserExExW `
            -CallingConvention StdCall `
            -CharSet Unicode `
            -Return bool `
            -Values @(
                $UserNamePtr, $DomainPtr, $PasswordPtr,
                0x02, # 0x02, 0x03, 0x07, 0x08
                0x00, # LOGON32_PROVIDER_DEFAULT
                [IntPtr]0, ([ref]$phToken), [IntPtr]0,
                [IntPtr]0, [IntPtr]0, [IntPtr]0))) {
                    throw "LogonUserExExW Failure .!"
                }
        #>

        # according to MS article, this is primary Token
        # we can return this --> $phToken, Directly.!
        # return $phToken

        #<#

        # Duplicate token to Primary
        $hToken = [IntPtr]0
        
        $ret = $USER::NtDuplicateToken(
                $phToken,        # Existing token
                0xF01FF,         # DesiredAccess: all rights needed
                [IntPtr]0,       # ObjectAttributes
                0x02,            # SECURITY_IMPERSONATION
                0x01,            # TOKEN_PRIMARY
                ([ref]$hToken)   # New token handle
            )

        if ($ret -ne 0) {
            Free-IntPtr -handle $hToken -Method NtHandle
            throw "Failed to Call NtDuplicateToken."
        }

        if (!$loadProfile) {
            return $hToken
        }

        $dwSize = if ([IntPtr]::Size -gt 4) { 0x38 } else { 0x20 }
        $lpProfileInfo = New-IntPtr -Size $dwSize -WriteSizeAtZero
        $lpUserName = [Marshal]::StringToCoTaskMemUni($UserName)
        [Marshal]::WriteIntPtr($lpProfileInfo, 0x08, $lpUserName)
        if (!($USER::LoadUserProfileW($hToken, $lpProfileInfo))) {
                throw "Failed to Load User profile."
            }

        Free-IntPtr -handle $sessionIdPtr
        Free-IntPtr -handle $phToken -Method NtHandle
        return $hToken
    }
    finally {
        ($lpProfileInfo, $lpUserName) | % { Free-IntPtr $_ }
        ($UserNamePtr, $PasswordPtr, $DomainPtr) | % { Free-IntPtr $_ }
    }
}
Function Process-UserToken {
    param (
        [PSObject]$Params = $null,
        [IntPtr]$hToken = [IntPtr]0
    )

    try {
        $Module = [AppDomain]::CurrentDomain.GetAssemblies()| ? { $_.ManifestModule.ScopeName -eq "Token" } | select -Last 1
        $Token = $Module.GetTypes()[0]
    }
    catch {
        $Module = [AppDomain]::CurrentDomain.DefineDynamicAssembly("null", 1).DefineDynamicModule("Token", $False).DefineType("null")
        @(
            @('null', 'null', [int], @()), # place holder
            @('OpenWindowStationW',           'User32.dll',   [intptr], @([string], [Int], [Int])),
            @('GetProcessWindowStation',      'User32.dll',   [intptr], @()),
            @('SetProcessWindowStation',      'User32.dll',   [bool],   @([IntPtr])),
            @('OpenDesktopW',                 'User32.dll',   [intptr], @([string], [int], [int], [int])),
            @('SetTokenInformation',          'advapi32.dll', [bool],   @([IntPtr], [Int], [IntPtr], [Int])),
            @('WTSGetActiveConsoleSessionId', 'Kernel32.dll', [int],    @())
        ) | % {
            $Module.DefinePInvokeMethod(($_[0]), ($_[1]), 22, 1, [Type]($_[2]), [Type[]]($_[3]), 1, 3).SetImplementationFlags(128) # Def` 128, fail-safe 0 
        }
        $Token = $Module.CreateType()
    }

    if ($Params -ne $null) {
        Invoke-UnmanagedMethod -Dll "$env:windir\temp\dacl.dll" -Function RemoveAccessAllowedAcesBasedSID -Return bool -Values @($Params.hWinSta, $Params.LogonSid) | Out-Null
        Invoke-UnmanagedMethod -Dll "$env:windir\temp\dacl.dll" -Function RemoveAccessAllowedAcesBasedSID -Return bool -Values @($Params.hDesktop, $Params.LogonSid) | Out-Null
        
        Free-IntPtr -handle $Params.hToken     -Method NtHandle
        Free-IntPtr -handle $Params.LogonSid   -Method NtHandle
        Free-IntPtr -handle ($Params.hDesktop) -Method Desktop
        Free-IntPtr -handle ($Params.hWinSta)  -Method WindowStation
    }
    elseif ($hToken -ne [IntPtr]::Zero) {
        $hDesktop, $hWinSta = [IntPtr]0, [IntPtr]0
        $activeSessionIdPtr, $LogonSid = [IntPtr]0, [IntPtr]0
        $hWinSta = $Token::OpenWindowStationW("winsta0", 0x00, (0x00020000 -bor 0x00040000L))
        if ($hWinSta -eq [IntPtr]::Zero) {
            throw "OpenWindowStationW failed .!"
        }

        $WinstaOld = $Token::GetProcessWindowStation()
        if (!($Token::SetProcessWindowStation($hWinSta))) {
            throw "SetProcessWindowStation failed .!"
        }
        $hDesktop = $Token::OpenDesktopW("default", 0x00, 0x00, (0x00020000 -bor 0x00040000 -bor 0x0080 -bor 0x0001))
        $Token::SetProcessWindowStation($WinstaOld) | Out-Null

        ## Call helper DLL
        if (!(Invoke-UnmanagedMethod -Dll "$env:windir\temp\dacl.dll" -Function GetLogonSidFromToken -Return bool -Values @($hToken, ([ref]$LogonSid)))) {
            throw "GetLogonSidFromToken helper failed .!"
        }

        ## Call helper DLL
        if (!(Invoke-UnmanagedMethod -Dll "$env:windir\temp\dacl.dll" -Function AddAceToWindowStation -Return bool -Values @($hWinSta, $LogonSid))) {
            throw "AddAceToWindowStation helper failed .!"
        }

        ## Call helper DLL
        if (!(Invoke-UnmanagedMethod -Dll "$env:windir\temp\dacl.dll" -Function AddAceToDesktop -Return bool -Values @($hDesktop, $LogonSid))) {
            throw "AddAceToWindowStation helper failed .!"
        }
        
        ## any other case will fail
        if (Check-AccountType -AccType System) {
            $activeSessionId = $Token::WTSGetActiveConsoleSessionId()
            $activeSessionIdPtr = New-IntPtr -Size 4 -InitialValue $activeSessionId
            if (!(Invoke-UnmanagedMethod -Dll ADVAPI32 -Function SetTokenInformation -Return bool -Values @(
                $hToken, 0xc, $activeSessionIdPtr, 4
                )))
            {
	            Write-Warning "Fail to Set Token Information SessionId.!"
            }
        }

        return [PSObject]@{
           hWinSta  = $hWinSta
           hDesktop = $hDesktop
           LogonSid = $LogonSid
           hToken   = $hToken
        }
    }
}
function Get-ProcessHandle {
    param (
        [int]    $ProcessId,
        [string] $ProcessName,
        [string] $ServiceName,
        [switch] $Impersonat
    )

    $scManager = $tiService = [IntPtr]::Zero
    $buffer = $clientIdPtr = $attributesPtr = [IntPtr]::Zero

    function Open-HandleFromPid($ProcID, $IsProcessToken) {
        try {
            
            $handle = [IntPtr]::Zero
            if ([IntPtr]::Size -eq 8) {
                # 64-bit sizes and layout
                $clientIdSize = 16
                $objectAttrSize = 48
            } else {
                # 32-bit sizes and layout (WOW64)
                $clientIdSize = 8
                $objectAttrSize = 24
            }
            $attributesPtr = New-IntPtr -Size $objectAttrSize -WriteSizeAtZero
            $clientIdPtr   = New-IntPtr -Size $clientIdSize   -InitialValue $ProcID -UsePointerSize
            $ntStatus = $Global:ntdll::NtOpenProcess(
                [ref]$handle, (0x0080 -bor 0x0800 -bor 0x0040 -bor 0x0400), $attributesPtr, $clientIdPtr)

            if (!$Impersonat) {
                return $handle
            }
            $tokenHandle = [IntPtr]::Zero
            $ret = $Global:ntdll::NtOpenProcessToken(
                $handle, (0x02 -bor 0x01 -bor 0x08), [ref]$tokenHandle
            )
            Free-IntPtr -handle $handle -Method NtHandle
            if ($tokenHandle -eq [IntPtr]::Zero) {
                throw "NtOpenProcessToken failue .!"
            }
            if (!($Global:kernel32::ImpersonateLoggedOnUser(
                $tokenHandle))) {
                throw "ImpersonateLoggedOnUser failue .!"
            }

            $NewTokenHandle = [IntPtr]0
            $ret = $Global:ntdll::NtDuplicateToken(
                $tokenHandle,
                (0x0080 -bor 0x0100 -bor 0x08 -bor 0x02 -bor 0x01),
                [IntPtr]0, $false, 0x01,
                [ref]$NewTokenHandle
            )
            Free-IntPtr -handle $tokenHandle -Method NtHandle

            $null = $Global:kernel32::RevertToSelf()
            return $NewTokenHandle
        }
        finally {
            Free-IntPtr -handle $clientIdPtr
            Free-IntPtr -handle $attributesPtr
        }
    }
    
    try {
            if ($ProcessId -ne 0) {
                if ($Impersonat) {
                    return Open-HandleFromPid -ProcID $ProcessId -IsProcessToken $true
                } else {
                    return Open-HandleFromPid -ProcID $ProcessId
                }
            }

            if (![string]::IsNullOrEmpty($ProcessName)) {
                $proc = Query-Process -ProcessName $ProcessName -First
                if ($proc -and $proc.UniqueProcessId) {
                    if ($Impersonat) {
                        return Open-HandleFromPid -ProcID $proc.UniqueProcessId -IsProcessToken $true
                    } else {
                        return Open-HandleFromPid -ProcID $proc.UniqueProcessId
                    }
                }
                throw "Error receive ID for selected Process"
            }

            if (![string]::IsNullOrEmpty($ServiceName)) {
                $ReturnLength = 0
                $hSCManager = $Global:advapi32::OpenSCManagerW(0,0, (0x0001 -bor 0x0002))

                if ($hSCManager -eq [IntPtr]::Zero) {
                    throw "OpenSCManagerW failed to open the service manger"
                }

                $lpServiceName = [Marshal]::StringToHGlobalAuto($ServiceName)
                $hService = $Global:advapi32::OpenServiceW(
                    $hSCManager, $lpServiceName, 0x0004 -bor 0x0010)

                if ($hService -eq [IntPtr]::Zero) {
                    throw "OpenServiceW failed"
                }

                $cbBufSize = 100;
                $pcbBytesNeeded = 0
                $dwCurrentState = 0
                $lpBuffer = New-IntPtr -Size $cbBufSize
                $ret = $Global:advapi32::QueryServiceStatusEx(
                    $hService, 0, $lpBuffer, $cbBufSize, [ref]$pcbBytesNeeded)
                if (!$ret) {
                    throw "QueryServiceStatusEx failed to query status of $ServiceName Service"
                }
                $dwCurrentState = [Marshal]::ReadInt32($lpBuffer, 4)
                Write-Warning ("Service State [Cur]: {0}" -f [SERVICE_STATUS]$dwCurrentState)

                if ($dwCurrentState -ne ([Int][SERVICE_STATUS]::RUNNING)) {
                    $Ret = $Global:advapi32::StartServiceW(
                        $hService, 0, 0)
                    if (!$Ret) {
                        throw "StartServiceW failed to start $ServiceName Service"
                    }
                }

                $svcLoadCount, $svcLoadMaxTries = 0, 8
                do {
                    Start-Sleep -Milliseconds 300
                    $ret = $Global:advapi32::QueryServiceStatusEx(
                        $hService, 0, $lpBuffer, $cbBufSize, [ref]$pcbBytesNeeded)

                    if (!$ret) {
                        throw "QueryServiceStatusEx failed to query status of $ServiceName Service"
                    }

                    if ($svcLoadCount++ -ge $svcLoadMaxTries) {
                        throw "Too many tries to load $ServiceName Service"
                    }

                    $dwCurrentState = [Marshal]::ReadInt32($lpBuffer, 4)

                } while ($dwCurrentState -ne ([Int][SERVICE_STATUS]::RUNNING))
                Write-Warning ("Service State [New]: {0}" -f [SERVICE_STATUS]$dwCurrentState)
                
                start-sleep -Seconds 1
                $svcProcID = [Marshal]::ReadInt32($lpBuffer, 28)
                if ($Impersonat) {
                    return Open-HandleFromPid -ProcID $svcProcID -IsProcessToken $true
                } else {
                    return Open-HandleFromPid -ProcID $svcProcID
                }
            }
    }
    finally {
        Free-IntPtr -handle $lpBuffer
        Free-IntPtr -handle $lpServiceName
        Free-IntPtr -handle $hSCManager -Method ServiceHandle
        Free-IntPtr -handle $hService   -Method ServiceHandle
        $lpBuffer = $lpServiceName = $hSCManager = $hService = $null
    }
}
Function Get-ProcessHelper {
        param (
            [int]    $ProcessId,
            [string] $ProcessName,
            [string] $ServiceName,
            [bool]   $Impersonat
        )
    <#
        Work with this Services:
        * cmd
        * lsass
        * spoolsv
        * Winlogon
        * TrustedInstaller
        * OfficeClickToRun

        And a lot of more, some service / Apps
        like Amd* GoodSync, etc, can be easyly use too.

    #>

    if ($ProcessName) {
        $Service = Get-CimInstance -ClassName Win32_Service -Filter "PathName LIKE '%$($ProcessName).exe%'" | Select-Object -Last 1
    }
    else {
        $Service = Get-CimInstance -ClassName Win32_Service -Filter "name = '$ServiceName'" | Select-Object -Last 1
    }
    if ($Service) {
                
        if ($Service.state -eq 'Running') {
            $ProcessId = [Int32]::Parse($Service.ProcessId)
            if ($Impersonat) {
                return Get-ProcessHandle -ProcessId $ProcessId -Impersonat
            } else {
                return Get-ProcessHandle -ProcessId $ProcessId
            }
        }
        else {
            # Managed Code fail to get handle,
            # if service need [re] Start. 
            # so, i have to call function twice
            # instead, this, work on first place.!
            $ServiceName = $Service.Name
            if ($Impersonat) {
                return Get-ProcessHandle -ServiceName $ServiceName -Impersonat
            } else {
                return Get-ProcessHandle -ServiceName $ServiceName
            }
        }
    } 
    elseif ($ProcessName) {
        if ($Impersonat) {
            return Get-ProcessHandle -ProcessName $ProcessName -Impersonat
        } else {
            return Get-ProcessHandle -ProcessName $ProcessName
        }
    }
}

<#
Privilege Escalation
https://www.ired.team/offensive-security/privilege-escalation

* Primary Access Token Manipulation
* https://www.ired.team/offensive-security/privilege-escalation/t1134-access-token-manipulation

* Windows NamedPipes 101 + Privilege Escalation
* https://www.ired.team/offensive-security/privilege-escalation/windows-namedpipes-privilege-escalation

* DLL Hijacking
* https://www.ired.team/offensive-security/privilege-escalation/t1038-dll-hijacking

* WebShells
* https://www.ired.team/offensive-security/privilege-escalation/t1108-redundant-access

* Image File Execution Options Injection
* https://www.ired.team/offensive-security/privilege-escalation/t1183-image-file-execution-options-injection

* Unquoted Service Paths
* https://www.ired.team/offensive-security/privilege-escalation/unquoted-service-paths

* Pass The Hash: Privilege Escalation with Invoke-WMIExec
* https://www.ired.team/offensive-security/privilege-escalation/pass-the-hash-privilege-escalation-with-invoke-wmiexec

* Environment Variable $Path Interception
* https://www.ired.team/offensive-security/privilege-escalation/environment-variable-path-interception

* Weak Service Permissions
* https://www.ired.team/offensive-security/privilege-escalation/weak-service-permissions

--------------

* Execute a command or a program with Trusted Installer privileges.
* Copyright (C) 2022  Matthieu `Rubisetcie` Carteron

* github Source C Code
* https://github.com/RubisetCie/god-mode

* fgsec (Felipe Gaspar) · GitHub
* https://github.com/fgsec/SharpGetSystem
* https://github.com/fgsec/Offensive/tree/master
* https://github.com/fgsec/SharpTokenTheft/tree/main

--------------

SERVICE_STATUS structure (winsvc.h)
https://learn.microsoft.com/en-us/windows/win32/api/winsvc/ns-winsvc-service_status

--------------

Clear-Host
Write-Host

Invoke-Process `
    -CommandLine "cmd /k whoami" `
    -RunAsConsole `
    -WaitForExit

Invoke-Process `
    -CommandLine "cmd /k whoami" `
    -ProcessName TrustedInstaller `
    -RunAsConsole `
    -RunAsParent

Invoke-Process `
    -CommandLine "cmd /k whoami" `
    -ProcessName winlogon `
    -RunAsConsole `
    -UseDuplicatedToken

# Could Fail to start from system/TI
Write-Host 'Invoke-ProcessAsUser, As Logon' -ForegroundColor Green
Invoke-ProcessAsUser `
    -Application cmd `
    -CommandLine "/k whoami" `
    -UserName user `
    -Password 0444 `
    -Mode Logon `
    -RunAsConsole

# Work From both Normal/Admin/System/TI Account
Write-Host 'Invoke-ProcessAsUser, As Token' -ForegroundColor Green
Invoke-ProcessAsUser `
    -Application cmd `
    -CommandLine "/k whoami" `
    -UserName user `
    -Password 0444 `
    -Mode Token `
    -RunAsConsole

# Could fail to start if not system Account
Write-Host 'Invoke-ProcessAsUser, As User' -ForegroundColor Green
Invoke-ProcessAsUser `
    -Application cmd `
    -CommandLine "/k whoami" `
    -UserName user `
    -Password 0444 `
    -Mode User `
    -RunAsConsole
#>
Function Invoke-Process {
    Param (
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string] $CommandLine,

        [Parameter(Mandatory=$false)]
        [ValidateSet('lsass', 'winlogon', 'TrustedInstaller', 'cmd', 'spoolsv', 'OfficeClickToRun')]
        [string] $ProcessName,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string] $ServiceName,

        [Parameter(Mandatory=$false)]
        [switch] $WaitForExit,

        [Parameter(Mandatory=$false)]
        [switch] $RunAsConsole,

        [Parameter(Mandatory=$false)]
        [switch] $RunAsParent,

        [Parameter(Mandatory=$false)]
        [switch] $UseDuplicatedToken
    )

    try {
        $tHandle = [IntPtr]::Zero

        if ($ProcessName -or $ServiceName -or $RunAsParent -or $UseDuplicatedToken) {
            if (!($ProcessName -xor $ServiceName)) {
                throw "Please Provide ProcessName -or ServiceName"
            }
            if (!($RunAsParent -xor $UseDuplicatedToken)) {
                throw "-ProcessName or -ServiceName Parameters, Must Run with -RunAsParent or -UseDuplicatedToken"
            }
            if ($UseDuplicatedToken) {
                Write-Warning "`nToken duplication may fail for highly privileged service processes (e.g., TrustedInstaller)`ndue to restrictive Access Control Lists (ACLs) or Protected Process status.`n"
            }
        }

        $ret = Adjust-TokenPrivileges -Privilege SeDebugPrivilege -SysCall
        if (!$ret) {
            return $false
        }
        $processInfoSize = if ([IntPtr]::Size -eq 8) { 24  } else { 16  }
        $startupInfoSize = if ([IntPtr]::Size -eq 8) { 112 } else { 104 }

        $startupInfo = New-IntPtr -Size $startupInfoSize -WriteSizeAtZero
        $processInfo = New-IntPtr -Size $processInfoSize

        $flags = if ($RunAsConsole) {
            (0x00000004 -bor 0x00080000) -bor 0x00000010
        } else { 
            (0x00000004 -bor 0x00080000) -bor 0x08000000
        }
        
        # Add flags -> STARTF_USESHOWWINDOW 0x00000001
        $dwFlagsOffset = if ([IntPtr]::Size -eq 8) {0x3C} else {0x2C}
        [Marshal]::WriteInt32($startupInfo, $dwFlagsOffset, 0x00000001)

        # Add flags -> SW_SHOWNORMAL 0x00000001
        $wShowWindowOffset = if ([IntPtr]::Size -eq 8) {0x40} else {0x30}
        [Marshal]::WriteInt16($startupInfo, $wShowWindowOffset, 0x00000001)

        # Init lpAttributeList, like InitializeProcThreadAttributeList Api .!
        # Clean List, Offset 4 -> Int32 -> Value 1
        # Populate List with 1 item, in both, x64, x86:
        # Offst 0,4,8 --> Int32 -> 1,1,1, As you see later
        $lpAttributeListSize = if ([IntPtr]::Size -eq 8) {0x30} else {0x20}
        $lpAttributeList = New-IntPtr -Size $lpAttributeListSize
        [Marshal]::WriteInt32($lpAttributeList, 4, 1)

        if ($RunAsParent) {
            
            $tHandle = Get-ProcessHelper `
                -ProcessId $ProcessId `
                -ProcessName $ProcessName `
                -ServiceName $ServiceName `
                -Impersonat $false

            # Allocate unmanaged memory for the handle pointer
            if (-not (IsValid-IntPtr $tHandle)) {
                throw "Invalid Service Handle"
            }
            $handlePtr = New-IntPtr -hHandle $tHandle -MakeRefType

            # Offset based on reverse-engineered memory layout
            # Set Parent Process as $tHandle, Work the same as
            # UpdateProcThreadAttribute Api .!
            0..2 | ForEach-Object { [Marshal]::WriteInt32($lpAttributeList, ($_ * 4), [Int32]1) }
            if ([IntPtr]::Size -eq 8) {
                [Marshal]::WriteInt32($lpAttributeList, 0x1a, 2)
                [Marshal]::WriteInt32($lpAttributeList, 0x20, 8)
                [Marshal]::WriteInt64($lpAttributeList, 0x28, $handlePtr)
            } else {
                [Marshal]::WriteByte($lpAttributeList, 0x16, 0x02)
                [Marshal]::WriteByte($lpAttributeList, 0x18, 0x04)
                [Marshal]::WriteInt32($lpAttributeList, 0x1C, $handlePtr)
            }
        }
        if ($UseDuplicatedToken) {
            $tHandle = Get-ProcessHelper `
                -ProcessId $ProcessId `
                -ProcessName $ProcessName `
                -ServiceName $ServiceName `
                -Impersonat $true
        }
        
        # Now, Update -> lpAttributeList
        if ([IntPtr]::Size -eq 8) {
            [Marshal]::WriteInt64(
                $startupInfo, 0x68, $lpAttributeList.ToInt64())
        } else {
            [Marshal]::WriteInt32(
                $startupInfo, 0x44, $lpAttributeList.ToInt32())
        }
      
        $CommandLinePtr = [Marshal]::StringToHGlobalUni($CommandLine)
        if ($UseDuplicatedToken) {
            $flags = 0x00000004 -bor 0x00000010
            #$ret = Invoke-UnmanagedMethod -Dll Advapi32 -Function CreateProcessWithTokenW -CallingConvention StdCall -Return bool -CharSet Unicode -Values @(
            $ret = $Global:advapi32::CreateProcessWithTokenW(
                $tHandle,        # handle from process / user
                0x00000001,      # LOGON_WITH_PROFILE
                [IntPtr]0,       # lpApplicationName
                $CommandLinePtr, # lpCommandLine
                $flags,          # dwCreationFlags
                [IntPtr]0,       # lpEnvironment
                [IntPtr]0,       # lpCurrentDirectory
                $startupInfo,    # lpStartupInfo
                $processInfo     # lpProcessInformation
            )
        } else {
            # Call CreateProcessAsUserW (needs the duplicated Primary Token)
            $ret = $Global:kernel32::CreateProcessW(
                0, $CommandLinePtr, 0, 0, $false, $flags, 0, 0, $startupInfo, $processInfo)
        }
        
        $err = [Marshal]::GetLastWin32Error()
        if (!$ret) {
            $msg = Parse-ErrorMessage -MessageId $err -Flags HRESULT
            Write-Warning "`nCreateProcessW fail with Error: $err`n$msg"
            return $false
        }

        $hProcess = [Marshal]::ReadIntPtr($processInfo, 0x0)
        $hThread  = [Marshal]::ReadIntPtr($processInfo, [IntPtr]::Size)
        $ret = $Global:kernel32::ResumeThread($hThread)
        if (!$ret) {
            return $false
        }

        if ($WaitForExit) {
            $null = $Global:kernel32::WaitForSingleObject(
                $hProcess, 0xFFFFFFFF)
        }
        return $true
    }
    Finally {
        # Free everything
        Free-IntPtr -handle $processInfo
        Free-IntPtr -handle $startupInfo
        Free-IntPtr -handle $lpAttributeList
        Free-IntPtr -handle $CommandLinePtr
        Free-IntPtr -handle $handlePtr
        Free-IntPtr -handle $buffer
        Free-IntPtr -handle $clientIdPtr
        Free-IntPtr -handle $attributesPtr
        Free-IntPtr -handle $hProcess  -Method Handle
        Free-IntPtr -handle $hThread   -Method Handle
        Free-IntPtr -handle $tHandle   -Method NtHandle

        # Nullify everything
        $buffer = $clientIdPtr = $attributesPtr = $null
        $processInfo = $startupInfo = $lpAttributeList = $null
        $CommandLinePtr = $hProcess = $hThread = $tHandle = $handlePtr = $null
    }
}
Function Invoke-ProcessAsUser {
    Param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string] $Application,

        [Parameter(Mandatory=$false)]
        [string] $CommandLine,

        [ValidateNotNullOrEmpty()]
        [String] $UserName,
        [ValidateNotNullOrEmpty()]
        [String] $Password,
        [String] $Domain,

        [ValidateSet("Token", "Logon", "User")]
        [String] $Mode = "Logon",

        [Parameter(Mandatory=$false)]
        [switch] $RunAsConsole
    )

<#
Custom Dll based on logue project.

# Starting an Interactive Client Process in C++
# https://learn.microsoft.com/en-us/previous-versions//aa379608(v=vs.85)?redirectedfrom=MSDN

# CreateProcessAsUser example
# https://github.com/msmania/logue/blob/master/logue.cpp
#>

try {
  if (!([Io.file]::Exists("$env:windir\temp\dacl.dll"))) {
$base64String = @'
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
'@
    $compressedBytes = [System.Convert]::FromBase64String($base64String)
    $memoryStream = New-Object System.IO.MemoryStream
    $memoryStream.Write($compressedBytes, 0, $compressedBytes.Length)
    $memoryStream.Position = 0

    $decompressedStream = New-Object System.IO.Compression.GZipStream($memoryStream, [System.IO.Compression.CompressionMode]::Decompress)
    $outputMemoryStream = New-Object System.IO.MemoryStream

    $decompressedStream.CopyTo($outputMemoryStream)
    $originalBytes = $outputMemoryStream.ToArray()

    $decompressedStream.Dispose()
    $memoryStream.Dispose()
    $outputMemoryStream.Dispose()
    $dllPath = "$env:windir\temp\dacl.dll"
    [System.IO.File]::WriteAllBytes($dllPath, $originalBytes)
  }
}
catch {
  throw "Can't load dacl.dll file .!"
}
    
    try {
        $hProcess, $hThread = [IntPtr]0, [IntPtr]0
        $infoSize = [PSCustomobject]@{
            # _STARTUPINFOW   struc
            lpStartupInfoSize = if ([IntPtr]::Size -gt 4) { 0x68 } else { 0x44 }
            # _PROCESS_INFORMATION struc
            ProcessInformationSize = if ([IntPtr]::Size -gt 4) { 0x18 } else { 0x10 }
        
        }
        
        $lpProcessInformation = New-IntPtr -Size $infoSize.ProcessInformationSize
        $lpStartupInfo = New-IntPtr -Size $infoSize.lpStartupInfoSize -WriteSizeAtZero
        $flags = if ($RunAsConsole) {
            # CREATE_NEW_CONSOLE / CREATE_UNICODE_ENVIRONMENT
            0x00000010 -bor 0x00000400
        } else {
            # CREATE_NO_WINDOW / CREATE_UNICODE_ENVIRONMENT
            0x08000000 -bor 0x00000400
        }

        $AppPath = (Get-Command $Application).Source
        $ApplicationPtr = [Marshal]::StringToHGlobalUni($AppPath)
        $FullCommandLine = """$AppPath"" $CommandLine"
        $CommandLinePtr = [Marshal]::StringToHGlobalUni($FullCommandLine)
        $lpEnvironment = [IntPtr]::Zero

        Adjust-TokenPrivileges -Privilege @("SeAssignPrimaryTokenPrivilege", "SeIncreaseQuotaPrivilege", "SeImpersonatePrivilege", "SeTcbPrivilege") -SysCall | Out-Null
        $hToken = Obtain-UserToken -UserName $UserName -Password $Password -Domain $Domain -loadProfile
            
        if (!(Invoke-UnmanagedMethod -Dll Userenv -Function CreateEnvironmentBlock -Return bool -Values @(([ref]$lpEnvironment), $hToken, $false))) {
            $lastError = [Marshal]::GetLastWin32Error()
            throw "Failed to create environment block. Last error: $lastError"
        }

        $OffsetList = [PSCustomObject]@{
            WindowFlags = if ([IntPtr]::Size -eq 8) { 0xA4 } else { 0x68 }
            ShowWindowFlags = if ([IntPtr]::Size -eq 8) { 0xA8 } else { 0x6C }
            lpDesktopOff = if ([IntPtr]::Size -gt 4) { 0x10 } else { 0x08 }
        }

        if ($Mode -eq 'Logon') {
            if (Check-AccountType -AccType System){
                Write-Warning "Could fail under system Account.!"
                #return $false
            }
            $UserNamePtr = [Marshal]::StringToHGlobalUni($UserName)
            $PasswordPtr = if ([string]::IsNullOrEmpty($Password)) { [IntPtr]::Zero } else { [Marshal]::StringToHGlobalUni($Password) }
            $DomainPtr   = if ([string]::IsNullOrEmpty($Domain)) { [IntPtr]::Zero } else { [Marshal]::StringToHGlobalUni($Domain) }

            # Call internally to Advapi32->CreateProcessWithLogonCommonW->RPC call
            $ret = Invoke-UnmanagedMethod -Dll Advapi32 -Function CreateProcessWithLogonW -CallingConvention StdCall -Return bool -CharSet Unicode -Values @(
                $UserNamePtr, $DomainPtr, $PasswordPtr,
                0x00000001,
                $ApplicationPtr, $CommandLinePtr,
                $flags, $lpEnvironment, "c:\", $lpStartupInfo, $lpProcessInformation
            )
        } elseif ($Mode -eq 'Token') {

            # Prefere hToken for current User
            $hInfo = Process-UserToken -hToken $hToken

            # Set lpDesktop Info
            $lpDesktopPtr = [Marshal]::StringToHGlobalUni("winsta0\default")
            [Marshal]::WriteIntPtr($lpStartupInfo, $OffsetList.lpDesktopOff, $lpDesktopPtr)

            # Set WindowFlags to STARTF_USESHOWWINDOW (0x00000001)
            [Marshal]::WriteInt32([IntPtr]::Add($lpStartupInfo, $OffsetList.WindowFlags), 0x01)

            # Set ShowWindowFlags to SW_SHOW (5)
            [Marshal]::WriteInt32([IntPtr]::Add($lpStartupInfo, $OffsetList.ShowWindowFlags), 0x05)

            # Call internally to Advapi32->CreateProcessWithLogonCommonW->RPC call
            $homeDrive = [marshal]::StringToCoTaskMemUni("c:\")
            $ret = $Global:advapi32::CreateProcessWithTokenW(
                $hToken,
                0x00000001,
                $ApplicationPtr, $CommandLinePtr,
                $flags, $lpEnvironment, $homeDrive, $lpStartupInfo, $lpProcessInformation
            )

            # Clean Params laters
            Process-UserToken -Params $hInfo

        } elseif ($Mode -eq 'User') {
            if (!(Check-AccountType -AccType System)) {
                Write-Warning "Could fail if not system Account.!"
                #return $false
            }
            
            # Prefere hToken for current User
            $hInfo = Process-UserToken -hToken $hToken
            
            # Impersonate the user
            if (!(Invoke-UnmanagedMethod -Dll Advapi32 -Function ImpersonateLoggedOnUser -Return bool -Values @($hToken))) {
                throw "ImpersonateLoggedOnUser failed.!"
            }

            # Set lpDesktop Info
            $lpDesktopPtr = [Marshal]::StringToHGlobalUni("winsta0\default")
            [Marshal]::WriteIntPtr($lpStartupInfo, $OffsetList.lpDesktopOff, $lpDesktopPtr)

            # Set WindowFlags to STARTF_USESHOWWINDOW (0x00000001)
            [Marshal]::WriteInt32([IntPtr]::Add($lpStartupInfo, $OffsetList.WindowFlags), 0x01)

            # Set ShowWindowFlags to SW_SHOW (5)
            [Marshal]::WriteInt32([IntPtr]::Add($lpStartupInfo, $OffsetList.ShowWindowFlags), 0x05)

            # Call internally to Advapi32->Kernel32->KernelBase->CreateProcessAsUserW->CreateProcessInternalW
            $ret = Invoke-UnmanagedMethod -Dll Kernel32 -Function CreateProcessAsUserW -CallingConvention StdCall -Return bool -CharSet Unicode -Values @(
                $hToken,
                $ApplicationPtr, $CommandLinePtr,
                [IntPtr]0, [IntPtr]0, 0x00,
                $flags, $lpEnvironment, "c:\", $lpStartupInfo, $lpProcessInformation
            )

            # Revert to your original, privileged context
            Invoke-UnmanagedMethod -Dll Advapi32 -Function RevertToSelf -Return bool | Out-Null

            # Clean Params laters
            Process-UserToken -Params $hInfo
        }
        
        if (!$ret) {
            $msg = Parse-ErrorMessage -LastWin32Error
            Write-Warning "Failed with Error: $err`n$msg"
            return $false
        }
        $hProcess = [Marshal]::ReadIntPtr(
            $lpProcessInformation, 0x0)
        $hThread  = [Marshal]::ReadIntPtr(
            $lpProcessInformation, [IntPtr]::Size)
        return $true
    }
    Finally {
        ($lpEnvironment) | % { Free-IntPtr $_ -Method Heap }
        ($hProcess, $hThread) | % { Free-IntPtr $_ -Method NtHandle }
        ($lpProcessInformation, $lpStartupInfo, $ApplicationPtr) | % { Free-IntPtr $_ }
        ($CommandLinePtr, $UserNamePtr, $PasswordPtr, $DomainPtr) | % { Free-IntPtr $_ }
        ($lpDesktopPtr, $homeDrive) | % { Free-IntPtr $_ }
    }
}

<#
Examples.

Invoke-NativeProcess `
    -ImageFile cmd `
    -commandLine "/k whoami"

try {
    $hProc = Get-ProcessHandle `
        -ProcessName 'TrustedInstaller.exe'
}
catch {
    $hProc = Get-ProcessHandle `
        -ServiceName 'TrustedInstaller'
}

if ($hProc -ne [IntPtr]::Zero) {
    Invoke-NativeProcess `
        -ImageFile cmd `
        -commandLine "/k whoami" `
        -hProc $hProc
}

Invoke-NativeProcess `
    -ImageFile "notepad.exe" `
    -Register

# Could fail to start if not system Account
Write-Host 'Invoke-NativeProcess, with hToken' -ForegroundColor Green
$hToken = Obtain-UserToken `
    -UserName user `
    -Password 0444 `
    -loadProfile
Invoke-NativeProcess `
    -ImageFile cmd `
    -commandLine "/k whoami" `
    -hToken $hToken

Free-IntPtr $hToken -Method NtHandle
Free-IntPtr $hProc  -Method NtHandle

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sources.
https://ntdoc.m417z.com/ntcreateuserprocess
https://ntdoc.m417z.com/rtl_user_process_parameters

https://github.com/capt-meelo/NtCreateUserProcess
https://github.com/BlackOfWorld/NtCreateUserProcess
https://github.com/Microwave89/createuserprocess
https://github.com/peta909/NtCreateUserProcess_
https://github.com/PorLaCola25/PPID-Spoofing

PPID Spoofing & BlockDLLs with NtCreateUserProcess
https://offensivedefence.co.uk/posts/nt-create-user-process/

Making NtCreateUserProcess Work
https://captmeelo.com/redteam/maldev/2022/05/10/ntcreateuserprocess.html

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

[Original] Native API Series- Using NtCreateUserProcess to Create a Normally Working Process - Programming Technology - Kanxu
https://bbs.kanxue.com/thread-272798.htm

Creating Processes Using System Calls - Core Labs
https://www.coresecurity.com/core-labs/articles/creating-processes-using-system-calls

GitHub - D0pam1ne705-Direct-NtCreateUserProcess- Call NtCreateUserProcess directly as normal-
https://github.com/D0pam1ne705/Direct-NtCreateUserProcess

GitHub - je5442804-NtCreateUserProcess-Post- NtCreateUserProcess with CsrClientCallServer for mainstream Windows x64 version
https://github.com/je5442804/NtCreateUserProcess-Post

PS_CREATE_INFO
https://www.geoffchappell.com/studies/windows/km/ntoskrnl/inc/api/ntpsapi/ps_create_info/index.htm

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

PEB
https://www.geoffchappell.com/studies/windows/km/ntoskrnl/inc/api/pebteb/peb/index.htm

RTL_USER_PROCESS_PARAMETERS
https://www.geoffchappell.com/studies/windows/km/ntoskrnl/inc/api/pebteb/rtl_user_process_parameters.htm

Must:
* 00000000 MaximumLength    0x440 [Int32]
* 00000004 Length           0x440 [Int32]
* 00000008 Flags            0x01  [Int32]
* 00000038 DosPath          UNICODE_STRING ?
* 00000060 ImagePathName    UNICODE_STRING ?
* 00000070 CommandLine      UNICODE_STRING ?
* 00000080 Environment      Pointer [Int64]
* 000003F0 EnvironmentSize  Size_T [Size]
#>
Function Invoke-NativeProcess {
    param (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$ImageFile = "C:\Windows\System32\cmd.exe",

        [Parameter(Mandatory = $false)]
        [string]$commandLine,

        [Parameter(Mandatory = $false)]
        [IntPtr]$hProc = [IntPtr]::Zero,

        [Parameter(Mandatory = $false)]
        [IntPtr]$hToken = [IntPtr]::Zero,

        [switch]$Register,
        [switch]$Auto,
        [switch]$Log
    )

function Write-AttributeEntry {
    param (
        [IntPtr] $BasePtr,
        [Int64]  $EntryType,
        [long]   $EntryLength,
        [IntPtr] $EntryBuffer
    )
    [Marshal]::WriteInt64($BasePtr,  0x00, $EntryType)
    if ([IntPtr]::Size -eq 8) {
        [Marshal]::WriteInt64($BasePtr,  0x08, $EntryLength)
        [Marshal]::WriteIntPtr($BasePtr, 0x10, $EntryBuffer)
    } else {
        [Marshal]::WriteInt32($BasePtr,  0x08, [int]$EntryLength)
        [Marshal]::WriteIntPtr($BasePtr, 0x0C, $EntryBuffer)
    }
}
function Get-EnvironmentBlockLength {
    param(
        [Parameter(Mandatory=$true)]
        [IntPtr]$lpEnvironment
    )

    <#
    ## Based on logic of KERNEL32 -> GetEnvironmentStrings
    ## LPWCH __stdcall GetEnvironmentStringsW()

    do
    {
    v3 = -1i64;
    while ( Environment[++v3] != 0 )
        ;
    Environment += v3 + 1;
    }
    while ( *Environment );
    #>

    $CurrentPtr = [IntPtr]$lpEnvironment
    $LengthInBytes = 0

    do {
        $Character = [Marshal]::ReadInt16($CurrentPtr)
        $CurrentPtr = [IntPtr]::Add($CurrentPtr, 2)
        $LengthInBytes += 2
        if ($Character -eq 0) {
            $NextCharacter = [Marshal]::ReadInt16($CurrentPtr)
            
            if ($NextCharacter -eq 0) {
                $LengthInBytes += 2
                break
            }
        }

    } while ($true)

    return $LengthInBytes
}

    try {
        if ($hToken -ne [IntPtr]::Zero -and (
            !(Check-AccountType -AccType System)
        )) {
            Write-Warning "Could fail if not system Account.!"
            #return $false
        }

        $hProcess, $hThread = [IntPtr]::Zero, [IntPtr]::Zero
        if (-not ([System.IO.Path]::IsPathRooted($ImageFile))) {
            try {
                $resolved = (Get-Command $ImageFile -ErrorAction Stop).Source
                $ImageFile = $resolved
            }
            catch {
                Write-Error "Could not resolve full path for '$ImageFile'"
                return
            }
        }

        $CreateInfoSize = if ([IntPtr]::Size -eq 8) { 0x58 } else { 0x48 }
        $CreateInfo = New-IntPtr -Size $CreateInfoSize -WriteSizeAtZero

        # 0x08 + (4x Size_T) [+ Optional (4x Size_T) ]
        # So, 0x0->0x8, Total size, Rest, Array[], each 32/24 BYTE'S

        $attrCount = 1
        $is64Bit    = ([IntPtr]::Size -eq 8)
        $SizeOfAtt  = if ($is64Bit) { 0x20 } else { 0x18 }
        if ($Register) { $attrCount += 2 }
        if ($hProc -ne [IntPtr]::Zero) { $attrCount += 1 }
        if ($hToken  -ne [IntPtr]::Zero) { $attrCount += 1 }
        
        # even for X86, 4 bytes for TotalLength + 4 padding
        $TotalLength = 0x08 + ($SizeOfAtt * $attrCount)
        $AttributeList = New-IntPtr -Size $TotalLength -InitialValue $TotalLength -UsePointerSize

        $ImagePath = Init-NativeString -Value $ImageFile -Encoding Unicode
        if ($commandLine) {
            $Params = Init-NativeString -Value "`"$ImageFile`" $commandLine" -Encoding Unicode
        }
        else {
            $Params = 0
        }

        $OffsetList = [PSCustomObject]@{
            
            # PEB Offset
            Params = if ([IntPtr]::Size -eq 8) {0x20} else {0x10}

            # RTL_USER_PROCESS_PARAMETERS Offset
            Length = if ([IntPtr]::Size -gt 4) { 0x440 } else { 0x2C0 }
            Cur = if ([IntPtr]::Size -eq 8) { 0x38 } else { 0x24 }
            Image = if ([IntPtr]::Size -eq 8) { 0x60 } else { 0x38 }
            CmdLine = if ([IntPtr]::Size -eq 8) { 0x70 } else { 0x40 }
            Env = if ([IntPtr]::Size -eq 8) { 0x80 } else { 0x48 }
            EnvSize = if ([IntPtr]::Size -eq 8) { 0x3F0 } else { 0x290 }
            DesktopInfo = if ([IntPtr]::Size -eq 8) { 0xC0 } else { 0x78 }
            WindowFlags = if ([IntPtr]::Size -eq 8) { 0xA4 } else { 0x68 }
            ShowWindowFlags = if ([IntPtr]::Size -eq 8) { 0xA8 } else { 0x6C }
        }

        # do not use,
        # it cause memory error messege, in console window
        # create WindowTitle, DesktopInfo, ShellInfo with fake value 00

        # Create Struct manually
        # to avoid memory error

        $CleanMode = "Auto"
        $Size_T = [UintPtr]::new(0x10)
        $Parameters = New-IntPtr -Size $OffsetList.Length
        ($paramSize, $paramSize, 0x01) | % -Begin { $i = 0 } -Process { [Marshal]::WriteInt32([IntPtr]::Add($Parameters, $i++ * 4), $_); }

        # RtlCreateEnvironmentEx(0), NtCurrentPeb()->ProcessParameters; -> Environment, EnvironmentSize
        [Marshal]::WriteIntPtr(
            $Parameters,
            $OffsetList.Env,
            ([Marshal]::ReadIntPtr(
                (NtCurrentTeb -Parameters), $OffsetList.Env))
        )
        if ([IntPtr]::Size -gt 4) {
            [Marshal]::WriteInt64(
                $Parameters,
                $OffsetList.EnvSize,
                ([Marshal]::ReadInt64(
                    (NtCurrentTeb -Parameters), $OffsetList.EnvSize))
            )
        } else {
            [Marshal]::WriteInt32(
                $Parameters,
                $OffsetList.EnvSize,
                ([Marshal]::ReadInt32(
                    (NtCurrentTeb -Parameters), $OffsetList.EnvSize))
            )
        }

        if ($hToken -ne [IntPtr]::Zero) {
            $lpEnvironment = [IntPtr]::Zero
            if (!(Invoke-UnmanagedMethod -Dll Userenv -Function CreateEnvironmentBlock -Return bool -Values @(([ref]$lpEnvironment), $hToken, $false))) {
                $lastError = [Marshal]::GetLastWin32Error()
                throw "Failed to create environment block. Last error: $lastError"
            }
            $lpLength = Get-EnvironmentBlockLength -lpEnvironment $lpEnvironment
            [Marshal]::WriteIntPtr(
                $Parameters, $OffsetList.Env, $lpEnvironment)
            if ([IntPtr]::Size -gt 4) {
                [Marshal]::WriteInt64(
                    $Parameters, $OffsetList.EnvSize, $lpLength
                )
            } else {
                [Marshal]::WriteInt32(
                    $Parameters, $OffsetList.EnvSize, $lpLength
                )
            }

            Adjust-TokenPrivileges -Privilege @("SeAssignPrimaryTokenPrivilege", "SeIncreaseQuotaPrivilege", "SeImpersonatePrivilege", "SeTcbPrivilege") -SysCall | Out-Null

            ## any other case will fail
            if (Check-AccountType -AccType System) {
                $activeSessionId = Invoke-UnmanagedMethod -Dll Kernel32 -Function WTSGetActiveConsoleSessionId -Return intptr
                $activeSessionIdPtr = New-IntPtr -Size 4 -InitialValue $activeSessionId
                if (!(Invoke-UnmanagedMethod -Dll ADVAPI32 -Function SetTokenInformation -Return bool -Values @(
                    $hToken, 0xc, $activeSessionIdPtr, 4
                    )))
                {
	                Write-Warning "Fail to Set Token Information SessionId.!"
                }
            }

            $hInfo = Process-UserToken -hToken $hToken
        }

        $DosPath = Init-NativeString -Value "$env:SystemDrive\" -Encoding Unicode
        [IntPtr]$CommandLine = if ($Params -ne 0) {[IntPtr]$Params} else {[IntPtr]$ImagePath}
        $ntdll::RtlMoveMemory(([IntPtr]::Add($Parameters, $OffsetList.Cur)),     $DosPath,     $Size_T)
        $ntdll::RtlMoveMemory(([IntPtr]::Add($Parameters, $OffsetList.Image)),   $ImagePath,   $Size_T)
        $ntdll::RtlMoveMemory(([IntPtr]::Add($Parameters, $OffsetList.CmdLine)), $CommandLine, $Size_T)

        $DesktopInfo = Init-NativeString -Value "winsta0\default" -Encoding Unicode
        $ntdll::RtlMoveMemory(([IntPtr]::Add($Parameters, $OffsetList.DesktopInfo)), $DesktopInfo, $Size_T)

        # Set WindowFlags to STARTF_USESHOWWINDOW (0x00000001)
        [Marshal]::WriteInt32([IntPtr]::Add($Parameters, $OffsetList.WindowFlags), 0x01)

        # Set ShowWindowFlags to SW_SHOW (5)
        [Marshal]::WriteInt32([IntPtr]::Add($Parameters, $OffsetList.ShowWindowFlags), 0x05)

        $NtImagePath = Init-NativeString -Value "\??\$ImageFile" -Encoding Unicode
        $Length = [Marshal]::ReadInt16($NtImagePath)
        $Buffer = [Marshal]::ReadIntPtr([IntPtr]::Add($NtImagePath, [IntPtr]::Size))

        <#
            * PS_ATTRIBUTE_NUM - NtDoc
            * https://ntdoc.m417z.com/ps_attribute_num

            PsAttributeToken, // in HANDLE
            PsAttributeClientId, // out PCLIENT_ID
            PsAttributeParentProcess, // in HANDLE

            * PsAttributeValue - NtDoc
            * https://ntdoc.m417z.com/psattributevalue
        
            PS_ATTRIBUTE_TOKEN = 0x60002;
            PS_ATTRIBUTE_PARENT_PROCESS = 0x60000;
            PS_ATTRIBUTE_CLIENT_ID = 0x10003;
            PS_ATTRIBUTE_IMAGE_NAME = 0x20005;
            PS_ATTRIBUTE_IMAGE_INFO = 0x00006;
        #>

        if ($Auto) {
            
            <#
            NTSTATUS
            NTAPI
            RtlCreateProcessParametersEx(
                _Out_ PRTL_USER_PROCESS_PARAMETERS* pProcessParameters,
                _In_ PUNICODE_STRING ImagePathName,
                _In_opt_ PUNICODE_STRING DllPath,
                _In_opt_ PUNICODE_STRING CurrentDirectory,
                _In_opt_ PUNICODE_STRING CommandLine,
                _In_opt_ PVOID Environment,
                _In_opt_ PUNICODE_STRING WindowTitle,
                _In_opt_ PUNICODE_STRING DesktopInfo,
                _In_opt_ PUNICODE_STRING ShellInfo,
                _In_opt_ PUNICODE_STRING RuntimeData,
                _In_ ULONG Flags
            );
            #>

            $CleanMode = "Heap"
            Free-IntPtr $Parameters
            $Parameters = [IntPtr]::Zero
            if (0 -ne $global:ntdll::RtlCreateProcessParametersEx(
                [ref]$Parameters, $ImagePath, 0,$DosPath, $Params, 0,0,0,0,0,0x01)) {
                return $false
            }
        }

        if ($Log) {
            
            <#
            Dump-MemoryAddress `
                -Pointer $Parameters `
                -Length ([Marshal]::ReadInt32($Parameters, 0x04))
            #>

            $MaximumLength = [marshal]::ReadInt32($Parameters, 0x00)
            $Length = [marshal]::ReadInt32($Parameters, 0x04)
            $Flags  = [marshal]::ReadInt32($Parameters, 0x08)
            $EnvStr =  [Marshal]::PtrToStringUni(
                [Marshal]::ReadIntPtr([IntPtr]::Add($Parameters, $OffsetList.Env)),
                (([marshal]::ReadInt64($Parameters, $OffsetList.EnvSize)) / 2)) 
            $DosPath =  Parse-NativeString -StringPtr ([IntPtr]::Add($Parameters, $OffsetList.Cur))     -Encoding Unicode
            $ExePath =  Parse-NativeString -StringPtr ([IntPtr]::Add($Parameters, $OffsetList.Image))   -Encoding Unicode
            $cmdLine =  Parse-NativeString -StringPtr ([IntPtr]::Add($Parameters, $OffsetList.CmdLine)) -Encoding Unicode

            write-warning "Flags = $Flags"
            write-warning "Dos Path = $($DosPath.StringData)"
            write-warning "Image Path = $($ExePath.StringData)"
            write-warning "Command Line= $($cmdLine.StringData)"
            write-warning "Length, MaximumLength = $Length, $MaximumLength"
            write-warning "Environment = $EnvStr"
        }
        
        # PS_ATTRIBUTE_IMAGE_NAME
        $AttributeListPtr = [IntPtr]::Add($AttributeList, 0x08)
        Write-AttributeEntry $AttributeListPtr 0x20005 $Length $Buffer

        if ($hProc -ne [IntPtr]::Zero) {
            # Parent Process, Jump offset +32/+24
            $AttributeListPtr = [IntPtr]::Add($AttributeListPtr, $SizeOfAtt)
            Write-AttributeEntry $AttributeListPtr 0x60000 ([IntPtr]::Size) $hProc
        }

        if ($hToken -ne [IntPtr]::Zero) {
            # Parent Process, Jump offset +32/+24
            $AttributeListPtr = [IntPtr]::Add($AttributeListPtr, $SizeOfAtt)
            Write-AttributeEntry $AttributeListPtr 0x60002 ([IntPtr]::Size) $hToken
        }

        if ($Register) {
            # CLIENT_ID, Jump offset +32/+24
            $ClientSize = if ([IntPtr]::Size -gt 4) { 0x10 } else { 0x08 }
            $ClientID = New-IntPtr -Size $ClientSize
            $AttributeListPtr = [IntPtr]::Add($AttributeListPtr, $SizeOfAtt)
            Write-AttributeEntry $AttributeListPtr 0x10003 $ClientSize $ClientID

            # SECTION_IMAGE_INFORMATION, Jump offset +32/+24
            $SectionImageSize = if ([IntPtr]::Size -gt 4) { 0x40 } else { 0x30 }
            $SectionImageInformation = New-IntPtr -Size $SectionImageSize
            $AttributeListPtr = [IntPtr]::Add($AttributeListPtr, $SizeOfAtt)
            Write-AttributeEntry $AttributeListPtr 0x06 $SectionImageSize $SectionImageInformation

            # PS_CREATE_INFO, InitFlags = 3
            $InitFlagsOffset = if ([IntPtr]::Size -gt 4) { 0x10 } else { 0x08 }
            [Marshal]::WriteInt32($CreateInfo, $InitFlagsOffset, 0x3)
            
            # PS_CREATE_INFO, AdditionalFileAccess = FILE_READ_ATTRIBUTES, FILE_READ_DATA
            $AdditionalFileAccessOffset = if ([IntPtr]::Size -gt 4) { 0x14 } else { 0xC }
            [Marshal]::WriteInt32($CreateInfo, $AdditionalFileAccessOffset, [Int32](0x0080 -bor 0x0001))

            # PROCESS_CREATE_FLAGS_SUSPENDED,       0x00000200
            # THREAD_CREATE_FLAGS_CREATE_SUSPENDED, 0x00000001
            $Ret = $global:ntdll::NtCreateUserProcess(
                [ref]$hProcess, [ref]$hThread,
                0x2000000, 0x2000000, 0, 0, 0x00000200, 0x00000001,
                $Parameters, $CreateInfo, $AttributeList)
            if ($Ret -ne 0) {
                return $false
            }

            try {
                return Send-CsrClientCall `
                    -hProcess $hProcess `
                    -hThread $hThread `
                    -ImagePath $ImagePath `
                    -NtImagePath $NtImagePath `
                    -ClientID $ClientID `
                    -CreateInfo $CreateInfo
            }
            finally {
                if ($hToken -ne [IntPtr]::Zero -and $hInfo -ne $null) {
                    Process-UserToken -Params $hInfo
                }
            }
        }

        <#
          NtCreateUserProcess - NtDoc
          https://ntdoc.m417z.com/ntcreateuserprocess
          
          Process Creation Flags
          https://learn.microsoft.com/en-us/windows/win32/procthread/process-creation-flags
        #>
        $Ret = $global:ntdll::NtCreateUserProcess(
            [ref]$hProcess, [ref]$hThread,
            0x2000000, 0x2000000,            # ACCESS_MASK
            [IntPtr]0, [IntPtr]0,            # ObjectAttributes -> Null
            0x00000200,                      # PROCESS_CREATE_FLAGS_* // -bor 0x00080000 // Fail under windows 10
            0x00000000,                      # THREAD_CREATE_FLAGS_
            $Parameters,                     # RTL_USER_PROCESS_PARAMETERS *ProcessParameters
            $CreateInfo,                     # PS_CREATE_INFO *CreateInfo
            $AttributeList                   # PS_ATTRIBUTE_LIST *AttributeList
        )
        if ($hToken -ne [IntPtr]::Zero -and $hInfo -ne $null) {
            Process-UserToken -Params $hInfo
        }

        if ($Ret -eq 0) {
            return $true
        }

        return $false
    }
    finally {
        if (![string]::IsNullOrEmpty($CleanMode)) {
            Free-IntPtr -handle $Parameters -Method $CleanMode
        }
        ($hProcess, $hThread) | % { Free-IntPtr -handle $_ -Method NtHandle }
        ($Params, $DosPath, $ImagePath, $NtImagePath, $Dummy, $DesktopInfo) | % { Free-IntPtr -handle $_ -Method UNICODE_STRING }
        ($EnvPtr, $CreateInfo, $AttributeList, $ClientID) | % { Free-IntPtr -handle $_ }
        ($pbiPtr, $baseApiMsg, $unicodeBuf, $FileInfoPtr, $SectionImageInformation) | % { Free-IntPtr -handle $_ }

        $Params = $Parameters = $hProcess = $NtImagePath = [IntPtr]0l
        $CreateInfo = $AttributeList = $ImagePath = $hThread = [IntPtr]0l
        $ClientID = $SectionImageInformation = $pbiPtr = $baseApiMsg = [IntPtr]0l
        $SpaceUnicode = $unicodeBuf = $FileInfoPtr = $AttributeListPtr = [IntPtr]0l
    }
}

<#
CsrClientCall Helper,
update system new process born.
#>
function Send-CsrClientCall {
    param (
        [IntPtr]$hProcess,
        [IntPtr]$hThread,
        [IntPtr]$ImagePath,
        [IntPtr]$NtImagePath,
        [IntPtr]$ClientID,
        [IntPtr]$CreateInfo
    )

    $baseApiMsg, $AssemblyName = [IntPtr]::Zero, [IntPtr]::Zero
    $pbiPtr, $unicodeBuf, $FileInfoPtr = [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero

    try {
        # === Query process info to get PEB ===
        $size, $retLen = 0x30, 0
        $pbiPtr = New-IntPtr -Size $size
        $status = $global:ntdll::NtQueryInformationProcess($hProcess, 0, $pbiPtr, [uint32]$size, [ref]$retLen)
        if ($status -ne 0) {
            Write-Error "NtQueryInformationProcess failed with status 0x{0:X}" -f $status
            return $false
        }

        $pidoffset = if ([IntPtr]::Size -eq 8) { 32 } else { 16 }
        $pidPtr = [Marshal]::ReadIntPtr($pbiPtr, $pidoffset)
        $Procid = if ([IntPtr]::Size -eq 8) { $pidPtr.ToInt64() } else { $pidPtr.ToInt32() }

        $pebOffset = if ([IntPtr]::Size -eq 8) { 8 } else { 4 }
        $Peb = [Marshal]::ReadIntPtr($pbiPtr, $pebOffset)

        if ([IntPtr]::Size -eq 8) {
            $UniqueProcess = [Marshal]::ReadIntPtr($ClientID, 0x00)
            $UniqueThread  = [Marshal]::ReadIntPtr($ClientID, 0x08)
        }
        else {
            $UniqueProcess = [Marshal]::ReadIntPtr($ClientID, 0x00)
            $UniqueThread  = [Marshal]::ReadIntPtr($ClientID, 0x04)
        }

        # === Prepare CSR API message ===

        $offsetOf = [PSCustomObject]@{
            Size       = if ([IntPtr]::Size -eq 8) { 0x258 } else { 0x1F0 }
            Stub       = if ([IntPtr]::Size -eq 8) { 0x218 } else { 0x1C0 }

            # 64-bit offsets
            Arch       = if ([IntPtr]::Size -eq 8) { 0x250 } else { 0x1E8 }
            Flags      = if ([IntPtr]::Size -eq 8) { 0x078 } else { 0x50 }
            PFlags     = if ([IntPtr]::Size -eq 8) { 0x07C } else { 0x54 }
            hProc      = if ([IntPtr]::Size -eq 8) { 0x040 } else { 0x30 }
            hThrd      = if ([IntPtr]::Size -eq 8) { 0x048 } else { 0x34 }
            UPID       = if ([IntPtr]::Size -eq 8) { 0x050 } else { 0x38 }
            UTID       = if ([IntPtr]::Size -eq 8) { 0x058 } else { 0x3C }
            PEB        = if ([IntPtr]::Size -eq 8) { 0x240 } else { 0x1D8 }
            FHand      = if ([IntPtr]::Size -eq 8) { 0x080 } else { 0x58 }
            MAddr      = if ([IntPtr]::Size -eq 8) { 0x0C8 } else { 0x84 }
            MSize      = if ([IntPtr]::Size -eq 8) { 0x0D0 } else { 0x88 }
    
            # cfHandle, cmAddress, cmSize are offsets within some internal struct
            cfHandle   = if ([IntPtr]::Size -eq 8) { 0x18 } else { 0x0C }
            cmAddress  = if ([IntPtr]::Size -eq 8) { 0x48 } else { 0x38 }
            cmSize     = if ([IntPtr]::Size -eq 8) { 0x50 } else { 0x40 }
        }

        $baseApiMsg = New-IntPtr -Size $offsetOf.Size                                       # SizeOf -> According to VS <.!>
        [Marshal]::WriteInt16($baseApiMsg, $offsetOf.Arch, 9)                               # ProcessorArchitecture, AMD64(=9)
        [Marshal]::WriteInt32($baseApiMsg, $offsetOf.Flags, 0x0040)                         # Sxs.Flags, Must
        [Marshal]::WriteInt32($baseApiMsg, $offsetOf.PFlags, 0x4001)                        # Sxs.ProcessParameterFlags, can be 0
        [Marshal]::WriteInt64($baseApiMsg, $offsetOf.hProc, ($hProcess -bor 2))             # hProcess, Not Must, can be 0
        [Marshal]::WriteInt64($baseApiMsg, $offsetOf.hThrd, $hThread)                       # hThread,  Not Must, can be 0
        [Marshal]::WriteInt64($baseApiMsg, $offsetOf.UPID, $UniqueProcess)                  # Unique Process ID, Must!
        [Marshal]::WriteInt64($baseApiMsg, $offsetOf.UTID, $UniqueThread)                   # Unique Thread  ID, Must!
        [Marshal]::WriteInt64($baseApiMsg, $offsetOf.PEB, $Peb)                             # Proc PEB Address,  Must!

        [Marshal]::WriteInt64($baseApiMsg, $offsetOf.FHand, ([Marshal]::ReadIntPtr([IntPtr]::Add($CreateInfo, $offsetOf.cfHandle))))   # createInfo.SuccessState.FileHandle
        [Marshal]::WriteInt64($baseApiMsg, $offsetOf.MAddr, ([Marshal]::ReadIntPtr([IntPtr]::Add($CreateInfo, $offsetOf.cmAddress))))  # createInfo.SuccessState.ManifestAddress
        [Marshal]::WriteInt64($baseApiMsg, $offsetOf.MSize, ([Marshal]::ReadIntPtr([IntPtr]::Add($CreateInfo, $offsetOf.cmSize))))     # createInfo.SuccessState.ManifestSize;

        # BaseCreateProcessMessage->Sxs.Win32ImagePath
        # BaseCreateProcessMessage->Sxs.NtImagePath
        # BaseCreateProcessMessage->Sxs.CultureFallBacks
        # BaseCreateProcessMessage->Sxs.AssemblyName

        $Size = [UIntPtr]::new(16)
        $AssemblyName = Init-NativeString -Value "Custom" -Encoding Unicode
        $FallBacks    = Init-NativeString -Value "en-US"  -Encoding Unicode -Length 0x10 -MaxLength 0x14 -BufferSize 0x28

        # Define the offset mapping based on pointer size
        $Offsets = if ([intPtr]::Size -eq 8) {
            @(
                @{ Offset = 0x088; Ptr = $ImagePath },
                @{ Offset = 0x098; Ptr = $NtImagePath },
                @{ Offset = 0x100; Ptr = $FallBacks },
                @{ Offset = 0x120; Ptr = $AssemblyName }
            )
        } else {
            @(
                @{ Offset = 0x5C; Ptr = $ImagePath },
                @{ Offset = 0x64; Ptr = $NtImagePath },
                @{ Offset = 0xAC; Ptr = $FallBacks },
                @{ Offset = 0xC4; Ptr = $AssemblyName }
            )
        }

        # Perform memory copy operation based on offsets
        $Offsets | ForEach-Object {
            $destPtr = [IntPtr]::Add($baseApiMsg, $_.Offset)
            $global:ntdll::RtlMoveMemory($destPtr, $_.Ptr, $Size)
        }

        # Cleanup vars
        $Size = $null
        $destPtr = $null

        # Define the FileInfo pointer array (same offsets for 32 and 64 bit, just handled differently)
        $FileInfoPtr = New-IntPtr -Size ([IntPtr]::Size * 4)
        $FileInfoData = $Offsets | ForEach-Object { $_.Offset }

        # Write the pointer values to $FileInfoPtr
        $FileInfoData | ForEach-Object -Begin { $i = -1 } -Process {
            $dest = [IntPtr]::Add($baseApiMsg, $_)
            $position = (++$i) * [IntPtr]::Size
            [Marshal]::WriteInt64($FileInfoPtr, $position, $dest)
            $dest = $null
        }

        # === Capture CSR message ===
        $bufferPtr = [IntPtr]::Zero
        $ret = $global:ntdll::CsrCaptureMessageMultiUnicodeStringsInPlace(
            [ref]$bufferPtr, 4, $FileInfoPtr)
        if ($ret -ne 0) {
            $ntLastError = Parse-ErrorMessage -MessageId $ret -Flags NTSTATUS
            Write-Error "CsrCaptureMessageMultiUnicodeStringsInPlace failure: $ntLastError"
            return $false
        }

        # === Send CSR message ===
        # CreateProcessInternalW, Reverse engineer code From IDA
        # CsrClientCallServer(ApiMessage, CaptureBuffer, (CSR_API_NUMBER)0x1001D, 0x218u);
        $ret = $global:ntdll::CsrClientCallServer(
            $baseApiMsg, $bufferPtr, 0x1001D, $offsetOf.Stub)

        if ($ret -ne 0) {
            $ntLastError = Parse-ErrorMessage -MessageId $ret -Flags NTSTATUS
            Write-Error "CsrClientCallServer failure: $ntLastError"
            return $false
        }

        # === Resume the thread ===
        $ret = $global:ntdll::NtResumeThread(
            $hThread, 0)
        if ($ret -ne 0) {
            $ntLastError = Parse-ErrorMessage -MessageId $ret -Flags NTSTATUS
            Write-Error "NtResumeThread failure: $ntLastError"
            return $false
        }

        return $true
    }
    finally {
        Free-IntPtr -handle $pbiPtr
        Free-IntPtr -handle $unicodeBuf
        Free-IntPtr -handle $FileInfoPtr
        Free-IntPtr -handle $baseApiMsg
        Free-IntPtr -handle $AssemblyName -Method UNICODE_STRING
        Free-IntPtr -handle $FallBacks    -Method UNICODE_STRING

        $pbiPtr = $unicodeBuf = $FileInfoPtr = $null
        $baseApiMsg = $AssemblyName = $FallBacks = $null
    }
}

<#
Examples.

.DESCRIPTION
    Creates a COM object to access licensing state and related properties.
    Parses various status enums into readable strings.
    Returns a PSCustomObject with detailed licensing info or $null if unable to create the COM object.

.EXAMPLE
    $licInfo = Get-LicensingInfo
    if ($licInfo) {
        $licInfo | Format-List
    } else {
        Write-Error "Failed to retrieve licensing info."
    }
#>
function Get-LicensingInfo {
    try {
        $clsid = "AA04CA0B-7597-4F3E-99A8-36712D13D676"
        $obj = [Activator]::CreateInstance([type]::GetTypeFromCLSID($clsid))
    }
    catch {
        return $null
    }
    try {
        
        # ENUM mappings
        $licensingStatusMap = @{
            0 = "Unlicensed (LICENSING_STATUS_UNLICENSED)"
            1 = "Licensed (LICENSING_STATUS_LICENSED)"
            2 = "In Grace Period (LICENSING_STATUS_IN_GRACE_PERIOD)"
            3 = "Notification Mode (LICENSING_STATUS_NOTIFICATION)"
        }

        $gracePeriodTypeMap = @{
            0   = "Out of Box Grace Period (E_GPT_OUT_OF_BOX)"
            1   = "Hardware Out-of-Tolerance Grace Period (E_GPT_HARDWARE_OOT)"
            2   = "Time-Based Validity Grace Period (E_GPT_TIMEBASED_VALIDITY)"
            255 = "Undefined Grace Period (E_GPT_UNDEFINED)"
        }

        $channelMap = @{
            0   = "Invalid License (LB_Invalid)"
            1   = "Hardware Bound (LB_HardwareId)"
            2   = "Environment-Based License (LB_Environment)"
            4   = "BIOS COA - Certificate of Authenticity (LB_BiosCOA)"
            8   = "BIOS SLP - System Locked Pre-installation (LB_BiosSLP)"
            16  = "BIOS Hardware ID License (LB_BiosHardwareID)"
            32  = "Token-Based Activation (LB_TokenActivation)"
            64  = "Automatic Virtual Machine Activation (LB_AutomaticVMActivation)"
            17  = "Hardware Binding - Any (LB_BindingHardwareAny)"
            12  = "BIOS Binding - Any (LB_BindingBiosAny)"
            28  = "BIOS Channel - Any (LB_ChannelBiosAny)"
            -1  = "Any Channel - Wildcard (LB_ChannelAny)"
        }

        $activationReasonMap = @{
            0   = "Generic Activation Error (E_AR_GENERIC_ERROR)"
            1   = "Activated Successfully (E_AR_ACTIVATED)"
            2   = "Invalid Product Key (E_AR_INVALID_PK)"
            3   = "Product Key Already Used (E_AR_USED_PRODUCT_KEY)"
            4   = "No Internet Connection (E_AR_NO_INTERNET)"
            5   = "Unexpected Error During Activation (E_AR_UNEXPECTED_ERROR)"
            6   = "Cannot Activate in Safe Mode (E_AR_SAFE_MODE_ERROR)"
            7   = "System State Error Preventing Activation (E_AR_SYSTEM_STATE_ERROR)"
            8   = "OEM COA Error (E_AR_OEM_COA_ERROR)"
            9   = "Expired License(s) (E_AR_EXPIRED_LICENSES)"
            10  = "No Product Key Installed (E_AR_NO_PKEY_INSTALLED)"
            11  = "Tampering Detected (E_AR_TAMPER_DETECTED)"
            12  = "Reinstallation Required for Activation (E_AR_REINSTALL_REQUIRED)"
            13  = "System Reboot Required (E_AR_REBOOT_REQUIRED)"
            14  = "Non-Genuine Windows Detected (E_AR_NON_GENUINE)"
            15  = "Token-Based Activation Error (E_AR_TOKENACTIVATION_ERROR)"
            16  = "Blocked Product Key Due to IP/Location (E_AR_BLOCKED_IPLOCATION_PK)"
            17  = "DNS Resolution Error (E_AR_DNS_ERROR)"
            18  = "Product Key Validation Error (E_VR_PRODUCTKEY_ERROR)"
            19  = "Raw Product Key Error (E_VR_PRODUCTKEY_RAW_ERROR)"
            20  = "Product Key Blocked by UI Policy (E_VR_PRODUCTKEY_UI_BLOCK)"
            255 = "Activation Reason Not Found (E_AR_NOT_FOUND)"
        }

        $systemStateFlagsMap = @{
            1  = "Reboot Policy Detected (SYSTEM_STATE_REBOOT_POLICY_FOUND)"
            2  = "System Tampering Detected (SYSTEM_STATE_TAMPERED)"
            8  = "Trusted Store Tampered (SYSTEM_STATE_TAMPERED_TRUSTED_STORE)"
            32 = "Kernel-Mode Cache Tampered (SYSTEM_STATE_TAMPERED_KM_CACHE)"
        }

        # Parse bitfield SystemStateFlags
        $stateFlags = $obj.SystemStateFlags
        $parsedStateFlags = @()
        foreach ($flag in $systemStateFlagsMap.Keys) {
            if ($stateFlags -band $flag) {
                $parsedStateFlags += $systemStateFlagsMap[$flag]
            }
        }

        $state = $obj.LicensingState
        $errMsg = Parse-ErrorMessage -MessageId $state.StatusReasonCode

        $result = [PSCustomObject]@{
            LicensingSystemDate         = $obj.LicensingSystemDate
            SystemStateFlags      = $parsedStateFlags -join ', '
            ActiveLicenseChannel  = $channelMap[$obj.ActiveLicenseChannel]
            ProductKeyType              = $obj.ProductKeyType
            IsTimebasedKeyInstalled     = [bool]$obj.IsTimebasedKeyInstalled
            DefaultKeyFromRegistry      = $obj.DefaultKeyFromRegistry
            IsLocalGenuine              = $obj.IsLocalGenuine
            skuId                   = $state.skuId
            Status            = $licensingStatusMap[$state.Status]
            StatusReasonCategory    = $activationReasonMap[$state.StatusReasonCategory]
            StatusReasonCode        = $errMsg
            Channel           = $channelMap[$state.Channel]
            GracePeriodType   = $gracePeriodTypeMap[$state.GracePeriodType]
            ValidityExpiration      = $state.ValidityExpiration
            KernelExpiration        = $state.KernelExpiration
        }

        return $result
    }
    catch {
    }
    finally {
        $null = [Marshal]::ReleaseComObject($obj)
        $null = [GC]::Collect()
        $null = [GC]::WaitForPendingFinalizers()
    }
}

<#
 Source ... 
 https://github.com/asdcorp/clic
 https://github.com/gravesoft/CAS
#>
function IsDigitalLicense {
    $interfaceSpec = Build-ComInterfaceSpec `
       -CLSID "17CCA47D-DAE5-4E4A-AC42-CC54E28F334A" `
       -IID "F2DCB80D-0670-44BC-9002-CD18688730AF" `
       -Index 5 `
       -Name AcquireModernLicenseForWindows `
       -Return int `
       -Params "int bAsync, out int lmReturnCode"

    try {
        $comObject = $interfaceSpec | Initialize-ComObject
        [int]$lmReturnCode = 0
        $hr = $comObject | Invoke-ComObject -Params (
            @(1, [ref]$lmReturnCode))

        if ($hr -eq 0) {
            return ($lmReturnCode -ne 1 -and $lmReturnCode -le [int32]::MaxValue)
        } else {
            return $false
        }

    } catch {
        Write-Warning "An error occurred: $($_.Exception.Message)"
        return $false
    } finally {
        $comObject | Release-ComObject
    }
}

<#
Source: Clic.C
Check if SubscriptionStatus
#>
function IsSubscriptionStatus {
    $dwSupported = 0 
    $ConsumeAddonPolicy = Get-ProductPolicy -Filter 'ConsumeAddonPolicySet' -UseApi
    if (-not $ConsumeAddonPolicy -or $ConsumeAddonPolicy.Value -eq $null) {
        return $false
    }
    
    $dwSupported = $ConsumeAddonPolicy.Value
    if ($dwSupported -eq 0) {
        return $false
    }

    $StatusPtr = [IntPtr]::Zero
    $ClipResult = $Global:CLIPC::ClipGetSubscriptionStatus([ref]$StatusPtr, [intPtr]::zero, [intPtr]::zero, [intPtr]::zero)
    if ($ClipResult -ne 0 -or $StatusPtr -eq [IntPtr]::Zero) {
        return $false
    }

    # so, it hold return data, no return data, no entiries
    # no entiries --> $False
    try {
        $dwStatus = [Marshal]::ReadInt32($StatusPtr)
        if ($dwStatus -and $dwStatus -gt 0) {
            return $true
        }
        return $false
    }
    finally {
        Free-IntPtr -handle $StatusPtr -Method Heap
    }
}

<#
Source: Clic.C
HRESULT WINAPI ClipGetSubscriptionStatus(
    SUBSCRIPTIONSTATUS **ppStatus
);

typedef struct _tagSUBSCRIPTIONSTATUS {
    DWORD dwEnabled;
    DWORD dwSku;
    DWORD dwState;
} SUBSCRIPTIONSTATUS;   

BOOL PrintSubscriptionStatus() {
    SUBSCRIPTIONSTATUS *pStatus;
    DWORD dwSupported = 0;

    if(SLGetWindowsInformationDWORD(L"ConsumeAddonPolicySet", &dwSupported))
        return FALSE;

    wprintf(L"SubscriptionSupportedEdition=%ws\n", BoolToWStr(dwSupported));

    if(ClipGetSubscriptionStatus(&pStatus))
        return FALSE;

    wprintf(L"SubscriptionEnabled=%ws\n", BoolToWStr(pStatus->dwEnabled));

    if(pStatus->dwEnabled == 0) {
        LocalFree(pStatus);
        return TRUE;
    }

    wprintf(L"SubscriptionSku=%d\n", pStatus->dwSku);
    wprintf(L"SubscriptionState=%d\n", pStatus->dwState);

    LocalFree(pStatus);
    return TRUE;
}

----------------------

typedef struct {
    int count;          // 4 bytes, at offset 0
    struct {
        int field1;     // 4 bytes
        int field2;     // 4 bytes
    } entries[];        // Followed by 'count' of these 8-byte pairs
} ClipSubscriptionData;

#>
function GetSubscriptionStatus {
    $StatusPtr = [IntPtr]::Zero
    $ClipResult = $Global:CLIPC::ClipGetSubscriptionStatus([ref]$StatusPtr, [intPtr]::zero, [intPtr]::zero, [intPtr]::zero)
    if ($ClipResult -ne 0 -or $StatusPtr -eq [IntPtr]::Zero) {
        return $false
    }

    try {
        $currentOffset = 4
        $subscriptionEntries = @()
        $dwStatus = [Marshal]::ReadInt32($StatusPtr)

        for ($i = 0; $i -lt $dwStatus; $i++) {
            $dwField1 = [Marshal]::ReadInt32([IntPtr]::Add($StatusPtr, $currentOffset))
            $currentOffset += 4

            $dwField2 = [Marshal]::ReadInt32([IntPtr]::Add($StatusPtr, $currentOffset))
            $currentOffset += 4

            $entry = [PSCustomObject]@{
                Sku   = $dwField1
                State = $dwField2
            }
            $subscriptionEntries += $entry
        }
        return $subscriptionEntries
    }
    finally {
        Free-IntPtr -handle $StatusPtr -Method Heap
    }
}

if ($null -eq $PSVersionTable -or $null -eq $PSVersionTable.PSVersion -or $null -eq $PSVersionTable.PSVersion.Major) {
    Clear-host
    Write-Host
    Write-Host "Unable to determine PowerShell version." -ForegroundColor Green
    Write-Host "This script requires PowerShell 5.0 or higher!" -ForegroundColor Green
    Write-Host
    Read-Host "Press Enter to exit..."
    Read-Host
    return
}

if ($PSVersionTable.PSVersion.Major -lt 5) {
    Clear-host
    Write-Host
    Write-Host "This script requires PowerShell 5.0 or higher!" -ForegroundColor Green
    Write-Host "Windows 10 & Above are supported." -ForegroundColor Green
    Write-Host
    Read-Host "Press Enter to exit..."
    Read-Host
    return
}

# Check if the current user is System or an Administrator
$isSystem = Check-AccountType -AccType System
$isAdmin  = Check-AccountType -AccType Administrator

if (![bool]$isSystem -and ![bool]$isAdmin) {
    Clear-host
    Write-Host
    if ($isSystem -eq $null -or $isAdmin -eq $null) {
        Write-Host "Unable to determine if the current user is System or Administrator." -ForegroundColor Yellow
        Write-Host "There may have been an internal error or insufficient permissions." -ForegroundColor Yellow
        return
    }
    Write-Host "This script must be run as Administrator or System!" -ForegroundColor Green
    Write-Host "Please run this script as Administrator." -ForegroundColor Green
    Write-Host "(Right-click and select 'Run as Administrator')" -ForegroundColor Green
    Write-Host
    Read-Host "Press Enter to exit..."
    Read-Host
    return
}

# LOAD DLL Function
$Global:SLC       = Init-SLC
$Global:ntdll     = Init-NTDLL
$Global:CLIPC     = Init-CLIPC
$Global:DismAPI   = Init-DismApi
$Global:PIDGENX   = Init-PIDGENX
$Global:kernel32  = Init-KERNEL32
$Global:advapi32  = Init-advapi32
$Global:PKHElper  = Init-PKHELPER
$Global:PKeyDatabase = Init-XMLInfo

# Instead of RtlGetCurrentPeb
$Global:PebPtr = NtCurrentTeb -Peb

# LOAD BASE ADDRESS for RtlFindMessage Api
$ApiMapList = @(
    # win32 errors
    "Kernel32.dll"
    "KernelBase.dll", 
    #"api-ms-win-core-synch-l1-2-0.dll",

    # NTSTATUS errors
    "ntdll.dll",

    # Activation errors
    "slc.dll",
    "sppc.dll"

    # Network Management errors
    "netmsg.dll",  # Network
    "winhttp.dll", # HTTP SERVICE
    "qmgr.dll"     # BITS
)
$baseMap = @{}
$global:LoadedModules = Get-LoadedModules -SortType Memory | 
    Select-Object BaseAddress, ModuleName, LoadAsData
$LoadedModules | Where-Object { $ApiMapList -contains $_.ModuleName } | 
    ForEach-Object { $baseMap[$_.ModuleName] = $_.BaseAddress
}
$flags = [LOAD_LIBRARY]::SEARCH_SYS32
$ApiMapList | Where-Object { $_ -notin $baseMap.Keys } | ForEach-Object {   
    $HResults = Ldr-LoadDll -dwFlags $flags -dll $_
    if ($HResults -ne [IntPtr]::Zero) {
        write-warning "LdrLoadDll Succedded to load $_"
    }
    else {
        write-warning "LdrLoadDll failed to load $_"
    }
    if ([IntPtr]::Zero -ne $HResults) {
        $baseMap[$_] = $HResults
    }
}

# Get Minimal Privileges To Load Some NtDll function
$PrivilegeList = @("SeDebugPrivilege", "SeImpersonatePrivilege", "SeIncreaseQuotaPrivilege", "SeAssignPrimaryTokenPrivilege", "SeSystemEnvironmentPrivilege")
Adjust-TokenPrivileges -Privilege $PrivilegeList -Log -SysCall

# INIT Global Variables
$Global:OfficeAppId  = '0ff1ce15-a989-479d-af46-f275c6370663'
$Global:windowsAppID  = '55c92734-d682-4d71-983e-d6ec3f16059f'
$Global:knownAppGuids = @($windowsAppID, $OfficeAppId)
$Global:CurrentVersion = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$Global:ubr = try { Get-LatestUBR } catch { 0 }
$Global:osVersion = Init-osVersion
$OperatingSystem = Get-CimInstance Win32_OperatingSystem -ea 0
$Global:OperatingSystemInfo = [PSCustomObject]@{
    dwOSMajorVersion = $Global:osVersion.Major
    dwOSMinorVersion = $Global:osVersion.Minor
    dwSpMajorVersion = try {$Global:osVersion.ServicePackMajor} catch {0};
    dwSpMinorVersion = try {$Global:osVersion.ServicePackMinor} catch {0};
}

# KeyInfo Part -->
$crc32_table = (
    0x0,
    0x04c11db7, 0x09823b6e, 0x0d4326d9, 0x130476dc, 0x17c56b6b,
    0x1a864db2, 0x1e475005, 0x2608edb8, 0x22c9f00f, 0x2f8ad6d6,
    0x2b4bcb61, 0x350c9b64, 0x31cd86d3, 0x3c8ea00a, 0x384fbdbd,
    0x4c11db70, 0x48d0c6c7, 0x4593e01e, 0x4152fda9, 0x5f15adac,
    0x5bd4b01b, 0x569796c2, 0x52568b75, 0x6a1936c8, 0x6ed82b7f,
    0x639b0da6, 0x675a1011, 0x791d4014, 0x7ddc5da3, 0x709f7b7a,
    0x745e66cd, 0x9823b6e0, 0x9ce2ab57, 0x91a18d8e, 0x95609039,
    0x8b27c03c, 0x8fe6dd8b, 0x82a5fb52, 0x8664e6e5, 0xbe2b5b58,
    0xbaea46ef, 0xb7a96036, 0xb3687d81, 0xad2f2d84, 0xa9ee3033,
    0xa4ad16ea, 0xa06c0b5d, 0xd4326d90, 0xd0f37027, 0xddb056fe,
    0xd9714b49, 0xc7361b4c, 0xc3f706fb, 0xceb42022, 0xca753d95,
    0xf23a8028, 0xf6fb9d9f, 0xfbb8bb46, 0xff79a6f1, 0xe13ef6f4,
    0xe5ffeb43, 0xe8bccd9a, 0xec7dd02d, 0x34867077, 0x30476dc0,
    0x3d044b19, 0x39c556ae, 0x278206ab, 0x23431b1c, 0x2e003dc5,
    0x2ac12072, 0x128e9dcf, 0x164f8078, 0x1b0ca6a1, 0x1fcdbb16,
    0x018aeb13, 0x054bf6a4, 0x0808d07d, 0x0cc9cdca, 0x7897ab07,
    0x7c56b6b0, 0x71159069, 0x75d48dde, 0x6b93dddb, 0x6f52c06c,
    0x6211e6b5, 0x66d0fb02, 0x5e9f46bf, 0x5a5e5b08, 0x571d7dd1,
    0x53dc6066, 0x4d9b3063, 0x495a2dd4, 0x44190b0d, 0x40d816ba,
    0xaca5c697, 0xa864db20, 0xa527fdf9, 0xa1e6e04e, 0xbfa1b04b,
    0xbb60adfc, 0xb6238b25, 0xb2e29692, 0x8aad2b2f, 0x8e6c3698,
    0x832f1041, 0x87ee0df6, 0x99a95df3, 0x9d684044, 0x902b669d,
    0x94ea7b2a, 0xe0b41de7, 0xe4750050, 0xe9362689, 0xedf73b3e,
    0xf3b06b3b, 0xf771768c, 0xfa325055, 0xfef34de2, 0xc6bcf05f,
    0xc27dede8, 0xcf3ecb31, 0xcbffd686, 0xd5b88683, 0xd1799b34,
    0xdc3abded, 0xd8fba05a, 0x690ce0ee, 0x6dcdfd59, 0x608edb80,
    0x644fc637, 0x7a089632, 0x7ec98b85, 0x738aad5c, 0x774bb0eb,
    0x4f040d56, 0x4bc510e1, 0x46863638, 0x42472b8f, 0x5c007b8a,
    0x58c1663d, 0x558240e4, 0x51435d53, 0x251d3b9e, 0x21dc2629,
    0x2c9f00f0, 0x285e1d47, 0x36194d42, 0x32d850f5, 0x3f9b762c,
    0x3b5a6b9b, 0x0315d626, 0x07d4cb91, 0x0a97ed48, 0x0e56f0ff,
    0x1011a0fa, 0x14d0bd4d, 0x19939b94, 0x1d528623, 0xf12f560e,
    0xf5ee4bb9, 0xf8ad6d60, 0xfc6c70d7, 0xe22b20d2, 0xe6ea3d65,
    0xeba91bbc, 0xef68060b, 0xd727bbb6, 0xd3e6a601, 0xdea580d8,
    0xda649d6f, 0xc423cd6a, 0xc0e2d0dd, 0xcda1f604, 0xc960ebb3,
    0xbd3e8d7e, 0xb9ff90c9, 0xb4bcb610, 0xb07daba7, 0xae3afba2,
    0xaafbe615, 0xa7b8c0cc, 0xa379dd7b, 0x9b3660c6, 0x9ff77d71,
    0x92b45ba8, 0x9675461f, 0x8832161a, 0x8cf30bad, 0x81b02d74,
    0x857130c3, 0x5d8a9099, 0x594b8d2e, 0x5408abf7, 0x50c9b640,
    0x4e8ee645, 0x4a4ffbf2, 0x470cdd2b, 0x43cdc09c, 0x7b827d21,
    0x7f436096, 0x7200464f, 0x76c15bf8, 0x68860bfd, 0x6c47164a,
    0x61043093, 0x65c52d24, 0x119b4be9, 0x155a565e, 0x18197087,
    0x1cd86d30, 0x029f3d35, 0x065e2082, 0x0b1d065b, 0x0fdc1bec,
    0x3793a651, 0x3352bbe6, 0x3e119d3f, 0x3ad08088, 0x2497d08d,
    0x2056cd3a, 0x2d15ebe3, 0x29d4f654, 0xc5a92679, 0xc1683bce,
    0xcc2b1d17, 0xc8ea00a0, 0xd6ad50a5, 0xd26c4d12, 0xdf2f6bcb,
    0xdbee767c, 0xe3a1cbc1, 0xe760d676, 0xea23f0af, 0xeee2ed18,
    0xf0a5bd1d, 0xf464a0aa, 0xf9278673, 0xfde69bc4, 0x89b8fd09,
    0x8d79e0be, 0x803ac667, 0x84fbdbd0, 0x9abc8bd5, 0x9e7d9662,
    0x933eb0bb, 0x97ffad0c, 0xafb010b1, 0xab710d06, 0xa6322bdf,
    0xa2f33668, 0xbcb4666d, 0xb8757bda, 0xb5365d03, 0xb1f740b4
);
$crc32_table = $crc32_table | % {
    $value = $_
    if ($value -lt 0) {
        $value += 0x100000000
    }
    [uint32]($value -band 0xFFFFFFFF)
}
enum SyncSource {
    U8 = 8
    U16 = 16
    U32 = 32
    U64 = 64
}
class UINT32u {
    [UInt32]   $u32
    [UInt16[]] $u16 = @(0, 0)
    [Byte[]]   $u8  = @(0, 0, 0, 0)

    [void] Sync([SyncSource]$source) {
        switch ($source.value__) {
            8 {
            	$this.u16 = 0..1 | % { [BitConverter]::ToUInt16($this.u8, $_ * 2) }
                $this.u32 = [BitConverter]::ToUInt32($this.u8, 0)
                #$this.u16 = [BitConverterHelper]::ToArrayOfType([UINT16],$this.u8)
                #$this.u32 = [BitConverterHelper]::ToArrayOfType([UINT32],$this.u8)[0]
            }
            16 {
                $this.u8 = $this.u16 | % { [BitConverter]::GetBytes($_) } | % {$_} 
                $this.u32 = [BitConverter]::ToUInt32($this.u8, 0)
                #$this.u8 = [BitConverterHelper]::ToByteArray($this.u16)
                #$this.u32 = [BitConverterHelper]::ToArrayOfType([UINT32],$this.u8)[0]
            }
            32 {
                $this.u8 = [BitConverter]::GetBytes($this.u32)
                $this.u16 = 0..1 | % { [BitConverter]::ToUInt16($this.u8, $_ * 2) }
                #$this.u8 = [BitConverter]::GetBytes($this.u32)
                #$this.u16 = [BitConverterHelper]::ToArrayOfType([UINT16],$this.u8)
            }
        }
    }
}
class UINT64u {
    [UInt64]   $u64
    [UInt32[]] $u32 = @(0, 0)
    [UInt16[]] $u16 = @(0, 0, 0, 0)
    [Byte[]]   $u8  = @(0, 0, 0, 0, 0, 0, 0, 0)

    [void] Sync([SyncSource]$source) {
        switch ($source.value__) {
            8 {
            	$this.u16 = 0..3 | % { [BitConverter]::ToUInt16($this.u8, $_ * 2) }
                $this.u32 = 0..1 | % { [BitConverter]::ToUInt32($this.u8, $_ * 4) }
                $this.u64 = [BitConverter]::ToUInt64($this.u8, 0)
                #$this.u16 = [BitConverterHelper]::ToArrayOfType([UINT16],$this.u8)
                #$this.u32 = [BitConverterHelper]::ToArrayOfType([UINT32],$this.u8)
                #$this.u64 = [BitConverterHelper]::ToArrayOfType([UINT64],$this.u8)[0]
            }
            16 {
            	$this.u8 = $this.u16 | % { [BitConverter]::GetBytes($_) } | % {$_}
                $this.u32 = 0..1 | % { [BitConverter]::ToUInt32($this.u8, $_ * 4) }
                $this.u64 = [BitConverter]::ToUInt64($this.u8, 0)
                #$this.u8 = [BitConverterHelper]::ToByteArray($this.u16)
                #$this.u32 = [BitConverterHelper]::ToArrayOfType([UINT32],$this.u8)
                #$this.u64 = [BitConverterHelper]::ToArrayOfType([UINT64],$this.u8)[0]
            }
            32 {
            	$this.u8 = $this.u32 | % { [BitConverter]::GetBytes($_) } | % {$_}
                $this.u16 = 0..3 | % { [BitConverter]::ToUInt16($this.u8, $_ * 2) }
                $this.u64 = [BitConverter]::ToUInt64($this.u8, 0)
                #$this.u8 = [BitConverterHelper]::ToByteArray($this.u32)
                #$this.u16 = [BitConverterHelper]::ToArrayOfType([UINT16],$this.u8)
                #$this.u64 = [BitConverterHelper]::ToArrayOfType([UINT64],$this.u8)[0]
            }
            64 {
                $this.u8 = [BitConverter]::GetBytes($this.u64)
                $this.u16 = 0..3 | % { [BitConverter]::ToUInt16($this.u8, $_ * 2) }
                $this.u32 = 0..1 | % { [BitConverter]::ToUInt32($this.u8, $_ * 4) }
                #$this.u8 = [BitConverterHelper]::ToByteArray($this.u64)
                #$this.u16 = [BitConverterHelper]::ToArrayOfType([UINT16],$this.u8)
                #$this.u32 = [BitConverterHelper]::ToArrayOfType([UINT32],$this.u8)
            }
        }
    }
}
class UINT128u {
    [UInt64[]] $u64 = @(0, 0)
    [UInt32[]] $u32 = @(0, 0, 0, 0)
    [UInt16[]] $u16 = @(0, 0, 0, 0, 0, 0, 0, 0)
    [Byte[]]   $u8  = @(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

    [void] Sync([SyncSource]$source) {
        switch ($source.value__) {
            8 {
                $this.u16 = 0..7 | % { [BitConverter]::ToUInt16($this.u8, $_ * 2) }
                $this.u32 = 0..3 | % { [BitConverter]::ToUInt32($this.u8, $_ * 4) }
                $this.u64 = 0..1 | % { [BitConverter]::ToUInt64($this.u8, $_ * 8) }
                #$this.u16 = [BitConverterHelper]::ToArrayOfType([UINT16],$this.u8)
                #$this.u32 = [BitConverterHelper]::ToArrayOfType([UINT32],$this.u8)
                #$this.u64 = [BitConverterHelper]::ToArrayOfType([UINT64],$this.u8)
            }
            16 {
                $this.u8 = $this.u16 | % { [BitConverter]::GetBytes($_) } | % {$_}
                $this.u32 = 0..3 | % { [BitConverter]::ToUInt32($this.u8, $_ * 4) }
                $this.u64 = 0..1 | % { [BitConverter]::ToUInt64($this.u8, $_ * 8) }
                #$this.u8 = [BitConverterHelper]::ToByteArray($this.u16)
                #$this.u32 = [BitConverterHelper]::ToArrayOfType([UINT32],$this.u8)
                #$this.u64 = [BitConverterHelper]::ToArrayOfType([UINT64],$this.u8)
            }
            32 {
                $this.u8 = $this.u32 | % { [BitConverter]::GetBytes($_) } | % {$_}
                $this.u16 = 0..7 | % { [BitConverter]::ToUInt16($this.u8, $_ * 2) }
                $this.u64 = 0..1 | % { [BitConverter]::ToUInt64($this.u8, $_ * 8) }
                #$this.u8 = [BitConverterHelper]::ToByteArray($this.u32)
                #$this.u16 = [BitConverterHelper]::ToArrayOfType([UINT16],$this.u8)
                #$this.u64 = [BitConverterHelper]::ToArrayOfType([UINT64],$this.u8)
            }
            64 {
                $this.u8 = [BitConverter]::GetBytes($this.u64[0]) + [BitConverter]::GetBytes($this.u64[1])
                $this.u16 = 0..7 | % { [BitConverter]::ToUInt16($this.u8, $_ * 2) }
                $this.u32 = 0..3 | % { [BitConverter]::ToUInt32($this.u8, $_ * 4) }
                #$this.u8 = [BitConverterHelper]::ToByteArray($this.u64)
                #$this.u16 = [BitConverterHelper]::ToArrayOfType([UINT16],$this.u8)
                #$this.u32 = [BitConverterHelper]::ToArrayOfType([UINT32],$this.u8)
            }
        }
    }
}
class BitConverterHelper {

    # Convert array of UInt16, Int16, UInt32, Int32, UInt64, or Int64 to byte array
    static [Byte[]] ToByteArray([Object[]] $values) {
        $byteList = New-Object List[Byte]

        foreach ($value in $values) {
            if ($value -is [UInt16]) {
                $byteList.AddRange([BitConverter]::GetBytes([UInt16]$value))
            } elseif ($value -is [Int16]) {
                $byteList.AddRange([BitConverter]::GetBytes([Int16]$value))
            } elseif ($value -is [UInt32]) {
                $byteList.AddRange([BitConverter]::GetBytes([UInt32]$value))
            } elseif ($value -is [Int32]) {
                $byteList.AddRange([BitConverter]::GetBytes([Int32]$value))
            } elseif ($value -is [UInt64]) {
                $byteList.AddRange([BitConverter]::GetBytes([UInt64]$value))
            } elseif ($value -is [Int64]) {
                $byteList.AddRange([BitConverter]::GetBytes([Int64]$value))
            } else {
                throw "Unsupported type: $($value.GetType().FullName)"
            }
        }

        return $byteList.ToArray()
    }

     # Convert byte array to an array of specified types (UInt16, Int16, UInt32, Int32, UInt64, Int64)
    static [Array] ToArrayOfType([Type] $type, [Byte[]] $bytes) {
        # Determine the size of each type in bytes
        $typeName = $type.FullName
        $size = switch ($typeName) {
            "System.UInt16" { 2 }
            "System.Int16"  { 2 }
            "System.UInt32" { 4 }
            "System.Int32"  { 4 }
            "System.UInt64" { 8 }
            "System.Int64"  { 8 }
            default { throw "Unsupported type: $type" }
        }
        
        # Validate byte array length
        if ($bytes.Length % $size -ne 0) {
            throw "Byte array length must be a multiple of $size for conversion to $type."
        }

        # Prepare result list
        $count = [math]::Floor($bytes.Length / $size)
        $result = New-Object 'System.Collections.Generic.List[Object]'

        # Convert bytes to the specified type
        for ($i = 0; $i -lt $count; $i++) {
            $index = $i * $size
            if ($typeName -eq "System.UInt16") {
                $result.Add([BitConverter]::ToUInt16($bytes, $index))
            } elseif ($typeName -eq "System.Int16") {
                $result.Add([BitConverter]::ToInt16($bytes, $index))
            } elseif ($typeName -eq "System.UInt32") {
                $result.Add([BitConverter]::ToUInt32($bytes, $index))
            } elseif ($typeName -eq "System.Int32") {
                $result.Add([BitConverter]::ToInt32($bytes, $index))
            } elseif ($typeName -eq "System.UInt64") {
                $result.Add([BitConverter]::ToUInt64($bytes, $index))
            } elseif ($typeName -eq "System.Int64") {
                $result.Add([BitConverter]::ToInt64($bytes, $index))
            }
        }

        return $result.ToArray()
    }
}

function Hash([UINT128u]$key) {
    $hash = -1
    for ($i = 0; $i -lt 16; $i++) {
        $index = (($hash -shr 24) -bxor $key.u8[$i]) -band 0xff
        $hash  = (($hash -shl 8) -bxor $crc32_table[$index]) -band 0xFFFFFFFF
    }
    return (-bnot $hash) -band 0x3ff
}
function SetHash {
    param (
        [UINT128u]$key3,
        [ref]$key2,
        [ref]$check
    )

    # Copy $key3 to $key2
    $key2.Value = [UINT128u]::new()
    [Array]::Copy($key3.u8, $key2.Value.u8, 16)

    # Compute the hash and set it in $check
    $check.Value.u8 = [BitConverter]::GetBytes([UINT32](Hash($key2.Value)))

    # Update $key2 with values from $check
    $key2.Value.u8[12] = [Byte]($key2.Value.u8[12] -bor ($check.Value.u8[0] -shl 7))
    $key2.Value.u8[13] = [Byte](($check.Value.u8[0] -shr 1) -bor ($check.Value.u8[1] -shl 7))
    $key2.Value.u8[14] = [Byte]($key2.Value.u8[14] -bor (($check.Value.u8[1] -shr 1) -band 0x1))
}
function SetInfo {
    param (
        [UINT32u]$groupid,
        [UINT32u]$keyid,
        [UINT64u]$secret,
        [ref]$key3
    )

    # Set bytes using groupid
    0..1 | % { $key3.Value.u8[$_] = [BYTE]$groupid.u8[$_] }
    $key3.Value.u8[2] = [BYTE]($key3.Value.u8[2] -bor ($groupid.u8[2] -band 0x0F))

    # Set bytes using keyid
    $key3.Value.u8[2] = [BYTE]($key3.Value.u8[2] -bor ($keyid.u8[0] -shl 4))
    3..5 | % { $key3.Value.u8[$_] = [BYTE](($keyid.u8[$_ - 3 + 1] -shl 4) -bor ($keyid.u8[$_ - 3] -shr 4) -band 0xFF) }
    $key3.Value.u8[6] = [BYTE]($key3.Value.u8[6] -bor (($keyid.u8[3] -shr 4) -band 0x03))

    # Set bytes using secret
    $key3.Value.u8[6] = [BYTE]($key3.Value.u8[6] -bor ($secret.u8[0] -shl 2))
    7..11 | % { $key3.Value.u8[$_] = [BYTE](($secret.u8[$_ - 7 + 1] -shl 2) -bor ($secret.u8[$_ - 7] -shr 6)) }
    $key3.Value.u8[12] = [BYTE](($key3.Value.u8[12] -bor (($secret.u8[6] -shl 2) -bor ($secret.u8[5] -shr 6))) -band 0x7F)
}
function Encode {
    param (
        [UINT128u]$key2,
        [ref]$key1
    )
    $data = 0..3 | % { [BitConverter]::ToUInt32($key2.u8, $_ * 4) }
    for ($i = 25; $i -gt 0; $i--) {
        for ($j = 3; $j -ge 0; $j--) {
            $tmp = if ($j -eq 3) { [UInt64]$data[$j] } else { ([UInt64]$last -shl 32) -bor [UInt64]$data[$j] }
            $data[$j], $last = [math]::Floor($tmp / 24), [UInt32]($tmp % 24)
        }
        $key1.Value[$i - 1] = [byte]$last }
}
function UnconvertChars([byte[]]$key1, [ref]$key0) {
    $n = $key1[0]
    $n += [math]::Floor($n / 5)

    $j = 1
    for ($i = 0; $i -lt 29; $i++) {
        if ($i -eq $n) {
            $key0.Value[$i] = 'N'
        }
        elseif ($i -eq 5 -or $i -eq 11 -or $i -eq 17 -or $i -eq 23) {
            $key0.Value[$i] = '-'
        }
        else {
            switch ($key1[$j++]) {
                0x00 { $key0.Value[$i] = 'B' }
                0x01 { $key0.Value[$i] = 'C' }
                0x02 { $key0.Value[$i] = 'D' }
                0x03 { $key0.Value[$i] = 'F' }
                0x04 { $key0.Value[$i] = 'G' }
                0x05 { $key0.Value[$i] = 'H' }
                0x06 { $key0.Value[$i] = 'J' }
                0x07 { $key0.Value[$i] = 'K' }
                0x08 { $key0.Value[$i] = 'M' }
                0x09 { $key0.Value[$i] = 'P' }
                0x0A { $key0.Value[$i] = 'Q' }
                0x0B { $key0.Value[$i] = 'R' }
                0x0C { $key0.Value[$i] = 'T' }
                0x0D { $key0.Value[$i] = 'V' }
                0x0E { $key0.Value[$i] = 'W' }
                0x0F { $key0.Value[$i] = 'X' }
                0x10 { $key0.Value[$i] = 'Y' }
                0x11 { $key0.Value[$i] = '2' }
                0x12 { $key0.Value[$i] = '3' }
                0x13 { $key0.Value[$i] = '4' }
                0x14 { $key0.Value[$i] = '6' }
                0x15 { $key0.Value[$i] = '7' }
                0x16 { $key0.Value[$i] = '8' }
                0x17 { $key0.Value[$i] = '9' }
                default { $key0.Value[$i] = '?' }
            }
        }
    }
}
function KeyEncode {
    param (
        # 'sgroupid' must be either a hexadecimal (e.g., 0xABC123) or an integer (e.g., 123456)
        [Parameter(Mandatory = $true)]
        [ValidatePattern('^(0x[0-9A-Fa-f]+|\d+)$')]
        [string]$sgroupid,

        [UInt32]$skeyid,
        [UInt64]$sunk
    )
   
    $sgroupid_f = if ($sgroupid -match '^0x') { [Convert]::ToUInt32($sgroupid.Substring(2), 16) } else { [UInt32]$sgroupid }

    if ($sgroupid_f -gt 0xffffff) {
        Write-Host "GroupId must be in the range 0-ffffff"
        return -1
    }
    if ($skeyid -gt 0x3fffffff) {
        Write-Host "KeyId must be in the range 0-3fffffff"
        return -1
    }
    if ($sunk -gt 0x1fffffffffffff) {
        Write-Host "Secret must be in the range 0-1fffffffffffff"
        return -1
    }

    $keyid     = [UINT32u]::new()
    $secret     = [UINT64u]::new()
    $groupid     = [UINT32u]::new()

    $secret.u8  = [BitConverter]::GetBytes($sunk)
    $keyid.u8  = [BitConverter]::GetBytes($skeyid)
    $groupid.u8  = [BitConverter]::GetBytes($sgroupid_f)

    $key3 = [UINT128u]::new()
    SetInfo -groupid $groupid -keyid $keyid -secret $secret -key3 ([ref]$key3)

    $key2 = [UINT128u]::new()
    $check = [UINT32u]::new()
    SetHash -key3 $key3 -key2 ([ref]$key2) -check ([ref]$check)

    $key1 = New-Object Byte[] 25
    Encode -key2 $key2 -key1 ([ref]$key1)

    $key0 = New-Object Char[] 29
    UnconvertChars -key1 $key1 -key0 ([ref]$key0)
   
    return (-join $key0)
}

function Get-Info {
    param (
        [Parameter(Mandatory=$true)]
        [UINT128u]$key3,

        [Parameter(Mandatory=$true)]
        [ref]$groupid,

        [Parameter(Mandatory=$true)]
        [ref]$keyid,

        [Parameter(Mandatory=$true)]
        [ref]$secret
    )

    $groupid.Value.u32 = 0
    $keyid.Value.u32 = 0
    $secret.Value.u64 = 0

    $groupid.Value.u8[0] = $key3.u8[0]
    $groupid.Value.u8[1] = $key3.u8[1]
    $groupid.Value.u8[2] = $key3.u8[2] -band 0x0f

    $keyid.Value.u8[0] = ($key3.u8[2] -shr 4) -bor ($key3.u8[3] -shl 4)
    $keyid.Value.u8[1] = ($key3.u8[3] -shr 4) -bor ($key3.u8[4] -shl 4)
    $keyid.Value.u8[2] = ($key3.u8[4] -shr 4) -bor ($key3.u8[5] -shl 4)
    $keyid.Value.u8[3] = (($key3.u8[5] -shr 4) -bor ($key3.u8[6] -shl 4)) -band 0x3f

    $secret.Value.u8[0] = ($key3.u8[6] -shr 2) -bor ($key3.u8[7] -shl 6)
    $secret.Value.u8[1] = ($key3.u8[7] -shr 2) -bor ($key3.u8[8] -shl 6)
    $secret.Value.u8[2] = ($key3.u8[8] -shr 2) -bor ($key3.u8[9] -shl 6)
    $secret.Value.u8[3] = ($key3.u8[9] -shr 2) -bor ($key3.u8[10] -shl 6)
    $secret.Value.u8[4] = ($key3.u8[10] -shr 2) -bor ($key3.u8[11] -shl 6)
    $secret.Value.u8[5] = ($key3.u8[11] -shr 2) -bor ($key3.u8[12] -shl 6)
    $secret.Value.u8[6] = ($key3.u8[12] -shr 2) -band 0x1f

    $groupid.Value.Sync([SyncSource]::U8)
    $keyid.Value.Sync([SyncSource]::U8)
    $secret.Value.Sync([SyncSource]::U8)

    return $true
}
function Check-Hash {
    param (
        [Parameter(Mandatory=$true)]
        [UINT128u]$key2,

        [Parameter(Mandatory=$true)]
        [ref]$key3,

        [Parameter(Mandatory=$true)]
        [ref]$check
    )

    # Reset the check value
    $check.Value.u32 = 0

    # Copy key2 to key3
    [Array]::Copy($key2.u8, $key3.Value.u8, $key2.u8.Length)

    # Modify key3 bytes
    $key3.Value.u8[12] = $key3.Value.u8[12] -band 0x7f
    $key3.Value.u8[13] = 0
    $key3.Value.u8[14] = $key3.Value.u8[14] -band 0xfe

    # Compute check bytes
    $check.Value.u8[0] = ($key2.u8[13] -shl 1) -bor ($key2.u8[12] -shr 7)
    $check.Value.u8[1] = (($key2.u8[14] -shl 1) -bor ($key2.u8[13] -shr 7)) -band 3

    # Compute hash
    $hash = Hash($key3.Value)
    $key3.Value.Sync([SyncSource]::U8)
    $check.Value.Sync([SyncSource]::U8)

    # Compare hash with check value
    if ($hash -ne $check.Value.u32) {
        Write-Output "Invalid key. The hash is incorrect."
        return $false
    }

    return $true
}
function ConvertTo-UInt32 {
    param (
        [Parameter(Mandatory = $true)]
        [BigInteger]$value
    )

    # Convert BigInteger to uint32 with proper masking
    return [uint32]($value % [BigInteger]0x100000000)
}
function Decode {
    param (
        [Parameter(Mandatory = $true)]
        [byte[]]$key1,

        [ref]$key2
    )

    # Initialize key2
    $key2.Value.u64[0] = 0
    $key2.Value.u64[1] = 0

    for ($ikey = 0; $ikey -lt 25; $ikey++) {
        $res = [BigInteger]24 * [BigInteger]$key2.Value.u32[0] + $key1[$ikey]
        $key2.Value.u32[0] = ConvertTo-UInt32 -value $res
        $res = [BigInteger]($res / [BigInteger]0x100000000)  # Handle overflow

        for ($i = 1; $i -lt 4; $i++) {
            $res += [BigInteger]24 * [BigInteger]$key2.Value.u32[$i]
            $key2.Value.u32[$i] = ConvertTo-UInt32 -value $res
            $res = [BigInteger]($res / [BigInteger]0x100000000)  # Handle overflow
        }
    }

    $key2.Value.Sync([SyncSource]::U32)

    return $true
}
function ConvertChars {
    param (
        [Parameter(Mandatory=$true)]
        [char[]]$key0,

        [ref]$key1
    )

    if ($key0.Length -ne 29) {
        Write-Output "Your key must be 29 characters long."
        return $false
    }

    if ($key0[5] -ne '-' -or $key0[11] -ne '-' -or $key0[17] -ne '-' -or $key0[23] -ne '-') {
        Write-Output "Incorrect hyphens."
        return $false
    }

    if ($key0[28] -eq 'N') {
        Write-Output "The last character must not be an N."
        return $false
    }

    $n = $false
    $j = 1
    $i = 0

    while ($j -lt 25 -and $i -lt $key0.Length) {
        switch ($key0[$i++]) {
            'N' {
                if ($n) {
                    throw "There may only be one N in a key."
                    return $false
                }
                $n = $true
                $key1.Value[0] = $j - 1
            }
            'B' { if ($j -lt 25) { $key1.Value[$j++] = 0x00 } }
            'C' { if ($j -lt 25) { $key1.Value[$j++] = 0x01 } }
            'D' { if ($j -lt 25) { $key1.Value[$j++] = 0x02 } }
            'F' { if ($j -lt 25) { $key1.Value[$j++] = 0x03 } }
            'G' { if ($j -lt 25) { $key1.Value[$j++] = 0x04 } }
            'H' { if ($j -lt 25) { $key1.Value[$j++] = 0x05 } }
            'J' { if ($j -lt 25) { $key1.Value[$j++] = 0x06 } }
            'K' { if ($j -lt 25) { $key1.Value[$j++] = 0x07 } }
            'M' { if ($j -lt 25) { $key1.Value[$j++] = 0x08 } }
            'P' { if ($j -lt 25) { $key1.Value[$j++] = 0x09 } }
            'Q' { if ($j -lt 25) { $key1.Value[$j++] = 0x0a } }
            'R' { if ($j -lt 25) { $key1.Value[$j++] = 0x0b } }
            'T' { if ($j -lt 25) { $key1.Value[$j++] = 0x0c } }
            'V' { if ($j -lt 25) { $key1.Value[$j++] = 0x0d } }
            'W' { if ($j -lt 25) { $key1.Value[$j++] = 0x0e } }
            'X' { if ($j -lt 25) { $key1.Value[$j++] = 0x0f } }
            'Y' { if ($j -lt 25) { $key1.Value[$j++] = 0x10 } }
            '2' { if ($j -lt 25) { $key1.Value[$j++] = 0x11 } }
            '3' { if ($j -lt 25) { $key1.Value[$j++] = 0x12 } }
            '4' { if ($j -lt 25) { $key1.Value[$j++] = 0x13 } }
            '6' { if ($j -lt 25) { $key1.Value[$j++] = 0x14 } }
            '7' { if ($j -lt 25) { $key1.Value[$j++] = 0x15 } }
            '8' { if ($j -lt 25) { $key1.Value[$j++] = 0x16 } }
            '9' { if ($j -lt 25) { $key1.Value[$j++] = 0x17 } }
            '-' { }
            default {
                throw "Invalid character in key."
                return $false
            }
        }
    }

    if (-not $n) {
        throw "The character N must be in the product key."
        return $false
    }

    return $true
}
function KeyDecode {
    param (
        [Parameter(Mandatory=$true)]
        [string]$key0
    )

    # Convert the string to a character array
    $key0Chars = $key0.ToCharArray()

    # Initialize $key1 array
    $key1 = New-Object byte[] 25
    
    # Convert characters to bytes
    if (-not (ConvertChars -key0 $key0Chars -key1 ([ref]$key1))) {
        return -1
    }
    
    # Initialize UINT128u structures
    $key2 = [UINT128u]::new()
    $key3 = [UINT128u]::new()
    $hash = [UINT32u]::new()
    
    # Decode the key
    if (-not (Decode -key1 $key1 -key2 ([ref]$key2))) {
        return -1
    }

    # Check the hash
    if (-not (Check-Hash -key2 $key2 -key3 ([ref]$key3) -check ([ref]$hash))) {
        return -1
    }
    
    # Initialize UINT32u and UINT64u structures
    $groupid = [UINT32u]::new()
    $keyid = [UINT32u]::new()
    $secret = [UINT64u]::new()
    
    # Get information
    if (-not (Get-Info -key3 $key3 -groupid ([ref]$groupid) -keyid ([ref]$keyid) -secret ([ref]$secret))) {
        return -1
    }
    
    return @(
    @{ Property = "KeyId";   Value = $keyid.u32 },
    @{ Property = "Hash";    Value = $hash.u32 },
    @{ Property = "GroupId"; Value = $groupid.u32 },
    @{ Property = "Secret";  Value = $secret.u64}
    )
}

# Adaption of "Licensing Stuff" from =awuctl=, "KeyInfo" from Bob65536
# https://github.com/awuctl/licensing-stuff/blob/main/keycutter.py
# https://forums.mydigitallife.net/threads/how-get-oem-key-system-key.87962/#post-1825092
# https://forums.mydigitallife.net/threads/we8industry-pro-wes8-activation.45312/#post-771802
# https://web.archive.org/web/20121026081005/http://forums.mydigitallife.info/threads/37590-Windows-8-Product-Key-Decoding

function Encode-Key {
    param(
        [Parameter(Mandatory=$true)]
        [UInt64]$group,

        [Parameter(Mandatory=$false)]
        [UInt64]$serial = 0,

        [Parameter(Mandatory=$false)]
        [UInt64]$security = 0,

        [Parameter(Mandatory=$false)]
        [int]$upgrade = 0,

        [Parameter(Mandatory=$false)]
        [int]$extra = 0,

        [Parameter(Mandatory=$false)]
        [int]$checksum = -1
    )

    # Alphabet used for encoding base24 digits (excluding 'N')
    $ALPHABET = 'BCDFGHJKMPQRTVWXY2346789'.ToCharArray()

    # Validate input ranges (equivalent to Python BOUNDS)
    if ($group -gt 0xFFFFF) {
        throw "Group value ($group) out of bounds (max 0xFFFFF)"
    }
    if ($serial -gt 0x3FFFFFFF) {
        throw "Serial value ($serial) out of bounds (max 0x3FFFFFFF)"
    }
    if ($security -gt 0x1FFFFFFFFFFFFF) {
        throw "Security value ($security) out of bounds (max 0x1FFFFFFFFFFFFF)"
    }
    if ($checksum -ne -1 -and $checksum -gt 0x3FF) {
        throw "Checksum value ($checksum) out of bounds (max 0x3FF)"
    }
    if ($upgrade -notin @(0, 1)) {
        throw "Upgrade value must be either 0 or 1"
    }
    if ($extra -notin @(0, 1)) {
        throw "Extra value must be either 0 or 1"
    }

    function Get-Checksum {
        param([byte[]]$data)

        [uint32]$crc = [uint32]::MaxValue
        foreach ($b in $data) {
            $index = (($crc -shr 24) -bxor $b) -band 0xFF
            $crc = ((($crc -shl 8) -bxor $crc32_table[$index]) -band 0xFFFFFFFF)
        }
        $crc = (-bnot $crc) -band 0xFFFFFFFF
        return $crc -band 0x3FF  # 10 bits checksum mask
    }
    function Encode-Base24 {
        param([System.Numerics.BigInteger]$num)
        $digits = New-Object byte[] 25
        for ($i = 24; $i -ge 0; $i--) {
            $digits[$i] = [byte]($num % 24)
            $num = [System.Numerics.BigInteger]::Divide($num, 24)
        }
        return $digits
    }
    function Format-5x5 {
        param([byte[]]$digits)

        # Calculate position for inserting 'N'
        $pos = $digits[0] #+ [math]::Floor($digits[0] / 5)

        $ALPHABET = @('B','C','D','F','G','H','J','K','M','P','Q','R','T','V','W','X','Y','2','3','4','6','7','8','9')

        $chars = @()
        for ($i = 1; $i -lt 25; $i++) {
            $chars += $ALPHABET[$digits[$i]]
        }

        # Insert 'N' at the calculated position
        if ($pos -le 0) {
            $chars = @('N') + $chars
        }
        elseif ($pos -ge $chars.Count) {
            $chars += 'N'
        }
        else {
            $chars = $chars[0..($pos - 1)] + 'N' + $chars[$pos..($chars.Count - 1)]
        }

        # Insert dashes every 5 characters to form groups
        return -join (
            ($chars[0..4] -join ''), '-',
            ($chars[5..9] -join ''), '-',
            ($chars[10..14] -join ''), '-',
            ($chars[15..19] -join ''), '-',
            ($chars[20..24] -join '')
        )
    }

    # Validate input ranges to avoid overflow
    if ($group -gt 0xFFFFF -or $serial -gt 0x3FFFFFFF -or $security -gt 0x1FFFFFFFFFFFFF) {
        throw "Field values out of range"
    }

    # Compose the key bits using BigInteger (64+ bit shifts)
    $key = [System.Numerics.BigInteger]::Zero
    $key = $key -bor ([System.Numerics.BigInteger]$extra -shl 114)
    $key = $key -bor ([System.Numerics.BigInteger]$upgrade -shl 113)
    $key = $key -bor ([System.Numerics.BigInteger]$security -shl 50)
    $key = $key -bor ([System.Numerics.BigInteger]$serial -shl 20)
    $key = $key -bor ([System.Numerics.BigInteger]$group)

    # Calculate checksum if not provided
    if ($checksum -lt 0) {
        $keyBytes = $key.ToByteArray()

        # Remove extra sign byte if present (BigInteger uses signed representation)
        if ($keyBytes.Length -gt 16) {
            if ($keyBytes[-1] -eq 0x00) {
                # Remove the last byte (sign byte)
                $keyBytes = $keyBytes[0..($keyBytes.Length - 2)]
            }
            else {
                throw "Key bytes length greater than 16 with unexpected data"
            }
        }

        # Pad with trailing zeros to get exactly 16 bytes (little-endian)
        if ($keyBytes.Length -lt 16) {
            $keyBytes += ,0 * (16 - $keyBytes.Length)
        }

        # No reversal needed ? checksum function expects little-endian bytes
        # [array]::Reverse($keyBytes)  # <-- removed

        $checksum = Get-Checksum $keyBytes
    }

    # Insert checksum bits at bit position 103
    $key = $key -bor ([System.Numerics.BigInteger]$checksum -shl 103)

    # Encode the final key to base24 digits
    $base24 = Encode-Base24 $key

    # Format into the 5x5 grouped string with 'N' insertion and dashes
    return Format-5x5 $base24
}
function Decode-Key {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Key
    )

    $ALPHABET = 'BCDFGHJKMPQRTVWXY2346789'.ToCharArray()

    # Remove hyphens and uppercase
    $k = $Key.Replace('-','').ToUpper()

    # Find 'N' position
    $ni = $k.IndexOf('N')
    if ($ni -lt 0) { throw "Invalid key (missing 'N')." }

    # Start digits array with position of 'N'
    $digits = @($ni)

    # Remove 'N' from key
    $rest = $k.Replace('N','')

    # Convert each character to index in alphabet
    foreach ($ch in $rest.ToCharArray()) {
        $idx = $Alphabet.IndexOf($ch)
        if ($idx -lt 0) { throw "Invalid character '$ch' in key." }
        $digits += $idx
    }

    [bigint]$value = 0
    foreach ($d in $digits) {
        $value = ($value * 24) + $d
    }

    # Extract bit fields
    $group    = [int]($value -band 0xfffff)                         # 20 bits decimal
    $serial   = [int](($value -shr 20) -band 0x3fffffff)             # 30 bits decimal
    $security = [bigint](($value -shr 50) -band 0x1fffffffffffff)    # 53 bits decimal (bigint)
    $checksum = [int](($value -shr 103) -band 0x3ff)                  # 10 bits decimal
    $upgrade  = [int](($value -shr 113) -band 0x1)                    # 1 bit decimal
    $extra    = [int](($value -shr 114) -band 0x1)                    # 1 bit decimal

    return [pscustomobject]@{
        Key      = $Key
        Integer  = $value
        Group    = $group
        Serial   = $serial
        Security = $security
        Checksum = $checksum
        Upgrade  = $upgrade
        Extra    = $extra
    }
}

<#
Generates product keys using:

1. Template mode:
   - Use -template (e.g. "NBBBB-BBBBB-") to find matching keys.
   - Stops when keys no longer match the given prefix.

2. Brute-force mode:
   - If no -template is given, generates up to -MaxTries.
   - Collects keys starting with "NBBBB".
   - If no valid keys are found, automatically retries up to 5 times,
     each time increasing the MaxTries limit by 5000.
     Stops once at least one key is found or the retry limit is reached.

Based on abbodi1406's logic:
https://forums.mydigitallife.net/threads/88595/page-6#post-1882091

Examples:
    Brute-force mode
    List-Keys -RefGroupId 2048
    List-Keys -RefGroupId 2048 -OffsetLimit 200000
    List-Keys -RefGroupId 2048 -StartAtOffset 120000 -KeysLimit 2
    List-Keys -RefGroupId 2048 -StartAtOffset 120000 -OffsetLimit 20000

    Template mode
    List-Keys -RefGroupId 2077 -template JDHD7-DHN6R-JDHD7
    List-Keys -RefGroupId 2048 -template NBBBB-BBBBB-BBBBB -KeysLimit 2
    List-Keys -RefGroupId 2048 -template NBBBB-BBBBB-BBBBB -OffsetLimit 20000
#>
function List-Keys {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateScript( { $_ -gt 0 } )]
        [int32]$RefGroupId,

        [Parameter(Mandatory=$false)]
        [string]$Template,

        [Parameter(Mandatory=$false)]
        [ValidateScript( { $_ -gt 0 } )]
        [int32]$OffsetLimit = 10000,

        [Parameter(Mandatory=$false)]
        [ValidateScript( { $_ -ge 0 } )]
        [int32]$KeysLimit = 0,

        [Parameter(Mandatory=$false)]
        [ValidateScript( { $_ -ge 0 } )]
        [int32]$StartAtOffset = 0
    )

    if ($template) {
        if ($template.Length -gt 21) {
            throw "Template too long"
        }        
        $paddingTemplate = 'NBBBB-BBBBB-BBBBB-BBBBB-BBBBB'

        # Pad the template to full length with padding string starting from template length
        $paddedTemplate = $template + $paddingTemplate.Substring($template.Length)

        # Decode the padded template into components
        $templateKey = Decode-Key -Key $paddedTemplate
        $serialIter = $templateKey.Serial
        $serialEdgeOffset = $serialIter + $OffsetLimit

    } 
    if (-not $template) {
        $serialIter = $StartAtOffset
    }

    # Initialize variables
    $keyArray = @()
    $attemptCount = 0

    while ($true) {
        if ($template) {   
            # ----- > Begin
            if ($serialIter -ge $serialEdgeOffset) {
                return $keyArray
            }
            $key = Encode-Key -group $RefGroupId -serial $serialIter -security $templateKey.Security -upgrade $templateKey.Upgrade -extra $templateKey.Extra
            $decodedKey = Decode-Key -Key $key
            if ($decodedKey.Checksum -ne $templateKey.Checksum) {
                $serialIter++
                continue
            }
            if (($key.Substring(0, $template.Length) -ne $template)) {
                break
            }
            # ----- > End
        }

        if (-not $template) {
            # ----- > Begin
            if ($serialIter -ge ($OffsetLimit+$StartAtOffset)) {
                if ($keyArray.Count -gt 0) {
                    return $keyArray
                }
                elseif ($attemptCount -lt 5) {
                    $OffsetLimit += 5000
                    $attemptCount++
                    continue 
                } else {
                    return $keyArray
                }
            }
            $key = Encode-Key -group $RefGroupId -serial $serialIter -security 0
            if ($key -notmatch "NBBBB") {
                $serialIter++
                continue 
            }
            # ----- > End
        }

        $keyArray += $key
        Write-Warning "StartAtOffset: $serialIter, Key: $key"

        # Check if MaxKeys is set and if we've reached the limit
        if ($KeysLimit -gt 0 -and $keyArray.Count -ge $KeysLimit) {
            return $keyArray
        }

        $serialIter++
    }

    # Just in case.
    # Should not arrived here.
    return $keyArray
}
# KeyInfo Part -->

# KMS Part -->
# define $Global's
$Global:LocalKms = $false
$Global:ConvertedMode = $true
$Global:Windows_7_Or_Earlier = $false
$Global:SupportedBuildYear = @('2019', '2021', '2024')
$Global:IP_ADDRESS = "192.168.$(Get-Random -Minimum 1 -Maximum 256).$(Get-Random -Minimum 1 -Maximum 256)"

# define Csv Data
$Global:Windows_Keys_List = @'
ID, KEY
00091344-1ea4-4f37-b789-01750ba6988c,W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9
01ef176b-3e0d-422a-b4f8-4ea880035e8f,4DWFP-JF3DJ-B7DTH-78FJB-PDRHK
034d3cbb-5d4b-4245-b3f8-f84571314078,WVDHN-86M7X-466P6-VHXV7-YY726
096ce63d-4fac-48a9-82a9-61ae9e800e5f,789NJ-TQK6T-6XTH8-J39CJ-J8D3P
0ab82d54-47f4-4acb-818c-cc5bf0ecb649,NMMPB-38DD4-R2823-62W8D-VXKJB
0df4f814-3f57-4b8b-9a9d-fddadcd69fac,NBTWJ-3DR69-3C4V8-C26MC-GQ9M6
10018baf-ce21-4060-80bd-47fe74ed4dab,RYXVT-BNQG7-VD29F-DBMRY-HT73M
113e705c-fa49-48a4-beea-7dd879b46b14,TT4HM-HN7YT-62K67-RGRQJ-JFFXW
18db1848-12e0-4167-b9d7-da7fcda507db,NKB3R-R2F8T-3XCDP-7Q2KW-XWYQ2
197390a0-65f6-4a95-bdc4-55d58a3b0253,8N2M2-HWPGY-7PGT9-HGDD8-GVGGY
1cb6d605-11b3-4e14-bb30-da91c8e3983a,YDRBP-3D83W-TY26F-D46B2-XCKRJ
21c56779-b449-4d20-adfc-eece0e1ad74b,CB7KF-BWN84-R7R2Y-793K2-8XDDG
21db6ba4-9a7b-4a14-9e29-64a60c59301d,KNC87-3J2TX-XB4WP-VCPJV-M4FWM
2401e3d0-c50a-4b58-87b2-7e794b7d2607,W7VD6-7JFBR-RX26B-YKQ3Y-6FFFJ
2b5a1b0f-a5ab-4c54-ac2f-a6d94824a283,JCKRF-N37P4-C2D82-9YXRT-4M63B
2c682dc2-8b68-4f63-a165-ae291d4cf138,HMBQG-8H2RH-C77VX-27R82-VMQBT
2d5a5a60-3040-48bf-beb0-fcd770c20ce0,DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ
2de67392-b7a7-462a-b1ca-108dd189f588,W269N-WFGWX-YVC9B-4J6C9-T83GX
32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee,M7XTQ-FN8P6-TTKYV-9D4CC-J462D
34e1ae55-27f8-4950-8877-7a03be5fb181,WMDGN-G9PQG-XVVXX-R3X43-63DFG
3c102355-d027-42c6-ad23-2e7ef8a02585,2WH4N-8QGBV-H22JP-CT43Q-MDWWJ
3dbf341b-5f6c-4fa7-b936-699dce9e263f,VP34G-4NPPG-79JTQ-864T4-R3MQX
3f1afc82-f8ac-4f6c-8005-1d233e606eee,6TP4R-GNPTD-KYYHQ-7B7DP-J447Y
43d9af6e-5e86-4be8-a797-d072a046896c,K9FYF-G6NCK-73M32-XMVPY-F9DRR
458e1bec-837a-45f6-b9d5-925ed5d299de,32JNW-9KQ84-P47T8-D8GGY-CWCK7
46bbed08-9c7b-48fc-a614-95250573f4ea,C29WB-22CC8-VJ326-GHFJW-H9DH4
4b1571d3-bafb-4b40-8087-a961be2caf65,9FNHH-K3HBT-3W4TD-6383H-6XYWF
4f3d1606-3fea-4c01-be3c-8d671c401e3b,YFKBB-PQJJV-G996G-VWGXY-2V3X8
5300b18c-2e33-4dc2-8291-47ffcec746dd,YVWGF-BXNMC-HTQYQ-CPQ99-66QFC
54a09a0d-d57b-4c10-8b69-a842d6590ad5,MRPKT-YTG23-K7D7T-X2JMM-QY7MG
58e97c99-f377-4ef1-81d5-4ad5522b5fd8,TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
59eb965c-9150-42b7-a0ec-22151b9897c5,KBN8V-HFGQ4-MGXVD-347P6-PDQGT
59eb965c-9150-42b7-a0ec-22151b9897c5,KBN8V-HFGQ4-MGXVD-347P6-PDQGT
5a041529-fef8-4d07-b06f-b59b573b32d2,W82YF-2Q76Y-63HXB-FGJG9-GF7QX
61c5ef22-f14f-4553-a824-c4b31e84b100,PTXN8-JFHJM-4WC78-MPCBR-9W4KR
620e2b3d-09e7-42fd-802a-17a13652fe7a,489J6-VHDMP-X63PK-3K798-CPX3Y
68531fb9-5511-4989-97be-d11a0f55633f,YC6KT-GKW9T-YTKYR-T4X34-R7VHC
68b6e220-cf09-466b-92d3-45cd964b9509,7M67G-PC374-GR742-YH8V4-TCBY3
7103a333-b8c8-49cc-93ce-d37c09687f92,92NFX-8DJQP-P6BBQ-THF9C-7CG2H
73111121-5638-40f6-bc11-f1d7b0d64300,NPPR9-FWDCX-D2C8J-H872K-2YT43
73e3957c-fc0c-400d-9184-5f7b6f2eb409,N2KJX-J94YW-TQVFB-DG9YT-724CC
7476d79f-8e48-49b4-ab63-4d0b813a16e4,HMCNV-VVBFX-7HMBH-CTY9B-B4FXY
7482e61b-c589-4b7f-8ecc-46d455ac3b87,74YFP-3QFB3-KQT8W-PMXWJ-7M648
78558a64-dc19-43fe-a0d0-8075b2a370a3,7B9N3-D94CG-YTVHR-QBPX3-RJP64
7afb1156-2c1d-40fc-b260-aab7442b62fe,RCTX3-KWVHP-BR6TB-RB6DM-6X7HP
7b4433f4-b1e7-4788-895a-c45378d38253,QN4C6-GBJD2-FB422-GHWJK-GJG2R
7b51a46c-0c04-4e8f-9af4-8496cca90d5e,WNMTR-4C88C-JK8YV-HQ7T2-76DF9
7b9e1751-a8da-4f75-9560-5fadfe3d8e38,3KHY7-WNT83-DGQKR-F7HPR-844BM
7d5486c7-e120-4771-b7f1-7b56c6d3170c,HM7DN-YVMH3-46JC3-XYTG7-CYQJJ
81671aaf-79d1-4eb1-b004-8cbbe173afea,MHF9N-XY6XB-WVXMC-BTDCT-MKKG7
8198490a-add0-47b2-b3ba-316b12d647b4,39BXF-X8Q23-P2WWT-38T2F-G3FPG
82bbc092-bc50-4e16-8e18-b74fc486aec3,NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J
87b838b7-41b6-4590-8318-5797951d8529,2F77B-TNFGY-69QQF-B8YKP-D69TJ
8860fcd4-a77b-4a20-9045-a150ff11d609,2WN2H-YGCQR-KFX6K-CD6TF-84YXQ
8a26851c-1c7e-48d3-a687-fbca9b9ac16b,GT63C-RJFQ3-4GMB6-BRFB9-CB83V
8c1c5410-9f39-4805-8c9d-63a07706358f,WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY
8c8f0ad3-9a43-4e05-b840-93b8d1475cbc,6N379-GGTMK-23C6M-XVVTC-CKFRQ
8de8eb62-bbe0-40ac-ac17-f75595071ea3,GRFBW-QNDC4-6QBHG-CCK3B-2PR88
90c362e5-0da1-4bfd-b53b-b87d309ade43,6NMRW-2C8FM-D24W7-TQWMY-CWH2D
95fd1c83-7df5-494a-be8b-1300e1c9d1cd,XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G
9bd77860-9b31-4b7b-96ad-2564017315bf,VDYBN-27WPP-V4HQT-9VMD4-VMK7H
9d5584a2-2d85-419a-982c-a00888bb9ddf,4K36P-JN4VD-GDC6V-KDT89-DYFKP
9f776d83-7156-45b2-8a5c-359b9c9f22a3,QFFDN-GRT3P-VKWWX-X7T3R-8B639
a00018a3-f20f-4632-bf7c-8daa5351c914,GNBB8-YVD74-QJHX6-27H4K-8QHDG
a78b8bd9-8017-4df5-b86a-09f756affa7c,6TPJF-RBVHG-WBW2R-86QPH-6RTM4
a80b5abf-76ad-428b-b05d-a47d2dffeebf,MH37W-N47XK-V7XM9-C7227-GCQG9
a9107544-f4a0-4053-a96a-1479abdef912,PVMJN-6DFY6-9CCP6-7BKTT-D3WVR
a98bcd6d-5343-4603-8afe-5908e4611112,NG4HW-VH26C-733KW-K6F98-J8CK4
a99cc1f0-7719-4306-9645-294102fbff95,FDNH6-VW9RW-BXPJ7-4XTYG-239TB
aa6dd3aa-c2b4-40e2-a544-a6bbb3f5c395,73KQT-CD9G6-K7TQG-66MRP-CQ22C
ad2542d4-9154-4c6d-8a44-30f11ee96989,TM24T-X9RMF-VWXK6-X8JC9-BFGM2
ae2ee509-1b34-41c0-acb7-6d4650168915,33PXH-7Y6KF-2VJC9-XBBR8-HVTHH
af35d7b7-5035-4b63-8972-f0b747b9f4dc,DXHJF-N9KQX-MFPVR-GHGQK-Y7RKV
b3ca044e-a358-4d68-9883-aaa2941aca99,D2N9P-3P6X9-2R39C-7RTCD-MDVJX
b743a2be-68d4-4dd3-af32-92425b7bb623,3NPTF-33KPT-GGBPR-YX76B-39KDD
b8f5e3a3-ed33-4608-81e1-37d6c9dcfd9c,KF37N-VDV38-GRRTV-XH8X6-6F3BB
b92e9980-b9d5-4821-9c94-140f632f6312,FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4
ba998212-460a-44db-bfb5-71bf09d1c68b,R962J-37N87-9VVK2-WJ74P-XTMHR
c04ed6bf-55c8-4b47-9f8e-5a1f31ceee60,BN3D2-R7TKB-3YPBD-8DRP2-27GG4
c06b6981-d7fd-4a35-b7b4-054742b7af67,GCRJD-8NW9H-F2CDX-CCM8D-9D6T9
c1af4d90-d1bc-44ca-85d4-003ba33db3b9,YQGMW-MPWTJ-34KDK-48M3W-X4Q6V
c6ddecd6-2354-4c19-909b-306a3058484e,Q6HTR-N24GM-PMJFP-69CD8-2GXKR
c72c6a1d-f252-4e7e-bdd1-3fca342acb35,BB6NG-PQ82V-VRDPW-8XVD2-V8P66
ca7df2e3-5ea0-47b8-9ac1-b1be4d8edd69,37D7F-N49CB-WQR8W-TBJ73-FM8RX
ca7df2e3-5ea0-47b8-9ac1-b1be4d8edd69,37D7F-N49CB-WQR8W-TBJ73-FM8RX
cab491c7-a918-4f60-b502-dab75e334f40,TNFGH-2R6PB-8XM3K-QYHX2-J4296
cd4e2d9f-5059-4a50-a92d-05d5bb1267c7,FNFKF-PWTVT-9RC8H-32HB2-JB34X
cd918a57-a41b-4c82-8dce-1a538e221a83,7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH
cda18cf3-c196-46ad-b289-60c072869994,TT8MH-CG224-D3D7Q-498W2-9QCTX
cfd8ff08-c0d7-452b-9f60-ef5c70c32094,VKK3X-68KWM-X2YGT-QR4M6-4BWMV
d30136fc-cb4b-416e-a23d-87207abc44a9,6XN7V-PCBDC-BDBRH-8DQY7-G6R44
d3643d60-0c42-412d-a7d6-52e6635327f6,48HP8-DN98B-MYWDG-T2DCC-8W83P
d4f54950-26f2-4fb4-ba21-ffab16afcade,VTC42-BM838-43QHV-84HX6-XJXKV
db537896-376f-48ae-a492-53d0547773d0,YBYF6-BHCR3-JPKRB-CDW7B-F9BK4
db78b74f-ef1c-4892-abfe-1e66b8231df6,NCTT7-2RGK8-WMHRF-RY7YQ-JTXG3
ddfa9f7c-f09e-40b9-8c1a-be877a9a7f4b,WYR28-R7TFJ-3X2YQ-YCY4H-M249D
de32eafd-aaee-4662-9444-c1befb41bde2,N69G4-B89J2-4G8F4-WWYCC-J464C
e0b2d383-d112-413f-8a80-97f373a5820c,YYVX9-NTFWV-6MDM3-9PT4T-4M68B
e0c42288-980c-4788-a014-c080d2e1926e,NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
e14997e7-800a-4cf7-ad10-de4b45b578db,JMNMF-RHW7P-DMY6X-RF3DR-X2BQT
e1a8296a-db37-44d1-8cce-7bc961d59c54,XGY72-BRBBT-FF8MH-2GG8H-W7KCW
e272e3e2-732f-4c65-a8f0-484747d0d947,DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4
e38454fb-41a4-4f59-a5dc-25080e354730,44RPN-FTY23-9VTTB-MP9BX-T84FV
e49c08e7-da82-42f8-bde2-b570fbcae76c,2HXDN-KRXHB-GPYC7-YCKFJ-7FVDG
e4db50ea-bda1-4566-b047-0ca50abc6f07,7NBT4-WGBQX-MP4H7-QXFF8-YP3KX
e58d87b5-8126-4580-80fb-861b22f79296,MX3RK-9HNGX-K3QKC-6PJ3F-W8D7B
e9942b32-2e55-4197-b0bd-5ff58cba8860,3PY8R-QHNP9-W7XQD-G6DPH-3J2C9
ebf245c1-29a8-4daf-9cb1-38dfc608a8c8,XCVCF-2NXM9-723PB-MHCB7-2RYQQ
ec868e65-fadf-4759-b23e-93fe37f2cc29,CPWHC-NT2C7-VYW78-DHDB2-PG3GK
ef6cfc9f-8c5d-44ac-9aad-de6a2ea0ae03,WX4NM-KYWYW-QJJR4-XV3QB-6VM33
f0f5ec41-0d55-4732-af02-440a44a3cf0f,XC9B7-NBPP2-83J2H-RHMBY-92BT4
f772515c-0e87-48d5-a676-e6962c3e1195,736RG-XDKJK-V34PF-BHK87-J6X3K
f7e88590-dfc7-4c78-bccb-6f3865b99d1a,VHXM3-NR6FT-RY6RT-CK882-KW2CJ
fd09ef77-5647-4eff-809c-af2b64659a45,22XQ2-VRXRG-P8D42-K34TD-G3QQC
fe1c3238-432a-43a1-8e25-97e7d1ef10f3,M9Q9P-WNJJT-6PXPY-DWX8H-6XWKK
ffee456a-cd87-4390-8e07-16146c672fd0,XYTND-K6QKT-K2MRH-66RTM-43JKP
7dc26449-db21-4e09-ba37-28f2958506a6,TVRH6-WHNXV-R9WG3-9XRFY-MY832
c052f164-cdf6-409a-a0cb-853ba0f0f55a,D764K-2NDRG-47T6Q-P8T8W-YP6DF
45b5aff2-60a0-42f2-bc4b-ec6e5f7b527e,FCNV3-279Q9-BQB46-FTKXX-9HPRH
c2e946d1-cfa2-4523-8c87-30bc696ee584,XGN3F-F394H-FD2MY-PP6FD-8MCRC
f57b5b6b-80c2-46e4-ae9d-9fe98e032cb7,GFMWN-WDHVB-4Y4XP-42WKM-RC6CQ
b1b1ef19-a088-4962-aedb-2a647a891104,XN3XP-QGKM4-KT7HM-6HC6T-H8V6F
1a716f14-0607-425f-a097-5f2f1f091315,QCQ4R-N2J93-PWMTK-G2BGF-BY82T
8f365ba6-c1b9-4223-98fc-282a0756a3ed,HTDQM-NBMMG-KGYDT-2DTKT-J2MPV
'@ | ConvertFrom-Csv
$Global:Office_Keys_List = @'
Product,Year,Key
Excel,2010,H62QG-HXVKF-PP4HP-66KMR-CW9BM
Excel,2013,VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB
Excel,2016,9C2PK-NWTVB-JMPW8-BFT28-7FTBF
Excel,2019,TMJWT-YYNMB-3BKTF-644FC-RVXBD
Excel,2021,NWG3X-87C9K-TC7YY-BC2G7-G6RVC
Excel,2024,F4DYN-89BP2-WQTWJ-GR8YC-CKGJG
PowerPoint,2010,RC8FX-88JRY-3PF7C-X8P67-P4VTT
PowerPoint,2013,4NT99-8RJFH-Q2VDH-KYG2C-4RD4F
PowerPoint,2016,J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6
PowerPoint,2019,RRNCX-C64HY-W2MM7-MCH9G-TJHMQ
PowerPoint,2021,TY7XF-NFRBR-KJ44C-G83KF-GX27K
PowerPoint,2024,CW94N-K6GJH-9CTXY-MG2VC-FYCWP
ProPlus,2010,VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB
ProPlus,2013,YC7DK-G2NP3-2QQC3-J6H88-GVGXT
ProPlus,2016,XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
ProPlus,2019,NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
ProPlus,2021,FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
ProPlus,2024,XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB
ProjectPro,2010,YGX6F-PGV49-PGW3J-9BTGG-VHKC6
ProjectPro,2013,FN8TT-7WMH6-2D4X9-M337T-2342K
ProjectPro,2016,YG9NW-3K39V-2T3HJ-93F3Q-G83KT
ProjectPro,2019,B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
ProjectPro,2021,FTNWT-C6WBT-8HMGF-K9PRX-QV9H8
ProjectPro,2024,FQQ23-N4YCY-73HQ3-FM9WC-76HF4
ProjectStd,2010,4HP3K-88W3F-W2K3D-6677X-F9PGB
ProjectStd,2013,6NTH3-CW976-3G3Y2-JK3TX-8QHTT
ProjectStd,2016,GNFHQ-F6YQM-KQDGJ-327XX-KQBVC
ProjectStd,2019,C4F7P-NCP8C-6CQPT-MQHV9-JXD2M
ProjectStd,2021,J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T
ProjectStd,2024,PD3TT-NTHQQ-VC7CY-MFXK3-G87F8
Publisher,2010,BFK7F-9MYHM-V68C7-DRQ66-83YTP
Publisher,2013,PN2WF-29XG2-T9HJ7-JQPJR-FCXK4
Publisher,2016,F47MM-N3XJP-TQXJ9-BP99D-8K837
Publisher,2019,G2KWX-3NW6P-PY93R-JXK2T-C9Y9V
Publisher,2021,2MW9D-N4BXM-9VBPG-Q7W6M-KFBGQ
SkypeforBusiness,2016,869NQ-FJ69K-466HW-QYCP2-DDBV6
SkypeforBusiness,2019,NCJ33-JHBBY-HTK98-MYCV8-HMKHJ
SkypeforBusiness,2021,HWCXN-K3WBT-WJBKY-R8BD9-XK29P
SkypeforBusiness,2024,4NKHF-9HBQF-Q3B6C-7YV34-F64P3
SmallBusBasics,2010,D6QFG-VBYP2-XQHM7-J97RH-VVRCK
Standard,2010,V7QKV-4XVVR-XYV4D-F7DFM-8R6BM
Standard,2013,KBKQT-2NMXY-JJWGP-M62JB-92CD4
Standard,2016,JNRGM-WHDWX-FJJG3-K47QV-DRTFM
Standard,2019,6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK
Standard,2021,KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3
Standard,2024,V28N4-JG22K-W66P8-VTMGK-H6HGR
VisioPrem,2010,D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ
VisioPro,2010,D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ
VisioPro,2013,C2FG9-N6J68-H8BTJ-BW3QX-RM3B3
VisioPro,2016,PD3PC-RHNGV-FXJ29-8JK7D-RJRJK
VisioPro,2019,9BGNQ-K37YR-RQHF2-38RQ3-7VCBB
VisioPro,2021,KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4
VisioPro,2024,B7TN8-FJ8V3-7QYCP-HQPMV-YY89G
VisioStd,2010,767HD-QGMWX-8QTDB-9G3R2-KHFGJ
VisioStd,2013,J484Y-4NKBF-W2HMG-DBMJC-PGWR7
VisioStd,2016,7WHWN-4T7MP-G96JF-G33KR-W8GF4
VisioStd,2019,7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2
VisioStd,2021,MJVNY-BYWPY-CWV6J-2RKRT-4M8QG
VisioStd,2024,JMMVY-XFNQC-KK4HK-9H7R3-WQQTV
Word,2010,HVHB3-C6FV7-KQX9W-YQG79-CRY7T
Word,2013,6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7
Word,2016,WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6
access,2010,V7Y44-9T38C-R2VJK-666HK-T7DDX
access,2013,NG2JY-H4JBT-HQXYP-78QH9-4JM2D
access,2016,GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW
access,2019,9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT
access,2021,WM8YG-YNGDD-4JHDC-PG3F4-FC4T4
access,2024,82FTR-NCHR7-W3944-MGRHM-JMCWD
mondo,2010,7TC2V-WXF6P-TD7RT-BQRXR-B8K32
mondo,2013,42QTK-RN8M7-J3C4G-BBGYM-88CYV
mondo,2016,HFTND-W9MK4-8B7MJ-B6C4G-XQBR2
outlook,2010,7YDC2-CWM8M-RRTJC-8MDVC-X3DWQ
outlook,2013,QPN8Q-BJBTJ-334K3-93TGY-2PMBT
outlook,2016,R69KK-NTPKF-7M3Q4-QYBHW-6MT9B
outlook,2019,7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK
outlook,2021,C9FM6-3N72F-HFJXB-TM3V9-T86R9
outlook,2024,D2F8D-N3Q3B-J28PV-X27HD-RJWB9
word,2019,PBX3G-NWMT6-Q7XBW-PYJGG-WXD33
word,2021,TN8H9-M34D3-Y64V9-TR72V-X79KV
word,2024,MQ84N-7VYDM-FXV7C-6K7CC-VFW9J
'@ | ConvertFrom-Csv
$Global:Kms_Servers_List = @'
Site
kms.digiboy.ir
hq1.chinancce.com
kms.cnlic.com
kms.chinancce.com
kms.ddns.net
franklv.ddns.net
k.zpale.com
m.zpale.com
mvg.zpale.com
kms.shuax.com
kensol263.imwork.net
annychen.pw
heu168.6655.la
xykz.f3322.org
kms789.com
dimanyakms.sytes.net
kms.03k.org
kms.lotro.cc
kms.didichuxing.com
zh.us.to
kms.aglc.cckms.aglc.cc
kms.xspace.in
winkms.tk
kms.srv.crsoo.com
kms.loli.beer
kms8.MSGuides.com
kms9.MSGuides.com
kms.zhuxiaole.org
kms.lolico.moe
kms.moeclub.org
'@ | ConvertFrom-Csv

# Base Ps1 operation
function Query-Basic {
    param (
        [Parameter(Mandatory = $true)]
        [string]$PropertyList,
        [Parameter(Mandatory = $true)]
        [string]$ClassName
    )
    try {
        $value = @($PropertyList).Replace(' ', '')
        $wmi_Object = Get-WmiObject -Query "SELECT $($value) FROM $($ClassName)" -ea 0
        if (-not $wmi_Object) { 
            return "Error:WMI_SEARCH_FAILURE"
        }
        return $wmi_Object
    }
    catch {
        return $null
    }
}
function Query-Advanced {
    param (
        [Parameter(Mandatory = $true)]
        [string]$PropertyList,
        [Parameter(Mandatory = $true)]
        [string]$ClassName,
        [Parameter(Mandatory = $true)]
        [string]$Filter
    )
    
    try {
        $value = @($PropertyList).Replace(' ', '')
        $Global:DBG = "SELECT $($value) FROM $($ClassName) WHERE ($($Filter))"
        $wmi_Object = Get-WmiObject -Query "SELECT $($value) FROM $($ClassName) WHERE ($($Filter))" -ea 0
        if (-not $wmi_Object) { 
            return "Error:WMI_SEARCH_FAILURE"
        }
        return $wmi_Object
    }
    catch {
        return $null
    }
}
function Activate-Class {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Class,
        [Parameter(Mandatory = $true)]
        [string]$Id
    )
    $Global:lastErr = $null
    try {
        (gwmi $Class -Filter "ID='$($Id)'").Activate()
        $Global:lastErr = 0
    }
    catch {
        $HResult = "0x{0:x}" -f @($_.Exception.InnerException).HResult
        $Global:lastErr = $HResult
    }
}
function Uninstall-ProductKey {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Class,
        [Parameter(Mandatory = $true)]
        [string]$Filter
    )
    try {
        Invoke-CimMethod -MethodName UninstallProductKey -Query "SELECT * FROM $($Class) WHERE ($($Filter))"
        $Global:lastErr = 0
    }
    catch {
        $HResult = "0x{0:x}" -f @($_.Exception.InnerException).HResult
        $Global:lastErr = $HResult
    }
}
function Install-ProductKey {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ProductKey
    )
    $ErrorActionPreference = "Stop"
    try {
        Invoke-CimMethod -MethodName InstallProductKey -Query "SELECT * FROM SoftwareLicensingService" -Arguments @{ ProductKey = $ProductKey }
        $Global:lastErr = 0
    }
    catch {
        $HResult = "0x{0:x}" -f @($_.Exception.InnerException).HResult
        $Global:lastErr = $HResult
    }
}
function Set-DefinedEntities {

    # --- Detect x64 paths ---
    if (Test-Path "$env:windir\SysWOW64\cscript.exe") {
        $global:cscript = "$env:windir\SysWOW64\cscript.exe"
    }
    if (Test-Path "$env:windir\SysWOW64\slmgr.vbs") {
        $global:slmgr = "$env:windir\SysWOW64\slmgr.vbs"
    }
    if (Test-Path "$env:ProgramFiles(x86)\Microsoft Office\office14\OSPP.vbs") {
        $global:OSPP_14 = "$env:ProgramFiles(x86)\Microsoft Office\office14\OSPP.vbs"
    }
    if (Test-Path "$env:ProgramFiles(x86)\Microsoft Office\office15\OSPP.vbs") {
        $global:OSPP_15 = "$env:ProgramFiles(x86)\Microsoft Office\office15\OSPP.vbs"
        $global:OSPP = $global:OSPP_15
    }
    if (Test-Path "$env:ProgramFiles(x86)\Microsoft Office\Office16\OSPP.vbs") {
        $global:OSPP_16 = "$env:ProgramFiles(x86)\Microsoft Office\Office16\OSPP.vbs"
        $global:OSPP = $global:OSPP_16
    }
    if (Test-Path "$env:ProgramFiles(x86)\Microsoft Office\root\Licenses16") {
        $global:licenceDir = "$env:ProgramFiles(x86)\Microsoft Office\root\Licenses16"
    }
    if (Test-Path "$env:ProgramFiles(x86)\Microsoft Office\root") {
        $global:root = "$env:ProgramFiles(x86)\Microsoft Office\root"
    }

    # Registry path checks (x64)
    $regPaths64 = @(
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\propertyBag",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\ClickToRunStore",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
    )
    foreach ($path in $regPaths64) {
        if (Get-ItemProperty -Path $path -Name ProductReleaseIds -ea 0) {
            $global:Key = $path
        }
    }

    if ($pkg = Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun" -Name PackageGUID -ea 0) {
        $global:guid = $pkg.PackageGUID
    }

    # --- Detect x86 paths ---
    if (Test-Path "$env:windir\System32\cscript.exe") {
        $global:cscript = "$env:windir\System32\cscript.exe"
    }
    if (Test-Path "$env:windir\System32\slmgr.vbs") {
        $global:slmgr = "$env:windir\System32\slmgr.vbs"
    }
    if (Test-Path "$env:ProgramFiles\Microsoft Office\office14\OSPP.vbs") {
        $global:OSPP_14 = "$env:ProgramFiles\Microsoft Office\office14\OSPP.vbs"
    }
    if (Test-Path "$env:ProgramFiles\Microsoft Office\Office15\OSPP.vbs") {
        $global:OSPP_15 = "$env:ProgramFiles\Microsoft Office\Office15\OSPP.vbs"
        $global:OSPP = $global:OSPP_15
    }
    if (Test-Path "$env:ProgramFiles\Microsoft Office\Office16\OSPP.vbs") {
        $global:OSPP_16 = "$env:ProgramFiles\Microsoft Office\Office16\OSPP.vbs"
        $global:OSPP = $global:OSPP_16
    }
    if (Test-Path "$env:ProgramFiles\Microsoft Office\root\Licenses16") {
        $global:licenceDir = "$env:ProgramFiles\Microsoft Office\root\Licenses16"
    }
    if (Test-Path "$env:ProgramFiles\Microsoft Office\root") {
        $global:root = "$env:ProgramFiles\Microsoft Office\root"
    }

    # Registry path checks (x86)
    $regPaths86 = @(
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\propertyBag",
        "HKLM:\SOFTWARE\Microsoft\Office\16.0\ClickToRunStore",
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    )
    foreach ($path in $regPaths86) {
        if (Get-ItemProperty -Path $path -Name ProductReleaseIds -ea 0) {
            $global:Key = $path
        }
    }

    if ($pkg = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun" -Name PackageGUID -ea 0) {
        $global:guid = $pkg.PackageGUID
    }

    # --- Registry constants ---
    $global:OSPP_HKLM     = 'HKLM:\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform'
    $global:OSPP_USER     = 'HKU\S-1-5-20\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform'
    $global:XSPP_HKLM_X32 = 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform'
    $global:XSPP_HKLM_X64 = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform'
    $global:XSPP_USER     = 'HKU\S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform'
}
function Clean-RegistryKeys {
    $ErrorActionPreference = 'SilentlyContinue'

    # List of value names to delete
    $valuesToDelete = @(
        'KeyManagementServiceName',
        'KeyManagementServicePort',
        'DisableDnsPublishing',
        'DisableKeyManagementServiceHostCaching'
    )

    # Delete values from OSPP paths
    foreach ($name in $valuesToDelete) {
        Remove-ItemProperty -Path $global:OSPP_USER -Name $name -Force
        Remove-ItemProperty -Path $global:OSPP_HKLM -Name $name -Force
    }

    # Delete values from XSPP paths (SLMGR.VBS)
    foreach ($name in $valuesToDelete) {
        Remove-ItemProperty -Path $global:XSPP_USER -Name $name -Force
        Remove-ItemProperty -Path $global:XSPP_HKLM_X32 -Name $name -Force
        Remove-ItemProperty -Path $global:XSPP_HKLM_X64 -Name $name -Force
    }

    # WMI Nethood subkeys to delete
    $subKeys = @(
        '55c92734-d682-4d71-983e-d6ec3f16059f',
        '0ff1ce15-a989-479d-af46-f275c6370663',
        '59a52881-a989-479d-af46-f275c6370663'
    )

    foreach ($subKey in $subKeys) {
        Remove-Item -Path "$global:XSPP_USER\$subKey" -Recurse -Force
        Remove-Item -Path "$global:XSPP_HKLM_X32\$subKey" -Recurse -Force
        Remove-Item -Path "$global:XSPP_HKLM_X64\$subKey" -Recurse -Force
    }

    $ErrorActionPreference = 'Continue'
}
function Update-RegistryKeys {
    param (
        [Parameter(Mandatory = $true)]
        [string]$KmsHost,
        [Parameter(Mandatory = $true)]
        [string]$KmsPort,

        [string]$SubKey,
        [string]$Id
    )

    # remove KMS38 lock --> From MAS PROJECT, KMS38_Activation.cmd
    $SID = New-Object SecurityIdentifier('S-1-5-32-544')
    $Admin = ($SID.Translate([NTAccount])).Value
    $ruleArgs = @("$Admin", "FullControl", "Allow")
    $path = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f'
    $regKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry64').OpenSubKey($path, 'ReadWriteSubTree', 'ChangePermissions')
    if ($regKey) {
        $acl = $regKey.GetAccessControl()
        $rule = [RegistryAccessRule]::new.Invoke($ruleArgs)
        $acl.ResetAccessRule($rule)
        $regKey.SetAccessControl($acl)
    }

    $osppPaths = @(
        'HKCU:\Software\Microsoft\OfficeSoftwareProtectionPlatform',
        'HKLM:\Software\Microsoft\OfficeSoftwareProtectionPlatform'
    )

    $xsppPaths = @(
        'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform',
        'HKLM:\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform'
    )

    # Apply to OSPP paths (Office)
    foreach ($path in $osppPaths) {
        if (-not (Test-Path $path)) {
            New-Item -Path $path -Force | Out-Null}
        New-ItemProperty -Path $path -Name 'KeyManagementServiceName' -Value $KmsHost -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $path -Name 'KeyManagementServicePort' -Value $KmsPort -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $path -Name 'DisableDnsPublishing' -Value 0 -PropertyType DWord -Force | Out-Null
        New-ItemProperty -Path $path -Name 'DisableKeyManagementServiceHostCaching' -Value 0 -PropertyType DWord -Force | Out-Null
    }

    # Apply to XSPP paths (Windows)
    foreach ($path in $xsppPaths) {
        New-Item -Path $path -Force -ea 0 | Out-Null
        New-ItemProperty -Path $path -Name 'KeyManagementServiceName' -Value $KmsHost -PropertyType String -Force -ea 0 | Out-Null
        New-ItemProperty -Path $path -Name 'KeyManagementServicePort' -Value $KmsPort -PropertyType String -Force -ea 0 | Out-Null
        New-ItemProperty -Path $path -Name 'DisableDnsPublishing' -Value 0 -PropertyType DWord -Force -ea 0 | Out-Null
        New-ItemProperty -Path $path -Name 'DisableKeyManagementServiceHostCaching' -Value 0 -PropertyType DWord -Force -ea 0 | Out-Null
    }
    if (!$SubKey -or !$Id) {
        return
    }
    # WMI Subkey paths (XSPP + subkey + id)
    foreach ($base in $xsppPaths) {
        $wmiPath = Join-Path -Path $base -ChildPath "$SubKey\$Id"
        try {
            if (-not (Test-Path $wmiPath)) {
               New-Item -Path $wmiPath -Force -ErrorAction Stop | Out-Null
            }
            New-ItemProperty -Path $wmiPath -Name 'KeyManagementServiceName' -Value $KmsHost -PropertyType String -Force -ea 0 | Out-Null
            New-ItemProperty -Path $wmiPath -Name 'KeyManagementServicePort' -Value $KmsPort -PropertyType String -Force -ea 0 | Out-Null
        }
        catch { 
            # should i do something here ?
        }
    }
}
function Update-Year {
    param (
        [Parameter(Mandatory = $true)]
        [ValidatePattern('^\d{4}$')]
        [string]$Year
    )

    $Global:ProductYear = $Year
}
function Get-ProductYear {

    $Global:OfficeC2R = $null
    $Global:OfficeMsi16 = $null

    if (-not $Global:key) {
        $uninstallPaths = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
        )

        foreach ($path in $uninstallPaths) {
            if (Get-ChildItem -Path $path -ea 0 | Where-Object { $_.PSChildName -like '*Office16*' }) {
                $Global:OfficeMsi16 = $true
                Update-Year 2016
                return
            }
        }

        return
    }

    # Extract ProductReleaseIds
    try {
        $productRelease = (Get-ItemProperty -Path $Global:key -Name ProductReleaseIds -ErrorAction Stop).ProductReleaseIds
        if (-not $productRelease) { return }
        $Global:OfficeC2R = $true

        foreach ($year in $Global:SupportedBuildYear) {
            if ($productRelease -like "*$year*") {
                Update-Year $year
                return
            }
        }
    } catch {
        return
    }

    # Default to 2016 if nothing matched
    Update-Year 2016
}
function Wmi-Activation {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$Grace,
        [Parameter(Mandatory = $true)]
        [string]$Ks,
        [Parameter(Mandatory = $true)]
        [string]$Kp,
        [Parameter(Mandatory = $true)]
        [string]$Km
    )

    # Wmi based activation
    # using SoftwareLicensingProduct class for office
    # using SoftwareLicensingService class for windows
    # using OfficeSoftwareProtectionProduct class for specific win7 case

    $subKey        = $null
    $SPP_ACT_CLASS = $null
    $SPP_KMS_Class = $null
    $SPP_KMS_Where = $null

    if ($Km -match 'Windows') {
        $subKey = '55c92734-d682-4d71-983e-d6ec3f16059f'
        $SPP_KMS_Class = 'SoftwareLicensingService'
        $SPP_KMS_Where = 'version is not null'
        $SPP_ACT_CLASS = 'SoftwareLicensingProduct'
    }

    if ($Km -match 'office') {
        $subKey = '0ff1ce15-a989-479d-af46-f275c6370663'

        if ($Global:Windows_7_Or_Earlier) {

            # Office in windows Windows 7 and less use 2 classes
  	        # OfficeSoftwareProtectionService for KMS settings 
  	        # OfficeSoftwareProtectionProduct for activation

            $SPP_KMS_Class = 'OfficeSoftwareProtectionService'
            $SPP_KMS_Where = 'version is not null'
            $SPP_ACT_CLASS = 'OfficeSoftwareProtectionProduct'
        }

        if ($Global:14_X_Mode) {

            # Office 2010 Classes
  	        # OfficeSoftwareProtectionService for KMS settings 
  	        # OfficeSoftwareProtectionProduct for activation

            $subKey = '59a52881-a989-479d-af46-f275c6370663'
            $SPP_KMS_Class = 'OfficeSoftwareProtectionService'
            $SPP_KMS_Where = 'version is not null'
            $SPP_ACT_CLASS = 'OfficeSoftwareProtectionProduct'
        }
    }

    if (-not $SPP_ACT_CLASS) {
        $SPP_ACT_CLASS = 'SoftwareLicensingProduct'
    }
    $Product_Licensing_Class = $SPP_ACT_CLASS
    $Product_Licensing_Where = "Id like '%$Id%'"
    if (-not $SPP_KMS_Class) {
        $SPP_KMS_Class = $Product_Licensing_Class
        $SPP_KMS_Where = $Product_Licensing_Where
    }

    Update-RegistryKeys -KmsHost $ks -KmsPort $Kp -SubKey $subKey -Id $Id
    Write-Host '+++ Activating +++'
    Write-Host '...................'
    $null = Activate-Class -Class $Product_Licensing_Class -Id $Id
    $wmi_Object = Query-Advanced -PropertyList 'GracePeriodRemaining' -ClassName $Product_Licensing_Class -Filter $Product_Licensing_Where
    if ($wmi_Object -and $wmi_Object.GracePeriodRemaining) {
        Write-Host "Old Grace               = $Grace"
        Write-Host "New Grace               = $($wmi_Object.GracePeriodRemaining)"
    }
    if ($Global:lastErr -eq 0 ) {
        Write-Host "Status                  = Succeeded (Error 0x$lastErr)"
    } else {
        Write-Host "Status                  = Failed (Error 0x$lastErr)"
    }
}
function Check-Activation {
    param (
        [Parameter(Mandatory=$true)]
        [string]$product,
        [Parameter(Mandatory=$true)]
        [string]$licenceType
    )

    $Global:ProductIsActivated = $false
    $LicensingProductClass = 'SoftwareLicensingProduct'
    if ($product -match 'Office' -and $Global:Windows_7_Or_Earlier) {
        $LicensingProductClass = 'OfficeSoftwareProtectionProduct'
    }
    $wmiSearch = "Name like '%$product%' and Description like '%$licenceType%' and PARTIALPRODUCTKEY IS NOT NULL and GenuineStatus = 0 and LicenseStatus = 1"
    $output = Query-Advanced -PropertyList "name" -ClassName $LicensingProductClass -Filter $wmiSearch

    # Evaluate output conditions
    if (-not $output -or ($output -is [string])) {
        return }

    if ($product -ieq 'Windows') {
        Write-Host "=== $product is $licenceType activated ==="
    } else {
        Write-Host "=== Office $product is $licenceType activated ==="
    }
    $Global:ProductIsActivated = $true
    Write-Host
}
function Search_VL_Products {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProductName
    )

    $ApplicationId = $null
    $LicensingProductClass = 'SoftwareLicensingProduct'

    if ($ProductName -match 'Office') {
        $ApplicationId = '0ff1ce15-a989-479d-af46-f275c6370663'
        if ($Global:Windows_7_Or_Earlier) {
            $LicensingProductClass = 'OfficeSoftwareProtectionProduct'
        }
    }
    elseif ($ProductName -match 'windows') {
        $ApplicationId = '55c92734-d682-4d71-983e-d6ec3f16059f'
    }

    $PropertyList = 'ID,LicenseStatus,PartialProductKey,GenuineStatus,Name,GracePeriodRemaining'
    $Filter = "ApplicationId like '%$ApplicationId%' and Description like '%KMSCLIENT%'"
    if (-not $Global:Windows_7_Or_Earlier -and ($ProductName -notmatch 'Office')) {
        $Filter += " and PartialProductKey is not null and LicenseFamily is not null and LicenseDependsOn is NULL"
    }

    # Assuming Query-Advanced is defined elsewhere in your script, you call it here:
    return Query-Advanced -PropertyList $PropertyList -ClassName $LicensingProductClass -Filter $Filter
}
function Search_14_X_VL_Products {
    $ApplicationId = '59a52881-a989-479d-af46-f275c6370663'
    $LicensingProductClass = 'OfficeSoftwareProtectionProduct'
    $PropertyList = 'ID,LicenseStatus,PartialProductKey,GenuineStatus,Name,GracePeriodRemaining'
    $Filter = "ApplicationId like '%$ApplicationId%' and Description like '%KMSCLIENT%' and PartialProductKey is not null"

    # Call your query function
    return Query-Advanced -PropertyList $PropertyList -ClassName $LicensingProductClass -Filter $Filter
}
function Search_Office_VL_Products {
    param (
        [Parameter(Mandatory=$true)]
        [string]$T_Year,
        [Parameter(Mandatory=$true)]
        [string]$T_Name
    )

    $LicensingProductClass = 'SoftwareLicensingProduct'
    $ApplicationId = '0ff1ce15-a989-479d-af46-f275c6370663'
    if ($Global:Windows_7_Or_Earlier) {
        $LicensingProductClass = 'OfficeSoftwareProtectionProduct'
    }

    # Extract last two digits of year
    $yearPart = $T_Year.Substring($T_Year.Length - 2)

    $Filter = "ApplicationId like '%$ApplicationId%' and Description like '%KMSCLIENT%' and PartialProductKey is not null"
    $filter += " and Name like '%office $yearPart%' and Name like '%$T_Name%'"
    return Query-Advanced -PropertyList "Name" -ClassName $LicensingProductClass -Filter $Filter
}
function Uninstall-PartialProductKey {
    param (
        [Parameter(Mandatory = $true)]
        [string]$PartialKey,
        [Parameter(Mandatory = $true)]
        [bool]$IsWindows7OrEarlier = $false
    )

    $LicensingProductClass = if ($IsWindows7OrEarlier) {
        'OfficeSoftwareProtectionProduct'
    } else {
        'SoftwareLicensingProduct'
    }

    # Create a WQL-compatible filter
    $Filter = "PartialProductKey LIKE '%$PartialKey%'"

    # Call the function to uninstall
    Uninstall-ProductKey -Class $LicensingProductClass -Filter $Filter
}
function Remove_Office_Products {
    param (
        [Parameter(Mandatory = $true)]
        [string]$T_Year,
        [Parameter(Mandatory = $true)]
        [string]$T_Name
    )

    $LicensingProductClass = 'SoftwareLicensingProduct'
    if ($Global:Windows_7_Or_Earlier) {
        $LicensingProductClass = 'OfficeSoftwareProtectionProduct'
    }

    $PropertyList = 'ID,LicenseStatus,PartialProductKey,GenuineStatus,Name,GracePeriodRemaining'

    # Extract last 2 digits of year
    $yearPart = $T_Year.Substring($T_Year.Length - 2)

    # Construct WMI filter
    $Filter = "PartialProductKey is not null and Name like '%office $yearPart%' and Name like '%$T_Name%'"

    # Call helper function to uninstall product keys matching filter
    Uninstall-ProductKey -Class $LicensingProductClass -Filter $Filter
}
function Integrate-License {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Product,

        [string]$Year
    )

    if (-not $global:root) {
        Write-Host "Root path not defined."
        return
    }

    # Determine product name
    if ($Year -match '2016') {
        $FinalProduct = "${Product}Volume.16"
    } else {
        $FinalProduct = "${Product}${Year}Volume.16"
    }

    # Construct the integrator path
    $IntegratorPath = Join-Path -Path $global:root -ChildPath "Integration\integrator.exe"

    # Build argument list
    $args = @(
        '/I',
        '/License',
        "PRIDName=$FinalProduct",
        "PackageGUID=$global:Guid",
        '/Global',
        '/C2R',
        "PackageRoot=$Root"
    )

    # Execute
    & $IntegratorPath @args
}
Function Service-Check {

    if ($Global:LocalKms) {

        # if defined LocalKms (
        #  set "bin=A64.dll,x64.dll,x86.dll"
        #  for %%# in (!bin!) do if not exist "%fs%\%%#" set "LocalKms="
        #  if not defined LocalKms (
        #  	echo.
        #  	echo Local Activation files Is Missing, Switch back to Online KMS
        #  	timeout 5
        #  	Clear-host
        #  )
        #)

    }

    try {
        # Query the Winmgmt service "Start" registry value
        $startValue = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Winmgmt" -Name "Start" -ErrorAction Stop).Start

    }
    catch {
        # If registry key not found or error occurs, just continue (no action)
        $startValue = $null
    }

    if ($startValue -eq 4) {
        # Start type 4 means "Disabled" - try to set to automatic
        $result = sc.exe config Winmgmt start=auto 2>&1

        if ($LASTEXITCODE -ne 0) {
            Clear-Host
            Write-Host ""
            Write-Host "#### ERROR:: WMI FAILURE"
            Write-Host ""
            Write-Host "- winmgmt service is not working"
            Write-Host ""
            Pause
            exit 1
        }
    }

    $wmi_check = (gwmi Win32_Processor -ea 0).AddressWidth -match "32|64"
    if ($wmi_check -eq $false) {
        Write-Host "#### ERROR:: WMI FAILURE"
        Write-Host ""
        Write-Host "- The WMI repository is broken"
        Write-Host "- winmgmt service is not working"
        Write-Host "- Script run in a sandbox/limited environment"
    }

    # Check if Windows 7 or earlier
    if ($Global:osVersion.Build -le 7601) {
        $Global:Windows_7_Or_Earlier = $true
    } else {
        $Global:Windows_7_Or_Earlier = $false
    }

    # Check if unsupported build (less than 2600)
    if ($Global:osVersion.Build -lt 2600) {
        Write-Host ""
        Write-Host "ERROR ### System not supported"
        Write-Host ""
        Read-Host -Prompt "Press Enter to exit"
        exit
    }

    $Global:xBit = if ([Environment]::Is64BitOperatingSystem) { 64 } else { 32 }

    #fix for IoTEnterpriseS new key
    #https://forums.mydigitallife.net/threads/windows-11-ltsc.87144/#post-1795872
}
Function LetsActivate {
    
    $Global:KmsServer = $null
    if ($Global:LocalKms -eq $false) {
        Write-Host "Start Activation Process"
        Write-Host "........................"
        Write-Host "Look For Online Kms Servers"
        foreach ($server in $Global:Kms_Servers_List.Site) {
            write-host "Check if $($server):1688 is Online"
            $connection = Test-NetConnection -ComputerName $server -Port 1688 -InformationAction SilentlyContinue
            if ($connection.TcpTestSucceeded) {
                $Global:KmsServer = $server
                break
            }
        }
        $ProgressPreference = 'Continue'
    }

    if (($Global:LocalKms -eq $false) -and (
    -not $Global:KmsServer)) {
        Write-Host
        write-host "ERROR ##### didnt found any available online kms server"
        Write-Host
        return
    }

    if (($Global:LocalKms -eq $false) -and $Global:KmsServer) {
        write-host "Winner Winner Chicken dinner"
        write-host
        Update-RegistryKeys -KmsHost $Global:KmsServer -KmsPort 1688 
    }
    if ($Global:LocalKms -eq $true) {
        #StartKMSActivation
    }

    WindowsHelper
    OfficeHelper

    if ($Global:LocalKms -eq $true) {
        #StopKMSActivation
    }
}
Function WindowsHelper {
    if (-not $global:slmgr) {
        Write-Host "ERROR ##### didnt found any Windows products / SLMGR.VBS IS Missing"
        Write-Host
        return
    }

    $global:VL_Product_Not_Found = $false
    $output = Search_VL_Products -ProductName windows
    if (-not $output -or ($output -is [string])) {
        $global:VL_Product_Not_Found = $true
    }

    if ($global:VL_Product_Not_Found) {
        
        $null = Check-Activation Windows RETAIL
        if ($global:ProductIsActivated) {
            return
        }

        $null = Check-Activation Windows MAK
        if ($global:ProductIsActivated) {
            return
        }

        $null = Check-Activation Windows OEM
        if ($global:ProductIsActivated) {
            return
        }
        
        Windows_Licence_Worker
        if (-not $global:serial) {
            return
        }
        $global:VL_Product_Not_Found = $false
        $null = Install-ProductKey -ProductKey $global:serial
        Write-Host
        $output = Search_VL_Products -ProductName windows
        if (-not $output -or ($output -is [string])) {
            Write-Host "ERROR ##### didnt found any windows volume products"
            Write-Host
        }
    }

    if ($output -and ($output -isnot [string])) {
      
      Write-Host $output.Name
      Write-Host "...................................."
      Write-Host "License / Genuine 	= $($output.LicenseStatus) / $($($output.GenuineStatus))"
      Write-Host "Period Remaining 	= $([Math]::Floor($output.GracePeriodRemaining/60/24))"
      Write-Host "Product ID 	        = $($output.ID)"

      if ($output.GracePeriodRemaining -gt 259200) {
        Write-Host
        Write-Host "=== Windows is KMS38/KMS4K activated ==="
        Write-Host
      }

      if ($output.GracePeriodRemaining -le 259200) {
        if ($Global:LocalKms -eq $false) {
           Write-Host
           Wmi-Activation -Id $output.ID -Grace $output.GracePeriodRemaining -Ks $Global:KmsServer -Kp 1688 -Km windows
        }
        if ($Global:LocalKms -eq $true) {
           Write-Host
           Wmi-Activation -Id $output.ID -Grace $output.GracePeriodRemaining -Ks $Global:IP_ADDRESS -Kp 1688 -Km windows
        }
      }
    }
}
Function OfficeHelper {
    param (
        [bool]$Is14XMode = $false
    )

    $Global:14_X_Mode = [bool]$Is14XMode
    if (!$global:OSPP_14 -and !$global:OSPP_15 -and !$global:OSPP_16 ) {
        Write-Host
        Write-Host "ERROR ##### didnt found any Office products / OSPP.VBS IS Missing"
        Write-Host
        return
    }

    $ohook_found = $false
    $paths = @(
        "$env:ProgramFiles\Microsoft Office\Office15\sppc*.dll",
        "$env:ProgramFiles\Microsoft Office\Office16\sppc*.dll",
        "$env:ProgramW6432\Microsoft Office\Office15\sppc*.dll",
        "$env:ProgramW6432\Microsoft Office\Office16\sppc*.dll",
        "$env:ProgramFiles(x86)\Microsoft Office\Office15\sppc*.dll",
        "$env:ProgramFiles(x86)\Microsoft Office\Office16\sppc*.dll"
    )

    foreach ($path in $paths) {
        if (Get-ChildItem -Path $path -Filter 'sppc*.dll' -Attributes ReparsePoint -ea 0) {
            $ohook_found = $true
            break
        }
    }

    # Also check the root\vfs paths
    $vfsPaths = @(
        "$env:ProgramFiles\Microsoft Office 15\root\vfs\System\sppc*.dll",
        "$env:ProgramFiles\Microsoft Office 15\root\vfs\SystemX86\sppc*.dll",
        "$env:ProgramFiles\Microsoft Office\root\vfs\System\sppc*.dll",
        "$env:ProgramFiles\Microsoft Office\root\vfs\SystemX86\sppc*.dll",
        "$env:ProgramW6432\Microsoft Office 15\root\vfs\System\sppc*.dll",
        "$env:ProgramW6432\Microsoft Office 15\root\vfs\SystemX86\sppc*.dll",
        "$env:ProgramW6432\Microsoft Office\root\vfs\System\sppc*.dll",
        "$env:ProgramW6432\Microsoft Office\root\vfs\SystemX86\sppc*.dll",
        "$env:ProgramFiles(x86)\Microsoft Office 15\root\vfs\System\sppc*.dll",
        "$env:ProgramFiles(x86)\Microsoft Office 15\root\vfs\SystemX86\sppc*.dll",
        "$env:ProgramFiles(x86)\Microsoft Office\root\vfs\System\sppc*.dll",
        "$env:ProgramFiles(x86)\Microsoft Office\root\vfs\SystemX86\sppc*.dll"
    )

    foreach ($path in $vfsPaths) {
        if (Get-ChildItem -Path $path -Filter 'sppc*.dll' -Attributes ReparsePoint -ea 0) {
            $ohook_found = $true
            break
        }
    }

    if ($ohook_found) {
        Write-Host
        Write-Host "=== Office is Ohook activated ==="
        Write-Host
        return
    }
    if (!$Is14XMode) {
       Office_Licence_Worker
    }

    $global:VL_Product_Not_Found = $false
    if ($Is14XMode) {
      $output = Search_14_X_VL_Products -ProductName office
    } else {
      $output = Search_VL_Products -ProductName office
    }

    if (-not $output -or ($output -is [string])) {
        $global:VL_Product_Not_Found = $true
    }

    if ($Global:VL_Product_Not_Found) {
        if ($Global:14_X_Mode) {
            Write-Host "ERROR ##### didn't find any Office 14.X volume products"
            return
        }

        if ($Global:OSPP_14) {
            Write-Host "ERROR ##### didn't find any Office 15.X 16.X volume products"
            Write-Host
            OfficeHelper -Is14XMode $true
            return
        }

        if (-not $Global:OSPP_14) {
            Write-Host "ERROR ##### didn't find any Office 14.X 15.X 16.X volume products"
            return
        }

        Write-Host "ERROR ##### Wtf happened now ??"
        return
    }

    if ($output -and ($output -isnot [string])) {
      foreach ($wmi_object in $output) {
          Write-Host
          Write-Host $wmi_object.Name
          Write-Host "...................................."
          Write-Host "License / Genuine 	= $($wmi_object.LicenseStatus) / $($($wmi_object.GenuineStatus))"
          Write-Host "Period Remaining 	= $([Math]::Floor($wmi_object.GracePeriodRemaining/60/24))"
          Write-Host "Product ID 	        = $($wmi_object.ID)"

          if ($wmi_object.GracePeriodRemaining -gt 259200) {
            Write-Host
            Write-Host "=== Office is KMS4K activated ==="
            Write-Host
          }

          if ($wmi_object.GracePeriodRemaining -le 259200) {
            if ($Global:LocalKms -eq $false) {
               Write-Host
               Wmi-Activation -Id $wmi_object.ID -Grace $wmi_object.GracePeriodRemaining -Ks $Global:KmsServer -Kp 1688 -Km windows
            }
            if ($Global:LocalKms -eq $true) {
               Write-Host
               Wmi-Activation -Id $wmi_object.ID -Grace $wmi_object.GracePeriodRemaining -Ks $Global:IP_ADDRESS -Kp 1688 -Km windows
            }
          }
      }
    }

    if (-not $Is14XMode -and ($Global:OSPP_14)) {
      OfficeHelper -Is14XMode $true
      return
    }

    write-host
    write-host "Search for 14.X Products"
    write-host "........................"
    write-host "--- 404 not found"
    return
}
Function Windows_Licence_Worker {

    $Global:serial = $null
    $EditionID = Get-ProductID
    $Global:VL_Product_Not_Found = $false
    $LicensingProductClass = 'SoftwareLicensingProduct'

    # fix for IoTEnterpriseS new key
    # https://forums.mydigitallife.net/threads/windows-11-ltsc.87144/#post-1795872

    if ($EditionID -eq 'ProfessionalSingleLanguage') {
        $EditionID = 'Professional' }
    elseif ($EditionID -eq 'ProfessionalCountrySpecific') {
        $EditionID = 'Professional' }
    elseif ($EditionID -eq 'IoTEnterprise') {
        $EditionID = 'Enterprise' }
    elseif ($EditionID -eq 'IoTEnterpriseK') {
        $EditionID = 'Enterprise' }
    elseif ($EditionID -eq 'IoTEnterpriseSK') {
        $EditionID = 'EnterpriseS' }
    elseif ($EditionID -eq 'IoTEnterpriseS') {
        if ($Global:osVersion.Build -lt 22610) {
            $EditionID = 'EnterpriseS'
            if ($Global:osVersion.Build -ge 19041 -and $Global:osVersion.UBR -ge 2788) {
                $EditionID = 'IoTEnterpriseS'
            }
        }
    }

    $wmiSearch = "Name like '%$EditionID%' and Description like '%VOLUME_KMSCLIENT%' and ApplicationId like '%55c92734-d682-4d71-983e-d6ec3f16059f%'"
    $output = Query-Advanced -PropertyList "ID" -ClassName $LicensingProductClass -Filter $wmiSearch

    # Evaluate output conditions
    if (-not $output -or ($output -is [string])) {
        Write-Host "ERROR ##### Couldn't find Any windows Supported ID"
        Write-Host
        return
    }

    # Blacklist IDs but you want to exclude these from output, so invert logic
    $blacklist = @(
        'b71515d9-89a2-4c60-88c8-656fbcca7f3a','af43f7f0-3b1e-4266-a123-1fdb53f4323b','075aca1f-05d7-42e5-a3ce-e349e7be7078',
        '11a37f09-fb7f-4002-bd84-f3ae71d11e90','43f2ab05-7c87-4d56-b27c-44d0f9a3dabd','2cf5af84-abab-4ff0-83f8-f040fb2576eb',
        '6ae51eeb-c268-4a21-9aae-df74c38b586d','ff808201-fec6-4fd4-ae16-abbddade5706','34260150-69ac-49a3-8a0d-4a403ab55763',
        '4dfd543d-caa6-4f69-a95f-5ddfe2b89567','5fe40dd6-cf1f-4cf2-8729-92121ac2e997','903663f7-d2ab-49c9-8942-14aa9e0a9c72',
        '2cc171ef-db48-4adc-af09-7c574b37f139','5b2add49-b8f4-42e0-a77c-adad4efeeeb1'
    )

    # Filter output to exclude blacklisted IDs
    $Vol_Products = $output | Where-Object { $blacklist -notcontains $_.ID }

    if (-not $Vol_Products) {
        Write-Host "ERROR ##### Couldn't find Any windows Supported ID"
        Write-Host
        return
    }

    $idLookup = @{}
    $Vol_Products | ForEach-Object { $idLookup[$_.ID] = $true }
    $MatchedIDs = $Global:Windows_Keys_List | Where-Object { $idLookup.ContainsKey($_.ID) }
    if (-not $MatchedIDs) {
        Write-Host "ERROR ##### Couldn't find Any windows Supported ID"
        Write-Host
        return
    }

    $Global:serial = $MatchedIDs[0].KEY
}
Function Office_Licence_Worker {
  Get-ProductYear
  if (-not $Global:ProductYear) {
    return
  }
  if ($Global:Officec2r) {
    $Global:ProductReleaseIds = (Get-ItemProperty -Path $Global:key -Name ProductReleaseIds -ea 0).ProductReleaseIds
  }
  if (!$Global:OfficeMsi16 -and !$Global:ProductReleaseIds) {
    return
  }
  $Global:ProductReleaseIds_ = $Global:ProductReleaseIds -split ','
  $ProductList = ("365","HOME","Professional","Private")
  
  $ConvertToMondo = $false
  $ConvertToMondo_ = $false
  $Global:Office_Product_Not_Found = $false

  # $ProductList Check [1]
  foreach ($product in $ProductList) {
    
    $SelectedX = $false

    # Start Loop #
    if ($Global:ProductReleaseIds -match $product) {
      $SelectedX = $true
    }
    if ($Global:OfficeMsi16) {
        # Query the registry for Office16 in both 64-bit and 32-bit paths
        $office16KeyPath1 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        $office16KeyPath2 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

        # Check for matches in the 64-bit registry path
        $matchFound = Get-ItemProperty -Path $office16KeyPath1 -ea 0 | ? { $_.DisplayName -like "*Office16.$product*" }
        if ($matchFound) {
            $SelectedX = $true
        }

        # Check for matches in the 32-bit registry path
        if (-not $matchFound) {
            $matchFound = Get-ItemProperty -Path $office16KeyPath2 -ea 0 | ? { $_.DisplayName -like "*Office16.$product*" }
            if ($matchFound) {
                $SelectedX = $true
            }
        }
    }
    if ($SelectedX -eq $true) {
      # $SelectedX Start
      $MoveToNext = $false
      if ($product -match 'HOME') {
        $Global:ProductReleaseIds_ | % {
          if (($_ -match "HOME") -and ($_ -match "365")) {
            $MoveToNext = $true
          }
        }
      }
      if (!$MoveToNext) {
        $Global:Office_Product_Not_Found = $true
        
        Check-Activation -product $product -licenceType RETAIL
        if ($Global:ProductIsActivated) {
          $Global:Office_Product_Not_Found = $false
        }

        Check-Activation -product $product -licenceType MAK
        if ($Global:ProductIsActivated) {
          $Global:Office_Product_Not_Found = $false
        }

        if ($Global:Office_Product_Not_Found -eq $true) {
          $ConvertToMondo = $true
          Remove_Office_Products -T_Year $Global:ProductYear -T_Name $product
        }
      }
      # $SelectedX End
    }
    # END Loop #
  }

  # Convert to mondo if needed.!
  if ($ConvertToMondo) {
    $tYear = '2016'
    $ProductYear = '2016'
    $tProduct = 'Mondo'
    
    $global:VL_Product_Not_Found = $false
    $output = Search_Office_VL_Products -T_Year $tYear -T_Name $tProduct
    if (-not $output -or ($output -is [string])) {
      $global:VL_Product_Not_Found = $true
    }

    if ($global:VL_Product_Not_Found) {
      Check-Activation -product $tProduct -licenceType RETAIL
      if ($global:ProductIsActivated) {
        $global:VL_Product_Not_Found = $false
      }
    }

    if ($global:VL_Product_Not_Found) {
      Check-Activation -product $tProduct -licenceType MAK
      if ($global:ProductIsActivated) {
        $global:VL_Product_Not_Found = $false
      }
    }

    if ($global:VL_Product_Not_Found) {
      Integrate-License -Product $tProduct -Year $ProductYear
      $files = Get-ChildItem -Path "$licenceDir\$tProduct*VL_KMS*.xrm-ms" -File -ea 0

      if ($files -and $files.FullName) {
          write-host
          Manage-SLHandle -Release | Out-null
          SL-InstallLicense -LicenseInput $files.FullName | Out-Null
      }
      $pInfo = $global:Office_Keys_List | ? Product -eq $tProduct | ? Year -EQ $tYear
      if ($pInfo) {
          write-host
          Manage-SLHandle -Release | Out-null
          SL-InstallProductKey -Keys ($pInfo.Key) | Out-Null
      }

    }
    $ConvertToMondo_ = $true
  }

  # $ProductList Check [2]
  if ($ConvertToMondo_) {
    $ProductList = @("publisher", "ProjectPro", "ProjectStd", "VisioStd", "VisioPro")
  } else {
    $ProductList = @("proplus", "Standard", "mondo", "word", "excel", "powerpoint", "Skype", "access", "outlook", "publisher", "ProjectPro", "ProjectStd", "VisioStd", "VisioPro", "OneNote")
  }
  foreach ($product in $ProductList) {
    # >> Start <<
    $SelectedX = $false
    if ($Global:OfficeMsi16) {
        # Query the registry for Office16 in both 64-bit and 32-bit paths
        $office16KeyPath1 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        $office16KeyPath2 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

        # Check for matches in the 64-bit registry path
        $matchFound = Get-ItemProperty -Path $office16KeyPath1 -ea 0 | ? { $_.DisplayName -like "*Office16.$product*" }
        if ($matchFound) {
            $SelectedX = $true
        }

        # Check for matches in the 32-bit registry path
        if (-not $matchFound) {
            $matchFound = Get-ItemProperty -Path $office16KeyPath2 -ea 0 | ? { $_.DisplayName -like "*Office16.$product*" }
            if ($matchFound) {
                $SelectedX = $true
            }
        }
    }
    if ($Global:Officec2r) {
        if ($Global:ProductReleaseIds -match $product) {
          $SelectedX = $true
          if (!$ConvertToMondo) {
            $Global:ProductReleaseIds_ | % {
                if ($_ -match $product) {
                  $ProductYear = $null
                  if ($_ -match '2024') {
                      $ProductYear = 2024
                  }
                  elseif ($_ -match '2021') {
                      $ProductYear = 2021
                  }
                  elseif ($_ -match '2019') {
                      $ProductYear = 2019
                  }
                  if (!$ProductYear -or ($product -eq "OneNote")) {
                      $ProductYear = 2016
                  }
               }
            }
          }
        }
    }

    if ($SelectedX) {

        ## >> Start
        $global:VL_Product_Not_Found = $false
        $output = Search_Office_VL_Products -T_Year $ProductYear -T_Name $product
        if (-not $output -or ($output -is [string])) {
          $global:VL_Product_Not_Found = $true
        }

        if ($global:VL_Product_Not_Found) {
          Check-Activation -product $product -licenceType RETAIL
          if ($global:ProductIsActivated) {
            $global:VL_Product_Not_Found = $false
          }
        }

        if ($global:VL_Product_Not_Found) {
          Check-Activation -product $product -licenceType MAK
          if ($global:ProductIsActivated) {
            $global:VL_Product_Not_Found = $false
          }
        }
        if ($global:VL_Product_Not_Found -eq $true) {
          @(2016, 2019, 2021, 2024) | % {
            Remove_Office_Products -T_Year $_ -T_Name $product
          }
          $year = if ($ProductYear -ne '2016') { $ProductYear } else { '2016' }
          $Client = Get-ChildItem -Path "$licenceDir\Client*.xrm-ms" -File -ErrorAction SilentlyContinue
          $Pkey = Get-ChildItem -Path "$licenceDir\pkeyconfig*.xrm-ms" -File -ErrorAction SilentlyContinue
          $productLicense = Get-ChildItem -Path "$licenceDir\$product*$year*VL_KMS*.xrm-ms" -File -ErrorAction SilentlyContinue | ? { $_.Name -notlike '*preview*' }
          $files = ($Client,$Pkey,$productLicense)

          # it also install preview, which i don't want
          # Integrate-License -Product $product -Year $year

          if ($files -and $files.FullName) {
            write-host
            Manage-SLHandle -Release | Out-null
            SL-InstallLicense -LicenseInput $files.FullName | Out-Null
          }
          $pInfo = $global:Office_Keys_List | ? Product -eq $product | ? Year -EQ $year
          if ($pInfo) {
            write-host
            Manage-SLHandle -Release | Out-null
            SL-InstallProductKey -Keys ($pInfo.Key) | Out-Null
          }
        }
        ## >> End
    }
    # >> End <<
  }
}
# KMS Part -->

# Tsforge Part -->
$base64EncodedDll = @"
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
"@  # Paste your full Base64 string here
$compressedDllBytes = [Convert]::FromBase64String($base64EncodedDll)
$compressedStream = [MemoryStream]::new($compressedDllBytes)
$gzipStream = New-Object Compression.GZipStream($compressedStream, [Compression.CompressionMode]::Decompress)
$decompressedStream = New-Object MemoryStream
$gzipStream.CopyTo($decompressedStream)
$gzipStream.Close()
$decompressedDllBytes = $decompressedStream.ToArray()
[Assembly]::Load($decompressedDllBytes)

Function GetRandomKey([String] $ProductID) {
    try {
        $guid = [Guid]::Parse($ProductID)
        $pkc = [LibTSforge.SPP.PKeyConfig]::new()
        try {
          $pkc.LoadConfig($guid) | Out-Null
        } catch {
          $pkc.LoadAllConfigs([LibTSforge.SPP.SLApi]::GetAppId($guid)) | Out-Null
        }
        $config = [LibTSforge.SPP.ProductConfig]::new()
        $refConfig = [ref] $config
        $pkc.Products.TryGetValue($guid, $refConfig) | Out-Null
        return $config.GetRandomKey().ToString()
    } catch {
        Write-Warning "Failed to retrieve key for Product ID: $ProductID"
        return $null
    }
}
function Activate-License([string]$desc, [string]$ver, [string]$prod, [string]$tsactid) {
    if ($desc -match 'KMS|KMSCLIENT') {
        [LibTSforge.Activators.KMS4k]::Activate($ver, $prod, $tsactid)
    }
    elseif ($desc -match 'VIRTUAL_MACHINE_ACTIVATION') {
        [LibTSforge.Activators.AVMA4K]::Activate($ver, $prod, $tsactid)
    }
    elseif ($desc -match 'MAK|RETAIL|OEM|KMS_R2|WS12|WS12_R2|WS16|WS19|WS22|WS25') {
        
        $isInsiderBuild = $Global:osVersion.Build -ge 26100 -and $Global:osVersion.UBR -ge 4188
        $serverAvailable = Test-Connection -ComputerName "activation.sls.microsoft.com" -Count 1 -Quiet

        Write-Warning "Insider build detected: $isInsiderBuild"
        Write-Warning "Activation server reachable: $serverAvailable"

        if ($isInsiderBuild -and $serverAvailable) {
            
            Write-Warning "Selected activation mode: Static_Cid"
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $attempts = @(
                @(100055, 1000043, 1338662172562478),
                @(1345, 1003020, 6311608238084405)
            )
            foreach ($params in $attempts) {
                Write-Warning "$ver, $prod, $tsactid"
                [LibTSforge.Modifiers.SetIIDParams]::SetParams($ver, $prod, $tsactid, [LibTSforge.SPP.PKeyAlgorithm]::PKEY2009, $params[0], $params[1], $params[2])
                $instId = [LibTSforge.SPP.SLApi]::GetInstallationID($tsactid)
                Write-Warning "GetInstallationID, $instId"
                $confId = Call-WebService -requestType 1 -installationId $instId -extendedProductId "31337-42069-123-456789-04-1337-2600.0000-2542001"
                Write-Warning "Call-WebService, $confId"
                $result = [LibTSforge.SPP.SLApi]::DepositConfirmationID($tsactid, $instId, $confId)
                Write-Warning "DepositConfirmationID, $result"
                if ($result -eq 0) { break }
            }
            [LibTSforge.SPP.SPPUtils]::RestartSPP($ver)
        } 
        else {
            if ($isInsiderBuild) {
                Write-Host
                Write-Host "Activation could fail, you should select Vol' products instead" -ForegroundColor Green
                Write-Host
            }
            Write-Warning "Selected activation mode: Zero_Cid"
            [LibTSforge.Activators.ZeroCID]::Activate($ver, $prod, $tsactid)
        }
    }
    else {
        Write-Warning "Unknown license type: $desc"
        return
    }

    $ProductInfo = gwmi SoftwareLicensingProduct -ErrorAction SilentlyContinue -Filter "ID='$tsactid'"
    if (-not $ProductInfo) {
        Write-Warning "Product not found"
        return
    }

    if ($desc -match 'KMS|KMSCLIENT') {
        if ($ProductInfo.GracePeriodRemaining -gt 259200) {
            if ($desc -match 'KMS' -and (
                $desc -notmatch 'CLIENT')) {
                [LibTSforge.Modifiers.KMSHostCharge]::Charge($ver, $prod, $tsactid)
            }
            return
        }

        Write-Warning "KMS4K activation failed"
        return
    }

    if ($desc -notmatch 'KMS|KMSCLIENT') {
        if ($ProductInfo.LicenseStatus -ne 1) {
            Write-Warning "Activation Failed [ZeroCid/StaticCid/AVMA4K]"   
            return
        }
    }

}
function Capture-ConsoleOutput {
    param (
        [ScriptBlock]$ScriptBlock
    )

    $stringWriter = New-Object StringWriter
    $originalOut = [Console]::Out
    $originalErr = [Console]::Error

    try {
        [Console]::SetOut($stringWriter)
        [Console]::SetError($stringWriter)

        & $ScriptBlock
    }
    finally {
        [Console]::SetOut($originalOut)
        [Console]::SetError($originalErr)
    }

    return $stringWriter.ToString()
}
# TSForge part -->

# ActivationWs project -->
<#
This code is adapted from the ActivationWs project.
Original Repository: https://github.com/dadorner-msft/activationws

MIT License

Copyright (c) Daniel Dorner

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is furnished
to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT, OR OTHERWISE, ARISING FROM,
OUT OF, OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>
function Call-WebService {
    param (
        [int]$requestType,
        [string]$installationId,
        [string]$extendedProductId
    )

function Parse-SoapResponse {
    param (
        [Parameter(Mandatory=$true)]
        [string]$soapResponse
    )

    # Unescape the HTML-encoded XML content
    $unescapedXml = [System.Net.WebUtility]::HtmlDecode($soapResponse)

    # Check for ErrorCode in the unescaped XML
    if ($unescapedXml -match "<ErrorCode>(.*?)</ErrorCode>") {
        $errorCode = $matches[1]

        # Handle known error codes
        switch ($errorCode) {
            "0x7F" { throw [System.Exception]::new("The Multiple Activation Key has exceeded its limit.") }
            "0x67" { throw [System.Exception]::new("The product key has been blocked.") }
            "0x68" { throw [System.Exception]::new("Invalid product key.") }
            "0x86" { throw [System.Exception]::new("Invalid key type.") }
            "0x90" { throw [System.Exception]::new("Please check the Installation ID and try again.") }
            default { throw [System.Exception]::new("The remote server reported an error ($errorCode).") }
        }
    }

    # Check for ResponseType in the unescaped XML and handle it
    if ($unescapedXml -match "<ResponseType>(.*?)</ResponseType>") {
        $responseType = $matches[1]

        switch ($responseType) {
            "1" {
                # Extract the CID value
                if ($unescapedXml -match "<CID>(.*?)</CID>") {
                    return $matches[1]
                } else {
                    throw "CID not found in the XML."
                }
            }
            "2" {
                # Extract the ActivationRemaining value
                if ($unescapedXml -match "<ActivationRemaining>(.*?)</ActivationRemaining>") {
                    return $matches[1]
                } else {
                    throw "ActivationRemaining not found in the XML."
                }
            }
            default {
                throw "The remote server returned an unrecognized response."
            }
        }
    } else {
        throw "ResponseType not found in the XML."
    }
}
function Create-WebRequest {
    param (
        [Parameter(Mandatory=$true)]
        [string]$soapRequest  # Expecting raw XML text as input
    )
    
    # Define the URI and the SOAPAction
    $Uri = New-Object Uri("https://activation.sls.microsoft.com/BatchActivation/BatchActivation.asmx")
    $Action = "http://www.microsoft.com/BatchActivationService/BatchActivate"  # Correct SOAPAction URL
    
    # Create the web request
    $webRequest = [System.Net.HttpWebRequest]::Create($Uri)
    
    # Set necessary headers and content type
    $webRequest.Accept = "text/xml"
    $webRequest.ContentType = "text/xml; charset=`"utf-8`""
    $webRequest.Headers.Add("SOAPAction", $Action)
    $webRequest.Host = "activation.sls.microsoft.com"
    $webRequest.Method = "POST"
    
    try {
        # Convert the string to a byte array and insert into the request stream
        $byteArray = [Encoding]::UTF8.GetBytes($soapRequest)
        $webRequest.ContentLength = $byteArray.Length
        
        $stream = $webRequest.GetRequestStream()
        $stream.Write($byteArray, 0, $byteArray.Length)  # Write the byte array to the stream
        $stream.Close()  # Close the stream after writing
        
        return $webRequest  # Return the webRequest object
        
    } catch {
        throw $_  # Catch any exceptions and rethrow
    }
}
function Create-SoapRequest {
    param (
        [int]$requestType,
        [string]$installationId,
        [string]$extendedProductId
    )

    $activationRequestXml = @"
<ActivationRequest xmlns="http://www.microsoft.com/DRM/SL/BatchActivationRequest/1.0">
  <VersionNumber>2.0</VersionNumber>
  <RequestType>$requestType</RequestType>
  <Requests>
    <Request>
      <PID>$extendedProductId</PID>
      <IID>$installationId</IID>
    </Request>
  </Requests>
</ActivationRequest>
"@
    
    if ($requestType -ne 1) {
        $activationRequestXml = $activationRequestXml -replace '\s*<IID>.*?</IID>\s*', ''
    }

    # Convert string to Base64-encoded Unicode bytes
    $base64RequestXml = [Convert]::ToBase64String([Encoding]::Unicode.GetBytes($activationRequestXml))

    # HMACSHA256 calculation with hardcoded MacKey
    $hmacSHA = New-Object System.Security.Cryptography.HMACSHA256
    $hmacSHA.Key = [byte[]]@(
        254, 49, 152, 117, 251, 72, 132, 134, 156, 243, 241, 206, 153, 168, 144, 100, 
        171, 87, 31, 202, 71, 4, 80, 88, 48, 36, 226, 20, 98, 135, 121, 160, 0, 0, 0, 0,
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
    )
    $digest = [Convert]::ToBase64String($hmacSHA.ComputeHash([Encoding]::Unicode.GetBytes($activationRequestXml)))

    # Create SOAP envelope with the necessary values
    return @"
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <soap:Body>
    <BatchActivate xmlns="http://www.microsoft.com/BatchActivationService">
      <request>
        <Digest>$digest</Digest>
        <RequestXml>$base64RequestXml</RequestXml>
      </request>
    </BatchActivate>
  </soap:Body>
</soap:Envelope>
"@
}

    # Create SOAP request
    Write-Warning "$requestType, $installationId, $extendedProductId"
    $soapRequest = Create-SoapRequest -requestType ([int]$requestType) -installationId $installationId -extendedProductId $extendedProductId

    # Create Web Request
    $webRequest = Create-WebRequest -soapRequest $soapRequest

    try {
        # Send the web request and get the response synchronously
        $webResponse = $webRequest.GetResponse()
        $streamReader = New-Object StreamReader($webResponse.GetResponseStream())
        $soapResponse = $streamReader.ReadToEnd()

        # Parse and return the response
        $Response = Parse-SoapResponse -soapResponse $soapResponse
        return $Response

    } catch {
       Write-Warning "Request response failue"
    }
    
    return 0
}
# ActivationWs project -->

<#
Based on idea from ->

# Old source, work on W7
# GetSLCertify.cs by laomms
# https://forums.mydigitallife.net/threads/open-source-windows-7-product-key-checker.10858/page-14#post-1531837

# new source, work on Windows 8 & up, N key's
# keycheck.py by WitherOrNot
# https://github.com/WitherOrNot/winkeycheck

#>
function Call-AltWebService ([string]$ProductKey, [Guid]$SkuID = [guid]::Empty) {
    if ([string]::IsNullOrEmpty($ProductKey) -or (
        $ProductKey.LastIndexOf("n",[StringComparison]::InvariantCultureIgnoreCase) -lt 0)) {
    }

    $keyInfo = Decode-Key -Key $ProductKey
    if ($SkuID -eq [guid]::Empty) {
        $SkuId = Retrieve-ProductKeyInfo -CdKey $ProductKey | select -ExpandProperty SkuId
    }
    if (!$SkuId -or !$keyInfo) {
        Write-warning "Possible Error: SkuId not found for the product key."
        Write-warning "Possible Error: Failed to decode product key."
        return
    }

    [long]$group    = $keyInfo.Group
    [long]$serial   = $keyInfo.Serial
    [long]$security = $keyInfo.Security
    [int32]$upgrade = $keyInfo.Upgrade

    [System.Numerics.BigInteger]$act_hash = [BigInteger]$upgrade -band 1
    $act_hash = $act_hash -bor (([BigInteger]$serial -band ((1L -shl 30) - 1)) -shl 1)
    $act_hash = $act_hash -bor (([BigInteger]$group -band ((1L -shl 20) - 1)) -shl 31)
    $act_hash = $act_hash -bor (([BigInteger]$security -band ((1L -shl 53) - 1)) -shl 51)
    $bytes = $act_hash.ToByteArray()
    $KeyData = New-Object 'Byte[]' 13
    [Array]::Copy($bytes, 0, $KeyData, 0, [Math]::Min(13, $bytes.Length))
    $act_data = [Convert]::ToBase64String($KeyData)
    # End of original Encode-KeyData logic

    $value = [HttpUtility]::HtmlEncode("msft2009:$SkuId&$act_data")
    $requestXml = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
    xmlns:soapenc="http://schemas.xmlsoap.org/soap/encoding/"
    xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
    xmlns:xsd="http://www.w3.org/2001/XMLSchema"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <soap:Body>
        <RequestSecurityToken
            xmlns="http://schemas.xmlsoap.org/ws/2004/04/security/trust">
            <TokenType>PKC</TokenType>
            <RequestType>http://schemas.xmlsoap.org/ws/2004/04/security/trust/Issue</RequestType>
            <UseKey>
                <Values xsi:nil="1"/>
            </UseKey>
            <Claims>
                <Values
                    xmlns:q1="http://schemas.xmlsoap.org/ws/2004/04/security/trust" soapenc:arrayType="q1:TokenEntry[3]">
                    <TokenEntry>
                        <Name>ProductKey</Name>
                        <Value>$ProductKey</Value>
                    </TokenEntry>
                    <TokenEntry>
                        <Name>ProductKeyType</Name>
                        <Value>msft:rm/algorithm/pkey/2009</Value>
                    </TokenEntry>
                    <TokenEntry>
                        <Name>ProductKeyActConfigId</Name>
                        <Value>$value</Value>
                    </TokenEntry>
                </Values>
            </Claims>
        </RequestSecurityToken>
    </soap:Body>
</soap:Envelope>
"@

    try {
        $response = $null
        $webRequest = [System.Net.HttpWebRequest]::Create('https://activation.sls.microsoft.com/slpkc/SLCertifyProduct.asmx')
        $webRequest.Method      = "POST"
        $webRequest.Accept      = 'text/*'
        $webRequest.UserAgent   = 'SLSSoapClient'
        $webRequest.ContentType = 'text/xml; charset=utf-8'
        $webRequest.Headers.Add("SOAPAction", "http://microsoft.com/SL/ProductCertificationService/IssueToken");

        try {
            $byteArray = [System.Text.Encoding]::UTF8.GetBytes($requestXml)
            $webRequest.ContentLength = $byteArray.Length
            $stream = $webRequest.GetRequestStream()
            $stream.Write($byteArray, 0, $byteArray.Length)
            $stream.Close()
            $httpResponse = $webRequest.GetResponse()
            $streamReader = New-Object System.IO.StreamReader($httpResponse.GetResponseStream())
            $response = $streamReader.ReadToEnd()
            $streamReader.Close()
        }
        catch [System.Net.WebException] {
            if ($_.Exception) {
                $stream = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($stream)
                $response = $reader.ReadToEnd().ToString()
                $reader.Close()
            }
        }
        catch {
            Write-Error "Error: $($_.Exception.Message)"
            $global:error = $_
            return $null
        }

    }
    catch {
        Write-Error "Error: $($_.Exception.Message)"
        return $null
    }

    if ($response -ne $null) {
        [xml]$xmlResponse = $response
        if ($xmlResponse.Envelope.Body.Fault -eq $null) {
            return "Valid Key"
        } else {
            return Parse-ErrorMessage -MessageId ($xmlResponse.Envelope.Body.Fault.detail.HRESULT) -Flags ACTIVATION
        }
    }

    return "Error: No response received.", "", $false
}

<#
Based on idea from ->

# Old source, work on W7
# GetSLCertify.cs by laomms
# https://forums.mydigitallife.net/threads/open-source-windows-7-product-key-checker.10858/page-14#post-1531837

# new source, work on Windows 8 & up, N key's
# keycheck.py by WitherOrNot
# https://github.com/WitherOrNot/winkeycheck

#>

function Consume-ProductKey {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ProductKey,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Retail', 'OEM', 'Volume', 'Volume:GVLK', 'Volume:MAK')]
        [string]$LicenseType = 'Retail',

        [Parameter(Mandatory = $false)]
        [Guid]$SkuID = [guid]::Empty
    )
    if ([string]::IsNullOrEmpty($ProductKey) -or (
        $ProductKey.LastIndexOf("n",[StringComparison]::InvariantCultureIgnoreCase) -lt 0)) {
    }

    $keyInfo = Decode-Key -Key $ProductKey
    if ($SkuID -eq [guid]::Empty) {
        $SkuId = Retrieve-ProductKeyInfo -CdKey $ProductKey | select -ExpandProperty SkuId
    }
    $LicenseXml = Get-LicenseData -SkuID $SkuID -Mode License
    $LicenseData = [HttpUtility]::HtmlEncode($LicenseXml)
    if (!$SkuId -or !$keyInfo -or !$LicenseData) {
        Clear-Host
        Write-Host
        Write-Host "** Consume process Failure:" -ForegroundColor Red
        Write-host "** Possible Error: Failed to decode product key." -ForegroundColor Green
        Write-host "** Possible Error: SkuId not found for the product key." -ForegroundColor Green
        Write-host "** Possible Error: Failed to Accuire License File for SKU Guid." -ForegroundColor Green
        Write-Host
        return
    }

    [long]$group    = $keyInfo.Group
    [long]$serial   = $keyInfo.Serial
    [long]$security = $keyInfo.Security
    [int32]$upgrade = $keyInfo.Upgrade
    [System.Numerics.BigInteger]$act_hash = [BigInteger]$upgrade -band 1
    $act_hash = $act_hash -bor (([BigInteger]$serial -band ((1L -shl 30) - 1)) -shl 1)
    $act_hash = $act_hash -bor (([BigInteger]$group -band ((1L -shl 20) - 1)) -shl 31)
    $act_hash = $act_hash -bor (([BigInteger]$security -band ((1L -shl 53) - 1)) -shl 51)
    $bytes = $act_hash.ToByteArray()
    $KeyData = New-Object 'Byte[]' 13
    [Array]::Copy($bytes, 0, $KeyData, 0, [Math]::Min(13, $bytes.Length))
    $act_data = [Convert]::ToBase64String($KeyData)

    $Hex = "2A0000000100020001000100000000000000010001000100"
    [byte[]]$Binding = @(
        for ($i=0; $i -lt $Hex.Length; $i+=2) {
            [byte]::Parse($Hex.Substring($i, 2), 'HexNumber')
        }
    )
    [byte[]]$RandomBytes = New-Object byte[] 18
    (New-Object System.Security.Cryptography.RNGCryptoServiceProvider).GetBytes($RandomBytes)
    $bindingData = [System.Convert]::ToBase64String((@($Binding) + @($RandomBytes)))

    $secure_store_id = [guid]::NewGuid()
    $act_config_id = [HttpUtility]::HtmlEncode("msft2009:$SkuId&$act_data")
    $systime = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:sszzz", [System.Globalization.CultureInfo]::InvariantCulture)
    $utctime = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:sszzz", [System.Globalization.CultureInfo]::InvariantCulture)

    $requestXml = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
    xmlns:soapenc="http://schemas.xmlsoap.org/soap/encoding/"
    xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
    xmlns:xsd="http://www.w3.org/2001/XMLSchema"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <soap:Body>
        <RequestSecurityToken
            xmlns="http://schemas.xmlsoap.org/ws/2004/04/security/trust">
            <TokenType>ProductActivation</TokenType>
            <RequestType>http://schemas.xmlsoap.org/ws/2004/04/security/trust/Issue</RequestType>
            <UseKey>
                <Values
                    xmlns:q1="http://schemas.xmlsoap.org/ws/2004/04/security/trust" soapenc:arrayType="q1:TokenEntry[1]">
                    <TokenEntry>
                        <Name>PublishLicense</Name>
                        <Value>$LicenseData</Value>
                    </TokenEntry>
                </Values>
            </UseKey>
            <Claims>
                <Values
                    xmlns:q1="http://schemas.xmlsoap.org/ws/2004/04/security/trust" soapenc:arrayType="q1:TokenEntry[14]">
                    <TokenEntry>
                        <Name>BindingType</Name>
                        <Value>msft:rm/algorithm/hwid/4.0</Value>
                    </TokenEntry>
                    <TokenEntry>
                        <Name>Binding</Name>
                        <Value>$bindingData</Value>
                    </TokenEntry>
                    <TokenEntry>
                        <Name>ProductKey</Name>
                        <Value>$ProductKey</Value>
                    </TokenEntry>
                    <TokenEntry>
                        <Name>ProductKeyType</Name>
                        <Value>msft:rm/algorithm/pkey/2009</Value>
                    </TokenEntry>
                    <TokenEntry>
                        <Name>ProductKeyActConfigId</Name>
                        <Value>$act_config_id</Value>
                    </TokenEntry>
                    <TokenEntry>
                        <Name>otherInfoPublic.licenseCategory</Name>
                        <Value>msft:sl/EUL/ACTIVATED/PUBLIC</Value>
                    </TokenEntry>
                    <TokenEntry>
                        <Name>otherInfoPrivate.licenseCategory</Name>
                        <Value>msft:sl/EUL/ACTIVATED/PRIVATE</Value>
                    </TokenEntry>
                    <TokenEntry>
                        <Name>otherInfoPublic.sysprepAction</Name>
                        <Value>rearm</Value>
                    </TokenEntry>
                    <TokenEntry>
                        <Name>otherInfoPrivate.sysprepAction</Name>
                        <Value>rearm</Value>
                    </TokenEntry>
                    <TokenEntry>
                        <Name>ClientInformation</Name>
                        <Value>SystemUILanguageId=1033;UserUILanguageId=1033;GeoId=244</Value>
                    </TokenEntry>
                    <TokenEntry>
                        <Name>ClientSystemTime</Name>
                        <Value>$systime</Value>
                    </TokenEntry>
                    <TokenEntry>
                        <Name>ClientSystemTimeUtc</Name>
                        <Value>$utctime</Value>
                    </TokenEntry>
                    <TokenEntry>
                        <Name>otherInfoPublic.secureStoreId</Name>
                        <Value>$secure_store_id</Value>
                    </TokenEntry>
                    <TokenEntry>
                        <Name>otherInfoPrivate.secureStoreId</Name>
                        <Value>$secure_store_id</Value>
                    </TokenEntry>
                </Values>
            </Claims>
        </RequestSecurityToken>
    </soap:Body>
</soap:Envelope>
"@

    try {
        $response = $null
        $webRequest = [System.Net.HttpWebRequest]::Create('https://activation.sls.microsoft.com/SLActivateProduct/SLActivateProduct.asmx?configextension=$LicenseType')
        $webRequest.Method      = "POST"
        $webRequest.Accept      = 'text/*'
        $webRequest.UserAgent   = 'SLSSoapClient'
        $webRequest.ContentType = 'text/xml; charset=utf-8'
        $webRequest.Headers.Add("SOAPAction", "http://microsoft.com/SL/ProductActivationService/IssueToken");

        try {
            $byteArray = [System.Text.Encoding]::UTF8.GetBytes($requestXml)
            $webRequest.ContentLength = $byteArray.Length
            $stream = $webRequest.GetRequestStream()
            $stream.Write($byteArray, 0, $byteArray.Length)
            $stream.Close()
            $httpResponse = $webRequest.GetResponse()
            $streamReader = New-Object System.IO.StreamReader($httpResponse.GetResponseStream())
            $response = $streamReader.ReadToEnd()
            $streamReader.Close()
        }
        catch [System.Net.WebException] {
            if ($_.Exception) {
                $stream = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($stream)
                $response = $reader.ReadToEnd().ToString()
                $reader.Close()
            }
        }
        catch {
            Write-Error "Error: $($_.Exception.Message)"
            $global:error = $_
            return $null
        }

    }
    catch {
        Write-Error "Error: $($_.Exception.Message)"
        return $null
    }

    if ($response -ne $null) {
        [xml]$xmlResponse = $response
        if ($xmlResponse.Envelope.Body.Fault -eq $null) {
            return "Valid Key"
        } else {
            return Parse-ErrorMessage -MessageId ($xmlResponse.Envelope.Body.Fault.detail.HRESULT) -Flags ACTIVATION
        }
    }

    return "Error: No response received.", "", $false
}

# oHook part -->
$KeyBlock = @'
SKU,KEY
O365BusinessRetail,Y9NF9-M2QWD-FF6RJ-QJW36-RRF2T
O365EduCloudRetail,W62NQ-267QR-RTF74-PF2MH-JQMTH
O365HomePremRetail,3NMDC-G7C3W-68RGP-CB4MH-4CXCH
O365ProPlusRetail,H8DN8-Y2YP3-CR9JT-DHDR9-C7GP3
O365SmallBusPremRetail,2QCNB-RMDKJ-GC8PB-7QGQV-7QTQJ
O365AppsBasicRetail,3HYJN-9KG99-F8VG9-V3DT8-JFMHV
AccessRetail,WHK4N-YQGHB-XWXCC-G3HYC-6JF94
AccessRuntimeRetail,RNB7V-P48F4-3FYY6-2P3R3-63BQV
ExcelRetail,RKJBN-VWTM2-BDKXX-RKQFD-JTYQ2
HomeBusinessPipcRetail,2WQNF-GBK4B-XVG6F-BBMX7-M4F2Y
HomeBusinessRetail,HM6FM-NVF78-KV9PM-F36B8-D9MXD
HomeStudentARMRetail,PBQPJ-NC22K-69MXD-KWMRF-WFG77
HomeStudentPlusARMRetail,6F2NY-7RTX4-MD9KM-TJ43H-94TBT
HomeStudentRetail,PNPRV-F2627-Q8JVC-3DGR9-WTYRK
HomeStudentVNextRetail,YWD4R-CNKVT-VG8VJ-9333B-RC3B8
MondoRetail,VNWHF-FKFBW-Q2RGD-HYHWF-R3HH2
OneNoteFreeRetail,XYNTG-R96FY-369HX-YFPHY-F9CPM
OneNoteRetail,FXF6F-CNC26-W643C-K6KB7-6XXW3
OutlookRetail,7N4KG-P2QDH-86V9C-DJFVF-369W9
PersonalPipcRetail,9CYB3-NFMRW-YFDG6-XC7TF-BY36J
PersonalRetail,FT7VF-XBN92-HPDJV-RHMBY-6VKBF
PowerPointRetail,N7GCB-WQT7K-QRHWG-TTPYD-7T9XF
ProPlusRetail,GM43N-F742Q-6JDDK-M622J-J8GDV
ProfessionalPipcRetail,CF9DD-6CNW2-BJWJQ-CVCFX-Y7TXD
ProfessionalRetail,NXFTK-YD9Y7-X9MMJ-9BWM6-J2QVH
ProjectProRetail,WPY8N-PDPY4-FC7TF-KMP7P-KWYFY
ProjectStdRetail,NTHQT-VKK6W-BRB87-HV346-Y96W8
PublisherRetail,WKWND-X6G9G-CDMTV-CPGYJ-6MVBF
SkypeServiceBypassRetail,6MDN4-WF3FV-4WH3Q-W699V-RGCMY
SkypeforBusinessEntryRetail,4N4D8-3J7Y3-YYW7C-73HD2-V8RHY
SkypeforBusinessRetail,PBJ79-77NY4-VRGFG-Y8WYC-CKCRC
StandardRetail,2FPWN-4H6CM-KD8QQ-8HCHC-P9XYW
VisioProRetail,NVK2G-2MY4G-7JX2P-7D6F2-VFQBR
VisioStdRetail,NCRB7-VP48F-43FYY-62P3R-367WK
WordRetail,P8K82-NQ7GG-JKY8T-6VHVY-88GGD
Access2019Retail,WRYJ6-G3NP7-7VH94-8X7KP-JB7HC
AccessRuntime2019Retail,FGQNJ-JWJCG-7Q8MG-RMRGJ-9TQVF
Excel2019Retail,KBPNW-64CMM-8KWCB-23F44-8B7HM
HomeBusiness2019Retail,QBN2Y-9B284-9KW78-K48PB-R62YT
HomeStudentARM2019Retail,DJTNY-4HDWM-TDWB2-8PWC2-W2RRT
HomeStudentPlusARM2019Retail,NM8WT-CFHB2-QBGXK-J8W6J-GVK8F
HomeStudent2019Retail,XNWPM-32XQC-Y7QJC-QGGBV-YY7JK
Outlook2019Retail,WR43D-NMWQQ-HCQR2-VKXDR-37B7H
Personal2019Retail,NMBY8-V3CV7-BX6K6-2922Y-43M7T
PowerPoint2019Retail,HN27K-JHJ8R-7T7KK-WJYC3-FM7MM
ProPlus2019Retail,BN4XJ-R9DYY-96W48-YK8DM-MY7PY
Professional2019Retail,9NXDK-MRY98-2VJV8-GF73J-TQ9FK
ProjectPro2019Retail,JDTNC-PP77T-T9H2W-G4J2J-VH8JK
ProjectStd2019Retail,R3JNT-8PBDP-MTWCK-VD2V8-HMKF9
Publisher2019Retail,4QC36-NW3YH-D2Y9D-RJPC7-VVB9D
SkypeforBusiness2019Retail,JBDKF-6NCD6-49K3G-2TV79-BKP73
SkypeforBusinessEntry2019Retail,N9722-BV9H6-WTJTT-FPB93-978MK
Standard2019Retail,NDGVM-MD27H-2XHVC-KDDX2-YKP74
VisioPro2019Retail,2NWVW-QGF4T-9CPMB-WYDQ9-7XP79
VisioStd2019Retail,263WK-3N797-7R437-28BKG-3V8M8
Word2019Retail,JXR8H-NJ3MK-X66W8-78CWD-QRVR2
Access2021Retail,P286B-N3XYP-36QRQ-29CMP-RVX9M
AccessRuntime2021Retail,MNX9D-PB834-VCGY2-K2RW2-2DP3D
Excel2021Retail,V6QFB-7N7G9-PF7W9-M8FQM-MY8G9
HomeBusiness2021Retail,JM99N-4MMD8-DQCGJ-VMYFY-R63YK
HomeStudent2021Retail,N3CWD-38XVH-KRX2Y-YRP74-6RBB2
OneNoteFree2021Retail,CNM3W-V94GB-QJQHH-BDQ3J-33Y8H
OneNote2021Retail,NB2TQ-3Y79C-77C6M-QMY7H-7QY8P
Outlook2021Retail,4NCWR-9V92Y-34VB2-RPTHR-YTGR7
Personal2021Retail,RRRYB-DN749-GCPW4-9H6VK-HCHPT
PowerPoint2021Retail,3KXXQ-PVN2C-8P7YY-HCV88-GVM96
ProPlus2021Retail,8WXTP-MN628-KY44G-VJWCK-C7PCF
Professional2021Retail,DJPHV-NCJV6-GWPT6-K26JX-C7PBG
ProjectPro2021Retail,QKHNX-M9GGH-T3QMW-YPK4Q-QRWMV
ProjectStd2021Retail,2B96V-X9NJY-WFBRC-Q8MP2-7CHRR
Publisher2021Retail,CDNFG-77T8D-VKQJX-B7KT3-KK28V
SkypeforBusiness2021Retail,DVBXN-HFT43-CVPRQ-J89TF-VMMHG
Standard2021Retail,HXNXB-J4JGM-TCF44-2X2CV-FJVVH
VisioPro2021Retail,T6P26-NJVBR-76BK8-WBCDY-TX3BC
VisioStd2021Retail,89NYY-KB93R-7X22F-93QDF-DJ6YM
Word2021Retail,VNCC4-CJQVK-BKX34-77Y8H-CYXMR
Access2024Retail,P6NMW-JMTRC-R6MQ6-HH3F2-BTHKB
Excel2024Retail,82CNJ-W82TW-BY23W-BVJ6W-W48GP
Home2024Retail,N69X7-73KPT-899FD-P8HQ4-QGTP4
HomeBusiness2024Retail,PRKQM-YNPQR-77QT6-328D7-BD223
Outlook2024Retail,2CFK4-N44KG-7XG89-CWDG6-P7P27
PowerPoint2024Retail,CT2KT-GTNWH-9HFGW-J2PWJ-XW7KJ
ProjectPro2024Retail,GNJ6P-Y4RBM-C32WW-2VJKJ-MTHKK
ProjectStd2024Retail,C2PNM-2GQFC-CY3XR-WXCP4-GX3XM
ProPlus2024Retail,VWCNX-7FKBD-FHJYG-XBR4B-88KC6
VisioPro2024Retail,HGRBX-N68QF-6DY8J-CGX4W-XW7KP
VisioStd2024Retail,VBXPJ-38NR3-C4DKF-C8RT7-RGHKQ
Word2024Retail,XN33R-RP676-GMY2F-T3MH7-GCVKR
ExcelVolume,9C2PK-NWTVB-JMPW8-BFT28-7FTBF
Excel2019Volume,TMJWT-YYNMB-3BKTF-644FC-RVXBD
Excel2021Volume,NWG3X-87C9K-TC7YY-BC2G7-G6RVC
Excel2024Volume,F4DYN-89BP2-WQTWJ-GR8YC-CKGJG
PowerPointVolume,J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6
PowerPoint2019Volume,RRNCX-C64HY-W2MM7-MCH9G-TJHMQ
PowerPoint2021Volume,TY7XF-NFRBR-KJ44C-G83KF-GX27K
PowerPoint2024Volume,CW94N-K6GJH-9CTXY-MG2VC-FYCWP
ProPlusVolume,XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
ProPlus2019Volume,NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
ProPlus2021Volume,FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
ProPlus2024Volume,XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB
ProjectProVolume,YG9NW-3K39V-2T3HJ-93F3Q-G83KT
ProjectPro2019Volume,B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
ProjectPro2021Volume,FTNWT-C6WBT-8HMGF-K9PRX-QV9H8
ProjectPro2024Volume,FQQ23-N4YCY-73HQ3-FM9WC-76HF4
ProjectStdVolume,GNFHQ-F6YQM-KQDGJ-327XX-KQBVC
ProjectStd2019Volume,C4F7P-NCP8C-6CQPT-MQHV9-JXD2M
ProjectStd2021Volume,J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T
ProjectStd2024Volume,PD3TT-NTHQQ-VC7CY-MFXK3-G87F8
PublisherVolume,F47MM-N3XJP-TQXJ9-BP99D-8K837
Publisher2019Volume,G2KWX-3NW6P-PY93R-JXK2T-C9Y9V
Publisher2021Volume,2MW9D-N4BXM-9VBPG-Q7W6M-KFBGQ
SkypeforBusinessVolume,869NQ-FJ69K-466HW-QYCP2-DDBV6
SkypeforBusiness2019Volume,NCJ33-JHBBY-HTK98-MYCV8-HMKHJ
SkypeforBusiness2021Volume,HWCXN-K3WBT-WJBKY-R8BD9-XK29P
SkypeforBusiness2024Volume,4NKHF-9HBQF-Q3B6C-7YV34-F64P3
StandardVolume,JNRGM-WHDWX-FJJG3-K47QV-DRTFM
Standard2019Volume,6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK
Standard2021Volume,KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3
Standard2024Volume,V28N4-JG22K-W66P8-VTMGK-H6HGR
VisioProVolume,PD3PC-RHNGV-FXJ29-8JK7D-RJRJK
VisioPro2019Volume,9BGNQ-K37YR-RQHF2-38RQ3-7VCBB
VisioPro2021Volume,KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4
VisioPro2024Volume,B7TN8-FJ8V3-7QYCP-HQPMV-YY89G
VisioStdVolume,7WHWN-4T7MP-G96JF-G33KR-W8GF4
VisioStd2019Volume,7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2
VisioStd2021Volume,MJVNY-BYWPY-CWV6J-2RKRT-4M8QG
VisioStd2024Volume,JMMVY-XFNQC-KK4HK-9H7R3-WQQTV
WordVolume,WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6
accessVolume,GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW
access2019Volume,9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT
access2021Volume,WM8YG-YNGDD-4JHDC-PG3F4-FC4T4
access2024Volume,82FTR-NCHR7-W3944-MGRHM-JMCWD
mondoVolume,HFTND-W9MK4-8B7MJ-B6C4G-XQBR2
outlookVolume,R69KK-NTPKF-7M3Q4-QYBHW-6MT9B
outlook2019Volume,7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK
outlook2021Volume,C9FM6-3N72F-HFJXB-TM3V9-T86R9
outlook2024Volume,D2F8D-N3Q3B-J28PV-X27HD-RJWB9
word2019Volume,PBX3G-NWMT6-Q7XBW-PYJGG-WXD33
word2021Volume,TN8H9-M34D3-Y64V9-TR72V-X79KV
word2024Volume,MQ84N-7VYDM-FXV7C-6K7CC-VFW9J
ProjectProXVolume,WGT24-HCNMF-FQ7XH-6M8K7-DRTW9
ProjectStdVolume,GNFHQ-F6YQM-KQDGJ-327XX-KQBVC
ProjectStdXVolume,D8NRQ-JTYM3-7J2DX-646CT-6836M
VisioProXVolume,69WXN-MBYV6-22PQG-3WGHK-RM6XC
VisioStdVolume,7WHWN-4T7MP-G96JF-G33KR-W8GF4
VisioStdXVolume,NY48V-PPYYH-3F4PX-XJRKJ-W4423
ProPlusSPLA2021Volume,JRJNJ-33M7C-R73X3-P9XF7-R9F6M
StandardSPLA2021Volume,BQWDW-NJ9YF-P7Y79-H6DCT-MKQ9C
'@ | ConvertFrom-Csv
function Install {

# Define registry path and value
$regPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"

# Get ProductReleaseIds registry value and process it
try {
    $productReleaseIds = (Get-ItemProperty -Path $regPath -ea 0).ProductReleaseIds
} catch {
    Write-Host "Error accessing registry: $_"
    return
}
if ($productReleaseIds) {
    
    $productReleaseIds -split ',' | ForEach-Object {
        $productKey = $_.Trim()
        $pKey = $KeyBlock | ? SKU -EQ $productKey | Select -ExpandProperty KEY
        if ($pKey) {
            write-host
            SL-InstallProductKey -Keys ($pKey)
        }
    }
} else {
    Write-Host
    Write-Host "ERROR: ProductReleaseIds not found." -ForegroundColor Red
    return
}

$base64 = @'
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
'@

# Decode the Base64 string back into a byte array
$compressedDllBytes = [Convert]::FromBase64String($base64)

# Step 2: Create a MemoryStream from the byte array
$compressedStream = [MemoryStream]::new($compressedDllBytes)

# Step 3: Decompress the byte array using GZip
$gzipStream = New-Object Compression.GZipStream($compressedStream, [Compression.CompressionMode]::Decompress)
$decompressedStream = New-Object MemoryStream

# Copy the decompressed data to a new MemoryStream
$gzipStream.CopyTo($decompressedStream)
$gzipStream.Close()

# Step 4: Get the decompressed byte array
$decompressedDllBytes = $decompressedStream.ToArray()
    
# Define the output file path with ProgramFiles environment variable
$sppc = [Path]::Combine($env:ProgramFiles, "Microsoft Office\root\vfs\System\sppc.dll")

# Define paths for symbolic link and target
$sppcs = [Path]::Combine($env:ProgramFiles, "Microsoft Office\root\vfs\System\sppcs.dll")
$System32 = [Path]::Combine($env:windir, "System32\sppc.dll")

# Step 1: Check if the symbolic link exists and remove it if necessary
if (Test-Path -Path $sppcs) {
    Write-Host "Symbolic link already exists at $sppcs. Attempting to remove..."

    try {
        # Remove the existing symbolic link
        Remove-Item -Path $sppcs -Force
        Write-Host "Existing symbolic link removed successfully."
    } catch {
        Write-Host "Failed to remove existing symbolic link: $_"
    }
} else {
    Write-Host "No symbolic link found at $sppcs."
}

try {
    # Attempt to write byte array to the file
    [System.IO.File]::WriteAllBytes($sppc, $decompressedDllBytes)
    Write-Host "Byte array written successfully to $sppc."
} 
catch {
    Write-Host "Failed to write byte array to ${sppc}: $_"

    # Inner try-catch to handle the case where the file is in use
    try {
        Write-Host "File is in use or locked. Attempting to move it to a temp file..."

        # Generate a random name for the temporary file in the temp folder
        $tempDir = [Path]::GetTempPath()
        $tempFileName = [Path]::Combine($tempDir, [Guid]::NewGuid().ToString() + ".bak")

        # Move the file to the temp location with a random name
        Move-Item -Path $sppc -Destination $tempFileName -Force
        Write-Host "Moved file to temporary location: $tempFileName"

        # Retry the write operation after moving the file
        [System.IO.File]::WriteAllBytes($sppc, $decompressedDllBytes)
        Write-Host "Byte array written successfully to $sppc after moving the file."

    } catch {
        Write-Host "Failed to move the file or retry the write operation: $_"
    }
}

# Step 3: Check if the symbolic link exists and create it if necessary
try {
    if (-not (Test-Path -Path $sppcs)) {
        # Create symbolic link only if it doesn't already exist
        New-Item -Path $sppcs -ItemType SymbolicLink -Target $System32 | Out-Null
        Write-Host "Symbolic link created successfully at $sppcs."
    } else {
        Write-Host "Symbolic link already exists at $sppcs."
    }
} catch {
    Write-Host "Failed to create symbolic link at ${sppcs}: $_"
}
	
# Define the target registry key path
$RegPath = "HKCU:\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency"

# Define the value name and data
$ValueName = "TimeOfLastHeartbeatFailure"
$ValueData = "2040-01-01T00:00:00Z"

# Check if the registry key exists. If not, create it.
# The -Force parameter on New-Item ensures the full path is created if necessary.
if (-not (Test-Path -Path $RegPath)) {
	Write-host "Registry key '$RegPath' not found. Creating it."
	# Use -Force to create the key and any missing parent keys
	# Out-Null is used to suppress the output object from New-Item
	New-Item -Path $RegPath -Force | Out-Null
}

# Set the registry value within the existing (or newly created) key.
# The -Force parameter on Set-ItemProperty ensures the value is created if it doesn't exist
# or updated if it does exist.
Write-host "Setting registry value '$ValueName' at '$RegPath'."
Set-ItemProperty -Path $RegPath -Name $ValueName -Value $ValueData -Type String -Force
}
function Remove {
# Remove the symbolic link if it exists
$sppcs = [Path]::Combine($env:ProgramFiles, "Microsoft Office\root\vfs\System\sppcs.dll")
if (Test-Path -Path $sppcs) {
    try {
        Remove-Item -Path $sppcs -Force
        Write-Host "Symbolic link '$sppcs' removed successfully."
    } catch {
        Write-Host "Failed to remove symbolic link '$sppcs': $_"
    }
} else {
    Write-Host "No symbolic link found at '$sppcs'."
}

# Remove the actual DLL file if it exists
$sppc = [Path]::Combine($env:ProgramFiles, "Microsoft Office\root\vfs\System\sppc.dll")
if (Test-Path -Path $sppc) {
    try {
        # Try to remove the file and handle any errors if they occur
        Remove-Item -Path $sppc -Force -ErrorAction Stop
        Write-Host "DLL file '$sppc' removed successfully."
    } catch {
        Write-Host "Failed to remove DLL file '$sppc': $_"
            
        # If removal failed, try to move the file to a temporary location
        try {
            # Generate a random name for the file in the temp directory
            $tempDir = [Path]::GetTempPath()
            $tempFileName = [Path]::Combine($tempDir, [Guid]::NewGuid().ToString() + ".bak")
            
            # Attempt to move the file to the temp directory with a random name
            Move-Item -Path $sppc -Destination $tempFileName -Force -ErrorAction Stop
            Write-Host "DLL file moved to Temp folder."
        } catch {
            Write-Host "Failed to move DLL file '$sppc' to temporary location: $_"
        }
    }
} else {
    Write-Host "No DLL file found at '$sppc'."
}
}
# oHook Part -->

# XML Parser Function's Part -->

$prefixes = [ordered]@{
  VisualStudio = 'ns1:'
  Office = 'pkc:'
  None = '' # keep last.
}
$validConfigTypes = @(
  'Office',
  'Windows',
  'VisualStudio'
)
Class KeyRange {
  [string] $RefActConfigId
  [string] $PartNumber
  [string] $EulaType
  [bool]   $IsValid
  [int]    $Start
  [int]    $End
}
Class Range {
  [KeyRange[]] $Ranges

  [String] ToString() {
    $output = $null
    $this.Ranges | Sort-Object -Property @{Expression = "Start"; Descending = $false } | Select-First 1 | % { 
      $keyInfo = $_ -as [KeyRange]
      $output += "[$($keyInfo.Start), $($keyInfo.End)], Number:$($keyInfo.PartNumber), Type:$($keyInfo.EulaType), IsValid:$($keyInfo.IsValid)`n" }
    return $output
  }
}
Class INFO_DATA {
  [string] $ActConfigId
  [int]    $RefGroupId
  [string] $EditionId
  [string] $ProductDescription
  [string] $ProductKeyType
  [bool]   $IsRandomized
  [Range]  $KeyRanges
  [string] $ProductKey
  [string] $Command

  [String] ToString() {
    $output = $null
    $GroupId = [String]::Format("{0:X}",$this.RefGroupId)
    $output  = " Ref: $($this.RefGroupId)`nType: $($this.ProductKeyType)`nEdit: $($this.EditionId)`n  ID: $($this.ActConfigId)`nName: $($this.ProductDescription)`n"
    $output += " Gen: (gwmi SoftwareLicensingService).InstallProductKey((KeyInfo $($GroupId) 0 0))`n"
    ($this.KeyRanges).Ranges | Sort-Object -Property @{Expression = "Start"; Descending = $false } | % { 
      $keyInfo = $_ -as [KeyRange]
      $output += "* Key Range: [$($keyInfo.Start)] => [$($keyInfo.End)], Number: $($keyInfo.PartNumber), Type: $($keyInfo.EulaType), IsValid: $($keyInfo.IsValid)`n" }
    return $output
  }
}
Function GenerateConfigList {
    param (
        [ValidateNotNullOrEmpty()]
        [Parameter(ValueFromPipeline)]
        [string] $pkeyconfig = "$env:windir\System32\spp\tokens\pkeyconfig\pkeyconfig.xrm-ms",

        [Parameter(Mandatory=$false)]
        [bool] $IgnoreAPI = $false,

        [Parameter(Mandatory=$false)]
        [bool] $SkipKey = $false,

        [Parameter(Mandatory=$false)]
        [bool] $SkipKeyRange = $false
    )

    function Get-XmlValue {
        param (
            [string]$Source,
            [string]$TagName
        )
    
        $startTag = "<$TagName>"
        $endTag = "</$TagName>"
        $iStart = $Source.IndexOf($startTag) + $startTag.Length
        $iEnd = $Source.IndexOf($endTag) - $iStart
        if ($iStart -ge 0 -and $iEnd -ge 0) {
            return $Source.Substring($iStart, $iEnd)
        }
        Write-Debug $TagName
        return $null
    }
    function Get-XmlSection {
        param (
            [string] $XmlContent,
            [string] $StartTag,
            [string] $EndTag,
            [string] $Delimiter
        )
    
        $iStart = $XmlContent.IndexOf($StartTag) + $StartTag.Length
        $iEnd = $XmlContent.IndexOf($EndTag)
        if ($iStart -ge $iEnd -or $iEnd -lt 0) {
            return @()
        }

        $length = $iEnd - $iStart
        $section = $XmlContent.Substring($iStart, $length)
        return ($section -split $Delimiter)
    }
    function Get-ConfigTags {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true, Position = 0, ParameterSetName = 'FromContent', HelpMessage = "Content source string.")]
            [ValidateNotNullOrEmpty()]
            [string]$Source,

            [Parameter(Mandatory = $true, Position = 0, ParameterSetName = 'FromType', HelpMessage = "Type of configuration.")]
            [ValidateScript({
                if ($_ -notin $validConfigTypes) {
                    throw "ERROR: Invalid ConfigType '$_'."
                }
                return $true
            })]
            [string]$ConfigType
        )

        # Initialize prefix
        $prefix = $prefixes.None

        # Determine the prefix based on the parameter set
        switch ($PSCmdlet.ParameterSetName) {
            'FromType' {
                $prefix = $prefixes[$ConfigType]
            }
            'FromContent' {
                if ($Source -match "r:grant|sl:policy") {
                    throw "ERROR: Source must contain valid content."
                }

                # Dynamically check for patterns, excluding None
                foreach ($key in $prefixes.Keys.Where({ $_ -ne 'None' })) {
                    if ($Source -match "$($prefixes[$key])ActConfigId") {
                        $prefix = $prefixes[$key]
                        break  # Exit loop on first match
                    }
                }
            }
        }

        # Create base tags using the determined prefix
        $baseConfigTag = "${prefix}Configuration"
        $baseKeyRangeTag = "${prefix}KeyRange"

        # Build and return XML tags as a custom object
        return [PSCustomObject]@{
            StartTagConfig    = "<${baseConfigTag}s>"
            EndTagConfig      = "</${baseConfigTag}s>"
            DelimiterConfig    = "<${baseConfigTag}>"
            StartTagKeyRange  = "<${baseKeyRangeTag}s>"
            EndTagKeyRange    = "</${baseKeyRangeTag}s>"
            DelimiterKeyRange  = "<${baseKeyRangeTag}>"
            TagPrefix          = $prefix
        }
    }

    if (-not [IO.FILE]::Exists($pkeyconfig)) {
        throw "ERROR: File not exist" }

    $data = Get-Content -Path $pkeyconfig
    $iStart = $data.IndexOf('<tm:infoBin name="pkeyConfigData">')
    if ($iStart -le 0) {
        throw "ERROR: FILE NOT SUPPORTED" }

    $iEnd = $data.Substring($iStart+34).IndexOf('</tm:infoBin>')
    $Conf = [Encoding]::UTF8.GetString(
      [Convert]::FromBase64String(
        $data.Substring(($iStart+34), $iEnd)))

    # Get configuration based on ConfigType
    $Config = Get-ConfigTags -Source $Conf

    # Process Configurations
    $Output = @{}
    $Configurations = Get-XmlSection -XmlContent $Conf -StartTag $Config.StartTagConfig -EndTag $Config.EndTagConfig -Delimiter $Config.DelimiterConfig
    $KeyRanges = Get-XmlSection -XmlContent $Conf -StartTag $Config.StartTagKeyRange -EndTag $Config.EndTagKeyRange -Delimiter $Config.DelimiterKeyRange

    $Configurations | ForEach-Object {
        
        try {
          $length = 0
          $Source = $_ | Out-String
          $length = $Source.Length
        }
        catch {
          # just in case of
          $length = 0
        }

        $ActConfigId = $null
        $RefGroupId = $null
        $EditionId = $null
        $ProductDescription = $null
        $ProductKeyType = $null
        $IsRandomized = $null

        if ($length -ge 5) {
          $ActConfigId = Get-XmlValue -Source $Source -TagName "$($Config.TagPrefix)ActConfigId"
          $RefGroupId = Get-XmlValue -Source $Source -TagName "$($Config.TagPrefix)RefGroupId"
          $EditionId = Get-XmlValue -Source $Source -TagName "$($Config.TagPrefix)EditionId"
          $ProductDescription = Get-XmlValue -Source $Source -TagName "$($Config.TagPrefix)ProductDescription"
          $ProductKeyType = Get-XmlValue -Source $Source -TagName "$($Config.TagPrefix)ProductKeyType"
          $IsRandomized = Get-XmlValue -Source $Source -TagName "$($Config.TagPrefix)IsRandomized" -as [BOOL]
        }

        if ($ActConfigId) {
            $cInfo = [INFO_DATA]::new()
            $cInfo.ActConfigId = $ActConfigId
            $cInfo.IsRandomized = $IsRandomized
            $cInfo.ProductDescription = $ProductDescription
            $cInfo.RefGroupId = $RefGroupId
            $cInfo.ProductKeyType = $ProductKeyType
            $cInfo.EditionId = $EditionId

            <#
            # Attempt to set ProductKey from the reference array
            if (-not $SkipKey) {
			    if ($Config.TagPrefix -and ($Config.TagPrefix -eq 'pkc:')) {
				    $cInfo.ProductKey = $OfficeOnlyKeys[[int]$RefGroupId]
			    } elseif ($Config.TagPrefix -and ($Config.TagPrefix -eq 'ns1:')){
                    $cInfo.ProductKey = ($VSOnlyKeys | ? { $_.Key -eq [int]$RefGroupId } | Get-Random -Count 1).Value
			    } else {
                    $cInfo.ProductKey = ($KeysRef | ? { $_.Key -eq [int]$RefGroupId } | Get-Random -Count 1).Value
                }
            }
            #>
            
            # Check if ProductKey is empty
            if (-not $SkipKey -and ([STRING]::IsNullOrEmpty($cInfo.ProductKey) -and ($RefGroupId -ne '999999'))) {
                # Check if ProductKeyType matches one of the groups
                if (![string]::IsNullOrEmpty($Config.TagPrefix)) {
                    # CASE OF --> Office & VS
                    $IgnoreAPI = $true }

                if (-not $IgnoreAPI -and ($groups -contains $ProductKeyType)) {
                    
                    # i don't think i need it any longer,
                    # since i extract all key's from pkhelper.dll

                    $value = Get-ProductKeys -EditionID $EditionId -ProductKeyType $ProductKeyType
                    if ($value) {
                        # Set ProductKey based on the result of Get-ProductKeys
                        $RefInfo = $null
                        try { $RefInfo = KeyDecode -key0 $value.ProductKey}
                        catch {}
                        if ($value -and $RefInfo -and ($RefInfo[2].Value -match $RefGroupId)) {
                          $cInfo.ProductKey = $value.ProductKey }}}}

            # Call Encode-Key only if ProductKey is still empty after the checks
            if (-not $SkipKey -and ([STRING]::IsNullOrEmpty($cInfo.ProductKey)-and ($RefGroupId -ne '999999'))) {
                $cInfo.ProductKey = Encode-Key $RefGroupId 0 0
            }
            # Call LibTSForge Generate key function only if ProductKey is still empty after the checks
            if (-not $SkipKey -and ([STRING]::IsNullOrEmpty($cInfo.ProductKey)-and ($RefGroupId -ne '999999'))) {
              $cInfo.ProductKey = GetRandomKey -ProductID (
                ([GUID]::Parse($cInfo.ActConfigId)).ToString())
            }

            $cInfo.Command = "(gwmi SoftwareLicensingService).InstallProductKey(""$($cInfo.ProductKey)"")"
            $cInfo.KeyRanges = [Range]::new()
            $Output[$ActConfigId] = $cInfo
        }
    }

    # Process Key Ranges
    if ($KeyRanges -and (-not $SkipKeyRange)) {
        $KeyRanges | ForEach-Object {
            $Source = $_ | Out-String
            $RefActConfigId = Get-XmlValue -Source $Source -TagName "$($Config.TagPrefix)RefActConfigId"
            $PartNumber = Get-XmlValue -Source $Source -TagName "$($Config.TagPrefix)PartNumber"
            $EulaType = Get-XmlValue -Source $Source -TagName "$($Config.TagPrefix)EulaType"
            $IsValid = Get-XmlValue -Source $Source -TagName "$($Config.TagPrefix)IsValid" -as [BOOL]
            $Start = Get-XmlValue -Source $Source -TagName "$($Config.TagPrefix)Start" -as [INT]
            $End = Get-XmlValue -Source $Source -TagName "$($Config.TagPrefix)End" -as [INT]

            if ($RefActConfigId) {
                $kRange = [KeyRange]::new()
                $kRange.End = $End
                $kRange.Start = $Start
                $kRange.IsValid = $IsValid
                $kRange.EulaType = $EulaType
                $kRange.PartNumber = $PartNumber
                $kRange.RefActConfigId = $RefActConfigId

                if ($Output[$RefActConfigId]) {
                    $cInfo = $Output[$RefActConfigId] -as [INFO_DATA]
                    $iInfo = $cInfo.KeyRanges -as [Range]
                    $iInfo.Ranges += $kRange
                }
            }
        }
    }

    return $Output.Values
}

# XML Parser Function's Part -->

# Start #
# License Info
# Begin #

# License Info Part -->

<#
TSforge
https://github.com/massgravel/TSforge

Open-source slc.dll patch for Windows 8 Milestone builds (7850, 795x, 7989)
Useful if you want to enable things such as Modern Task Manager, Ribbon Explorer, etc.
https://github.com/LBBNetwork/openredpill/blob/master/private.c

Open-source slc.dll patch for Windows 8 Milestone builds (7850, 795x, 7989)
https://github.com/LBBNetwork/openredpill/blob/master/slpublic.h

slpublic.h header
https://learn.microsoft.com/en-us/windows/win32/api/slpublic/
#>

<#
.SYNOPSIS
Opens or closes the global SLC handle,
and optionally closes a specified $hSLC handle.

#>
function Manage-SLHandle {
    [CmdletBinding()]
    param(
        [IntPtr]$hSLC = [IntPtr]::Zero,
        [switch]$Create,
        [switch]$Release,
        [switch]$Force
    )

    # Initialize global variables
    if (-not $global:Status_)      { $global:Status_ = 0 }
    if (-not $global:hSLC_)        { $global:hSLC_   = [IntPtr]::Zero }
    if (-not $global:TrackedSLCs)  { $global:TrackedSLCs = [System.Collections.Generic.HashSet[IntPtr]]::new() }
    if (-not $global:SLC_Lock)        { $global:SLC_Lock = New-Object Object }


    # Helper: Check if handle is tracked
    function Is-HandleTracked([IntPtr]$handle) {
        return $global:TrackedSLCs.Contains($handle)
    }

    # Create new handle
    [System.Threading.Monitor]::Enter($global:SLC_Lock)
    try {
        if ($Create) {
            $newHandle = [IntPtr]::Zero
            $hr = $Global:SLC::SLOpen([ref]$newHandle)
            if ($hr -ne 0) {
                throw "SLOpen failed with HRESULT 0x{0:X8}" -f $hr
            }
            $global:TrackedSLCs.Add($newHandle) | Out-Null
            Write-Verbose "New handle created and tracked."
            return $newHandle
        }

        # Release handle
        if ($Release) {
            # Release specific handle if valid
            if ($hSLC -and $hSLC -ne [IntPtr]::Zero) {
                if (-not (Is-HandleTracked $hSLC) -and -not $Force) {
                    Write-Warning "Handle not tracked or already released. Use -Force to override."
                    return
                }
                Write-Verbose "Releasing specified handle."
                Free-IntPtr -handle $hSLC -Method License
                $global:TrackedSLCs.Remove($hSLC) | Out-Null
                return $hr
            }

            # Release global handle
            if ($global:Status_ -eq 0 -and -not $Force) {
                Write-Warning "Global handle already closed. Use -Force to override."
                return
            }

            Write-Verbose "Releasing global handle."
            Free-IntPtr -handle $hSLC_ -Method License
            $global:TrackedSLCs.Remove($global:hSLC_) | Out-Null
            $global:hSLC_ = [IntPtr]::Zero
            $global:Status_ = 0
            return $hr
        }

        # Return existing global handle if already open
        if ($global:Status_ -eq 1 -and $global:hSLC_ -ne [IntPtr]::Zero -and -not $Force) {
            Write-Verbose "Returning existing global handle."
            return $global:hSLC_
        }

        # Open or reopen global handle
        if ($Force -and $global:hSLC_ -ne [IntPtr]::Zero) {
            Write-Verbose "Force-closing previously open global handle."
            Free-IntPtr -handle $hSLC_ -Method License
            $global:TrackedSLCs.Remove($global:hSLC_) | Out-Null
        }

        Write-Verbose "Opening new global handle."
        $global:hSLC_ = [IntPtr]::Zero
        $hr = $Global:SLC::SLOpen([ref]$global:hSLC_)
        if ($hr -ne 0) {
            throw "SLOpen failed with HRESULT 0x{0:X8}" -f $hr
        }
        $global:TrackedSLCs.Add($global:hSLC_) | Out-Null
        $global:Status_ = 1
        return $global:hSLC_
    }
    finally {
        [System.Threading.Monitor]::Exit($global:SLC_Lock)
    }
}

<#
typedef enum _tagSLDATATYPE {

SL_DATA_NONE = REG_NONE,      // 0
SL_DATA_SZ = REG_SZ,          // 1
SL_DATA_DWORD = REG_DWORD,    // 4
SL_DATA_BINARY = REG_BINARY,  // 3
SL_DATA_MULTI_SZ,             // 7
SL_DATA_SUM = 100             // 100

} SLDATATYPE;

#define REG_NONE		0	/* no type */
#define REG_SZ			1	/* string type (ASCII) */
#define REG_EXPAND_SZ	2	/* string, includes %ENVVAR% (expanded by caller) (ASCII) */
#define REG_BINARY		3	/* binary format, callerspecific */
#define REG_DWORD		4	/* DWORD in little endian format */
#define REG_DWORD_LITTLE_ENDIAN	4	/* DWORD in little endian format */
#define REG_DWORD_BIG_ENDIAN	5	/* DWORD in big endian format  */
#define REG_LINK		6	/* symbolic link (UNICODE) */
#define REG_MULTI_SZ	7	/* multiple strings, delimited by \0, terminated by \0\0 (ASCII) */
#define REG_RESOURCE_LIST	8	/* resource list? huh? */
#define REG_FULL_RESOURCE_DESCRIPTOR	9	/* full resource descriptor? huh? */
#define REG_RESOURCE_REQUIREMENTS_LIST	10
#define REG_QWORD		11	/* QWORD in little endian format */
#>
$SLDATATYPE = @{
    SL_DATA_NONE       = 0   # REG_NONE
    SL_DATA_SZ         = 1   # REG_SZ
    SL_DATA_DWORD      = 4   # REG_DWORD
    SL_DATA_BINARY     = 3   # REG_BINARY
    SL_DATA_MULTI_SZ   = 7   # REG_MULTI_SZ
    SL_DATA_SUM        = 100 # Custom value
}
function Parse-RegistryData {
    param (
        # Data type (e.g., $SLDATATYPE.SL_DATA_NONE, $SLDATATYPE.SL_DATA_SZ, etc.)
        [Parameter(Mandatory=$true)]
        [int]$dataType,

        # Pointer to the data (e.g., registry value pointer)
        [Parameter(Mandatory=$false)]
        [IntPtr]$ptr,

        # Size of the data (in bytes)
        [Parameter(Mandatory=$true)]
        [int]$valueSize,

        # Optional, for special cases (e.g., ProductSkuId)
        [Parameter(Mandatory=$false)]
        [string]$valueName,

        [Parameter(Mandatory=$false)]
        [byte[]]$blob,

        [Parameter(Mandatory=$false)]
        [int]$dataOffset = 0
    )

    # Treat IntPtr.Zero as null for XOR logic
    $ptrIsSet = ($ptr -ne [IntPtr]::Zero) -and ($ptr -ne $null)
    $blobIsSet = ($blob -ne $null)

    if (-not ($ptrIsSet -xor $blobIsSet)) {
        Write-Warning "Exactly one of 'ptr' or 'blob' must be provided, not both or neither."
        return $null
    }

    if ($valueSize -le 0) {
        Write-Warning "Data size is zero or negative for valueName '$valueName'. Returning null."
        return $null
    }

    if ($blobIsSet) {
        if ($dataOffset -lt 0 -or ($dataOffset + $valueSize) -gt $blob.Length) {
            Write-Warning "Invalid dataOffset ($dataOffset) or valueSize ($valueSize) exceeds blob length ($($blob.Length)) for valueName '$valueName'. Returning null."
            return $null
        }
    }

    $result = $null

    $uint32Names = @(
        'SL_LAST_ACT_ATTEMPT_HRESULT',
        'SL_LAST_ACT_ATTEMPT_SERVER_FLAGS',
        'Security-SPP-LastWindowsActivationHResult'
    )

    $datetimeNames = @(
        'SL_LAST_ACT_ATTEMPT_TIME',
        'EvaluationEndDate',
        'TrustedTime',
        'Security-SPP-LastWindowsActivationTime'
    )

    switch ($dataType) {
        $SLDATATYPE.SL_DATA_NONE { 
            $result = $null 
        }

        $SLDATATYPE.SL_DATA_SZ {
            # SL_DATA_SZ = Unicode string
            if ($ptr) {
                # PtrToStringUni expects length in characters, valueSize is in bytes, so divide by 2
                $result = [Marshal]::PtrToStringUni($ptr, $valueSize / 2).TrimEnd([char]0)
            }
            else {
                $buffer = New-Object byte[] $valueSize
                [Buffer]::BlockCopy($blob, $dataOffset, $buffer, 0, $valueSize)
                $result = [Encoding]::Unicode.GetString($buffer).TrimEnd([char]0)
            }
        }

        $SLDATATYPE.SL_DATA_DWORD {
            # SL_DATA_DWORD = DWORD (4 bytes)
            if ($valueSize -ne 4) {
                $result = $null
            }
            elseif ($ptr) {
                # Allocate 4-byte array
                $bytes = New-Object byte[] 4
                [Marshal]::Copy($ptr, $bytes, 0, 4)
                $result = [BitConverter]::ToInt32($bytes, 0)    # instead ToUInt32
            }
            else {
                $buffer = New-Object byte[] $valueSize
                [Buffer]::BlockCopy($blob, $dataOffset, $buffer, 0, $valueSize)
                $result = [BitConverter]::ToInt32($buffer, 0)  # instead ToUInt32
            }
        }

        $SLDATATYPE.SL_DATA_BINARY {
            # SL_DATA_BINARY = Binary blob
            if ($valueName -eq 'ProductSkuId' -and $valueSize -eq 16) {
                # If it's ProductSkuId and the buffer is 16 bytes, treat it as a GUID
                $bytes = New-Object byte[] 16
                if ($ptr) {
                    [Marshal]::Copy($ptr, $bytes, 0, 16)
                }
                else {
                    [Buffer]::BlockCopy($blob, $dataOffset, $bytes, 0, $valueSize)
                }
                $result = [Guid]::new($bytes)
            }
            elseif ($datetimeNames -contains $valueName -and $valueSize -eq 8) {
                $bytes = New-Object byte[] 8
                if ($ptr) {
                    [Marshal]::Copy($ptr, $bytes, 0, 8)
                }
                else {
                    [Buffer]::BlockCopy($blob, $dataOffset, $bytes, 0, 8)
                }
                $fileTime = [BitConverter]::ToInt64($bytes, 0)
                $result = [DateTime]::FromFileTimeUtc($fileTime)
            }
            elseif ($uint32Names -contains $valueName -and $valueSize -eq 4) {
                $bytes = New-Object byte[] 4
                if ($ptr) {
                    [Marshal]::Copy($ptr, $bytes, 0, 4)
                }
                else {
                    [Buffer]::BlockCopy($blob, $dataOffset, $bytes, 0, 4)
                }
                $result = [BitConverter]::ToInt32($bytes, 0) # instead ToUInt32
            }
            else {
                # Otherwise, just copy the binary data
                $result = New-Object byte[] $valueSize
                if ($ptr) {
                    [Marshal]::Copy($ptr, $result, 0, $valueSize)
                    $result = ($result | ForEach-Object { $_.ToString("X2") }) -join "-"
                }
                else {
                    [Buffer]::BlockCopy($blob, $dataOffset, $result, 0, $valueSize)
                }
            }
        }

        $SLDATATYPE.SL_DATA_MULTI_SZ {
            # SL_DATA_MULTI_SZ = Multi-string
            if ($ptr) {
               $raw = [Marshal]::PtrToStringUni($ptr, $valueSize / 2)
               $result = $raw -split "`0" | Where-Object { $_ -ne '' }
            }
            else {
               $buffer = New-Object byte[] $valueSize
               [Buffer]::BlockCopy($blob, $dataOffset, $buffer, 0, $valueSize)
               $raw = [Encoding]::Unicode.GetString($buffer)
               $result = $raw -split "`0" | Where-Object { $_ -ne '' }
            }
        }

        $SLDATATYPE.SL_DATA_SUM { # SL_DATA_SUM = Custom (100)
            # Handle this case accordingly (based on your logic)
            $result = $null
        }

        default {
            # Return null for any unsupported data types
            $result = $null
        }
    }

    return $result
}

<#
Check if a specific Sku is token based edition
#>
Function IsTokenBasedEdition {
    param (
        [Parameter(Mandatory=$false)]
        [GUID]$SkuId,

        [Parameter(Mandatory=$false)]
        [GUID]$LicenseFileId,

        [Parameter(Mandatory=$false)]
        [Intptr]$hSLC = [IntPtr]::Zero
    )

    if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
        $hSLC = if ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
            $global:hSLC_
        } else {
            Manage-SLHandle
        }
    }

    try {
        $closeHandle = $true
        if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
            $hr = $Global:SLC::SLOpen([ref]$hSLC)
            if ($hr -ne 0) {
                throw "SLOpen failed: HRESULT 0x{0:X8}" -f $hr
            }
        } else {
            $closeHandle = $false
        }
    }
    catch {
    }

    try {
        if ((-not $LicenseFileId -and -not $SkuId) -or (
            $LicenseFileId -and $SkuId)) {
                throw "Not a valid choice."
        }

        [Guid]$LicenseFile = [guid]::Empty

        if ($SkuId) {
            $LicenseFile = Retrieve-SKUInfo -SkuId $SkuId -eReturnIdType SL_ID_LICENSE_FILE
        }
        else {
            $LicenseFile = $LicenseFileId
        }

        [IntPtr]$TokenActivationGrants = [IntPtr]::Zero
        if ($LicenseFile -ne ([guid]::empty)) {
            $hrsults = $Global:slc::SLGetTokenActivationGrants(
                $hSLC, [ref]$LicenseFile, [ref]$TokenActivationGrants
            )
                    
            if ($hrsults -ne 0) {
                $errorMessege = Parse-ErrorMessage -MessageId $hrsults -Flags ([ErrorMessageType]::ACTIVATION -bor [ErrorMessageType]::HRESULT)
                Write-Warning "$($hrsults): $($errorMessege)"
                $result = $false
            }
            else {
                $null = $Global:slc::SLFreeTokenActivationGrants(
                    $TokenActivationGrants)
                $result = $true
            }

            return $result
        }
        throw "cant parse GUID"
    }
    finally {
        if ($closeHandle) {
            Write-Warning "Consider Open handle Using Manage-SLHandle"
            Free-IntPtr -handle $hSLC -Method License
        }
    }
}

<#
.SYNOPSIS

$fileId = '?'
$LicenseId = '?'
$OfficeAppId  = '0ff1ce15-a989-479d-af46-f275c6370663'
$windowsAppID  = '55c92734-d682-4d71-983e-d6ec3f16059f'
$enterprisesn = '7103a333-b8c8-49cc-93ce-d37c09687f92'

# should return $OfficeAppId & $windowsAppID
Write-Warning 'Get all installed application IDs.'
Get-SLIDList -eQueryIdType SL_ID_APPLICATION -eReturnIdType SL_ID_APPLICATION
Read-Host

# should return All Office & windows installed SKU
Write-Warning 'Get all installed product SKU IDs.'
Get-SLIDList -eQueryIdType SL_ID_PRODUCT_SKU -eReturnIdType SL_ID_PRODUCT_SKU
Read-Host

# should return $SKU per group <Office -or windows>
Write-Warning 'Get SKU IDs according to the input application ID.'
Get-SLIDList -eQueryIdType SL_ID_APPLICATION -eReturnIdType SL_ID_PRODUCT_SKU -pQueryId $OfficeAppId
Get-SLIDList -eQueryIdType SL_ID_APPLICATION -eReturnIdType SL_ID_PRODUCT_SKU -pQueryId $windowsAppID
Read-Host

# should return $windowsAppID or $OfficeAppId
Write-Warning 'Get application IDs according to the input SKU ID.'
Get-SLIDList -eQueryIdType SL_ID_PRODUCT_SKU -eReturnIdType SL_ID_APPLICATION -pQueryId $enterprisesn
Read-Host

# Same As SLGetInstalledProductKeyIds >> SL_ID_PKEY >> SLGetPKeyInformation >> BLOB
Write-Warning 'Get license PKey IDs according to the input SKU ID.'
Get-SLIDList -eQueryIdType SL_ID_PRODUCT_SKU -eReturnIdType SL_ID_PKEY -pQueryId $enterprisesn 
Read-Host

Write-Warning 'Get license file Ids according to the input SKU ID.'
Get-SLIDList -eQueryIdType SL_ID_PRODUCT_SKU -eReturnIdType SL_ID_LICENSE_FILE -pQueryId $enterprisesn 
Read-Host

Write-Warning 'Get license IDs according to the input license file ID.'
Get-SLIDList -eQueryIdType SL_ID_LICENSE_FILE -eReturnIdType SL_ID_LICENSE -pQueryId $fileId 
Read-Host

Write-Warning 'Get license file ID according to the input license ID.'
Get-SLIDList -eQueryIdType SL_ID_LICENSE -pQueryId $LicenseId -eReturnIdType SL_ID_LICENSE_FILE
Read-Host

Write-Warning 'Get License File Id according to the input License Id'
Get-SLIDList -eQueryIdType SL_ID_LICENSE -pQueryId $LicenseId -eReturnIdType SL_ID_LICENSE_FILE
Read-Host

write-warning "Get union of all application IDs or SKU IDs from all grants of a token activation license."
write-warning "Returns SL_E_NOT_SUPPORTED if the license ID is valid but doesn't refer to a token activation license."
Get-SLIDList -eQueryIdType SL_ID_LICENSE -pQueryId $LicenseId -eReturnIdType SL_ID_APPLICATION

write-warning "Get union of all application IDs or SKU IDs from all grants of a token activation license."
write-warning "Returns SL_E_NOT_SUPPORTED if the license ID is valid but doesn't refer to a token activation license."
Get-SLIDList -eQueryIdType SL_ID_LICENSE -pQueryId $LicenseId -eReturnIdType SL_ID_PRODUCT_SKU

# SLUninstallLicense >> [in] const SLID *pLicenseFileId
Write-Warning 'Get License File IDs associated with a specific Application ID:'
Get-SLIDList -eQueryIdType SL_ID_APPLICATION -pQueryId $OfficeAppId -eReturnIdType SL_ID_ALL_LICENSE_FILES
Read-Host

Write-Warning 'Get License File IDs associated with a specific Application ID:'
Get-SLIDList -eQueryIdType SL_ID_APPLICATION -pQueryId $OfficeAppId -eReturnIdType SL_ID_ALL_LICENSES
Read-Host

$LicensingProducts = (
    Get-SLIDList -eQueryIdType SL_ID_APPLICATION -eReturnIdType SL_ID_PRODUCT_SKU -pQueryId $windowsAppID | ? { Retrieve-SKUInfo -SkuId $_ -eReturnIdType SL_ID_PKEY }
    ) | % {
    [PSCustomObject]@{
        ID            = $_
        Description   = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'Description'
        Name          = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'productName'
        LicenseFamily = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'Family'
    }
}
#>
enum eQueryIdType {
    SL_ID_APPLICATION = 0
    SL_ID_PRODUCT_SKU = 1
    SL_ID_LICENSE_FILE = 2
    SL_ID_LICENSE = 3
}
enum eReturnIdType {
    SL_ID_APPLICATION = 0
    SL_ID_PRODUCT_SKU = 1
    SL_ID_LICENSE_FILE = 2
    SL_ID_LICENSE = 3
    SL_ID_PKEY = 4
    SL_ID_ALL_LICENSES = 5
    SL_ID_ALL_LICENSE_FILES = 6
}
function Get-SLIDList {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("SL_ID_APPLICATION", "SL_ID_PRODUCT_SKU", "SL_ID_LICENSE_FILE", "SL_ID_LICENSE")]
        [string]$eQueryIdType,

        [Parameter(Mandatory=$true)]
        [ValidateSet("SL_ID_APPLICATION", "SL_ID_PRODUCT_SKU", "SL_ID_LICENSE", "SL_ID_PKEY", "SL_ID_ALL_LICENSES", "SL_ID_ALL_LICENSE_FILES", "SL_ID_LICENSE_FILE")]
        [string]$eReturnIdType,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$pQueryId = $null,

        [Parameter(Mandatory=$false)]
        [Intptr]$hSLC = [IntPtr]::Zero
    )
    
    $dummyGuid = [Guid]::Empty
    $QueryIdValidation = ($eQueryIdType -ne $eReturnIdType) -and [string]::IsNullOrWhiteSpace($pQueryId)
    $GuidValidation = (-not [string]::IsNullOrWhiteSpace($pQueryId)) -and (
        -not [Guid]::TryParse($pQueryId, [ref]$dummyGuid) -or
        ($dummyGuid -eq [Guid]::Empty)
    )
    $AppIDValidation = ($eQueryIdType -ne [eQueryIdType]::SL_ID_APPLICATION) -and ($eReturnIdType.ToString() -match '_ALL_')
    $AppGUIDValidation = ($eQueryIdType -eq [eQueryIdType]::SL_ID_APPLICATION) -and ($eReturnIdType -ne $eQueryIdType) -and
                        (-not ($knownAppGuids -contains $pQueryId))

    if ($AppIDValidation -or $QueryIdValidation -or $GuidValidation -or $AppGUIDValidation) {
        Write-Warning "Invalid parameters:"

        if ($AppIDValidation) {
            "  - _ALL_ types are allowed only with SL_ID_APPLICATION"
            return  }

        if ($QueryIdValidation -and $GuidValidation) {
            "  - A valid, non-empty pQueryId is required when source and target types differ"
            return }

        if ($QueryIdValidation -and $AppGUIDValidation) {
            if ($eQueryIdType -eq [eQueryIdType]::SL_ID_APPLICATION) {
                try {
                    $output = foreach ($appId in $Global:knownAppGuids) {
                        Get-SLIDList -eQueryIdType $eQueryIdType -eReturnIdType $eReturnIdType -pQueryId $appId
                    }
                }
                catch {
                    Write-Warning "An error occurred while attempting to retrieve results with known GUIDs: $_"
                }

                if ($output) {
                    return $output.Guid
                } else {
                    Write-Warning "No valid results returned for the known Application GUIDs."
                    return
                }
            }

            Write-Warning "  - pQueryId must be a known Application GUID when source is SL_ID_APPLICATION and target differs"
            return
        }

        if ($QueryIdValidation) {
            "  - A valid pQueryId is required when source and target types differ"
            return }

        if ($GuidValidation) {
            "  - pQueryId must be a non-empty valid GUID"
            return }

        if ($AppGUIDValidation) {
            "  - pQueryId must match a known Application GUID when source is SL_ID_APPLICATION and target differs"
            return }
    }


    $eQueryIdTypeInt = [eQueryIdType]::$eQueryIdType
    $eReturnIdTypeInt = [eReturnIdType]::$eReturnIdType

    $queryIdPtr = [IntPtr]::Zero 
    $gch = $null                 

    $pnReturnIds = 0
    $ppReturnIds = [IntPtr]::Zero
    
    $needToCloseLocalHandle = $true
    $currentHSLC = if ($hSLC -and $hSLC -ne [IntPtr]::Zero -and $hSLC -ne 0) {
        $hSLC
    } elseif ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
        $global:hSLC_
    } else {
        Manage-SLHandle
    }

    try {
        if (-not $currentHSLC -or $currentHSLC -eq [IntPtr]::Zero -or $currentHSLC -eq 0) {
            $hresult = $Global:SLC::SLOpen([ref]$currentHSLC)
            
            if ($hresult -ne 0) {
                $uint32Value = $hresult -band 0xFFFFFFFF
                $hexString = "0x{0:X8}" -f $uint32Value
                throw "Failed to open SLC handle. HRESULT: $hexString"
            }
        } else {
            $needToCloseLocalHandle = $false
        }
        
        if ($pQueryId) {
            if ($pQueryId -match '^[{]?[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}[}]?$') {
                $queryGuid = [Guid]$pQueryId
                $bytes = $queryGuid.ToByteArray()
                $gch = [GCHandle]::Alloc($bytes, [GCHandleType]::Pinned)
                $queryIdPtr = $gch.AddrOfPinnedObject()
            } else {
                $queryIdPtr = [Marshal]::StringToHGlobalUni($pQueryId)
            }
        } else {
            $queryIdPtr = [IntPtr]::Zero
        }

        $result = $Global:SLC::SLGetSLIDList($currentHSLC, $eQueryIdTypeInt, $queryIdPtr, $eReturnIdTypeInt, [ref]$pnReturnIds, [ref]$ppReturnIds)
        if ($result -eq 0 -and $pnReturnIds -gt 0 -and $ppReturnIds -ne [IntPtr]::Zero) {
            $guidList = @()

            foreach ($i in 0..($pnReturnIds - 1)) {
                $currentPtr = [IntPtr]([Int64]$ppReturnIds + [Int64]16 * $i)
                $guidBytes = New-Object byte[] 16
                [Marshal]::Copy($currentPtr, $guidBytes, 0, 16)
                $guidList += (New-Object Guid (,$guidBytes))
            }
            return $guidList
        } else {
            $uint32Value = $result -band 0xFFFFFFFF
            $hexString = "0x{0:X8}" -f $uint32Value
            if ($result -eq 0xC004F012) {
                return @()
            } else {
                throw "Failed to retrieve ID list. HRESULT: $hexString"
            }
        }
    } catch {
        Write-Warning "Error in Get-SLIDList (QueryIdType: $($eQueryIdType), ReturnIdType: $($eReturnIdType), pQueryId: $($pQueryId)): $($_.Exception.Message)"
        throw $_
    } finally {

        if ($ppReturnIds -ne [IntPtr]::Zero) {
            $null = $Global:kernel32::LocalFree($ppReturnIds)
            $ppReturnIds = [IntPtr]::Zero
        }
        if ($queryIdPtr -ne [IntPtr]::Zero -and $gch -eq $null) {
            [Marshal]::FreeHGlobal($queryIdPtr)
            $queryIdPtr = [IntPtr]::Zero
        }
        if ($gch -ne $null -and $gch.IsAllocated) {
            $gch.Free()
            $gch = $null
        }

        if ($needToCloseLocalHandle -and $currentHSLC -ne [IntPtr]::Zero) {
            Free-IntPtr -handle $currentHSLC -Method License
            $currentHSLC = [IntPtr]::Zero
        }
    }
}

<#
.SYNOPSIS
Function Retrieve-SKUInfo retrieves related licensing IDs for a given SKU GUID.
Also, Support for SL_ID_ALL_LICENSES & SL_ID_ALL_LICENSE_FILES, Only for Application-ID

Specific SKUs require particular IDs:
- The SKU for SLUninstallLicense requires the ID_LICENSE_FILE GUID.
- The SKU for SLUninstallProofOfPurchase requires the ID_PKEY GUID.

Optional Pointer: Handle to the Software Licensing Service (SLC).
Optional eReturnIdType: Type of ID to return (e.g., SL_ID_APPLICATION, SL_ID_PKEY, etc.).
#>
function Retrieve-SKUInfo {
    param(
        [Parameter(Mandatory = $true)]
        [ValidatePattern('^[{]?[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}[}]?$')]
        [string]$SkuId,

        [Parameter(Mandatory = $false)]
        [ValidateSet("SL_ID_APPLICATION", "SL_ID_PRODUCT_SKU", "SL_ID_LICENSE", "SL_ID_PKEY", "SL_ID_ALL_LICENSES", "SL_ID_ALL_LICENSE_FILES", "SL_ID_LICENSE_FILE")]
        [string]$eReturnIdType,

        [Parameter(Mandatory=$false)]
        [Intptr]$hSLC = [IntPtr]::Zero
    )

    # Define once at the top
    $Is__ALL = $eReturnIdType -match '_ALL_'
    $IsAppID = $Global:knownAppGuids -contains $SkuId

    # XOR Case, Check if Both Valid, if one valid, exit
    if ($Is__ALL -xor $IsAppID) {
        Write-Warning "ApplicationID Work with SL_ID_ALL_LICENSES -or SL_ID_ALL_LICENSE_FILES Only!"
        return $null
    }

    function Get-IDs {
        param (
            [string]$returnType,
            [Intptr]$hSLC
        )
        try {
            if ($IsAppID) {
                return Get-SLIDList -eQueryIdType SL_ID_APPLICATION -eReturnIdType $returnType -pQueryId $SkuId -hSLC $hSLC
            } else {
                return Get-SLIDList -eQueryIdType SL_ID_PRODUCT_SKU -eReturnIdType $returnType -pQueryId $SkuId -hSLC $hSLC
            }
        } catch {
            Write-Warning "Get-SLIDList call failed for $returnType and $SkuId"
            return $null
        }
    }

    $product = [Guid]$SkuId

    if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
        $hSLC = if ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
            $global:hSLC_
        } else {
            Manage-SLHandle
        }
    }

    try {
        $closeHandle = $true
        if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
            $hr = $Global:SLC::SLOpen([ref]$hSLC)
            if ($hr -ne 0) {
                throw "SLOpen failed: HRESULT 0x{0:X8}" -f $hr
            }
        } else {
            $closeHandle = $false
        }
    }
    catch {
        return $null
    }

    try {
        # [SL_ID_LICENSE_FILE] Case
        [Guid]$fileId = try {
            [Guid]::Parse((Get-LicenseDetails -ActConfigId $product -pwszValueName fileId -hSLC $hSLC).Trim().Substring(0,36))
        }
        catch {
            [GUID]::Empty
        }

        # [SL_ID_LICENSE] Case **Alternative**
        [Guid]$licenseId = try {
            [Guid]::Parse((Get-LicenseDetails -ActConfigId $product -pwszValueName licenseId -hSLC $hSLC).Trim().Substring(0,36))
        } catch {
            [Guid]::Empty
        }

        [Guid]$privateCertificateId = try {
            [Guid]::Parse((Get-LicenseDetails -ActConfigId $product -pwszValueName privateCertificateId -hSLC $hSLC).Trim().Substring(0,36))
        } catch {
            [Guid]::Empty
        }

        # [SL_ID_APPLICATION] Case **Alternative**
        [Guid]$applicationId = try {
            [Guid]::Parse((Get-LicenseDetails -ActConfigId $product -pwszValueName applicationId -hSLC $hSLC).Trim().Substring(0,36))
        } catch {
            [Guid]::Empty
        }

        # [SL_ID_PKEY] Case **Alternative**
        [Guid]$pkId = try {
            [Guid]::Parse((Get-LicenseDetails -ActConfigId $product -pwszValueName pkeyIdList -hSLC $hSLC).Trim().Substring(0,36)) # Instead `pkeyId`
        } catch {
            [Guid]::Empty
        }

        [uint32]$countRef = 0
        [IntPtr]$ppKeyIds = [intPtr]::Zero
        [GUID]$pKeyId = [GUID]::Empty
        [uint32]$hresults = $Global:SLC::SLGetInstalledProductKeyIds(
            $hSLC, [ref]$product, [ref]$countRef, [ref]$ppKeyIds)
        if ($hresults -eq 0) {
            if ($countRef -gt 0 -and (
                $ppKeyIds -ne [IntPtr]::Zero)) {
                    if ($ppKeyIds.ToInt64() -gt 0) {
                        try {
                            $buffer = New-Object byte[] 16
                            [Marshal]::Copy($ppKeyIds, $buffer, 0, 16)
                            $pKeyId = [Guid]::new($buffer)
                        }
                        catch {
                            $pKeyId = $null
                        }
        }}}

        # -------------------------------------------------

        if (-not $eReturnIdType) {
            $SKU_DATA = [pscustomobject]@{
                ID_SKU          = $SkuId
                ID_APPLICATION  = if ($applicationId -and $applicationId -ne [Guid]::Empty) { $applicationId } else { try { Get-IDs SL_ID_APPLICATION -hSLC $hSLC } catch { [Guid]::Empty } }
                ID_PKEY         = if ($pkId -and $pkId -ne [Guid]::Empty) { $pkId } elseif ($Product_SKU_ID -and $Product_SKU_ID -ne [Guid]::Empty) { $Product_SKU_ID } else { try { Get-IDs SL_ID_PKEY -hSLC $hSLC } catch { [Guid]::Empty } }
                ID_LICENSE_FILE = if ($fileId -and $fileId -ne [Guid]::Empty) { $fileId } else { try { Get-IDs SL_ID_LICENSE_FILE -hSLC $hSLC } catch { [Guid]::Empty } }
                ID_LICENSE      = if (($licenseId -and $privateCertificateId) -and ($licenseId -ne [Guid]::Empty -and $privateCertificateId -ne [Guid]::Empty)) { @($licenseId, $privateCertificateId) } else { try { Get-IDs SL_ID_LICENSE -hSLC $hSLC } catch { [Guid]::Empty } }
            }
            return $SKU_DATA
        }

        switch ($eReturnIdType) {
            "SL_ID_APPLICATION" {
                if ($applicationId -and $applicationId -ne [Guid]::Empty) {
                    return $applicationId
                }
                try { return Get-IDs SL_ID_APPLICATION -hSLC $hSLC } catch {}
                return [Guid]::Empty
            }

            "SL_ID_PRODUCT_SKU" {
                return $SkuId
            }

            "SL_ID_LICENSE" {
                if ($licenseId -and $privateCertificateId -and $licenseId -ne [Guid]::Empty -and $privateCertificateId -ne [Guid]::Empty) {
                    return @($licenseId, $privateCertificateId)
                }
                try { return Get-IDs SL_ID_LICENSE -hSLC $hSLC } catch {}
                return [Guid]::Empty
            }

            "SL_ID_PKEY" {
                if ($pkId -and $pkId -ne [Guid]::Empty) {
                    return $pkId
                }
                if ($pKeyId -and $pKeyId -ne [Guid]::Empty) {
                    return $pKeyId
                }
                try { return Get-IDs SL_ID_PKEY -hSLC $hSLC } catch {}
                return [Guid]::Empty
            }

            "SL_ID_ALL_LICENSES" {
                try { return Get-IDs SL_ID_ALL_LICENSES -hSLC $hSLC } catch {}
                return [Guid]::Empty
            }

            "SL_ID_ALL_LICENSE_FILES" {
                try { return Get-IDs SL_ID_ALL_LICENSE_FILES -hSLC $hSLC } catch {}
                return [Guid]::Empty
            }

            "SL_ID_LICENSE_FILE" {
                if ($fileId -and $fileId -ne [Guid]::Empty) {
                    return $fileId
                }

                # it possible using Get-SLIDList to convert SKU > ID_LICENSE > ID_LICENSE_FILE, but not directly.!
                try { return Get-SLIDList -eQueryIdType SL_ID_LICENSE -eReturnIdType SL_ID_LICENSE_FILE -pQueryId $licenseId } catch {}
                try { return Get-SLIDList -eQueryIdType SL_ID_LICENSE -eReturnIdType SL_ID_LICENSE_FILE -pQueryId $privateCertificateId } catch {}
                return [Guid]::Empty
            }
            default {
                return [Guid]::Empty
            }
        }
    }
    finally {

        if ($null -ne $ppKeyIds -and (
            $ppKeyIds -ne [IntPtr]::Zero) -and (
                $ppKeyIds -ne 0)) {
                    $null = $Global:kernel32::LocalFree($ppKeyIds)
        }

        if ($closeHandle) {
            Write-Warning "Consider Open handle Using Manage-SLHandle"
            Free-IntPtr -handle $hSLC -Method License
        }
    }
}

<#
.SYNOPSIS
Function Receive license data as Config or License file
#>
function Get-LicenseData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [Guid]$SkuID,

        [Parameter(Mandatory=$false)]
        [Intptr]$hSLC = [IntPtr]::Zero,

        [ValidateNotNullOrEmpty()]
        [ValidateSet("License", "Config")]
        [string]$Mode
    )

    if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
        $hSLC = if ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
            $global:hSLC_
        } else {
            Manage-SLHandle
        }
    }

    try {
        $closeHandle = $true
        if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
            $hr = $Global:SLC::SLOpen([ref]$hSLC)
            if ($hr -ne 0) {
                throw "SLOpen failed: HRESULT 0x{0:X8}" -f $hr
            }
        } else {
            $closeHandle = $false
        }
    }
    catch {
        return $null
    }
	
    try {
        $fileGuid = [guid]::Empty
        if ($Mode -eq 'License') {
            $fileGuid = Retrieve-SKUInfo -SkuId $SkuID -eReturnIdType SL_ID_LICENSE_FILE
        }
        if ($Mode -eq 'Config') {
            $LicenseId = Get-LicenseDetails -ActConfigId $SkuID -pwszValueName pkeyConfigLicenseId
            $fileGuid = Retrieve-SKUInfo -SkuId $LicenseId -eReturnIdType SL_ID_LICENSE_FILE
        }
        if (-not $fileGuid -or (
	        [guid]$fileGuid -eq [GUID]::Empty)) {
		        return $null
        }

        $count = 0
        $ppbLicenseFile = [IntPtr]::Zero
        $res = $global:SLC::SLGetLicense($hSLC, [ref]$fileGuid, [ref]$count, [ref]$ppbLicenseFile)
        if ($res -ne 0) { throw "SLGetLicense failed (code $res)" }
        $blob = New-Object byte[] $count
        [Marshal]::Copy($ppbLicenseFile, $blob, 0, $count)
        $content = [Text.Encoding]::UTF8.GetString($blob)
        return $content

    }
    finally {
        Free-IntPtr -handle ppbLicenseFile -Method Local
        if ($closeHandle) {
            Write-Warning "Consider Open handle Using Manage-SLHandle"
            Free-IntPtr -handle $hSLC -Method License
        }
    }
}

<#
.SYNOPSIS
Function Retrieve-SKUInfo retrieves related licensing IDs for a given SKU GUID.
Convert option, CD-KEY->Ref/ID, Ref->SKU, SKU->Ref
#>
function Retrieve-ProductKeyInfo {
    param (
        [ValidateScript({ $_ -ne $null -and $_ -ne [guid]::Empty })]
        [guid]$SkuId,

        [ValidateScript({ $_ -ne $null -and $_ -gt 0 })]
        [int]$RefGroupId,

        [ValidateScript({ $_ -match '^(?i)[A-Z0-9]{5}(-[A-Z0-9]{5}){4}$' })]
        [string]$CdKey
    )

    # Validate only one parameter
    $paramsProvided = @($SkuId, $RefGroupId, $CdKey) | Where-Object { $_ }
    if ($paramsProvided.Count -ne 1) {
        Write-Warning "Please specify exactly one of -SkuId, -RefGroupId, or -CdKey"
        return $null
    }

    # SkuId to RefGroupId
    if ($SkuId) {
        $entry = $Global:PKeyDatabase | Where-Object { $_.ActConfigId -eq "{$SkuId}" } | Select-Object -First 1
        if ($entry) {
            return $entry.RefGroupId
        } else {
            Write-Warning "RefGroupId not found for SkuId: $SkuId"
            return $null
        }
    }

    # RefGroupId to SkuId
    elseif ($RefGroupId) {
        $entry = $Global:PKeyDatabase | Where-Object { $_.RefGroupId -eq $RefGroupId } | Select-Object -First 1
        if ($entry) {
            return [guid]($entry.ActConfigId -replace '[{}]', '')
        } else {
            Write-Warning "ActConfigId not found for RefGroupId: $RefGroupId"
            return $null
        }
    }

    # CdKey to RefGroupId to SkuId
    elseif ($CdKey) {
        try {
            $decoded = KeyDecode -key0 $CdKey.Substring(0,29)
            $refGroupFromKey = [int]$decoded[2].Value

            $entry = $Global:PKeyDatabase | Where-Object { $_.RefGroupId -eq $refGroupFromKey } | Select-Object -First 1
            if ($entry) {
                return [PSCustomObject]@{
                    RefGroupId = $refGroupFromKey
                    SkuId      = [guid]($entry.ActConfigId -replace '[{}]', '')
                }
            } else {
                Write-Warning "SKU not found for RefGroupId $refGroupFromKey extracted from CD Key"
                return $null
            }
        } catch {
            Write-Warning "Failed to decode CD Key: $_"
            return $null
        }
    }
}

<#
.SYNOPSIS
    Fires a licensing state change event after installing or removing a license/key.
#>
function Fire-LicensingStateChangeEvent {
    param (
        [Parameter(Mandatory=$true)]
        [IntPtr]$hSLC
    )
    
    if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
        $hSLC = if ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
            $global:hSLC_
        } else {
            Manage-SLHandle
        }
    }

    try {
        $closeHandle = $true
        if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
            $hr = $Global:SLC::SLOpen([ref]$hSLC)
            if ($hr -ne 0) {
                throw "SLOpen failed: HRESULT 0x{0:X8}" -f $hr
            }
        } else {
            $closeHandle = $false
        }
    }
    catch {
        return $null
    }

    try {
        $SlEvent = "msft:rm/event/licensingstatechanged"
        $WindowsSlid = New-Object Guid($Global:windowsAppID)
        $OfficeSlid  = New-Object Guid($Global:OfficeAppId)
        ($WindowsSlid, $OfficeSlid) | % {
            $hrEvent = $Global:SLC::SLFireEvent(
                $hSLC,  # Using the IntPtr (acting like a pointer)
                $SlEvent, 
                [ref]$_
            )

            # Check if the event firing was successful (HRESULT 0 means success)
            if ($hrEvent -eq 0) {
                Write-Host "Licensing state change event fired successfully."
            } else {
                Write-Host "Failed to fire licensing state change event. HRESULT: $hrEvent"
            }
        }
    }
    finally {
        if ($closeHandle) {
            Write-Warning "Consider Open handle Using Manage-SLHandle"
            Free-IntPtr -handle $hSLC -Method License
        }
    }
}

<#
.SYNOPSIS
    Re-Arm Specific ID <> SKU.
#>
Function SL-ReArm {
    param (
        [Parameter(Mandatory=$false)]
        [ValidateSet(
            '0ff1ce15-a989-479d-af46-f275c6370663',
            '55c92734-d682-4d71-983e-d6ec3f16059f'
        )]
        [GUID]$AppID,

        [Parameter(Mandatory=$false)]
        [GUID]$skuID,

        [Parameter(Mandatory=$false)]
        [Intptr]$hSLC = [IntPtr]::Zero
    )

    if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
        $hSLC = if ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
            $global:hSLC_
        } else {
            Manage-SLHandle
        }
    }

    try {
        $closeHandle = $true
        if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
            $hr = $Global:SLC::SLOpen([ref]$hSLC)
            if ($hr -ne 0) {
                throw "SLOpen failed: HRESULT 0x{0:X8}" -f $hr
            }
        } else {
            $closeHandle = $false
        }
    }
    catch {
    }

    try {
        if (-not $AppID -or -not $skuID) {
            $hrsults = $Global:slc::SLReArmWindows()
        }
        elseif ($AppID -and $skuID) {
            $AppID_ = [GUID]::new($AppID)
            $skuID_ = [GUID]::new($skuID)
            $hrsults = $Global:slc::SLReArm(
                $hSLC, [ref]$AppID_, [REF]$skuID_, 0)
        }        
        if ($hrsults -ne 0) {
            $errorMessege = Parse-ErrorMessage -MessageId $hrsults -Flags ([ErrorMessageType]::ACTIVATION -bor [ErrorMessageType]::HRESULT)
            Write-Warning "$($hrsults): $($errorMessege)"
        }
        return $hrsults
    }
    finally {
        if ($closeHandle) {
            Write-Warning "Consider Open handle Using Manage-SLHandle"
            Free-IntPtr -handle $hSLC -Method License
        }
    }
}

<#
.SYNOPSIS
    Activate Specific SKU.
#>
Function SL-Activate {
    param (
        [Parameter(Mandatory=$true)]
        [GUID]$skuID,

        [Parameter(Mandatory=$false)]
        [Intptr]$hSLC = [IntPtr]::Zero
    )

    if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
        $hSLC = if ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
            $global:hSLC_
        } else {
            Manage-SLHandle
        }
    }

    try {
        $closeHandle = $true
        if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
            $hr = $Global:SLC::SLOpen([ref]$hSLC)
            if ($hr -ne 0) {
                throw "SLOpen failed: HRESULT 0x{0:X8}" -f $hr
            }
        } else {
            $closeHandle = $false
        }
    }
    catch {
    }

    try {
        $skuID_ = [GUID]::new($skuID)
        $hrsults = $Global:slc::SLActivateProduct(
            $hSLC, [REF]$skuID_, 0,[IntPtr]::Zero,[IntPtr]::Zero,$null,0)

        if ($hrsults -ne 0) {
            $errorMessege = Parse-ErrorMessage -MessageId $hrsults -Flags ([ErrorMessageType]::ACTIVATION -bor [ErrorMessageType]::HRESULT)
            Write-Warning "$errorMessege, $hresult"
        }
        return $hrsults
    }
    finally {
        if ($closeHandle) {
            Write-Warning "Consider Open handle Using Manage-SLHandle"
            Free-IntPtr -handle $hSLC -Method License
        }
    }
}

<#
.SYNOPSIS
   WMI -> RefreshLicenseStatus
#>
Function SL-RefreshLicenseStatus {
    param (
        [Parameter(Mandatory=$false)]
        [ValidateSet(
            '0ff1ce15-a989-479d-af46-f275c6370663',
            '55c92734-d682-4d71-983e-d6ec3f16059f'
        )]
        [GUID]$AppID,

        [Parameter(Mandatory=$false)]
        [GUID]$skuID,

        [Parameter(Mandatory=$false)]
        [Intptr]$hSLC = [IntPtr]::Zero
    )

    if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
        $hSLC = if ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
            $global:hSLC_
        } else {
            Manage-SLHandle
        }
    }

    try {
        $closeHandle = $true
        if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
            $hr = $Global:SLC::SLOpen([ref]$hSLC)
            if ($hr -ne 0) {
                throw "SLOpen failed: HRESULT 0x{0:X8}" -f $hr
            }
        } else {
            $closeHandle = $false
        }
    }
    catch {
    }

    try {
        if (-not $AppID -and -not $skuID) {
            $hrsults = $Global:slc::SLConsumeWindowsRight($hSLC)
        }
        elseif ($AppID) {
            $AppID_ = [GUID]::new($AppID)
            if (-not $skuID) {
                $hrsults = $Global:slc::SLConsumeRight(
                    $hSLC, [ref]$AppID_, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero)
            }
            else {
                $skuID_ = [GUID]::new($skuID)
                $skuIDPtr = New-IntPtr -Size 16
                try {
                    [Marshal]::StructureToPtr($skuID_, $skuIDPtr, $false)
                    $hrsults = $Global:slc::SLConsumeRight(
                        $hSLC, [ref]$AppID_, $skuIDPtr, [IntPtr]::Zero, [IntPtr]::Zero)
                }
                finally {
                    New-IntPtr -hHandle $skuIDPtr -Release
                }
            }
        }
        elseif ($skuID) {
            $skuID_ = [GUID]::new($skuID)
            $AppID_ = Retrieve-SKUInfo -SkuId $skuID -eReturnIdType SL_ID_APPLICATION
            if (-not $AppID_) {
                throw "Couldn't retrieve AppId for SKU: $skuID"
            }
            $skuIDPtr = New-IntPtr -Size 16
            try {
                [Marshal]::StructureToPtr($skuID_, $skuIDPtr, $false)
                $hrsults = $Global:slc::SLConsumeRight(
                    $hSLC, [ref]$AppID_, $skuIDPtr, [IntPtr]::Zero, [IntPtr]::Zero)
            }
            finally {
                New-IntPtr -hHandle $skuIDPtr -Release
            }
        }

        if ($hrsults -ne 0) {
            $errorMessege = Parse-ErrorMessage -MessageId $hrsults -Flags ([ErrorMessageType]::ACTIVATION -bor [ErrorMessageType]::HRESULT)
            Write-Warning "$($hrsults): $($errorMessege)"
        }
        return $hrsults
    }
    finally {
        if ($closeHandle) {
            Write-Warning "Consider Open handle Using Manage-SLHandle"
            Free-IntPtr -handle $hSLC -Method License
        }
    }
}

<#
.SYNOPSIS

Usage: KEY
SL-InstallProductKey -Keys "3HYJN-9KG99-F8VG9-V3DT8-JFMHV"
SL-InstallProductKey -Keys ("BW9HJ-N9HF7-7M9PW-3PBJR-37DCT","NJ8QJ-PYYXJ-F6HVQ-RYPFK-BKQ86","K8BH4-6TN3G-YXVMY-HBMMF-KBXPT","GMN9H-QCX29-F3JWJ-RYPKC-DDD86","TN6YY-MWHCT-T6PK2-886FF-6RBJ6")
#>
function SL-InstallProductKey {
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Keys,

        [Parameter(Mandatory=$false)]
        [Intptr]$hSLC = [IntPtr]::Zero
    )

    if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
        $hSLC = if ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
            $global:hSLC_
        } else {
            Manage-SLHandle
        }
    }

    try {
        $closeHandle = $true
        if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
            $hr = $Global:SLC::SLOpen([ref]$hSLC)
            if ($hr -ne 0) {
                throw "SLOpen failed: HRESULT 0x{0:X8}" -f $hr
            }
        } else {
            $closeHandle = $false
        }
    }
    catch {
        return $null
    }

    [guid[]]$PKeyIdLst = @()

    try {
        if (-not $Keys) {
            Write-Warning "No product keys provided. Please provide at least one key."
            return $null
        }

        $invalidKeys = $Keys | Where-Object { [string]::IsNullOrWhiteSpace($_) }
        if ($invalidKeys.Count -gt 0) {
            Write-Warning "The following keys are invalid (empty or whitespace): $($invalidKeys -join ', ')"
            return $null
        }

        foreach ($key in $Keys) {
            $KeyBlob = [System.Text.Encoding]::UTF8.GetBytes($key)
            $KeyTypes = @(
                "msft:rm/algorithm/pkey/detect",
                "msft:rm/algorithm/pkey/2009",
                "msft:rm/algorithm/pkey/2007",
                "msft:rm/algorithm/pkey/2005"
            )

            $PKeyIdOut = [Guid]::NewGuid()
            $installationSuccess = $false

            foreach ($KeyType in $KeyTypes) {
                $hrInstall = $Global:SLC::SLInstallProofOfPurchase(
                    $hSLC,
                    $KeyType,
                    $key,            # Directly using the key string
                    0,               # PKeyDataSize is 0 (no additional data)
                    [IntPtr]::Zero,  # No additional data (zero pointer)
                    [ref]$PKeyIdOut
                )

                if ($hrInstall -eq 0) {
                    Write-Host "Proof of purchase installed successfully with KeyType: $KeyType. PKeyId: $PKeyIdOut"
                    $PKeyIdLst += $PKeyIdOut  # Add the successful GUID to the list
                    $installationSuccess = $true  # Mark success for this key
                    break
                }
            }

            if (-not $installationSuccess) {
                $errorMessege = Parse-ErrorMessage -MessageId $hrInstall -Flags ([ErrorMessageType]::ACTIVATION -bor [ErrorMessageType]::HRESULT)
                Write-Warning "Failed to install the proof of purchase for key $key. HRESULT: $hrInstall"
                Write-Warning "$($hrInstall): $($errorMessege)"
            }
        }
    }
    finally {

        Fire-LicensingStateChangeEvent -hSLC $hSLC     
        if ($closeHandle) {
            Write-Warning "Consider Open handle Using Manage-SLHandle"
            Free-IntPtr -handle $hSLC -Method License
        }
    }

    # Return list of successfully installed PKeyIds
    # return $PKeyIdLst
}

<#
.SYNOPSIS
Usage: KEY -OR SKU -OR PKEY
Usage: Remove Current Windows KEY

Example.
SL-UninstallProductKey -ClearKey $true
SL-UninstallProductKey -KeyLst ("3HYJN-9KG99-F8VG9-V3DT8-JFMHV", "JFMHV") -skuList @("dabaa1f2-109b-496d-bf49-1536cc862900") -pkeyList @("e953e4ac-7ce5-0401-e56c-70c13b8e5a82")
#>
function SL-UninstallProductKey {
    param (
        [Parameter(Mandatory = $false)]
        [string[]]$KeyLst,  # List of partial product keys (optional)

        [Parameter(Mandatory = $false)]
        [GUID[]]$skuList,  # List of GUIDs (optional)

        [Parameter(Mandatory = $false)]
        [GUID[]]$pkeyList,  # List of GUIDs (optional),

        [Parameter(Mandatory=$false)]
        [Intptr]$hSLC = [IntPtr]::Zero,

        [Parameter(Mandatory=$false)]
        [switch]$ClearKey
    )

    if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
        $hSLC = if ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
            $global:hSLC_
        } else {
            Manage-SLHandle
        }
    }

    try {
        $closeHandle = $true
        if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
            $hr = $Global:SLC::SLOpen([ref]$hSLC)
            if ($hr -ne 0) {
                throw "SLOpen failed: HRESULT 0x{0:X8}" -f $hr
            }
        } else {
            $closeHandle = $false
        }
    }
    catch {
        return $null
    }

    try {
        # Initialize the list to hold GUIDs
        $guidList = @()

        if (-not ($skuList -or $KeyLst -or $pkeyList) -and !$ClearKey) {
            Write-Warning "No provided SKU or Key"
            return
        }
        
        $validSkuList = @()
        $validCdKeys  = @()
        $validPkeyList = @()
        if ($KeyLst) {
            foreach ($key in $KeyLst) {
                if ($key.Length -eq 5 -or $key -match '^[A-Z0-9]{5}(-[A-Z0-9]{5}){4}$') {
                    $validCdKeys += $key
                }}}
        if ($skuList) {
            foreach ($sku in $skuList) {
                if ($sku -match '^[{]?[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}[}]?$') {
                    $validSkuList += [guid]$sku  }}}
        if ($pkeyList) {
            foreach ($pkey in $pkeyList) {
                if ($pkey -match '^[{]?[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}[}]?$') {
                    $validPkeyList += [guid]$pkey }}}

        $results = @()
        foreach ($guid in (
            Get-SLIDList -eQueryIdType SL_ID_PRODUCT_SKU -eReturnIdType SL_ID_PRODUCT_SKU)) {
                $ID_PKEY = Retrieve-SKUInfo -SkuId $guid -eReturnIdType SL_ID_PKEY
                if ($ID_PKEY) {
                    $results += [PSCustomObject]@{
                        SL_ID_PRODUCT_SKU = $guid
                        SL_ID_PKEY = $ID_PKEY
                        PartialProductKey = Get-SLCPKeyInfo -PKEY $ID_PKEY -pwszValueName PartialProductKey
                    }}}

        if ($KeyLst) {
            foreach ($key in $KeyLst) {
                if ($key.Length -eq 29) {
                    $pPKeyId = [GUID]::Empty
                    $AlgoTypes = @(
                        "msft:rm/algorithm/pkey/detect",
                        "msft:rm/algorithm/pkey/2009",
                        "msft:rm/algorithm/pkey/2007",
                        "msft:rm/algorithm/pkey/2005"
                    )
                    foreach ($type in $AlgoTypes) {
                        $hresults = $Global:SLC::SLGetPKeyId(
                            $hSLC,$type,$key,[Intptr]::Zero,[Intptr]::Zero,[ref]$pPKeyId)
                        if ($hresults -eq 0) {
                            break;
                        }
                    }
                    if ($hresults -eq 0 -and (
                        $pPKeyId -ne [GUID]::Empty)) {
                            $results += [PSCustomObject]@{
                                SL_ID_PRODUCT_SKU = $null
                                SL_ID_PKEY = $pPKeyId
                                PartialProductKey = $key
                            }
                    }
                }
            }
        }

        # Initialize filtered results array
        $filteredResults = @()
        foreach ($item in $results) {          
            $isValidPkey = $validPkeyList -contains $item.SL_ID_PKEY
            $isValidKey  = $validCdKeys -contains $item.PartialProductKey
            $isValidSKU  = $validSkuList -contains $item.SL_ID_PRODUCT_SKU
            if ($isValidKey){
                write-warning "Valid key found, $($item.PartialProductKey)"
            }
            if ($isValidSKU){
                write-warning "Valid SKU found, $($item.SL_ID_PRODUCT_SKU)"
            }
            if ($isValidPkey){
                write-warning "Valid PKEY found, $($item.SL_ID_PKEY)"
            }
            if ($isValidKey -or $isValidSKU -or $isValidPkey) {
                $filteredResults += $item
            }
        }

        # Step 3: Retrieve unique SL_ID_PKEY values from the filtered results
        $SL_ID_PKEY_list = $filteredResults | Select-Object -ExpandProperty SL_ID_PKEY | Sort-Object | Select-Object -Unique

        if ($ClearKey) {
            $pPKeyId = [GUID]::Empty
            $AlgoTypes = @(
                "msft:rm/algorithm/pkey/detect",
                "msft:rm/algorithm/pkey/2009",
                "msft:rm/algorithm/pkey/2007",
                "msft:rm/algorithm/pkey/2005"
            )
            $DigitalKey  = $(Parse-DigitalProductId).DigitalKey
            if (-not $DigitalKey) {
                $DigitalKey = $(Parse-DigitalProductId4).DigitalKey
            }
            if ($DigitalKey) {
                foreach ($type in $AlgoTypes) {
                    $hresults = $Global:SLC::SLGetPKeyId(
                        $hSLC,$type,$DigitalKey,[Intptr]::Zero,[Intptr]::Zero,[ref]$pPKeyId)
                    if ($hresults -eq 0) {
                        break;
                    }
                }
                if ($hresults -eq 0 -and (
                    $pPKeyId -ne [GUID]::Empty)) {
                        $SL_ID_PKEY_list = @($pPKeyId)
                }
            }
        }

        # Proceed to uninstall each product key using its GUID
        foreach ($guid in $SL_ID_PKEY_list) {
            if ($guid) {
                Write-Host "Attempting to uninstall product key with GUID: $guid"
                $hrUninstall = $Global:SLC::SLUninstallProofOfPurchase($hSLC, $guid)

                if ($hrUninstall -eq 0) {
                    Write-Host "Product key uninstalled successfully: $guid"
                } else {
                    $uint32Value = $hrUninstall -band 0xFFFFFFFF
                    $hexString = "0x{0:X8}" -f $uint32Value
                    Write-Warning "Failed to uninstall product key with HRESULT: $hexString for GUID: $guid"
                }
            } else {
                Write-Warning "Skipping invalid GUID: $guid"
            }
        }
    }
    catch {
        Write-Warning "An unexpected error occurred: $_"
    }
    finally {
        
        # Launch event of license status change after license/key install/remove
        Fire-LicensingStateChangeEvent -hSLC $hSLC

        if ($closeHandle) {
            Write-Warning "Consider Open handle Using Manage-SLHandle"
            Free-IntPtr -handle $hSLC -Method License
        }
    }
}

<#
.SYNOPSIS

# Path to license file
$licensePath = 'C:\Program Files\Microsoft Office\root\Licenses16\client-issuance-bridge-office.xrm-ms'

if (-not (Test-Path $licensePath)) {
    Write-Warning "License file not found: $licensePath"
    return
}

# 1. Install license from file path (string)
Write-Host "`n--- Installing from file path ---"
$result1 = SL-InstallLicense -LicenseInput $licensePath
Write-Host "Result (file path): $result1`n"

# 2. Install license from byte array
Write-Host "--- Installing from byte array ---"
$bytes = [System.IO.File]::ReadAllBytes($licensePath)
$result2 = SL-InstallLicense -LicenseInput $bytes
Write-Host "Result (byte array): $result2`n"

# 3. Install license from text string
Write-Host "--- Installing from text string ---"
$licenseText = Get-Content $licensePath -Raw
$result3 = SL-InstallLicense -LicenseInput $licenseText
Write-Host "Result (text): $result3`n"
#>
function SL-InstallLicense {
    param (
        # Can be string (file path or raw text) or byte[]
        [Parameter(Mandatory = $true)]
        [object[]]$LicenseInput,

        [Parameter(Mandatory=$false)]
        [Intptr]$hSLC = [IntPtr]::Zero
    )
    
    # Prepare to install license
    $LicenseFileIdOut = [Guid]::NewGuid()
    # Store the file IDs for all successfully installed licenses
    $LicenseFileIds = @()

    if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
        $hSLC = if ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
            $global:hSLC_
        } else {
            Manage-SLHandle
        }
    }

    try {
        $closeHandle = $true
        if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
            $hr = $Global:SLC::SLOpen([ref]$hSLC)
            if ($hr -ne 0) {
                throw "SLOpen failed: HRESULT 0x{0:X8}" -f $hr
            }
        } else {
            $closeHandle = $false
        }
    }
    catch {
        return $null
    }

    try {
        # Initialize an array to store blobs
        $LicenseBlobs = @()

        # Loop through each license input
        foreach ($input in $LicenseInput) {
            # Determine the type of input and process accordingly
            if ($input -is [byte[]]) {
                # If input is already a byte array, use it directly
                $LicenseBlob = $input
            }
            elseif ($input -is [string]) {
                if (Test-Path $input) {
                    # If it's a file path, read the file and get its byte array
                    $LicenseBlob = [System.IO.File]::ReadAllBytes($input)
                }
                else {
                    # If it's plain text, convert the text to a byte array
                    $LicenseBlob = [Encoding]::UTF8.GetBytes($input)
                }
            }
            else {
                Write-Warning "Invalid input type. Provide a file path, byte array, or text string."
                continue
            }

            if ($LicenseBlob) {
                # Pin the current blob in memory
                $blobPtr = [Marshal]::UnsafeAddrOfPinnedArrayElement($LicenseBlob, 0)
        
                # Call the installation API for the current LicenseBlob
                $hrInstall = $Global:SLC::SLInstallLicense($hSLC, $LicenseBlob.Length, $blobPtr, [ref]$LicenseFileIdOut)

                # Check if the installation was successful (HRESULT 0 means success)
                if ($hrInstall -ne 0) {
                    $errorMessege = Parse-ErrorMessage -MessageId $hrInstall -Flags ([ErrorMessageType]::ACTIVATION -bor [ErrorMessageType]::HRESULT)
                    if ($errorMessege) {
                        Write-Warning "$($hrInstall): $($errorMessege)"
                    } else {
                        Write-Warning "Unknown error HRESULT $hexString"
                    }
                    Write-Warning "Failed to install the proof of purchase for key $key. HRESULT: $hrInstall"
                    continue  # Skip to the next blob if the current installation fails
                }

                # If successful, add the LicenseFileIdOut to the array
                $LicenseFileIds += $LicenseFileIdOut
                Write-Host "Successfully installed license with FileId: $LicenseFileIdOut"
            }

        }
    }
    finally {
        # Launch event of license status change after license/key install/remove
        Fire-LicensingStateChangeEvent -hSLC $hSLC

        if ($closeHandle) {
            Write-Warning "Consider Open handle Using Manage-SLHandle"
            Free-IntPtr -handle $hSLC -Method License
        }
    }

    # Return all the File IDs that were successfully installed
    #return $LicenseFileIds
}

<#
.SYNOPSIS
Uninstalls the license specified by the license file ID and target user option.

By --> SL_ID_ALL_LICENSE_FILES
$OfficeAppId  = '0ff1ce15-a989-479d-af46-f275c6370663'
$SL_ID_List = Get-SLIDList -eQueryIdType SL_ID_APPLICATION -pQueryId $OfficeAppId -eReturnIdType SL_ID_ALL_LICENSE_FILES
SL-UninstallLicense -LicenseFileIds $SL_ID_List

By --> SL_ID_PRODUCT_SKU -->
$OfficeAppId  = '0ff1ce15-a989-479d-af46-f275c6370663'
$WMI_QUERY = Get-SLIDList -eQueryIdType SL_ID_APPLICATION -eReturnIdType SL_ID_PRODUCT_SKU -pQueryId $OfficeAppId
SL-UninstallLicense -ProductSKUs $WMI_QUERY

>>> Results >> 

ActConfigId                            ProductDescription                  
-----------                            ------------------                  
{DABAA1F2-109B-496D-BF49-1536CC862900} Office16_O365AppsBasicR_Subscription

>>> Command >>
SL-UninstallLicense -ProductSKUs ('DABAA1F2-109B-496D-BF49-1536CC862900' -as [GUID])
#>
function SL-UninstallLicense {
    param (
        [Parameter(Mandatory=$false)]
        [Guid[]]$ProductSKUs,

        [Parameter(Mandatory=$false)]
        [Guid[]]$LicenseFileIds,

        [Parameter(Mandatory=$false)]
        [Intptr]$hSLC = [IntPtr]::Zero
    )

    if (-not $ProductSKUs -and -not $LicenseFileIds) {
        throw "You must provide at least one of -ProductSKUs or -LicenseFileIds."
    }

    if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
        $hSLC = if ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
            $global:hSLC_
        } else {
            Manage-SLHandle
        }
    }

    try {
        $closeHandle = $true
        if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
            $hr = $Global:SLC::SLOpen([ref]$hSLC)
            if ($hr -ne 0) {
                throw "SLOpen failed: HRESULT 0x{0:X8}" -f $hr
            }
        } else {
            $closeHandle = $false
        }
    }
    catch {
        return $null
    }

    try {
        
        $LicenseFileIdsLst = @()

        # Add valid LicenseFileIds directly
        if ($LicenseFileIds) {
            foreach ($lfid in $LicenseFileIds) {
                if ($lfid -is [Guid]) {
                    $LicenseFileIdsLst += $lfid }}}

        # Convert each ProductSKU to LicenseFileId and add it
        if ($ProductSKUs) {
            foreach ($sku in $ProductSKUs) {
                if ($sku -isnot [Guid]) { continue }
                $fileGuid = Retrieve-SKUInfo -SkuId $sku -eReturnIdType SL_ID_LICENSE_FILE -hSLC $hSLC
                if ($fileGuid -and ($fileGuid -is [Guid]) -and ($fileGuid -ne [Guid]::Empty)) {
                    $LicenseFileIdsLst += $fileGuid }}}

        foreach ($LicenseFileId in ($LicenseFileIdsLst | Sort-Object -Unique)) {
            $hresult = $Global:SLC::SLUninstallLicense($hSLC, [ref]$LicenseFileId)
            if ($hresult -ne 0) {
                $errorMessege = Parse-ErrorMessage -MessageId $hresult -Flags ([ErrorMessageType]::ACTIVATION -bor [ErrorMessageType]::HRESULT)
                Write-Warning "$errorMessege, $hresult"
            } 
            else {
                Write-Warning "License File ID: $LicenseFileId was removed."
            }
        }
        
    }
    catch {
        # Convert to unsigned 32-bit int (number)
        $hresult = $_.Exception.HResult
        if ($hresult -ne 0) {
            $errorMessege = Parse-ErrorMessage -MessageId $hresult -Flags ([ErrorMessageType]::ACTIVATION -bor [ErrorMessageType]::HRESULT)
            Write-Warning "$errorMessege, $hresult"
        }
    }
    finally {
        
        # Launch event of license status change after license/key install/remove
        Fire-LicensingStateChangeEvent -hSLC $hSLC

        if ($closeHandle) {
            Write-Warning "Consider Open handle Using Manage-SLHandle"
            Free-IntPtr -handle $hSLC -Method License
        }
    }
}

<#
.SYNOPSIS
Retrieves Software Licensing Client status for application and product SkuID.

Example:
$LicensingProducts = (
    Get-SLIDList -eQueryIdType SL_ID_APPLICATION -eReturnIdType SL_ID_PRODUCT_SKU -pQueryId $windowsAppID | ? { Retrieve-SKUInfo -SkuId $_ -eReturnIdType SL_ID_PKEY }
    ) | % {
    [PSCustomObject]@{
        ID            = $_
        PKEY          = Retrieve-SKUInfo -SkuId $_ -eReturnIdType SL_ID_PKEY
        Description   = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'Description'
        Name          = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'productName'
        LicenseFamily = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'Family'
    }
}

Clear-Host
$LicensingProducts | % { 
    Write-Host
    Write-Warning "Get-SLCPKeyInfo Function"
    Get-SLCPKeyInfo -PKEY ($_).PKEY -loopAllValues

    Write-Host
    Write-Warning "Get-SLLicensingStatus"
    Get-SLLicensingStatus -ApplicationID 55c92734-d682-4d71-983e-d6ec3f16059f -SkuID ($_).ID

    Write-Host
    Write-Warning "Get-GenuineInformation"
    Write-Host
    Get-GenuineInformation -QueryId ($_).ID -loopAllValues

    Write-Host
    Write-Warning "Get-ApplicationInformation"
    Write-Host
    Get-ApplicationInformation -ApplicationId ($_).ID -loopAllValues
}
#>
enum LicenseStatusEnum {
    Unlicensed        = 0
    Licensed          = 1
    OOBGrace          = 2
    OOTGrace          = 3
    NonGenuineGrace   = 4
    Notification      = 5
    ExtendedGrace     = 6
}
enum LicenseCategory {
    KMS38        # Valid until 2038
    KMS4K        # Beyond 2038
    ShortTermVL  # Volume expiring within 6 months
    Unknown
    NonKMS
}
function Get-SLLicensingStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("55c92734-d682-4d71-983e-d6ec3f16059f", "0ff1ce15-a989-479d-af46-f275c6370663")]
        [Guid]$ApplicationID,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [Guid]$SkuID,

        [Parameter(Mandatory=$false)]
        [Intptr]$hSLC = [IntPtr]::Zero
    )

    if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
        $hSLC = if ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
            $global:hSLC_
        } else {
            Manage-SLHandle
        }
    }

    try {
        $closeHandle = $true
        if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
            $hr = $Global:SLC::SLOpen([ref]$hSLC)
            if ($hr -ne 0) {
                throw "SLOpen failed: HRESULT 0x{0:X8}" -f $hr
            }
        } else {
            $closeHandle = $false
        }
    }
    catch {
        return $null
    }

    try {
        $guidPattern = '^(?im)[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$'
        if ($SkuID -notmatch $guidPattern) {
            return $null  }

        $pnStatusCount = [uint32]0
        $ppLicensingStatus = [IntPtr]::Zero
        $pAppID = $ApplicationID
        $pProductSkuId = $SkuID

        $result = $global:slc::SLGetLicensingStatusInformation(
            $hSLC,
            [ref]$pAppID,
            [ref]$pProductSkuId,
            [IntPtr]::Zero,
            [ref]$pnStatusCount,
            [ref]$ppLicensingStatus
        )

        $licensingInfo = $null

        if ($result -eq 0 -and $pnStatusCount -gt 0 -and $ppLicensingStatus -ne [IntPtr]::Zero) {
            $eStatus = [Marshal]::ReadInt32($ppLicensingStatus, 16)
            $dwGraceTime = [Marshal]::ReadInt32($ppLicensingStatus, 20)
            $dwTotalGraceDays = [Marshal]::ReadInt32($ppLicensingStatus, 24)
            $hrReason = [Marshal]::ReadInt32($ppLicensingStatus, 28)
            $qwValidityExpiration = [Marshal]::ReadInt64($ppLicensingStatus, 32)

            if (($null -eq $eStatus) -or ($null -eq $dwGraceTime)) {
                return $null
            }

            $expirationDateTime = $null
            if ($qwValidityExpiration -gt 0) {
                try {
                    $expirationDateTime = [DateTime]::FromFileTimeUtc($qwValidityExpiration)
                } catch {}
            }

            $now = Get-Date
            $graceExpirationDate = $now.AddMinutes($dwGraceTime)
            $graceUpToYear = $graceExpirationDate.Year

            $daysLeft = ($graceExpirationDate - $now).Days
            $year = $graceExpirationDate.Year

            $licenseCategory = $Global:PKeyDatabase | Where-Object { $_.ActConfigId -eq "{$skuId}" } | Select-Object -First 1 -ExpandProperty ProductKeyType

            if ($licenseCategory -ieq 'Volume:GVLK') {
                if ($year -gt 2038) {
                    $typeKMS = [LicenseCategory]::KMS4K
                }
                elseif ($year -in 2037, 2038) {
                    $typeKMS = [LicenseCategory]::KMS38
                }
                elseif ($daysLeft -le 180 -and $daysLeft -ge 0) {
                    $typeKMS = [LicenseCategory]::ShortTermVL
                }
                else {
                    $typeKMS = [LicenseCategory]::Unknown
                }
            } else {
                $typeKMS = [LicenseCategory]::NonKMS
            }

            $errorMessage = Parse-ErrorMessage -MessageId $hrReason -Flags ACTIVATION
            $hrHex = '0x{0:X8}' -f ($hrReason -band 0xFFFFFFFF)

            return [PSCustomObject]@{
                ID                   = $SkuID
                LicenseStatus        = [LicenseStatusEnum]$eStatus
                GracePeriodRemaining = $dwGraceTime
                TotalGraceDays       = $dwTotalGraceDays
                EvaluationEndDate    = $expirationDateTime
                LicenseStatusReason  = $hrHex
                LicenseChannel       = $licenseCategory
                LicenseTier          = $typeKMS
                ApiCallHResult       = '0x{0:X8}' -f $result
                ErrorMessege         = $errorMessage
            }
        }
    }
    catch{}
    finally {
        
        Free-IntPtr -handle $ppLicensingStatus -Method Local
        if ($closeHandle) {
            Write-Warning "Consider Open handle Using Manage-SLHandle"
            Free-IntPtr -handle $hSLC -Method License
        }
    }
}

<#
.SYNOPSIS
Gets the information of the specified product key.
work for all sku that are activated AKA [SL_ID_PKEY],
else, no results.

how to receive them --> demo code -->
Get-SLIDList -eQueryIdType SL_ID_PRODUCT_SKU -eReturnIdType SL_ID_PRODUCT_SKU | ? {Retrieve-SKUInfo -SkuId $_ -eReturnIdType SL_ID_PKEY}
Get-SLIDList -eQueryIdType SL_ID_PRODUCT_SKU -eReturnIdType SL_ID_PRODUCT_SKU | ? {Retrieve-SKUInfo -SkuId $_ -eReturnIdType SL_ID_PKEY} | % { Get-SLCPKeyInfo $_ -loopAllValues }

Example:
$LicensingProducts = (
    Get-SLIDList -eQueryIdType SL_ID_APPLICATION -eReturnIdType SL_ID_PRODUCT_SKU -pQueryId $windowsAppID | ? { Retrieve-SKUInfo -SkuId $_ -eReturnIdType SL_ID_PKEY }
    ) | % {
    [PSCustomObject]@{
        ID            = $_
        PKEY          = Retrieve-SKUInfo -SkuId $_ -eReturnIdType SL_ID_PKEY
        Description   = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'Description'
        Name          = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'productName'
        LicenseFamily = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'Family'
    }
}

Clear-Host
$LicensingProducts | % { 
    Write-Host
    Write-Warning "Get-SLCPKeyInfo Function"
    Get-SLCPKeyInfo -PKEY ($_).PKEY -loopAllValues

    Write-Host
    Write-Warning "Get-SLLicensingStatus"
    Get-SLLicensingStatus -ApplicationID 55c92734-d682-4d71-983e-d6ec3f16059f -SkuID ($_).ID

    Write-Host
    Write-Warning "Get-GenuineInformation"
    Write-Host
    Get-GenuineInformation -QueryId ($_).ID -loopAllValues

    Write-Host
    Write-Warning "Get-ApplicationInformation"
    Write-Host
    Get-ApplicationInformation -ApplicationId ($_).ID -loopAllValues
}
#>
function Get-SLCPKeyInfo {
    param(
        [Parameter(Mandatory = $false)]
        [Guid] $SKU,

        [Parameter(Mandatory = $false)]
        [Guid] $PKEY,

        [Parameter(Mandatory = $false)]
        [ValidateSet("DigitalPID", "DigitalPID2", "PartialProductKey", "ProductSkuId", "Channel")]
        [string] $pwszValueName,

        [Parameter(Mandatory = $false)]
        [switch]$loopAllValues,

        [Parameter(Mandatory=$false)]
        [Intptr]$hSLC = [IntPtr]::Zero
    )

    # Oppsite XOR Case, Validate it Not both
    if (!($pwszValueName -xor $loopAllValues)) {
        Write-Warning "Choice 1 option only, can't use both / none"
        return
    }

    if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
        $hSLC = if ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
            $global:hSLC_
        } else {
            Manage-SLHandle
        }
    }

    try {
        $closeHandle = $true
        if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
            $hr = $Global:SLC::SLOpen([ref]$hSLC)
            if ($hr -ne 0) {
                throw "SLOpen failed: HRESULT 0x{0:X8}" -f $hr
            }
        } else {
            $closeHandle = $false
        }
    }
    catch {
        return $null
    }

    try {
        
        # If loopAllValues is true, loop through all values in the ValidateSet and fetch the details for each
        $allValues = @{}
        $allValueNames = @("DigitalPID", "DigitalPID2", "PartialProductKey", "ProductSkuId", "Channel" )
        foreach ($valueName in $allValueNames) {
            $allValues[$valueName] = $null
        }

        if ($PKEY -and $PKEY -ne [GUID]::Empty) {
            $PKeyId = $PKEY
        }
        elseif ($SKU -and $SKU -ne [GUID]::Empty) {
            $PKeyId = Retrieve-SKUInfo -SkuId $SKU -eReturnIdType SL_ID_PKEY -hSLC $hSLC
        }        
        if (-not $PKeyId) {
            return ([GUID]::Empty)
        }

        if ($loopAllValues) {
            foreach ($valueName in $allValueNames) {
                $dataType = 0
                $bufferSize = 0
                $bufferPtr = [IntPtr]::Zero

                $hr = $Global:SLC::SLGetPKeyInformation(
                    $hSLC, [ref]$PKeyId, $valueName, [ref]$dataType, [ref]$bufferSize, [ref]$bufferPtr )

                if ($hr -ne 0) {
                    continue;
                }
                $allValues[$valueName] = Parse-RegistryData -dataType $dataType -ptr $bufferPtr -valueSize $bufferSize -valueName $valueName
            }
            return $allValues
        }

        $dataType = 0
        $bufferSize = 0
        $bufferPtr = [IntPtr]::Zero

        $hr = $Global:SLC::SLGetPKeyInformation(
            $hSLC, [ref]$PKeyId, $pwszValueName, [ref]$dataType, [ref]$bufferSize, [ref]$bufferPtr )

        if ($hr -ne 0) {
            throw "SLGetPKeyInformation failed: HRESULT 0x{0:X8}" -f $hr
        }
        return Parse-RegistryData -dataType $dataType -ptr $bufferPtr -valueSize $bufferSize -valueName $pwszValueName
    }
    catch { }
    finally {
        if ($null -ne $bufferPtr -and (
            $bufferPtr -ne [IntPtr]::Zero)) {
                $null = $Global:kernel32::LocalFree($bufferPtr)
        }
        if ($closeHandle) {
            Write-Warning "Consider Open handle Using Manage-SLHandle"
            Free-IntPtr -handle $hSLC -Method License
        }
    }
}

<#
.SYNOPSIS
Gets information about the genuine state of a Windows computer.

[in] pQueryId
A pointer to an SLID structure that specifies the *application* to check.

pQueryId
pQueryId can be one of the following.  

ApplicationId in case of querying following property values.
    SL_PROP_BRT_DATA
    SL_PROP_BRT_COMMIT

SKUId in case of querying following property values.
    SL_PROP_LAST_ACT_ATTEMPT_HRESULT
    SL_PROP_LAST_ACT_ATTEMPT_TIME
    SL_PROP_LAST_ACT_ATTEMPT_SERVER_FLAGS
    SL_PROP_ACTIVATION_VALIDATION_IN_PROGRESS

Example:
$LicensingProducts = (
    Get-SLIDList -eQueryIdType SL_ID_APPLICATION -eReturnIdType SL_ID_PRODUCT_SKU -pQueryId $windowsAppID | ? { Retrieve-SKUInfo -SkuId $_ -eReturnIdType SL_ID_PKEY }
    ) | % {
    [PSCustomObject]@{
        ID            = $_
        PKEY          = Retrieve-SKUInfo -SkuId $_ -eReturnIdType SL_ID_PKEY
        Description   = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'Description'
        Name          = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'productName'
        LicenseFamily = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'Family'
    }
}

Clear-Host
$LicensingProducts | % { 
    Write-Host
    Write-Warning "Get-SLCPKeyInfo Function"
    Get-SLCPKeyInfo -PKEY ($_).PKEY -loopAllValues

    Write-Host
    Write-Warning "Get-SLLicensingStatus"
    Get-SLLicensingStatus -ApplicationID 55c92734-d682-4d71-983e-d6ec3f16059f -SkuID ($_).ID

    Write-Host
    Write-Warning "Get-GenuineInformation"
    Write-Host
    Get-GenuineInformation -QueryId ($_).ID -loopAllValues

    Write-Host
    Write-Warning "Get-ApplicationInformation"
    Write-Host
    Get-ApplicationInformation -ApplicationId ($_).ID -loopAllValues
}
 #>
function Get-GenuineInformation {
    param (
        [Parameter(Mandatory)]
        [string]$QueryId,

        [Parameter(Mandatory=$false)]
        [ValidateSet(
            'SL_BRT_DATA',
            'SL_BRT_COMMIT',
            'SL_GENUINE_RESULT',
            'SL_GET_GENUINE_AUTHZ',
            'SL_NONGENUINE_GRACE_FLAG',
            'SL_LAST_ACT_ATTEMPT_TIME',
            'SL_LAST_ACT_ATTEMPT_HRESULT',
            'SL_LAST_ACT_ATTEMPT_SERVER_FLAGS',
            'SL_ACTIVATION_VALIDATION_IN_PROGRESS'
        )]
        [string]$ValueName,

        [Parameter(Mandatory = $false)]
        [switch]$loopAllValues,

        [Parameter(Mandatory=$false)]
        [Intptr]$hSLC = [IntPtr]::Zero
    )

    # Oppsite XOR Case, Validate it Not both
    if (!($ValueName -xor $loopAllValues)) {
        Write-Warning "Choice 1 option only, can't use both / none"
        return
    }

    # Cast ApplicationId to Guid
    $appGuid = [Guid]::Parse($QueryId)
    $IsAppID = $Global:knownAppGuids -contains $appGuid

    if ($IsAppID -and (-not $loopAllValues) -and ($ValueName -notmatch "_BRT_|GENUINE")) {
        Write-Warning "The selected property '$ValueName' is not valid for an ApplicationId."
        return
    }
    elseif ((-not $IsAppID) -and (-not $loopAllValues) -and ($ValueName -match '_BRT_')) {
        Write-Warning "The selected property '$ValueName' is not valid for a SKUId."
        return
    }

    if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
        $hSLC = if ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
            $global:hSLC_
        } else {
            Manage-SLHandle
        }
    }

    try {
        $closeHandle = $true
        if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
            $hr = $Global:SLC::SLOpen([ref]$hSLC)
            if ($hr -ne 0) {
                throw "SLOpen failed: HRESULT 0x{0:X8}" -f $hr
            }
        } else {
            $closeHandle = $false
        }
    }
    catch {
        return $null
    }

    # Prepare variables for out params
    $dataType = 0
    $valueSize = 0
    $ptrValue = [IntPtr]::Zero

    if ($loopAllValues) {
          
        # Combine all arrays and remove duplicates
        $allValues = @{}
        $allValueNames = if ($IsAppID) {
            @(
                'SL_BRT_DATA',
                'SL_BRT_COMMIT'
                'SL_GENUINE_RESULT',
                'SL_GET_GENUINE_AUTHZ',
                'SL_NONGENUINE_GRACE_FLAG'
            )
        } else {
            @(
                'SL_LAST_ACT_ATTEMPT_HRESULT',
                'SL_LAST_ACT_ATTEMPT_TIME',
                'SL_LAST_ACT_ATTEMPT_SERVER_FLAGS',
                'SL_ACTIVATION_VALIDATION_IN_PROGRESS'
            )
        }

        foreach ($Name in $allValueNames) {
            
            # Clear value
            $allValues[$Name] = $null

            $dataType = 0
            $valueSize = 0

            $hresult = $global:SLC::SLGetGenuineInformation(
                [ref] $appGuid,
                $Name,
                [ref] $dataType,
                [ref] $valueSize,
                [ref] $ptrValue
            )

            if ($hresult -ne 0) {
                continue
            }
            if ($valueSize -eq 0 -or $ptrValue -eq [IntPtr]::Zero) {
                continue
            }

            $allValues[$Name] = Parse-RegistryData -dataType $dataType -ptr $ptrValue -valueSize $valueSize -valueName $Name
            Free-IntPtr -handle $ptrValue -Method Local
            
            if ($allValues[$Name] -and (
                $Name -eq 'SL_LAST_ACT_ATTEMPT_HRESULT')) {
                $hrReason = $allValues[$Name]
                $errorMessage = Parse-ErrorMessage -MessageId $hrReason -Flags ACTIVATION
                $hrHex = '0x{0:X8}' -f ($hrReason -band 0xFFFFFFFF)
                $allValues[$Name] = $hrHex
            }
            
        }
        if ($errorMessage) {
            $allValues.Add("SL_LAST_ACT_ATTEMPT_MESSEGE",$errorMessage)
        }
        return $allValues
    }

    try {
        # Call SLGetGenuineInformation - pass [ref] for out params
        $hresult = $global:SLC::SLGetGenuineInformation(
            [ref] $appGuid,
            $ValueName,
            [ref] $dataType,
            [ref] $valueSize,
            [ref] $ptrValue
        )

        if ($hresult -ne 0) {
            $errorMessege = Parse-ErrorMessage -MessageId $hresult -Flags ([ErrorMessageType]::ACTIVATION -bor [ErrorMessageType]::HRESULT)
            Write-Warning "$errorMessege, $hresult"
            throw "SLGetGenuineInformation failed with HRESULT: $hresult"
        }

        if ($valueSize -eq 0 -or $ptrValue -eq [IntPtr]::Zero) {
            return $null
        }

        try {
            return Parse-RegistryData -dataType $dataType -ptr $ptrValue -valueSize $valueSize -valueName $ValueName
        }
        finally {
            Free-IntPtr -handle $ptrValue -Method Local
        }
    }
    finally {
        if ($closeHandle) {
            Write-Warning "Consider Open handle Using Manage-SLHandle"
            Free-IntPtr -handle $hSLC -Method License
        }
    }
}

<#
.SYNOPSIS
Gets information about the specified application.

Example:
$LicensingProducts = (
    Get-SLIDList -eQueryIdType SL_ID_APPLICATION -eReturnIdType SL_ID_PRODUCT_SKU -pQueryId $windowsAppID | ? { Retrieve-SKUInfo -SkuId $_ -eReturnIdType SL_ID_PKEY }
    ) | % {
    [PSCustomObject]@{
        ID            = $_
        PKEY          = Retrieve-SKUInfo -SkuId $_ -eReturnIdType SL_ID_PKEY
        Description   = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'Description'
        Name          = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'productName'
        LicenseFamily = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'Family'
    }
}

Clear-Host
$LicensingProducts | % { 
    Write-Host
    Write-Warning "Get-SLCPKeyInfo Function"
    Get-SLCPKeyInfo -PKEY ($_).PKEY -loopAllValues

    Write-Host
    Write-Warning "Get-SLLicensingStatus"
    Get-SLLicensingStatus -ApplicationID 55c92734-d682-4d71-983e-d6ec3f16059f -SkuID ($_).ID

    Write-Host
    Write-Warning "Get-GenuineInformation"
    Write-Host
    Get-GenuineInformation -QueryId ($_).ID -loopAllValues

    Write-Host
    Write-Warning "Get-ApplicationInformation"
    Write-Host
    Get-ApplicationInformation -ApplicationId ($_).ID -loopAllValues
}
#>
function Get-ApplicationInformation {
    param (
        [Parameter(Mandatory)]
        [string]$ApplicationId,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet(
            "TrustedTime",
            "IsKeyManagementService",
            "KeyManagementServiceCurrentCount",
            "KeyManagementServiceRequiredClientCount",
            "KeyManagementServiceUnlicensedRequests",
            "KeyManagementServiceLicensedRequests",
            "KeyManagementServiceOOBGraceRequests",
            "KeyManagementServiceOOTGraceRequests",
            "KeyManagementServiceNonGenuineGraceRequests",
            "KeyManagementServiceNotificationRequests",
            "KeyManagementServiceTotalRequests",
            "KeyManagementServiceFailedRequests"
        )]
        [string]$PropertyName,

        [Parameter(Mandatory = $false)]
        [switch]$loopAllValues,

        [Parameter(Mandatory=$false)]
        [Intptr]$hSLC = [IntPtr]::Zero
    )

    # Oppsite XOR Case, Validate it Not both
    if (!($PropertyName -xor $loopAllValues)) {
        Write-Warning "Choice 1 option only, can't use both / none"
        return
    }

    # Cast ApplicationId to Guid
    $appGuid = [Guid]$ApplicationId

    if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
        $hSLC = if ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
            $global:hSLC_
        } else {
            Manage-SLHandle
        }
    }

    try {
        $closeHandle = $true
        if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
            $hr = $Global:SLC::SLOpen([ref]$hSLC)
            if ($hr -ne 0) {
                throw "SLOpen failed: HRESULT 0x{0:X8}" -f $hr
            }
        } else {
            $closeHandle = $false
        }
    }
    catch {
        return $null
    }

    if ($loopAllValues) {
          
        # Combine all arrays and remove duplicates
        $allValues = @{}
        $allValueNames = (
            "TrustedTime",
            "IsKeyManagementService",
            "KeyManagementServiceCurrentCount",
            "KeyManagementServiceRequiredClientCount",
            "KeyManagementServiceUnlicensedRequests",
            "KeyManagementServiceLicensedRequests",
            "KeyManagementServiceOOBGraceRequests",
            "KeyManagementServiceOOTGraceRequests",
            "KeyManagementServiceNonGenuineGraceRequests",
            "KeyManagementServiceNotificationRequests",
            "KeyManagementServiceTotalRequests",
            "KeyManagementServiceFailedRequests"
        )

        $dataTypePtr = [Marshal]::AllocHGlobal(4)
        $valueSizePtr = [Marshal]::AllocHGlobal(4)
        $ptrPtr = [Marshal]::AllocHGlobal([IntPtr]::Size)

        foreach ($valueName in $allValueNames) {
            # Clear value
            $allValues[$valueName] = $null

            # Initialize the out params to zero/null
            [Marshal]::WriteInt32($dataTypePtr, 0)
            [Marshal]::WriteInt32($valueSizePtr, 0)
            [Marshal]::WriteIntPtr($ptrPtr, [IntPtr]::Zero)

            $res = $global:SLC::SLGetApplicationInformation(
                $hSLC,
                [ref]$appGuid,
                $valueName,
                $dataTypePtr,
                $valueSizePtr,
                $ptrPtr
            )

            if ($res -ne 0) {
                continue
            }

            # Read the outputs from the unmanaged memory pointers
            $dataType = [Marshal]::ReadInt32($dataTypePtr)
            $valueSize = [Marshal]::ReadInt32($valueSizePtr)
        
            # Dereference the pointer-to-pointer to get actual buffer pointer
            $ptr = [Marshal]::ReadIntPtr($ptrPtr)

            if ($ptr -eq [IntPtr]::Zero) {
                continue
            }

            if ($valueSize -eq 0) {
                continue
            }

            $allValues[$valueName] = Parse-RegistryData -dataType $dataType -ptr $ptr -valueSize $valueSize -valueName $valueName
            Free-IntPtr -handle $ptr -Method Local
        }
        return $allValues
    }


    # Allocate memory for dataType (optional out parameter)
    $dataTypePtr = [Marshal]::AllocHGlobal(4)
    # Allocate memory for valueSize (UINT* out param)
    $valueSizePtr = [Marshal]::AllocHGlobal(4)
    # Allocate memory for pointer to byte buffer (PBYTE* out param)
    $ptrPtr = [Marshal]::AllocHGlobal([IntPtr]::Size)

    try {
        # Initialize the out params to zero/null
        [Marshal]::WriteInt32($dataTypePtr, 0)
        [Marshal]::WriteInt32($valueSizePtr, 0)
        [Marshal]::WriteIntPtr($ptrPtr, [IntPtr]::Zero)

        $hresult = $global:SLC::SLGetApplicationInformation(
            $hSLC,
            [ref]$appGuid,
            $PropertyName,
            $dataTypePtr,
            $valueSizePtr,
            $ptrPtr
        )

        if ($hresult -ne 0) {
            $errorMessege = Parse-ErrorMessage -MessageId $hresult -Flags ([ErrorMessageType]::ACTIVATION -bor [ErrorMessageType]::HRESULT)
            Write-Warning "$errorMessege, $hresult"
            throw "SLGetApplicationInformation failed (code $hresult)"
        }

        # Read the outputs from the unmanaged memory pointers
        $dataType = [Marshal]::ReadInt32($dataTypePtr)
        $valueSize = [Marshal]::ReadInt32($valueSizePtr)
        
        # Dereference the pointer-to-pointer to get actual buffer pointer
        $ptr = [Marshal]::ReadIntPtr($ptrPtr)

        if ($ptr -eq [IntPtr]::Zero) {
            throw "Pointer to data buffer is null"
        }

        if ($valueSize -eq 0) {
            throw "Returned value size is zero"
        }

        try {
            return Parse-RegistryData -dataType $dataType -ptr $ptr -valueSize $valueSize -valueName $PropertyName
        }
        finally {
            Free-IntPtr -handle $ptr -Method Local
        }
    }
    finally {
        if ($closeHandle) {
            Write-Warning "Consider Open handle Using Manage-SLHandle"
            Free-IntPtr -handle $hSLC -Method License
        }
        if ($dataTypePtr -and $dataTypePtr -ne [IntPtr]::Zero) {
            [Marshal]::FreeHGlobal($dataTypePtr)
        }
        if ($valueSizePtr -and $valueSizePtr -ne [IntPtr]::Zero) {
            [Marshal]::FreeHGlobal($valueSizePtr)
        }
        if ($ptrPtr -and $ptrPtr -ne [IntPtr]::Zero) {
            [Marshal]::FreeHGlobal($ptrPtr)
        }
    }
}

<#
.SYNOPSIS
Gets information about the specified product SKU.

"Description", "Name", "Author", 
Taken from Microsoft Offical Documentation.

# ----------------------------------------------

fileId                # SL_ID_LICENSE_FILE
pkeyId                # SL_ID_PKEY
productSkuId          # SL_ID_PRODUCT_SKU
applicationId         # SL_ID_APPLICATION
licenseId             # SL_ID_LICENSE 
privateCertificateId  # SL_ID_LICENSE 

# ------>>>> More info ------>>>

https://github.com/LBBNetwork/openredpill/blob/master/slpublic.h
https://learn.microsoft.com/en-us/windows/win32/api/slpublic/nf-slpublic-slgetslidlist

SL_ID_APPLICATION,  appId        X
SL_ID_PRODUCT_SKU,  skuId        ?
SL_ID_LICENSE_FILE, fileId       V
SL_ID_LICENSE,      LicenseId    V

# ----------------------------------------------

"msft:sl/EUL/GENERIC/PUBLIC",    "msft:sl/EUL/GENERIC/PRIVATE",
"msft:sl/EUL/PHONE/PUBLIC",      "msft:sl/EUL/PHONE/PRIVATE",
"msft:sl/EUL/STORE/PUBLIC",      "msft:sl/EUL/STORE/PRIVATE",
"msft:sl/EUL/ACTIVATED/PRIVATE", "msft:sl/EUL/ACTIVATED/PUBLIC",

# ----------------------------------------------

Also, 
if you read String's from sppwmi.dll [hint by abbody1406]
you will find more data.

!Jackpot!
inside sppsvc.exe, lot of data, we can search,
include fileId & more info.

# ----------------------------------------------

Also, some properties from *SoftwareLicensingProduct* WMI Class
can be enum too, some with diffrent name.

class SoftwareLicensingProduct
{
  string   ID;                                            --> Function * SLGetProductSkuInformation  --> productSkuId
  string   Name;                                          --> Function * SLGetProductSkuInformation 
  string   Description;                                   --> Function * SLGetProductSkuInformation 
  string   ApplicationID;                                 --> Function * SLGetProductSkuInformation 
  string   ProcessorURL;                                  --> Function * SLGetProductSkuInformation 
  string   MachineURL;                                    --> Function * SLGetProductSkuInformation 
  string   ProductKeyURL;                                 --> Function * SLGetProductSkuInformation 
  
  sppcomapi.dll
  __int64 __fastcall SPPGetServerAddresses(HSLC hSLC, struct SActivationServerAddress **a2, unsigned int *a3)
  v58[0] = (__int64)L"SPCURL"; // GetProcessorURL
  v58[1] = (__int64)L"RACURL"; // GetMachineURL
  v51 = L"PAURL";              // GetUseLicenseURL
  v58[2] = (__int64)L"PKCURL"; // GetProductKeyURL
  v58[3] = (__int64)L"EULURL"; // GetUseLicenseURL

  string   UseLicenseURL;                                 --> Function * SLGetProductSkuInformation  --> PAUrl [-or EULURL, By abbody1406]

  uint32   LicenseStatus;                                 --> Function * SLGetLicensingStatusInformation
  uint32   LicenseStatusReason;                           --> Function * SLGetLicensingStatusInformation
  uint32   GracePeriodRemaining;                          --> Function * SLGetLicensingStatusInformation
  datetime EvaluationEndDate;                             --> Function * SLGetLicensingStatusInformation
  string   OfflineInstallationId;                         --> Function * SLGenerateOfflineInstallationId
  string   PartialProductKey;                             --> Function * SLGetPKeyInformation
  string   ProductKeyID;                                  --> Function * SLGetPKeyInformation
  string   ProductKeyID2;                                 --> Function * SLGetPKeyInformation
  string   ProductKeyChannel;                             --> Function * SLGetPKeyInformation
  string   LicenseFamily;                                 --> Function * SLGetProductSkuInformation  --> Family
  string   LicenseDependsOn;                              --> Function * SLGetProductSkuInformation  --> DependsOn
  string   ValidationURL;                                 --> Function * SLGetProductSkuInformation  --> ValUrl
  boolean  LicenseIsAddon;                                --> Function * SLGetProductSkuInformation  --> [BOOL](DependsOn) // From TSforge project
  uint32   VLActivationInterval;                          --> Function * SLGetProductSkuInformation
  uint32   VLRenewalInterval;                             --> Function * SLGetProductSkuInformation
  string   KeyManagementServiceProductKeyID;              --> Function * SLGetProductSkuInformation  --> CustomerPID
  string   KeyManagementServiceMachine;                   --> Function * SLGetProductSkuInformation  --> KeyManagementServiceName
  uint32   KeyManagementServicePort;                      --> Function * SLGetProductSkuInformation  --> KeyManagementServicePort
  string   DiscoveredKeyManagementServiceMachineName;     --> Function * SLGetProductSkuInformation  --> DiscoveredKeyManagementServiceName
  uint32   DiscoveredKeyManagementServiceMachinePort;     --> Function * SLGetProductSkuInformation  --> DiscoveredKeyManagementServicePort
  BOOL     IsKeyManagementServiceMachine;                 --> Function * SLGetApplicationInformation --> Key: "IsKeyManagementService"                      (SL_INFO_KEY_IS_KMS)
  uint32   KeyManagementServiceCurrentCount;              --> Function * SLGetApplicationInformation --> Key: "KeyManagementServiceCurrentCount"            (SL_INFO_KEY_KMS_CURRENT_COUNT)
  uint32   RequiredClientCount;                           --> Function * SLGetApplicationInformation --> Key: "KeyManagementServiceRequiredClientCount"     (SL_INFO_KEY_KMS_REQUIRED_CLIENT_COUNT)
  uint32   KeyManagementServiceUnlicensedRequests;        --> Function * SLGetApplicationInformation --> Key: "KeyManagementServiceUnlicensedRequests"      (SL_INFO_KEY_KMS_UNLICENSED_REQUESTS)
  uint32   KeyManagementServiceLicensedRequests;          --> Function * SLGetApplicationInformation --> Key: "KeyManagementServiceLicensedRequests"        (SL_INFO_KEY_KMS_LICENSED_REQUESTS)
  uint32   KeyManagementServiceOOBGraceRequests;          --> Function * SLGetApplicationInformation --> Key: "KeyManagementServiceOOBGraceRequests"        (SL_INFO_KEY_KMS_OOB_GRACE_REQUESTS)
  uint32   KeyManagementServiceOOTGraceRequests;          --> Function * SLGetApplicationInformation --> Key: "KeyManagementServiceOOTGraceRequests"        (SL_INFO_KEY_KMS_OOT_GRACE_REQUESTS)
  uint32   KeyManagementServiceNonGenuineGraceRequests;   --> Function * SLGetApplicationInformation --> Key: "KeyManagementServiceNonGenuineGraceRequests" (SL_INFO_KEY_KMS_NON_GENUINE_GRACE_REQUESTS)
  uint32   KeyManagementServiceTotalRequests;             --> Function * SLGetApplicationInformation --> Key: "KeyManagementServiceTotalRequests"           (SL_INFO_KEY_KMS_TOTAL_REQUESTS)
  uint32   KeyManagementServiceFailedRequests;            --> Function * SLGetApplicationInformation --> Key: "KeyManagementServiceFailedRequests"          (SL_INFO_KEY_KMS_FAILED_REQUESTS)
  uint32   KeyManagementServiceNotificationRequests;      --> Function * SLGetApplicationInformation --> Key: "KeyManagementServiceNotificationRequests"    (SL_INFO_KEY_KMS_NOTIFICATION_REQUESTS)
  uint32   GenuineStatus;                                 --> Function * SLGetLicensingStatusInformation
  uint32   ExtendedGrace;                                 --> Function * SLGetProductSkuInformation  --> TimeBasedExtendedGrace
  string   TokenActivationILID;                           --> Function * SLGetProductSkuInformation
  uint32   TokenActivationILVID;                          --> Function * SLGetProductSkuInformation
  uint32   TokenActivationGrantNumber;                    --> Function * SLGetProductSkuInformation 
  string   TokenActivationCertificateThumbprint;          --> Function * SLGetProductSkuInformation
  string   TokenActivationAdditionalInfo;                 --> Function * SLGetProductSkuInformation
  datetime TrustedTime;                                   --> Function * SLGetProductSkuInformation [Licensing System Date]
};

# ----------------------------------------------

Now found that it read --> r:otherInfo Section
So it support for Any <tm:infoStr name=

<r:otherInfo xmlns:r="urn:mpeg:mpeg21:2003:01-REL-R-NS">
	<tm:infoTables xmlns:tm="http://www.microsoft.com/DRM/XrML2/TM/v2">
		<tm:infoList tag="#global">
		<tm:infoStr name="licenseType">msft:sl/PL/GENERIC/PUBLIC</tm:infoStr>
		<tm:infoStr name="licenseVersion">2.0</tm:infoStr>
		<tm:infoStr name="licensorUrl">http://licensing.microsoft.com</tm:infoStr>
		<tm:infoStr name="licenseCategory">msft:sl/PL/GENERIC/PUBLIC</tm:infoStr>
		<tm:infoStr name="productSkuId">{2c060131-0e43-4e01-adc1-cf5ad1100da8}</tm:infoStr>
		<tm:infoStr name="privateCertificateId">{274ff0e9-dfec-43e7-b675-67e61645b6a9}</tm:infoStr>
		<tm:infoStr name="applicationId">{55c92734-d682-4d71-983e-d6ec3f16059f}</tm:infoStr>
		<tm:infoStr name="productName">Windows(R), EnterpriseSN edition</tm:infoStr>
		<tm:infoStr name="Family">EnterpriseSN</tm:infoStr>
		<tm:infoStr name="productAuthor">Microsoft Corporation</tm:infoStr>
		<tm:infoStr name="productDescription">Windows(R) Operating System</tm:infoStr>
		<tm:infoStr name="clientIssuanceCertificateId">{4961cc30-d690-43be-910c-8e2db01fc5ad}</tm:infoStr>
		<tm:infoStr name="hwid:ootGrace">0</tm:infoStr>
		</tm:infoList>
	</tm:infoTables>
</r:otherInfo>
<r:otherInfo xmlns:r="urn:mpeg:mpeg21:2003:01-REL-R-NS">
	<tm:infoTables xmlns:tm="http://www.microsoft.com/DRM/XrML2/TM/v2">
		<tm:infoList tag="#global">
		<tm:infoStr name="licenseType">msft:sl/PL/GENERIC/PRIVATE</tm:infoStr>
		<tm:infoStr name="licenseVersion">2.0</tm:infoStr>
		<tm:infoStr name="licensorUrl">http://licensing.microsoft.com</tm:infoStr>
		<tm:infoStr name="licenseCategory">msft:sl/PL/GENERIC/PRIVATE</tm:infoStr>
		<tm:infoStr name="publicCertificateId">{0f6421d2-b7ea-45e0-b87d-773975685c35}</tm:infoStr>
		<tm:infoStr name="clientIssuanceCertificateId">{4961cc30-d690-43be-910c-8e2db01fc5ad}</tm:infoStr>
		<tm:infoStr name="hwid:ootGrace">0</tm:infoStr>
		<tm:infoStr name="win:branding">126</tm:infoStr>
		</tm:infoList>
	</tm:infoTables>
</r:otherInfo>
#>
<#
Clear-host
$WMI_QUERY = Get-SLIDList -eQueryIdType SL_ID_PRODUCT_SKU -eReturnIdType SL_ID_PRODUCT_SKU
$XML_Table = $WMI_QUERY | % { Get-LicenseDetails $_ -ReturnRawData}

write-host
$uniqueNamesHashTable = @{}
foreach ($xmlEntry in $XML_Table) {
    # Ensure the licenseGroup property exists on the current object
    if ($xmlEntry.licenseGroup) {
        # The 'license' property within licenseGroup can be a single object or an array of objects.
        # Use @() to ensure it's treated as an array, making iteration consistent.
        foreach ($licenseObject in @($xmlEntry.licenseGroup.license)) {
            # Safely navigate the object path to get to the 'infoStr' elements.
            # We check each step to prevent errors if a property is missing.
            if ($licenseObject.otherInfo -and
                $licenseObject.otherInfo.infoTables -and
                $licenseObject.otherInfo.infoTables.infoList -and
                $licenseObject.otherInfo.infoTables.infoList.infoStr) {

                # Extract all 'name' attributes from the 'infoStr' elements for the current license.
                # If infoStr is a single object, .name will work. If it's an array, it will automatically
                # collect all names.
                $names = $licenseObject.otherInfo.infoTables.infoList.infoStr.name

                # Add each extracted name to our hashtable.
                # The 'ContainsKey' check ensures that only unique names are added.
                foreach ($name in @($names)) { # @($names) ensures we iterate even if $names is a single string
                    if (-not $uniqueNamesHashTable.ContainsKey($name)) {
                        $uniqueNamesHashTable[$name] = $true # Value can be anything; we only care about the key
                    }
                }
            }
        }
    }
}
$uniqueNamesHashTable.keys

$xmlStrings = $XML_Table | ForEach-Object { $_.OuterXml }
$regexPattern = ">(msft.*?)<"
$extractedContent = New-Object System.Collections.ArrayList # Use ArrayList for efficient adding
foreach ($xmlString in $xmlStrings) {
    $matchesInCurrentString = [regex]::Matches($xmlString, $regexPattern)
    if ($matchesInCurrentString.Count -gt 0) {
        foreach ($match in $matchesInCurrentString) {
            # Add the content of the first capturing group (index 1) to our ArrayList
            # $match.Groups[0] would be the entire match (>msft...<)
            # $match.Groups[1] is the content of the first capturing group (msft...)
            [void]$extractedContent.Add($match.Groups[1].Value)
        }
    }
}

write-host
if ($extractedContent.Count -gt 0) {
    $extractedContent | Select-Object -Unique | Sort-Object | ForEach-Object {
        Write-Host $_ }}
Read-Host
#>
function Get-LicenseDetails {
    param (
        [Parameter(Mandatory)]
        [Guid]$ActConfigId,

        [Parameter(Mandatory = $false)]
        [ValidateSet(
        "fileId", "pkeyId", "productSkuId", "applicationId",
        "licenseId", "privateCertificateId", "pkeyIdList",

        "msft:sl/EUL/GENERIC/PUBLIC", "msft:sl/EUL/GENERIC/PRIVATE",
        "msft:sl/EUL/PHONE/PUBLIC", "msft:sl/EUL/PHONE/PRIVATE",
        "msft:sl/EUL/STORE/PUBLIC", "msft:sl/EUL/STORE/PRIVATE",
        "msft:sl/EUL/ACTIVATED/PRIVATE", "msft:sl/EUL/ACTIVATED/PUBLIC",
        "msft:sl/PL/GENERIC/PUBLIC",    "msft:sl/PL/GENERIC/PRIVATE",

        "Description", "Name", "Author",
		
        "TokenActivationILID", "TokenActivationILVID","TokenActivationGrantNumber",
        "TokenActivationCertificateThumbprint", "TokenActivationAdditionalInfo",
        "pkeyConfigLicenseId", "licenseType", "licenseVersion", "licensorUrl", "licenseNamespace",
        "productName", "Family", "productAuthor", "productDescription", "licenseCategory",
        "hwid:ootGrace", "issuanceCertificateId", "ValUrl", "PAUrl", "ActivationSequence", 
        "UXDifferentiator", "ProductKeyGroupUniqueness", "EnableNotificationMode", "EULURL", 
        "GraceTimerUniqueness", "ValidityTimerUniqueness", "EnableActivationValidation", "PKCURL",
        "DependsOn", "phone:policy", "licensorKeyIndex", "BuildVersion", "ValidationTemplateId",
        "ProductUniquenessGroupId", "ApplicationBitmap", "migratable",
        "ProductKeyID", "VLActivationInterval", "VLRenewalInterval", "KeyManagementServiceProductKeyID", 
        "KeyManagementServicePort", "TrustedTime", "CustomerPID", "KeyManagementServiceName", 
        "KeyManagementServicePort", "DiscoveredKeyManagementServiceName", 
        "DiscoveredKeyManagementServicePort", "TimeBasedExtendedGrace",
        "ADActivationObjectDN", "ADActivationObjectName", "DiscoveredKeyManagementServiceIpAddress",
        "KeyManagementServiceLookupDomain", "RemainingRearmCount", "VLActivationType",
        "TokenActivationCertThumbprint", "RearmCount", "ADActivationCsvlkPID", "ADActivationCsvlkSkuID",
        "fileIndex", "licenseDescription", "metaInfoType", "DigitalEncryptedPID",
        "InheritedActivationId", "InheritedActivationHostMachineName", "InheritedActivationHostDigitalPid2",
        "InheritedActivationActivationTime"
        )]
        [String]$pwszValueName,

        [Parameter(Mandatory = $false)]
        [switch]$loopAllValues,

        [Parameter(Mandatory = $false)]
        [switch]$ReturnRawData,

        [Parameter(Mandatory=$false)]
        [Intptr]$hSLC = [IntPtr]::Zero
    )

    # 3 CASES OF XOR
    if (@([BOOL]$pwszValueName + [BOOL]$loopAllValues + [BOOL]$ReturnRawData) -ne 1) {
        Write-Warning "Exactly one of -pwszValueName, -loopAllValues, or -ReturnRawData must be specified."
        return
    }

    if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
        $hSLC = if ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
            $global:hSLC_
        } else {
            Manage-SLHandle
        }
    }

    try {
        $closeHandle = $true
        if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
            $hr = $Global:SLC::SLOpen([ref]$hSLC)
            if ($hr -ne 0) {
                throw "SLOpen failed: HRESULT 0x{0:X8}" -f $hr
            }
        } else {
            $closeHandle = $false
        }
    }
    catch {
        return $null
    }

    try {
        if ($loopAllValues) {
            $allValues = @{}
            $SL_ID = @(
                "fileId",               # SL_ID_LICENSE_FILE
                "pkeyId",               # SL_ID_PKEY
                "pkeyIdList"            # SL_ID_PKEY [Same]
                "productSkuId",         # SL_ID_PRODUCT_SKU
                "applicationId"         # SL_ID_APPLICATION
                "licenseId",            # SL_ID_LICENSE
                "privateCertificateId"  # SL_ID_LICENSE
            )

            # un offical, intersting pattern
            # first saw in MAS AIO file, TSforge project}
            $MSFT = @(
                "msft:sl/EUL/GENERIC/PUBLIC",    "msft:sl/EUL/GENERIC/PRIVATE",    # un-offical
                "msft:sl/EUL/PHONE/PUBLIC",      "msft:sl/EUL/PHONE/PRIVATE",      # un-offical
                "msft:sl/EUL/STORE/PUBLIC",      "msft:sl/EUL/STORE/PRIVATE",      # un-offical
                "msft:sl/EUL/ACTIVATED/PRIVATE", "msft:sl/EUL/ACTIVATED/PUBLIC",   # extract from SPP* dll/exe files
                "msft:sl/PL/GENERIC/PUBLIC",    "msft:sl/PL/GENERIC/PRIVATE"       # extract from SPP* dll/exe files
            )
            
            # Offical from MS
            $OfficalPattern = @("Description", "Name", "Author")

            # the rest, from XML Blobs --> <infoStr>
            $xml = @("pkeyConfigLicenseId", "privateCertificateId", "licenseType", 
                "licensorUrl",  "licenseCategory", "productName", "Family","licenseVersion",
                "productAuthor", "productDescription",  "hwid:ootGrace", "issuanceCertificateId", "PAUrl",
                "ActivationSequence", "ValidationTemplateId", "ValUrl", "UXDifferentiator",
				"ProductKeyGroupUniqueness", "EnableNotificationMode", "GraceTimerUniqueness",
                "ValidityTimerUniqueness", "EnableActivationValidation",
                "DependsOn", "phone:policy", "licensorKeyIndex", "BuildVersion",
                "ProductUniquenessGroupId", "ApplicationBitmap", "migratable")
            
            # SoftwareLicensingProduct class (WMI)
            $SoftwareLicensingProduct = @(
                "Name",  "Description", "ApplicationID", "VLActivationInterval", "VLRenewalInterval",
                "ProductKeyID",  "KeyManagementServiceProductKeyID", "KeyManagementServicePort", 
                "RequiredClientCount", "TrustedTime", "TokenActivationILID", "TokenActivationILVID",
                "TokenActivationGrantNumber", "TokenActivationCertificateThumbprint", "CustomerPID",
                "KeyManagementServiceName", "KeyManagementServicePort","TimeBasedExtendedGrace", 
                "DiscoveredKeyManagementServiceName", "DiscoveredKeyManagementServicePort",
                "PKCURL", "EULURL"
            )

            # SPP* DLL/EXE file's
            $sppwmi = @(
                "ADActivationObjectDN", "ADActivationObjectName", "DiscoveredKeyManagementServiceIpAddress",
                "KeyManagementServiceLookupDomain", "RemainingRearmCount", "TokenActivationAdditionalInfo",
                "TokenActivationCertThumbprint", "VLActivationType", "RearmCount", "ADActivationCsvlkPID", 
                "fileIndex", "licenseDescription", "metaInfoType", "DigitalEncryptedPID", "ADActivationCsvlkSkuID",
				"InheritedActivationId", "InheritedActivationHostMachineName", "InheritedActivationHostDigitalPid2",
				"InheritedActivationActivationTime", "licenseNamespace"
            )

            # Combine all arrays and remove duplicates
            $allValueNames = ($SL_ID + $MSFT + $OfficalPattern + $xml + $SoftwareLicensingProduct + $sppwmi) | Sort-Object -Unique

            foreach ($valueName in $allValueNames) {
                $dataType = 0
                $valueSize = 0
                $ptr = [IntPtr]::Zero
                $res = $global:SLC::SLGetProductSkuInformation(
                    $hSLC,
                    [ref]$ActConfigId,
                    $valueName,
                    [ref]$dataType,
                    [ref]$valueSize,
                    [ref]$ptr
                )

                if ($res -ne 0) {
                    #Write-Warning "fail to process Name: $valueName"
                    $allValues[$valueName] = $null
                    continue;
                }
 
                $allValues[$valueName] = Parse-RegistryData -dataType $dataType -ptr $ptr -valueSize $valueSize -valueName $valueName
                Free-IntPtr -handle $ptr -Method Local
            }

            return $allValues
        }

        if ($pwszValueName) {
            $dataType = 0
            $valueSize = 0
            $ptr = [IntPtr]::Zero
            $res = $global:SLC::SLGetProductSkuInformation(
                $hSLC,
                [ref]$ActConfigId,
                $pwszValueName,
                [ref]$dataType,
                [ref]$valueSize,
                [ref]$ptr
            )

            if ($res -ne 0) {
                #$messege = $(Parse-ErrorMessage -MessageId $res)
                #Write-Warning "ERROR $res, $messege, Value: $pwszValueName"
                throw
            }

            try {
                return Parse-RegistryData -dataType $dataType -ptr $ptr -valueSize $valueSize -valueName $pwszValueName
            }
            finally {
                Free-IntPtr -handle $ptr -Method Local
            }
        }

        try {
            $content = Get-LicenseData -SkuID $ActConfigId -Mode License
            $xmlContent = $content.Substring($content.IndexOf('<r'))
            $xml = [xml]$xmlContent
        }
        catch {
        }

        if ($ReturnRawData) {
            return $xml
        }

        # Transform into detailed custom objects
        $licenseObjects = @()

        foreach ($license in $xml.licenseGroup.license) {
            $policyList = @()
            foreach ($policy in $license.grant.allConditions.allConditions.productPolicies.policyStr) {
                $policyList += [PSCustomObject]@{
                    Name  = $policy.name
                    Value = $policy.InnerText
                }
            }

            if (-not $policyList) {
                continue;
            }

            $licenseObjects += [PSCustomObject]@{
                LicenseId  = $license.licenseId
                GrantName  = $license.grant.name
                Policies   = $policyList
            }
        }

        return $licenseObjects
    }
    catch { }
    finally {
        if ($closeHandle) {
            Write-Warning "Consider Open handle Using Manage-SLHandle"
            Free-IntPtr -handle $hSLC -Method License
        }
    }
}

<#
.SYNOPSIS
    Retrieves various system and service information,
    based on the specified value name or fetches all available information if requested.

.Source
   sppwmi.dll.! 
   *GetServiceInformation*

.Usage

Example Code:
~~~~~~~~~~~~

Get-ServiceInfo -loopAllValues
Get-ServiceInfo -pwszValueName SecureStoreId

~~~~~~~~~~~~~~~~~~~~~~~~~

Clear-Host

Write-Host
Write-Host "Get-OA3xOriginalProductKey" -ForegroundColor Green
Get-OA3xOriginalProductKey

Write-Host
Write-Host "Get-ServiceInfo" -ForegroundColor Green
Get-ServiceInfo -loopAllValues | Format-Table -AutoSize

Write-Host
Write-Host "Get-ActiveLicenseInfo" -ForegroundColor Green
Get-ActiveLicenseInfo | Format-List
#>
function Get-ServiceInfo {
    param (
        [Parameter(Mandatory = $false)]
        [ValidateSet(
            "ActivePlugins", "CustomerPID", "SystemState", "Version",
            "BiosOA2MinorVersion", "BiosProductKey", "BiosSlicState",
            "BiosProductKeyDescription", "BiosProductKeyPkPn",
            "ClientMachineID", "SecureStoreId", "SessionMachineId",
            "DiscoveredKeyManagementPort",
            "DiscoveredKeyManagementServicePort",
            "DiscoveredKeyManagementServiceIpAddress",
            "DiscoveredKeyManagementServiceName",
            "IsKeyManagementService",
            "KeyManagementServiceCurrentCount",
            "KeyManagementServiceFailedRequests",
            "KeyManagementServiceLicensedRequests",
            "KeyManagementServiceNonGenuineGraceRequests",
            "KeyManagementServiceNotificationRequests",
            "KeyManagementServiceOOBGraceRequests",
            "KeyManagementServiceOOTGraceRequests",
            "KeyManagementServiceRequiredClientCount",
            "KeyManagementServiceTotalRequests",
            "KeyManagementServiceUnlicensedRequests",
            "TokenActivationAdditionalInfo",
            "TokenActivationCertThumbprint",
            "TokenActivationGrantNumber",
            "TokenActivationILID",
            "TokenActivationILVID"
        )]
        [String]$pwszValueName,

        [Parameter(Mandatory = $false)]
        [switch]$loopAllValues,

        [Parameter(Mandatory=$false)]
        [Intptr]$hSLC = [IntPtr]::Zero
    )

    # !Xor Case
    if (!($pwszValueName -xor [BOOL]$loopAllValues)) {
        Write-Warning "Exactly one of -pwszValueName, -loopAllValues, must be specified."
        return
    }

    if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
        $hSLC = if ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
            $global:hSLC_
        } else {
            Manage-SLHandle
        }
    }

    try {
        $closeHandle = $true
        if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
            $hr = $Global:SLC::SLOpen([ref]$hSLC)
            if ($hr -ne 0) {
                throw "SLOpen failed: HRESULT 0x{0:X8}" -f $hr
            }
        } else {
            $closeHandle = $false
        }
    }
    catch {
        return $null
    }

    try {
        $allValues = @{}
        if ($loopAllValues) {
            $allValueNames = @(
                "ActivePlugins", "CustomerPID", "SystemState", "Version",
                "BiosOA2MinorVersion", "BiosProductKey", "BiosSlicState",
                "BiosProductKeyDescription", "BiosProductKeyPkPn",
                "ClientMachineID", "SecureStoreId", "SessionMachineId",
                "DiscoveredKeyManagementPort",
                "DiscoveredKeyManagementServicePort",
                "DiscoveredKeyManagementServiceIpAddress",
                "DiscoveredKeyManagementServiceName",
                "IsKeyManagementService",
                "KeyManagementServiceCurrentCount",
                "KeyManagementServiceFailedRequests",
                "KeyManagementServiceLicensedRequests",
                "KeyManagementServiceNonGenuineGraceRequests",
                "KeyManagementServiceNotificationRequests",
                "KeyManagementServiceOOBGraceRequests",
                "KeyManagementServiceOOTGraceRequests",
                "KeyManagementServiceRequiredClientCount",
                "KeyManagementServiceTotalRequests",
                "KeyManagementServiceUnlicensedRequests",
                "TokenActivationAdditionalInfo",
                "TokenActivationCertThumbprint",
                "TokenActivationGrantNumber",
                "TokenActivationILID",
                "TokenActivationILVID"
            )

            foreach ($valueName in $allValueNames) {
                $dataType = 0
                $valueSize = 0
                $ptr = [IntPtr]::Zero
                $res = $global:SLC::SLGetServiceInformation(
                    $hSLC,
                    $valueName,
                    [ref]$dataType,
                    [ref]$valueSize,
                    [ref]$ptr
                )

                if ($res -ne 0) {
                    #Write-Warning "fail to process Name: $valueName"
                    $allValues[$valueName] = $null
                    continue;
                }
 
                $allValues[$valueName] = Parse-RegistryData -dataType $dataType -ptr $ptr -valueSize $valueSize -valueName $valueName
                Free-IntPtr -handle $ptr -Method Local
            }

            return $allValues
        }

        if ($pwszValueName) {
            $dataType = 0
            $valueSize = 0
            $ptr = [IntPtr]::Zero
            $res = $global:SLC::SLGetServiceInformation(
                $hSLC,
                $pwszValueName,
                [ref]$dataType,
                [ref]$valueSize,
                [ref]$ptr
            )

            if ($res -ne 0) {
                #$messege = $(Parse-ErrorMessage -MessageId $res)
                #Write-Warning "ERROR $res, $messege, Value: $pwszValueName"
                throw
            }

            # Parse value based on data type
            try {
                return Parse-RegistryData -dataType $dataType -ptr $ptr -valueSize $valueSize -valueName $pwszValueName
            }
            finally {
                Free-IntPtr -handle $ptr -Method Local
            }
        }
    }
    catch { }
    finally {
        if ($closeHandle) {
            Write-Warning "Consider Open handle Using Manage-SLHandle"
            Free-IntPtr -handle $hSLC -Method License
        }
    }
}

<#
.SYNOPSIS
Get active information using SLGetActiveLicenseInfo API
return value is struct DigitalProductId4

Also, Parse-DigitalProductId4 function, read same results just from registry
"HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -> DigitalProductId4 [Propertie]

*** not always, in case of ==> Get-OA3xOriginalProductKey == TRUE
Get-ActiveLicenseInfo ==> IS EQAL too ==>
$GroupID = Decode-Key -Key (Get-OA3xOriginalProductKey) | Select-Object -ExpandProperty Group
$Global:PKeyDatabase | ? RefGroupId -eq $GroupID | Select-Object -First 1

it also mask the key, like it was mak key [it's bios key !]
and, not even active, so, why mark it as active ?

.Usage

Example Code:
~~~~~~~~~~~~

Clear-Host

Write-Host
Write-Host "Get-OA3xOriginalProductKey" -ForegroundColor Green
Get-OA3xOriginalProductKey

Write-Host
Write-Host "Get-ServiceInfo" -ForegroundColor Green
Get-ServiceInfo -loopAllValues | Format-Table -AutoSize

Write-Host
Write-Host "Get-ActiveLicenseInfo" -ForegroundColor Green
Get-ActiveLicenseInfo | Format-List
#>
function Get-ActiveLicenseInfo {
    param (
        [Guid]$SkuID = [Guid]::Empty
    )

    # Initialize all pointers to zero
    $hSLC = $ptr = $contextPtr = [IntPtr]::Zero
    $size = 0

    try {
        # If GUID is provided, allocate and copy
        if ($SkuID -ne [Guid]::Empty) {
            $guidBytes = $SkuID.ToByteArray()
            $contextPtr = New-IntPtr -Size 16
            [Marshal]::Copy($guidBytes, 0, $contextPtr, 16)
        }

        Manage-SLHandle -Release | Out-Null
        $hSLC = Manage-SLHandle

        if ($hSLC -eq [IntPtr]::Zero) {
            throw "Fail to get handle from SLOPEN API"
        }

        $res = $Global:SLC::SLGetActiveLicenseInfo($hSLC, $contextPtr, [ref]$size, [ref]$ptr)
        if ($res -ne 0 -or $ptr -eq [IntPtr]::Zero -or $size -le 0) {
            $ErrorMessage = Parse-ErrorMessage -MessageId $res
            throw "SLGetActiveLicenseInfo failed with HRESULT: $res, $ErrorMessage."
        }

        if ($size -lt 1280) {
            throw "Returned license buffer too small ($size bytes). Expected >= 1280 for DigitalProductId4."
        }
        
        Parse-DigitalProductId4 -Pointer $ptr -Length $size -FromIntPtr
    }
    catch {
        Write-Warning "Failed to get active license info: $_"
    }
    finally {
        if ($hSLC -ne [IntPtr]::Zero) {
            Manage-SLHandle -Release | Out-Null
        }
        if ($contextPtr -ne [IntPtr]::Zero) {
            New-IntPtr -hHandle $contextPtr -Release
        }
        if ($ptr -ne [IntPtr]::Zero) {
            $Global:kernel32::LocalFree($ptr) | Out-Null
        }
    }
}

<#
.SYNOPSIS
Get Info per license, using Pkeyconfig [XML] & low level API

$WMI_QUERY = Get-SLIDList -eQueryIdType SL_ID_PRODUCT_SKU -eReturnIdType SL_ID_PRODUCT_SKU
$WMI_SQL = $WMI_QUERY | % { Get-LicenseInfo -ActConfigId $_ }
$WMIINFO = $WMI_SQL | Select-Object * -ExcludeProperty Policies | ? EditionID -NotMatch 'ESU'
Manage-SLHandle -Release | Out-null
#>
function Get-LicenseInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string] $ActConfigId,

        [Parameter(Mandatory=$false)]
        [Intptr]$hSLC = [IntPtr]::Zero
    )
    function Get-BrandingValue {
        param (
            [Parameter(Mandatory=$true)]
            [guid]$sku
        )

        try {

            # Fetch license details for the SKU
            $xml = Get-LicenseDetails -ActConfigId $sku -ReturnRawData -hSLC $hSLC
            if (-not $xml) {
                return;  }

            $BrandingValue = $xml.licenseGroup.license[1].otherInfo.infoTables.infoList.infoStr | Where-Object Name -EQ 'win:branding'
            return $BrandingValue.'#text'

            #$match = $Global:productTypeTable | Where-Object {
            #    [Convert]::ToInt32($_.DWORD, 16) -eq $BrandingValue.'#text'
            #}
            #return $match.ProductID

        } catch {
            Write-Warning "An error occurred: $_"
            return $null
        }
    }
    Function Get-KeyManagementServiceInfo {
        param (
            [Parameter(Mandatory=$true)]
            [STRING]$SKU_ID
        )

        if ([STRING]::IsNullOrWhiteSpace($SKU_ID) -or (
        $SKU_ID -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')) {
            Return @();
        }

        $Base = "HKLM:SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
        $Application_ID = Retrieve-SKUInfo -SkuId 7103a333-b8c8-49cc-93ce-d37c09687f92 -eReturnIdType SL_ID_APPLICATION | select -ExpandProperty Guid

        if ($Application_ID) {
            $KeyManagementServiceName = Get-ItemProperty -Path "$Base\$Application_ID\$SKU_ID" -Name "KeyManagementServiceName" -ea 0 | select -ExpandProperty KeyManagementServiceName
            $KeyManagementServicePort = Get-ItemProperty -Path "$Base\$Application_ID\$SKU_ID" -Name "KeyManagementServicePort" -ea 0 | select -ExpandProperty KeyManagementServicePort
        }
        if (-not $KeyManagementServiceName) {
            $KeyManagementServiceName = Get-ItemProperty -Path "$Base" -Name "KeyManagementServiceName" -ea 0 | select -ExpandProperty KeyManagementServiceName
        }
        if (-not $KeyManagementServicePort) {
            $KeyManagementServicePort = Get-ItemProperty -Path "$Base" -Name "KeyManagementServicePort" -ea 0 | select -ExpandProperty KeyManagementServicePort
        }
        return @{
            KeyManagementServiceName = $KeyManagementServiceName
            KeyManagementServicePort = $KeyManagementServicePort
        }
    }

    if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
        $hSLC = if ($global:hSLC_ -and $global:hSLC_ -ne [IntPtr]::Zero -and $global:hSLC_ -ne 0) {
            $global:hSLC_
        } else {
            Manage-SLHandle
        }
    }

    try {
        $closeHandle = $true
        if (-not $hSLC -or $hSLC -eq [IntPtr]::Zero -or $hSLC -eq 0) {
            $hr = $Global:SLC::SLOpen([ref]$hSLC)
            if ($hr -ne 0) {
                throw "SLOpen failed: HRESULT 0x{0:X8}" -f $hr
            }
        } else {
            $closeHandle = $false
        }
    }
    catch {
        return $null
    }
    
    # Normalize GUID (no braces for WMI)
    $guidNoBraces = $ActConfigId.Trim('{}')

    # Get WMI data filtered by ID
    #$wmiData = Get-WmiObject -Query "SELECT * FROM SoftwareLicensingProduct WHERE ID='$guidNoBraces'"

    $Policies = Get-LicenseDetails -ActConfigId $ActConfigId -hSLC $hSLC
    if ($Policies) {
        $policiesArray = foreach ($item in $Policies) {
            $LicenseId = $item.LicenseId
            foreach ($policy in $item.Policies) {
                if ($policy.Name -and $policy.Value) {
                    [PSCustomObject]@{
                        ID    = $LicenseId
                        Name  = $policy.Name
                        Value = $policy.Value
                    }
                }
            }
        }
    }

    # Gets the information of the specified product key.
    $SLCPKeyInfo = Get-SLCPKeyInfo -SKU $ActConfigId -loopAllValues -hSLC $hSLC

    # Define your ValidateSet values for license details
    $info = Get-LicenseDetails -ActConfigId $ActConfigId -loopAllValues -hSLC $hSLC

    if ($closeHandle) {
        Write-Warning "Consider Open handle Using Manage-SLHandle"
        Free-IntPtr -handle $hSLC -Method License
    }

    # Extract XML data filtered by ActConfigId
    $xmlData = $Global:PKeyDatabase | ? { $_.ActConfigId -eq $ActConfigId -or $_.ActConfigId -eq "{$guidNoBraces}" }
    $KeyManagementServiceInfo = Get-KeyManagementServiceInfo -SKU_ID $ActConfigId

    $ppwszInstallation = $null
    $ppwszInstallationIdPtr = [IntPtr]::Zero
    $pProductSkuId = [GUID]::new($ActConfigId)
    $null = $Global:SLC::SLGenerateOfflineInstallationIdEx(
        $hSLC, [ref]$pProductSkuId, 0, [ref]$ppwszInstallationIdPtr)
    if ($ppwszInstallationIdPtr -ne [IntPtr]::Zero) {
        $ppwszInstallation = [marshal]::PtrToStringAuto($ppwszInstallationIdPtr)
    }
    Free-IntPtr -handle $ppwszInstallationIdPtr -Method Local
    $ppwszInstallationIdPtr = 0

    $Branding = Get-BrandingValue -sku $ActConfigId

    return [PSCustomObject]@{
        # Policies
        Branding = $Branding
        Policies = $policiesArray

        # XML data properties, with safety checks
        ActConfigId        = if ($xmlData.ActConfigId) { $xmlData.ActConfigId } else { $null }
        RefGroupId         = if ($xmlData.RefGroupId) { $xmlData.RefGroupId } else { $null }
        EditionId          = if ($xmlData.EditionId) { $xmlData.EditionId } else { $null }
        #ProductDescription = if ($xmlData.ProductDescription) { $xmlData.ProductDescription } else { $null }
        ProductKeyType     = if ($xmlData.ProductKeyType) { $xmlData.ProductKeyType } else { $null }
        IsRandomized       = if ($xmlData.IsRandomized) { $xmlData.IsRandomized } else { $null }

        # License Details (from ValidateSet)
        Description          = $info["Description"]
        Name                 = $info["Name"]
        Author               = $info["Author"]
        licenseType         = $info["licenseType"]
        licenseVersion      = $info["licenseVersion"]
        licensorUrl         = $info["licensorUrl"]
        licenseCategory     = $info["licenseCategory"]
        ID                  = $info["productSkuId"]
        privateCertificateId = $info["privateCertificateId"]
        applicationId       = $info["applicationId"]
        productName         = $info["productName"]
        LicenseFamily       = $info["Family"]
        productAuthor       = $info["productAuthor"]
        productDescription  = $info["productDescription"]
        hwidootGrace        = $info["hwid:ootGrace"]
        TrustedTime         = $info["TrustedTime"]
        ProductUniquenessGroupId  = $info["ProductUniquenessGroupId"]
        issuanceCertificateId         = $info["issuanceCertificateId"]
        pkeyConfigLicenseId         = $info["pkeyConfigLicenseId"]
        ValidationURL         = $info["ValUrl"]
        BuildVersion         = $info["BuildVersion"]
        ActivationSequence         = $info["ActivationSequence"]
        EnableActivationValidation         = $info["EnableActivationValidation"]
        ValidityTimerUniqueness         = $info["ValidityTimerUniqueness"]
        ApplicationBitmap         = $info["ApplicationBitmap"]
        
        # Abbody1406 suggestion PAUrl -or EULURL
        UseLicenseURL         = if ($info["PAUrl"]) {$info["PAUrl"]} else {if ($info["EULURL"]) {$info["EULURL"]} else {$null}}
        ExtendedGrace         = $info["TimeBasedExtendedGrace"]
        phone_policy         = $info["phone:policy"]
        UXDifferentiator         = $info["UXDifferentiator"] # WindowsSkuCategory
        ProductKeyGroupUniqueness         = $info["ProductKeyGroupUniqueness"]
        migratable         = $info["migratable"]
        LicenseDependsOn         = $info["DependsOn"]
        LicenseIsAddon           = [BOOL]($info["DependsOn"])
        EnableNotificationMode         = $info["EnableNotificationMode"]
        GraceTimerUniqueness         = $info["GraceTimerUniqueness"]
        VLActivationInterval = $info["VLActivationInterval"]
        ValidationTemplateId = $info["ValidationTemplateId"]      
        licensorKeyIndex = $info["licensorKeyIndex"]
        TokenActivationILID = $info["TokenActivationILID"]
        TokenActivationILVID = $info["TokenActivationILVID"]
        TokenActivationGrantNumber = $info["TokenActivationGrantNumber"]
        TokenActivationCertificateThumbprint = $info["TokenActivationCertificateThumbprint"]
        OfflineInstallationId = $ppwszInstallation

        #KeyManagementServicePort = $info["KeyManagementServicePort"]
        #KeyManagementServiceName = if ($KeyManagementServiceInfo.KeyManagementServiceName) { $KeyManagementServiceInfo.KeyManagementServiceName } else { $null }
        #KeyManagementServicePort = if ($KeyManagementServiceInfo.KeyManagementServicePort) { $KeyManagementServiceInfo.KeyManagementServicePort } else { $null }

        # thank's abbody1406 for last 5 item's
        KeyManagementServiceProductKeyID = $info["CustomerPID"]
        KeyManagementServiceMachine = $info["KeyManagementServiceName"]
        KeyManagementServicePort = $info["KeyManagementServicePort"]
        DiscoveredKeyManagementServiceMachineName = $info["DiscoveredKeyManagementServiceName"]
        DiscoveredKeyManagementServiceMachinePort = $info["DiscoveredKeyManagementServicePort"]

        # another new list from sppwmi.dll
        ADActivationObjectDN = $info["ADActivationObjectDN"]
        ADActivationObjectName = $info["ADActivationObjectName"]
        ADActivationCsvlkPID = $info["ADActivationCsvlkPID"]
        ADActivationCsvlkSkuID = $info["ADActivationCsvlkSkuID"]
        DiscoveredKeyManagementServiceIpAddress = $info["DiscoveredKeyManagementServiceIpAddress"]
        KeyManagementServiceLookupDomain = $info["KeyManagementServiceLookupDomain"]
        TokenActivationAdditionalInfo = $info["TokenActivationAdditionalInfo"]
        TokenActivationCertThumbprint = $info["TokenActivationCertThumbprint"]
        VLActivationType = $info["VLActivationType"]
        RearmCount = $info["RearmCount"]
        RemainingRearmCount = $info["RemainingRearmCount"]

        # CPKey Info
        ProductKeyChannel    = $SLCPKeyInfo["Channel"]
        ProductKeyID        = $SLCPKeyInfo["DigitalPID"]
        ProductKeyID2       = $SLCPKeyInfo["DigitalPID2"]
        #ProductSkuId      = $SLCPKeyInfo["ProductSkuId"]
        PartialProductKey = $SLCPKeyInfo["PartialProductKey"]
    }
}

<#
Service & Active Lisence Info
~ Get-ServiceInfo >> SLGetServiceInformation
~ Get-ActiveLicenseInfo >> SLGetActiveLicenseInfo

mostly good for oem information
in case of oem license not exist,
SLGetActiveLicenseInfo will output current active license
#>
Function Query-ActiveLicenseInfo {
    $Info = @()
    $licInput, $serInput = @{}, @{}

    try {
        $ActiveLicenseInfo = Get-ActiveLicenseInfo
        @("ActivationID", "AdvancedPID", "DigitalKey",
            "EditionID", "EditionType", "EULA", "KeyType",
            "MajorVersion", "MinorVersion" ) | % { $licInput.Add($_,$ActiveLicenseInfo.$_)}

        $licInput.Keys | Sort | % {
            $Info += [PSCustomObject]@{
                Name = $_
                Value = $licInput[$_]
            }
        }

        $ServiceInfo = Get-ServiceInfo -loopAllValues
        $ServiceInfo.Keys | % { $serInput.Add($_,$ServiceInfo[$_])}

        $serInput.Keys | Sort | % {
            $Info += [PSCustomObject]@{
                Name = $_
                Value = $serInput[$_]
            }
        }
    }
    catch {
    }

    return $Info
}
# License Info Part -->

# Run part -->
function Run-Tsforge {
Write-host
Write-host
$selected = $null
$ver   = [LibTSforge.Utils]::DetectVersion()
$prod  = [LibTSforge.SPP.SPPUtils]::DetectCurrentKey()

Manage-SLHandle -Release | Out-null
$LicensingProducts = Get-SLIDList -eQueryIdType SL_ID_APPLICATION -eReturnIdType SL_ID_PRODUCT_SKU -pQueryId $windowsAppID | % {
    [PSCustomObject]@{
        ID            = $_
        Description   = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'Description'
        Name          = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'productName'
        LicenseFamily = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'Family'
    }
}

$products = $LicensingProducts | Where-Object { $_.Description -notmatch 'DEMO|MSDN|PIN|FREE|TIMEBASED|GRACE|W10' } | Select ID,Description,Name
$selected = $products | Sort-Object @{Expression='Name';Descending=$false}, @{Expression='Description';Descending=$true} | Out-GridView -Title 'Select Products to activate' -OutputMode Multiple
if (-not $selected) {
    Write-Host
    Write-Host "ERROR: No matching product found" -ForegroundColor Red
    Write-Host
    return
}

if ($selected -and @($selected).Count -ge 1) {
	foreach ($item in $selected) {
		$tsactid = $item.ID
		Write-Host "ID:          $tsactid" -ForegroundColor DarkGreen
		Write-Host "Name:        $($item.Name)" -ForegroundColor DarkGreen
		Write-Host "Description: $($item.Description)" -ForegroundColor White

        $name = $($item.Name)
        $description = $($item.Description)
        $key = GetRandomKey -ProductID $tsactid
        Write-Warning "GetRandomKey, $key"

        if (-not $key) {
            if ($description -match "Windows") {
                $windowsPath = Join-Path $env:windir "System32\spp\tokens\pkeyconfig\pkeyconfig.xrm-ms"
                $xmlData = Extract-Base64Xml -FilePath $windowsPath | Where-Object ActConfigId -Match "{$tsactid}"
            }
            elseif ($description -match "office") {
                $registryPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun"
                $officeInstallRoot = (Get-ItemProperty -Path $registryPath -ea 0).InstallPath
                if ($officeInstallRoot) {
                    $pkeyconfig = Join-Path $officeInstallRoot "\root\Licenses16\pkeyconfig-office.xrm-ms"
                    if ($pkeyconfig -and [System.IO.File]::Exists($pkeyconfig)) {
                        $xmlData = Extract-Base64Xml -FilePath $pkeyconfig | Where-Object ActConfigId -Match "{$tsactid}"
                    }
                }
            }
            if ($xmlData -and $xmlData.RefGroupId) {
                $key = Encode-Key $xmlData.RefGroupId 0 0
                Write-Warning "Encode-Key, $key"
            }
        }

		# Check if the product is VIRTUAL_MACHINE_ACTIVATION
		if ($item.Description -match 'VIRTUAL_MACHINE_ACTIVATION') {
			Write-Host "REQUIRES: Windows Server Datacenter as HOST + hyper-V or QEMU to work," -ForegroundColor Yellow
            Write-Host "by design output indicate success but slmgr.vbs -dlv indicate real state" -ForegroundColor Yellow
		}
		Write-Host

		if ($key) {
			SL-InstallProductKey -Keys @($key)
			Write-Host "Install key: $key"
		} else {
			Write-Warning "No key generated for: $($item.Name)"
			continue
		}

		Activate-License -desc $item.Description -ver $ver -prod $prod -tsactid $tsactid
	}
}
}
function Run-oHook {

if ($AutoMode) {
    Install
    return
}

Write-Host
Write-Host "Welcome to the oHook DLL Installtion Script" -ForegroundColor Cyan
Write-Host "-------------------------------------------"
Write-Host

# Prompt the user for action (I for Install, R for Remove)
$action = Read-Host "Do you want to Install (I) or Remove (R)? (Enter 'I' or 'R')"

# Normalize the input to uppercase for better consistency
$action = $action.ToUpper()

Write-Host

# Run the appropriate function based on user input
switch ($action) {
    'I' {
        Install
        break
    }
    'R' {
        Remove
        break
    }
    default {
        Write-Host "Invalid choice. Please enter either 'I' for Install or 'R' for Remove." -ForegroundColor Red
        break
    }
}

Write-Host "--------------------------------------"
Write-Host "Script execution completed."
return
}
function Run-HWID {
    param (
        [bool]$ForceVolume = $false
    )

Write-Host
Write-Host "Notice:" -ForegroundColor Magenta
Write-Host
Write-Host "HWID activation isn't supported for Evaluation or Server versions." -ForegroundColor Yellow
Write-Host "If HWID activation isn't possible, KMS38 will be used." -ForegroundColor Yellow
Write-Host "For Evaluation and Server, the script uses alternative methods:" -ForegroundColor Yellow
Write-Host
Write-Host "* KMS38   for {Server}`n* TSForge for {Evaluation}" -ForegroundColor Green
Write-Host

$sandbox = "Null" | Get-Service -ea 0
if (-not $sandbox) {
    Write-Host "'Null' service found! Possible sandbox environment." -ForegroundColor Red
    return
}

# remvoe KMS38 lock --> From MAS PROJECT, KMS38_Activation.cmd
$SID = New-Object SecurityIdentifier('S-1-5-32-544')
$Admin = ($SID.Translate([NTAccount])).Value
$ruleArgs = @("$Admin", "FullControl", "Allow")
$path = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f'
$regkey = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry64').OpenSubKey($path, 'ReadWriteSubTree', 'ChangePermissions')
if ($regkey) {
    $acl = $regkey.GetAccessControl()
    $rule = [RegistryAccessRule]::new.Invoke($ruleArgs)
    $acl.ResetAccessRule($rule)
    $regkey.SetAccessControl($acl)
}

$ClipUp = Get-Command ClipUp -ea 0
if (-not $ClipUp) {
  Write-Host
  Write-Host "ClipUp.exe is missing.!" -ForegroundColor Yellow
  Write-Host "Attemp to download ClipUp.exe from remote server" -ForegroundColor Yellow
  iwr "https://github.com/BlueOnBLack/Misc/raw/refs/heads/main/ClipUp.exe" -OutFile "$env:windir\ClipUp.exe" -ea 0
  if ([IO.FILE]::Exists("$env:windir\ClipUp.exe")) {
    if (@(Get-AuthenticodeSignature "$env:windir\ClipUp.exe" -ea 0).Status -eq 'Valid') {
      Write-Host "File was download & verified, at location: $env:windir\ClipUp.exe" -ForegroundColor Yellow
    }
    else {
      ri "$env:windir\ClipUp.exe" -Force -ea 0
    }
  }
}

$ClipUp = Get-Command ClipUp -ea 0
if (-not $ClipUp) {
  Write-Host "ClipUp.exe not found" -ForegroundColor Yellow
  return
}

$osInfo = Get-CimInstance Win32_OperatingSystem
$server = $osInfo.Caption -match "Server"
$evaluation = $osInfo.Caption -match "Evaluation"

if ($server) {
  Write-Host
  Write-Host "Server edition found" -ForegroundColor Yellow
  Write-Host "KMS38 will use instead" -ForegroundColor Yellow
}
elseif ($evaluation) {
  Write-Host
  Write-Host "evaluation edition found" -ForegroundColor Yellow
  Write-Host "use TSFORGE to Remove-Reset Evaluation Lock" -ForegroundColor Yellow
  Write-Host
  $version = [LibTSforge.Utils]::DetectVersion();
  $production = [LibTSforge.SPP.SPPUtils]::DetectCurrentKey();
  
  try {
    # Update from latest TSforge_Activation.cmd
    [LibTSforge.Modifiers.TamperedFlagsDelete]::DeleteTamperFlags($ver, $prod)
    [LibTSforge.SPP.SLApi]::RefreshLicenseStatus()
    [LibTSforge.Modifiers.RearmReset]::Reset($ver, $prod)
    [LibTSforge.Modifiers.GracePeriodReset]::Reset($version,$production)
    [LibTSforge.Modifiers.KeyChangeLockDelete]::Delete($version,$production)
  }
  catch {
  }

  Write-Host
  Write-Host "Done." -ForegroundColor Green
  return
}

# Check if the build is too old
if ($Global:osVersion.Build -lt 10240) {
    Write-Host "`n[!] Unsupported OS version detected: $buildNum" -ForegroundColor Red
    Write-Host "HWID Activation is only supported on Windows 10 or 11." -ForegroundColor DarkYellow
    Write-Host "Use the TSforge activation option from the main menu." -ForegroundColor Cyan
    return
}

("ClipSVC","wlidsvc","sppsvc","KeyIso","LicenseManager","Winmgmt") | % { Start-Service $_ -ea 0}
$EditionID = Get-ProductID
if (!$EditionID) {
  throw "EditionID Variable not found" }
$hashTable = @'
ID,KEY,SKU_ID,Key_part,value,Status,Type,Product
8b351c9c-f398-4515-9900-09df49427262,XGVPP-NMH47-7TTHJ-W3FW7-8HV2C,4,X19-99683,HGNKjkKcKQHO6n8srMUrDh/MElffBZarLqCMD9rWtgFKf3YzYOLDPEMGhuO/auNMKCeiU7ebFbQALS/MyZ7TvidMQ2dvzXeXXKzPBjfwQx549WJUU7qAQ9Txg9cR9SAT8b12Pry2iBk+nZWD9VtHK3kOnEYkvp5WTCTsrSi6Re4,0,OEM:NONSLP,Enterprise
c83cef07-6b72-4bbc-a28f-a00386872839,3V6Q6-NQXCX-V8YXR-9QCYV-QPFCT,27,X19-98746,NHn2n0N1UfVf00CfaI5LCDMDsKdVAWpD/HAfUrcTAKsw9d2Sks4h5MhyH/WUx+B6dFi8ol7D3AHorR8y9dqVS1Bd2FdZNJl/tTR1PGwYn6KL88NS19aHmFNdX8s4438vaa+Ty8Qk8EDcwm/wscC8lQmi3/RgUKYdyGFvpbGSVlk,0,Volume:MAK,EnterpriseN
4de7cb65-cdf1-4de9-8ae8-e3cce27b9f2c,VK7JG-NPHTM-C97JM-9MPGT-3V66T,48,X19-98841,Yl/jNfxJ1SnaIZCIZ4m6Pf3ySNoQXifNeqfltNaNctx+onwiivOx7qcSn8dFtURzgMzSOFnsRQzb5IrvuqHoxWWl1S3JIQn56FvKsvSx7aFXIX3+2Q98G1amPV/WEQ0uHA5d7Ya6An+g0Z0zRP7evGoomTs4YuweaWiZQjQzSpA,0,Retail,Professional
9fbaf5d6-4d83-4422-870d-fdda6e5858aa,2B87N-8KFHP-DKV6R-Y2C8J-PKCKT,49,X19-98859,Ge0mRQbW8ALk7T09V+1k1yg66qoS0lhkgPIROOIOgxKmWPAvsiLAYPKDqM4+neFCA/qf1dHFmdh0VUrwFBPYsK251UeWuElj4bZFVISL6gUt1eZwbGfv5eurQ0i+qZiFv+CcQOEFsd5DD4Up6xPLLQS3nAXODL5rSrn2sHRoCVY,0,Retail,ProfessionalN
f742e4ff-909d-4fe9-aacb-3231d24a0c58,4CPRK-NM3K3-X6XXQ-RXX86-WXCHW,98,X19-98877,vel4ytVtnE8FhvN87Cflz9sbh5QwHD1YGOeej9QP7hF3vlBR4EX2/S/09gRneeXVbQnjDOCd2KFMKRUWHLM7ZhFBk8AtlG+kvUawPZ+CIrwrD3mhi7NMv8UX/xkLK3HnBupMEuEwsMJgCUD8Pn6om1mEiQebHBAqu4cT7GN9Y0g,0,Retail,CoreN
1d1bac85-7365-4fea-949a-96978ec91ae0,N2434-X9D7W-8PF6X-8DV9T-8TYMD,99,X19-99652,Nv17eUTrr1TmUX6frlI7V69VR6yWb7alppCFJPcdjfI+xX4/Cf2np3zm7jmC+zxFb9nELUs477/ydw2KCCXFfM53bKpBQZKHE5+MdGJGxebOCcOtJ3hrkDJtwlVxTQmUgk5xnlmpk8PHg82M2uM5B7UsGLxGKK4d3hi0voSyKeI,0,Retail,CoreCountrySpecific
3ae2cc14-ab2d-41f4-972f-5e20142771dc,BT79Q-G7N6G-PGBYW-4YWX6-6F4BT,100,X19-99661,FV2Eao/R5v8sGrfQeOjQ4daokVlNOlqRCDZXuaC45bQd5PsNU3t1b4AwWeYM8TAwbHauzr4tPG0UlsUqUikCZHy0poROx35bBBMBym6Zbm9wDBVyi7nCzBtwS86eOonQ3cU6WfZxhZRze0POdR33G3QTNPrnVIM2gf6nZJYqDOA,0,Retail,CoreSingleLanguage
2b1f36bb-c1cd-4306-bf5c-a0367c2d97d8,YTMG3-N6DKC-DKB77-7M9GH-8HVX7,101,X19-98868,GH/jwFxIcdQhNxJIlFka8c1H48PF0y7TgJwaryAUzqSKXynONLw7MVciDJFVXTkCjbXSdxLSWpPIC50/xyy1rAf8aC7WuN/9cRNAvtFPC1IVAJaMeq1vf4mCqRrrxJQP6ZEcuAeHFzLe/LLovGWCd8rrs6BbBwJXCvAqXImvycQ,0,Retail,Core
2a6137f3-75c0-4f26-8e3e-d83d802865a4,XKCNC-J26Q9-KFHD2-FKTHY-KD72Y,119,X19-99606,hci78IRWDLBtdbnAIKLDgV9whYgtHc1uYyp9y6FszE9wZBD5Nc8CUD2pI2s2RRd3M04C4O7M3tisB3Ov/XVjpAbxlX3MWfUR5w4MH0AphbuQX0p5MuHEDYyfqlRgBBRzOKePF06qfYvPQMuEfDpKCKFwNojQxBV8O0Arf5zmrIw,0,OEM:NONSLP,PPIPro
e558417a-5123-4f6f-91e7-385c1c7ca9d4,YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY,121,X19-98886,x9tPFDZmjZMf29zFeHV5SHbXj8Wd8YAcCn/0hbpLcId4D7OWqkQKXxXHIegRlwcWjtII0sZ6WYB0HQV2KH3LvYRnWKpJ5SxeOgdzBIJ6fhegYGGyiXsBv9sEb3/zidPU6ZK9LugVGAcRZ6HQOiXyOw+Yf5H35iM+2oDZXSpjvJw,0,Retail,Education
c5198a66-e435-4432-89cf-ec777c9d0352,84NGF-MHBT6-FXBX8-QWJK7-DRR8H,122,X19-98892,jkL4YZkmBCJtvL1fT30ZPBcjmzshBSxjwrE0Q00AZ1hYnhrH+npzo1MPCT6ZRHw19ZLTz7wzyBb0qqcBVbtEjZW0Xs2MYLxgriyoONkhnPE6KSUJBw7C0enFVLHEqnVu/nkaOFfockN3bc+Eouw6W2lmHjklPHc9c6Clo04jul0,0,Retail,EducationN
f6e29426-a256-4316-88bf-cc5b0f95ec0c,PJB47-8PN2T-MCGDY-JTY3D-CBCPV,125,X23-50331,OPGhsyx+Ctw7w/KLMRNrY+fNBmKPjUG0R9RqkWk4e8ez+ExSJxSLLex5WhO5QSNgXLmEra+cCsN6C638aLjIdH2/L7D+8z/C6EDgRvbHMmidHg1lX3/O8lv0JudHkGtHJYewjorn/xXGY++vOCTQdZNk6qzEgmYSvPehKfdg8js,1,Volume:MAK,EnterpriseS,Ge
cce9d2de-98ee-4ce2-8113-222620c64a27,KCNVH-YKWX8-GJJB9-H9FDT-6F7W2,125,X22-66075,GCqWmJOsTVun9z4QkE9n2XqBvt3ZWSPl9QmIh9Q2mXMG/QVt2IE7S+ES/NWlyTSNjLVySr1D2sGjxgEzy9kLwn7VENQVJ736h1iOdMj/3rdqLMSpTa813+nPSQgKpqJ3uMuvIvRP0FdB7Y4qt8qf9kNKK25A1QknioD/6YubL/4,1,Volume:MAK,EnterpriseS,VB
d06934ee-5448-4fd1-964a-cd077618aa06,43TBQ-NH92J-XKTM7-KT3KK-P39PB,125,X21-83233,EpB6qOCo8pRgO5kL4vxEHck2J1vxyd9OqvxUenDnYO9AkcGWat/D74ZcFg5SFlIya1U8l5zv+tsvZ4wAvQ1IaFW1PwOKJLOaGgejqZ41TIMdFGGw+G+s1RHsEnrWr3UOakTodby1aIMUMoqf3NdaM5aWFo8fOmqWC5/LnCoighs,0,OEM:NONSLP,EnterpriseS,RS5
706e0cfd-23f4-43bb-a9af-1a492b9f1302,NK96Y-D9CD8-W44CQ-R8YTK-DYJWX,125,X21-05035,ntcKmazIvLpZOryft28gWBHu1nHSbR+Gp143f/BiVe+BD2UjHBZfSR1q405xmQZsygz6VRK6+zm8FPR++71pkmArgCLhodCQJ5I4m7rAJNw/YX99pILphi1yCRcvHsOTGa825GUVXgf530tHT6hr0HQ1lGeGgG1hPekpqqBbTlg,0,OEM:NONSLP,EnterpriseS,RS1
faa57748-75c8-40a2-b851-71ce92aa8b45,FWN7H-PF93Q-4GGP8-M8RF3-MDWWW,125,X19-99617,Fe9CDClilrAmwwT7Yhfx67GafWRQEpwyj8R+a4eaTqbpPcAt7d1hv1rx8Sa9AzopEGxIrb7IhiPoDZs0XaT1HN0/olJJ/MnD73CfBP4sdQdLTsSJE3dKMWYTQHpnjqRaS/pNBYRr8l9Mv8yfcP8uS2MjIQ1cRTqRmC7WMpShyCg,0,OEM:NONSLP,EnterpriseS,TH
3d1022d8-969f-4222-b54b-327f5a5af4c9,2DBW3-N2PJG-MVHW3-G7TDK-9HKR4,126,X21-04921,zLPNvcl1iqOefy0VLg+WZgNtRNhuGpn8+BFKjMqjaNOSKiuDcR6GNDS5FF1Aqk6/e6shJ+ohKzuwrnmYq3iNQ3I2MBlYjM5kuNfKs8Vl9dCjSpQr//GBGps6HtF2xrG/2g/yhtYC7FbtGDIE16uOeNKFcVg+XMb0qHE/5Etyfd8,0,Volume:MAK,EnterpriseSN,RS1
60c243e1-f90b-4a1b-ba89-387294948fb6,NTX6B-BRYC2-K6786-F6MVQ-M7V2X,126,X19-98770,kbXfe0z9Vi1S0yfxMWzI5+UtWsJKzxs7wLGUDLjrckFDn1bDQb4MvvuCK1w+Qrq33lemiGpNDspa+ehXiYEeSPFcCvUBpoMlGBFfzurNCHWiv3o1k3jBoawJr/VoDoVZfxhkps0fVoubf9oy6C6AgrkZ7PjCaS58edMcaUWvYYg,0,Volume:MAK,EnterpriseSN,TH
01eb852c-424d-4060-94b8-c10d799d7364,3XP6D-CRND4-DRYM2-GM84D-4GG8Y,139,X23-37869,PVW0XnRJnsWYjTqxb6StCi2tge/uUwegjdiFaFUiZpwdJ620RK+MIAsSq5S+egXXzIWNntoy2fB6BO8F1wBFmxP/mm/3rn5C33jtF5QrbNqY7X9HMbqSiC7zhs4v4u2Xa4oZQx8JQkwr8Q2c/NgHrOJKKRASsSckhunxZ+WVEuM,1,Retail,ProfessionalCountrySpecific,Zn
eb6d346f-1c60-4643-b960-40ec31596c45,DXG7C-N36C4-C4HTG-X4T3X-2YV77,161,X21-43626,MaVqTkRrGnOqYizl15whCOKWzx01+BZTVAalvEuHXM+WV55jnIfhWmd/u1GqCd5OplqXdU959zmipK2Iwgu2nw/g91nW//sQiN/cUcvg1Lxo6pC3gAo1AjTpHmGIIf9XlZMYlD+Vl6gXsi/Auwh3yrSSFh5s7gOczZoDTqQwHXA,0,Retail,ProfessionalWorkstation
89e87510-ba92-45f6-8329-3afa905e3e83,WYPNQ-8C467-V2W6J-TX4WX-WT2RQ,162,X21-43644,JVGQowLiCcPtGY9ndbBDV+rTu/q5ljmQTwQWZgBIQsrAeQjLD8jLEk/qse7riZ7tMT6PKFVNXeWqF7PhLAmACbE8O3Lvp65XMd/Oml9Daynj5/4n7unsffFHIHH8TGyO5j7xb4dkFNqC5TX3P8/1gQEkTIdZEOTQQXFu0L2SP5c,0,Retail,ProfessionalWorkstationN
62f0c100-9c53-4e02-b886-a3528ddfe7f6,8PTT6-RNW4C-6V7J2-C2D3X-MHBPB,164,X21-04955,CEDgxI8f/fxMBiwmeXw5Of55DG32sbGALzHihXkdbYTDaE3pY37oAA4zwGHALzAFN/t254QImGPYR6hATgl+Cp804f7serJqiLeXY965Zy67I4CKIMBm49lzHLFJeDnVTjDB0wVyN29pvgO3+HLhZ22KYCpkRHFFMy2OKxS68Yc,0,Retail,ProfessionalEducation
13a38698-4a49-4b9e-8e83-98fe51110953,GJTYN-HDMQY-FRR76-HVGC7-QPF8P,165,X21-04956,r35zp9OfxKSBcTxKWon3zFtbOiCufAPo6xRGY5DJqCRFKdB0jgZalNQitvjmaZ/Rlez2vjRJnEart4LrvyW4d9rrukAjR3+c3UkeTKwoD3qBl9AdRJbXCa2BdsoXJs1WVS4w4LuVzpB/SZDuggZt0F2DlMB427F5aflook/n1pY,0,Retail,ProfessionalEducationN
df96023b-dcd9-4be2-afa0-c6c871159ebe,NJCF7-PW8QT-3324D-688JX-2YV66,175,X21-41295,rVpetYUmiRB48YJfCvJHiaZapJ0bO8gQDRoql+rq5IobiSRu//efV1VXqVpBkwILQRKgKIVONSTUF5y2TSxlDLbDSPKp7UHfbz17g6vRKLwOameYEz0ZcK3NTbApN/cMljHvvF/mBag1+sHjWu+eoFzk8H89k9nw8LMeVOPJRDc,0,Retail,ServerRdsh
d4ef7282-3d2c-4cf0-9976-8854e64a8d1e,V3WVW-N2PV2-CGWC3-34QGF-VMJ2C,178,X21-32983,Xzme9hDZR6H0Yx0deURVdE6LiTOkVqWng5W/OTbkxRc0rq+mSYpo/f/yqhtwYlrkBPWx16Yok5Bvcb34vbKHvEAtxfYp4te20uexLzVOtBcoeEozARv4W/6MhYfl+llZtR5efsktj4N4/G4sVbuGvZ9nzNfQO9TwV6NGgGEj2Ec,0,Retail,Cloud
af5c9381-9240-417d-8d35-eb40cd03e484,NH9J3-68WK7-6FB93-4K3DF-DJ4F6,179,X21-32987,QGRDZOU/VZhYLOSdp2xDnFs8HInNZctcQlWCIrORVnxTQr55IJwN4vK3PJHjkfRLQ/bgUrcEIhyFbANqZFUq8yD1YNubb2bjNORgI/m8u85O9V7nDGtxzO/viEBSWyEHnrzLKKWYqkRQKbbSW3ungaZR0Ti5O2mAUI4HzAFej50,0,Retail,CloudN
8ab9bdd1-1f67-4997-82d9-8878520837d9,XQQYW-NFFMW-XJPBH-K8732-CKFFD,188,X21-99378,djy0od0uuKd2rrIl+V1/2+MeRltNgW7FEeTNQsPMkVSL75NBphgoso4uS0JPv2D7Y1iEEvmVq6G842Kyt52QOwXgFWmP/IQ6Sq1dr+fHK/4Et7bEPrrGBEZoCfWqk0kdcZRPBij2KN6qCRWhrk1hX2g+U40smx/EYCLGh9HCi24,0,OEM:DM,IoTEnterprise
ed655016-a9e8-4434-95d9-4345352c2552,QPM6N-7J2WJ-P88HH-P3YRH-YY74H,191,X21-99682,qHs/PzfhYWdtSys2edzcz4h+Qs8aDqb8BIiQ/mJ/+0uyoJh1fitbRCIgiFh2WAGZXjdgB8hZeheNwHibd8ChXaXg4u+0XlOdFlaDTgTXblji8fjETzDBk9aGkeMCvyVXRuUYhTSdp83IqGHz7XuLwN2p/6AUArx9JZCoLGV8j3w,0,OEM:NONSLP,IoTEnterpriseS,VB
6c4de1b8-24bb-4c17-9a77-7b939414c298,CGK42-GYN6Y-VD22B-BX98W-J8JXD,191,X23-12617,J/fpIRynsVQXbp4qZNKp6RvOgZ/P2klILUKQguMlcwrBZybwNkHg/kM5LNOF/aDzEktbPnLnX40GEvKkYT6/qP4cMhn/SOY0/hYOkIdR34ilzNlVNq5xP7CMjCjaUYJe+6ydHPK6FpOuEoWOYYP5BZENKNGyBy4w4shkMAw19mA,0,OEM:NONSLP,IoTEnterpriseS,Ge
d4bdc678-0a4b-4a32-a5b3-aaa24c3b0f24,K9VKN-3BGWV-Y624W-MCRMQ-BHDCD,202,X22-53884,kyoNx2s93U6OUSklB1xn+GXcwCJO1QTEtACYnChi8aXSoxGQ6H2xHfUdHVCwUA1OR0UeNcRrMmOzZBOEUBtdoGWSYPg9AMjvxlxq9JOzYAH+G6lT0UbCWgMSGGrqdcIfmshyEak3aUmsZK6l+uIAFCCZZ/HbbCRkkHC5rWKstMI,0,Retail,CloudEditionN
92fb8726-92a8-4ffc-94ce-f82e07444653,KY7PN-VR6RX-83W6Y-6DDYQ-T6R4W,203,X22-53847,gD6HnT4jP4rcNu9u83gvDiQq1xs7QSujcDbo60Di5iSVa9/ihZ7nlhnA0eDEZfnoDXriRiPPqc09T6AhSnFxLYitAkOuPJqL5UMobIrab9dwTKlowqFolxoHhLOO4V92Hsvn/9JLy7rEzoiAWHhX/0cpMr3FCzVYPeUW1OyLT1A,0,Retail,CloudEdition
5a85300a-bfce-474f-ac07-a30983e3fb90,N979K-XWD77-YW3GB-HBGH6-D32MH,205,X23-15042,blZopkUuayCTgZKH4bOFiisH9GTAHG5/js6UX/qcMWWc3sWNxKSX1OLp1k3h8Xx1cFuvfG/fNAw/I83ssEtPY+A0Gx1JF4QpRqsGOqJ5ruQ2tGW56CJcCVHkB+i46nJAD759gYmy3pEYMQbmpWbhLx3MJ6kvwxKfU+0VCio8k50,0,OEM:DM,IoTEnterpriseSK
80083eae-7031-4394-9e88-4901973d56fe,P8Q7T-WNK7X-PMFXY-VXHBG-RRK69,206,X23-62084,habUJ0hhAG0P8iIKaRQ74/wZQHyAdFlwHmrejNjOSRG08JeqilJlTM6V8G9UERLJ92/uMDVHIVOPXfN8Zdh8JuYO8oflPnqymIRmff/pU+Gpb871jV2JDA4Cft5gmn+ictKoN4VoSfEZRR+R5hzF2FsoCExDNNw6gLdjtiX94uA,0,OEM:DM,IoTEnterpriseK
'@
$customObjectArray = $hashTable | ConvertFrom-Csv

Manage-SLHandle -Release | Out-null
$LicensingProducts = Get-SLIDList -eQueryIdType SL_ID_PRODUCT_SKU -eReturnIdType SL_ID_PRODUCT_SKU | % {
    try {
        $Branding = $null
        [XML]$licenseData = Get-LicenseDetails $_ -ReturnRawData $true
        $Branding = ($licenseData.licenseGroup.license[1].otherInfo.infoTables.infoList.infoStr | ? Name -EQ win:branding).'#text'
    }
    catch {
        $Branding = $null
    }
    [PSCustomObject]@{
        ID            = $_
        Description   = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'Description'
        Name          = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'productName'
        LicenseFamily = Get-LicenseDetails -ActConfigId $_ -pwszValueName 'Family'
        Branding      = $Branding
    }
}
$SupportedProducts = $LicensingProducts | 
    ? { $customObjectArray.ID -contains $_.ID } | ? {
        ($customObjectArray | ? ID -EQ $_.ID | select -ExpandProperty SKU_ID ) -match $_.Branding }

if ($ForceVolume -eq $true) {
   $SupportedProducts = $null
}

if ($server -or !$SupportedProducts) {

  Write-Host
  Write-Host "ERROR: No matching product found" -ForegroundColor Red
  Write-Host "Trying to use KMS38 instead." -ForegroundColor Red

  try {
    $product = $LicensingProducts | ? Description -Match 'VOLUME_KMSCLIENT' | ? LicenseFamily -EQ $EditionID
    if (-not $product) {
      $product = $LicensingProducts | ? Description -Match 'VOLUME_KMSCLIENT' | ? { $_.LicenseFamily } | Out-GridView -Title "Select prefered product's" -OutputMode Single
    }
    if ($product){
        $Vol_Key = GetRandomKey -ProductID $product.ID
        if (-not $Vol_Key) {
            $refSku = Retrieve-ProductKeyInfo -SkuId $product.ID
            $Vol_Key = Encode-Key $refSku 0 0
            Write-Warning "Encode-Key, $Vol_Key"
        }
    }
    else {
      Write-Host
      Write-Host "ERROR: No matching product found" -ForegroundColor Red
    }
  } catch {
    Write-Host "ERROR: fetch product - Key for VOLUME_KMSCLIENT version" -ForegroundColor Red
    return
  }
  if ([STRING]::IsNullOrWhiteSpace($Vol_Key) -or [STRING]::IsNullOrEmpty(($Vol_Key))) {
    return
  }
}
else {
  $filter = ($customObjectArray | ? Status -EQ 0 | select ID).ID
  $products =  $SupportedProducts | ? {$filter -contains $_.ID}

  $product = $null
  $product = $products | ? {$_.LicenseFamily -match $EditionID} | select -First 1
  if (-not $product) {
    $product = $products | Out-GridView -Title "Select prefered product's" -OutputMode Single
  }

  if (-not $product) {
    return  }
}

Function Encode-Blob {
    param (
        $SessionIdStr
    )
    function Sign {
        param (
            $Properties,
            $rsa
        )

        $sha256 = [Security.Cryptography.SHA256]::Create()
        $bytes = [Text.Encoding]::UTF8.GetBytes($Properties)
        $hash = $sha256.ComputeHash($bytes)

        $signature = $rsa.SignHash($hash, [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1)
        return [Convert]::ToBase64String($signature)
    }
    [byte[]] $key = 0x07,0x02,0x00,0x00,0x00,0xA4,0x00,0x00,0x52,0x53,0x41,0x32,0x00,0x04,0x00,0x00,
                    0x01,0x00,0x01,0x00,0x29,0x87,0xBA,0x3F,0x52,0x90,0x57,0xD8,0x12,0x26,0x6B,0x38,
                    0xB2,0x3B,0xF9,0x67,0x08,0x4F,0xDD,0x8B,0xF5,0xE3,0x11,0xB8,0x61,0x3A,0x33,0x42,
                    0x51,0x65,0x05,0x86,0x1E,0x00,0x41,0xDE,0xC5,0xDD,0x44,0x60,0x56,0x3D,0x14,0x39,
                    0xB7,0x43,0x65,0xE9,0xF7,0x2B,0xA5,0xF0,0xA3,0x65,0x68,0xE9,0xE4,0x8B,0x5C,0x03,
                    0x2D,0x36,0xFE,0x28,0x4C,0xD1,0x3C,0x3D,0xC1,0x90,0x75,0xF9,0x6E,0x02,0xE0,0x58,
                    0x97,0x6A,0xCA,0x80,0x02,0x42,0x3F,0x6C,0x15,0x85,0x4D,0x83,0x23,0x6A,0x95,0x9E,
                    0x38,0x52,0x59,0x38,0x6A,0x99,0xF0,0xB5,0xCD,0x53,0x7E,0x08,0x7C,0xB5,0x51,0xD3,
                    0x8F,0xA3,0x0D,0xA0,0xFA,0x8D,0x87,0x3C,0xFC,0x59,0x21,0xD8,0x2E,0xD9,0x97,0x8B,
                    0x40,0x60,0xB1,0xD7,0x2B,0x0A,0x6E,0x60,0xB5,0x50,0xCC,0x3C,0xB1,0x57,0xE4,0xB7,
                    0xDC,0x5A,0x4D,0xE1,0x5C,0xE0,0x94,0x4C,0x5E,0x28,0xFF,0xFA,0x80,0x6A,0x13,0x53,
                    0x52,0xDB,0xF3,0x04,0x92,0x43,0x38,0xB9,0x1B,0xD9,0x85,0x54,0x7B,0x14,0xC7,0x89,
                    0x16,0x8A,0x4B,0x82,0xA1,0x08,0x02,0x99,0x23,0x48,0xDD,0x75,0x9C,0xC8,0xC1,0xCE,
                    0xB0,0xD7,0x1B,0xD8,0xFB,0x2D,0xA7,0x2E,0x47,0xA7,0x18,0x4B,0xF6,0x29,0x69,0x44,
                    0x30,0x33,0xBA,0xA7,0x1F,0xCE,0x96,0x9E,0x40,0xE1,0x43,0xF0,0xE0,0x0D,0x0A,0x32,
                    0xB4,0xEE,0xA1,0xC3,0x5E,0x9B,0xC7,0x7F,0xF5,0x9D,0xD8,0xF2,0x0F,0xD9,0x8F,0xAD,
                    0x75,0x0A,0x00,0xD5,0x25,0x43,0xF7,0xAE,0x51,0x7F,0xB7,0xDE,0xB7,0xAD,0xFB,0xCE,
                    0x83,0xE1,0x81,0xFF,0xDD,0xA2,0x77,0xFE,0xEB,0x27,0x1F,0x10,0xFA,0x82,0x37,0xF4,
                    0x7E,0xCC,0xE2,0xA1,0x58,0xC8,0xAF,0x1D,0x1A,0x81,0x31,0x6E,0xF4,0x8B,0x63,0x34,
                    0xF3,0x05,0x0F,0xE1,0xCC,0x15,0xDC,0xA4,0x28,0x7A,0x9E,0xEB,0x62,0xD8,0xD8,0x8C,
                    0x85,0xD7,0x07,0x87,0x90,0x2F,0xF7,0x1C,0x56,0x85,0x2F,0xEF,0x32,0x37,0x07,0xAB,
                    0xB0,0xE6,0xB5,0x02,0x19,0x35,0xAF,0xDB,0xD4,0xA2,0x9C,0x36,0x80,0xC6,0xDC,0x82,
                    0x08,0xE0,0xC0,0x5F,0x3C,0x59,0xAA,0x4E,0x26,0x03,0x29,0xB3,0x62,0x58,0x41,0x59,
                    0x3A,0x37,0x43,0x35,0xE3,0x9F,0x34,0xE2,0xA1,0x04,0x97,0x12,0x9D,0x8C,0xAD,0xF7,
                    0xFB,0x8C,0xA1,0xA2,0xE9,0xE4,0xEF,0xD9,0xC5,0xE5,0xDF,0x0E,0xBF,0x4A,0xE0,0x7A,
                    0x1E,0x10,0x50,0x58,0x63,0x51,0xE1,0xD4,0xFE,0x57,0xB0,0x9E,0xD7,0xDA,0x8C,0xED,
                    0x7D,0x82,0xAC,0x2F,0x25,0x58,0x0A,0x58,0xE6,0xA4,0xF4,0x57,0x4B,0xA4,0x1B,0x65,
                    0xB9,0x4A,0x87,0x46,0xEB,0x8C,0x0F,0x9A,0x48,0x90,0xF9,0x9F,0x76,0x69,0x03,0x72,
                    0x77,0xEC,0xC1,0x42,0x4C,0x87,0xDB,0x0B,0x3C,0xD4,0x74,0xEF,0xE5,0x34,0xE0,0x32,
                    0x45,0xB0,0xF8,0xAB,0xD5,0x26,0x21,0xD7,0xD2,0x98,0x54,0x8F,0x64,0x88,0x20,0x2B,
                    0x14,0xE3,0x82,0xD5,0x2A,0x4B,0x8F,0x4E,0x35,0x20,0x82,0x7E,0x1B,0xFE,0xFA,0x2C,
                    0x79,0x6C,0x6E,0x66,0x94,0xBB,0x0A,0xEB,0xBA,0xD9,0x70,0x61,0xE9,0x47,0xB5,0x82,
                    0xFC,0x18,0x3C,0x66,0x3A,0x09,0x2E,0x1F,0x61,0x74,0xCA,0xCB,0xF6,0x7A,0x52,0x37,
                    0x1D,0xAC,0x8D,0x63,0x69,0x84,0x8E,0xC7,0x70,0x59,0xDD,0x2D,0x91,0x1E,0xF7,0xB1,
                    0x56,0xED,0x7A,0x06,0x9D,0x5B,0x33,0x15,0xDD,0x31,0xD0,0xE6,0x16,0x07,0x9B,0xA5,
                    0x94,0x06,0x7D,0xC1,0xE9,0xD6,0xC8,0xAF,0xB4,0x1E,0x2D,0x88,0x06,0xA7,0x63,0xB8,
                    0xCF,0xC8,0xA2,0x6E,0x84,0xB3,0x8D,0xE5,0x47,0xE6,0x13,0x63,0x8E,0xD1,0x7F,0xD4,
                    0x81,0x44,0x38,0xBF

    $rsa = New-Object Security.Cryptography.RSACryptoServiceProvider
    $rsa.ImportCspBlob($key)
    $SessionId = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($SessionIdStr + [char]0))
    $PropertiesStr = "OA3xOriginalProductId=;OA3xOriginalProductKey=;SessionId=$SessionId;TimeStampClient=2022-10-11T12:00:00Z"
    $SignatureStr = Sign $PropertiesStr $rsa
    return @"
<?xml version="1.0" encoding="utf-8"?><genuineAuthorization xmlns="http://www.microsoft.com/DRM/SL/GenuineAuthorization/1.0"><version>1.0</version><genuineProperties origin="sppclient"><properties>$PropertiesStr</properties><signatures><signature name="clientLockboxKey" method="rsa-sha256">$SignatureStr</signature></signatures></genuineProperties></genuineAuthorization>
"@
}

$outputPath = Join-Path "C:\ProgramData\Microsoft\Windows\ClipSVC\GenuineTicket" "GenuineTicket.xml"
if ($Vol_Key) {
    $SessionID = 'OSMajorVersion=5;OSMinorVersion=1;OSPlatformId=2;PP=0;GVLKExp=2038-01-19T03:14:07Z;DownlevelGenuineState=1;'
    $signature = Encode-Blob -SessionIdStr $SessionID
    #$signature = '<?xml version="1.0" encoding="utf-8"?><genuineAuthorization xmlns="http://www.microsoft.com/DRM/SL/GenuineAuthorization/1.0"><version>1.0</version><genuineProperties origin="sppclient"><properties>OA3xOriginalProductId=;OA3xOriginalProductKey=;SessionId=TwBTAE0AYQBqAG8AcgBWAGUAcgBzAGkAbwBuAD0ANQA7AE8AUwBNAGkAbgBvAHIAVgBlAHIAcwBpAG8AbgA9ADEAOwBPAFMAUABsAGEAdABmAG8AcgBtAEkAZAA9ADIAOwBQAFAAPQAwADsARwBWAEwASwBFAHgAcAA9ADIAMAAzADgALQAwADEALQAxADkAVAAwADMAOgAxADQAOgAwADcAWgA7AEQAbwB3AG4AbABlAHYAZQBsAEcAZQBuAHUAaQBuAGUAUwB0AGEAdABlAD0AMQA7AAAA;TimeStampClient=2022-10-11T12:00:00Z</properties><signatures><signature name="clientLockboxKey" method="rsa-sha256">C52iGEoH+1VqzI6kEAqOhUyrWuEObnivzaVjyef8WqItVYd/xGDTZZ3bkxAI9hTpobPFNJyJx6a3uriXq3HVd7mlXfSUK9ydeoUdG4eqMeLwkxeb6jQWJzLOz41rFVSMtBL0e+ycCATebTaXS4uvFYaDHDdPw2lKY8ADj3MLgsA=</signature></signatures></genuineProperties></genuineAuthorization>'
}

$cProduct = $customObjectArray | ? ID -EQ $product.ID
if (-not $Vol_Key) {
    $SessionID = 'OSMajorVersion=5;OSMinorVersion=1;OSPlatformId=2;PP=0;Pfn=Microsoft.Windows.'+$($cProduct.SKU_ID)+'.'+$($cProduct.Key_part)+
        '_8wekyb3d8bbwe;PKeyIID=465145217131314304264339481117862266242033457260311819664735280;'
    $signature = Encode-Blob -SessionIdStr $SessionID

    <#
    $SessionID += [char]0
    $encoded = [convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($SessionID))

    if ($encoded -notmatch 'AAAA$') {
       Write-Warning "Base64 string doesn't contain 'AAAA'"
    }

    $signature = '<?xml version="1.0" encoding="utf-8"?><genuineAuthorization xmlns="http://www.microsoft.com/DRM/SL/GenuineAuthorization/1.0">'+
      '<version>1.0</version><genuineProperties origin="sppclient"><properties>OA3xOriginalProductId=;OA3xOriginalProductKey=;SessionId=' +
      $encoded + ';TimeStampClient=2022-10-11T12:00:00Z</properties><signatures><signature name="clientLockboxKey" method="rsa-sha256">' +
      $cProduct.value + '=</signature></signatures></genuineProperties></genuineAuthorization>'
    #>

    $geoName = (Get-ItemProperty -Path "HKCU:\Control Panel\International\Geo").Name
    $geoNation = (Get-ItemProperty -Path "HKCU:\Control Panel\International\Geo").Nation
}

$tdir = "$env:ProgramData\Microsoft\Windows\ClipSVC\GenuineTicket"

# Create directory if it doesn't exist
if (-not (Test-Path -Path $tdir)) {
    New-Item -ItemType Directory -Path $tdir | Out-Null
}

# Delete files starting with "Genuine" in $tdir
Get-ChildItem -Path $tdir -Filter "Genuine*" -File -ea 0 | Remove-Item -Force -ea 0

# Delete .xml files in $tdir
Get-ChildItem -Path $tdir -Filter "*.xml" -File -ea 0 | Remove-Item -Force -ea 0

# Delete all files in the Migration folder
$migrationPath = "$env:ProgramData\Microsoft\Windows\ClipSVC\Install\Migration"
if (Test-Path -Path $migrationPath) {
    Get-ChildItem -Path $migrationPath -File -ea 0 | Remove-Item -Force -ea 0
}

if ($Vol_Key) {
    # Remove registry keys
    Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f" -Force -Recurse -ea 0
    Remove-Item -Path "HKU:\S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f" -Force -Recurse -ea 0

    # Registry path for new entries
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f\$($product.ID)"

    # Create new registry values
    New-Item -Path $regPath -Force -ea 0 | Out-Null
    New-ItemProperty -Path $regPath -Name "KeyManagementServiceName" -PropertyType String -Value "127.0.0.2" -Force -ea 0
    New-ItemProperty -Path $regPath -Name "KeyManagementServicePort" -PropertyType String -Value "1688" -Force -ea 0
}

try {
    if ($geoName -and $geoNation -and ($geoName -ne 'US')){
      Set-WinHomeLocation -GeoId 244 -ea 0 }
    if ($Vol_Key) {
        $tsactid = $product.ID
        Manage-SLHandle -Release | Out-null
        Write-Warning "SL-InstallProductKey -Keys $Vol_Key"
        SL-InstallProductKey -Keys @($Vol_Key)
    }
    else {
        $tsactid = $cProduct.ID

        Manage-SLHandle -Release | Out-null
        $HWID_KEY = $cProduct.Key
        Write-Warning "SL-InstallProductKey -Keys $HWID_KEY"
        SL-InstallProductKey -Keys @($HWID_KEY)
    }
    
    $ID_PKEY = Retrieve-SKUInfo -SkuId $tsactid -eReturnIdType SL_ID_PKEY
    if ($ID_PKEY -eq $null) {
        $RefGroupId = $Global:PKeyDatabase | ? ActConfigId -Match "{$tsactid}" | select -ExpandProperty RefGroupId
        if (-not $RefGroupId) {
           Write-Warning "Fail to receive RefGroupId for $tsactid"
		   if ($HWID_KEY) {
			   Clear-host
			   Write-host
			   Run-HWID -ForceVolume $true
			   return
		   }
        }
        if ($RefGroupId) {
            $key = Encode-Key $RefGroupId
            if ($key) {
                $null = SL-InstallProductKey -Keys $key
                $ID_PKEY = Retrieve-SKUInfo -SkuId $tsactid -eReturnIdType SL_ID_PKEY
                if ($ID_PKEY -eq $null) {
                    Write-Warning "Fail to install key for $tsactid"
                    return
                }}}}

    [System.IO.File]::WriteAllText($outputPath, $signature, [Encoding]::UTF8)
    Write-Host
    clipup -v -o
    [System.IO.File]::WriteAllText($outputPath, $signature, [Encoding]::UTF8)
    Write-Host
    if ($Vol_Key) {
      Stop-Service sppsvc -force -ea 0
    }
    Restart-Service ClipSVC
    Write-Host

    if ($Vol_Key) {
        Manage-SLHandle -Release | Out-null
        $null = SL-ReArm -AppID 55c92734-d682-4d71-983e-d6ec3f16059f -skuID $product.ID
    }
    else {
        Manage-SLHandle -Release | Out-null
        $null = SL-Activate -skuID $product.ID
    }
   
    Manage-SLHandle -Release | Out-null
    $null = SL-RefreshLicenseStatus -AppID 55c92734-d682-4d71-983e-d6ec3f16059f -skuID $product.ID
}
catch {
    Write-Host "ERROR: Failed to activate. Operation aborted." -ForegroundColor Red
    Write-Host
    Write-Host $_.Exception.Message
    Write-Host
    return
}
Finally {
  if ($geoNation) {
    Set-WinHomeLocation -GeoId $geoNation -ea 0 }
}

Manage-SLHandle -Release | Out-null
$StatusInfo = Get-SLLicensingStatus -ApplicationID 55c92734-d682-4d71-983e-d6ec3f16059f -SkuID $product.ID

if (-not $StatusInfo) {
    Write-Warning "Fail to fetch status data"
    return
}
if ($Vol_Key -and (
    $StatusInfo.LicenseTier -ne [LicenseCategory]::KMS38)) {
        Write-Host
        Write-Host "KMS38 Activation Failed." -ForegroundColor Red
        Write-Host "Try re-apply Activation again later" -ForegroundColor Red
        Write-Host
        Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f" -Force -Recurse -ea 0
        Remove-Item -Path "HKU:\S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f" -Force -Recurse -ea 0
}
elseif (-not $Vol_Key -and (
    $StatusInfo.LicenseStatus -ne [LicenseStatusEnum]::Licensed)) {
    Write-Host
    Write-Host "HWID Activation Failed." -ForegroundColor Red
    Write-Host "Try re-apply Activation again later" -ForegroundColor Red
    Write-Host
}
else {
    Write-Host
    Write-Host "everything is Well Done" -ForegroundColor Yellow
    Write-Host "Go Home & Rest. !" -ForegroundColor Yellow
    Write-Host

    if ($Vol_Key) {

        # enable KMS38 lock --> From MAS PROJECT, KMS38_Activation.cmd
        $SID = New-Object SecurityIdentifier('S-1-5-32-544')
        $Admin = ($SID.Translate([NTAccount])).Value
        $ruleArgs = @("$Admin", "Delete, SetValue", "ContainerInherit", "None", "Deny")
        $path = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f'
        $key = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry64').OpenSubKey($path, 'ReadWriteSubTree', 'ChangePermissions')
        if ($key) {
            $acl = $key.GetAccessControl()
            $rule = [RegistryAccessRule]::new.Invoke($ruleArgs)
            $acl.ResetAccessRule($rule)
            $key.SetAccessControl($acl)
        }
    }
}

}
function Run-KMS {
    Set-DefinedEntities
    Clean-RegistryKeys
    Service-Check

    # Windows_Addict Not genuine fix
    $regPaths = @($Global:XSPP_USER, $Global:XSPP_HKLM_X32, $Global:XSPP_HKLM_X64)
    foreach ($path in $regPaths) {
        try {
            New-Item -Path $path -Force -ea 0 | Out-Null
            if (Test-Path $path) {
                Set-ItemProperty -Path $path -Name 'KeyManagementServiceName' -Value $Global:IP_ADDRESS -Type String -Force -ea 0
            }
        } catch {
            #Write-Host "Failed to write to $path" -ForegroundColor Red
        }
    }

    write-host
    write-host "Convert & Activate Smart Solution For Office / Windows Products"
    write-host "Support Activation for 			 :: Office 2010 --> late Office 2021"
    write-host "Support Convert    For 			 :: Office 2016 [MSI], 2016 --> 2021 [C2R]"
    write-host "Support Convert / Activation for :: Windows Vista --> Late Windows 11"
    write-host
    write-host "** Keep Origional Activated OEM / Retail / Mak Licences"
    write-host "** Clean Duplicated Licences Of same Products, different year"
    write-host "** Clean Unused Product Licences like :: 365, Home, Professional, Private"
    write-host
    LetsActivate
    Clean-RegistryKeys

    # Windows_Addict Not genuine fix
    $regPaths = @($Global:XSPP_USER, $Global:XSPP_HKLM_X32, $Global:XSPP_HKLM_X64)
    foreach ($path in $regPaths) {
        try {
            New-Item -Path $path -Force -ea 0 | Out-Null
            if (Test-Path $path) {
                Set-ItemProperty -Path $path -Name 'KeyManagementServiceName' -Value $Global:IP_ADDRESS -Type String -Force -ea 0
            }
        } catch {
            #Write-Host "Failed to write to $path" -ForegroundColor Red
        }
    }
}
function Run-Troubleshoot {
param (
    [bool]$AutoMode = $false,
    [bool]$RunUpgrade = $false,
    [bool]$RunWmiRepair = $false,
    [bool]$RunTokenStoreReset = $false,
    [bool]$RunUninstallLicenses = $false,
    [bool]$RunScrubOfficeC2R = $false,
    [bool]$RunOfficeLicenseInstaller = $false,
    [bool]$RunOfficeOnlineInstallation = $false
)
# --> Start
$dicKeepSku = @{}
$Start_Time = $(Get-Date -Format hh:mm:ss)
$IgnoreCase = [Text.RegularExpressions.RegexOptions]::IgnoreCase

Set-Location "HKLM:\"
$sPackageGuid = $null

@("SOFTWARE\Microsoft\Office\15.0\ClickToRun",
  "SOFTWARE\Microsoft\Office\16.0\ClickToRun",
  "SOFTWARE\Microsoft\Office\ClickToRun" ) | % {
    try {
      $sPackageGuid = gpv $_ PackageGUID -ea 0
    } catch{}}
Function Convert-To-System {
   param (
     [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
     [string] $NAME
   )
  
  switch ($NAME){
   "ARM"      {return "ARM"}
   "CHPE"     {return "CHPE"}
   "Win7"     {return "7.0"}
   "Win8"     {return "8.0"}
   "Win8.0"   {return "8.0"}
   "Win8.1"   {return "8.1"}
   "Default"  {return "10.0"}
   "RDX Test" {return "RDX"}
  }
  return "Null"
}
Function Convert-To-Channel {
   param (
     [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
     [string] $FFN
   )
  
  switch ($FFN){
   "492350f6-3a01-4f97-b9c0-c7c6ddf67d60" {return "Current"}
   "64256afe-f5d9-4f86-8936-8840a6a4f5be" {return "CurrentPreview"}
   "5440fd1f-7ecb-4221-8110-145efaa6372f" {return "BetaChannel"}
   "55336b82-a18d-4dd6-b5f6-9e5095c314a6" {return "MonthlyEnterprise"}
   "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" {return "SemiAnnual"}
   "b8f9b850-328d-4355-9145-c59439a0c4cf" {return "SemiAnnualPreview"}
   "f2e724c1-748f-4b47-8fb8-8e0d210e9208" {return "PerpetualVL2019"}
   "5030841d-c919-4594-8d2d-84ae4f96e58e" {return "PerpetualVL2021"}
   "7983BAC0-E531-40CF-BE00-FD24FE66619C" {return "PerpetualVL2024"}
   "ea4a4090-de26-49d7-93c1-91bff9e53fc3" {return "DogfoodDevMain"}
   "f3260cf1-a92c-4c75-b02e-d64c0a86a968" {return "DogfoodCC"}
   "c4a7726f-06ea-48e2-a13a-9d78849eb706" {return "DogfoodDCEXT"}
   "834504cc-dc55-4c6d-9e71-e024d0253f6d" {return "DogfoodFRDC"}
   "5462eee5-1e97-495b-9370-853cd873bb07" {return "MicrosoftCC"}
   "f4f024c8-d611-4748-a7e0-02b6e754c0fe" {return "MicrosoftDC"}
   "b61285dd-d9f7-41f2-9757-8f61cba4e9c8" {return "MicrosoftDevMain"}
   "9a3b7ff2-58ed-40fd-add5-1e5158059d1c" {return "MicrosoftFRDC"}
   "1d2d2ea6-1680-4c56-ac58-a441c8c24ff9" {return "MicrosoftLTSC"}
   "86752282-5841-4120-ac80-db03ae6b5fdb" {return "MicrosoftLTSC2021"}
   "C02D8FE6-5242-4DA8-972F-82EE55E00671" {return "MicrosoftLTSC2024"}
   "2e148de9-61c8-4051-b103-4af54baffbb4" {return "InsidersLTSC"}
   "12f4f6ad-fdea-4d2a-a90f-17496cc19a48" {return "InsidersLTSC2021"}
   "20481F5C-C268-4624-936C-52EB39DDBD97" {return "InsidersLTSC2024"}
   "0002c1ba-b76b-4af9-b1ee-ae2ad587371f" {return "InsidersMEC"}
  }
  return "Null"
}
Function Get-Office-Apps {
   param (
     [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
     [string] $FFN
   )
  
  $ProgressPreference = 'SilentlyContinue'    # Subsequent calls do not display UI.
  $URI = 'https://clients.config.office.net/releases/v1.0/OfficeReleases'
  $URI = 'https://mrodevicemgr.officeapps.live.com/mrodevicemgrsvc/api/v2/C2RReleaseData'
  $REQ = IWR $URI -ea 0

  if (-not $REQ) {
    return $null
  }

  $Json = $REQ.Content | ConvertFrom-Json
  $Json|Sort-Object FFN|select @{Name='Channel'; Expr={$_.FFN|Convert-To-Channel}},FFN,@{Name='Build'; Expr={$_.AvailableBuild}},@{Name='System'; Expr={$_.Name|Convert-To-System}}
}
Function IsC2R {
   param (
     [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
     [string] $Value,
	 
	 [parameter(Mandatory=$false)]
     [bool] $FastSearch
   )

  $OREF          = "^(.*)(\\ROOT\\OFFICE1)(.*)$"
  $MSOFFICE      = "^(.*)(\\Microsoft Office)(.*)$"
  $OREFROOT      = "^(.*)(Microsoft Office\\root\\)(.*)$"
  $OCOMMON	     = "^(.*)(\\microsoft shared\\ClickToRun)(.*)$"

  
  if (($FastSearch -ne $null) -and ($FastSearch -eq $true)) {
	if ([REGEX]::IsMatch(
      $Value,$MSOFFICE,$IgnoreCase)) {
        return $true }
	return $false
  }
  
  if ([REGEX]::IsMatch(
    $Value,$OREF,$IgnoreCase)) {
      return $true }
  if ([REGEX]::IsMatch(
    $Value,$MSOFFICE,$IgnoreCase)) {
      return $true }
  if ([REGEX]::IsMatch(
    $Value,$OREFROOT,$IgnoreCase)) {
      return $true }
  if ([REGEX]::IsMatch(
    $Value,$OCOMMON,$IgnoreCase)) {
      return $true }
             
  return $false
}
Function GetExpandedGuid {
   param (
     [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
     [ValidatePattern('^[0-9a-fA-F]{32}$')]
     [ValidateScript( { [Guid]::Parse($_) -is [Guid] })]
     [string] $sGuid
   )

if (($sGuid.Length -ne 32) -or (
  $sGuid -notmatch '00F01FEC')) {
    return $null }

$output = ""
([ordered]@{
1=$sGuid.ToCharArray(0,8)
2=$sGuid.ToCharArray(8,4)
3=$sGuid.ToCharArray(12,4)}).GetEnumerator() | % {
  [ARRAY]::Reverse($_.Value)
  $output += (-join $_.Value) }
$sArr = $sGuid.ToCharArray()
([ordered]@{
17=20
21=32 }).GetEnumerator() | % {
$_.Key..$_.Value | % {
  if ($_ % 2) {
    $output += $sArr[$_]
} else {
    $output += $sArr[$_-2] }} }
return [Guid]::Parse(
  -join $output).ToString().ToUpper()
}
Function CheckDelete {
   param (
     [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
     [string] $sProductCode
   )

   # FOR GUID FORMAT
   # {90160000-008C-0000-1000-0000000FF1CE}

   # ensure valid GUID length
   if ($sProductCode.Length -ne 38) {
     return $false }	

    # only care if it's in the expected ProductCode pattern
	if (-not(
	  InScope $sProductCode)) {
        return $false }
	
    # check if it's a known product that should be kept
    if ($dicKeepSku.ContainsKey($sProductCode)) {
      return $false }
	
  return $True
}
Function InScope {
   param (
     [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
     [string] $sProductCode
   )
   
   $PRODLEN = 13
   $OFFICEID = "0000000FF1CE}"
   if ($sProductCode.Length -ne 38) {
     return $false }

   $sProd = $sProductCode.ToUpper()
   if ($sProd.Substring($sProd.Length-13,$PRODLEN) -ne $OFFICEID ) {
     if ($sPackageGuid -and ($sProd -eq $sPackageGuid.ToUpper())) {
       return $True }
     switch ($sProductCode)
     {
       "{6C1ADE97-24E1-4AE4-AEDD-86D3A209CE60}" {return $True}
       "{9520DDEB-237A-41DB-AA20-F2EF2360DCEB}" {return $True}
       "{9AC08E99-230B-47e8-9721-4577B7F124EA}" {return $True}
     }
     return $false }
   
   if ([INT]$sProd.Substring(3,2) -gt 14) {
     switch ($sProd.Substring(10,4))
     {
       "007E" {return $True}
       "008F" {return $True}
       "008C" {return $True}
       "24E1" {return $True}
       "237A" {return $True}
       "00DD" {return $True}
       Default {return $false}
     }
   }

    return $false
}
function Get-Shortcut {
<#
.SYNOPSIS
    Get information about a Shortcut (.lnk file)
.DESCRIPTION
    Get information about a Shortcut (.lnk file)
.PARAMETER Path
    File
.EXAMPLE
    Get-Shortcut -Path 'C:\Portable\Test.lnk'
 
    Link : Test.lnk
    TargetPath : C:\Portable\PortableApps\Notepad++Portable\Notepad++Portable.exe
    WindowStyle : 1
    IconLocation : ,0
    Hotkey :
    Target : Notepad++Portable.exe
    Arguments :
    LinkPath : C:\Portable\Test.lnk
#>

    [CmdletBinding(ConfirmImpact='None')]
    param(
        [string] $path
    )

    begin {
        Write-Verbose -Message "Starting [$($MyInvocation.Mycommand)]"
        $obj = New-Object -ComObject WScript.Shell
    }

    process {
        if (Test-Path -Path $Path) {
            $ResolveFile = Resolve-Path -Path $Path
            if ($ResolveFile.count -gt 1) {
                Write-Warning -Message "ERROR: File specification [$File] resolves to more than 1 file."
            } else {
                Write-Verbose -Message "Using file [$ResolveFile] in section [$Section], getting comments"
                $ResolveFile = Get-Item -Path $ResolveFile
                if ($ResolveFile.Extension -eq '.lnk') {
                    $link = $obj.CreateShortcut($ResolveFile.FullName)

                    $info = @{}
                    $info.Hotkey = $link.Hotkey
                    $info.TargetPath = $link.TargetPath
                    $info.LinkPath = $link.FullName
                    $info.Arguments = $link.Arguments
                    $info.Target = try {Split-Path -Path $info.TargetPath -Leaf } catch { 'n/a'}
                    $info.Link = try { Split-Path -Path $info.LinkPath -Leaf } catch { 'n/a'}
                    $info.WindowStyle = $link.WindowStyle
                    $info.IconLocation = $link.IconLocation

                    New-Object -TypeName PSObject -Property $info
                } else {
                    Write-Warning -Message 'Extension is not .lnk'
                }
            }
        } else {
            Write-Warning -Message "ERROR: File [$Path] does not exist"
        }
    }

    end {
        Write-Verbose -Message "Ending [$($MyInvocation.Mycommand)]"
    }
}
Function CleanShortcuts {
   param (
     [parameter(Mandatory=$True)]
     [string] $sFolder
   )

 Set-Location "c:\"

 if (-not (
   Test-Path $sFolder )) {
     return; }

 dir $sFolder -Filter *.lnk -Recurse -ea 0 | % {
    $Shortcut = Get-Shortcut(
      $_.FullName) -ea 0
    if ($Shortcut -and $Shortcut.TargetPath -and (
      $Shortcut.TargetPath|IsC2R)) {
          RI $_.FullName -Force -ea 0  }}
}
function UninstallOfficeC2R {
$URL = 
  "http://officecdn.microsoft.com/pr/wsus/setup.exe"

$Path = 
  "$env:WINDIR\temp\setup.exe"

$XML = 
  "$env:WINDIR\temp\RemoveAll.xml"


$CODE = @"
<Configuration> 
  <Remove All="TRUE"> 
</Remove> 
  <Display Level="None" AcceptEULA="TRUE" />   
</Configuration>
"@

try {
  "*** -- build the remove.xml"
  $CODE | Out-File $XML
  "*** -- ODT not available. Try to download"
  (New-Object WebClient).DownloadFile($URL, $Path)
}
catch { }

Set-Location "$env:SystemDrive\"
Push-Location "$env:WINDIR\temp\"
if ([IO.FILE]::Exists(
  $Path)) {
    $Proc = start $Path -arg "/configure RemoveAll.xml" -Wait -WindowStyle Hidden -PassThru -ea 0
    "*** -- ODT uninstall for OfficeC2R returned with value:$($Proc.ExitCode)" }

if ($Proc -and $Proc.ExitCode -eq 0) {
  "*** -- Use unified ARP uninstall command [No-Need]"
  return }

"*** -- Use unified ARP uninstall command"

try {
  $HashList = GetUninstall }
catch {
  $HashList = $null }

$arrayList = @{}
$OfficeClickToRun = $null

if ($HashList) {
  foreach ($key in $HashList.keys) {
    $value = $HashList[$key]
    if (($value -notlike "*OfficeClickToRun.exe*") -and (
      $false -eq ($value|CheckDelete) )) {
      continue }
    $data  = $value.Split( )
    if ($data) {
      0..$data.Count | % {
        if ($data[$_] -match 'productstoremove=') {
          $data[$_] = "productstoremove=AllProducts" }}
    
    $value   = $data -join (' ')
    $value  += ' displaylevel=false'
    $prefix  = $value.Split('"')
    try {
      $OfficeClickToRun = $prefix[1]
      $arrayList.Add($key,$prefix[2]) }
    catch {}
}}}

foreach ($key in $arrayList.Keys) {
  if ([IO.FILE]::Exists($OfficeClickToRun)) {
    $value = $arrayList[$key]
    $Proc = start $OfficeClickToRun -Arg $value -Wait -WindowStyle Hidden -PassThru -ea 0
    "*** -- uninstall command: $arg, exit code value: $($Proc.ExitCode)"
}}

return
}
Function CloseOfficeApps {
$dicApps = @{}
$dicApps.Add("appvshnotify.exe","appvshnotify.exe")
$dicApps.Add("integratedoffice.exe","integratedoffice.exe")
$dicApps.Add("integrator.exe","integrator.exe")
$dicApps.Add("firstrun.exe","firstrun.exe")
$dicApps.Add("communicator.exe","communicator.exe")
$dicApps.Add("msosync.exe","msosync.exe")
$dicApps.Add("OneNoteM.exe","OneNoteM.exe")
$dicApps.Add("iexplore.exe","iexplore.exe")
$dicApps.Add("mavinject32.exe","mavinject32.exe")
$dicApps.Add("werfault.exe","werfault.exe")
$dicApps.Add("perfboost.exe","perfboost.exe")
$dicApps.Add("roamingoffice.exe","roamingoffice.exe")
$dicApps.Add("officeclicktorun.exe","officeclicktorun.exe")
$dicApps.Add("officeondemand.exe","officeondemand.exe")
$dicApps.Add("OfficeC2RClient.exe","OfficeC2RClient.exe")
$dicApps.Add("explorer.exe","explorer.exe")
$dicApps.Add("msiexec.exe","msiexec.exe")
$dicApps.Add("ose.exe","ose.exe")
$dicList = $dicApps.Values -join "|"

$Process = gwmi -Query "Select * From Win32_Process"
$Process | ? {
  [REGEX]::IsMatch($_.Name,$dicList, $IgnoreCase)} | % {
    try {($_).Terminate()|Out-Null} catch {} }

$Process = gwmi -Query "Select * From Win32_Process"
$Process | % {
  $ExecuePath = ($_).Properties | ? Name -EQ ExecutablePath | select Value
  if ($ExecuePath -and $ExecuePath.Value) {
    if ($ExecuePath.Value|IsC2R) {
        try {
          ($_).Terminate()|Out-Null}
        catch {} }}}
}
Function DelSchtasks {
SCHTASKS /Delete /F /TN "C2RAppVLoggingStart" *>$null
SCHTASKS /Delete /F /TN "FF_INTEGRATEDstreamSchedule" *>$null
SCHTASKS /Delete /F /TN "Microsoft Office 15 Sync Maintenance for {d068b555-9700-40b8-992c-f866287b06c1}" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\Office Automatic Updates 2.0" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\Office Automatic Updates" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\Office ClickToRun Service Monitor" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\Office Feature Updates Logon" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\Office Feature Updates" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\Office Performance Monitor" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\OfficeInventoryAgentFallBack" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\OfficeInventoryAgentLogOn" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\OfficeTelemetryAgentFallBack" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\OfficeTelemetryAgentLogOn" *>$null
SCHTASKS /Delete /F /TN "Office 15 Subscription Heartbeat" *>$null
SCHTASKS /Delete /F /TN "Office Background Streaming" *>$null
SCHTASKS /Delete /F /TN "Office Subscription Maintenance" *>$null
}
Function ClearShellIntegrationReg {
Set-Location "HKLM:\"
RI "HKLM:SOFTWARE\Classes\Protocols\Handler\osf" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Classes\CLSID\{573FFD05-2805-47C2-BCE0-5F19512BEB8D}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Classes\CLSID\{8BA85C75-763B-4103-94EB-9470F12FE0F7}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Classes\CLSID\{CD55129A-B1A1-438E-A425-CEBC7DC684EE}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Classes\CLSID\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Classes\CLSID\{E768CD3B-BDDC-436D-9C13-E1B39CA257B1}" -Force -ea 0 -Recurse

RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 1 (ErrorConflict)" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 2 (SyncInProgress)" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 3 (InSync)" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 1 (ErrorConflict)" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 2 (SyncInProgress)" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 3 (InSync)" -Force -ea 0 -Recurse
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{B28AA736-876B-46DA-B3A8-84C5E30BA492}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{8B02D659-EBBB-43D7-9BBA-52CF22C5B025}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{42042206-2D85-11D3-8CFF-005004838597}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{993BE281-6695-4BA5-8A2A-7AACBFAAB69E}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{C41662BB-1FA0-4CE0-8DC5-9B7F8279FF97}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{506F4668-F13E-4AA1-BB04-B43203AB3CC0}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{D66DC78C-4F61-447F-942B-3FB6980118CF}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{46137B78-0EC3-426D-8B89-FF7C3A458B5E}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{8BA85C75-763B-4103-94EB-9470F12FE0F7}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{CD55129A-B1A1-438E-A425-CEBC7DC684EE}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{E768CD3B-BDDC-436D-9C13-E1B39CA257B1}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{E768CD3B-BDDC-436D-9C13-E1B39CA257B1}" -Force -ea 0
RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\Namespace\{B28AA736-876B-46DA-B3A8-84C5E30BA492}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\NetworkNeighborhood\Namespace\{46137B78-0EC3-426D-8B89-FF7C3A458B5E}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Microsoft Office Temp Files" -Force -ea 0 -Recurse
}
function Get-MsiProducts {
  
  # PowerShell: Get-MsiProducts / Get Windows Installer Products
  # https://gist.github.com/MyITGuy/153fc0f553d840631269720a56be5136#file-file01-ps1

    function Get-MsiUpgradeCode {
        [CmdletBinding()]
        param (
            [Guid]$ProductCode,
            [Guid]$UpgradeCode
        )
        function ConvertFrom-CompressedGuid {
            <#
	        .SYNOPSIS
		        Converts a compressed globally unique identifier (GUID) string into a GUID string.
	        .DESCRIPTION
            Takes a compressed GUID string and breaks it into 6 parts. It then loops through the first five parts and reversing the order. It loops through the sixth part and reversing the order of every 2 characters. It then joins the parts back together and returns a GUID.
	        .EXAMPLE
		        ConvertFrom-CompressedGuid -CompressedGuid '2820F6C7DCD308A459CABB92E828C144'
	
		        The output of this example would be: {7C6F0282-3DCD-4A80-95AC-BB298E821C44}
	        .PARAMETER CompressedGuid
		        A compressed globally unique identifier (GUID) string.
	        #>
            [CmdletBinding()]
            [OutputType([String])]
            param (
                [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
                [ValidatePattern('^[0-9a-fA-F]{32}$')]
                [ValidateScript( { [Guid]::Parse($_) -is [Guid] })]
                [String]$CompressedGuid
            )
            process {
                Write-Verbose "CompressedGuid: $($CompressedGuid)"
                $GuidString = ([Guid]$CompressedGuid).ToString('N')
                Write-Verbose "GuidString: $($GuidString)"
                $Indexes = [ordered]@{
                    0  = 8
                    8  = 4
                    12 = 4
                    16 = 2
                    18 = 2
                    20 = 12
                }
                $Guid = ''
                foreach ($key in $Indexes.Keys) {
                    $value = $Indexes[$key]
                    $Substring = $GuidString.Substring($key, $value)
                    Write-Verbose "Substring: $($Substring)"
                    switch ($key) {
                        20 {
                            $parts = $Substring -split '(.{2})' | Where-Object { $_ }
                            foreach ($part In $parts) {
                                $part = $part -split '(.{1})'
                                Write-Verbose "Part: $($part)"
                                [Array]::Reverse($part)
                                Write-Verbose "Reversed: $($part)"
                                $Guid += $part -join ''
                            }
                        }
                        default {
                            $part = $Substring.ToCharArray()
                            Write-Verbose "Part: $($part)"
                            [Array]::Reverse($part)
                            Write-Verbose "Reversed: $($part)"
                            $Guid += $part -join ''
                        }
                    }
                }
                [Guid]::Parse($Guid).ToString('B').ToUpper()
            }
        }

        function ConvertTo-CompressedGuid {
            <#
	        .SYNOPSIS
		        Converts a GUID string into a compressed globally unique identifier (GUID) string.
	        .DESCRIPTION
		        Takes a GUID string and breaks it into 6 parts. It then loops through the first five parts and reversing the order. It loops through the sixth part and reversing the order of every 2 characters. It then joins the parts back together and returns a compressed GUID string.
	        .EXAMPLE
		        ConvertTo-CompressedGuid -Guid '{7C6F0282-3DCD-4A80-95AC-BB298E821C44}'
	
            The output of this example would be: 2820F6C7DCD308A459CABB92E828C144
	        .PARAMETER Guid
            A globally unique identifier (GUID).
	        #>
            [CmdletBinding()]
            [OutputType([String])]
            param (
                [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
                [ValidateScript( { [Guid]::Parse($_) -is [Guid] })]
                [Guid]$Guid
            )
            process {
                Write-Verbose "Guid: $($Guid)"
                $GuidString = $Guid.ToString('N')
                Write-Verbose "GuidString: $($GuidString)"
                $Indexes = [ordered]@{
                    0  = 8
                    8  = 4
                    12 = 4
                    16 = 2
                    18 = 2
                    20 = 12
                }
                $CompressedGuid = ''
                foreach ($key in $Indexes.keys) {
                    $value = $Indexes[$key]
                    $Substring = $GuidString.Substring($key, $value)
                    Write-Verbose "Substring: $($Substring)"
                    switch ($key) {
                        20 {
                            $parts = $Substring -split '(.{2})' | Where-Object { $_ }
                            foreach ($part In $parts) {
                                $part = $part -split '(.{1})'
                                Write-Verbose "Part: $($part)"
                                [Array]::Reverse($part)
                                Write-Verbose "Reversed: $($part)"
                                $CompressedGuid += $part -join ''
                            }
                        }
                        default {
                            $part = $Substring.ToCharArray()
                            Write-Verbose "Part: $($part)"
                            [Array]::Reverse($part)
                            Write-Verbose "Reversed: $($part)"
                            $CompressedGuid += $part -join ''
                        }
                    }
                }
                [Guid]::Parse($CompressedGuid).ToString('N').ToUpper()
            }
        }

        filter ByProductCode {
            $Object = $_
            Write-Verbose "ProductCode: $($ProductCode)"
            if ($ProductCode) {
                $Object | Where-Object { [Guid]($_.ProductCode) -eq [Guid]($ProductCode) }
                break
            }
            $Object
        }

        $Path = "Registry::HKEY_CLASSES_ROOT\Installer\UpgradeCodes\*"
        if ($UpgradeCode) {
            $CompressedUpgradeCode = ConvertTo-CompressedGuid -Guid $UpgradeCode -Verbose:$false
            Write-Verbose "CompressedUpgradeCode: $($CompressedUpgradeCode)"
            $Path = "Registry::HKEY_CLASSES_ROOT\Installer\UpgradeCodes\$($CompressedUpgradeCode)"
        }

        Get-Item -Path $Path -ErrorAction SilentlyContinue | ForEach-Object {
            $UpgradeCodeFromCompressedGuid = ConvertFrom-CompressedGuid -CompressedGuid $_.PSChildName -Verbose:$false
            foreach ($ProductCodeCompressedGuid in ($_.GetValueNames())) {
                $Properties = [ordered]@{
                    ProductCode = ConvertFrom-CompressedGuid -CompressedGuid $ProductCodeCompressedGuid -Verbose:$false
                    UpgradeCode = [Guid]::Parse($UpgradeCodeFromCompressedGuid).ToString('B').ToUpper()
                }
                [PSCustomObject]$Properties | ByProductCode
            }
        }
    }

    $MsiUpgradeCodes = Get-MsiUpgradeCode

    $Installer = New-Object -ComObject WindowsInstaller.Installer
	$Type = $Installer.GetType()
	$Products = $Type.InvokeMember('Products', [BindingFlags]::GetProperty, $null, $Installer, $null)
	foreach ($Product In $Products) {
		$hash = @{}
		$hash.ProductCode = $Product
		$Attributes = @('Language', 'ProductName', 'PackageCode', 'Transforms', 'AssignmentType', 'PackageName', 'InstalledProductName', 'VersionString', 'RegCompany', 'RegOwner', 'ProductID', 'ProductIcon', 'InstallLocation', 'InstallSource', 'InstallDate', 'Publisher', 'LocalPackage', 'HelpLink', 'HelpTelephone', 'URLInfoAbout', 'URLUpdateInfo')		
		foreach ($Attribute In $Attributes) {
			$hash."$($Attribute)" = $null
		}
		foreach ($Attribute In $Attributes) {
			try {
				$hash."$($Attribute)" = $Type.InvokeMember('ProductInfo', [BindingFlags]::GetProperty, $null, $Installer, @($Product, $Attribute))
			} catch [Exception] {
				#$error[0]|format-list -force
			}
		}
        
        # UpgradeCode
        $hash.UpgradeCode = $MsiUpgradeCodes | Where-Object ProductCode -eq ($hash.ProductCode) | Select-Object -ExpandProperty UpgradeCode

		New-Object -TypeName PSObject -Property $hash
	}
}
function UninstallLicenses($DllPath) {
  
  # https://github.com/ave9858
  # https://gist.github.com/ave9858/9fff6af726ba3ddc646285d1bbf37e71

    $DynAssembly = New-Object AssemblyName('Win32Lib')
    $AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly($DynAssembly, [AssemblyBuilderAccess]::Run)
    $ModuleBuilder = $AssemblyBuilder.DefineDynamicModule('Win32Lib', $False)
    $TypeBuilder = $ModuleBuilder.DefineType('sppc', 'Public, Class')
    $DllImportConstructor = [DllImportAttribute].GetConstructor(@([String]))
    $FieldArray = [Reflection.FieldInfo[]] @([DllImportAttribute].GetField('EntryPoint'))

    $Open = $TypeBuilder.DefineMethod('SLOpen', [Reflection.MethodAttributes] 'Public, Static', [int], @([IntPtr].MakeByRefType()))
    $Open.SetCustomAttribute((New-Object CustomAttributeBuilder(
                $DllImportConstructor,
                @($DllPath),
                $FieldArray,
                @('SLOpen'))))

    $GetSLIDList = $TypeBuilder.DefineMethod('SLGetSLIDList', [Reflection.MethodAttributes] 'Public, Static', [int], @([IntPtr], [int], [guid].MakeByRefType(), [int], [int].MakeByRefType(), [IntPtr].MakeByRefType()))
    $GetSLIDList.SetCustomAttribute((New-Object CustomAttributeBuilder(
                $DllImportConstructor,
                @($DllPath),
                $FieldArray,
                @('SLGetSLIDList'))))

    $UninstallLicense = $TypeBuilder.DefineMethod('SLUninstallLicense', [Reflection.MethodAttributes] 'Public, Static', [int], @([IntPtr], [IntPtr]))
    $UninstallLicense.SetCustomAttribute((New-Object CustomAttributeBuilder(
                $DllImportConstructor,
                @($DllPath),
                $FieldArray,
                @('SLUninstallLicense'))))

    $SPPC = $TypeBuilder.CreateType()
    $Handle = [IntPtr]::Zero
    $SPPC::SLOpen([ref]$handle) | Out-Null
    $pnReturnIds = 0
    $ppReturnIds = [IntPtr]::Zero

    if (!$SPPC::SLGetSLIDList($handle, 0, [ref][guid]"0ff1ce15-a989-479d-af46-f275c6370663", 6, [ref]$pnReturnIds, [ref]$ppReturnIds)) {
        foreach ($i in 0..($pnReturnIds - 1)) {
            $SPPC::SLUninstallLicense($handle, [Int64]$ppReturnIds + [Int64]16 * $i) | Out-Null
        }    
    }
}
function GetUninstall {
$UninstallArr  = @{}
$UninstallKeys = @{}
$UninstallKeys.Add(1,"HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
$UninstallKeys.Add(2,"HKLM:SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall")

foreach ($sKey in $UninstallKeys.Values) {
  Set-Location "HKLM:\"
  Push-Location $sKey -ea 0

  if (@(Get-Location).Path -NE 'HKLM:\') {
    $children = gci .
    $children | % {
      $sName = $_.Name.Replace('HKEY_LOCAL_MACHINE','HKLM:')
      $sGuid = $sName.Split('\')|select -Last 1
      Set-Location "HKLM:\"; Push-Location "$sName" -ea 0
      if (@(Get-Location).Path -NE 'HKLM:\') {
        try {
          $UninstallString = $null
          $UninstallString = gpv . -Name 'UninstallString' -ea 0 }
        catch {}
        if ($UninstallString -and (
          $UninstallString|IsC2R)) {

            try {
              $UninstallArr.Add(
                $sGuid, $UninstallString)}
            catch {}}}}}
}

return $UninstallArr
}
function CleanUninstall {

$UninstallKeys = @{}
$UninstallKeys.Add(1,"HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
$UninstallKeys.Add(2,"HKLM:SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall")

foreach ($sKey in $UninstallKeys.Values) {
Set-Location "HKLM:\"
Push-Location $sKey -ea 0

if (@(Get-Location).Path -NE 'HKLM:\') {
  $children = gci .
  $children | % {
    $sName = $_.Name.Replace('HKEY_LOCAL_MACHINE','HKLM:')
    $sGuid = $sName.Split('\')|select -Last 1
    
    Set-Location "HKLM:\"
    Push-Location "$sName" -ea 0
    if (@(Get-Location).Path -NE 'HKLM:\') {
      try {
        $InstallLocation = $null
        $InstallLocation = gpv . -Name 'InstallLocation' -ea 0 }
      catch {}

      if (($sGuid -and ($sGuid|CheckDelete)) -or (
        $InstallLocation -and ($InstallLocation|IsC2R))) {
          Set-Location "HKLM:\"
          RI $sName -Recurse -Force }
    }}}}
}
Function RegWipe {

CloseOfficeApps

"*** -- C2R specifics"
"*** -- Virtual InstallRoot"
"*** -- Mapi Search reg"
"*** -- Office key in HKLM"

Set-Location "HKLM:\"
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Run" Lync15 -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Run" Lync16 -Force -ea 0
RI "HKLM:SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot\Virtual" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot\Virtual" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\Common\InstallRoot\Virtual" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Classes\CLSID\{2027FC3B-CF9D-4ec7-A823-38BA308625CC}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\15.0\ClickToRun" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\15.0\ClickToRunStore" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\16.0\ClickToRun" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\16.0\ClickToRunStore" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\ClickToRun" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\ClickToRunStore" -Force -ea 0 -Recurse
RI "HKLM:Software\Microsoft\Office\15.0" -Force -ea 0 -Recurse
RI "HKLM:Software\Microsoft\Office\16.0" -Force -ea 0 -Recurse

"*** -- HKCU Registration"
Set-Location "HKCU:\"
RI "HKCU:Software\Microsoft\Office\15.0\Registration" -Force -ea 0 -Recurse
RI "HKCU:Software\Microsoft\Office\16.0\Registration" -Force -ea 0 -Recurse
RI "HKCU:Software\Microsoft\Office\Registration" -Force -ea 0 -Recurse
RI "HKCU:SOFTWARE\Microsoft\Office\15.0\ClickToRun" -Force -ea 0 -Recurse
RI "HKCU:SOFTWARE\Microsoft\Office\16.0\ClickToRun" -Force -ea 0 -Recurse
RI "HKCU:SOFTWARE\Microsoft\Office\ClickToRun" -Force -ea 0 -Recurse
RI "HKCU:Software\Microsoft\Office\15.0" -Force -ea 0 -Recurse
RI "HKCU:Software\Microsoft\Office\16.0" -Force -ea 0 -Recurse

"*** -- App Paths"
$Keys = reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths" 2>$null
$keys | % {
$value = reg query "$_" /ve /t REG_SZ 2>$null
if ($value -match "\\Microsoft Office") {
  reg delete $_ /f | Out-Null }}

"*** -- Run key"
$hDefKey = "HKLM"
$sSubKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
Set-Location "$($hDefKey):\"
Push-Location "$($hDefKey):$($sSubKeyName)" -ea 0

if (@(Get-Location).Path -ne "$($hDefKey):\") {
  $arrNames = gi .
  if ($arrNames)  {
    $arrNames.Property | % { 
      $name = GPV . $_
      if ($name -and (
        $Name|IsC2R)) {
          RP . $_ -Force
}}}}

"*** -- Un-install Keys"
CleanUninstall

"*** -- UpgradeCodes, WI config, WI global config"
"*** -- msiexec based uninstall [Fail-Safe]"

# First here ... 
# HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products

$hash     = $null;
$HashList = $null;
$sKey     = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products"

Set-Location 'HKLM:\'
$sProducts = 
  GCI $sKey -ea 0
$HashList = $sProducts | % {
  ($_).PSPath.Split('\') | select -Last 1 | % {
    [PSCustomObject]@{
    cGuid = $_
    sGuid = ($_|GetExpandedGuid) }}}
$GuidList = 
  $HashList | ? sGuid

if ($GuidList) {
  $GuidList | ? sGuid | % {
    $Proc = $null
    $ProductCode = $_.sGuid
    $sMsiProp = "REBOOT=ReallySuppress NOREMOVESPAWN=True"
    $sUninstallCmd = "/x {$($ProductCode)} $($sMsiProp) /q"

    if ($ProductCode) {
      $Proc = start msiexec.exe -Args $sUninstallCmd -Wait -WindowStyle Hidden -ea 0 -PassThru
      "*** -- Msiexec $($sUninstallCmd) ,End with value: $($proc.ExitCode)" }

    Set-Location 'HKLM:\'
    RI "$sKey\$($_.sGuid)" -Force -Recurse -ea 0 | Out-Null
    Set-Location 'HKCR:\'
    RI "HKCR:\Installer\Products\$($_.sGuid)" -Force -Recurse -ea 0 | Out-Null }}

# Second here ... 
# HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes

$hash     = $null;
$HashList = $null;
$sKey     = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes"

Set-Location 'HKLM:\'
$sUpgradeCodes = 
  GCI $sKey -ea 0
$HashList = $sUpgradeCodes | % {
  ($_).PSPath.Split('\') | select -Last 1 | % {
    [PSCustomObject]@{
    cGuid = $_
    sGuid = ($_|GetExpandedGuid) }}}
$GuidList = 
  $HashList | ? sGuid

if ($GuidList) {
  $GuidList | % {
    Set-Location 'HKLM:\'
    RI "$sKey\$($_.sGuid)" -Force -Recurse -ea 0 | Out-Null
    Set-Location 'HKCR:\'
    RI "HKCR:\Installer\UpgradeCodes\$($_.sGuid)" -Force -Recurse -ea 0 | Out-Null }}

# make sure we clean everything
$sKeyToRe = @{}
$sKeyList = (
  "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes",
  "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products" )

foreach ($sKey in $sKeyList)
{
  Set-Location "HKLM:\"
  $sKey = "HKLM:" + $sKey

  Set-Location "HKLM:\"
  Push-Location $sKey -ea 0

  if (@(Get-Location).Path -NE 'HKLM:\') {
    $children = gci .
    $children | % {
      $sName = $_.Name.Replace('HKEY_LOCAL_MACHINE','HKLM:')
      $sGuid = $sName.Split('\')|select -Last 1

      Set-Location "HKLM:\"
      Push-Location $sName -ea 0
      if (@(Get-Location).Path -NE 'HKLM:\') {

      $InstallSource   = $null
      $UninstallString = $null
    
      try {
        $InstallSource   = GPV "InstallProperties" -Name InstallSource   -ea 0
        $UninstallString = GPV "InstallProperties" -Name UninstallString -ea 0 }
      catch { }
    
      $CheckOfficeApp = $null
      $CheckOfficeApp = ($sGuid -and ($sGuid|CheckDelete)) -or (
        $InstallSource -and $UninstallString -and ($InstallSource|ISC2R) -and (
        [REGEX]::Match($UninstallString, "^.*{(.*)}.*$",$IgnoreCase)))

      if ($CheckOfficeApp -eq $true) {
        $Matches = [REGEX]::Matches($UninstallString,"^.*{(.*)}.*$",
          $IgnoreCase)
        try {
          $ProductCode = $null
          $ProductCode = $Matches[0].Groups[1].Value }
        catch {}

        $proc = $null
        $sMsiProp = "REBOOT=ReallySuppress NOREMOVESPAWN=True"
        $sUninstallCmd = "/x {$($ProductCode)} $($sMsiProp) /q"

        if ($ProductCode) {
          $proc = start msiexec.exe -Args $sUninstallCmd -Wait -WindowStyle Hidden -ea 0 -PassThru
          "*** -- mSiexec $($sUninstallCmd) ,End with value: $($proc.ExitCode)"
		  $sKeyToRe.Add($sName,$sName) }
}}}}}

Set-Location "HKLM:\"
$sKeyToRe.Keys | % {
  RI $_ -Force -Recurse -ea 0 | Out-Null }

Set-Location "HKCR:\"
$sKeyToRe.Keys | % {
  $GUID = ($_).Split('\') | Select-Object -Last 1
  if ($GUID) {
    RI "HKCR:\Installer\Products\$GUID" -Force -Recurse -ea 0 | Out-Null }}

"*** -- Known Typelib Registration"
RegWipeTypeLib

"*** -- Published Components [JAWOT]"

"*** -- ActiveX/COM Components [JAWOT]"
$COM = (
"{00020800-0000-0000-C000-000000000046}","{00020803-0000-0000-C000-000000000046}",
"{00020812-0000-0000-C000-000000000046}","{00020820-0000-0000-C000-000000000046}",
"{00020821-0000-0000-C000-000000000046}","{00020827-0000-0000-C000-000000000046}",
"{00020830-0000-0000-C000-000000000046}","{00020832-0000-0000-C000-000000000046}",
"{00020833-0000-0000-C000-000000000046}","{00020906-0000-0000-C000-000000000046}",
"{00020907-0000-0000-C000-000000000046}","{000209F0-0000-0000-C000-000000000046}",
"{000209F4-0000-0000-C000-000000000046}","{000209F5-0000-0000-C000-000000000046}",
"{000209FE-0000-0000-C000-000000000046}","{000209FF-0000-0000-C000-000000000046}",
"{00024500-0000-0000-C000-000000000046}","{00024502-0000-0000-C000-000000000046}",
"{00024505-0016-0000-C000-000000000046}","{048EB43E-2059-422F-95E0-557DA96038AF}",
"{18A06B6B-2F3F-4E2B-A611-52BE631B2D22}","{1B261B22-AC6A-4E68-A870-AB5080E8687B}",
"{1CDC7D25-5AA3-4DC4-8E0C-91524280F806}","{3C18EAE4-BC25-4134-B7DF-1ECA1337DDDC}",
"{64818D10-4F9B-11CF-86EA-00AA00B929E8}","{64818D11-4F9B-11CF-86EA-00AA00B929E8}",
"{65235197-874B-4A07-BDC5-E65EA825B718}","{73720013-33A0-11E4-9B9A-00155D152105}",
"{75D01070-1234-44E9-82F6-DB5B39A47C13}","{767A19A0-3CC7-415B-9D08-D48DD7B8557D}",
"{84F66100-FF7C-4fb4-B0C0-02CD7FB668FE}","{8A624388-AA27-43E0-89F8-2A12BFF7BCCD}",
"{912ABC52-36E2-4714-8E62-A8B73CA5E390}","{91493441-5A91-11CF-8700-00AA0060263B}",
"{AA14F9C9-62B5-4637-8AC4-8F25BF29D5A7}","{C282417B-2662-44B8-8A94-3BFF61C50900}",
"{CF4F55F4-8F87-4D47-80BB-5808164BB3F8}","{DC020317-E6E2-4A62-B9FA-B3EFE16626F4}",
"{EABCECDB-CC1C-4A6F-B4E3-7F888A5ADFC8}","{F4754C9B-64F5-4B40-8AF4-679732AC0607}")

#Set-Location "HKCR:\"
$COM | % {
  # will not work .. why ? don't know
  # ri "HKCR:CLSID\$_" -Recurse -Force -ea 0 
}

"*** -- TypeLib Interface [JAWOT]"
$interface = @(
"{000672AC-0000-0000-C000-000000000046}","{000C0300-0000-0000-C000-000000000046}"
"{000C0301-0000-0000-C000-000000000046}","{000C0302-0000-0000-C000-000000000046}"
"{000C0304-0000-0000-C000-000000000046}","{000C0306-0000-0000-C000-000000000046}"
"{000C0308-0000-0000-C000-000000000046}","{000C030A-0000-0000-C000-000000000046}"
"{000C030C-0000-0000-C000-000000000046}","{000C030D-0000-0000-C000-000000000046}"
"{000C030E-0000-0000-C000-000000000046}","{000C0310-0000-0000-C000-000000000046}"
"{000C0311-0000-0000-C000-000000000046}","{000C0312-0000-0000-C000-000000000046}"
"{000C0313-0000-0000-C000-000000000046}","{000C0314-0000-0000-C000-000000000046}"
"{000C0315-0000-0000-C000-000000000046}","{000C0316-0000-0000-C000-000000000046}"
"{000C0317-0000-0000-C000-000000000046}","{000C0318-0000-0000-C000-000000000046}"
"{000C0319-0000-0000-C000-000000000046}","{000C031A-0000-0000-C000-000000000046}"
"{000C031B-0000-0000-C000-000000000046}","{000C031C-0000-0000-C000-000000000046}"
"{000C031D-0000-0000-C000-000000000046}","{000C031E-0000-0000-C000-000000000046}"
"{000C031F-0000-0000-C000-000000000046}","{000C0320-0000-0000-C000-000000000046}"
"{000C0321-0000-0000-C000-000000000046}","{000C0322-0000-0000-C000-000000000046}"
"{000C0324-0000-0000-C000-000000000046}","{000C0326-0000-0000-C000-000000000046}"
"{000C0328-0000-0000-C000-000000000046}","{000C032E-0000-0000-C000-000000000046}"
"{000C0330-0000-0000-C000-000000000046}","{000C0331-0000-0000-C000-000000000046}"
"{000C0332-0000-0000-C000-000000000046}","{000C0333-0000-0000-C000-000000000046}"
"{000C0334-0000-0000-C000-000000000046}","{000C0337-0000-0000-C000-000000000046}"
"{000C0338-0000-0000-C000-000000000046}","{000C0339-0000-0000-C000-000000000046}"
"{000C033A-0000-0000-C000-000000000046}","{000C033B-0000-0000-C000-000000000046}"
"{000C033D-0000-0000-C000-000000000046}","{000C033E-0000-0000-C000-000000000046}"
"{000C0340-0000-0000-C000-000000000046}","{000C0341-0000-0000-C000-000000000046}"
"{000C0353-0000-0000-C000-000000000046}","{000C0356-0000-0000-C000-000000000046}"
"{000C0357-0000-0000-C000-000000000046}","{000C0358-0000-0000-C000-000000000046}"
"{000C0359-0000-0000-C000-000000000046}","{000C035A-0000-0000-C000-000000000046}"
"{000C0360-0000-0000-C000-000000000046}","{000C0361-0000-0000-C000-000000000046}"
"{000C0362-0000-0000-C000-000000000046}","{000C0363-0000-0000-C000-000000000046}"
"{000C0364-0000-0000-C000-000000000046}","{000C0365-0000-0000-C000-000000000046}"
"{000C0366-0000-0000-C000-000000000046}","{000C0367-0000-0000-C000-000000000046}"
"{000C0368-0000-0000-C000-000000000046}","{000C0369-0000-0000-C000-000000000046}"
"{000C036A-0000-0000-C000-000000000046}","{000C036C-0000-0000-C000-000000000046}"
"{000C036D-0000-0000-C000-000000000046}","{000C036E-0000-0000-C000-000000000046}"
"{000C036F-0000-0000-C000-000000000046}","{000C0370-0000-0000-C000-000000000046}"
"{000C0371-0000-0000-C000-000000000046}","{000C0372-0000-0000-C000-000000000046}"
"{000C0373-0000-0000-C000-000000000046}","{000C0375-0000-0000-C000-000000000046}"
"{000C0376-0000-0000-C000-000000000046}","{000C0377-0000-0000-C000-000000000046}"
"{000C0379-0000-0000-C000-000000000046}","{000C037A-0000-0000-C000-000000000046}"
"{000C037B-0000-0000-C000-000000000046}","{000C037C-0000-0000-C000-000000000046}"
"{000C037D-0000-0000-C000-000000000046}","{000C037E-0000-0000-C000-000000000046}"
"{000C037F-0000-0000-C000-000000000046}","{000C0380-0000-0000-C000-000000000046}"
"{000C0381-0000-0000-C000-000000000046}","{000C0382-0000-0000-C000-000000000046}"
"{000C0385-0000-0000-C000-000000000046}","{000C0386-0000-0000-C000-000000000046}"
"{000C0387-0000-0000-C000-000000000046}","{000C0388-0000-0000-C000-000000000046}"
"{000C0389-0000-0000-C000-000000000046}","{000C038A-0000-0000-C000-000000000046}"
"{000C038B-0000-0000-C000-000000000046}","{000C038C-0000-0000-C000-000000000046}"
"{000C038E-0000-0000-C000-000000000046}","{000C038F-0000-0000-C000-000000000046}"
"{000C0390-0000-0000-C000-000000000046}","{000C0391-0000-0000-C000-000000000046}"
"{000C0392-0000-0000-C000-000000000046}","{000C0393-0000-0000-C000-000000000046}"
"{000C0395-0000-0000-C000-000000000046}","{000C0396-0000-0000-C000-000000000046}"
"{000C0397-0000-0000-C000-000000000046}","{000C0398-0000-0000-C000-000000000046}"
"{000C0399-0000-0000-C000-000000000046}","{000C039A-0000-0000-C000-000000000046}"
"{000C03A0-0000-0000-C000-000000000046}","{000C03A1-0000-0000-C000-000000000046}"
"{000C03A2-0000-0000-C000-000000000046}","{000C03A3-0000-0000-C000-000000000046}"
"{000C03A4-0000-0000-C000-000000000046}","{000C03A5-0000-0000-C000-000000000046}"
"{000C03A6-0000-0000-C000-000000000046}","{000C03A7-0000-0000-C000-000000000046}"
"{000C03B2-0000-0000-C000-000000000046}","{000C03B9-0000-0000-C000-000000000046}"
"{000C03BA-0000-0000-C000-000000000046}","{000C03BB-0000-0000-C000-000000000046}"
"{000C03BC-0000-0000-C000-000000000046}","{000C03BD-0000-0000-C000-000000000046}"
"{000C03BE-0000-0000-C000-000000000046}","{000C03BF-0000-0000-C000-000000000046}"
"{000C03C0-0000-0000-C000-000000000046}","{000C03C1-0000-0000-C000-000000000046}"
"{000C03C2-0000-0000-C000-000000000046}","{000C03C3-0000-0000-C000-000000000046}"
"{000C03C4-0000-0000-C000-000000000046}","{000C03C5-0000-0000-C000-000000000046}"
"{000C03C6-0000-0000-C000-000000000046}","{000C03C7-0000-0000-C000-000000000046}"
"{000C03C8-0000-0000-C000-000000000046}","{000C03C9-0000-0000-C000-000000000046}"
"{000C03CA-0000-0000-C000-000000000046}","{000C03CB-0000-0000-C000-000000000046}"
"{000C03CC-0000-0000-C000-000000000046}","{000C03CD-0000-0000-C000-000000000046}"
"{000C03CE-0000-0000-C000-000000000046}","{000C03CF-0000-0000-C000-000000000046}"
"{000C03D0-0000-0000-C000-000000000046}","{000C03D1-0000-0000-C000-000000000046}"
"{000C03D2-0000-0000-C000-000000000046}","{000C03D3-0000-0000-C000-000000000046}"
"{000C03D4-0000-0000-C000-000000000046}","{000C03D5-0000-0000-C000-000000000046}"
"{000C03D6-0000-0000-C000-000000000046}","{000C03D7-0000-0000-C000-000000000046}"
"{000C03E0-0000-0000-C000-000000000046}","{000C03E1-0000-0000-C000-000000000046}"
"{000C03E2-0000-0000-C000-000000000046}","{000C03E3-0000-0000-C000-000000000046}"
"{000C03E4-0000-0000-C000-000000000046}","{000C03E5-0000-0000-C000-000000000046}"
"{000C03E6-0000-0000-C000-000000000046}","{000C03F0-0000-0000-C000-000000000046}"
"{000C03F1-0000-0000-C000-000000000046}","{000C0410-0000-0000-C000-000000000046}"
"{000C0411-0000-0000-C000-000000000046}","{000C0913-0000-0000-C000-000000000046}"
"{000C0914-0000-0000-C000-000000000046}","{000C0936-0000-0000-C000-000000000046}"
"{000C1530-0000-0000-C000-000000000046}","{000C1531-0000-0000-C000-000000000046}"
"{000C1532-0000-0000-C000-000000000046}","{000C1533-0000-0000-C000-000000000046}"
"{000C1534-0000-0000-C000-000000000046}","{000C1709-0000-0000-C000-000000000046}"
"{000C170B-0000-0000-C000-000000000046}","{000C170F-0000-0000-C000-000000000046}"
"{000C1710-0000-0000-C000-000000000046}","{000C1711-0000-0000-C000-000000000046}"
"{000C1712-0000-0000-C000-000000000046}","{000C1713-0000-0000-C000-000000000046}"
"{000C1714-0000-0000-C000-000000000046}","{000C1715-0000-0000-C000-000000000046}"
"{000C1716-0000-0000-C000-000000000046}","{000C1717-0000-0000-C000-000000000046}"
"{000C1718-0000-0000-C000-000000000046}","{000C171B-0000-0000-C000-000000000046}"
"{000C171C-0000-0000-C000-000000000046}","{000C1723-0000-0000-C000-000000000046}"
"{000C1724-0000-0000-C000-000000000046}","{000C1725-0000-0000-C000-000000000046}"
"{000C1726-0000-0000-C000-000000000046}","{000C1727-0000-0000-C000-000000000046}"
"{000C1728-0000-0000-C000-000000000046}","{000C1729-0000-0000-C000-000000000046}"
"{000C172A-0000-0000-C000-000000000046}","{000C172B-0000-0000-C000-000000000046}"
"{000C172C-0000-0000-C000-000000000046}","{000C172D-0000-0000-C000-000000000046}"
"{000C172E-0000-0000-C000-000000000046}","{000C172F-0000-0000-C000-000000000046}"
"{000C1730-0000-0000-C000-000000000046}","{000C1731-0000-0000-C000-000000000046}"
"{000CD100-0000-0000-C000-000000000046}","{000CD101-0000-0000-C000-000000000046}"
"{000CD102-0000-0000-C000-000000000046}","{000CD6A1-0000-0000-C000-000000000046}"
"{000CD6A2-0000-0000-C000-000000000046}","{000CD6A3-0000-0000-C000-000000000046}"
"{000CD706-0000-0000-C000-000000000046}","{000CD809-0000-0000-C000-000000000046}"
"{000CD900-0000-0000-C000-000000000046}","{000CD901-0000-0000-C000-000000000046}"
"{000CD902-0000-0000-C000-000000000046}","{000CD903-0000-0000-C000-000000000046}"
"{000CDB00-0000-0000-C000-000000000046}","{000CDB01-0000-0000-C000-000000000046}"
"{000CDB02-0000-0000-C000-000000000046}","{000CDB03-0000-0000-C000-000000000046}"
"{000CDB04-0000-0000-C000-000000000046}","{000CDB05-0000-0000-C000-000000000046}"
"{000CDB06-0000-0000-C000-000000000046}","{000CDB09-0000-0000-C000-000000000046}"
"{000CDB0A-0000-0000-C000-000000000046}","{000CDB0E-0000-0000-C000-000000000046}"
"{000CDB0F-0000-0000-C000-000000000046}","{000CDB10-0000-0000-C000-000000000046}"
"{00194002-D9C3-11D3-8D59-0050048384E3}","{4291224C-DEFE-485B-8E69-6CF8AA85CB76}"
"{4B0F95AC-5769-40E9-98DF-80CDD086612E}","{4CAC6328-B9B0-11D3-8D59-0050048384E3}"
"{55F88890-7708-11D1-ACEB-006008961DA5}","{55F88892-7708-11D1-ACEB-006008961DA5}"
"{55F88896-7708-11D1-ACEB-006008961DA5}","{6EA00553-9439-4D5A-B1E6-DC15A54DA8B2}"
"{88FF5F69-FACF-4667-8DC8-A85B8225DF15}","{8A64A872-FC6B-4D4A-926E-3A3689562C1C}"
"{919AA22C-B9AD-11D3-8D59-0050048384E3}","{A98639A1-CB0C-4A5C-A511-96547F752ACD}"
"{ABFA087C-F703-4D53-946E-37FF82B2C994}","{D996597A-0E80-4753-81FC-DCF16BDF4947}"
"{DE9CD4FF-754A-49DD-A0DC-B787DA2DB0A1}","{DFD3BED7-93EC-4BCE-866C-6BAB41D28621}"
)

#Set-Location "HKCR:\"
$interface | % {
  # will not work .. why ? don't know
  # RI "HKCR\Interface\$_" -Recurse -Force -ea 0
}

"*** -- Components in Global [ & Could take 2-3 minutes & ]"
$Keys = reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components" 2>$null
$keys | % {
 $data = reg query $_ /t REG_SZ 2>$null
 if (($data -ne $null) -and (
   $data -match "\\Microsoft Office")) {
     reg delete $_ /f | Out-Null }}

"*** -- Components in CLSID [ & Could take 2-3 minutes & ]"
$Keys = reg query "HKLM\SOFTWARE\Classes\CLSID" 2>$null
$keys | % {
  $LocalServer32 = reg query "$_\LocalServer32" /ve /t REG_SZ 2>$null
  if (($LocalServer32 -ne $null) -and (
    $LocalServer32[2] -match "\\Microsoft Office")) {
      reg delete $_ /f | Out-Null }
  if ($LocalServer32 -eq $null) {
    $InprocServer32 = reg query "$_\InprocServer32" /ve /t REG_SZ 2>$null
    if (($InprocServer32 -ne $null) -and (
      $InprocServer32[2] -match "\\Microsoft Office")) {
        reg delete $_ /f | Out-Null }}}

<#
-- reg query "HKEY_CLASSES_ROOT\CLSID\
-- "HKEY_CLASSES_ROOT\CLSID\{C282417B-2662-44B8-8A94-3BFF61C50900}"

-- reg query "HKEY_CLASSES_ROOT\CLSID\{C282417B-2662-44B8-8A94-3BFF61C50900}\LocalServer32"
-- ERROR: The system was unable to find the specified registry key or value. [ACCESS DENIED ERROR]

$Keys = reg query "HKCR\CLSID" 2>$null
$keys | % {
  $LocalServer32 = reg query "$_\LocalServer32" /ve /t REG_SZ 2>$null
  if (($LocalServer32 -ne $null) -and (
    $LocalServer32[2] -match "\\Microsoft Office")) {
      reg delete $_ /f | Out-Null }
  if ($LocalServer32 -eq $null) {
    $InprocServer32 = reg query "$_\InprocServer32" /ve /t REG_SZ 2>$null
    if (($InprocServer32 -ne $null) -and (
      $InprocServer32[2] -match "\\Microsoft Office")) {
        reg delete $_ /f | Out-Null }}}
#>
}
Function FileWipe {

"*** -- remove the OfficeSvc service"
$service = $null
$service = Get-WmiObject Win32_Service -Filter "Name='OfficeSvc'" -ea 0
if ($service) { 
  try {
    $service.delete()|out-null}
  catch {} }

"*** -- remove the ClickToRunSvc service"
$service = $null
$service = Get-WmiObject Win32_Service -Filter "Name='ClickToRunSvc'" -ea 0
if ($service) { 
  try {
    $service.delete()|out-null}
  catch {} }

"*** -- delete C2R package files"
Set-Location "$($env:SystemDrive)\"

RI @(Join-Path $env:ProgramFiles "Microsoft Office\Office16") -Recurse -force -ea 0
RI @(Join-Path $env:ProgramData "Microsoft\office\FFPackageLocker") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramData "Microsoft\office\FFStatePBLocker") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramFiles "Microsoft Office\AppXManifest.xml") -force -ea 0 
RI @(Join-Path $env:ProgramFiles "Microsoft Office\FileSystemMetadata.xml") -force -ea 0 

RI @(Join-Path $env:ProgramData "Microsoft\ClickToRun") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramData "Microsoft\office\Heartbeat") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramData "Microsoft\office\FFPackageLocker") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramData "Microsoft\office\ClickToRunPackageLocker") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramFiles "Microsoft Office 15") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramFiles "Microsoft Office 16") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramFiles "Microsoft Office\root") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramFiles "Microsoft Office\Office16") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramFiles "Microsoft Office\Office15") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramFiles "Microsoft Office\PackageManifests") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramFiles "Microsoft Office\PackageSunrisePolicies") -Recurse -force -ea 0 
RI @(Join-Path $env:CommonProgramFiles "microsoft shared\ClickToRun") -Recurse -force -ea 0 

if ($env:ProgramFilesX86) {
  RI @(Join-Path $env:ProgramFilesX86 "Microsoft Office\AppXManifest.xml") -force -ea 0 
  RI @(Join-Path $env:ProgramFilesX86 "Microsoft Office\FileSystemMetadata.xml") -force -ea 0 
  RI @(Join-Path $env:ProgramFilesX86 "Microsoft Office\root") -Recurse -force -ea 0 
  RI @(Join-Path $env:ProgramFilesX86 "Microsoft Office\Office16") -Recurse -force -ea 0 
  RI @(Join-Path $env:ProgramFilesX86 "Microsoft Office\Office15") -Recurse -force -ea 0 
  RI @(Join-Path $env:ProgramFilesX86 "Microsoft Office\PackageManifests") -Recurse -force -ea 0 
  RI @(Join-Path $env:ProgramFilesX86 "Microsoft Office\PackageSunrisePolicies") -Recurse -force -ea 0 
}

RI @(Join-Path $env:userprofile "Microsoft Office") -Recurse -force -ea 0 
RI @(Join-Path $env:userprofile "Microsoft Office 15") -Recurse -force -ea 0 
RI @(Join-Path $env:userprofile "Microsoft Office 16") -Recurse -force -ea 0
}
Function RestoreExplorer {
$wmiInfo = gwmi -Query "Select * From Win32_Process Where Name='explorer.exe'"
if (-not $wmiInfo) {
  start "explorer"}
}
Function Uninstall {

"*** -- remove the published component registration for C2R packages"
$Location = (
  "SOFTWARE\Microsoft\Office\ClickToRun",
  "SOFTWARE\Microsoft\Office\16.0\ClickToRun",
  "SOFTWARE\Microsoft\Office\15.0\ClickToRun" )

Foreach ($Loc in $Location) {
  Set-Location "HKLM:\"
  Set-Location $Loc -ea 0
  if (@(Get-Location).Path -ne 'HKLM:\') {
    try {
      $sPkgFld  = $null; $sPkgGuid = $null;
      $sPkgFld  = GPV . -Name PackageFolder
      $sPkgGuid = GPV . -Name PackageGUID
      HandlePakage $sPkgFld $sPkgGuid
    }
    catch {
      $sPkgFld  = $null
      $sPkgGuid = $null
    }
}}

"*** -- delete potential blocking registry keys for msiexec based tasks"
Set-Location "HKLM:\"
RI "HKLM:SOFTWARE\Microsoft\Office\15.0\ClickToRun" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\16.0\ClickToRun" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\ClickToRun" -Force -ea 0 -Recurse

Set-Location "HKCU:\"
RI "HKCU:SOFTWARE\Microsoft\Office\15.0\ClickToRun" -Force -ea 0 -Recurse
RI "HKCU:SOFTWARE\Microsoft\Office\16.0\ClickToRun" -Force -ea 0 -Recurse
RI "HKCU:SOFTWARE\Microsoft\Office\ClickToRun" -Force -ea 0 -Recurse

"*** -- AppV keys"
$hDefKey_List = @(
  "HKCU", "HKLM" )
$sSubKeyName_List = @(
  "SOFTWARE\Microsoft\AppV\ISV",
  "SOFTWARE\Microsoft\AppVISV" )

foreach ($hDefKey in $hDefKey_List) {
  foreach ($sSubKeyName in $sSubKeyName_List) {
    Set-Location "$($hDefKey):\"
    Push-Location "$($hDefKey):$($sSubKeyName)" -ea 0
    if (@(Get-Location).Path -ne "$($hDefKey):\") {
      $arrNames = gi .
      if ($arrNames)  {
        $arrNames.Property | % { 
          $name = GPV . $_
          if ($name -and (
            $Name|IsC2R)) {
              RP . $_ -Force }}}}}}	
	
"*** -- msiexec based uninstall"
try {
  $omsi = Get-MsiProducts }
catch { 
 return }

 if (!($omsi)) { # ! same as -not
   return }
 
$sUninstallCmd = $null
$sMsiProp = "REBOOT=ReallySuppress NOREMOVESPAWN=True"

 $omsi | % {
  $ProductCode   = $_.ProductCode
  $InstallSource = $_.InstallSource

  if (($ProductCode -and ($ProductCode|CheckDelete)) -or (
    $InstallSource -and ($InstallSource|IsC2R))) {
        $sUninstallCmd = "/x $($ProductCode) $($sMsiProp) /q"
	    $proc = start msiexec.exe -Args $sUninstallCmd -Wait -WindowStyle Hidden -ea 0 -PassThru
        "*** -- msIexec $($sUninstallCmd) ,End with value: $($proc.ExitCode)"

 }}
 net stop msiserver *>$null
}
Function RegWipeTypeLib {
$sTLKey = 
"Software\Classes\TypeLib\"

$RegLibs = @(
"\0\Win32\","\0\Win64\","\9\Win32\","\9\Win64\")

$arrTypeLibs = @(
"{000204EF-0000-0000-C000-000000000046}","{000204EF-0000-0000-C000-000000000046}",
"{00020802-0000-0000-C000-000000000046}","{00020813-0000-0000-C000-000000000046}",
"{00020905-0000-0000-C000-000000000046}","{0002123C-0000-0000-C000-000000000046}",
"{00024517-0000-0000-C000-000000000046}","{0002E157-0000-0000-C000-000000000046}",
"{00062FFF-0000-0000-C000-000000000046}","{0006F062-0000-0000-C000-000000000046}",
"{0006F080-0000-0000-C000-000000000046}","{012F24C1-35B0-11D0-BF2D-0000E8D0D146}",
"{06CA6721-CB57-449E-8097-E65B9F543A1A}","{07B06096-5687-4D13-9E32-12B4259C9813}",
"{0A2F2FC4-26E1-457B-83EC-671B8FC4C86D}","{0AF7F3BE-8EA9-4816-889E-3ED22871FE05}",
"{0D452EE1-E08F-101A-852E-02608C4D0BB4}","{0EA692EE-BB50-4E3C-AEF0-356D91732725}",
"{1F8E79BA-9268-4889-ADF3-6D2AABB3C32C}","{2374F0B1-3220-4c71-B702-AF799F31ABB4}",
"{238AA1AC-786F-4C17-BAAB-253670B449B9}","{28DD2950-2D4A-42B5-ABBF-500AA42E7EC1}",
"{2A59CA0A-4F1B-44DF-A216-CB2C831E5870}","{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}",
"{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}","{2F7FC181-292B-11D2-A795-DFAA798E9148}",
"{3120BA9F-4FC8-4A4F-AE1E-02114F421D0A}","{31411197-A502-11D2-BBCA-00C04F8EC294}",
"{3B514091-5A69-4650-87A3-607C4004C8F2}","{47730B06-C23C-4FCA-8E86-42A6A1BC74F4}",
"{49C40DDF-1B04-4868-B3B5-E49F120E4BFA}","{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}",
"{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}","{4D95030A-A3A9-4C38-ACA8-D323A2267698}",
"{55A108B0-73BB-43db-8C03-1BEF4E3D2FE4}","{56D04F5D-964F-4DBF-8D23-B97989E53418}",
"{5B87B6F0-17C8-11D0-AD41-00A0C90DC8D9}","{66CDD37F-D313-4E81-8C31-4198F3E42C3C}",
"{6911FD67-B842-4E78-80C3-2D48597C2ED0}","{698BB59C-38F1-4CEF-92F9-7E3986E708D3}",
"{6DDCE504-C0DC-4398-8BDB-11545AAA33EF}","{6EFF1177-6974-4ED1-99AB-82905F931B87}",
"{73720002-33A0-11E4-9B9A-00155D152105}","{759EF423-2E8F-4200-ADF0-5B6177224BEE}",
"{76F6F3F5-9937-11D2-93BB-00105A994D2C}","{773F1B9A-35B9-4E95-83A0-A210F2DE3B37}",
"{7D868ACD-1A5D-4A47-A247-F39741353012}","{7E36E7CB-14FB-4F9E-B597-693CE6305ADC}",
"{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}","{8404DD0E-7A27-4399-B1D9-6492B7DD7F7F}",
"{8405D0DF-9FDD-4829-AEAD-8E2B0A18FEA4}","{859D8CF5-7ADE-4DAB-8F7D-AF171643B934}",
"{8E47F3A2-81A4-468E-A401-E1DEBBAE2D8D}","{91493440-5A91-11CF-8700-00AA0060263B}",
"{9A8120F2-2782-47DF-9B62-54F672075EA1}","{9B7C3E2E-25D5-4898-9D85-71CEA8B2B6DD}",
"{9B92EB61-CBC1-11D3-8C2D-00A0CC37B591}","{9D58B963-654A-4625-86AC-345062F53232}",
"{9DCE1FC0-58D3-471B-B069-653CE02DCE88}","{A4D51C5D-F8BF-46CC-92CC-2B34D2D89716}",
"{A717753E-C3A6-4650-9F60-472EB56A7061}","{AA53E405-C36D-478A-BBFF-F359DF962E6D}",
"{AAB9C2AA-6036-4AE1-A41C-A40AB7F39520}","{AB54A09E-1604-4438-9AC7-04BE3E6B0320}",
"{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}","{AC2DE821-36A2-11CF-8053-00AA006009FA}",
"{B30CDC65-4456-4FAA-93E3-F8A79E21891C}","{B8812619-BDB3-11D0-B19E-00A0C91E29D8}",
"{B9164592-D558-4EE7-8B41-F1C9F66D683A}","{B9AA1F11-F480-4054-A84E-B5D9277E40A8}",
"{BA35B84E-A623-471B-8B09-6D72DD072F25}","{BDEADE33-C265-11D0-BCED-00A0C90AB50F}",
"{BDEADEF0-C265-11D0-BCED-00A0C90AB50F}","{BDEADEF0-C265-11D0-BCED-00A0C90AB50F}",
"{C04E4E5E-89E6-43C0-92BD-D3F2C7FBA5C4}","{C3D19104-7A67-4EB0-B459-D5B2E734D430}",
"{C78F486B-F679-4af5-9166-4E4D7EA1CEFC}","{CA973FCA-E9C3-4B24-B864-7218FC1DA7BA}",
"{CBA4EBC4-0C04-468d-9F69-EF3FEED03236}","{CBBC4772-C9A4-4FE8-B34B-5EFBD68F8E27}",
"{CD2194AA-11BE-4EFD-97A6-74C39C6508FF}","{E0B12BAE-FC67-446C-AAE8-4FA1F00153A7}",
"{E985809A-84A6-4F35-86D6-9B52119AB9D7}","{ECD5307E-4419-43CF-8BDA-C9946AC375CF}",
"{EDCD5812-6A06-43C3-AFAC-46EF5D14E22C}","{EDCD5812-6A06-43C3-AFAC-46EF5D14E22C}",
"{EDCD5812-6A06-43C3-AFAC-46EF5D14E22C}","{EDDCFF16-3AEE-4883-BD91-0F3978640DFB}",
"{EE9CFA8C-F997-4221-BE2F-85A5F603218F}","{F2A7EE29-8BF6-4a6d-83F1-098E366C709C}",
"{F3685D71-1FC6-4CBD-B244-E60D8C89990B}")

    foreach ($tl in $arrTypeLibs) {
  
      Set-Location "HKLM:\"
      $sKey = "HKLM:" + $sTLKey + $tl

      Set-Location "HKLM:\"
      Push-Location $sKey -ea 0
      if (@(Get-Location).Path -eq 'HKLM:\') {
        continue
      }

      $children   = GCI .
      $fCanDelete = $false

      if (-not $children) {
        Set-Location "HKLM:\"
        Push-Location "HKLM:$($sTLKey)" -ea 0
        if (@(Get-Location).Path -ne 'HKLM:\') {
          RI $tl -Recurse -Force }
        continue
      }
  
      foreach ($K in $children) {
    
        $sTLVerKey = $sKey + "\" + $K.PSChildName
        $PSChildName = GCI $K.PSChildName -ea 0
        if ($PSChildName) {
          $fCanDelete = $true }
    
        Set-Location "HKLM:\"
        Push-Location $sKey -ea 0
        if (@(Get-Location).Path -eq 'HKLM:\') {
          continue }
    
        $RegLibs | % {
          Set-Location "HKLM:\"
          Push-Location "$($sTLVerKey)$($_)" -ea 0
          if (@(Get-Location).Path -ne 'HKLM:\') {
            try {
              $Default = gpv . -Name '(Default)' -ea 0 }
            catch {}
            if ($Default -and (
              [IO.FILE]::Exists($Default))) {
                $fCanDelete = $false }}}

        if ($fCanDelete) {
          Set-Location "HKLM:\"
          Push-Location $sKey -ea 0
          if (@(Get-Location).Path -ne 'HKLM:\') {
	      RI $K.PSChildName -Recurse -Force }}
      }
    }
}
Function CleanOSPP {
    $OfficeAppId  = '0ff1ce15-a989-479d-af46-f275c6370663'
    $SL_ID_PRODUCT_SKU = Get-SLIDList -eQueryIdType SL_ID_APPLICATION -pQueryId $OfficeAppId -eReturnIdType SL_ID_PRODUCT_SKU
    $SL_ID_ALL_LICENSE_FILES = Get-SLIDList -eQueryIdType SL_ID_APPLICATION -pQueryId $OfficeAppId -eReturnIdType SL_ID_ALL_LICENSE_FILES
    if ($SL_ID_PRODUCT_SKU) {
       SL-UninstallProductKey -skuList $SL_ID_PRODUCT_SKU
    }
    if ($SL_ID_ALL_LICENSE_FILES) {
       SL-UninstallLicense -LicenseFileIds $SL_ID_ALL_LICENSE_FILES
    }
}
Function ClearVNextLicCache {

$Licenses = Join-Path $ENV:localappdata "Microsoft\Office\Licenses"
    if (Test-Path $Licenses) {
      Set-Location "$($env:SystemDrive)\"
      RI $Licenses -Recurse -Force -ea 0 }
}
Function HandlePakage {
  param (
   [parameter(Mandatory=$True)]
   [string]$sPkgFldr,

   [parameter(Mandatory=$True)]
   [string]$sPkgGuid
  )

  $RootPath =
    Join-Path $sPkgFldr "\root"
  $IntegrationPath =
    Join-Path $sPkgFldr "\root\Integration"
  $Integrator =
    Join-Path $sPkgFldr "\root\Integration\Integrator.exe"
  $Integrator_ =
    "$env:ProgramData\Microsoft\ClickToRun\{$sPkgGuid}\integrator.exe"

  if (-not (
      Test-Path ($IntegrationPath ))) {
        return }
  
  Set-Location 'c:\'
  Push-Location $RootPath

  #Remove `Root`->`Integration\C2RManifest*.xml`
  if (@(Get-Location).Path -ne 'c:\') {
    RI .\Integration\ -Filter "C2RManifest*.xml" -Recurse -Force -ea 0
  }
  
  if ([IO.FILE]::Exists(
    $Integrator)) {
      $Args = "/U /Extension PackageRoot=""$($RootPath)"" PackageGUID=""$($sPkgGuid)"""
      $Proc = start $Integrator -arg $Args -Wait -WindowStyle Hidden -PassThru -ea 0
	  "*** -- Uninstall ID: $sPkgGuid with Full Args, returned with value:$($Proc.ExitCode)"
      $Args = "/U"
      $Proc = start $Integrator -arg $Args -Wait -WindowStyle Hidden -PassThru  -ea 0
	  "*** -- Uninstall ID: $sPkgGuid with Minimum Args, returned with value:$($Proc.ExitCode)" }

  if ([IO.FILE]::Exists(
    $Integrator_)) {
      $Args = "/U /Extension PackageRoot=""$($RootPath)"" PackageGUID=""$($sPkgGuid)"""
      $Proc = start $Integrator_ -arg $Args -Wait -WindowStyle Hidden -PassThru -ea 0
	  "*** -- Uninstall ID: $sPkgGuid with Full Args, returned with value:$($Proc.ExitCode)"
      $Args = "/U"
      $Proc = start $Integrator_ -arg $Args -Wait -WindowStyle Hidden -PassThru  -ea 0
	  "*** -- Uninstall ID: $sPkgGuid with Minimum Args, returned with value:$($Proc.ExitCode)" }
}
Function Office_Online_Install (
  [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
  [String] $Channel) {
Function Get-Lang {
  $oLang = @{}
  $oLang.Add(1033,"English")
  $oLang.Add(1078,"Afrikaans")
  $oLang.Add(1052,"Albanian")
  $oLang.Add(1118,"Amharic")
  $oLang.Add(1025,"Arabic")
  $oLang.Add(1067,"Armenian")
  $oLang.Add(1101,"Assamese")
  $oLang.Add(1068,"Azerbaijani Latin")
  $oLang.Add(2117,"Bangla Bangladesh")
  $oLang.Add(1093,"Bangla Bengali India")
  $oLang.Add(1069,"Basque Basque")
  $oLang.Add(1059,"Belarusian")
  $oLang.Add(5146,"Bosnian")
  $oLang.Add(1026,"Bulgarian")
  $oLang.Add(2051,"Catalan Valencia")
  $oLang.Add(1027,"Catalan")
  $oLang.Add(2052,"Chinese Simplified")
  $oLang.Add(1028,"Chinese Traditional")
  $oLang.Add(1050,"Croatian")
  $oLang.Add(1029,"Czech")
  $oLang.Add(1030,"Danish")
  $oLang.Add(1164,"Dari")
  $oLang.Add(1043,"Dutch")
  $oLang.Add(2057,"English UK")
  $oLang.Add(1061,"Estonian")
  $oLang.Add(1124,"Filipino")
  $oLang.Add(1035,"Finnish")
  $oLang.Add(3084,"French Canada")
  $oLang.Add(1036,"French")
  $oLang.Add(1110,"Galician")
  $oLang.Add(1079,"Georgian")
  $oLang.Add(1031,"German")
  $oLang.Add(1032,"Greek")
  $oLang.Add(1095,"Gujarati")
  $oLang.Add(1128,"Hausa Nigeria")
  $oLang.Add(1037,"Hebrew")
  $oLang.Add(1081,"Hindi")
  $oLang.Add(1038,"Hungarian")
  $oLang.Add(1039,"Icelandic")
  $oLang.Add(1136,"Igbo")
  $oLang.Add(1057,"Indonesian")
  $oLang.Add(2108,"Irish")
  $oLang.Add(1040,"Italian")
  $oLang.Add(1041,"Japanese")
  $oLang.Add(1099,"Kannada")
  $oLang.Add(1087,"Kazakh")
  $oLang.Add(1107,"Khmer")
  $oLang.Add(1089,"KiSwahili")
  $oLang.Add(1159,"Kinyarwanda")
  $oLang.Add(1111,"Konkani")
  $oLang.Add(1042,"Korean")
  $oLang.Add(1088,"Kyrgyz")
  $oLang.Add(1062,"Latvian")
  $oLang.Add(1063,"Lithuanian")
  $oLang.Add(1134,"Luxembourgish")
  $oLang.Add(1071,"Macedonian")
  $oLang.Add(1086,"Malay Latin")
  $oLang.Add(1100,"Malayalam")
  $oLang.Add(1082,"Maltese")
  $oLang.Add(1153,"Maori")
  $oLang.Add(1102,"Marathi")
  $oLang.Add(1104,"Mongolian")
  $oLang.Add(1121,"Nepali")
  $oLang.Add(2068,"Norwedian Nynorsk")
  $oLang.Add(1044,"Norwegian Bokmal")
  $oLang.Add(1096,"Odia")
  $oLang.Add(1123,"Pashto")
  $oLang.Add(1065,"Persian")
  $oLang.Add(1045,"Polish")
  $oLang.Add(1046,"Portuguese Brazilian")
  $oLang.Add(2070,"Portuguese Portugal")
  $oLang.Add(1094,"Punjabi Gurmukhi")
  $oLang.Add(3179,"Quechua")
  $oLang.Add(1048,"Romanian")
  $oLang.Add(1047,"Romansh")
  $oLang.Add(1049,"Russian")
  $oLang.Add(1074,"Setswana")
  $oLang.Add(1169,"Scottish Gaelic")
  $oLang.Add(7194,"Serbian Bosnia")
  $oLang.Add(10266,"Serbian Serbia")
  $oLang.Add(9242,"Serbian")
  $oLang.Add(1132,"Sesotho sa Leboa")
  $oLang.Add(2137,"Sindhi Arabic")
  $oLang.Add(1115,"Sinhala")
  $oLang.Add(1051,"Slovak")
  $oLang.Add(1060,"Slovenian")
  $oLang.Add(3082,"Spanish")
  $oLang.Add(2058,"Spanish Mexico")
  $oLang.Add(1053,"Swedish")
  $oLang.Add(1097,"Tamil")
  $oLang.Add(1092,"Tatar Cyrillic")
  $oLang.Add(1098,"Telugu")
  $oLang.Add(1054,"Thai")
  $oLang.Add(1055,"Turkish")
  $oLang.Add(1090,"Turkmen")
  $oLang.Add(1058,"Ukrainian")
  $oLang.Add(1056,"Urdu")
  $oLang.Add(1152,"Uyghur")
  $oLang.Add(1091,"Uzbek")
  $oLang.Add(1066,"Vietnamese")
  $oLang.Add(1106,"Welsh")
  $oLang.Add(1160,"Wolof")
  $oLang.Add(1130,"Yoruba")
  $oLang.Add(1076,"isiXhosa")
  $oLang.Add(1077,"isiZulu")
  return $oLang
}
Function Get-Culture {
  $oLang = @{}
  $oLang.Add(1033,"en-us")
  $oLang.Add(1078,"af-za")
  $oLang.Add(1052,"sq-al")
  $oLang.Add(1118,"am-et")
  $oLang.Add(1025,"ar-sa")
  $oLang.Add(1067,"hy-am")
  $oLang.Add(1101,"as-in")
  $oLang.Add(1068,"az-latn-az")
  $oLang.Add(2117,"bn-bd")
  $oLang.Add(1093,"bn-in")
  $oLang.Add(1069,"eu-es")
  $oLang.Add(1059,"be-by")
  $oLang.Add(5146,"bs-latn-ba")
  $oLang.Add(1026,"bg-bg")
  $oLang.Add(2051,"ca-es-valencia")
  $oLang.Add(1027,"ca-es")
  $oLang.Add(2052,"zh-cn")
  $oLang.Add(1028,"zh-tw")
  $oLang.Add(1050,"hr-hr")
  $oLang.Add(1029,"cs-cz")
  $oLang.Add(1030,"da-dk")
  $oLang.Add(1164,"prs-af")
  $oLang.Add(1043,"nl-nl")
  $oLang.Add(2057,"en-GB")
  $oLang.Add(1061,"et-ee")
  $oLang.Add(1124,"fil-ph")
  $oLang.Add(1035,"fi-fi")
  $oLang.Add(3084,"fr-CA")
  $oLang.Add(1036,"fr-fr")
  $oLang.Add(1110,"gl-es")
  $oLang.Add(1079,"ka-ge")
  $oLang.Add(1031,"de-de")
  $oLang.Add(1032,"el-gr")
  $oLang.Add(1095,"gu-in")
  $oLang.Add(1128,"ha-Latn-NG")
  $oLang.Add(1037,"he-il")
  $oLang.Add(1081,"hi-in")
  $oLang.Add(1038,"hu-hu")
  $oLang.Add(1039,"is-is")
  $oLang.Add(1136,"ig-NG")
  $oLang.Add(1057,"id-id")
  $oLang.Add(2108,"ga-ie")
  $oLang.Add(1040,"it-it")
  $oLang.Add(1041,"ja-jp")
  $oLang.Add(1099,"kn-in")
  $oLang.Add(1087,"kk-kz")
  $oLang.Add(1107,"km-kh")
  $oLang.Add(1089,"sw-ke")
  $oLang.Add(1159,"rw-RW")
  $oLang.Add(1111,"kok-in")
  $oLang.Add(1042,"ko-kr")
  $oLang.Add(1088,"ky-kg")
  $oLang.Add(1062,"lv-lv")
  $oLang.Add(1063,"lt-lt")
  $oLang.Add(1134,"lb-lu")
  $oLang.Add(1071,"mk-mk")
  $oLang.Add(1086,"ms-my")
  $oLang.Add(1100,"ml-in")
  $oLang.Add(1082,"mt-mt")
  $oLang.Add(1153,"mi-nz")
  $oLang.Add(1102,"mr-in")
  $oLang.Add(1104,"mn-mn")
  $oLang.Add(1121,"ne-np")
  $oLang.Add(2068,"nn-no")
  $oLang.Add(1044,"nb-no")
  $oLang.Add(1096,"or-in")
  $oLang.Add(1123,"ps-AF")
  $oLang.Add(1065,"fa-ir")
  $oLang.Add(1045,"pl-pl")
  $oLang.Add(1046,"pt-br")
  $oLang.Add(2070,"pt-pt")
  $oLang.Add(1094,"pa-in")
  $oLang.Add(3179,"quz-pe")
  $oLang.Add(1048,"ro-ro")
  $oLang.Add(1047,"rm-CH")
  $oLang.Add(1049,"ru-ru")
  $oLang.Add(1074,"tn-ZA")
  $oLang.Add(1169,"gd-gb")
  $oLang.Add(7194,"sr-cyrl-ba")
  $oLang.Add(10266,"sr-cyrl-rs")
  $oLang.Add(9242,"sr-latn-rs")
  $oLang.Add(1132,"nso-ZA")
  $oLang.Add(2137,"sd-arab-pk")
  $oLang.Add(1115,"si-lk")
  $oLang.Add(1051,"sk-sk")
  $oLang.Add(1060,"sl-si")
  $oLang.Add(3082,"es-es")
  $oLang.Add(2058,"es-MX")
  $oLang.Add(1053,"sv-se")
  $oLang.Add(1097,"ta-in")
  $oLang.Add(1092,"tt-ru")
  $oLang.Add(1098,"te-in")
  $oLang.Add(1054,"th-th")
  $oLang.Add(1055,"tr-tr")
  $oLang.Add(1090,"tk-tm")
  $oLang.Add(1058,"uk-ua")
  $oLang.Add(1056,"ur-pk")
  $oLang.Add(1152,"ug-cn")
  $oLang.Add(1091,"uz-latn-uz")
  $oLang.Add(1066,"vi-vn")
  $oLang.Add(1106,"cy-gb")
  $oLang.Add(1160,"wo-SN")
  $oLang.Add(1130,"yo-NG")
  $oLang.Add(1076,"xh-ZA")
  $oLang.Add(1077,"zu-ZA")
  return $oLang
}
Function Get-Channels {
  $oProd = @{}
  $oProd.Add("BetaChannel","5440fd1f-7ecb-4221-8110-145efaa6372f")
  $oProd.Add("Current","492350f6-3a01-4f97-b9c0-c7c6ddf67d60")
  $oProd.Add("CurrentPreview","64256afe-f5d9-4f86-8936-8840a6a4f5be")
  $oProd.Add("DogfoodCC","f3260cf1-a92c-4c75-b02e-d64c0a86a968")
  $oProd.Add("DogfoodDCEXT","c4a7726f-06ea-48e2-a13a-9d78849eb706")
  $oProd.Add("DogfoodDevMain","ea4a4090-de26-49d7-93c1-91bff9e53fc3")
  $oProd.Add("DogfoodFRDC","834504cc-dc55-4c6d-9e71-e024d0253f6d")
  $oProd.Add("InsidersLTSC","2e148de9-61c8-4051-b103-4af54baffbb4")
  $oProd.Add("InsidersLTSC2021","12f4f6ad-fdea-4d2a-a90f-17496cc19a48")
  $oProd.Add("InsidersLTSC2024","20481F5C-C268-4624-936C-52EB39DDBD97")
  $oProd.Add("InsidersMEC","0002c1ba-b76b-4af9-b1ee-ae2ad587371f")   
  $oProd.Add("MicrosoftCC","5462eee5-1e97-495b-9370-853cd873bb07")
  $oProd.Add("MicrosoftDC","f4f024c8-d611-4748-a7e0-02b6e754c0fe")
  $oProd.Add("MicrosoftDevMain","b61285dd-d9f7-41f2-9757-8f61cba4e9c8")
  $oProd.Add("MicrosoftFRDC","9a3b7ff2-58ed-40fd-add5-1e5158059d1c")
  $oProd.Add("MicrosoftLTSC","1d2d2ea6-1680-4c56-ac58-a441c8c24ff9")
  $oProd.Add("MicrosoftLTSC2021","86752282-5841-4120-ac80-db03ae6b5fdb")
  $oProd.Add("MicrosoftLTSC2024","C02D8FE6-5242-4DA8-972F-82EE55E00671")
  $oProd.Add("MonthlyEnterprise","55336b82-a18d-4dd6-b5f6-9e5095c314a6")
  $oProd.Add("PerpetualVL2019","f2e724c1-748f-4b47-8fb8-8e0d210e9208")
  $oProd.Add("PerpetualVL2021","5030841d-c919-4594-8d2d-84ae4f96e58e")
  $oProd.Add("PerpetualVL2024","7983BAC0-E531-40CF-BE00-FD24FE66619C")
  $oProd.Add("SemiAnnual","7ffbc6bf-bc32-4f92-8982-f9dd17fd3114")
  $oProd.Add("SemiAnnualPreview","b8f9b850-328d-4355-9145-c59439a0c4cf")
  return $oProd
}

  $IsX32=$Null
  $IsX64=$Null
  $file = $Null

  $IgnoreCase = [Text.RegularExpressions.RegexOptions]::IgnoreCase

  $oProductsId = Get-Channels
  if ($Channel -and ($oProductsId[$Channel] -eq $null)) {
    throw "ERROR: BAD CHANNEL"
  }

  if ($Channel) {
    $sChannel = $oProductsId.GetEnumerator()|? {$_.key -eq $Channel}
  }

  if (-not($sChannel)){
    $sChannel = $oProductsId | OGV -Title "Select Channel" -OutputMode Single
    if (-not ($sChannel)){
      return;
  }}

  # find FFNRoot value
  $FFNRoot = $sChannel.value
  $sUrl="http://officecdn.microsoft.com/pr/$FFNRoot"

  $Build = Get-Office-Apps|?{($_.Channel -eq $sChannel.Key) -and ($_.System -eq '10.0')}|select Build
  if (-not ($Build)) {
    throw "ERROR: FAIL TO GET BUILD VERSION"
  }
  $oVer = $Build.Build

  ri @(Join-Path $env:TEMP VersionDescriptor.xml) -Force -ea 0
  try {
        Switch ([intptr]::Size) {
            4 { 
                $IsX32 = $true
                $IsX64 = $Null
                $file = "$ENV:TEMP\v32.cab"
                ri $file -Force -ea 0
            
                # Attempt to download the v32.cab file
                try {
                    Write-Warning "Cab File, $sUrl/Office/Data/v32.cab"
                    iwr -Uri "$sUrl/Office/Data/v32.cab" -OutFile $file -ErrorAction Stop
                } catch {
                    Write-Warning "ERROR: FAIL DOWNLOAD CAB FILE for v32.cab ($_)"
                    return
                }

                # Attempt to extract the CAB file
                try {
                    Expand $file -f:VersionDescriptor.xml $env:TEMP *>$Null
                } catch {
                    Write-Warning "ERROR: FAIL EXTRACT XML FILE for v32.cab ($_)"
                    return
                }
            }

            8 {
                $IsX32 = $Null
                $IsX64 = $true
                $file = "$ENV:TEMP\v64.cab"
                ri $file -Force -ea 0
            
                # Attempt to download the v64.cab file
                try {
                    Write-Warning "Cab File, $sUrl/Office/Data/v64.cab"
                    iwr -Uri "$sUrl/Office/Data/v64.cab" -OutFile $file -ErrorAction Stop
                } catch {
                    Write-Warning "ERROR: FAIL DOWNLOAD CAB FILE for v64.cab ($_)"
                    return
                }

                # Attempt to extract the CAB file
                try {
                    Expand $file -f:VersionDescriptor.xml $env:TEMP *>$Null
                } catch {
                    Write-Warning "ERROR: FAIL EXTRACT XML FILE for v64.cab ($_)"
                    return
                }
            }
        }
    }
    catch {
        Write-Warning "Script failed during processing. Exiting."
        return
    }

  if (!(Test-path(
    @(Join-Path $env:TEMP VersionDescriptor.xml)))) {
      throw "ERROR: FAIL EXTRACT XML FILE"
  }

  $oXml = Get-Content @(Join-Path $env:TEMP VersionDescriptor.xml) -ea 0
  if (!$oXml) {
    throw "ERROR: FAIL READ XML FILE"
  }

  $rPat = '^(.*)(ProductReleaseId Name=)(.*)>$'
  $oApps = $oXml|?{[REGEX]::IsMatch($_,$rPat,$IgnoreCase)}|%{$_.SubString(28,$_.Length-28-2)}|sort|OGV -title "Found Office Apps" -OutputMode Multiple
  if (-not $oApps) {
    return;
  }

  $LangList = Get-Lang
 
  $mLang = $LangList |  OGV -Title "Select the Main Language" -OutputMode Single
  if (-not $mLang) {
    return;
  }
  $aLang = $LangList |  OGV -Title "Select Additional Language[s]" -PassThru

  $culture = ''
  $oCul = Get-Culture
  
  $mLang | % {
    $culture += $oCul[[INT]$_.Key]
  }

  if ($aLang) {
    $aLang | % {
      $culture += '_' + $oCul[[INT]$_.Key]
  }}
  
  # Start set values
  $type = "CDN"
  $bUrl = $sUrl
  $misc = "flt.useoutlookshareaddon=unknown flt.useofficehelperaddon=unknown"

  $sCulture = $culture
  $mCulture = $oCul[[INT]$mLang[0].Name]

  $AppList = ''
  $oApps | % {$AppList += "$($_).16_$($sCulture)_x-none|"}
  $sAppList = $AppList.TrimEnd('|')
  
  $services = @("WSearch", "ClickToRunSvc")

  foreach ($svcName in $services) {
    try {
        $svc = Get-Service -Name $svcName -ErrorAction Stop
        if ($svc.Status -eq 'Running') {
            $svc.Stop() | Out-Null
        }
    } catch {
        # Silently continue; service may not exist or already be stopped
    }
  }
  
  if ($IsX32) {
	$vSys = "x86"
    $c2r = "$ENV:ProgramFiles(x86)\Common Files\Microsoft Shared\ClickToRun"
    MD $c2r -ea 0 | Out-Null
    $file = "$ENV:TEMP\i320.cab"
    ri $file -Force -ea 0
    iwr -Uri "$sUrl/Office/Data/$oVer/i320.cab" -OutFile $file -ea 0
    if (-not(Test-Path($file))){throw "ERROR: FAIL DOWNLOAD CAB FILE"}
    Expand $file -f:* $c2r *>$Null
    Push-Location $c2r
  }

  if ($IsX64) {
	$vSys = "x64"
    $c2r = "$ENV:ProgramFiles\Common Files\Microsoft Shared\ClickToRun"
    MD $c2r -ea 0 | Out-Null
    $file = "$ENV:TEMP\i640.cab"
    ri $file -Force -ea 0
    iwr -Uri "$sUrl/Office/Data/$oVer/i640.cab" -OutFile $file -ea 0
    if (-not(Test-Path($file))){throw "ERROR: FAIL DOWNLOAD CAB FILE"}
    Expand $file -f:* $c2r *>$Null
    Push-Location $c2r
  }
  
  $args = @(
    "platform=$vSys"
    "culture=$mCulture"
    "productstoadd=$sAppList"
    "cdnbaseurl.16=$sUrl"
    "baseurl.16=$bUrl"
    "version.16=$oVer"
    "mediatype.16=$type"
    "sourcetype.16=$type"
    "updatesenabled.16=True"
    "acceptalleulas.16=True"
    "displaylevel=True"
    "bitnessmigration=False"
    "deliverymechanism=$FFNRoot"
    "$misc"
  )
  $OfficeClickToRun = Join-Path $c2r OfficeClickToRun.exe
  if (-not (Test-Path -Path $OfficeClickToRun)) {
    Write-Warning "Missing file, $OfficeClickToRun"
    return
  }
  $process = Start-Process -FilePath $OfficeClickToRun -ArgumentList $args -NoNewWindow -PassThru
  $process.WaitForExit()
  if ($process.ExitCode -eq 0) {
     Write-Host "OfficeClickToRun.exe ran successfully."
  } else {
      Write-Host "There was an error. Exit code: $($process.ExitCode)"
  }
  return
}
function Uninstall-Licenses {
    
    Manage-SLHandle -Release | Out-null
    $WMI_QUERY = @()
    $WMI_QUERY = Get-SLIDList -eQueryIdType SL_ID_PRODUCT_SKU -eReturnIdType SL_ID_PRODUCT_SKU
    $WMI_SQL = foreach ($iid in $WMI_QUERY) {
        Get-LicenseInfo -ActConfigId $iid
    }

    # Filter the results to ensure only items with EditionId
    $filteredResults = $WMI_SQL | Where-Object { $_.EditionId } | Select-Object ActConfigId, EditionId, ProductDescription

    # Show the GridView where user can select multiple items, with all columns visible
    $selectedItems = $filteredResults | Out-GridView -Title "Select Prodouct SKU To Remove" -PassThru

    # If any items are selected
    if ($selectedItems) {
        # Extract only ActConfigId GUIDs from the selected items
        $GUID_ARRAY = $selectedItems | Select-Object -ExpandProperty ActConfigId

        # Uninstall the products using the GUID array
        #SL-UninstallProductKey $skuList $GUID_ARRAY
        SL-UninstallLicense -ProductSKUs $GUID_ARRAY
    } else {
        Write-Host "No items selected."
    }
    Manage-SLHandle -Release | Out-null
}
Function Reset-Store {
    Write-Host
    Write-Host "##### Running :: slmgr.vbs /rilc"
    Stop-Service -Name sppsvc -Force
    $networkServicePath = "$env:SystemDrive\Windows\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform"
    if (Test-Path "$networkServicePath\tokens.bar") {
        Remove-Item "$networkServicePath\tokens.bar" -Force
    }
    if (Test-Path "$networkServicePath\tokens.dat") {
        Rename-Item "$networkServicePath\tokens.dat" -NewName "tokens.bar"
    }
    $storePath = "$env:SystemDrive\Windows\System32\spp\store"
    if (Test-Path "$storePath\tokens.bar") {
        Remove-Item "$storePath\tokens.bar" -Force
    }
    if (Test-Path "$storePath\tokens.dat") {
        Rename-Item "$storePath\tokens.dat" -NewName "tokens.bar"
    }
    $storePath2 = "$env:SystemDrive\Windows\System32\spp\store\2.0"
    if (Test-Path "$storePath2\tokens.bar") {
        Remove-Item "$storePath2\tokens.bar" -Force
    }
    if (Test-Path "$storePath2\tokens.dat") {
        Rename-Item "$storePath2\tokens.dat" -NewName "tokens.bar"
    }
    Start-Service -Name sppsvc
  
    $pathList = (
    (Join-Path $ENV:SystemRoot "system32\oem"),
    (Join-Path $ENV:SystemRoot "system32\spp\tokens"))

    $Selection = @(
        foreach ($loc in $pathList) {
            if (Test-Path $loc) {
                # Find all .xrm-ms files in the location and add them to $LicenseFiles
                Get-ChildItem -Path $loc -Filter *.xrm-ms -Recurse -Force | ForEach-Object {
                    $_.FullName } } })
    
    Manage-SLHandle -Release | Out-null
    SL-InstallLicense -LicenseInput $Selection
}
Function Office-License-Installer {
    Manage-SLHandle -Release | Out-null
    $targetPath = "$env:ProgramFiles\Microsoft Office\root\Licenses16"
    Set-Location $targetPath -ErrorAction SilentlyContinue
    if ($PWD.Path -ieq $targetPath) {
    } else {
        Write-Warning "Failed to change to the target folder."
        return
    }

    $LicensingService = gwmi SoftwareLicensingService -ErrorAction Stop
    if (-not $LicensingService) {
        return }

    $file_list = dir * -Name
    $loc = (Get-Location).Path
    if ($AutoMode -and $LicensePattern -and (-not [string]::IsNullOrWhiteSpace($LicensePattern))) {
        $Selection = dir "*$LicensePattern*" -Name
    } else {
        $Selection = $file_list | ogv -Title "License installer - Helper" -OutputMode Multiple
    }
    if ($Selection) {
    
        $AllLicenseFiles = @(
            $Selection + ($file_list | Where-Object { $_ -like "pkeyconfig*" -or $_ -like "Client*" })
        ) | ForEach-Object { Join-Path $loc $_ }

        #$AllLicenseFiles | ForEach-Object { Write-Host "Install License: $_" }
        SL-installLicense -LicenseInput $AllLicenseFiles
    }
     Manage-SLHandle -Release | Out-null
    return
}
Function OffScrubc2r {
# ---------------------- #
# Begin of main function #
# ---------------------- #

"*** $(Get-Date -Format hh:mm:ss): Load HKCR Hive"
if ($null -eq (Get-PSDrive HKCR -ea 0)) {
    New-PSDrive HKCR Registry HKEY_CLASSES_ROOT -ErrorAction Stop | Out-Null }

"*** $(Get-Date -Format hh:mm:ss): Clean OSPP"
CleanOSPP

"*** $(Get-Date -Format hh:mm:ss): Clean vNext Licenses"
ClearVNextLicCache

"*** $(Get-Date -Format hh:mm:ss): End running processes"
ClearShellIntegrationReg
CloseOfficeApps

"*** $(Get-Date -Format hh:mm:ss): Clean Scheduler tasks"
DelSchtasks

"*** $(Get-Date -Format hh:mm:ss): Clean Office shortcuts"
CleanShortcuts -sFolder "$env:AllusersProfile"
CleanShortcuts -sFolder "$env:SystemDrive\Users"

"*** $(Get-Date -Format hh:mm:ss): Remove Office C2R / O365"
Uninstall

"*** $(Get-Date -Format hh:mm:ss): call odt based uninstall"
UninstallOfficeC2R

"*** $(Get-Date -Format hh:mm:ss): CleanUp"
FileWipe
RegWipe

"*** $(Get-Date -Format hh:mm:ss): Ensure Explorer runs"
RestoreExplorer

"*** $(Get-Date -Format hh:mm:ss): Un-Load HKCR Hive"
Set-Location "HKLM:\"
Remove-PSDrive -Name HKCR -ea 0 | Out-Null

write-host "Begin: $($Start_Time), End: $(Get-Date -Format hh:mm:ss)"
Write-Host
timeout 3 *>$null
return
}
function Test-WMIHealth {
    Write-Host "`nTesting WMI health..." -ForegroundColor Cyan
    $WmiFailure = $false

    try {
        $null = Get-Disk -ErrorAction Stop
        $null = Get-Partition -ErrorAction Stop
        $arch = (Get-CimInstance Win32_Processor).AddressWidth
        if ($arch -match "64|32") {
            # Nothing here
        } else {
            $WmiFailure = $true
        }
    } catch {
        $WmiFailure = $true
    }

    if ($WmiFailure) {
        Write-Host "`n*** WMI STATUS = FAIL ***" -ForegroundColor Red
    } else {
        Write-Host "`n*** WMI STATUS = OK ***" -ForegroundColor Green
    }
}
function Invoke-WMIRepair {
    param (
        [ValidateSet("Soft", "Hard")]
        [string]$Mode = "Soft"
    )

    if ($Mode -eq "Soft") {
        Write-Host "`n[Soft Repair] Starting..." -ForegroundColor Yellow
        winmgmt /verifyrepository
        winmgmt /salvagerepository

        # Restart WMI service after soft repair
        Restart-Service -Name winmgmt -Force
    } else {
        Write-Host "`n[Hard Repair] Starting..." -ForegroundColor Red

        # Stop and disable winmgmt service before repair
        Stop-Service -Name winmgmt -Force -ea 0
        Set-Service -Name winmgmt -StartupType Disabled

        $basePaths = @(
            "$env:windir\System32",
            "$env:windir\SysWOW64"
        )

        foreach ($base in $basePaths) {
            $wbem = Join-Path $base "wbem"
            if (Test-Path $wbem) {
                Push-Location $wbem

                winmgmt /resetrepository
                winmgmt /resyncperf

                if (Test-Path "$wbem\Repos_bakup") {
                    Remove-Item "$wbem\Repos_bakup" -Recurse -Force
                }
                if (Test-Path "$wbem\Repository") {
                    Rename-Item "$wbem\Repository" "Repos_bakup"
                }

                # Re-register key DLLs
                $dlls = @("scecli.dll", "userenv.dll")
                foreach ($dll in $dlls) {
                    $dllPath = Join-Path $base $dll
                    if (Test-Path $dllPath) {
                        Start-Process regsvr32 -ArgumentList "/s", $dllPath -Wait -NoNewWindow
                    }
                }

                # Register all DLLs in wbem folder
                Get-ChildItem -Filter *.dll | ForEach-Object {
                    Start-Process regsvr32 -ArgumentList "/s", $_.FullName -Wait -NoNewWindow
                }

                # Recompile MOFs and MFLs in wbem root folder
                Get-ChildItem -Filter *.mof | ForEach-Object { mofcomp $_.FullName | Out-Null }
                Get-ChildItem -Filter *.mfl | ForEach-Object { mofcomp $_.FullName | Out-Null }

                # === NEW: Recompile MOFs recursively in all wbem subfolders ===
                Write-Host "[INFO] Recursively recompiling MOFs in wbem subfolders..." -ForegroundColor Cyan
                Get-ChildItem -Recurse -Filter *.mof -ea 0 | ForEach-Object {
                    try {
                        mofcomp $_.FullName | Out-Null
                        Write-Host "Compiled: $($_.FullName)"
                    } catch {
                        Write-Warning "Failed to compile: $($_.FullName)"
                    }
                }

                # === NEW: Explicit Storage MOFs and DLLs ===
                Write-Host "[INFO] Recompiling storage-related MOFs and registering DLLs..." -ForegroundColor Cyan

                $criticalMofs = @(
                    # Storage-related
                    "$env:windir\System32\wbem\storage.mof",
                    "$env:windir\System32\wbem\disk.mof",
                    "$env:windir\System32\wbem\volume.mof",

                    # Core system management
                    "$env:windir\System32\wbem\cimwin32.mof",
                    "$env:windir\System32\wbem\netevent.mof",
                    "$env:windir\System32\wbem\wmipicmp.mof",
                    "$env:windir\System32\wbem\msiprov.mof",
                    "$env:windir\System32\wbem\wmi.mof",
                    "$env:windir\System32\wbem\eventlog.mof",
                    "$env:windir\System32\wbem\perf.mof",
                    "$env:windir\System32\wbem\perfproc.mof",
                    "$env:windir\System32\wbem\perfdisk.mof",
                    "$env:windir\System32\wbem\perfnet.mof",

                    # Networking
                    "$env:windir\System32\wbem\netbios.mof",
                    "$env:windir\System32\wbem\network.mof",

                    # Other potentially important MOFs
                    "$env:windir\System32\wbem\swprv.mof",
                    "$env:windir\System32\wbem\vsprov.mof"
                )

                foreach ($mof in $criticalMofs) {
                    if (Test-Path $mof) {
                        mofcomp $mof | Out-Null
                        Write-Host "Compiled storage MOF: $mof"
                    } else {
                        Write-Warning "Storage MOF not found: $mof"
                    }
                }

                $storageDlls = @(
                    "$env:windir\System32\storprov.dll",
                    "$env:windir\System32\vmstorfl.dll"
                )

                foreach ($dll in $storageDlls) {
                    if (Test-Path $dll) {
                        Start-Process regsvr32 -ArgumentList "/s", $dll -Wait -NoNewWindow
                        Write-Host "Registered DLL: $dll"
                    } else {
                        Write-Warning "Storage DLL not found: $dll"
                    }
                }

                # Restart winmgmt service and set to Automatic
                Set-Service -Name winmgmt -StartupType Automatic
                Start-Service -Name winmgmt

                # Re-register wmiprvse.exe
                Start-Process "wmiprvse.exe" -ArgumentList "/regserver" -Wait -NoNewWindow

                Pop-Location
            }
        }
    }

    Write-Host "`n[$Mode Repair] Completed." -ForegroundColor Green
}
Function WMI_Reset_Main {

    Clear-Host
    Write-Host "* Make sure to run as administrator"
    Write-Host "* Please disable any antivirus temporarily`n"
    Pause

    Test-WMIHealth
    Write-Host
    Write-Host "`nChoose repair mode:"
    Write-Host "[1] Soft Repair (Safe, Recommended First)"
    Write-Host "[2] Hard Repair (Full Reset, use only if Soft Repair fails)"
    $modeInput = Read-Host "Enter 1 or 2"

    switch ($modeInput) {
        '1' { Invoke-WMIRepair -Mode Soft }
        '2' { Invoke-WMIRepair -Mode Hard }
        default { Write-Host "Invalid selection. Exiting." -ForegroundColor Red; Pause; return }
    }

    # Ask for reboot after repair
    Write-Host "`nA system restart is strongly recommended to complete the WMI repair." -ForegroundColor Cyan
    $reboot = Read-Host "Do you want to restart now? (Y/N)"

    if ($reboot -match '^[Yy]$') {
        Write-Host "`nRestarting system..." -ForegroundColor Yellow
        Restart-Computer -Force
    } else {
        Write-Host "`nPlease remember to restart the system manually later." -ForegroundColor Red
        return
    }
}

if ($AutoMode) {
    $actionMap = [ordered]@{
        "RunWmiRepair"            = $RunWmiRepair;
        "RunTokenStoreReset"      = $RunTokenStoreReset;
        "RunUninstallLicenses"    = $RunUninstallLicenses;
        "RunScrubOfficeC2R"       = $RunScrubOfficeC2R;
        "RunOfficeLicenseInstaller" = $RunOfficeLicenseInstaller;
        "RunOfficeOnlineInstallation" = $RunOfficeOnlineInstallation
    }

    foreach ($action in $actionMap.GetEnumerator()) {
        if ($action.Value) {
            switch ($action.Name) {
                "RunWmiRepair" {
                    try {
                        WMI_Reset_Main
                    }
                    catch {
                        Write-Warning "$($action.Name) Mode Fail"
                    }
                    break
                }
                "RunTokenStoreReset" {
                    try {
                        Reset-Store
                    }
                    catch {
                        Write-Warning "$($action.Name) Mode Fail"
                    }
                    break
                }
                "RunUninstallLicenses" {
                    try {
                        Uninstall-Licenses
                    }
                    catch {
                        Write-Warning "$($action.Name) Mode Fail"
                    }
                    break
                }
                "RunScrubOfficeC2R" {
                    try {
                        OffScrubc2r
                    }
                    catch {
                        Write-Warning "$($action.Name) Mode Fail"
                    }
                    break
                }
                "RunOfficeLicenseInstaller" {
                    try {
                        Office-License-Installer
                    }
                    catch {
                        Write-Warning "$($action.Name) Mode Fail"
                    }
                    break
                }
                "RunOfficeOnlineInstallation" {
                    try {
                        Office_Online_Install
                    }
                    catch {
                        Write-Warning "$($action.Name) Mode Fail"
                    }
                    break
                }
                default {
                }
            }
        }
    }

    return
}
  
Clear-host
Write-Host
Write-Host "Troubleshoot Menu" -ForegroundColor Green
Write-Host
Write-Host "1. Wmi Repair / Reset" -ForegroundColor Green
write-host "2. Token Store Reset" -ForegroundColor Green
Write-Host "3. Uninstall Licenses" -ForegroundColor Green
write-host "4. Scrub Office C2R" -ForegroundColor Green
write-host "5. Office License Installer" -ForegroundColor Green
write-host "6. Office Online Installation" -ForegroundColor Green
write-host "7. Upgrade Windows edition" -ForegroundColor Green
Write-Host
$choice = Read-Host "Please choose an option:"
Write-Host

switch ($choice) {
    "1" {
        WMI_Reset_Main
    }
    "2" {
        Reset-Store
    }
    "3" {
        Uninstall-Licenses
    }
    "4" {
        OffScrubc2r
    }
    "5" {
        Office-License-Installer
    }
    "6" {
        Office_Online_Install
    }
    "7" {
        Get-EditionTargetsFromMatrix -UpgradeFrom $true
    }    
}
# --> End
}
Function Check-Status {

    Clear-host
    Write-Host
    Write-Host "Start Check ..."
    Write-Host

    $ohook_found = $false
    $paths = @(
        "$env:ProgramFiles\Microsoft Office\Office15\sppc*.dll",
        "$env:ProgramFiles\Microsoft Office\Office16\sppc*.dll",
        "$env:ProgramW6432\Microsoft Office\Office15\sppc*.dll",
        "$env:ProgramW6432\Microsoft Office\Office16\sppc*.dll",
        "$env:ProgramFiles(x86)\Microsoft Office\Office15\sppc*.dll",
        "$env:ProgramFiles(x86)\Microsoft Office\Office16\sppc*.dll"
    )

    foreach ($path in $paths) {
        if (Get-ChildItem -Path $path -Filter 'sppc*.dll' -Attributes ReparsePoint -ea 0) {
            $ohook_found = $true
            break
        }
    }

    # Also check the root\vfs paths
    $vfsPaths = @(
        "$env:ProgramFiles\Microsoft Office 15\root\vfs\System\sppc*.dll",
        "$env:ProgramFiles\Microsoft Office 15\root\vfs\SystemX86\sppc*.dll",
        "$env:ProgramFiles\Microsoft Office\root\vfs\System\sppc*.dll",
        "$env:ProgramFiles\Microsoft Office\root\vfs\SystemX86\sppc*.dll",
        "$env:ProgramW6432\Microsoft Office 15\root\vfs\System\sppc*.dll",
        "$env:ProgramW6432\Microsoft Office 15\root\vfs\SystemX86\sppc*.dll",
        "$env:ProgramW6432\Microsoft Office\root\vfs\System\sppc*.dll",
        "$env:ProgramW6432\Microsoft Office\root\vfs\SystemX86\sppc*.dll",
        "$env:ProgramFiles(x86)\Microsoft Office 15\root\vfs\System\sppc*.dll",
        "$env:ProgramFiles(x86)\Microsoft Office 15\root\vfs\SystemX86\sppc*.dll",
        "$env:ProgramFiles(x86)\Microsoft Office\root\vfs\System\sppc*.dll",
        "$env:ProgramFiles(x86)\Microsoft Office\root\vfs\SystemX86\sppc*.dll"
    )

    foreach ($path in $vfsPaths) {
        if (Get-ChildItem -Path $path -Filter 'sppc*.dll' -Attributes ReparsePoint -ea 0) {
            $ohook_found = $true
            break
        }
    }

    if ($ohook_found) {
        Write-Host "=== Office Ohook Bypass found ==="
        Write-Host
    }

    Check-Activation Windows RETAIL
    Check-Activation Windows MAK
    Check-Activation Windows OEM
    Check-Activation Office RETAIL
    Check-Activation Office MAK

    $output = Search_VL_Products -ProductName windows
    if ($output -and ($output -isnot [string])) {
      foreach ($obj in $output) {

        Write-Host "=== Volume License Found: $($obj.Name) === "
        if ($obj.GracePeriodRemaining -gt 259200) {
            Write-Host "=== Windows is KMS38/KMS4K activated ==="
        }
        Write-Host
      }
    }

    $output = Search_VL_Products -ProductName office
    if ($output -and ($output -isnot [string])) {
      foreach ($obj in $output) {

        Write-Host "=== Volume License Found: $($obj.Name) === "
        if ($obj.GracePeriodRemaining -gt 259200) {
            Write-Host "=== Office is KMS4K activated ==="
        }
        Write-Host
      }
    }
    Write-Host "End Check ..."
    return
}
# Run part -->

if ($AutoMode) {
    $actionMap = [ordered]@{
        "RunWmiRepair"            = $RunWmiRepair;
        "RunTokenStoreReset"      = $RunTokenStoreReset;
        "RunUninstallLicenses"    = $RunUninstallLicenses;
        "RunScrubOfficeC2R"       = $RunScrubOfficeC2R;
        "RunOfficeLicenseInstaller" = $RunOfficeLicenseInstaller;
        "RunOfficeOnlineInstallation" = $RunOfficeOnlineInstallation
        "RunUpgrade"              = $RunUpgrade
        "RunHWID"                 = $RunHWID;
        "RunoHook"                = $RunoHook;
        "RunVolume"               = $RunVolume;
        "RunTsforge"              = $RunTsforge;       
        "RunCheckActivation"      = $RunCheckActivation;
    }

    foreach ($action in $actionMap.GetEnumerator()) {
        if ($action.Value) {
            switch ($action.Name) {
                "RunHWID" {
                    try {
                        Run-HWID
                    }
                    catch {
                        Write-Warning "$($action.Name) Mode Fail"
                    }
                }
                "RunoHook" {
                    try {
                        Run-oHook
                    }
                    catch {
                        Write-Warning "$($action.Name) Mode Fail"
                    }
                }
                "RunVolume" {
                    try {
                        Run-KMS
                    }
                    catch {
                        Write-Warning "$($action.Name) Mode Fail"
                    }
                }
                "RunTsforge" {
                    try {
                        Run-Tsforge
                    }
                    catch {
                        Write-Warning "$($action.Name) Mode Fail"
                    }
                }
                "RunCheckActivation" {
                    try {
                        Check-Status
                    }
                    catch {
                        Write-Warning "$($action.Name) Mode Fail"
                    }
                }
                "RunWmiRepair" {
                    try {
                        Run-Troubleshoot -AutoMode $true -RunWmiRepair $true
                    }
                    catch {
                        Write-Warning "$($action.Name) Mode Fail"
                    }
                }
                "RunTokenStoreReset" {
                    try {
                        Run-Troubleshoot -AutoMode $true -RunTokenStoreReset $true
                    }
                    catch {
                        Write-Warning "$($action.Name) Mode Fail"
                    }
                }
                "RunUninstallLicenses" {
                    try {
                        Run-Troubleshoot -AutoMode $true -RunUninstallLicenses $true
                    }
                    catch {
                        Write-Warning "$($action.Name) Mode Fail"
                    }
                }
                "RunScrubOfficeC2R" {
                    try {
                        Run-Troubleshoot -AutoMode $true -RunScrubOfficeC2R $true
                    }
                    catch {
                        Write-Warning "$($action.Name) Mode Fail"
                    }
                }
                "RunOfficeLicenseInstaller" {
                    try {
                        Run-Troubleshoot -AutoMode $true -RunOfficeLicenseInstaller $true
                    }
                    catch {
                        Write-Warning "$($action.Name) Mode Fail"
                    }
                }
                "RunOfficeOnlineInstallation" {
                    try {
                        Run-Troubleshoot -AutoMode $true -RunOfficeOnlineInstallation $true
                    }
                    catch {
                        Write-Warning "$($action.Name) Mode Fail"
                    }
                }
                "RunUpgrade" {
                    try {
                        Get-EditionTargetsFromMatrix -UpgradeFrom $true
                    }
                    catch {
                        Write-Warning "$($action.Name) Mode Fail"
                    }
                }
                default {
                }
            }
        }
    }
    Write-Host
    Manage-SLHandle -Release | Out-null
    Read-Host "Press any key to close"
    return
}

do {
    Clear-Host
    Write-Host
    Write-Host "Welcome to Darki Activation p``s`` services" -ForegroundColor Yellow
    Write-Host "Choose a tool to activate:" -ForegroundColor Yellow
    Write-Host
    Write-Host "[H] HWID \ KMS38     {for Windows products only}" -ForegroundColor Yellow
    Write-Host "[O] oHook            {bypass for Office products only}" -ForegroundColor Yellow
    Write-Host "[K] Volume           {For Office -and windows products}" -ForegroundColor Yellow
    Write-Host "[T] Tsforge          {For Office -and windows (+esu) products}" -ForegroundColor Yellow
    Write-Host "[S] Troubleshoot     {For Office -and windows Products}" -ForegroundColor Yellow
    Write-Host "[C] Check Activation {For Office -and windows Products}" -ForegroundColor Yellow
    Write-Host "[E] Exit             {Exit the program}" -ForegroundColor Yellow
    Write-Host
        
    $choice = Read-Host "Enter {H} or {O} or {K} or {T} or {S} or {C} or {E} to Exit"
    Write-Host
    switch ($choice.ToUpper()) {
        "T" {
            Run-Tsforge
        }
        "O" {
            Run-oHook
        }
        "H" {
            Run-HWID
        }
        "K" {
            Run-KMS
        }
        "S" {
            Run-Troubleshoot
        }
        "C" {
            Check-Status
        }
        "E" {
            Write-Host "Exiting the program..." -ForegroundColor Red
            Manage-SLHandle -Release | Out-null
            break
        }
        default {
            Write-Host "Invalid selection. Please enter a valid option."
        }
    }
        
    Read-Host "Press Enter to continue..."  

} until ($choice.ToUpper() -eq "E")