using namespace System
using namespace System.Collections.Generic
using namespace System.Drawing
using namespace System.IO
using namespace System.IO.Compression
using namespace System.Management.Automation
using namespace System.Net
using namespace System.Diagnostics
using namespace System.Numerics
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

$ProgressPreference = 'SilentlyContinue'

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

        [ValidateSet("HGlobal", "Handle", "NtHandle", "ServiceHandle", "Heap", "STRING", "UNICODE_STRING", "BSTR", "VARIANT" ,"Local" ,"Auto", "Desktop", "WindowStation")]
        [string]$Method = "HGlobal"
    )
    $IsValidPointer = IsValid-IntPtr $handle
    if (!$IsValidPointer) {
        return
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
            $null = Invoke-UnmanagedMethod -Dll win32U -Function NtUserCloseDesktop -Values @($ptrToFree) -SysCall
        }
        "WindowStation" {
            $null = Invoke-UnmanagedMethod -Dll win32U -Function NtUserCloseWindowStation -Values @($ptrToFree) -SysCall
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
        @{ Name = "QueryServiceStatusEx";    Dll = "advapi32.dll"; ReturnType = [BOOL];   Parameters = [Type[]]@([IntPtr],[Int32],[IntPtr],[Int32],[UInt32].MakeByRefType()) }
    )
    return Register-NativeMethods $functions
}
Function Init-KERNEL32 {

    $functions = @(
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

[DllImport("kernel32.dll", CharSet=CharSet.Unicode)]
[DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
public static extern IntPtr LoadLibraryW(string lpLibFileName);

[DllImport("kernel32.dll")]
[DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
public static extern IntPtr GetProcAddress(
    IntPtr hModule,
    string lpProcName);

[DllImport("ntdll.dll", SetLastError = true)]
public static extern int NtProtectVirtualMemory(
    IntPtr ProcessHandle,           // Handle to the process
    ref IntPtr BaseAddress,         // Base address of the memory region
    ref UIntPtr RegionSize,         // Size of the region to protect
    uint NewProtection,             // New protection (e.g., PAGE_EXECUTE_READWRITE)
    out uint OldProtection          // Old protection (output)
);

public static byte[] GenerateSyscall (byte[] syscall)
{
    return  new byte[]
    {
        0x4C, 0x8B, 0xD1,                                       // mov r10, rcx
        0xB8, syscall[0], syscall[1], syscall[2], syscall[3],   // mov eax, syscall
        0x0F, 0x05,                                             // syscall
        0xC3                                                    // ret
    };
}

[UnmanagedFunctionPointer(CallingConvention.StdCall)]
public delegate int NtGetNextProcessDelegate(
    IntPtr ProcessHandle,       // HANDLE
    uint DesiredAccess,         // ACCESS_MASK
    uint HandleAttributes,      // ULONG
    uint Flags,                 // ULONG
    ref IntPtr NewProcessHandle // PHANDLE (output)
);

static void Main(string[] args)
{
    byte[] shellcode = null;
    uint syscallNumber = 0;
    IntPtr processHandle = IntPtr.Zero;
    IntPtr newProcessHandle = IntPtr.Zero;

    var Address = GetProcAddress(LoadLibraryW("ntdll.dll"), "NtGetNextProcess");
    if (Address != IntPtr.Zero)
    {
        unsafe
        {
            // Cast the IntPtr to a byte pointer
            byte* pFunc = (byte*)Address;

            // Read the first few bytes. 
            // This is the common syscall stub for 64-bit Windows.
            byte byte1 = pFunc[0]; // should be 0x4C
            byte byte2 = pFunc[1]; // should be 0x8B
            byte byte3 = pFunc[2]; // should be 0xD1
            byte byte4 = pFunc[3]; // should be 0xB8

            // The syscall number is in the next 4 bytes
            syscallNumber = *(uint*)(pFunc + 4);
        }
    }

    if (syscallNumber == 0)
    {
        return;
    }

    var syscallBytes = BitConverter.GetBytes(syscallNumber);
    shellcode  = GenerateSyscall(syscallBytes);

    unsafe
    {
        fixed (byte* ptr = shellcode)
        {
            // Store the original address of the shellcode
            uint lpflOldProtect = 0;
            var hProc = new IntPtr(-1);
            var originalBaseAddress = (IntPtr)ptr;
            var tempBaseAddress = originalBaseAddress;
            var tempLength = new UIntPtr((uint)shellcode.Length);

            // Call NtProtectVirtualMemory with the temporary variables
            if (NtProtectVirtualMemory(hProc, ref tempBaseAddress, ref tempLength, (uint)0x00000040, out lpflOldProtect) != 0)
            {
                throw new Win32Exception();
            }

            // Use the ORIGINAL address to create the delegate
            var ntGetNextProcessDel = Marshal.GetDelegateForFunctionPointer<NtGetNextProcessDelegate>(originalBaseAddress);

            // Example of calling the dynamic assembly via delegate
            processHandle = IntPtr.Zero; 
            newProcessHandle = IntPtr.Zero;
            while (0 == ntGetNextProcessDel(processHandle, 0x02000000, 0x00, 0x00, ref newProcessHandle))
            {
                Console.WriteLine($"NtGetNextProcess succeeded, NewProcessHandle: {newProcessHandle}");
                processHandle = newProcessHandle;
            }
        }
    }
}

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Clear-Host

Write-Host
write-host 'NtGetNextProcess Test'
$hProc = [IntPtr]::Zero
$hProcNext = [IntPtr]::Zero
Invoke-UnmanagedMethod `
    -Dll NTDLL `
    -Function ZwGetNextProcess `
    -Values @($hProc, [UInt32]0x02000000, [UInt32]0x00, [UInt32]0x00, ([ref]$hProcNext)) `
    -SysCall
$hProcNext

Write-Host
write-host 'NtGetNextThread Test'
$hThread = [IntPtr]::Zero
$hThreadNext = [IntPtr]::Zero
Invoke-UnmanagedMethod `
    -Dll NTDLL `
    -Function ZwGetNextThread `
    -Values @([IntPtr]::new(-1), $hThread, 0x0040, 0x00, 0x00, ([ref]$hThreadNext)) `
    -SysCall
$hThreadNext

Write-Host
write-host 'NtClose Test'
Invoke-UnmanagedMethod `
    -Dll NTDLL `
    -Function NtClose `
    -Values @([IntPtr]$hProcNext) `
    -SysCall

Write-Host
write-host 'NtCreateFile Test'

$FileHandle = [IntPtr]::Zero
$IoStatusBlock    = New-IntPtr -Size 16
$ObjectAttributes = New-IntPtr -Size 48 -WriteSizeAtZero

$filePath = ("\??\{0}\test.txt" -f [Environment]::GetFolderPath('Desktop'))
$ObjectName = Init-NativeString -Encoding Unicode -Value $filePath
[Marshal]::WriteIntPtr($ObjectAttributes, 0x10, $ObjectName)
[Marshal]::WriteInt32($ObjectAttributes,  0x18, 0x40)

# Call syscall
Invoke-UnmanagedMethod `
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
    -SysCall

Free-NativeString -StringPtr $ObjectName

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
        
        $shellcode = if ([IntPtr]::Size -gt 4) { 
            [TEB]::GenerateSyscallx64(
                ([BitConverter]::GetBytes($sysCallID)))
        } else {
            [TEB]::GenerateSyscallx86($funcAddress)
        }
        $baseAddress = [IntPtr]::Zero;
        $regionSize = [UIntPtr]::new($shellcode.Length);

        if ([TEB]::ZwAllocateVirtualMemory(
            [IntPtr]::new(-1),
            [ref] $baseAddress,
            [UIntPtr]::new(0x00),
            [ref] $regionSize,
            0x3000, 
            0x40 
        ) -ne 0) {
            throw "Fail to allocate Memory for SysCall"
        }

        [Marshal]::Copy(
            $shellcode, 0, $baseAddress, $shellcode.Length)
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
        BaseAddress      = $baseAddress
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
            if ($ApiObject.BaseAddress -and $ApiObject.BaseAddress -ne [IntPtr]::Zero) {
                $baseAddressToFree = $ApiObject.BaseAddress
                [TEB]::ZwFreeVirtualMemory(
                    [IntPtr]::new(-1),
                    [ref]($ApiObject.BaseAddress),
                    [ref]($ApiObject.RegionSize),
                    0x8000 ## MEM_RELEASE
                );
            }
            $ApiObject.DelegateInstance = $null
            $ApiObject.DelegateType = $null
            $ApiObject.FunctionPtr = [IntPtr]::Zero
            $ApiObject.DelegateCode = $null
            $ApiObject.Dll = $null
            $ApiObject.Function = $null
            $ApiObject.BaseAddress = $null
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
        Initialize-ApiObject -ApiSpec $apiSpec -SysCall
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
        [ValidateSet("Base", "Extend")]
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
        $callback = [Teb]::GetCallback();
        $callbackPtr = [Marshal]::GetFunctionPointerForDelegate($callback);

        if ([IntPtr]::Size -eq 8)
        {
            $shellcode = [byte[]]@(
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
            $shellcode = [byte[]]@(
                0x64, 0xA1, 0x18, 0x00, 0x00, 0x00, ## mov eax, fs:[0x18] ## Get a specific address from the Thread Information Block.
                0x50,                               ## push eax           ## Push this address onto the stack to use later.
                0x8B, 0x44, 0x24, 0x08,             ## mov eax, [esp + 8] ## Get a second value (a function pointer or argument) from the stack.
                0x50,                               ## push eax           ## Push this value onto the stack as well.
                0xFF, 0x14, 0x24,                   ## call [esp]         ## Call the function whose address is now at the top of the stack.
                0x83, 0xC4, 0x08,                   ## add esp, 8         ## Clean up the two values we pushed on the stack.
                0xC3                                ## ret                ## Return to the calling code.
            )
        }
        $baseAddress = [IntPtr]::Zero;
        $regionSize = [UIntPtr]::new($shellcode.Length);
        $status = [TEB]::ZwAllocateVirtualMemory(
            [IntPtr]::new(-1),
            [ref] $baseAddress,
            [UIntPtr]::new(0x00),
            [ref] $regionSize,
            0x3000,   ## MEM_COMMIT | MEM_RESERVE
            0x40      ## PAGE_EXECUTE_READWRITE
        );

        if ($status -ne 0)
        {
            return;
        }

        ## Copy shellcode to allocated memory
        [Marshal]::Copy($shellcode, 0, $baseAddress, $shellcode.Length);

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
        }

        $status = [TEB]::ZwFreeVirtualMemory(
            [IntPtr]::new(-1),
            [ref] $baseAddress,
            [ref] $regionSize,
            0x8000 ## MEM_RELEASE
        );

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
    
    $shellcode = Build-ASM-Shell
    $len = $shellcode.Length
    $baseAddress = [IntPtr]::Zero
    $regionSize = [uintptr]::new($len)
    
    $ntStatus = if ($Method -eq "Base") {
        [TEB]::ZwAllocateVirtualMemory(
           [IntPtr]::new(-1),
           [ref]$baseAddress,
           [UIntPtr]::new(0x00),
           [ref]$regionSize,
           0x3000, 0x40)
    } else {
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

    try {
        $Address = [IntPtr]::Zero
        [marshal]::Copy($shellcode, 0x00, $baseAddress, $len)

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
        [TEB]::ZwFreeVirtualMemory(
            [IntPtr]::new(-1),
            [ref]$baseAddress,
            [ref]$regionSize,
            0x4000) | Out-Null
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

        $ret = Invoke-UnmanagedMethod `
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
                [IntPtr]0, [IntPtr]0, [IntPtr]0
            )

        # according to MS article, this is primary Token
        # we can return this --> $phToken, Directly.!
        # return $phToken

        #<#
        if (!$ret) {
            throw "Failed to Call LogonUserExExW."
        }

        # Duplicate token to Primary
        $hToken = [IntPtr]0
        $ret = Invoke-UnmanagedMethod `
            -Dll NTDLL `
            -Function NtDuplicateToken `
            -CallingConvention StdCall `
            -CharSet Unicode `
            -Return int32 `
            -Values @(
                $phToken,        # Existing token
                0xF01FF,         # DesiredAccess: all rights needed
                [IntPtr]0,       # ObjectAttributes
                0x02,            # SECURITY_IMPERSONATION
                0x01,            # TOKEN_PRIMARY
                ([ref]$hToken)   # New token handle
            ) `
            -SysCall

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
        if (!(Invoke-UnmanagedMethod `
            -Dll Userenv.dll `
            -Function LoadUserProfileW `
            -Return bool `
            -CharSet Unicode `
            -Values @($hToken, $lpProfileInfo))) {
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
        $hWinSta = Invoke-UnmanagedMethod -Dll User32 -Function OpenWindowStationW -Return intptr -Values @(
            "winsta0", 0x00, (0x00020000 -bor 0x00040000L))
        if ($hWinSta -eq [IntPtr]::Zero) {
            throw "OpenWindowStationW failed .!"
        }

        $WinstaOld = Invoke-UnmanagedMethod -Dll User32 -Function GetProcessWindowStation -Return intptr
        if (!(Invoke-UnmanagedMethod -Dll User32 -Function SetProcessWindowStation -Return bool -Values @($hWinSta))) {
            throw "SetProcessWindowStation failed .!"
        }
        $hDesktop = Invoke-UnmanagedMethod -Dll User32 -Function OpenDesktopW -Values @(
            "default", 0x00, 0x00, (0x00020000 -bor 0x00040000 -bor 0x0080 -bor 0x0001))

        Invoke-UnmanagedMethod -Dll User32 -Function SetProcessWindowStation -Return bool -Values @($WinstaOld) | Out-Null

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
            $activeSessionId = Invoke-UnmanagedMethod -Dll Kernel32 -Function WTSGetActiveConsoleSessionId -Return intptr
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
    [CmdletBinding(DefaultParameterSetName = 'ByService')]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = 'ByProcessId')]
        [int] $ProcessId,

        [Parameter(Mandatory = $true, ParameterSetName = 'ByProcessName')]
        [string] $ProcessName,

        [Parameter(Mandatory = $true, ParameterSetName = 'ByService')]
        [string] $ServiceName
    )

    # Constants
    $SC_MANAGER_CONNECT        = 0x0001
    $SC_MANAGER_CREATE_SERVICE = 0x0002
    $SERVICE_QUERY_STATUS      = 0x0004
    $PROCESS_ALL_ACCESS        = 0x1F0FFF
    $PROCESS_CREATE_PROCESS    = 0x0080
    $SERVICE_START             = 0x0010
    $SERVICE_QUERY_STATUS      = 0x0004

    $scManager = $tiService = [IntPtr]::Zero
    $buffer = $clientIdPtr = $attributesPtr = [IntPtr]::Zero

    function Open-HandleFromPid($ProcID) {
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
                [ref]$handle, $PROCESS_ALL_ACCESS, $attributesPtr, $clientIdPtr)

            if ($ntStatus -ne 0) {
                Start-Sleep 1
                write-warning "get handle using PROCESS_ALL_ACCESS failed, re-try with PROCESS_CREATE_PROCESS"
                $ntStatus = $Global:ntdll::NtOpenProcess(
                    [ref]$handle, $PROCESS_CREATE_PROCESS, $attributesPtr, $clientIdPtr)
            }

            return $handle
        }
        finally {
            Free-IntPtr -handle $clientIdPtr
            Free-IntPtr -handle $attributesPtr
        }
    }

    try {
        switch ($PSCmdlet.ParameterSetName) {
            'ByProcessId' {
                return Open-HandleFromPid -ProcID $ProcessId
            }

            'ByProcessName' {
                $proc = Query-Process -ProcessName $ProcessName -First
                if ($proc -and $proc.UniqueProcessId) {
                    return Open-HandleFromPid -ProcID $proc.UniqueProcessId
                }
                throw "Error receive ID for selected Process"
            }

            'ByService' {
                $ReturnLength = 0
                $dwDesiredAccess = $SC_MANAGER_CONNECT -bor $SC_MANAGER_CREATE_SERVICE
                $hSCManager = $Global:advapi32::OpenSCManagerW(0,0,$dwDesiredAccess)

                if ($hSCManager -eq [IntPtr]::Zero) {
                    throw "OpenSCManagerW failed to open the service manger"
                }

                $dwDesiredAccess = $SERVICE_START -bor $SERVICE_QUERY_STATUS
                $lpServiceName = [Marshal]::StringToHGlobalAuto($ServiceName)
                $hService = $Global:advapi32::OpenServiceW(
                    $hSCManager, $lpServiceName, $dwDesiredAccess)

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
                $svcHandle = Open-HandleFromPid -ProcID ($svcProcID)
                return $svcHandle
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

--------------

SERVICE_STATUS structure (winsvc.h)
https://learn.microsoft.com/en-us/windows/win32/api/winsvc/ns-winsvc-service_status

--------------

# Create Process AS Normal USER
$ret = Invoke-Process `
    -CommandLine "cmd /k whoami" `
    -RunAsConsole `
    -WaitForExit

# Create Process AS System, using TrustedInstaller service
$ret = Invoke-Process `
    -CommandLine "cmd /k whoami" `
    -ServiceName TrustedInstaller `
    -RunAsConsole `
    -WaitForExit `
    -RunAsTI

# Create Process AS System, using winlogon Process
$ret = Invoke-Process `
    -CommandLine "cmd /k whoami" `
    -ProcessName winlogon `
    -RunAsConsole `
    -WaitForExit `
    -RunAsTI

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
        [switch] $RunAsTI
    )

    try {
        $tHandle = [IntPtr]::Zero

        if (-not ($ProcessName -xor $ServiceName)) {
            if (-not (!$RunAsTI -and !$ProcessName -and !$ServiceName)) {
                throw "Please Use -ProcessName or -ServiceName Parameters"
            }
        }

        if (($ProcessName -or $ServiceName) -and !$RunAsTI) {
            throw "-ProcessName or -ServiceName Parameters, Must Run with -RunAsTI"
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

        if ($RunAsTI) {
            
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
                    $tHandle = Get-ProcessHandle -ProcessId $ProcessId
                }
                else {
                    # Managed Code fail to get handle,
                    # if service need [re] Start. 
                    # so, i have to call function twice
                    # instead, this, work on first place.!
                    $ServiceName = $Service.Name
                    $tHandle = Get-ProcessHandle -ServiceName $ServiceName
                }
            } 
            elseif ($ProcessName) {
                $tHandle = Get-ProcessHandle -ProcessName $ProcessName
            }

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
        
        # Now, Update -> lpAttributeList
        if ([IntPtr]::Size -eq 8) {
            [Marshal]::WriteInt64(
                $startupInfo, 0x68, $lpAttributeList.ToInt64())
        } else {
            [Marshal]::WriteInt32(
                $startupInfo, 0x44, $lpAttributeList.ToInt32())
        }
      
        $CommandLinePtr = [Marshal]::StringToHGlobalUni($CommandLine)

        $Global:kernel32::CreateProcessW(
            0, $CommandLinePtr, 0, 0, $false, $flags, 0, 0, $startupInfo, $processInfo)
        
        $err = [Marshal]::GetLastWin32Error()
        if (!$ret) {
            $msg = Parse-ErrorMessage -MessageId $err -Flags HRESULT
            Write-Warning "`nCreateProcessW fail with Error: $err`n$msg"
            return $false
        }

        $hProcess = [Marshal]::ReadIntPtr($processInfo, 0x0)
        $hThread  = [Marshal]::ReadIntPtr($processInfo, [IntPtr]::Size)
        
        $ret = Adjust-TokenPrivileges -hProcess $hProcess -AdjustAll -SysCall
        if (!$ret) {
            Write-Host "Fail to Adjust Token Privileges for $hProcess"
        }
        
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
H4sIAAAAAAAEAO07DXhTVZY3bdI/qCnQQCkIAVJbxHbSpmBLiyTS4KuTau2PFi2kIX2l0TSJyQsURteyoaPhmRl0ddZZHVcZnHVmHe2o31Cqs6YUaSswFuTTIs7QGR1NLTpFWf4cfXvOvS9tWuAbv911d+db78fNuffcc849995zzz339lFx+w4STwhRQpYkQjoJS0byNZKCkCvmd11BXk4+tKBTYTm0oKbZ4dN6vO4NXluL1m5zudyCdj2v9fpdWodLW3ZztbbF3cjnpaam6GQRf7GpUjrJXWui+Vet59a8CnBl8M66Lgqb616h0F3XQaGr7iWAB7feWfca5blzTRvAxcH1lH5x8K66PRQ+soZBB61XOezNKH/yECrNhDR+P4HcvePt5ihulCzUTom74hqSCRVZ0Zq58JNGi20KIpfjCEmQeaKQ7JAnkzY3KKJMUXBxnRUrSwh5D+A1pYQ8iEgPIecQNoB+cTEKtwEN8iylJJdNAwZYQ0UMQk9Ih+Ky5CRP4FsFgFyGrFDm+DiiSQva5HkbbYKNkL3TmUySDnneRDoj/MtjZKQNJ6aS0LkiBRfRhfM8jJCOsUHus/AS8rw+r53Ic+KR5RVdio53uu1sjnCuKN3yi+iuv/xM/P9KXEjlWE1Iz0E5cUGLLokL1ujSLMEyXYYFqtrqW2/jAif1nHieCwm6uj1oRpJm/QLgFgcjXZIkmQ39XFDQaS1iLye+xYlJkqYEmwP79NZ1d4wJh9SJzDEIYy1IN9WYak23mm7jQk7dMLf15GOomKhqWAnAsI8LmiVONN9nkKD0Qrl4iguZe7lg7W6zOMSBklpOfJcLVrwGSvSaT6D9mA1hLnjvO1yo9kTw3l6LeKpM/HOZKEmaq7SEtIf96yTNy/MJCZzbom5fBQYiru2VNNsBUyb2dCXRzvskzVZAWMQI1x5Wb7syDhsrerHrCvF11i2ILxd/K2kG5qNU9bYDCuRcu9sSqviiXDxkCZlRhbOS5sQ8RvAMJah4B9Uzf8T1mkdoZ+39wjzTnji6GrUfmTqn0vl9UWbyUKYDkuYWQATCSSb1C/vizSNB85vioKQ5Om+S1gfmodajVOulwCrWvmnqjIu2tshCE6Gl+N6zQr1o/qg9LNSKJ9Tz26g9gHqg+JuwisF735Q0c2WGtwmKrXizC0WZOmHNJRi7Sb3bFIdzIWlGrmSEv0Tle0rWfuR9WTxhbv9AuA7W7gVuyX5cvKYdIH83lVEmTk0oExPQ3rR0FndSAYIHyaUeLrBfaSp5zwsqnIIGOj5TZ050HH+6EtU5B6MUVr6iYFjxLUlTy4QsAVuoEPeLb5WLIKCIIWdC36/holWIsNwz2MDMX0iaHdAO1IYjVPQWrLV/IqRLmk204XW5wUYbJGxYRxt65QYOamj4FRJnOBh5QwXabN2HNmyymtaZ1prqYQ+s7bnI3jnxJNh6G7N1oTRq66No5xUdFvFjWIEwGFxn8N5OUHuAGh018gHZyMvEj8SzYN6SpoIO0X+7pDk+Vzbsq6lhd0qaX82dZCLPzsXJG6ImomSG3Yn9VIj7ooY9QJfk7Fy2pi/CDINRRyzBV3UDaAghc7hcPAh0RyVNnEyEx2Kx/6h6m5dZeRjsWr2Nh0pT8d1K9bY7oASGH6odlG09aB5kth40D0manjlMTLECFfK/J2nOA8Ys7h/Xe3ROdNHV26Yr0B6oMYpv09aHZQG/I1QPYZNoHkR7QtM2EvX8bUS27zCo30ftG7hKZa5/YfbdZyrK92vKQ7ckwQaWNGcy6bRm0Ra09Rjzx5mi5nWKEgnNaPWD3icljXHOxQZbOAcLoxcZ7K7MGIPtRd8J835/5pjBduBiWMT94G/njBnsicxY8+uH2ojqNQUdAVugCth3YKa7Kd0+me7pTPQ1Z7DhyUykjVr8A5my/Y6i/abEU/tFu7zIeuF8CJws4kJluspeyJJCDbZg0dVRszFR3w74OpzoyMILkgQHQFGs95/ED2SjsfxxjL/mvng1lTOF1evR3CNHzl9KnrEWz6Zm2EEFxfIOKtMZDZ/gcdSAHektwXt0lRbxQtCp03cxw3PqiuiJRfdTmBNPiV+CHc+ma62TNI/NpltImC1pQlAUB7hQKq5f5MQ5SRregeYdfEKHgZ5Yr9NLmsrZkzbYDbOxgIbqBxnlURl1KONxlFFGjdyi0zMNYfJRF6YdGsDGDKoLME+NMlcj8zpkRt8t/s4CA6wrnq/+/qeUyaIrwnMkAc4N4BU/50I1gOn+UMnFn+XEm+BAzsxAOsSGUl0obBoIE1enZQF/eAX+CLOg6QVsOnMWDLynZL73Z8M/pOJRwOFZqPSopGnImLQxbwNEeTARxztT0qzJkFVuBruK/CuIGqmyiKuTOHHAhIP7wyw6uFxJ850o5adIufUsninzJQ3JQOvska3zYyAfmQZECtSsHoh+w0wdxs+JbB3o+C0iLDQn4rqzmXh2Fho8ddi/nIUc0R3zj1CDQxyYjGjxGPKCZTVbx8y8qWnMX5SkHltGiD+N6+5OazrbLUl+RQ/X3Z82vB32CbXnnPYjQnHgY4WQgz8zAxcUQgrdCiAzpydyYwJuz8gcAOpfhymuHAYNcFhNt2pY/Wg3rWZQy6rH+EvQpXFgt1o4JQInIcD5DLy8oT/SBUK2gbfqww1hgoNm++ADYO7G3YrAinI4QYi6/XXcM6riUno7I5EHVcgh3Bz5YTLUfgs7KfKJEgMNTRzaQyj1c5iLyPNT6ZIsjixm5HAEaI7Mou092P7AVBT6QQkLp40F0vYDkR8kQmmb5F8ZOZpII0EuUEQEHUxw5DhVU8gAF0NtGg4TToy3iKk3gTCTdFhSqWBO6RyNxOEGhG2EC1fEwRJyMBda07qePYlUf/jdQeNSmBcahhq3nxdVry9Fhe9LZcxGGmj2SH1B1XPQEMlWyjOjDawoQKX9yyIlMLLI72HNIt1JOJhZJSwAhpHsA3UNR4zb34z8ni7SYOQ6IBZ7Rn7KdEhCHUAJDs/HOq0laMwIVsL6cElwlOMCGcvFUfECLFH7Eb+6WJMDOtyXaAgP/wm4Q9dLgXOKTUbwUOGZaI9hvypUpxiBDWWEujgIA9O3H1dve54a9X7xHZityDNfwt4cBKegx8PpH2Kbuk5HmwKSwr8M5BVQV45tT0IbuHgwqMiBr6AoqiysT+GKKA2cFxqYO0lQBaR4f+mYWENMj0IWcE6P1TYtSgdxsiaq9UiC4ThCsQcXoRUWwWhaZ13XM27DuGBwSJ6DAF08BRvDr4o8rqA7kZ50MQsPa24dfhx0Rp9ezYz+mKFf0sSny+Hv3hkYFStpta8rWUlwL2nv4KSsczNw3+C9BQ6XPbNo9HxyBj1CE2mocWgxGPL1cKuM/Jw6izLoMKhavBQNHMsQVUN9TyGNxKYjOqgKGBilEcr3F6LFrAUMdturus3AdlevyiaXOqkvvCuMO2YNcvYqFXET8GLqQwVUTaU2ilQA8tUoEhifhbFFJJiDHeycM7KLF5jOMUlTjYMUX8L7r0HCUwscfAM9LPZMZ0tcit4dzjyLrpmDU5Neji1iH84MHDH7scGDlRw8+uoxZt8EnFJvQIq753lcPLyqYUtGLYam9Manei8fhL9SsPNE7o+XoDfs8W/iQhVpcHehDxiS5iXsXTSnYeQqaR6Fmhjm8rG8mZVDFRmsbsO6OQMQo1z3kJYzQNlghpijm9sj0QSSFkHFMN4bpy7rhgW4FbUAY6nkzg7geuSzuyX4a5i24muhJmUtnE79cSh1MdSH30ikd9tcKPegj47MOIXhQ5Iu8qFEC3HUE1P/Hbk/kRrLAp4TlVl6aiJc0oiKK+kT0g3hUU79q9TTeAh8WsD8d4di5CyLZXIi32W8iZFfgNMZmR2ZA/2AC2xmaGVBeCQxYkvExxfKijzgyHMV1JE/yuRFZVUnItaPTNMjq1klMbICCiMnY/gjSwATuWpcaNQ/YmSThhsvQ95459Exwd4bjuxKoB5+ZuALhX+aHPod4vAi1DWNxQV1cBxWyjvSia5YoH4VtuWPEnC6oV8rFSKkwSSfXsrOt+GHABd5kklXRYIJ8rEnj3OrXB9ulWW09/sT96t+CdyKSCt6Y2vC2JjTI1WskgycN+Cx+fHYAJlPaFshLMWw65h/Q+C8YtOdkUMqqlFO+3G/DrRah97/PMX5UcsmrA8zmjzowdCkdqvW4IADfdLn6vtUuWyzLwOA1WWsuhLAflUyKkkV0N7Rs0eFJ8FXRD7z4Wjr7qy4He82qj2/h6CglbOnnoQCHJw/RxD/xtZifFv0WzunxCGZKcNfY1myV737lnQusDeDiz8M94kELjS1zRKa+gYXTNeVlxwWMsQbp1pKur0p4vVJ8d2WkrAXnNrenJETsF5wuNBZCmzRkXuxlAJ2mAAwjpplRg+bo+3HIruVeKDBmKdsO+5P5O7XLCukywDjiM7jEnxj235MUG474p8aeVdJD8K3lLJJaembDeyfWiDr6VUVyC9yPczOarnQczrjFzDMrSfDeFQek73uxiuY1YsHDi3eg8/KkTDsNtAd9rypE6+tkWtp9FEB/mDrFWh4P4nQd49HB5XUz4NnkzQfpTKPtpITGT4UGkJwsVfrgaaIkrn/HOpZ8D0Bz43HUYbY3kfZ0bKD7dgT6kKRpk76ykRdf/sDqOpP4sc54BxswBsKnHVGvJGwJ68XUvHoA/eJFxh60YBr8ygWOElzeyrryShpZqVSC0yFzSYkhbj4SBc9jJ/QDbC3AZyztT3jMQ04/ysXaZB9xVwAI9MhWsLxQww7Fycp0JPElZzwDrPdiXuyZwLv2XTKO5o+kXco9dK8e8GIkT8N/UUGfbhDSWmGsKFfvXPrKZfAO7d+5nDx5VMOixEapm39/Abe5Z/yVoiT1DvF8/51WaMrJTXXq8rAh9Y2XBuV5hr8E4IkrQgnKIiQs6IhIY4IC1Z4EGSqOs7/BQIh7SYdt0dBFPTQ5NQ/7fall4mqE+DPTYE+RVlQdRSKI4mAO5CPTxz9JvGUWTxsFo8Fvky8x2gKmRJRR/GzMvGAuuv9ZN8U4IsDvify8c6muCe1E0O2EHdBvbNMPNKJr/clkXsoU5l4HA60O/BgCQx91auqXRI9xW9ewgJb2PqlUFR3fZTumw9Ef+5V6cebFixB2kVL2GM/EM1RBwz4qNGvVgxw3X/UclMG8OVUC01z1YG/yKGGVpdQmqBu/wT9lmr4akJeHAr0JfWq3oYibhEYM5RMZ7q1wvpAn7ZX9RuoqmjDK1DaA1QDoPdPoVy2qB+0+hJ0eRRrJf3+a2kPC/cu3+vHmO3vAB1IDcAvuO6h472qu69m2gLL+qtxZoFozdWo/YezfFM5ddcorL/qBlRA3fXHGb4bx+WtNIkHTLDBuofTsiSiJIE/JARV+Xo04P7FMHPc2cPcotSpV6PwVBWAUsXGKVzgD69B9cxiJlDj01CB0NHQDF8ygKXvQlMiGmMOxu7O8fgvzYo+xxAuVg0CCRyN9DImZT2cAmaVtY3+bqS/HfT3Dfr7Kv19nv4+RX+3TKI/eLBD0QNy2VlfIZrgaHujXDwcSSXyHSR6hhqrzWIGJ75rCrx/ziL2m84QpUU8LMwwicakCnul8uygJf4wZ++zLDpcbu+5UUxP48RVaRAqAEg6U6aIVwtp6l8D3DOKMcaibkt8n8UAffXfMfwyhAVS1m+TUZ+L0/h9Uxr6K+1ZKTAmYy2LkD/eXkHP6LXD5AKEzjFobeTT8yyQ1IKnGWtgxNZzUWI94kSFmEYdKBqvJZSaA7paRLNHrG3mxIqGyFNnmSQ9k0Q5DP1bi8BIw+pHusVuwJNv099kylnJ4BEZvirDDhk+IkNBhvUyrJIhJ8Nz1zGYLtdPyrBo5cT+OkwMPiHDv5fhTpn/RzLcJsNWGR6U6Zrl+sdy/SsZTpf/opcnQ06Gu2S4ToY/kKEgQ5dpon5HZfmvXjcRX2pkcK4MlTLsnPSXxE9l/EJZbooMvy/jD8rwRRk+JcOHjRPl1Ml1YRL+9CL29+co3JHNYMckGE3OSfX/rrSjksnV3zJRvjQp3UB4IpAa4iZ3QclFyiE3Qc1LWogNWhxQdhEtySH5xAd1LbGTZmhxAeTJYqg3Qc1BnFBrhFouZD1pJVnwWwQwRe6Xg3Yb8RATUDpBpv1rcn5d/SbKyiEFgG28hLaX068aJJeR5dCaBeMkZBXI9ZDNgHdQOV9vlIT2bQHeDVSvxbQlJWb+jd/EYmOqvLQddeQzfDj/Mnb2TelzmdS29NJ6dMr69clwQIaDMhyK0T/6LQjGfbdA/gAugh8smtiGf7fCTx+KlkHWTWzDwKIOmC3QZpnUdsUkvfDtOb4tvu2xpeyzh4NLGe7zuYS4rwQ/CDl1PiGLIL+/gJDfLYTYCHS5OQt8GuQw5LSrCLFBNs4mpAE6/wDKeKP5dTb4FYBV1WXV3SNXblx496JVz/3z6bZPph35GMe6anl9rY/3+upNjS0Ol8MneG2C21tfxvvuEtye+kab3Vnfuqywvop38jYfTxF5nsb1Ud01csZ5mgH5hlU1Fvx2pG2W/D2JrsVF1wT0gKkYw+n0+jHdGK517PuUPAd+D6LDWz7aFv3EI0+vtzdtgBsa1PHpLG9VVY2ubhX1rKMTcbcj7twEXDnzwAUTcJQuaQKuktKlTcRRuowJuBpKpy1gn8nIOEqnB1wlhPTRb2TAFmE3R+u6jW6n0NIo29O+uDH8li1bGtfD8MizRXI/XsGuKzfRfjom4G6n/XTG4moYXXgCjtH1Ae4B7KeV6RP9xiiPZ3X81qiVjM05/TpnVymzaxlnoP2VTlgb+haA5/DtyjHcMjrHleyTm7Fvf/APfHD7yFvv89H2BvadT/SbH/yOp4Hq7PPadfp8SgK4x8Zxk74YunxSzFGSOVWzPDMbNPrZ2SlktmIGSR2dMpQ8kBhWeZQN8fq4ypNXMf1mL0oims8ubsPPihYkxZGk05nEAwadCxNxGuZwAaoaU1coFUQJu2D2zHiSuDfeEwfjRH4j8quVRH06hagapxJ87x+EXFosyy2cTuvLYH9UIk4VR1QfTKF1fEBpLmZ9Ta5P4MG1jpGL+3iaBsYjJBNNYyLRFCYQTcG0oaeuYvuyD3zCtZnMF2D6Jyg/E1NXwO08oSCuclqKkqQUJpGUqgQP8uKfXPHPzPhHt0My7VGoa6D+oVyfB+XFMe1Yvzqmnc4TxEfJmQkks1BNMr0zh9I9MxpQPuptA9rrwM9dyGaQ6lIVV6lQJxA1jEUNY1EXTPEokhUkGeZbkaQkSV6lJ74hTi/rrVekKklqYQpJrUqoTMb+wD9G5eMnSK9D1ucwqIhTEFwrRXoSSW+EXJNI0qGP9II0j0KjgrnrIhpFJ0lAGjofCSSlIMHD6IG2BmgLVSR9eponelYgpHL147bInrIJ+UUJa1fJuRXqD0J+CnIbOA89rKFnJltPtK+XAP8m5BOQ/x1ySinrIY7Ew3pQ79vodBJTY6PJbud9PpPT6d7EQ4W/HtxzY3V5GWvja9yyDx+r3+ZwNbo3VQs2weF2kRt4weLe4HZVOxpXe90tNe67eBep4lvcG/nJkn1jouGusGJ8jM/D/vfAmPticCcBl5bP/EU0ncH4H+iGYuhWw0bJAbrKmBj7GvBij6HfjsFZAFeZf6nd/u295W/t3vJgHBqdzSeYvV63l5DvxXO8zYN2ZifvY1ul142Wh1hC7qOtq708j17E4rbbnLTyXXPVTWaLoYDuA/KhAvgwgrl5/Z28Xajm7X6vQ9hMHoivviS+ttpcFeXlkDfaArvF7nV4IPgpg01GFmGbye4sdzW5vS1syyxDnIV3bRCaYdcQsllR7nIIDpvTsQX2jJMsZDyg70x504FfvtRWNbeSLTG8F6tAyB/jqi+rm0phvttvc1Id7sA+6eaN1fR7ZJXbtZH3CkBT464WvA7XBijeRkg7tHg2I6up7FZTZXl0LpKI1brK6vPwdkeTw25ttrkanTzokQV4n9BoFTZ7eKsDurA28hAkujdbnRAsgs8mLXyLj4fSrauqam+qKa8w5xfqmcxExutwW+3ulha3y7qxaZMHVBGa0BqsVpvdK1gd7vXWJr/LDtEhyHcIAu9tIdeOl608WUmsPr7Z2uRwAsKKojMIiHQ1OTb4vbzVZQNz2mS1eTdsJMTAONnERpt410aH1+1q4V0CRtixFG4X3+oQrIJtvRNWayGx8q0w6cKkhpnQH1bhvPI4clt8uZscrlzQPpcOL9eZn5ufywY9qd3rdwmOFj6WIklVJThX2TwC6A6rRL+5TkOcxe2+y+9ZDXOBa2h2Cd7NhMzCllsdXgEWvNYFYmHNL6hqXWyBGs2tdt6D5Kvp5BCyW4V2f7lmI+6yVX6vF2ZC3mzkHVUNTLPDZRP4KIrsji/3yRW3dzVvQ10rvbwPJ7BNeYuf926u5L3U3lx2GAWMEqSbLpZe3kjMMdiaZi9vawQkmR2Pe2+zT+BbamCCTD7QkMcS+XF8zM6wgJGBG2gkL4BGZfx6/4YNvDeqSSnant2zmXyb/ksp+pf9/2w6tFZ7pOlo9M/34+8/GPew/7aQFCWl1f/h54Bv019JaRCH8mnsrogx6Z0Qhx4tmnjfQXgcLvppxQzmQLyqK2bxqx1yYTGLY1+Sy7H3k9i7S+y9BuG/ZRByTzGD+yHvKGbwXchPFTOohvtLB5S1AK+DHC5m7xw7INcvZ/BncC/ZB/hnAS6BO0US4K+Zw/7PTuNyBh+C/AiUHwHYB7lzOYOPXwmxFpSfAJg2j5DnljP4GuQwlD8AKM1jfZH5hBTPZ+VSgB65jLBdLj8AsEMuI3xdLvcBzNHCfC5nsE7L8Aidchnhz+TyswANC5iehQAPYl7O4KiMRzh7ISvrAX5vEYwd7g33AOyEnFHC4Gm5fA5gCO5GuhJ2X03OZn1NzabvxKSshL4nk4chF5XQt2R6TzSWMEjvecsZ/FAuI7wglxHiXevB5QxykE8uv7TdfZv+LyYFfXvMYM8kE/DovPWXwCfj1xKEvds8dAmJpStbW5xaCEd9EImsyM7P02dreZfd3QhR6Yrs2prVuUXZWp8AIYvNCWHXiuzNvC975XWpKaU2n49vWe/crAUBLt+KbL/Xtdxnb+ZbbL7cFofd6/a5m4RciC6X23wteRvzs7UQkDiaIEK9NbY3FPWdqCyofBPT9rec9Oz/ynU83fl0+Om+pweeHnq6aKdxZ84u4y5u18CuyK7RXed2kWeSnvnfVvTb9E2k/wDnBaePADwAAA==
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
            $ret = Invoke-UnmanagedMethod -Dll Advapi32 -Function CreateProcessWithTokenW -CallingConvention StdCall -Return bool -CharSet Unicode -Values @(
                $hToken,
                0x00000001,
                $ApplicationPtr, $CommandLinePtr,
                $flags, $lpEnvironment, "c:\", $lpStartupInfo, $lpProcessInformation
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
        ($lpDesktopPtr) | % { Free-IntPtr $_ }
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

# work - job Here.

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
#$Global:PKeyDatabase = Init-XMLInfo

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