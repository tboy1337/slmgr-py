"""
Windows Software Licensing Management Tool - Python Version

Copyright (c) Microsoft Corporation. All rights reserved.
Converted from VBScript to Python.

This script provides command-line management of Windows licensing and activation.
"""

import os
import sys
import winreg
from datetime import datetime, timedelta
from typing import Callable, Dict, List, Optional, Protocol, Tuple, Union, cast

try:
    import pythoncom
    import win32com.client
    import wmi
except ImportError as e:
    print(f"Error: Required Windows modules not available: {e}")
    print("Please install: pip install pywin32 wmi")
    sys.exit(1)


# Protocol classes for WMI/COM objects (to satisfy mypy strict mode)
class WMIServiceProtocol(Protocol):
    """Protocol for WMI Service objects (dynamic attributes from WMI)"""

    def query(self, query: str) -> List[object]:
        """Execute WMI query"""


class RegistryProtocol(Protocol):
    """Protocol for WMI Registry objects"""

    def SetStringValue(
        self, hkey: int, key_path: str, value_name: str, value: str
    ) -> int:
        """Set string value in registry"""

    def DeleteValue(self, hkey: int, key_path: str, value_name: str) -> int:
        """Delete value from registry"""

    def CheckAccess(self, hkey: int, key_path: str, access: int) -> Tuple[int, ...]:
        """Check registry access"""


class SoftwareLicensingServiceProtocol(Protocol):
    """Protocol for SoftwareLicensingService WMI object"""

    Version: str
    ClientMachineID: str
    KeyManagementServiceHostCaching: bool
    KeyManagementServiceListeningPort: int
    KeyManagementServiceDnsPublishing: bool
    KeyManagementServiceLowPriority: bool
    RemainingWindowsReArmCount: int
    VLActivationTypeEnabled: int

    def InstallProductKey(self, product_key: str) -> None:
        """Install a product key"""

    def RefreshLicenseStatus(self) -> None:
        """Refresh license status"""

    def ReArmWindows(self) -> None:
        """Rearm Windows"""

    def SetKeyManagementServiceMachine(self, machine: str) -> None:
        """Set KMS machine"""

    def ClearKeyManagementServiceMachine(self) -> None:
        """Clear KMS machine"""

    def SetKeyManagementServicePort(self, port: int) -> None:
        """Set KMS port"""

    def ClearKeyManagementServicePort(self) -> None:
        """Clear KMS port"""

    def SetVLActivationTypeEnabled(self, activation_type: int) -> None:
        """Set volume activation type"""

    def ClearVLActivationTypeEnabled(self) -> None:
        """Clear volume activation type"""

    def DisableKeyManagementServiceHostCaching(self, disable: bool) -> None:
        """Disable KMS host caching"""

    def SetVLActivationInterval(self, interval: int) -> None:
        """Set VL activation interval"""

    def SetVLRenewalInterval(self, interval: int) -> None:
        """Set VL renewal interval"""

    def SetKeyManagementServiceListeningPort(self, port: int) -> None:
        """Set KMS listening port"""

    def DisableKeyManagementServiceDnsPublishing(self, disable: bool) -> None:
        """Disable KMS DNS publishing"""

    def EnableKeyManagementServiceLowPriority(self, enable: bool) -> None:
        """Enable KMS low priority"""

    def SetKeyManagementServiceLookupDomain(self, domain: str) -> None:
        """Set KMS lookup domain"""

    def ClearKeyManagementServiceLookupDomain(self) -> None:
        """Clear KMS lookup domain"""

    def DoActiveDirectoryOnlineActivation(self, product_key: str, ao_name: str) -> None:
        """Do AD online activation"""

    def GenerateActiveDirectoryOfflineActivationId(self, product_key: str) -> str:
        """Generate AD offline activation ID"""

    def DepositActiveDirectoryOfflineActivationConfirmation(
        self, product_key: str, cid: str, ao_name: str
    ) -> None:
        """Deposit AD offline activation confirmation"""

    def ClearProductKeyFromRegistry(self) -> None:
        """Clear product key from registry"""

    def InstallLicense(self, license_data: str) -> None:
        """Install license"""

    def ReArmApp(self, app_id: str) -> None:
        """Rearm application"""


class SoftwareLicensingProductProtocol(Protocol):
    """Protocol for SoftwareLicensingProduct WMI object"""

    ID: str
    Name: str
    Description: str
    ApplicationId: str
    PartialProductKey: str
    ProductKeyID: str
    ProductKeyChannel: str
    LicenseStatus: int
    LicenseStatusReason: int
    GracePeriodRemaining: int
    LicenseIsAddon: bool
    KeyManagementServiceMachine: str
    KeyManagementServicePort: int
    KeyManagementServiceLookupDomain: str
    DiscoveredKeyManagementServiceMachineName: str
    DiscoveredKeyManagementServiceMachinePort: int
    DiscoveredKeyManagementServiceMachineIpAddress: str
    KeyManagementServiceProductKeyID: str
    VLActivationInterval: int
    VLRenewalInterval: int
    VLActivationType: int
    VLActivationTypeEnabled: int
    TokenActivationILVID: int
    TokenActivationILID: str
    TokenActivationGrantNumber: int
    TokenActivationCertificateThumbprint: str
    TokenActivationAdditionalInfo: str
    OfflineInstallationId: str
    RemainingAppReArmCount: int
    RemainingSkuReArmCount: int
    EvaluationEndDate: str
    TrustedTime: str
    IAID: str
    IsKeyManagementServiceMachine: int
    KeyManagementServiceCurrentCount: int
    KeyManagementServiceTotalRequests: int
    KeyManagementServiceFailedRequests: int
    KeyManagementServiceUnlicensedRequests: int
    KeyManagementServiceLicensedRequests: int
    KeyManagementServiceOOBGraceRequests: int
    KeyManagementServiceOOTGraceRequests: int
    KeyManagementServiceNonGenuineGraceRequests: int
    KeyManagementServiceNotificationRequests: int
    ADActivationObjectName: str
    ADActivationObjectDN: str
    ADActivationCsvlkPid: str
    ADActivationCsvlkSkuId: str
    AutomaticVMActivationHostMachineName: str
    AutomaticVMActivationLastActivationTime: str
    AutomaticVMActivationHostDigitalPid2: str

    def Activate(self) -> None:
        """Activate product"""

    def DepositOfflineConfirmationId(
        self, installation_id: str, confirmation_id: str
    ) -> None:
        """Deposit offline confirmation ID"""

    def UninstallProductKey(self) -> None:
        """Uninstall product key"""

    def ReArmSKU(self) -> None:
        """Rearm SKU"""

    def GenerateTokenActivationChallenge(self) -> str:
        """Generate token activation challenge"""

    def DepositTokenActivationResponse(
        self, challenge: str, auth_info1: str, auth_info2: str
    ) -> None:
        """Deposit token activation response"""

    def GetTokenActivationGrants(self) -> object:
        """Get token activation grants"""

    def SetKeyManagementServiceMachine(self, machine: str) -> None:
        """Set KMS machine"""

    def ClearKeyManagementServiceMachine(self) -> None:
        """Clear KMS machine"""

    def SetKeyManagementServicePort(self, port: int) -> None:
        """Set KMS port"""

    def ClearKeyManagementServicePort(self) -> None:
        """Clear KMS port"""

    def SetKeyManagementServiceLookupDomain(self, domain: str) -> None:
        """Set KMS lookup domain"""

    def ClearKeyManagementServiceLookupDomain(self) -> None:
        """Clear KMS lookup domain"""

    def ReArmsku(self) -> None:
        """Rearm SKU (note: typo in WMI method name)"""

    def SetVLActivationTypeEnabled(self, activation_type: int) -> None:
        """Set VL activation type enabled"""

    def ClearVLActivationTypeEnabled(self) -> None:
        """Clear VL activation type enabled"""


class ADObjectProtocol(Protocol):
    """Protocol for Active Directory objects"""

    Class: str
    Name: str
    Parent: str

    def Get(self, attribute: str) -> object:
        """Get AD attribute"""

    def GetInfoEx(self, attributes: List[str], reserved: int) -> None:
        """Load AD object info"""


class ADRootDSEProtocol(Protocol):
    """Protocol for Active Directory RootDSE"""

    def Get(self, attribute: str) -> str:
        """Get RootDSE attribute"""


class ADNamespaceProtocol(Protocol):
    """Protocol for Active Directory namespace"""

    def OpenDSObject(self, path: str, user: str, password: str, flags: int) -> object:
        """Open DS object"""


class TokenActivationLicenseProtocol(Protocol):
    """Protocol for Token Activation License objects"""

    ID: str
    ILID: str
    ILVID: int
    ExpirationDate: str
    AdditionalInfo: str
    AuthorizationStatus: int
    Description: str

    def Uninstall(self) -> None:
        """Uninstall license"""


# Constants
WINDOWS_APP_ID = "55c92734-d682-4d71-983e-d6ec3f16059f"
DEFAULT_PORT = 1688
HKEY_LOCAL_MACHINE = winreg.HKEY_LOCAL_MACHINE
SL_KEY_PATH = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
SL_KEY_PATH_32 = r"SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
NS_KEY_PATH = (
    r"S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
)

# Error codes
HR_S_OK = 0
HR_ERROR_FILE_NOT_FOUND = 0x80070002
HR_SL_E_GRACE_TIME_EXPIRED = 0xC004F009
HR_SL_E_NOT_GENUINE = 0xC004F200
HR_SL_E_PKEY_NOT_INSTALLED = 0xC004F014
HR_INVALID_ARG = 0x80070057
HR_ERROR_DS_NO_SUCH_OBJECT = 0x80072030

# AD Constants
AD_LDAP_PROVIDER = "LDAP:"
AD_LDAP_PROVIDER_PREFIX = "LDAP://"
AD_ROOT_DSE = "rootDSE"
AD_CONFIGURATION_NC = "configurationNamingContext"
AD_ACT_OBJ_CONTAINER = "CN=Activation Objects,CN=Microsoft SPP,CN=Services,"
AD_ACT_OBJ_CLASS = "msSPP-ActivationObject"
AD_ACT_OBJ_ATTRIB_SKU_ID = "msSPP-CSVLKSkuId"
AD_ACT_OBJ_ATTRIB_PID = "msSPP-CSVLKPid"
AD_ACT_OBJ_ATTRIB_PARTIAL_PKEY = "msSPP-CSVLKPartialProductKey"
AD_ACT_OBJ_DISPLAY_NAME = "displayName"
AD_ACT_OBJ_ATTRIB_DN = "distinguishedName"

# WMI class names
SERVICE_CLASS = "SoftwareLicensingService"
PRODUCT_CLASS = "SoftwareLicensingProduct"
TKA_LICENSE_CLASS = "SoftwareLicensingTokenActivationLicense"

# WMI Query clauses
PRODUCT_IS_PRIMARY_SKU_SELECT = (
    "ID, ApplicationId, PartialProductKey, LicenseIsAddon, Description, Name"
)
KMS_CLIENT_LOOKUP_CLAUSE = "KeyManagementServiceMachine, KeyManagementServicePort, KeyManagementServiceLookupDomain"
PARTIAL_PRODUCT_KEY_NON_NULL_WHERE = "PartialProductKey <> null"
EMPTY_WHERE_CLAUSE = ""


# Error messages
ERROR_MESSAGES: Dict[int, str] = {
    0xC004C001: "The activation server determined the specified product key is invalid",
    0xC004C003: "The activation server determined the specified product key is blocked",
    0xC004C017: "The activation server determined the specified product key has been blocked for this geographic location.",
    0xC004B100: "The activation server determined that the computer could not be activated",
    0xC004C008: "The activation server determined that the specified product key could not be used",
    0xC004C020: "The activation server reported that the Multiple Activation Key has exceeded its limit",
    0xC004C021: "The activation server reported that the Multiple Activation Key extension limit has been exceeded",
    0xC004D307: "The maximum allowed number of re-arms has been exceeded. You must re-install the OS before trying to re-arm again",
    0xC004F009: "The software Licensing Service reported that the grace period expired",
    0xC004F00F: "The Software Licensing Server reported that the hardware ID binding is beyond level of tolerance",
    0xC004F014: "The Software Licensing Service reported that the product key is not available",
    0xC004F025: "Access denied: the requested action requires elevated privileges",
    0xC004F02C: "The software Licensing Service reported that the format for the offline activation data is incorrect",
    0xC004F035: "The software Licensing Service reported that the computer could not be activated with a Volume license product key. Volume licensed systems require upgrading from a qualified operating system. Please contact your system administrator or use a different type of key",
    0xC004F038: "The software Licensing Service reported that the computer could not be activated. The count reported by your Key Management Service (KMS) is insufficient. Please contact your system administrator",
    0xC004F039: "The software Licensing Service reported that the computer could not be activated. The Key Management Service (KMS) is not enabled",
    0xC004F041: "The software Licensing Service determined that the Key Management Server (KMS) is not activated. KMS needs to be activated",
    0xC004F042: "The software Licensing Service determined that the specified Key Management Service (KMS) cannot be used",
    0xC004F050: "The Software Licensing Service reported that the product key is invalid",
    0xC004F051: "The software Licensing Service reported that the product key is blocked",
    0xC004F064: "The software Licensing Service reported that the non-Genuine grace period expired",
    0xC004F065: "The software Licensing Service reported that the application is running within the valid non-genuine period",
    0xC004F066: "The Software Licensing Service reported that the product SKU is not found",
    0xC004F06B: "The software Licensing Service determined that it is running in a virtual machine. The Key Management Service (KMS) is not supported in this mode",
    0xC004F074: "The Software Licensing Service reported that the computer could not be activated. No Key Management Service (KMS) could be contacted. Please see the Application Event Log for additional information.",
    0xC004F075: "The Software Licensing Service reported that the operation cannot be completed because the service is stopping",
    0xC004F304: "The Software Licensing Service reported that required license could not be found.",
    0xC004F305: "The Software Licensing Service reported that there are no certificates found in the system that could activate the product.",
    0xC004F30A: "The Software Licensing Service reported that the computer could not be activated. The certificate does not match the conditions in the license.",
    0xC004F30D: "The Software Licensing Service reported that the computer could not be activated. The thumbprint is invalid.",
    0xC004F30E: "The Software Licensing Service reported that the computer could not be activated. A certificate for the thumbprint could not be found.",
    0xC004F30F: "The Software Licensing Service reported that the computer could not be activated. The certificate does not match the criteria specified in the issuance license.",
    0xC004F310: "The Software Licensing Service reported that the computer could not be activated. The certificate does not match the trust point identifier (TPID) specified in the issuance license.",
    0xC004F311: "The Software Licensing Service reported that the computer could not be activated. A soft token cannot be used for activation.",
    0xC004F312: "The Software Licensing Service reported that the computer could not be activated. The certificate cannot be used because its private key is exportable.",
    0x5: "Access denied: the requested action requires elevated privileges",
    0x80070005: "Access denied: the requested action requires elevated privileges",
    0x80070057: "The parameter is incorrect",
    0x8007232A: "DNS server failure",
    0x8007232B: "DNS name does not exist",
    0x800706BA: "The RPC server is unavailable",
    0x8007251D: "No records found for DNS query",
}


class SLMgrError(Exception):
    """Base exception for slmgr errors"""

    def __init__(self, message: str, error_code: Optional[int] = None):
        self.message = message
        self.error_code = error_code
        super().__init__(self.message)


class OutputManager:
    """Manages output buffering similar to VBScript's LineOut/LineFlush"""

    def __init__(self) -> None:
        self.buffer: List[str] = []

    def line_out(self, text: str) -> None:
        """Add line to output buffer"""
        self.buffer.append(text)

    def line_flush(self, text: str = "") -> None:
        """Flush buffer to stdout with optional additional line"""
        if text:
            self.buffer.append(text)
        if self.buffer:
            print("\n".join(self.buffer))
        self.buffer = []

    def get_output(self) -> str:
        """Get current buffer content"""
        return "\n".join(self.buffer)


class WMIConnection:
    """Manages WMI connections to local or remote machines"""

    def __init__(self, computer: str = ".", username: str = "", password: str = ""):
        self.computer = computer
        self.username = username
        self.password = password
        self.is_remote = computer != "."
        self.wmi_service: Optional[WMIServiceProtocol] = None
        self.registry: Optional[RegistryProtocol] = None

    def connect(self, _output: OutputManager) -> None:
        """Establish WMI connection"""
        try:
            if not self.is_remote:
                # Local connection
                wmi_obj: object = wmi.WMI(computer=self.computer)
                self.wmi_service = cast(WMIServiceProtocol, wmi_obj)
                wmi_reg: object = wmi.WMI(computer=self.computer, namespace="default")
                self.registry = cast(RegistryProtocol, wmi_reg.StdRegProv)  # type: ignore[attr-defined]
            else:
                # Remote connection
                _: object = pythoncom.CoInitialize()

                # Connect to WMI
                if self.username and self.password:
                    wmi_obj = wmi.WMI(
                        computer=self.computer,
                        user=self.username,
                        password=self.password,
                        namespace="root\\cimv2",
                    )
                    self.wmi_service = cast(WMIServiceProtocol, wmi_obj)
                else:
                    wmi_obj = wmi.WMI(computer=self.computer)
                    self.wmi_service = cast(WMIServiceProtocol, wmi_obj)

                # Check version compatibility
                service = self.get_service_object("Version")
                version = service.Version
                if version and version.startswith(("6.0", "6.1")):
                    raise SLMgrError(
                        "The remote machine does not support this version of SLMgr.py"
                    )

                # Connect to registry
                if self.username and self.password:
                    locator: object = win32com.client.Dispatch(
                        "WbemScripting.SWbemLocator"
                    )
                    obj_server: object = locator.ConnectServer(  # type: ignore[attr-defined]
                        self.computer,
                        "root\\default:StdRegProv",
                        self.username,
                        self.password,
                    )
                    obj_server.Security_.ImpersonationLevel = 3  # type: ignore[attr-defined]
                    reg_prov: object = obj_server.Get("StdRegProv")  # type: ignore[attr-defined]
                    self.registry = cast(RegistryProtocol, reg_prov)
                else:
                    wmi_reg = wmi.WMI(computer=self.computer, namespace="default")
                    self.registry = cast(RegistryProtocol, wmi_reg.StdRegProv)  # type: ignore[attr-defined]

        except Exception as e:
            error_code_attr: object = getattr(e, "com_error", None)
            error_code: Optional[int] = cast(Optional[int], error_code_attr)
            if self.is_remote:
                error_str = str(error_code) if error_code else "unknown"
                msg = f"Error 0x{error_str} occurred in connecting to server {self.computer}."
            else:
                error_str = str(error_code) if error_code else "unknown"
                msg = f"Error 0x{error_str} occurred in connecting to the local WMI provider."
            raise SLMgrError(msg, error_code) from e

    def get_service_object(self, properties: str) -> SoftwareLicensingServiceProtocol:
        """Get SoftwareLicensingService object"""
        if not self.wmi_service:
            raise SLMgrError("WMI service not connected")

        query = f"SELECT {properties} FROM {SERVICE_CLASS}"
        results_obj: object = self.wmi_service.query(query)
        results: List[object] = cast(List[object], results_obj)

        if not results:
            raise SLMgrError("Failed to query Software Licensing Service")

        return cast(SoftwareLicensingServiceProtocol, results[0])

    def get_product_collection(
        self, select_clause: str, where_clause: str
    ) -> List[SoftwareLicensingProductProtocol]:
        """Get collection of SoftwareLicensingProduct objects"""
        if not self.wmi_service:
            raise SLMgrError("WMI service not connected")

        if where_clause and where_clause != EMPTY_WHERE_CLAUSE:
            query = f"SELECT {select_clause} FROM {PRODUCT_CLASS} WHERE {where_clause}"
        else:
            query = f"SELECT {select_clause} FROM {PRODUCT_CLASS}"

        results_obj: object = self.wmi_service.query(query)
        results_list: object = list(results_obj)  # type: ignore[call-overload]
        return cast(List[SoftwareLicensingProductProtocol], results_list)

    def get_product_object(
        self, select_clause: str, where_clause: str
    ) -> SoftwareLicensingProductProtocol:
        """Get single SoftwareLicensingProduct object"""
        products = self.get_product_collection(select_clause, where_clause)

        if len(products) == 0:
            raise SLMgrError("Error: product not found.", HR_SL_E_PKEY_NOT_INSTALLED)
        if len(products) > 1:
            raise SLMgrError("Invalid arguments", HR_INVALID_ARG)

        return products[0]


class RegistryManager:
    """Manages registry operations"""

    def __init__(self, connection: WMIConnection):
        self.connection = connection

    def set_string_value(
        self, hkey: int, key_path: str, value_name: str, value: str
    ) -> int:
        """Set a registry string value"""
        try:
            if not self.connection.is_remote:
                # Local registry access
                with winreg.OpenKey(hkey, key_path, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, value)
                return 0
            # Remote registry access via WMI
            if self.connection.registry:
                result_obj: object = self.connection.registry.SetStringValue(
                    hkey, key_path, value_name, value
                )
                result: int = cast(int, result_obj)
                return result if result else 0
            return 1
        except Exception:  # pylint: disable=broad-exception-caught
            return 1

    def delete_value(self, hkey: int, key_path: str, value_name: str) -> int:
        """Delete a registry value"""
        try:
            if not self.connection.is_remote:
                with winreg.OpenKey(hkey, key_path, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.DeleteValue(key, value_name)
                return 0
            if self.connection.registry:
                result_obj: object = self.connection.registry.DeleteValue(
                    hkey, key_path, value_name
                )
                result: int = cast(int, result_obj)
                return result if result else 0
            return 1
        except FileNotFoundError:
            return 2
        except Exception:  # pylint: disable=broad-exception-caught
            return 1

    def key_exists(self, hkey: int, key_path: str) -> bool:
        """Check if a registry key exists"""
        try:
            if not self.connection.is_remote:
                winreg.OpenKey(hkey, key_path, 0, winreg.KEY_QUERY_VALUE)
                return True
            if self.connection.registry:
                result_obj: object = self.connection.registry.CheckAccess(
                    hkey, key_path, 1
                )
                result: Tuple[int, ...] = cast(Tuple[int, ...], result_obj)
                return result[0] != 2 if result else False
            return False
        except Exception:  # pylint: disable=broad-exception-caught
            return False


def get_error_message(error_code: int) -> str:
    """Get friendly error message for error code"""
    if error_code in ERROR_MESSAGES:
        return ERROR_MESSAGES[error_code]
    return f"On a computer running Microsoft Windows non-core edition, run 'slui.exe 0x2a 0x{error_code:X}' to display the error text."


def show_error(message: str, error_code: Optional[int], description: str = "") -> None:
    """Display error message"""
    if error_code is not None:
        error_desc = get_error_message(error_code & 0xFFFFFFFF)

        if not description:
            description = error_desc
        elif description and not error_desc.startswith("On a computer"):
            description = f"{description} ({error_desc})"

        full_message = f"{message}0x{error_code & 0xFFFFFFFF:X} {description}"
    else:
        full_message = f"{message} {description}" if description else message

    print(full_message, file=sys.stderr)


def quit_with_error(message: str, error_code: int) -> None:
    """Display error and exit"""
    show_error(message, error_code)
    sys.exit(error_code if error_code > 0 else 1)


def is_kms_client(description: str) -> bool:
    """Check if license is KMS client"""
    return "VOLUME_KMSCLIENT" in description


def is_kms_server(description: str) -> bool:
    """Check if license is KMS server"""
    if is_kms_client(description):
        return False
    return "VOLUME_KMS" in description


def is_tbl(description: str) -> bool:
    """Check if license is time-based"""
    return "TIMEBASED_" in description


def is_avma(description: str) -> bool:
    """Check if license is AVMA"""
    return "VIRTUAL_MACHINE_ACTIVATION" in description


def is_mak(description: str) -> bool:
    """Check if license is MAK"""
    return "MAK" in description


def is_token_activated(product: SoftwareLicensingProductProtocol) -> bool:
    """Check if product is token-activated"""
    try:
        ilvid = product.TokenActivationILVID
        return ilvid is not None and ilvid != 0xFFFFFFFF
    except Exception:  # pylint: disable=broad-exception-caught
        return False


def is_ad_activated(product: SoftwareLicensingProductProtocol) -> bool:
    """Check if product is AD-activated"""
    try:
        return product.VLActivationType == 1
    except Exception:  # pylint: disable=broad-exception-caught
        return False


def get_is_primary_windows_sku(product: SoftwareLicensingProductProtocol) -> int:
    """
    Returns 0 if not primary SKU, 1 if it is, 2 if uncertain
    """
    is_primary = 0

    if product.ApplicationId.lower() == WINDOWS_APP_ID and product.PartialProductKey:
        try:
            is_addon = product.LicenseIsAddon
            is_primary = 0 if is_addon else 1
        except Exception:  # pylint: disable=broad-exception-caught
            # Property not available - check if KMS
            if is_kms_client(product.Description) or is_kms_server(product.Description):
                is_primary = 1
            else:
                is_primary = 2  # Indeterminate

    return is_primary


def check_product_for_command(
    product: SoftwareLicensingProductProtocol, activation_id: str
) -> bool:
    """Check if product matches the command criteria"""
    if (
        not activation_id
        and product.ApplicationId.lower() == WINDOWS_APP_ID
        and not product.LicenseIsAddon
    ):
        return True
    if product.ID.lower() == activation_id.lower():
        return True
    return False


def get_days_from_mins(minutes: int) -> int:
    """Convert minutes to days (ceiling operation)"""
    mins_in_day = 24 * 60
    return (minutes + mins_in_day - 1) // mins_in_day


def guid_to_string(byte_array: bytes) -> str:
    """Convert GUID byte array to string format"""
    if len(byte_array) < 16:
        return ""

    def hex_byte(b: int) -> str:
        return f"{b:02X}"

    s = "{"
    s += hex_byte(byte_array[3])
    s += hex_byte(byte_array[2])
    s += hex_byte(byte_array[1])
    s += hex_byte(byte_array[0])
    s += "-"
    s += hex_byte(byte_array[5])
    s += hex_byte(byte_array[4])
    s += "-"
    s += hex_byte(byte_array[7])
    s += hex_byte(byte_array[6])
    s += "-"
    s += hex_byte(byte_array[8])
    s += hex_byte(byte_array[9])
    s += "-"
    s += hex_byte(byte_array[10])
    s += hex_byte(byte_array[11])
    s += hex_byte(byte_array[12])
    s += hex_byte(byte_array[13])
    s += hex_byte(byte_array[14])
    s += hex_byte(byte_array[15])
    s += "}"
    return s


def wmi_date_to_datetime(wmi_date: str) -> Optional[datetime]:
    """Convert WMI datetime string to Python datetime"""
    if not wmi_date or wmi_date == "00000000000000.000000+000":
        return None

    try:
        # WMI format: YYYYMMDDHHMMSS.mmmmmm+UUU
        year = int(wmi_date[0:4])
        month = int(wmi_date[4:6])
        day = int(wmi_date[6:8])
        hour = int(wmi_date[8:10])
        minute = int(wmi_date[10:12])
        second = int(wmi_date[12:14])

        return datetime(year, month, day, hour, minute, second)
    except Exception:  # pylint: disable=broad-exception-caught
        return None


def output_indeterminate_operation_warning(
    product: SoftwareLicensingProductProtocol, output: OutputManager
) -> None:
    """Output warning when primary key determination is uncertain"""
    output.line_out(
        "Warning: SLMGR was not able to validate the current product key for Windows. Please upgrade to the latest service pack."
    )
    output.line_out(f"Processing the license for {product.Description} ({product.ID}).")


def fail_remote_exec(is_remote: bool) -> None:
    """Fail if command doesn't support remote execution"""
    if is_remote:
        raise SLMgrError(
            "This command of SLMgr.py is not supported for remote execution"
        )


# ============================================================================
# Command Handler Functions
# ============================================================================


def install_product_key(
    conn: WMIConnection, reg: RegistryManager, output: OutputManager, product_key: str
) -> None:
    """Install product key"""
    try:
        service = conn.get_service_object("Version")
        version = service.Version
        service.InstallProductKey(product_key)

        # Refresh license status
        service.RefreshLicenseStatus()

        # Check if this makes the system a KMS server
        is_kms = False
        products = conn.get_product_collection(
            PRODUCT_IS_PRIMARY_SKU_SELECT, PARTIAL_PRODUCT_KEY_NON_NULL_WHERE
        )

        for product in products:
            is_primary = get_is_primary_windows_sku(product)
            if is_primary == 2:
                output_indeterminate_operation_warning(product, output)

            if is_kms_server(product.Description):
                is_kms = True
                break

        if is_kms:
            # Set KMS version in registry
            ret = reg.set_string_value(
                HKEY_LOCAL_MACHINE, SL_KEY_PATH, "KeyManagementServiceVersion", version
            )
            if ret != 0:
                raise SLMgrError("Failed to set registry value", ret)

            if reg.key_exists(HKEY_LOCAL_MACHINE, SL_KEY_PATH_32):
                ret = reg.set_string_value(
                    HKEY_LOCAL_MACHINE,
                    SL_KEY_PATH_32,
                    "KeyManagementServiceVersion",
                    version,
                )
                if ret != 0:
                    raise SLMgrError("Failed to set registry value", ret)
        else:
            # Clear KMS version from registry
            reg.delete_value(
                HKEY_LOCAL_MACHINE, SL_KEY_PATH, "KeyManagementServiceVersion"
            )
            reg.delete_value(
                HKEY_LOCAL_MACHINE, SL_KEY_PATH_32, "KeyManagementServiceVersion"
            )

        output.line_out(f"Installed product key {product_key} successfully.")

    except Exception as e:
        raise SLMgrError(f"Error installing product key: {e}") from e


def uninstall_product_key(
    conn: WMIConnection,
    reg: RegistryManager,
    output: OutputManager,
    activation_id: str = "",
) -> None:
    """Uninstall product key"""
    activation_id = activation_id.lower()
    kms_server_found = False
    uninstall_done = False

    try:
        service = conn.get_service_object("Version")
        version = service.Version

        products = conn.get_product_collection(
            PRODUCT_IS_PRIMARY_SKU_SELECT + ", ProductKeyID",
            PARTIAL_PRODUCT_KEY_NON_NULL_WHERE,
        )

        for product in products:
            if check_product_for_command(product, activation_id):
                is_primary = get_is_primary_windows_sku(product)
                if not activation_id and is_primary == 2:
                    output_indeterminate_operation_warning(product, output)

                product.UninstallProductKey()
                service.RefreshLicenseStatus()

                if activation_id or is_primary == 1:
                    uninstall_done = True

                output.line_out("Uninstalled product key successfully.")

            elif is_kms_server(product.Description):
                kms_server_found = True

            if kms_server_found and uninstall_done:
                break

        if kms_server_found:
            ret = reg.set_string_value(
                HKEY_LOCAL_MACHINE, SL_KEY_PATH, "KeyManagementServiceVersion", version
            )
            if ret != 0:
                raise SLMgrError("Failed to set registry value", ret)

            ret = reg.set_string_value(
                HKEY_LOCAL_MACHINE,
                SL_KEY_PATH_32,
                "KeyManagementServiceVersion",
                version,
            )
            if ret != 0:
                raise SLMgrError("Failed to set registry value", ret)
        else:
            reg.delete_value(
                HKEY_LOCAL_MACHINE, SL_KEY_PATH, "KeyManagementServiceVersion"
            )
            reg.delete_value(
                HKEY_LOCAL_MACHINE, SL_KEY_PATH_32, "KeyManagementServiceVersion"
            )

        if not uninstall_done:
            output.line_out("Error: product key not found.")

    except Exception as e:
        raise SLMgrError(f"Error uninstalling product key: {e}") from e


def display_installation_id(
    conn: WMIConnection, output: OutputManager, activation_id: str = ""
) -> None:
    """Display installation ID for offline activation"""
    activation_id = activation_id.lower()
    found_at_least_one = False

    products = conn.get_product_collection(
        PRODUCT_IS_PRIMARY_SKU_SELECT + ", OfflineInstallationId",
        PARTIAL_PRODUCT_KEY_NON_NULL_WHERE,
    )

    for product in products:
        if check_product_for_command(product, activation_id):
            is_primary = get_is_primary_windows_sku(product)
            if not activation_id and is_primary == 2:
                output_indeterminate_operation_warning(product, output)

            output.line_out(f"Installation ID: {product.OfflineInstallationId}")
            found_at_least_one = True

            if activation_id or is_primary == 1:
                if found_at_least_one:
                    output.line_out("")
                    output.line_out(
                        "Product activation telephone numbers can be obtained by searching the phone.inf file for the appropriate phone number for your location/country. You can open the phone.inf file from a Command Prompt or the Start Menu by running: notepad %systemroot%\\system32\\sppui\\phone.inf"
                    )
                return

    if found_at_least_one:
        output.line_out("")
        output.line_out(
            "Product activation telephone numbers can be obtained by searching the phone.inf file for the appropriate phone number for your location/country. You can open the phone.inf file from a Command Prompt or the Start Menu by running: notepad %systemroot%\\system32\\sppui\\phone.inf"
        )
    else:
        output.line_out("Error: product not found.")


def activate_product(
    conn: WMIConnection, output: OutputManager, activation_id: str = ""
) -> None:
    """Activate Windows"""
    activation_id = activation_id.lower()
    found_at_least_one = False

    service = conn.get_service_object("Version")

    products = conn.get_product_collection(
        PRODUCT_IS_PRIMARY_SKU_SELECT + ", LicenseStatus, VLActivationTypeEnabled",
        PARTIAL_PRODUCT_KEY_NON_NULL_WHERE,
    )

    for product in products:
        if check_product_for_command(product, activation_id):
            is_primary = get_is_primary_windows_sku(product)
            if not activation_id and is_primary == 2:
                output_indeterminate_operation_warning(product, output)

            # Check if configured for token-based activation only
            if product.VLActivationTypeEnabled == 3:
                output.line_out(
                    "This system is configured for Token-based activation only. Use slmgr.py /fta to initiate Token-based activation, or slmgr.py /act-type to change the activation type setting."
                )
                return

            output.line_out(f"Activating {product.Name} ({product.ID}) ...")

            # Avoid using MAK count unless needed
            if not is_mak(product.Description) or product.LicenseStatus != 1:
                product.Activate()
                service.RefreshLicenseStatus()
                # Refresh product info
                product = conn.get_product_object(
                    PRODUCT_IS_PRIMARY_SKU_SELECT
                    + ", LicenseStatus, LicenseStatusReason",
                    f"ID = '{product.ID}'",
                )

            # Display status
            if product.LicenseStatus == 1:
                output.line_out("Product activated successfully.")
            elif product.LicenseStatus == 4:
                output.line_out(
                    "Error: The machine is running within the non-genuine grace period. Run 'slui.exe' to go online and make the machine genuine."
                )
            elif (
                product.LicenseStatus == 5
                and product.LicenseStatusReason == HR_SL_E_NOT_GENUINE
            ):
                output.line_out(
                    "Error: Windows is running within the non-genuine notification period. Run 'slui.exe' to go online and validate Windows."
                )
            elif product.LicenseStatus == 6:
                output.line_out("Product activated successfully.")
                output.line_out("License Status: Extended grace period")
            else:
                output.line_out("Error: Product activation failed.")

            found_at_least_one = True

            if activation_id or is_primary == 1:
                return

    if not found_at_least_one:
        output.line_out("Error: product not found.")


def phone_activate_product(
    conn: WMIConnection,
    output: OutputManager,
    confirmation_id: str,
    activation_id: str = "",
) -> None:
    """Activate product with user-provided Confirmation ID"""
    activation_id = activation_id.lower()
    found_at_least_one = False

    service = conn.get_service_object("Version")

    products = conn.get_product_collection(
        PRODUCT_IS_PRIMARY_SKU_SELECT
        + ", OfflineInstallationId, LicenseStatus, LicenseStatusReason",
        PARTIAL_PRODUCT_KEY_NON_NULL_WHERE,
    )

    for product in products:
        if check_product_for_command(product, activation_id):
            is_primary = get_is_primary_windows_sku(product)
            if not activation_id and is_primary == 2:
                output_indeterminate_operation_warning(product, output)

            product.DepositOfflineConfirmationId(
                product.OfflineInstallationId, confirmation_id
            )
            service.RefreshLicenseStatus()
            # Refresh product
            product = conn.get_product_object(
                PRODUCT_IS_PRIMARY_SKU_SELECT + ", LicenseStatus, LicenseStatusReason",
                f"ID = '{product.ID}'",
            )

            if product.LicenseStatus == 1:
                output.line_out(
                    f"Confirmation ID for product {product.ID} deposited successfully."
                )
            elif product.LicenseStatus == 4:
                output.line_out(
                    "Error: The machine is running within the non-genuine grace period. Run 'slui.exe' to go online and make the machine genuine."
                )
            elif (
                product.LicenseStatus == 5
                and product.LicenseStatusReason == HR_SL_E_NOT_GENUINE
            ):
                output.line_out(
                    "Error: Windows is running within the non-genuine notification period. Run 'slui.exe' to go online and validate Windows."
                )
            elif product.LicenseStatus == 6:
                output.line_out("Product activated successfully.")
                output.line_out("License Status: Extended grace period")
            else:
                output.line_out("Error: Product activation failed.")

            found_at_least_one = True

            if activation_id or is_primary == 1:
                return

    if not found_at_least_one:
        output.line_out("Error: product not found.")


def clear_product_key_from_registry(conn: WMIConnection, output: OutputManager) -> None:
    """Clear product key from the registry"""
    service = conn.get_service_object("Version")
    service.ClearProductKeyFromRegistry()
    output.line_out("Product key from registry cleared successfully.")


def install_license(
    conn: WMIConnection, output: OutputManager, license_file: str
) -> None:
    """Install license file"""
    try:
        with open(license_file, "r", encoding="utf-8") as f:
            license_data = f.read()
    except UnicodeDecodeError:
        # Try with different encodings
        try:
            with open(license_file, "r", encoding="utf-16") as f:
                license_data = f.read()
        except Exception:  # pylint: disable=broad-exception-caught
            with open(license_file, "r", encoding="ascii") as f:
                license_data = f.read()

    service = conn.get_service_object("Version")
    service.InstallLicense(license_data)
    output.line_out(f"License file {license_file} installed successfully.")


def reinstall_licenses(conn: WMIConnection, output: OutputManager) -> None:
    """Re-install system license files"""
    output.line_out("Re-installing license files ...")

    system_root = os.environ.get("SystemRoot", "C:\\Windows")
    spp_tokens_folder = os.path.join(system_root, "system32", "spp", "tokens")
    oem_folder = os.path.join(system_root, "system32", "oem")

    service = conn.get_service_object("Version")

    # Install from spp\tokens subfolders
    if os.path.exists(spp_tokens_folder):
        for root, _dirs, files in os.walk(spp_tokens_folder):
            for file in files:
                if file.endswith(".xrm-ms"):
                    license_path = os.path.join(root, file)
                    try:
                        with open(license_path, "r", encoding="utf-8") as f:
                            license_data = f.read()
                        service.InstallLicense(license_data)
                    except Exception:  # pylint: disable=broad-exception-caught
                        pass  # Continue with other files

    # Install from oem folder
    if os.path.exists(oem_folder):
        for root, _dirs, files in os.walk(oem_folder):
            for file in files:
                if file.endswith(".xrm-ms"):
                    license_path = os.path.join(root, file)
                    try:
                        with open(license_path, "r", encoding="utf-8") as f:
                            license_data = f.read()
                        service.InstallLicense(license_data)
                    except Exception:  # pylint: disable=broad-exception-caught
                        pass

    output.line_out("License files re-installed successfully.")


def rearm_windows(conn: WMIConnection, output: OutputManager) -> None:
    """Reset the licensing status of the machine"""
    service = conn.get_service_object("Version")
    service.ReArmWindows()
    output.line_out("Command completed successfully.")
    output.line_out("Please restart the system for the changes to take effect.")


def rearm_app(conn: WMIConnection, output: OutputManager, app_id: str) -> None:
    """Reset the licensing status of the given app"""
    service = conn.get_service_object("Version")
    service.ReArmApp(app_id)
    output.line_out("Command completed successfully.")


def rearm_sku(conn: WMIConnection, output: OutputManager, activation_id: str) -> None:
    """Reset the licensing status of the given sku"""
    activation_id = activation_id.lower()
    sku_found = False

    products = conn.get_product_collection("ID", f"ID = '{activation_id}'")

    for product in products:
        if product.ID.lower() == activation_id:
            sku_found = True
            product.ReArmsku()
            output.line_out("Command completed successfully.")
            break

    if not sku_found:
        output.line_out("Error: product not found.")


def _get_expiration_message(
    license_status: int,
    grace_remaining: int,
    description: str,
    ends: datetime,
) -> str:
    """Get expiration message for a given license status"""
    # Define grace period status messages
    grace_status_messages = {
        2: "Initial grace period",
        3: "Additional grace period",
        4: "Non-genuine grace period",
        6: "Extended grace period",
    }

    if license_status == 0:
        return "Unlicensed"

    if license_status == 1:
        if grace_remaining != 0:
            if is_tbl(description):
                return f"Timebased activation will expire {ends}"
            if is_avma(description):
                return f"Automatic VM activation will expire {ends}"
            return f"Volume activation will expire {ends}"
        return "The machine is permanently activated."

    if license_status in grace_status_messages:
        return f"{grace_status_messages[license_status]} ends {ends}"

    if license_status == 5:
        return "Windows is in Notification mode"

    return ""


def expiration_datetime(
    conn: WMIConnection, output: OutputManager, activation_id: str = ""
) -> None:
    """Display expiration date for current license state"""
    activation_id = activation_id.lower()
    found = False

    if activation_id:
        where_clause = (
            f"ID = '{activation_id}' AND {PARTIAL_PRODUCT_KEY_NON_NULL_WHERE}"
        )
    else:
        where_clause = f"ApplicationId = '{WINDOWS_APP_ID}' AND {PARTIAL_PRODUCT_KEY_NON_NULL_WHERE}"

    products = conn.get_product_collection(
        PRODUCT_IS_PRIMARY_SKU_SELECT + ", LicenseStatus, GracePeriodRemaining",
        where_clause,
    )

    for product in products:
        found = True
        is_primary = get_is_primary_windows_sku(product)
        if not activation_id and is_primary == 2:
            output_indeterminate_operation_warning(product, output)

        license_status = product.LicenseStatus
        grace_remaining = product.GracePeriodRemaining
        ends = datetime.now() + timedelta(minutes=grace_remaining)

        result = _get_expiration_message(
            license_status, grace_remaining, product.Description, ends
        )

        if result:
            output.line_out(f"{product.Name}:")
            output.line_out(f"    {result}")

    if not found:
        output.line_out("Error: product key not found.")


def _get_select_strings(verbose: bool) -> Tuple[str, str]:
    """Get service and product select strings based on verbose flag"""
    service_select = (
        "KeyManagementServiceListeningPort, KeyManagementServiceDnsPublishing, "
        "KeyManagementServiceLowPriority, ClientMachineId, KeyManagementServiceHostCaching, Version"
    )

    product_select = (
        PRODUCT_IS_PRIMARY_SKU_SELECT
        + ", "
        + "ProductKeyID, ProductKeyChannel, OfflineInstallationId, "
        "ProcessorURL, MachineURL, UseLicenseURL, ProductKeyURL, ValidationURL, "
        "GracePeriodRemaining, LicenseStatus, LicenseStatusReason, EvaluationEndDate, "
        "VLRenewalInterval, VLActivationInterval, KeyManagementServiceLookupDomain, KeyManagementServiceMachine, "
        "KeyManagementServicePort, DiscoveredKeyManagementServiceMachineName, "
        "DiscoveredKeyManagementServiceMachinePort, DiscoveredKeyManagementServiceMachineIpAddress, KeyManagementServiceProductKeyID, "
        "TokenActivationILID, TokenActivationILVID, TokenActivationGrantNumber, "
        "TokenActivationCertificateThumbprint, TokenActivationAdditionalInfo, TrustedTime, "
        "ADActivationObjectName, ADActivationObjectDN, ADActivationCsvlkPid, ADActivationCsvlkSkuId, VLActivationTypeEnabled, VLActivationType, "
        "IAID, AutomaticVMActivationHostMachineName, AutomaticVMActivationLastActivationTime, AutomaticVMActivationHostDigitalPid2"
    )

    if verbose:
        service_select = "RemainingWindowsReArmCount, " + service_select
        product_select = (
            "RemainingAppReArmCount, RemainingSkuReArmCount, " + product_select
        )

    return service_select, product_select


def _should_show_sku(
    product_iter: SoftwareLicensingProductProtocol, param: str
) -> Tuple[bool, bool]:
    """Determine if SKU should be shown and if using default"""
    is_primary = get_is_primary_windows_sku(product_iter)
    use_default = False
    show_sku_info = False

    if not param and is_primary in (1, 2):
        use_default = True
        show_sku_info = True

    if not param and product_iter.LicenseIsAddon and product_iter.PartialProductKey:
        show_sku_info = True

    if param == "all":
        show_sku_info = True

    if param == product_iter.ID.lower():
        show_sku_info = True

    return show_sku_info, use_default


def _display_product_basic_info(
    product: SoftwareLicensingProductProtocol, output: OutputManager
) -> None:
    """Display basic product information"""
    output.line_out("")
    output.line_out(f"Name: {product.Name}")
    output.line_out(f"Description: {product.Description}")

    token_additional_info: object = getattr(
        product, "TokenActivationAdditionalInfo", None
    )
    if token_additional_info:
        output.line_out(f"Additional Information: {token_additional_info}")


def _display_product_verbose_info(
    product: SoftwareLicensingProductProtocol,
    output: OutputManager,
    b_kms_client: bool,
    b_avma: bool,
) -> None:
    """Display verbose product information"""
    output.line_out(f"Activation ID: {product.ID}")
    output.line_out(f"Application ID: {product.ApplicationId}")
    output.line_out(f"Extended PID: {product.ProductKeyID}")
    output.line_out(f"Product Key Channel: {product.ProductKeyChannel}")
    output.line_out(f"Installation ID: {product.OfflineInstallationId}")

    if not b_kms_client and not b_avma:
        processor_url: object = getattr(product, "ProcessorURL", None)
        if processor_url:
            output.line_out(f"Processor Certificate URL: {processor_url}")
        machine_url: object = getattr(product, "MachineURL", None)
        if machine_url:
            output.line_out(f"Machine Certificate URL: {machine_url}")
        use_license_url: object = getattr(product, "UseLicenseURL", None)
        if use_license_url:
            output.line_out(f"Use License URL: {use_license_url}")
        product_key_url: object = getattr(product, "ProductKeyURL", None)
        if product_key_url:
            output.line_out(f"Product Key Certificate URL: {product_key_url}")
        validation_url: object = getattr(product, "ValidationURL", None)
        if validation_url:
            output.line_out(f"Validation URL: {validation_url}")


def _display_license_status(
    product: SoftwareLicensingProductProtocol,
    output: OutputManager,
    b_tbl: bool,
    b_avma: bool,
) -> None:
    """Display license status information"""
    ls = product.LicenseStatus

    if ls == 0:
        output.line_out("License Status: Unlicensed")
        return

    if ls == 1:
        output.line_out("License Status: Licensed")
        gp_min = product.GracePeriodRemaining
        if gp_min != 0:
            gp_day = get_days_from_mins(gp_min)
            if b_tbl:
                output.line_out(
                    f"Timebased activation expiration: {gp_min} minute(s) ({gp_day} day(s))"
                )
            elif b_avma:
                output.line_out(
                    f"Automatic VM activation expiration: {gp_min} minute(s) ({gp_day} day(s))"
                )
            else:
                output.line_out(
                    f"Volume activation expiration: {gp_min} minute(s) ({gp_day} day(s))"
                )
        return

    if ls == 2:
        output.line_out("License Status: Initial grace period")
        gp_min = product.GracePeriodRemaining
        gp_day = get_days_from_mins(gp_min)
        output.line_out(f"Time remaining: {gp_min} minute(s) ({gp_day} day(s))")
        return

    if ls == 3:
        output.line_out(
            "License Status: Additional grace period (KMS license expired or hardware out of tolerance)"
        )
        gp_min = product.GracePeriodRemaining
        gp_day = get_days_from_mins(gp_min)
        output.line_out(f"Time remaining: {gp_min} minute(s) ({gp_day} day(s))")
        return

    if ls == 4:
        output.line_out("License Status: Non-genuine grace period.")
        gp_min = product.GracePeriodRemaining
        gp_day = get_days_from_mins(gp_min)
        output.line_out(f"Time remaining: {gp_min} minute(s) ({gp_day} day(s))")
        return

    if ls == 5:
        output.line_out("License Status: Notification")
        err_code = product.LicenseStatusReason & 0xFFFFFFFF
        if err_code == HR_SL_E_NOT_GENUINE:
            output.line_out(f"Notification Reason: 0x{err_code:X} (non-genuine).")
        elif err_code == HR_SL_E_GRACE_TIME_EXPIRED:
            output.line_out(
                f"Notification Reason: 0x{err_code:X} (grace time expired)."
            )
        else:
            output.line_out(f"Notification Reason: 0x{err_code:X}.")
        return

    if ls == 6:
        output.line_out("License Status: Extended grace period")
        gp_min = product.GracePeriodRemaining
        gp_day = get_days_from_mins(gp_min)
        output.line_out(f"Time remaining: {gp_min} minute(s) ({gp_day} day(s))")
        return

    output.line_out("License Status: Unknown")


def _display_verbose_extras(
    product: SoftwareLicensingProductProtocol,
    service: SoftwareLicensingServiceProtocol,
    output: OutputManager,
) -> None:
    """Display verbose extra information"""
    eval_end = (
        wmi_date_to_datetime(product.EvaluationEndDate)
        if hasattr(product, "EvaluationEndDate")
        else None
    )
    if eval_end:
        output.line_out(f"Evaluation End Date: {eval_end}")

    if product.ApplicationId.lower() == WINDOWS_APP_ID:
        output.line_out(
            f"Remaining Windows rearm count: {service.RemainingWindowsReArmCount}"
        )
    else:
        output.line_out(f"Remaining App rearm count: {product.RemainingAppReArmCount}")
    output.line_out(f"Remaining SKU rearm count: {product.RemainingSkuReArmCount}")

    trusted_time = (
        wmi_date_to_datetime(product.TrustedTime)
        if hasattr(product, "TrustedTime")
        else None
    )
    if trusted_time:
        output.line_out(f"Trusted time: {trusted_time}")


def _display_activation_type(
    product: SoftwareLicensingProductProtocol, output: OutputManager
) -> None:
    """Display activation type configuration"""
    vl_type = (
        product.VLActivationTypeEnabled
        if hasattr(product, "VLActivationTypeEnabled")
        else 0
    )
    if vl_type == 1:
        output.line_out("Configured Activation Type: AD")
    elif vl_type == 2:
        output.line_out("Configured Activation Type: KMS")
    elif vl_type == 3:
        output.line_out("Configured Activation Type: Token")
    else:
        output.line_out("Configured Activation Type: All")


def _display_activation_info(
    product: SoftwareLicensingProductProtocol,
    service: SoftwareLicensingServiceProtocol,
    conn: WMIConnection,
    output: OutputManager,
    b_kms_client: bool,
    b_kms_server: bool,
    b_avma: bool,
    is_primary: int,
) -> None:
    """Display activation-specific information (KMS/AD/TKA/AVMA)"""
    if b_kms_client:
        _display_activation_type(product, output)

        if is_ad_activated(product):
            display_ad_client_info(service, product, output)
        elif is_token_activated(product):
            display_tka_client_info(service, product, output)
        elif product.LicenseStatus != 1:
            output.line_out(
                "Please use slmgr.py /ato to activate and update KMS client information in order to update values."
            )
        else:
            display_kms_client_info(service, product, output)

    if b_kms_server or is_primary in (1, 2):
        display_kms_info(service, product, conn, output)

    if b_avma:
        if hasattr(product, "IAID") and product.IAID:
            output.line_out(f"Guest IAID: {product.IAID}")
        else:
            output.line_out("Guest IAID: Not Available")

        display_avma_client_info(product, output)


def display_all_information(
    conn: WMIConnection, output: OutputManager, param: str = "", verbose: bool = False
) -> None:
    """Display license information (/dli and /dlv)"""
    param = param.lower()
    product_key_found = False

    service_select, product_select = _get_select_strings(verbose)
    service = conn.get_service_object(service_select)

    if verbose:
        output.line_out(f"Software licensing service version: {service.Version}")

    # Determine which products to display
    if param == "all":
        iter_select = product_select
        where_clause = EMPTY_WHERE_CLAUSE
    else:
        iter_select = PRODUCT_IS_PRIMARY_SKU_SELECT
        where_clause = EMPTY_WHERE_CLAUSE

    products_iter = conn.get_product_collection(iter_select, where_clause)

    for product_iter in products_iter:
        show_sku_info, use_default = _should_show_sku(product_iter, param)

        if show_sku_info:
            # Get full product object if not "all"
            if param == "all":
                product = product_iter
            else:
                product = conn.get_product_object(
                    product_select, f"ID = '{product_iter.ID}'"
                )

            is_primary = get_is_primary_windows_sku(product_iter)
            if use_default and is_primary == 2:
                output_indeterminate_operation_warning(product, output)

            product_key_found = True

            # Display basic info
            _display_product_basic_info(product, output)

            # Determine activation types
            b_kms_server = is_kms_server(product.Description)
            b_kms_client = is_kms_client(product.Description)
            b_tbl = is_tbl(product.Description)
            b_avma = is_avma(product.Description)

            # Display verbose info
            if verbose:
                _display_product_verbose_info(product, output, b_kms_client, b_avma)

            # Display partial product key
            if product.PartialProductKey:
                output.line_out(f"Partial Product Key: {product.PartialProductKey}")
            else:
                output.line_out("This license is not in use.")

            # Display license status
            _display_license_status(product, output, b_tbl, b_avma)

            # Display verbose extras
            if product.LicenseStatus != 0 and verbose:
                _display_verbose_extras(product, service, output)

            # Display activation-specific info
            _display_activation_info(
                product,
                service,
                conn,
                output,
                b_kms_client,
                b_kms_server,
                b_avma,
                is_primary,
            )

            # Break if specific product found
            if param != "all" and param == product.ID.lower():
                break

    if not product_key_found:
        output.line_out("Error: product key not found.")


def display_kms_client_info(
    service: SoftwareLicensingServiceProtocol,
    product: SoftwareLicensingProductProtocol,
    output: OutputManager,
) -> None:
    """Display KMS client information"""
    output.line_out("")
    output.line_out("Most recent activation information:")
    output.line_out("Key Management Service client information")
    output.line_out(f"    Client Machine ID (CMID): {service.ClientMachineID}")

    if (
        hasattr(product, "KeyManagementServiceLookupDomain")
        and product.KeyManagementServiceLookupDomain
    ):
        output.line_out(
            f"    Registered KMS SRV record lookup domain: {product.KeyManagementServiceLookupDomain}"
        )

    kms_machine = (
        product.KeyManagementServiceMachine
        if hasattr(product, "KeyManagementServiceMachine")
        else ""
    )
    if kms_machine:
        port = (
            product.KeyManagementServicePort
            if hasattr(product, "KeyManagementServicePort")
            else 0
        )
        if port == 0:
            port = DEFAULT_PORT
        output.line_out(f"    Registered KMS machine name: {kms_machine}:{port}")
    else:
        disc_name = (
            product.DiscoveredKeyManagementServiceMachineName
            if hasattr(product, "DiscoveredKeyManagementServiceMachineName")
            else ""
        )
        disc_port = (
            product.DiscoveredKeyManagementServiceMachinePort
            if hasattr(product, "DiscoveredKeyManagementServiceMachinePort")
            else 0
        )

        if disc_name and disc_port:
            output.line_out(f"    KMS machine name from DNS: {disc_name}:{disc_port}")
        else:
            output.line_out("    DNS auto-discovery: KMS name not available")

    ip_addr = (
        product.DiscoveredKeyManagementServiceMachineIpAddress
        if hasattr(product, "DiscoveredKeyManagementServiceMachineIpAddress")
        else ""
    )
    if ip_addr:
        output.line_out(f"    KMS machine IP address: {ip_addr}")
    else:
        output.line_out("    KMS machine IP address: not available")

    output.line_out(
        f"    KMS machine extended PID: {product.KeyManagementServiceProductKeyID}"
    )
    output.line_out(f"    Activation interval: {product.VLActivationInterval} minutes")
    output.line_out(f"    Renewal interval: {product.VLRenewalInterval} minutes")

    if service.KeyManagementServiceHostCaching:
        output.line_out("    KMS host caching is enabled")
    else:
        output.line_out("    KMS host caching is disabled")

    if (
        kms_machine
        and hasattr(product, "KeyManagementServiceLookupDomain")
        and product.KeyManagementServiceLookupDomain
    ):
        port = (
            product.KeyManagementServicePort
            if hasattr(product, "KeyManagementServicePort")
            else 0
        )
        if port == 0:
            port = DEFAULT_PORT
        output.line_out("")
        output.line_out(
            f"Warning: /skms setting overrides the /skms-domain setting. {kms_machine}:{port} will be used for activation."
        )


def display_kms_info(
    service: SoftwareLicensingServiceProtocol,
    product: SoftwareLicensingProductProtocol,
    conn: WMIConnection,
    output: OutputManager,
) -> None:
    """Display KMS server information"""
    # Get extended KMS properties
    kms_product = conn.get_product_object(
        "IsKeyManagementServiceMachine, KeyManagementServiceCurrentCount, "
        + "KeyManagementServiceTotalRequests, KeyManagementServiceFailedRequests, "
        + "KeyManagementServiceUnlicensedRequests, KeyManagementServiceLicensedRequests, "
        + "KeyManagementServiceOOBGraceRequests, KeyManagementServiceOOTGraceRequests, "
        + "KeyManagementServiceNonGenuineGraceRequests, KeyManagementServiceNotificationRequests",
        f"ID = '{product.ID}'",
    )

    if kms_product.IsKeyManagementServiceMachine > 0:
        output.line_out("")
        output.line_out("Key Management Service is enabled on this machine")
        output.line_out(
            f"    Current count: {kms_product.KeyManagementServiceCurrentCount}"
        )

        port = service.KeyManagementServiceListeningPort
        if port == 0:
            port = DEFAULT_PORT
        output.line_out(f"    Listening on Port: {port}")

        if service.KeyManagementServiceDnsPublishing:
            output.line_out("    DNS publishing enabled")
        else:
            output.line_out("    DNS publishing disabled")

        if service.KeyManagementServiceLowPriority:
            output.line_out("    KMS priority: Low")
        else:
            output.line_out("    KMS priority: Normal")

        # Display cumulative requests if available
        try:
            total_req = kms_product.KeyManagementServiceTotalRequests
            if total_req is not None:
                output.line_out("")
                output.line_out(
                    "Key Management Service cumulative requests received from clients"
                )
                output.line_out(
                    f"    Total requests received: {kms_product.KeyManagementServiceTotalRequests}"
                )
                output.line_out(
                    f"    Failed requests received: {kms_product.KeyManagementServiceFailedRequests}"
                )
                output.line_out(
                    f"    Requests with License Status Unlicensed: {kms_product.KeyManagementServiceUnlicensedRequests}"
                )
                output.line_out(
                    f"    Requests with License Status Licensed: {kms_product.KeyManagementServiceLicensedRequests}"
                )
                output.line_out(
                    f"    Requests with License Status Initial grace period: {kms_product.KeyManagementServiceOOBGraceRequests}"
                )
                output.line_out(
                    f"    Requests with License Status License expired or Hardware out of tolerance: {kms_product.KeyManagementServiceOOTGraceRequests}"
                )
                output.line_out(
                    f"    Requests with License Status Non-genuine grace period: {kms_product.KeyManagementServiceNonGenuineGraceRequests}"
                )
                output.line_out(
                    f"    Requests with License Status Notification: {kms_product.KeyManagementServiceNotificationRequests}"
                )
        except Exception:  # pylint: disable=broad-exception-caught
            pass


def display_ad_client_info(
    _service: SoftwareLicensingServiceProtocol,
    product: SoftwareLicensingProductProtocol,
    output: OutputManager,
) -> None:
    """Display AD activation client information"""
    output.line_out("")
    output.line_out("Most recent activation information:")
    output.line_out("AD Activation client information")
    output.line_out(f"    Activation Object name: {product.ADActivationObjectName}")
    output.line_out(f"    AO DN: {product.ADActivationObjectDN}")
    output.line_out(f"    AO extended PID: {product.ADActivationCsvlkPid}")
    output.line_out(f"    AO activation ID: {product.ADActivationCsvlkSkuId}")


def display_tka_client_info(
    _service: SoftwareLicensingServiceProtocol,
    product: SoftwareLicensingProductProtocol,
    output: OutputManager,
) -> None:
    """Display Token-based Activation client information"""
    output.line_out("")
    output.line_out("Most recent activation information:")
    output.line_out("Token-based Activation information")
    output.line_out(f"    License ID (ILID): {product.TokenActivationILID}")
    output.line_out(f"    Version ID (ILvID): {product.TokenActivationILVID}")
    output.line_out(f"    Grant Number: {product.TokenActivationGrantNumber}")
    output.line_out(
        f"    Certificate Thumbprint: {product.TokenActivationCertificateThumbprint}"
    )


def display_avma_client_info(
    product: SoftwareLicensingProductProtocol, output: OutputManager
) -> None:
    """Display AVMA client information"""
    host_name_obj: object = getattr(product, "AutomaticVMActivationHostMachineName", "")
    host_name: str = str(host_name_obj) if host_name_obj else ""

    act_time_obj: object = getattr(
        product, "AutomaticVMActivationLastActivationTime", ""
    )
    act_time: str = str(act_time_obj) if act_time_obj else ""

    host_pid_obj: object = getattr(product, "AutomaticVMActivationHostDigitalPid2", "")
    host_pid: str = str(host_pid_obj) if host_pid_obj else ""

    if host_name or act_time or host_pid:
        output.line_out("")
        output.line_out("Most recent activation information:")
        output.line_out("Automatic VM Activation client information")

        if host_name:
            output.line_out(f"    Host machine name: {host_name}")
        else:
            output.line_out("    Host machine name: Not Available")

        if act_time:
            time_obj = wmi_date_to_datetime(act_time)
            output.line_out(f"    Activation time: {time_obj}")
        else:
            output.line_out("    Activation time: Not Available")

        if host_pid:
            output.line_out(f"    Host Digital PID2: {host_pid}")
        else:
            output.line_out("    Host Digital PID2: Not Available")


# ============================================================================
# KMS Management Commands
# ============================================================================


def get_kms_client_object_by_activation_id(
    conn: WMIConnection, activation_id: str
) -> Union[SoftwareLicensingServiceProtocol, SoftwareLicensingProductProtocol]:
    """Get KMS client object (service or product) by activation ID"""
    activation_id = activation_id.lower()

    if not activation_id:
        return conn.get_service_object(f"Version, {KMS_CLIENT_LOOKUP_CLAUSE}")
    products = conn.get_product_collection(
        f"ID, {KMS_CLIENT_LOOKUP_CLAUSE}", EMPTY_WHERE_CLAUSE
    )
    for product in products:
        if product.ID.lower() == activation_id:
            return product
    raise SLMgrError(f"Error: Activation ID ({activation_id}) not found.")


def set_kms_machine_name(
    conn: WMIConnection,
    output: OutputManager,
    kms_name_port: str,
    activation_id: str = "",
) -> None:
    """Set the name and/or port for the KMS computer"""
    # Parse name and port
    kms_name = ""
    kms_port = ""

    # Check for IPv6 address
    if kms_name_port.startswith("["):
        bracket_end = kms_name_port.find("]")
        if bracket_end > 1:
            if len(kms_name_port) == bracket_end + 1:
                kms_name = kms_name_port
            else:
                kms_name = kms_name_port[: bracket_end + 1]
                kms_port = kms_name_port[bracket_end + 2 :]
    elif ":" in kms_name_port:
        # IPv4 address with port
        colon_pos = kms_name_port.find(":")
        kms_name = kms_name_port[:colon_pos]
        kms_port = kms_name_port[colon_pos + 1 :]
    else:
        # Just hostname/IP without port
        kms_name = kms_name_port

    target: Union[SoftwareLicensingServiceProtocol, SoftwareLicensingProductProtocol]
    target = get_kms_client_object_by_activation_id(conn, activation_id)

    if kms_name:
        target.SetKeyManagementServiceMachine(kms_name)

    if kms_port:
        target.SetKeyManagementServicePort(int(kms_port))
    else:
        target.ClearKeyManagementServicePort()

    output.line_out(
        f"Key Management Service machine name set to {kms_name_port} successfully."
    )

    lookup_domain: object = getattr(target, "KeyManagementServiceLookupDomain", None)
    if lookup_domain:
        output.line_out(
            f"Warning: /skms setting overrides the /skms-domain setting. {kms_name_port} will be used for activation."
        )


def clear_kms_name(
    conn: WMIConnection, output: OutputManager, activation_id: str = ""
) -> None:
    """Clear name of KMS computer used"""
    target: Union[SoftwareLicensingServiceProtocol, SoftwareLicensingProductProtocol]
    target = get_kms_client_object_by_activation_id(conn, activation_id)

    target.ClearKeyManagementServiceMachine()
    target.ClearKeyManagementServicePort()
    output.line_out("Key Management Service machine name cleared successfully.")

    lookup_domain: object = getattr(target, "KeyManagementServiceLookupDomain", None)
    if lookup_domain:
        output.line_out(
            f"Warning: /skms-domain setting is in effect. {lookup_domain} will be used for DNS SRV record lookup."
        )


def set_kms_lookup_domain(
    conn: WMIConnection, output: OutputManager, fqdn: str, activation_id: str = ""
) -> None:
    """Set the specific DNS domain in which all KMS SRV records can be found"""
    target: Union[SoftwareLicensingServiceProtocol, SoftwareLicensingProductProtocol]
    target = get_kms_client_object_by_activation_id(conn, activation_id)

    target.SetKeyManagementServiceLookupDomain(fqdn)
    output.line_out(f"Key Management Service lookup domain set to {fqdn} successfully.")

    kms_machine_obj: object = getattr(target, "KeyManagementServiceMachine", None)
    if kms_machine_obj:
        kms_machine: str = str(kms_machine_obj)
        port_obj: object = getattr(target, "KeyManagementServicePort", 0)
        try:
            port: int = int(port_obj) if port_obj else 0  # type: ignore[call-overload]
        except (ValueError, TypeError):
            port = 0
        if port == 0:
            port = DEFAULT_PORT
        output.line_out(
            f"Warning: /skms setting overrides the /skms-domain setting. {kms_machine}:{port} will be used for activation."
        )


def clear_kms_lookup_domain(
    conn: WMIConnection, output: OutputManager, activation_id: str = ""
) -> None:
    """Clear the specific DNS domain in which all KMS SRV records can be found"""
    target: Union[SoftwareLicensingServiceProtocol, SoftwareLicensingProductProtocol]
    target = get_kms_client_object_by_activation_id(conn, activation_id)

    target.ClearKeyManagementServiceLookupDomain()
    output.line_out("Key Management Service lookup domain cleared successfully.")

    kms_machine_obj: object = getattr(target, "KeyManagementServiceMachine", None)
    if kms_machine_obj:
        kms_machine: str = str(kms_machine_obj)
        port_obj: object = getattr(target, "KeyManagementServicePort", 0)
        try:
            port: int = int(port_obj) if port_obj else 0  # type: ignore[call-overload]
        except (ValueError, TypeError):
            port = 0
        if port == 0:
            port = DEFAULT_PORT
        output.line_out(
            f"Warning: /skms setting is in effect. {kms_machine}:{port} will be used for activation."
        )


def set_host_caching_disable(
    conn: WMIConnection, output: OutputManager, disable: bool
) -> None:
    """Enable or disable KMS host caching"""
    service = conn.get_service_object("Version")
    service.DisableKeyManagementServiceHostCaching(disable)

    if disable:
        output.line_out("KMS host caching is disabled")
    else:
        output.line_out("KMS host caching is enabled")


def set_activation_interval(
    conn: WMIConnection, output: OutputManager, interval: int
) -> None:
    """Set interval for unactivated clients to attempt KMS connection"""
    if interval < 0:
        raise SLMgrError("Error: The data is invalid")

    service = conn.get_service_object("Version")
    products = conn.get_product_collection(
        "ID, IsKeyManagementServiceMachine", PARTIAL_PRODUCT_KEY_NON_NULL_WHERE
    )

    kms_flag = False
    for product in products:
        if product.IsKeyManagementServiceMachine == 1:
            kms_flag = True
            service.SetVLActivationInterval(interval)
            output.line_out(
                f"Volume activation interval set to {interval} minutes successfully."
            )
            output.line_out(
                "Warning: a KMS reboot is needed for this setting to take effect."
            )
            break

    if not kms_flag:
        output.line_out(
            "Warning: Volume activation interval can only be set on a KMS machine that is also activated."
        )


def set_renewal_interval(
    conn: WMIConnection, output: OutputManager, interval: int
) -> None:
    """Set renewal interval for activated clients to attempt KMS connection"""
    if interval < 0:
        raise SLMgrError("Error: The data is invalid")

    service = conn.get_service_object("Version")
    products = conn.get_product_collection(
        "ID, IsKeyManagementServiceMachine", PARTIAL_PRODUCT_KEY_NON_NULL_WHERE
    )

    kms_flag = False
    for product in products:
        if product.IsKeyManagementServiceMachine:
            kms_flag = True
            service.SetVLRenewalInterval(interval)
            output.line_out(
                f"Volume renewal interval set to {interval} minutes successfully."
            )
            output.line_out(
                "Warning: a KMS reboot is needed for this setting to take effect."
            )
            break

    if not kms_flag:
        output.line_out(
            "Warning: Volume renewal interval can only be set on a KMS machine that is also activated."
        )


def set_kms_listen_port(conn: WMIConnection, output: OutputManager, port: int) -> None:
    """Set TCP port KMS will use to communicate with clients"""
    service = conn.get_service_object("Version")
    products = conn.get_product_collection(
        "ID, IsKeyManagementServiceMachine", PARTIAL_PRODUCT_KEY_NON_NULL_WHERE
    )

    kms_flag = False
    for product in products:
        if product.IsKeyManagementServiceMachine:
            kms_flag = True
            service.SetKeyManagementServiceListeningPort(port)
            output.line_out(f"KMS port set to {port} successfully.")
            output.line_out(
                "Warning: a KMS reboot is needed for this setting to take effect."
            )
            break

    if not kms_flag:
        output.line_out(
            "Warning: KMS port can only be set on a KMS machine that is also activated."
        )


def set_dns_publishing_disabled(
    conn: WMIConnection, output: OutputManager, disable: bool
) -> None:
    """Enable or disable DNS publishing by KMS"""
    service = conn.get_service_object("Version")
    products = conn.get_product_collection(
        "ID, IsKeyManagementServiceMachine", PARTIAL_PRODUCT_KEY_NON_NULL_WHERE
    )

    kms_flag = False
    for product in products:
        if product.IsKeyManagementServiceMachine:
            kms_flag = True
            service.DisableKeyManagementServiceDnsPublishing(disable)
            if disable:
                output.line_out("DNS publishing disabled")
            else:
                output.line_out("DNS publishing enabled")
            output.line_out(
                "Warning: a KMS reboot is needed for this setting to take effect."
            )
            break

    if not kms_flag:
        output.line_out(
            "Warning: DNS Publishing can only be set on a KMS machine that is also activated."
        )


def set_kms_low_priority(
    conn: WMIConnection, output: OutputManager, low_priority: bool
) -> None:
    """Set KMS priority to normal or low"""
    service = conn.get_service_object("Version")
    products = conn.get_product_collection(
        "ID, IsKeyManagementServiceMachine", PARTIAL_PRODUCT_KEY_NON_NULL_WHERE
    )

    kms_flag = False
    for product in products:
        if product.IsKeyManagementServiceMachine:
            kms_flag = True
            service.EnableKeyManagementServiceLowPriority(low_priority)
            if low_priority:
                output.line_out("KMS priority set to Low")
            else:
                output.line_out("KMS priority set to Normal")
            output.line_out(
                "Warning: a KMS reboot is needed for this setting to take effect."
            )
            break

    if not kms_flag:
        output.line_out(
            "Warning: Priority can only be set on a KMS machine that is also activated."
        )


def set_vl_activation_type(
    conn: WMIConnection,
    output: OutputManager,
    act_type: Optional[int],
    activation_id: str = "",
) -> None:
    """Set activation type to 1 (for AD) or 2 (for KMS) or 3 (for Token) or 0 (for all)"""
    if act_type is None:
        act_type = 0

    if act_type < 0 or act_type > 3:
        raise SLMgrError("Error: The data is invalid")

    target: Union[SoftwareLicensingServiceProtocol, SoftwareLicensingProductProtocol]
    target = get_kms_client_object_by_activation_id(conn, activation_id)

    if act_type != 0:
        target.SetVLActivationTypeEnabled(act_type)
    else:
        target.ClearVLActivationTypeEnabled()

    output.line_out("Volume activation type set successfully.")


# ============================================================================
# Token-based Activation Commands
# ============================================================================


def tka_list_ils(conn: WMIConnection, output: OutputManager) -> None:
    """List installed Token-based Activation Issuance Licenses"""
    output.line_out("Token-based Activation Issuance Licenses:")
    output.line_out("")

    query = f"SELECT * FROM {TKA_LICENSE_CLASS}"
    licenses_obj: object = conn.wmi_service.query(query) if conn.wmi_service else []
    licenses_list: object = list(licenses_obj)  # type: ignore[call-overload]
    licenses: List[TokenActivationLicenseProtocol] = cast(
        List[TokenActivationLicenseProtocol], licenses_list
    )

    n_listed = 0
    for license_obj in licenses:
        ilid_val: str = str(license_obj.ILID)
        ilvid_val: int = int(license_obj.ILVID)
        header = f"{ilid_val}    {ilvid_val}"
        output.line_out(header)
        output.line_out(f"    License ID (ILID): {ilid_val}")
        output.line_out(f"    Version ID (ILvID): {ilvid_val}")

        exp_date_str: object = getattr(license_obj, "ExpirationDate", None)
        if exp_date_str:
            exp_date = wmi_date_to_datetime(str(exp_date_str))
            if exp_date:
                output.line_out(f"    Valid to: {exp_date}")

        additional_info: object = getattr(license_obj, "AdditionalInfo", None)
        if additional_info:
            output.line_out(f"    Additional Information: {additional_info}")

        auth_status: object = getattr(license_obj, "AuthorizationStatus", 0)
        auth_status_int: int = int(str(auth_status)) if auth_status else 0
        if auth_status_int != 0:
            output.line_out(f"    Error: 0x{auth_status_int:X}")
        else:
            description: object = getattr(license_obj, "Description", None)
            if description:
                output.line_out(f"    Description: {description}")

        output.line_out("")
        n_listed += 1

    if n_listed == 0:
        output.line_out("No licenses found.")


def tka_remove_il(
    conn: WMIConnection, output: OutputManager, ilid: str, ilvid: str
) -> None:
    """Remove installed Token-based Activation Issuance License"""
    output.line_out("Removing Token-based Activation License ...")
    output.line_out("")

    ilvid_int = int(ilvid)

    query = f"SELECT * FROM {TKA_LICENSE_CLASS}"
    licenses_obj: object = conn.wmi_service.query(query) if conn.wmi_service else []
    licenses_list: object = list(licenses_obj)  # type: ignore[call-overload]
    licenses: List[TokenActivationLicenseProtocol] = cast(
        List[TokenActivationLicenseProtocol], licenses_list
    )

    n_removed = 0
    for license_obj in licenses:
        license_ilid: str = str(license_obj.ILID)
        license_ilvid: int = int(license_obj.ILVID)
        if ilid == license_ilid and ilvid_int == license_ilvid:
            license_obj.Uninstall()
            license_id: str = str(license_obj.ID)
            output.line_out(f"Removed license with SLID={license_id}.")
            n_removed += 1

    if n_removed == 0:
        output.line_out("No licenses found.")


def tka_list_certs(conn: WMIConnection, output: OutputManager) -> None:
    """List Token-based Activation Certificates"""
    try:
        signer: object = win32com.client.Dispatch("SPPWMI.SppWmiTokenActivationSigner")
    except Exception as e:
        raise SLMgrError(f"Error creating token activation signer: {e}") from e

    # Get Windows product
    product = conn.get_product_object(
        "ID, Name, ApplicationId, PartialProductKey, Description, LicenseIsAddon",
        f"ApplicationId = '{WINDOWS_APP_ID}' AND PartialProductKey <> NULL AND LicenseIsAddon = FALSE",
    )

    grants: object = product.GetTokenActivationGrants()
    thumbprints: object = signer.GetCertificateThumbprints(grants)  # type: ignore[attr-defined]
    for thumbprint_obj in thumbprints:  # type: ignore[attr-defined]
        thumbprint_str: str = str(thumbprint_obj)
        parts_obj: object = thumbprint_str.split("|")
        parts: List[str] = cast(List[str], parts_obj)
        if len(parts) >= 5:
            output.line_out("")
            output.line_out(f"Thumbprint: {parts[0]}")
            output.line_out(f"Subject: {parts[1]}")
            output.line_out(f"Issuer: {parts[2]}")
            output.line_out(f"Valid from: {parts[3]}")
            output.line_out(f"Valid to: {parts[4]}")


def tka_activate(
    conn: WMIConnection, output: OutputManager, thumbprint: str, pin: str = ""
) -> None:
    """Force Token-based Activation"""
    try:
        signer: object = win32com.client.Dispatch("SPPWMI.SppWmiTokenActivationSigner")
    except Exception as e:
        raise SLMgrError(f"Error creating token activation signer: {e}") from e

    product = conn.get_product_object(
        "ID, Name, ApplicationId, PartialProductKey, Description, LicenseIsAddon",
        f"ApplicationId = '{WINDOWS_APP_ID}' AND PartialProductKey <> NULL AND LicenseIsAddon = FALSE",
    )

    service = conn.get_service_object("Version")

    output.line_out(f"Activating {product.Name} ({product.ID}) ...")

    challenge_obj: object = product.GenerateTokenActivationChallenge()
    challenge: str = str(challenge_obj)
    auth_info1_obj: object = signer.Sign(challenge, thumbprint, pin)  # type: ignore[attr-defined]
    auth_info1: str = str(auth_info1_obj)
    auth_info2 = ""  # Returned by Sign in VBScript version

    product.DepositTokenActivationResponse(challenge, auth_info1, auth_info2)
    service.RefreshLicenseStatus()

    # Refresh product
    product = conn.get_product_object(
        "ID, Name, LicenseStatus, LicenseStatusReason", f"ID = '{product.ID}'"
    )

    # Display status
    if product.LicenseStatus == 1:
        output.line_out("Product activated successfully.")
    elif product.LicenseStatus == 4:
        output.line_out(
            "Error: The machine is running within the non-genuine grace period. Run 'slui.exe' to go online and make the machine genuine."
        )
    elif (
        product.LicenseStatus == 5
        and product.LicenseStatusReason == HR_SL_E_NOT_GENUINE
    ):
        output.line_out(
            "Error: Windows is running within the non-genuine notification period. Run 'slui.exe' to go online and validate Windows."
        )
    elif product.LicenseStatus == 6:
        output.line_out("Product activated successfully.")
        output.line_out("License Status: Extended grace period")
    else:
        output.line_out("Error: Product activation failed.")


# ============================================================================
# Active Directory Activation Commands
# ============================================================================


def ad_activate_online(
    conn: WMIConnection, output: OutputManager, product_key: str, ao_name: str = ""
) -> None:
    """Activate AD forest with user-provided product key"""
    fail_remote_exec(conn.is_remote)

    service = conn.get_service_object("Version")
    service.DoActiveDirectoryOnlineActivation(product_key, ao_name)
    output.line_out("Product activated successfully.")


def ad_get_iid(conn: WMIConnection, output: OutputManager, product_key: str) -> None:
    """Display Installation ID for AD forest"""
    fail_remote_exec(conn.is_remote)

    service = conn.get_service_object("Version")
    iid = service.GenerateActiveDirectoryOfflineActivationId(product_key)
    output.line_out(f"Installation ID: {iid}")
    output.line_out("")
    output.line_out(
        "Product activation telephone numbers can be obtained by searching the phone.inf file for the appropriate phone number for your location/country. You can open the phone.inf file from a Command Prompt or the Start Menu by running: notepad %systemroot%\\system32\\sppui\\phone.inf"
    )


def ad_activate_phone(
    conn: WMIConnection,
    output: OutputManager,
    product_key: str,
    cid: str,
    ao_name: str = "",
) -> None:
    """Activate AD forest with user-provided product key and Confirmation ID"""
    fail_remote_exec(conn.is_remote)

    service = conn.get_service_object("Version")
    service.DepositActiveDirectoryOfflineActivationConfirmation(
        product_key, cid, ao_name
    )
    output.line_out("Product activated successfully.")


def _get_ad_connection_info() -> Tuple[str, ADNamespaceProtocol, str]:
    """Get AD connection information (machine domain, namespace, config NC)"""
    ad_sys_info: object = win32com.client.Dispatch("ADSystemInfo")
    machine_domain: str = str(ad_sys_info.DomainDNSName) + "/"  # type: ignore[attr-defined]

    namespace: ADNamespaceProtocol = cast(
        ADNamespaceProtocol, win32com.client.GetObject(AD_LDAP_PROVIDER)
    )
    root_dse: ADRootDSEProtocol = cast(
        ADRootDSEProtocol,
        namespace.OpenDSObject(
            f"{AD_LDAP_PROVIDER_PREFIX}{machine_domain}{AD_ROOT_DSE}", "", "", 4
        ),
    )
    config_nc: str = str(root_dse.Get(AD_CONFIGURATION_NC))
    return machine_domain, namespace, config_nc


def _display_activation_object_info(
    child_obj: ADObjectProtocol, output: OutputManager
) -> None:
    """Display information for a single activation object"""
    child_obj.GetInfoEx(
        [
            AD_ACT_OBJ_DISPLAY_NAME,
            AD_ACT_OBJ_ATTRIB_DN,
            AD_ACT_OBJ_ATTRIB_SKU_ID,
            AD_ACT_OBJ_ATTRIB_PID,
        ],
        0,
    )

    display_name: str = str(child_obj.Get(AD_ACT_OBJ_DISPLAY_NAME))
    sku_id_bytes: object = child_obj.Get(AD_ACT_OBJ_ATTRIB_SKU_ID)
    sku_id = (
        guid_to_string(sku_id_bytes)
        if isinstance(sku_id_bytes, bytes)
        else str(sku_id_bytes)
    )
    partial_key: str = str(child_obj.Get(AD_ACT_OBJ_ATTRIB_PARTIAL_PKEY))
    pid: str = str(child_obj.Get(AD_ACT_OBJ_ATTRIB_PID))
    dn: str = str(child_obj.Get(AD_ACT_OBJ_ATTRIB_DN))

    output.line_out(f"    Activation Object name: {display_name}")
    output.line_out(f"        Activation ID: {sku_id}")
    output.line_out(f"        Partial Product Key: {partial_key}")
    output.line_out(f"        AO extended PID: {pid}")
    output.line_out(f"        AO DN: {dn}")
    output.line_out("")


def ad_list_activation_objects(conn: WMIConnection, output: OutputManager) -> None:
    """Display Activation Objects in AD"""
    fail_remote_exec(conn.is_remote)

    try:
        machine_domain, namespace, config_nc = _get_ad_connection_info()

        # Try to open activation objects container
        container_path = f"{AD_LDAP_PROVIDER_PREFIX}{machine_domain}{AD_ACT_OBJ_CONTAINER}{config_nc}"
        try:
            container: object = namespace.OpenDSObject(container_path, "", "", 4)
        except Exception:  # pylint: disable=broad-exception-caught
            output.line_out(
                "Active Directory-Based Activation is not supported in the current Active Directory schema."
            )
            return

        output.line_out("Activation Objects")

        found = False
        for child in container:  # type: ignore[attr-defined]
            child_obj: ADObjectProtocol = cast(ADObjectProtocol, child)
            if str(child_obj.Class) == AD_ACT_OBJ_CLASS:
                found = True
                _display_activation_object_info(child_obj, output)

        if not found:
            output.line_out("    No objects found")

    except Exception as e:
        raise SLMgrError(f"Error querying Active Directory: {e}") from e


def _construct_ao_dn(ao_name: str, config_nc: str) -> str:
    """Construct the distinguished name for an activation object"""
    if ",cn=" in ao_name.lower():
        return ao_name
    if ao_name.lower().startswith("cn="):
        return f"{ao_name},{AD_ACT_OBJ_CONTAINER}{config_nc}"
    return f"CN={ao_name},{AD_ACT_OBJ_CONTAINER}{config_nc}"


def ad_delete_activation_object(
    conn: WMIConnection, output: OutputManager, ao_name: str
) -> None:
    """Delete Activation Objects in AD"""
    fail_remote_exec(conn.is_remote)

    try:
        machine_domain, namespace, config_nc = _get_ad_connection_info()

        # Check if AD schema supports Activation Objects
        container_path = f"{AD_LDAP_PROVIDER_PREFIX}{machine_domain}{AD_ACT_OBJ_CONTAINER}{config_nc}"
        try:
            _container: object = namespace.OpenDSObject(container_path, "", "", 4)
        except Exception:  # pylint: disable=broad-exception-caught
            output.line_out(
                "Active Directory-Based Activation is not supported in the current Active Directory schema."
            )
            return

        # Construct and display DN
        dn = _construct_ao_dn(ao_name, config_nc)
        if ",cn=" not in ao_name.lower():
            output.line_out(f"    AO DN: {dn}")
            output.line_out("")

        # Delete object
        obj: ADObjectProtocol = cast(
            ADObjectProtocol,
            win32com.client.GetObject(f"{AD_LDAP_PROVIDER_PREFIX}{dn}"),
        )
        parent_obj: object = win32com.client.GetObject(obj.Parent)

        if str(obj.Class) == AD_ACT_OBJ_CLASS:
            parent_obj.Delete(str(obj.Class), str(obj.Name))  # type: ignore[attr-defined]

        output.line_out("Operation completed successfully.")

    except Exception as e:
        raise SLMgrError(f"Error deleting activation object: {e}") from e


# ============================================================================
# CLI Parser and Main Entry Point
# ============================================================================


def _print_usage_header() -> None:
    """Print usage header"""
    print("Windows Software Licensing Management Tool")
    print("Usage: slmgr.py [MachineName [User Password]] [<Option>]")
    print("       MachineName: Name of remote machine (default is local machine)")
    print("       User:        Account with required privilege on remote machine")
    print("       Password:    password for the previous account")
    print("")


def _print_global_options() -> None:
    """Print global options"""
    print("Global Options:")
    print("/ipk <Product Key>")
    print("    Install product key (replaces existing key)")
    print("/ato [Activation ID]")
    print("    Activate Windows")
    print("/dli [Activation ID | All]")
    print("    Display license information (default: current license)")
    print("/dlv [Activation ID | All]")
    print("    Display detailed license information (default: current license)")
    print("/xpr [Activation ID]")
    print("    Expiration date for current license state")
    print("")


def _print_advanced_options() -> None:
    """Print advanced options"""
    print("Advanced Options:")
    print("/cpky")
    print("    Clear product key from the registry (prevents disclosure attacks)")
    print("/ilc <License file>")
    print("    Install license")
    print("/rilc")
    print("    Re-install system license files")
    print("/rearm")
    print("    Reset the licensing status of the machine")
    print("/rearm-app <Application ID>")
    print("    Reset the licensing status of the given app")
    print("/rearm-sku <Activation ID>")
    print("    Reset the licensing status of the given sku")
    print("/upk [Activation ID]")
    print("    Uninstall product key")
    print("")
    print("/dti [Activation ID]")
    print("    Display Installation ID for offline activation")
    print("/atp <Confirmation ID> [Activation ID]")
    print("    Activate product with user-provided Confirmation ID")
    print("")


def _print_kms_client_options() -> None:
    """Print KMS client options"""
    print("Volume Licensing: Key Management Service (KMS) Client Options:")
    print("/skms <Name[:Port] | : port> [Activation ID]")
    print(
        "    Set the name and/or the port for the KMS computer this machine will use. IPv6 address must be specified in the format [hostname]:port"
    )
    print("/ckms [Activation ID]")
    print("    Clear name of KMS computer used (sets the port to the default)")
    print("/skms-domain <FQDN> [Activation ID]")
    print(
        "    Set the specific DNS domain in which all KMS SRV records can be found. This setting has no effect if the specific single KMS host is set via /skms option."
    )
    print("/ckms-domain [Activation ID]")
    print(
        "    Clear the specific DNS domain in which all KMS SRV records can be found. The specific KMS host will be used if set via /skms. Otherwise default KMS auto-discovery will be used."
    )
    print("/skhc")
    print("    Enable KMS host caching")
    print("/ckhc")
    print("    Disable KMS host caching")
    print("")


def _print_token_options() -> None:
    """Print token-based activation options"""
    print("Volume Licensing: Token-based Activation Options:")
    print("/lil")
    print("    List installed Token-based Activation Issuance Licenses")
    print("/ril <ILID> <ILvID>")
    print("    Remove installed Token-based Activation Issuance License")
    print("/ltc")
    print("    List Token-based Activation Certificates")
    print("/fta <Certificate Thumbprint> [<PIN>]")
    print("    Force Token-based Activation")
    print("")


def _print_kms_server_options() -> None:
    """Print KMS server options"""
    print("Volume Licensing: Key Management Service (KMS) Options:")
    print("/sprt <Port>")
    print("    Set TCP port KMS will use to communicate with clients")
    print("/sai <Activation Interval>")
    print(
        "    Set interval (minutes) for unactivated clients to attempt KMS connection. The activation interval must be between 15 minutes (min) and 30 days (max) although the default (2 hours) is recommended."
    )
    print("/sri <Renewal Interval>")
    print(
        "    Set renewal interval (minutes) for activated clients to attempt KMS connection. The renewal interval must be between 15 minutes (min) and 30 days (max) although the default (7 days) is recommended."
    )
    print("/sdns")
    print("    Enable DNS publishing by KMS (default)")
    print("/cdns")
    print("    Disable DNS publishing by KMS")
    print("/spri")
    print("    Set KMS priority to normal (default)")
    print("/cpri")
    print("    Set KMS priority to low")
    print("/act-type [Activation-Type] [Activation ID]")
    print(
        "    Set activation type to 1 (for AD) or 2 (for KMS) or 3 (for Token) or 0 (for all)."
    )
    print("")


def _print_ad_options() -> None:
    """Print AD activation options"""
    print("Volume Licensing: Active Directory (AD) Activation Options:")
    print("/ad-activation-online <Product Key> [Activation Object name]")
    print("    Activate AD (Active Directory) forest with user-provided product key")
    print("/ad-activation-get-iid <Product Key>")
    print("    Display Installation ID for AD (Active Directory) forest")
    print(
        "/ad-activation-apply-cid <Product Key> <Confirmation ID> [Activation Object name]"
    )
    print(
        "    Activate AD (Active Directory) forest with user-provided product key and Confirmation ID"
    )
    print("/ao-list")
    print("    Display Activation Objects in AD (Active Directory)")
    print("/del-ao <Activation Object DN | Activation Object RDN>")
    print(
        "    Delete Activation Objects in AD (Active Directory) for user-provided Activation Object"
    )


def display_usage() -> None:
    """Display usage information"""
    _print_usage_header()
    _print_global_options()
    _print_advanced_options()
    _print_kms_client_options()
    _print_token_options()
    _print_kms_server_options()
    _print_ad_options()


def parse_arguments() -> Tuple[str, str, str, List[str]]:
    """Parse command-line arguments to extract connection info and command"""
    args = sys.argv[1:]

    if not args:
        display_usage()
        sys.exit(1)

    computer = "."
    username = ""
    password = ""
    command_args = []

    # Check for remote connection parameters (first 3 args before / or -)
    remote_params = []

    for i, arg in enumerate(args):
        if i >= 3:
            break
        if arg.startswith("/") or arg.startswith("-"):
            break
        remote_params.append(arg)

    # Parse remote connection info
    if len(remote_params) == 3:
        computer = remote_params[0]
        username = remote_params[1]
        password = remote_params[2]
        command_args = args[3:]
    elif len(remote_params) == 1 and not (
        remote_params[0].startswith("/") or remote_params[0].startswith("-")
    ):
        computer = remote_params[0]
        command_args = args[1:]
    else:
        command_args = args

    return computer, username, password, command_args


class _CommandHandler:
    """Handler for command execution with validation"""

    def __init__(
        self, conn: WMIConnection, reg: RegistryManager, output: OutputManager
    ) -> None:
        self.conn = conn
        self.reg = reg
        self.output = output

    def handle_ipk(self, params: List[str]) -> None:
        """Install product key"""
        if len(params) < 1:
            raise SLMgrError("Error: option /ipk needs <Product Key>")
        install_product_key(self.conn, self.reg, self.output, params[0])

    def handle_upk(self, params: List[str]) -> None:
        """Uninstall product key"""
        activation_id = params[0] if params else ""
        uninstall_product_key(self.conn, self.reg, self.output, activation_id)

    def handle_dti(self, params: List[str]) -> None:
        """Display installation ID"""
        activation_id = params[0] if params else ""
        display_installation_id(self.conn, self.output, activation_id)

    def handle_ato(self, params: List[str]) -> None:
        """Activate product"""
        activation_id = params[0] if params else ""
        activate_product(self.conn, self.output, activation_id)

    def handle_atp(self, params: List[str]) -> None:
        """Phone activate product"""
        if len(params) < 1:
            raise SLMgrError("Error: option /atp needs <Confirmation ID>")
        cid = params[0]
        activation_id = params[1] if len(params) > 1 else ""
        phone_activate_product(self.conn, self.output, cid, activation_id)

    def handle_dli(self, params: List[str]) -> None:
        """Display license information"""
        param = params[0] if params else ""
        display_all_information(self.conn, self.output, param, False)

    def handle_dlv(self, params: List[str]) -> None:
        """Display verbose license information"""
        param = params[0] if params else ""
        display_all_information(self.conn, self.output, param, True)

    def handle_xpr(self, params: List[str]) -> None:
        """Display expiration datetime"""
        activation_id = params[0] if params else ""
        expiration_datetime(self.conn, self.output, activation_id)

    def handle_cpky(self, params: List[str]) -> None:  # pylint: disable=unused-argument
        """Clear product key from registry"""
        clear_product_key_from_registry(self.conn, self.output)

    def handle_ilc(self, params: List[str]) -> None:
        """Install license"""
        if len(params) < 1:
            raise SLMgrError("Error: option /ilc needs <License file>")
        install_license(self.conn, self.output, params[0])

    def handle_rilc(self, params: List[str]) -> None:  # pylint: disable=unused-argument
        """Reinstall licenses"""
        reinstall_licenses(self.conn, self.output)

    def handle_rearm(
        self, params: List[str]
    ) -> None:  # pylint: disable=unused-argument
        """Rearm Windows"""
        rearm_windows(self.conn, self.output)

    def handle_rearm_app(self, params: List[str]) -> None:
        """Rearm application"""
        if len(params) < 1:
            raise SLMgrError("Error: option /rearm-app needs <Application ID>")
        rearm_app(self.conn, self.output, params[0])

    def handle_rearm_sku(self, params: List[str]) -> None:
        """Rearm SKU"""
        if len(params) < 1:
            raise SLMgrError("Error: option /rearm-sku needs <Activation ID>")
        rearm_sku(self.conn, self.output, params[0])

    def handle_skms(self, params: List[str]) -> None:
        """Set KMS machine name"""
        if len(params) < 1:
            raise SLMgrError("Error: option /skms needs <Name[:Port] | : port>")
        kms_name = params[0]
        activation_id = params[1] if len(params) > 1 else ""
        set_kms_machine_name(self.conn, self.output, kms_name, activation_id)

    def handle_ckms(self, params: List[str]) -> None:
        """Clear KMS name"""
        activation_id = params[0] if params else ""
        clear_kms_name(self.conn, self.output, activation_id)

    def handle_skms_domain(self, params: List[str]) -> None:
        """Set KMS lookup domain"""
        if len(params) < 1:
            raise SLMgrError("Error: option /skms-domain needs <FQDN>")
        fqdn = params[0]
        activation_id = params[1] if len(params) > 1 else ""
        set_kms_lookup_domain(self.conn, self.output, fqdn, activation_id)

    def handle_ckms_domain(self, params: List[str]) -> None:
        """Clear KMS lookup domain"""
        activation_id = params[0] if params else ""
        clear_kms_lookup_domain(self.conn, self.output, activation_id)

    def handle_skhc(self, params: List[str]) -> None:  # pylint: disable=unused-argument
        """Enable KMS host caching"""
        set_host_caching_disable(self.conn, self.output, False)

    def handle_ckhc(self, params: List[str]) -> None:  # pylint: disable=unused-argument
        """Disable KMS host caching"""
        set_host_caching_disable(self.conn, self.output, True)

    def handle_sprt(self, params: List[str]) -> None:
        """Set KMS listen port"""
        if len(params) < 1:
            raise SLMgrError("Error: option /sprt needs <Port>")
        set_kms_listen_port(self.conn, self.output, int(params[0]))

    def handle_sai(self, params: List[str]) -> None:
        """Set activation interval"""
        if len(params) < 1:
            raise SLMgrError("Error: option /sai needs <Activation Interval>")
        set_activation_interval(self.conn, self.output, int(params[0]))

    def handle_sri(self, params: List[str]) -> None:
        """Set renewal interval"""
        if len(params) < 1:
            raise SLMgrError("Error: option /sri needs <Renewal Interval>")
        set_renewal_interval(self.conn, self.output, int(params[0]))

    def handle_sdns(self, params: List[str]) -> None:  # pylint: disable=unused-argument
        """Enable DNS publishing"""
        set_dns_publishing_disabled(self.conn, self.output, False)

    def handle_cdns(self, params: List[str]) -> None:  # pylint: disable=unused-argument
        """Disable DNS publishing"""
        set_dns_publishing_disabled(self.conn, self.output, True)

    def handle_spri(self, params: List[str]) -> None:  # pylint: disable=unused-argument
        """Set KMS priority to normal"""
        set_kms_low_priority(self.conn, self.output, False)

    def handle_cpri(self, params: List[str]) -> None:  # pylint: disable=unused-argument
        """Set KMS priority to low"""
        set_kms_low_priority(self.conn, self.output, True)

    def handle_act_type(self, params: List[str]) -> None:
        """Set VL activation type"""
        act_type = int(params[0]) if params and params[0].isdigit() else None
        activation_id = (
            params[1]
            if len(params) > 1
            else (params[0] if params and not params[0].isdigit() else "")
        )
        set_vl_activation_type(self.conn, self.output, act_type, activation_id)

    def handle_lil(self, params: List[str]) -> None:  # pylint: disable=unused-argument
        """List token-based activation licenses"""
        tka_list_ils(self.conn, self.output)

    def handle_ril(self, params: List[str]) -> None:
        """Remove token-based activation license"""
        if len(params) < 2:
            raise SLMgrError("Error: option /ril needs <ILID> <ILvID>")
        tka_remove_il(self.conn, self.output, params[0], params[1])

    def handle_ltc(self, params: List[str]) -> None:  # pylint: disable=unused-argument
        """List token-based activation certificates"""
        tka_list_certs(self.conn, self.output)

    def handle_fta(self, params: List[str]) -> None:
        """Force token-based activation"""
        if len(params) < 1:
            raise SLMgrError(
                "Error: option /fta needs <Certificate Thumbprint> [<PIN>]"
            )
        thumbprint = params[0]
        pin = params[1] if len(params) > 1 else ""
        tka_activate(self.conn, self.output, thumbprint, pin)

    def handle_ad_activation_online(self, params: List[str]) -> None:
        """AD activate online"""
        if len(params) < 1:
            raise SLMgrError("Error: option /ad-activation-online needs <Product Key>")
        product_key = params[0]
        ao_name = params[1] if len(params) > 1 else ""
        ad_activate_online(self.conn, self.output, product_key, ao_name)

    def handle_ad_activation_get_iid(self, params: List[str]) -> None:
        """Get AD installation ID"""
        if len(params) < 1:
            raise SLMgrError("Error: option /ad-activation-get-iid needs <Product Key>")
        ad_get_iid(self.conn, self.output, params[0])

    def handle_ad_activation_apply_cid(self, params: List[str]) -> None:
        """Apply AD confirmation ID"""
        if len(params) < 2:
            raise SLMgrError(
                "Error: option /ad-activation-apply-cid needs <Product Key> <Confirmation ID>"
            )
        product_key = params[0]
        cid = params[1]
        ao_name = params[2] if len(params) > 2 else ""
        ad_activate_phone(self.conn, self.output, product_key, cid, ao_name)

    def handle_ao_list(
        self, params: List[str]
    ) -> None:  # pylint: disable=unused-argument
        """List AD activation objects"""
        ad_list_activation_objects(self.conn, self.output)

    def handle_del_ao(self, params: List[str]) -> None:
        """Delete AD activation object"""
        if len(params) < 1:
            raise SLMgrError(
                "Error: option /del-ao needs <Activation Object DN | Activation Object RDN>"
            )
        ad_delete_activation_object(self.conn, self.output, params[0])


def _get_command_registry() -> Dict[str, Callable[[List[str]], None]]:
    """Build command registry mapping option names to handler methods"""
    # This will be populated at runtime with handler instance
    return {}


def execute_command(
    conn: WMIConnection,
    reg: RegistryManager,
    output: OutputManager,
    command_args: List[str],
) -> None:
    """Execute the specified command"""
    if not command_args:
        display_usage()
        return

    option = command_args[0].lstrip("/-").lower()
    params = command_args[1:]

    try:
        # Create handler instance
        handler = _CommandHandler(conn, reg, output)

        # Build command registry
        command_registry: Dict[str, Callable[[List[str]], None]] = {
            "ipk": handler.handle_ipk,
            "upk": handler.handle_upk,
            "dti": handler.handle_dti,
            "ato": handler.handle_ato,
            "atp": handler.handle_atp,
            "dli": handler.handle_dli,
            "dlv": handler.handle_dlv,
            "xpr": handler.handle_xpr,
            "cpky": handler.handle_cpky,
            "ilc": handler.handle_ilc,
            "rilc": handler.handle_rilc,
            "rearm": handler.handle_rearm,
            "rearm-app": handler.handle_rearm_app,
            "rearm-sku": handler.handle_rearm_sku,
            "skms": handler.handle_skms,
            "ckms": handler.handle_ckms,
            "skms-domain": handler.handle_skms_domain,
            "ckms-domain": handler.handle_ckms_domain,
            "skhc": handler.handle_skhc,
            "ckhc": handler.handle_ckhc,
            "sprt": handler.handle_sprt,
            "sai": handler.handle_sai,
            "sri": handler.handle_sri,
            "sdns": handler.handle_sdns,
            "cdns": handler.handle_cdns,
            "spri": handler.handle_spri,
            "cpri": handler.handle_cpri,
            "act-type": handler.handle_act_type,
            "lil": handler.handle_lil,
            "ril": handler.handle_ril,
            "ltc": handler.handle_ltc,
            "fta": handler.handle_fta,
            "ad-activation-online": handler.handle_ad_activation_online,
            "ad-activation-get-iid": handler.handle_ad_activation_get_iid,
            "ad-activation-apply-cid": handler.handle_ad_activation_apply_cid,
            "ao-list": handler.handle_ao_list,
            "del-ao": handler.handle_del_ao,
        }

        # Lookup and execute command
        command_func = command_registry.get(option)
        if command_func:
            command_func(params)
        else:
            print(f"Unrecognized option: {command_args[0]}", file=sys.stderr)
            print("", file=sys.stderr)
            display_usage()
            sys.exit(1)

    except SLMgrError as e:
        show_error("Error: ", e.error_code, str(e))
        sys.exit(e.error_code if e.error_code else 1)
    except Exception as e:  # pylint: disable=broad-exception-caught
        show_error("Error: ", None, str(e))
        sys.exit(1)


def main() -> None:
    """Main entry point"""
    try:
        # Parse arguments
        computer, username, password, command_args = parse_arguments()

        # Check for help
        if command_args and command_args[0].lstrip("/-").lower() in ["?", "h", "help"]:
            display_usage()
            return

        # Create output manager
        output = OutputManager()

        # Connect to WMI
        conn = WMIConnection(computer, username, password)
        conn.connect(output)

        # Create registry manager
        reg = RegistryManager(conn)

        # Execute command
        execute_command(conn, reg, output, command_args)

        # Flush output
        output.line_flush()

    except KeyboardInterrupt:
        print("\nOperation cancelled by user.", file=sys.stderr)
        sys.exit(1)
    except SLMgrError as e:
        show_error("Error: ", e.error_code, str(e))
        sys.exit(e.error_code if e.error_code else 1)
    except Exception as e:  # pylint: disable=broad-exception-caught
        show_error("Error: ", None, str(e))
        sys.exit(1)


if __name__ == "__main__":
    main()
