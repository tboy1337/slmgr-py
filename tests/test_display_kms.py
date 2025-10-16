"""Tests for display and KMS functions in slmgr.py"""

from typing import Any
from unittest.mock import Mock, patch

import pytest

import slmgr


class TestDisplayFunctions:
    """Test display information functions"""

    def test_display_all_information_basic(self) -> None:
        """Test display_all_information with basic product"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.KeyManagementServiceListeningPort = 1688
        mock_service.KeyManagementServiceDnsPublishing = True
        mock_service.KeyManagementServiceLowPriority = False
        mock_service.KeyManagementServiceHostCaching = True
        mock_service.ClientMachineId = "test-id"

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 1
        mock_product.GracePeriodRemaining = 0

        conn.wmi_service = Mock()
        call_count = [0]

        def query_side_effect(q: str) -> list:
            call_count[0] += 1
            if "SoftwareLicensingService" in q:
                return [mock_service]
            return [mock_product]

        conn.wmi_service.query.side_effect = query_side_effect

        # Mock the KMS product check
        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 0

        def get_product_side_effect(select: str, where: str = "") -> list:
            if "IsKeyManagementServiceMachine" in select:
                return [mock_kms_product]
            return [mock_product]

        conn.get_product_collection = Mock(side_effect=get_product_side_effect)

        slmgr.display_all_information(conn, output)

        result = output.get_output()
        assert "Test Product" in result
        assert "Licensed" in result

    def test_display_all_information_verbose(self) -> None:
        """Test display_all_information with verbose mode"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.KeyManagementServiceListeningPort = 1688
        mock_service.KeyManagementServiceDnsPublishing = True
        mock_service.KeyManagementServiceLowPriority = False
        mock_service.KeyManagementServiceHostCaching = True
        mock_service.ClientMachineId = "test-id"
        mock_service.RemainingWindowsReArmCount = 3

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 1
        mock_product.GracePeriodRemaining = 0
        mock_product.ProductKeyID = "extended-pid"
        mock_product.ProductKeyChannel = "Retail"
        mock_product.OfflineInstallationId = "123456"
        mock_product.EvaluationEndDate = "20241225143000.000000+000"
        mock_product.TrustedTime = "20231225143000.000000+000"
        mock_product.RemainingAppReArmCount = 2
        mock_product.RemainingSkuReArmCount = 1

        conn.wmi_service = Mock()
        call_count = [0]

        def query_side_effect(q: str) -> list:
            call_count[0] += 1
            if "SoftwareLicensingService" in q:
                return [mock_service]
            return [mock_product]

        conn.wmi_service.query.side_effect = query_side_effect

        # Mock the KMS product check
        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 0

        def get_product_side_effect(select: str, where: str = "") -> list:
            if "IsKeyManagementServiceMachine" in select:
                return [mock_kms_product]
            return [mock_product]

        conn.get_product_collection = Mock(side_effect=get_product_side_effect)

        slmgr.display_all_information(conn, output, "", True)

        result = output.get_output()
        assert "Software licensing service version" in result
        assert "Remaining Windows rearm count" in result

    def test_display_all_information_notification_status(self) -> None:
        """Test display_all_information with notification status"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.KeyManagementServiceListeningPort = 1688
        mock_service.KeyManagementServiceDnsPublishing = True
        mock_service.KeyManagementServiceLowPriority = False
        mock_service.KeyManagementServiceHostCaching = True
        mock_service.ClientMachineId = "test-id"

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 5
        mock_product.LicenseStatusReason = slmgr.HR_SL_E_NOT_GENUINE
        mock_product.GracePeriodRemaining = 0

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        # Mock the KMS product check
        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 0

        def get_product_side_effect(select: str, where: str = "") -> list:
            if "IsKeyManagementServiceMachine" in select:
                return [mock_kms_product]
            return [mock_product]

        conn.get_product_collection = Mock(side_effect=get_product_side_effect)

        slmgr.display_all_information(conn, output)

        result = output.get_output()
        assert "Notification" in result
        assert "non-genuine" in result.lower()

    def test_display_all_information_addon(self) -> None:
        """Test display_all_information with addon product"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.KeyManagementServiceListeningPort = 1688
        mock_service.KeyManagementServiceDnsPublishing = True
        mock_service.KeyManagementServiceLowPriority = False
        mock_service.KeyManagementServiceHostCaching = True
        mock_service.ClientMachineId = "test-id"

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Addon"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = True
        mock_product.LicenseStatus = 1
        mock_product.GracePeriodRemaining = 0

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        slmgr.display_all_information(conn, output)

        result = output.get_output()
        assert "Test Addon" in result

    def test_display_all_information_vl_activation_type_ad(self) -> None:
        """Test display_all_information with VL activation type AD"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.KeyManagementServiceListeningPort = 1688
        mock_service.RemainingWindowsReArmCount = 3

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "VOLUME_KMSCLIENT"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 1
        mock_product.GracePeriodRemaining = 0
        mock_product.VLActivationTypeEnabled = 1
        mock_product.RemainingSkuReArmCount = 1

        conn.wmi_service = Mock()

        def query_side_effect(q: str) -> list:
            if "SoftwareLicensingService" in q:
                return [mock_service]
            return [mock_product]

        conn.wmi_service.query.side_effect = query_side_effect

        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 0

        def get_product_side_effect(select: str, where: str = "") -> Any:
            if "IsKeyManagementServiceMachine" in select:
                return [mock_kms_product]
            return [mock_product]

        conn.get_product_collection = Mock(side_effect=get_product_side_effect)

        # Mock AD activation
        with patch("slmgr.is_ad_activated", return_value=True):
            with patch("slmgr.display_ad_client_info"):
                slmgr.display_all_information(conn, output, "", True)

        result = output.get_output()
        assert "Configured Activation Type: AD" in result

    def test_display_all_information_vl_activation_type_kms(self) -> None:
        """Test display_all_information with VL activation type KMS"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.KeyManagementServiceListeningPort = 1688
        mock_service.RemainingWindowsReArmCount = 3
        mock_service.ClientMachineID = "test-cmid"

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "VOLUME_KMSCLIENT"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 1
        mock_product.GracePeriodRemaining = 0
        mock_product.VLActivationTypeEnabled = 2
        mock_product.RemainingSkuReArmCount = 1

        conn.wmi_service = Mock()

        def query_side_effect(q: str) -> list:
            if "SoftwareLicensingService" in q:
                return [mock_service]
            return [mock_product]

        conn.wmi_service.query.side_effect = query_side_effect

        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 0

        def get_product_side_effect(select: str, where: str = "") -> Any:
            if "IsKeyManagementServiceMachine" in select:
                return [mock_kms_product]
            return [mock_product]

        conn.get_product_collection = Mock(side_effect=get_product_side_effect)

        # Mock not AD/token activated
        with patch("slmgr.is_ad_activated", return_value=False):
            with patch("slmgr.is_token_activated", return_value=False):
                with patch("slmgr.display_kms_client_info"):
                    slmgr.display_all_information(conn, output, "", True)

        result = output.get_output()
        assert "Configured Activation Type: KMS" in result

    def test_display_all_information_vl_activation_type_token(self) -> None:
        """Test display_all_information with VL activation type Token"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.KeyManagementServiceListeningPort = 1688
        mock_service.RemainingWindowsReArmCount = 3

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "VOLUME_KMSCLIENT"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 1
        mock_product.GracePeriodRemaining = 0
        mock_product.VLActivationTypeEnabled = 3
        mock_product.RemainingSkuReArmCount = 1

        conn.wmi_service = Mock()

        def query_side_effect(q: str) -> list:
            if "SoftwareLicensingService" in q:
                return [mock_service]
            return [mock_product]

        conn.wmi_service.query.side_effect = query_side_effect

        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 0

        def get_product_side_effect(select: str, where: str = "") -> Any:
            if "IsKeyManagementServiceMachine" in select:
                return [mock_kms_product]
            return [mock_product]

        conn.get_product_collection = Mock(side_effect=get_product_side_effect)

        # Mock token activation
        with patch("slmgr.is_ad_activated", return_value=False):
            with patch("slmgr.is_token_activated", return_value=True):
                with patch("slmgr.display_tka_client_info"):
                    slmgr.display_all_information(conn, output, "", True)

        result = output.get_output()
        assert "Configured Activation Type: Token" in result

    def test_display_all_information_vl_activation_type_all(self) -> None:
        """Test display_all_information with VL activation type All"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.KeyManagementServiceListeningPort = 1688
        mock_service.RemainingWindowsReArmCount = 3

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "VOLUME_KMSCLIENT"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 0
        mock_product.GracePeriodRemaining = 0
        mock_product.VLActivationTypeEnabled = 0
        mock_product.RemainingSkuReArmCount = 1

        conn.wmi_service = Mock()

        def query_side_effect(q: str) -> list:
            if "SoftwareLicensingService" in q:
                return [mock_service]
            return [mock_product]

        conn.wmi_service.query.side_effect = query_side_effect

        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 0

        def get_product_side_effect(select: str, where: str = "") -> Any:
            if "IsKeyManagementServiceMachine" in select:
                return [mock_kms_product]
            return [mock_product]

        conn.get_product_collection = Mock(side_effect=get_product_side_effect)

        # Mock not activated
        with patch("slmgr.is_ad_activated", return_value=False):
            with patch("slmgr.is_token_activated", return_value=False):
                slmgr.display_all_information(conn, output, "", True)

        result = output.get_output()
        assert "Configured Activation Type: All" in result
        assert "Please use slmgr.py /ato" in result

    def test_display_all_information_avma_with_iaid(self) -> None:
        """Test display_all_information with AVMA and IAID"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.KeyManagementServiceListeningPort = 1688
        mock_service.RemainingWindowsReArmCount = 3

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "VIRTUAL_MACHINE_ACTIVATION"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 1
        mock_product.GracePeriodRemaining = 1440
        mock_product.RemainingSkuReArmCount = 1
        mock_product.IAID = "12345678-1234-1234-1234-123456789012"

        conn.wmi_service = Mock()

        def query_side_effect(q: str) -> list:
            if "SoftwareLicensingService" in q:
                return [mock_service]
            return [mock_product]

        conn.wmi_service.query.side_effect = query_side_effect

        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 0

        def get_product_side_effect(select: str, where: str = "") -> Any:
            if "IsKeyManagementServiceMachine" in select:
                return [mock_kms_product]
            return [mock_product]

        conn.get_product_collection = Mock(side_effect=get_product_side_effect)

        with patch("slmgr.display_avma_client_info"):
            slmgr.display_all_information(conn, output, "", True)

        result = output.get_output()
        assert "Guest IAID:" in result
        assert "12345678-1234-1234-1234-123456789012" in result

    def test_display_all_information_avma_no_iaid(self) -> None:
        """Test display_all_information with AVMA but no IAID"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.KeyManagementServiceListeningPort = 1688
        mock_service.RemainingWindowsReArmCount = 3

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "VIRTUAL_MACHINE_ACTIVATION"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 1
        mock_product.GracePeriodRemaining = 1440
        mock_product.RemainingSkuReArmCount = 1
        mock_product.IAID = None

        conn.wmi_service = Mock()

        def query_side_effect(q: str) -> list:
            if "SoftwareLicensingService" in q:
                return [mock_service]
            return [mock_product]

        conn.wmi_service.query.side_effect = query_side_effect

        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 0

        def get_product_side_effect(select: str, where: str = "") -> Any:
            if "IsKeyManagementServiceMachine" in select:
                return [mock_kms_product]
            return [mock_product]

        conn.get_product_collection = Mock(side_effect=get_product_side_effect)

        with patch("slmgr.display_avma_client_info"):
            slmgr.display_all_information(conn, output, "", True)

        result = output.get_output()
        assert "Guest IAID: Not Available" in result

    def test_display_all_information_with_product_id(self) -> None:
        """Test display_all_information with specific product ID"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.KeyManagementServiceListeningPort = 1688

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 1
        mock_product.GracePeriodRemaining = 0
        mock_product.RemainingSkuReArmCount = 1

        conn.wmi_service = Mock()

        def query_side_effect(q: str) -> list:
            if "SoftwareLicensingService" in q:
                return [mock_service]
            return [mock_product]

        conn.wmi_service.query.side_effect = query_side_effect

        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 0

        def get_product_side_effect(select: str, where: str = "") -> Any:
            if "IsKeyManagementServiceMachine" in select:
                return [mock_kms_product]
            return [mock_product]

        conn.get_product_collection = Mock(side_effect=get_product_side_effect)

        slmgr.display_all_information(conn, output, "test-id", True)

        result = output.get_output()
        assert "Test Product" in result

    def test_display_all_information_no_product_key_found(self) -> None:
        """Test display_all_information when no product key found"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"

        conn.wmi_service = Mock()

        def query_side_effect(q: str) -> list:
            if "SoftwareLicensingService" in q:
                return [mock_service]
            return []

        conn.wmi_service.query.side_effect = query_side_effect

        conn.get_product_collection = Mock(return_value=[])

        slmgr.display_all_information(conn, output, "", True)

        result = output.get_output()
        assert "not found" in result.lower()

    def test_display_kms_client_info(self) -> None:
        """Test display_kms_client_info"""
        service = Mock()
        service.ClientMachineID = "client-id"
        service.KeyManagementServiceHostCaching = True

        product = Mock()
        product.KeyManagementServiceLookupDomain = "example.com"
        product.KeyManagementServiceMachine = "kms.example.com"
        product.KeyManagementServicePort = 1688
        product.DiscoveredKeyManagementServiceMachineName = "discovered.example.com"
        product.DiscoveredKeyManagementServiceMachinePort = 1688
        product.DiscoveredKeyManagementServiceMachineIpAddress = "192.168.1.1"
        product.KeyManagementServiceProductKeyID = "extended-pid"
        product.VLActivationInterval = 120
        product.VLRenewalInterval = 10080

        output = slmgr.OutputManager()
        slmgr.display_kms_client_info(service, product, output)

        result = output.get_output()
        assert "client-id" in result
        assert "kms.example.com" in result

    def test_display_kms_client_info_no_machine(self) -> None:
        """Test display_kms_client_info without registered machine"""
        service = Mock()
        service.ClientMachineID = "client-id"
        service.KeyManagementServiceHostCaching = False

        product = Mock()
        product.KeyManagementServiceMachine = ""
        product.DiscoveredKeyManagementServiceMachineName = "discovered.example.com"
        product.DiscoveredKeyManagementServiceMachinePort = 1688
        product.DiscoveredKeyManagementServiceMachineIpAddress = ""
        product.KeyManagementServiceProductKeyID = "extended-pid"
        product.VLActivationInterval = 120
        product.VLRenewalInterval = 10080

        output = slmgr.OutputManager()
        slmgr.display_kms_client_info(service, product, output)

        result = output.get_output()
        assert "caching is disabled" in result.lower()
        assert "discovered.example.com" in result

    def test_display_kms_info(self) -> None:
        """Test display_kms_info"""
        service = Mock()
        service.KeyManagementServiceListeningPort = 1688
        service.KeyManagementServiceDnsPublishing = True
        service.KeyManagementServiceLowPriority = False

        product = Mock()
        product.ID = "test-id"

        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 1
        mock_kms_product.KeyManagementServiceCurrentCount = 25
        mock_kms_product.KeyManagementServiceTotalRequests = 100
        mock_kms_product.KeyManagementServiceFailedRequests = 5
        mock_kms_product.KeyManagementServiceUnlicensedRequests = 10
        mock_kms_product.KeyManagementServiceLicensedRequests = 80
        mock_kms_product.KeyManagementServiceOOBGraceRequests = 5
        mock_kms_product.KeyManagementServiceOOTGraceRequests = 0
        mock_kms_product.KeyManagementServiceNonGenuineGraceRequests = 0
        mock_kms_product.KeyManagementServiceNotificationRequests = 0

        conn = slmgr.WMIConnection()
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_kms_product]

        output = slmgr.OutputManager()
        slmgr.display_kms_info(service, product, conn, output)

        result = output.get_output()
        assert "Key Management Service is enabled" in result
        assert "Current count: 25" in result
        assert "Total requests received: 100" in result

    def test_display_kms_info_not_kms(self) -> None:
        """Test display_kms_info when not a KMS machine"""
        service = Mock()
        product = Mock()
        product.ID = "test-id"

        mock_product = Mock()
        mock_product.IsKeyManagementServiceMachine = 0

        conn = slmgr.WMIConnection()
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        output = slmgr.OutputManager()
        slmgr.display_kms_info(service, product, conn, output)

        result = output.get_output()
        assert result == ""

    def test_display_ad_client_info(self) -> None:
        """Test display_ad_client_info"""
        service = Mock()
        product = Mock()
        product.ADActivationObjectName = "AO-Name"
        product.ADActivationObjectDN = "CN=AO,DC=example,DC=com"
        product.ADActivationCsvlkPid = "extended-pid"
        product.ADActivationCsvlkSkuId = "sku-id"

        output = slmgr.OutputManager()
        slmgr.display_ad_client_info(service, product, output)

        result = output.get_output()
        assert "AO-Name" in result
        assert "CN=AO,DC=example,DC=com" in result

    def test_display_tka_client_info(self) -> None:
        """Test display_tka_client_info"""
        service = Mock()
        product = Mock()
        product.TokenActivationILID = "ilid"
        product.TokenActivationILVID = 1
        product.TokenActivationGrantNumber = "grant-123"
        product.TokenActivationCertificateThumbprint = "thumbprint"

        output = slmgr.OutputManager()
        slmgr.display_tka_client_info(service, product, output)

        result = output.get_output()
        assert "Token-based Activation" in result
        assert "ilid" in result

    def test_display_avma_client_info(self) -> None:
        """Test display_avma_client_info"""
        product = Mock()
        product.AutomaticVMActivationHostMachineName = "host-machine"
        product.AutomaticVMActivationLastActivationTime = "20231225143000.000000+000"
        product.AutomaticVMActivationHostDigitalPid2 = "host-pid"

        output = slmgr.OutputManager()
        slmgr.display_avma_client_info(product, output)

        result = output.get_output()
        assert "Automatic VM Activation" in result
        assert "host-machine" in result

    def test_display_avma_client_info_no_data(self) -> None:
        """Test display_avma_client_info without data"""
        product = Mock()
        product.AutomaticVMActivationHostMachineName = ""
        product.AutomaticVMActivationLastActivationTime = ""
        product.AutomaticVMActivationHostDigitalPid2 = ""

        output = slmgr.OutputManager()
        slmgr.display_avma_client_info(product, output)

        result = output.get_output()
        assert result == ""


class TestKMSFunctions:
    """Test KMS management functions"""

    def test_get_kms_client_object_by_activation_id_service(self) -> None:
        """Test get_kms_client_object_by_activation_id returns service"""
        conn = slmgr.WMIConnection()
        mock_service = Mock()
        mock_service.Version = "10.0"
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        result = slmgr.get_kms_client_object_by_activation_id(conn, "")
        assert result == mock_service

    def test_get_kms_client_object_by_activation_id_product(self) -> None:
        """Test get_kms_client_object_by_activation_id returns product"""
        conn = slmgr.WMIConnection()
        mock_product = Mock()
        mock_product.ID = "test-id"
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        result = slmgr.get_kms_client_object_by_activation_id(conn, "test-id")
        assert result == mock_product

    def test_get_kms_client_object_by_activation_id_not_found(self) -> None:
        """Test get_kms_client_object_by_activation_id not found"""
        conn = slmgr.WMIConnection()
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = []

        with pytest.raises(slmgr.SLMgrError):
            slmgr.get_kms_client_object_by_activation_id(conn, "test-id")

    def test_set_kms_machine_name_ipv4(self) -> None:
        """Test set_kms_machine_name with IPv4"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.KeyManagementServiceMachine = ""
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.set_kms_machine_name(conn, output, "kms.example.com:1688")

        mock_service.SetKeyManagementServiceMachine.assert_called_once_with(
            "kms.example.com"
        )
        mock_service.SetKeyManagementServicePort.assert_called_once_with(1688)

    def test_set_kms_machine_name_ipv6(self) -> None:
        """Test set_kms_machine_name with IPv6"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.KeyManagementServiceMachine = ""
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.set_kms_machine_name(conn, output, "[::1]:1688")

        mock_service.SetKeyManagementServiceMachine.assert_called_once_with("[::1]")
        mock_service.SetKeyManagementServicePort.assert_called_once_with(1688)

    def test_set_kms_machine_name_no_port(self) -> None:
        """Test set_kms_machine_name without port"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.KeyManagementServiceMachine = ""
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.set_kms_machine_name(conn, output, "kms.example.com")

        mock_service.SetKeyManagementServiceMachine.assert_called_once_with(
            "kms.example.com"
        )
        mock_service.ClearKeyManagementServicePort.assert_called_once()

    def test_clear_kms_name(self) -> None:
        """Test clear_kms_name"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.KeyManagementServiceMachine = ""
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.clear_kms_name(conn, output)

        mock_service.ClearKeyManagementServiceMachine.assert_called_once()
        mock_service.ClearKeyManagementServicePort.assert_called_once()

    def test_set_kms_lookup_domain(self) -> None:
        """Test set_kms_lookup_domain"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.KeyManagementServiceMachine = ""
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.set_kms_lookup_domain(conn, output, "example.com")

        mock_service.SetKeyManagementServiceLookupDomain.assert_called_once_with(
            "example.com"
        )

    def test_clear_kms_lookup_domain(self) -> None:
        """Test clear_kms_lookup_domain"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.KeyManagementServiceMachine = ""
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.clear_kms_lookup_domain(conn, output)

        mock_service.ClearKeyManagementServiceLookupDomain.assert_called_once()

    def test_set_host_caching_disable(self) -> None:
        """Test set_host_caching_disable"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.set_host_caching_disable(conn, output, True)

        mock_service.DisableKeyManagementServiceHostCaching.assert_called_once_with(
            True
        )

    def test_set_activation_interval(self) -> None:
        """Test set_activation_interval"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_product = Mock()
        mock_product.IsKeyManagementServiceMachine = 1

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        slmgr.set_activation_interval(conn, output, 120)

        mock_service.SetVLActivationInterval.assert_called_once_with(120)

    def test_set_activation_interval_negative(self) -> None:
        """Test set_activation_interval with negative value"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        with pytest.raises(slmgr.SLMgrError):
            slmgr.set_activation_interval(conn, output, -1)

    def test_set_renewal_interval(self) -> None:
        """Test set_renewal_interval"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_product = Mock()
        mock_product.IsKeyManagementServiceMachine = 1

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        slmgr.set_renewal_interval(conn, output, 10080)

        mock_service.SetVLRenewalInterval.assert_called_once_with(10080)

    def test_set_kms_listen_port(self) -> None:
        """Test set_kms_listen_port"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_product = Mock()
        mock_product.IsKeyManagementServiceMachine = 1

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        slmgr.set_kms_listen_port(conn, output, 1688)

        mock_service.SetKeyManagementServiceListeningPort.assert_called_once_with(1688)

    def test_set_dns_publishing_disabled(self) -> None:
        """Test set_dns_publishing_disabled"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_product = Mock()
        mock_product.IsKeyManagementServiceMachine = 1

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        slmgr.set_dns_publishing_disabled(conn, output, False)

        mock_service.DisableKeyManagementServiceDnsPublishing.assert_called_once_with(
            False
        )

    def test_set_kms_low_priority(self) -> None:
        """Test set_kms_low_priority"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_product = Mock()
        mock_product.IsKeyManagementServiceMachine = 1

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        slmgr.set_kms_low_priority(conn, output, True)

        mock_service.EnableKeyManagementServiceLowPriority.assert_called_once_with(True)

    def test_set_vl_activation_type(self) -> None:
        """Test set_vl_activation_type"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.set_vl_activation_type(conn, output, 2)

        mock_service.SetVLActivationTypeEnabled.assert_called_once_with(2)

    def test_set_vl_activation_type_invalid(self) -> None:
        """Test set_vl_activation_type with invalid type"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        with pytest.raises(slmgr.SLMgrError):
            slmgr.set_vl_activation_type(conn, output, 5)

    def test_display_all_information_verbose_licensed_with_grace_period(self) -> None:
        """Test display_all_information verbose with licensed status and grace period"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.RemainingWindowsReArmCount = 3
        mock_service.RemainingAppReArmCount = 5

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.ApplicationID = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 1  # Licensed
        mock_product.GracePeriodRemaining = 43200  # 30 days
        mock_product.EvaluationEndDate = "20251116000000.000000-000"
        mock_product.TrustedTime = "20251016000000.000000-000"
        mock_product.IsKeyManagementServiceMachine = 0

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        # Mock the KMS product check
        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 0

        def get_product_side_effect(select: str, where: str = "") -> list:
            if "IsKeyManagementServiceMachine" in select:
                return [mock_kms_product]
            return [mock_product]

        conn.get_product_collection = Mock(side_effect=get_product_side_effect)
        conn.get_product_object = Mock(return_value=mock_product)

        slmgr.display_all_information(conn, output, "", True)

        result = output.get_output()
        assert "Licensed" in result
        assert "Volume activation expiration" in result

    def test_display_all_information_verbose_initial_grace(self) -> None:
        """Test display_all_information verbose with initial grace period"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.RemainingWindowsReArmCount = 3
        mock_service.RemainingAppReArmCount = 5

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.ApplicationID = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 2  # Initial grace period
        mock_product.GracePeriodRemaining = 43200
        mock_product.EvaluationEndDate = "20251116000000.000000-000"
        mock_product.TrustedTime = "20251016000000.000000-000"
        mock_product.IsKeyManagementServiceMachine = 0

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        # Mock the KMS product check
        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 0

        def get_product_side_effect(select: str, where: str = "") -> list:
            if "IsKeyManagementServiceMachine" in select:
                return [mock_kms_product]
            return [mock_product]

        conn.get_product_collection = Mock(side_effect=get_product_side_effect)
        conn.get_product_object = Mock(return_value=mock_product)

        slmgr.display_all_information(conn, output, "", True)

        result = output.get_output()
        assert "Initial grace period" in result

    def test_display_all_information_verbose_additional_grace(self) -> None:
        """Test display_all_information verbose with additional grace period"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.RemainingWindowsReArmCount = 3
        mock_service.RemainingAppReArmCount = 5

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.ApplicationID = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 3  # Additional grace period
        mock_product.GracePeriodRemaining = 43200
        mock_product.EvaluationEndDate = "20251116000000.000000-000"
        mock_product.TrustedTime = "20251016000000.000000-000"
        mock_product.IsKeyManagementServiceMachine = 0

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        # Mock the KMS product check
        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 0

        def get_product_side_effect(select: str, where: str = "") -> list:
            if "IsKeyManagementServiceMachine" in select:
                return [mock_kms_product]
            return [mock_product]

        conn.get_product_collection = Mock(side_effect=get_product_side_effect)
        conn.get_product_object = Mock(return_value=mock_product)

        slmgr.display_all_information(conn, output, "", True)

        result = output.get_output()
        assert "Additional grace period" in result

    def test_display_all_information_verbose_non_genuine_grace(self) -> None:
        """Test display_all_information verbose with non-genuine grace period"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.RemainingWindowsReArmCount = 3
        mock_service.RemainingAppReArmCount = 5

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.ApplicationID = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 4  # Non-genuine grace period
        mock_product.GracePeriodRemaining = 43200
        mock_product.EvaluationEndDate = "20251116000000.000000-000"
        mock_product.TrustedTime = "20251016000000.000000-000"
        mock_product.IsKeyManagementServiceMachine = 0

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        # Mock the KMS product check
        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 0

        def get_product_side_effect(select: str, where: str = "") -> list:
            if "IsKeyManagementServiceMachine" in select:
                return [mock_kms_product]
            return [mock_product]

        conn.get_product_collection = Mock(side_effect=get_product_side_effect)
        conn.get_product_object = Mock(return_value=mock_product)

        slmgr.display_all_information(conn, output, "", True)

        result = output.get_output()
        assert "Non-genuine grace period" in result

    def test_display_all_information_verbose_notification_grace_time_expired(
        self,
    ) -> None:
        """Test display_all_information verbose with notification status (grace time expired)"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.RemainingWindowsReArmCount = 3
        mock_service.RemainingAppReArmCount = 5

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.ApplicationID = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 5  # Notification
        mock_product.LicenseStatusReason = slmgr.HR_SL_E_GRACE_TIME_EXPIRED
        mock_product.GracePeriodRemaining = 0
        mock_product.EvaluationEndDate = "20251116000000.000000-000"
        mock_product.TrustedTime = "20251016000000.000000-000"
        mock_product.IsKeyManagementServiceMachine = 0

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        # Mock the KMS product check
        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 0

        def get_product_side_effect(select: str, where: str = "") -> list:
            if "IsKeyManagementServiceMachine" in select:
                return [mock_kms_product]
            return [mock_product]

        conn.get_product_collection = Mock(side_effect=get_product_side_effect)
        conn.get_product_object = Mock(return_value=mock_product)

        slmgr.display_all_information(conn, output, "", True)

        result = output.get_output()
        assert "Notification" in result
        assert "grace time expired" in result.lower()

    def test_display_all_information_verbose_extended_grace(self) -> None:
        """Test display_all_information verbose with extended grace period"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.RemainingWindowsReArmCount = 3
        mock_service.RemainingAppReArmCount = 5

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.ApplicationID = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 6  # Extended grace period
        mock_product.GracePeriodRemaining = 43200
        mock_product.EvaluationEndDate = "20251116000000.000000-000"
        mock_product.TrustedTime = "20251016000000.000000-000"
        mock_product.IsKeyManagementServiceMachine = 0

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        # Mock the KMS product check
        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 0

        def get_product_side_effect(select: str, where: str = "") -> list:
            if "IsKeyManagementServiceMachine" in select:
                return [mock_kms_product]
            return [mock_product]

        conn.get_product_collection = Mock(side_effect=get_product_side_effect)
        conn.get_product_object = Mock(return_value=mock_product)

        slmgr.display_all_information(conn, output, "", True)

        result = output.get_output()
        assert "Extended grace period" in result


class TestKMSManagementFunctions:
    """Test KMS management functions"""
