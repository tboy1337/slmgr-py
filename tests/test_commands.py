"""Tests for command handler functions in slmgr.py"""

from typing import Any
from unittest.mock import MagicMock, Mock, mock_open, patch

import pytest

import slmgr


class TestCommandHandlers:
    """Test command handler functions"""

    def test_install_product_key(self) -> None:
        """Test install_product_key"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        # Mock no KMS products
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else []
        )

        reg.set_string_value = Mock(return_value=0)
        reg.delete_value = Mock(return_value=0)

        slmgr.install_product_key(conn, reg, output, "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX")

        result = output.get_output()
        assert "successfully" in result.lower()

    def test_install_product_key_kms_server(self) -> None:
        """Test install_product_key with KMS server"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"

        mock_product = Mock()
        mock_product.Description = "VOLUME_KMS"
        mock_product.PartialProductKey = "XXXXX"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.LicenseIsAddon = False

        conn.wmi_service = Mock()
        call_count = [0]

        def query_side_effect(q: str) -> list:
            call_count[0] += 1
            if "SoftwareLicensingService" in q:
                return [mock_service]
            return [mock_product]

        conn.wmi_service.query.side_effect = query_side_effect

        reg.set_string_value = Mock(return_value=0)
        reg.key_exists = Mock(return_value=False)

        slmgr.install_product_key(conn, reg, output, "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX")

        assert reg.set_string_value.called

    def test_install_product_key_registry_error(self) -> None:
        """Test install_product_key with registry error"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"

        mock_product = Mock()
        mock_product.Description = "VOLUME_KMS"
        mock_product.PartialProductKey = "XXXXX"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.LicenseIsAddon = False

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        reg.set_string_value = Mock(return_value=1)

        with pytest.raises(slmgr.SLMgrError):
            slmgr.install_product_key(
                conn, reg, output, "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"
            )

    def test_uninstall_product_key(self) -> None:
        """Test uninstall_product_key"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.Description = "Test"

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        reg.set_string_value = Mock(return_value=0)
        reg.delete_value = Mock(return_value=0)

        slmgr.uninstall_product_key(conn, reg, output)

        mock_product.UninstallProductKey.assert_called_once()

    def test_uninstall_product_key_not_found(self) -> None:
        """Test uninstall_product_key when not found"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else []
        )

        reg.delete_value = Mock(return_value=0)

        slmgr.uninstall_product_key(conn, reg, output)

        result = output.get_output()
        assert "not found" in result.lower()

    def test_uninstall_product_key_kms_server_registry_error(self) -> None:
        """Test uninstall_product_key with KMS server and registry error"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"

        mock_product_kms = Mock()
        mock_product_kms.ID = "kms-id"
        mock_product_kms.Description = "VOLUME_KMS"
        mock_product_kms.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product_kms.PartialProductKey = "XXXXX"

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.Description = "Test"

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service]
            if "SoftwareLicensingService" in q
            else [mock_product, mock_product_kms]
        )

        reg.set_string_value = Mock(return_value=1)

        with pytest.raises(slmgr.SLMgrError) as exc_info:
            slmgr.uninstall_product_key(conn, reg, output)
        assert "Failed to set registry value" in str(exc_info.value)

    def test_uninstall_product_key_kms_server_registry_error_32bit(self) -> None:
        """Test uninstall_product_key with KMS server and 32-bit registry error"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"

        mock_product_kms = Mock()
        mock_product_kms.ID = "kms-id"
        mock_product_kms.Description = "VOLUME_KMS"
        mock_product_kms.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product_kms.PartialProductKey = "XXXXX"

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.Description = "Test"

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service]
            if "SoftwareLicensingService" in q
            else [mock_product, mock_product_kms]
        )

        call_count = [0]

        def set_string_side_effect(*args: Any, **kwargs: Any) -> int:
            call_count[0] += 1
            if call_count[0] == 1:
                return 0
            return 1

        reg.set_string_value = Mock(side_effect=set_string_side_effect)

        with pytest.raises(slmgr.SLMgrError) as exc_info:
            slmgr.uninstall_product_key(conn, reg, output)
        assert "Failed to set registry value" in str(exc_info.value)

    def test_display_installation_id(self) -> None:
        """Test display_installation_id"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.OfflineInstallationId = "123456789"

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.display_installation_id(conn, output)

        result = output.get_output()
        assert "123456789" in result
        assert "phone.inf" in result

    def test_display_installation_id_with_activation_id(self) -> None:
        """Test display_installation_id with activation_id"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.OfflineInstallationId = "123456789"

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.display_installation_id(conn, output, "test-id")

        result = output.get_output()
        assert "123456789" in result
        assert "phone.inf" in result

    def test_display_installation_id_indeterminate_operation(self) -> None:
        """Test display_installation_id with indeterminate operation warning"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.OfflineInstallationId = "123456789"
        mock_product.Description = "VOLUME_KMS"

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.display_installation_id(conn, output)

        result = output.get_output()
        assert "123456789" in result

    def test_display_installation_id_not_found(self) -> None:
        """Test display_installation_id when product not found"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = []

        slmgr.display_installation_id(conn, output)

        result = output.get_output()
        assert "not found" in result.lower()

    def test_activate_product(self) -> None:
        """Test activate_product"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 1
        mock_product.VLActivationTypeEnabled = 0
        mock_product.Description = "Test"

        conn.wmi_service = Mock()
        call_count = [0]

        def query_side_effect(q: str) -> list:
            call_count[0] += 1
            if "SoftwareLicensingService" in q:
                return [mock_service]
            return [mock_product]

        conn.wmi_service.query.side_effect = query_side_effect

        slmgr.activate_product(conn, output)

        result = output.get_output()
        assert "activated" in result.lower()

    def test_activate_product_token_only(self) -> None:
        """Test activate_product with token-only"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.VLActivationTypeEnabled = 3

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        slmgr.activate_product(conn, output)

        result = output.get_output()
        assert "Token-based activation" in result

    def test_activate_product_mak_already_active(self) -> None:
        """Test activate_product with already activated MAK"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 1
        mock_product.VLActivationTypeEnabled = 0
        mock_product.Description = "MAK"

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        slmgr.activate_product(conn, output)

        result = output.get_output()
        assert "activated" in result.lower()

    def test_activate_product_non_genuine(self) -> None:
        """Test activate_product non-genuine"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 4
        mock_product.VLActivationTypeEnabled = 0
        mock_product.Description = "Test"

        conn.wmi_service = Mock()
        call_count = [0]

        def query_side_effect(q: str) -> list:
            call_count[0] += 1
            if "SoftwareLicensingService" in q:
                return [mock_service]
            return [mock_product]

        conn.wmi_service.query.side_effect = query_side_effect

        slmgr.activate_product(conn, output)

        result = output.get_output()
        assert "non-genuine" in result.lower()

    def test_phone_activate_product(self) -> None:
        """Test phone_activate_product"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.OfflineInstallationId = "123456789"
        mock_product.LicenseStatus = 1

        conn.wmi_service = Mock()
        call_count = [0]

        def query_side_effect(q: str) -> list:
            call_count[0] += 1
            if "SoftwareLicensingService" in q:
                return [mock_service]
            return [mock_product]

        conn.wmi_service.query.side_effect = query_side_effect

        slmgr.phone_activate_product(conn, output, "111111-222222-333333")

        result = output.get_output()
        assert "deposited successfully" in result.lower()

    def test_clear_product_key_from_registry(self) -> None:
        """Test clear_product_key_from_registry"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.clear_product_key_from_registry(conn, output)

        mock_service.ClearProductKeyFromRegistry.assert_called_once()
        result = output.get_output()
        assert "cleared" in result.lower()

    @patch("builtins.open", mock_open(read_data="<License></License>"))
    def test_install_license(self) -> None:
        """Test install_license"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.install_license(conn, output, "test.xrm-ms")

        mock_service.InstallLicense.assert_called_once()
        result = output.get_output()
        assert "installed" in result.lower()

    @patch("builtins.open")
    def test_install_license_encoding_fallback(self, mock_file: Mock) -> None:
        """Test install_license with encoding fallback"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        # First call fails with UTF-8, second with UTF-16, third succeeds with ASCII
        mock_file.side_effect = [
            UnicodeDecodeError("utf-8", b"", 0, 1, "test"),
            mock_open(read_data="<License></License>")(),
        ]

        slmgr.install_license(conn, output, "test.xrm-ms")

        result = output.get_output()
        assert "installed" in result.lower()

    @patch("builtins.open")
    def test_install_license_encoding_fallback_to_ascii(self, mock_file: Mock) -> None:
        """Test install_license with UTF-16 failure, fallback to ASCII"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        # First call fails with UTF-8, second fails with UTF-16, third succeeds with ASCII
        mock_utf8_file = MagicMock()
        mock_utf8_file.__enter__.side_effect = UnicodeDecodeError(
            "utf-8", b"", 0, 1, "test"
        )

        mock_utf16_file = MagicMock()
        mock_utf16_file.__enter__.side_effect = Exception("UTF-16 error")

        mock_ascii_file = MagicMock()
        mock_ascii_file.__enter__.return_value.read.return_value = "<License></License>"

        mock_file.side_effect = [mock_utf8_file, mock_utf16_file, mock_ascii_file]

        slmgr.install_license(conn, output, "test.xrm-ms")

        result = output.get_output()
        assert "installed" in result.lower()

    @patch("os.path.exists")
    @patch("os.walk")
    @patch("builtins.open", mock_open(read_data="<License></License>"))
    def test_reinstall_licenses(self, mock_walk: Mock, mock_exists: Mock) -> None:
        """Test reinstall_licenses"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        mock_exists.return_value = True
        mock_walk.return_value = [
            ("c:\\windows\\system32\\spp\\tokens", [], ["test.xrm-ms"]),
        ]

        slmgr.reinstall_licenses(conn, output)

        result = output.get_output()
        assert "re-installed" in result.lower()

    @patch("os.path.exists")
    @patch("os.walk")
    @patch("builtins.open")
    def test_reinstall_licenses_with_errors(
        self, mock_file: Mock, mock_walk: Mock, mock_exists: Mock
    ) -> None:
        """Test reinstall_licenses with file read errors"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        mock_exists.return_value = True
        mock_walk.return_value = [
            (
                "c:\\windows\\system32\\spp\\tokens",
                [],
                ["test1.xrm-ms", "test2.xrm-ms"],
            ),
            ("c:\\windows\\system32\\oem", [], ["test3.xrm-ms"]),
        ]

        # Make some files fail to read
        mock_file.side_effect = [
            Exception("File read error"),
            mock_open(read_data="<License></License>")(),
            mock_open(read_data="<License></License>")(),
        ]

        slmgr.reinstall_licenses(conn, output)

        result = output.get_output()
        assert "re-installed" in result.lower()

    def test_rearm_windows(self) -> None:
        """Test rearm_windows"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.rearm_windows(conn, output)

        mock_service.ReArmWindows.assert_called_once()
        result = output.get_output()
        assert "restart" in result.lower()

    def test_rearm_app(self) -> None:
        """Test rearm_app"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.rearm_app(conn, output, "test-app-id")

        mock_service.ReArmApp.assert_called_once_with("test-app-id")

    def test_rearm_sku_found(self) -> None:
        """Test rearm_sku when SKU found"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.rearm_sku(conn, output, "test-id")

        mock_product.ReArmsku.assert_called_once()

    def test_rearm_sku_not_found(self) -> None:
        """Test rearm_sku when SKU not found"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = []

        slmgr.rearm_sku(conn, output, "test-id")

        result = output.get_output()
        assert "not found" in result.lower()

    def test_expiration_datetime_unlicensed(self) -> None:
        """Test expiration_datetime with unlicensed status"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 0
        mock_product.GracePeriodRemaining = 0
        mock_product.Description = "Test"

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.expiration_datetime(conn, output)

        result = output.get_output()
        assert "Unlicensed" in result

    def test_expiration_datetime_permanently_activated(self) -> None:
        """Test expiration_datetime with permanently activated"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 1
        mock_product.GracePeriodRemaining = 0
        mock_product.Description = "Test"

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.expiration_datetime(conn, output)

        result = output.get_output()
        assert "permanently activated" in result.lower()

    def test_expiration_datetime_timebased(self) -> None:
        """Test expiration_datetime with timebased activation"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 1
        mock_product.GracePeriodRemaining = 1440  # 1 day
        mock_product.Description = "TIMEBASED_"

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.expiration_datetime(conn, output)

        result = output.get_output()
        assert "Timebased activation" in result

    def test_expiration_datetime_notification(self) -> None:
        """Test expiration_datetime with notification status"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 5
        mock_product.GracePeriodRemaining = 0
        mock_product.Description = "Test"

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.expiration_datetime(conn, output)

        result = output.get_output()
        assert "Notification" in result

    def test_expiration_datetime_not_found(self) -> None:
        """Test expiration_datetime when not found"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = []

        slmgr.expiration_datetime(conn, output)

        result = output.get_output()
        assert "not found" in result.lower()

    def test_install_product_key_kms_server_32bit_key_exists(self) -> None:
        """Test install_product_key with KMS server and 32-bit registry key exists"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"

        mock_product = Mock()
        mock_product.Description = "VOLUME_KMS"
        mock_product.PartialProductKey = "XXXXX"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.LicenseIsAddon = False

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        reg.set_string_value = Mock(return_value=0)
        reg.delete_value = Mock(return_value=0)
        reg.key_exists = Mock(return_value=True)  # 32-bit key exists

        slmgr.install_product_key(conn, reg, output, "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX")

        result = output.get_output()
        assert "successfully" in result.lower()
        # Verify that set_string_value was called twice (once for 64-bit, once for 32-bit)
        assert reg.set_string_value.call_count == 2

    def test_install_product_key_kms_server_registry_error_32bit(self) -> None:
        """Test install_product_key with KMS server and registry error on 32-bit key"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"

        mock_product = Mock()
        mock_product.Description = "VOLUME_KMS"
        mock_product.PartialProductKey = "XXXXX"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.LicenseIsAddon = False

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        # First call succeeds, second call fails
        reg.set_string_value = Mock(side_effect=[0, 1])
        reg.delete_value = Mock(return_value=0)
        reg.key_exists = Mock(return_value=True)

        with pytest.raises(slmgr.SLMgrError) as exc_info:
            slmgr.install_product_key(
                conn, reg, output, "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"
            )

        assert "Failed to set registry value" in str(exc_info.value)

    def test_install_product_key_indeterminate_operation(self) -> None:
        """Test install_product_key with indeterminate operation warning"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"

        mock_product = Mock()
        mock_product.Description = "Test Product"
        mock_product.PartialProductKey = "XXXXX"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.LicenseIsAddon = False
        # Make get_is_primary_windows_sku return 2 (indeterminate)
        mock_product.LicenseStatus = 1

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        reg.set_string_value = Mock(return_value=0)
        reg.delete_value = Mock(return_value=0)
        reg.key_exists = Mock(return_value=False)

        # Mock get_is_primary_windows_sku to return 2
        with patch("slmgr.get_is_primary_windows_sku", return_value=2):
            slmgr.install_product_key(
                conn, reg, output, "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"
            )

        result = output.get_output()
        assert "successfully" in result.lower()

    def test_uninstall_product_key_kms_server_registry_error_alternate(self) -> None:
        """Test uninstall_product_key with KMS server and registry error"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"

        mock_product = Mock()
        mock_product.Description = "VOLUME_KMS"
        mock_product.PartialProductKey = "XXXXX"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.Uninstall = Mock()

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        # First call succeeds, second call fails
        reg.set_string_value = Mock(side_effect=[0, 1])

        with pytest.raises(slmgr.SLMgrError) as exc_info:
            slmgr.uninstall_product_key(conn, reg, output)

        assert "Failed to set registry value" in str(exc_info.value)

    def test_activate_product_non_genuine_grace_period(self) -> None:
        """Test activate_product with non-genuine grace period"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.RefreshLicenseStatus = Mock()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 4  # Non-genuine grace period
        mock_product.LicenseStatusReason = slmgr.HR_SL_E_GRACE_TIME_EXPIRED
        mock_product.VLActivationTypeEnabled = 0
        mock_product.Activate = Mock()

        conn.wmi_service = Mock()
        conn.get_service_object = Mock(return_value=mock_service)
        conn.get_product_collection = Mock(return_value=[mock_product])
        conn.get_product_object = Mock(return_value=mock_product)

        slmgr.activate_product(conn, output)

        result = output.get_output()
        assert "non-genuine grace period" in result.lower()

    def test_activate_product_non_genuine_notification(self) -> None:
        """Test activate_product with non-genuine notification"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.RefreshLicenseStatus = Mock()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 5  # Non-genuine notification
        mock_product.LicenseStatusReason = slmgr.HR_SL_E_NOT_GENUINE
        mock_product.VLActivationTypeEnabled = 0
        mock_product.Activate = Mock()

        conn.wmi_service = Mock()
        conn.get_service_object = Mock(return_value=mock_service)
        conn.get_product_collection = Mock(return_value=[mock_product])
        conn.get_product_object = Mock(return_value=mock_product)

        slmgr.activate_product(conn, output)

        result = output.get_output()
        assert "non-genuine notification" in result.lower()

    def test_activate_product_extended_grace(self) -> None:
        """Test activate_product with extended grace period"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.RefreshLicenseStatus = Mock()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 6  # Extended grace period
        mock_product.VLActivationTypeEnabled = 0
        mock_product.Activate = Mock()

        conn.wmi_service = Mock()
        conn.get_service_object = Mock(return_value=mock_service)
        conn.get_product_collection = Mock(return_value=[mock_product])
        conn.get_product_object = Mock(return_value=mock_product)

        slmgr.activate_product(conn, output)

        result = output.get_output()
        assert "Extended grace period" in result

    def test_activate_product_failure(self) -> None:
        """Test activate_product with activation failure"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.RefreshLicenseStatus = Mock()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 99  # Unknown status
        mock_product.VLActivationTypeEnabled = 0
        mock_product.Activate = Mock()

        conn.wmi_service = Mock()
        conn.get_service_object = Mock(return_value=mock_service)
        conn.get_product_collection = Mock(return_value=[mock_product])
        conn.get_product_object = Mock(return_value=mock_product)

        slmgr.activate_product(conn, output)

        result = output.get_output()
        assert "activation failed" in result.lower()

    def test_phone_activate_product_non_genuine_grace_period(self) -> None:
        """Test phone_activate_product with non-genuine grace period"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.RefreshLicenseStatus = Mock()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 4  # Non-genuine grace period
        mock_product.LicenseStatusReason = slmgr.HR_SL_E_GRACE_TIME_EXPIRED
        mock_product.DepositOfflineConfirmationId = Mock()

        conn.wmi_service = Mock()
        conn.get_service_object = Mock(return_value=mock_service)
        conn.get_product_collection = Mock(return_value=[mock_product])
        conn.get_product_object = Mock(return_value=mock_product)

        slmgr.phone_activate_product(conn, output, "123456")

        result = output.get_output()
        assert "non-genuine grace period" in result.lower()

    def test_phone_activate_product_non_genuine_notification(self) -> None:
        """Test phone_activate_product with non-genuine notification"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.RefreshLicenseStatus = Mock()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 5  # Non-genuine notification
        mock_product.LicenseStatusReason = slmgr.HR_SL_E_NOT_GENUINE
        mock_product.DepositOfflineConfirmationId = Mock()

        conn.wmi_service = Mock()
        conn.get_service_object = Mock(return_value=mock_service)
        conn.get_product_collection = Mock(return_value=[mock_product])
        conn.get_product_object = Mock(return_value=mock_product)

        slmgr.phone_activate_product(conn, output, "123456")

        result = output.get_output()
        assert "non-genuine notification" in result.lower()

    def test_phone_activate_product_extended_grace(self) -> None:
        """Test phone_activate_product with extended grace period"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.RefreshLicenseStatus = Mock()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 6  # Extended grace period
        mock_product.DepositOfflineConfirmationId = Mock()

        conn.wmi_service = Mock()
        conn.get_service_object = Mock(return_value=mock_service)
        conn.get_product_collection = Mock(return_value=[mock_product])
        conn.get_product_object = Mock(return_value=mock_product)

        slmgr.phone_activate_product(conn, output, "123456")

        result = output.get_output()
        assert "Extended grace period" in result

    def test_phone_activate_product_failure(self) -> None:
        """Test phone_activate_product with activation failure"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.RefreshLicenseStatus = Mock()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseIsAddon = False
        mock_product.LicenseStatus = 99  # Unknown status
        mock_product.DepositOfflineConfirmationId = Mock()

        conn.wmi_service = Mock()
        conn.get_service_object = Mock(return_value=mock_service)
        conn.get_product_collection = Mock(return_value=[mock_product])
        conn.get_product_object = Mock(return_value=mock_product)

        slmgr.phone_activate_product(conn, output, "123456")

        result = output.get_output()
        assert "activation failed" in result.lower()

    def test_expiration_datetime_with_activation_id(self) -> None:
        """Test expiration_datetime with activation_id"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseStatus = 0  # Unlicensed
        mock_product.GracePeriodRemaining = 43200  # 30 days in minutes

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.expiration_datetime(conn, output, "test-id")

        result = output.get_output()
        assert "unlicensed" in result.lower()

    def test_expiration_datetime_notification_status(self) -> None:
        """Test expiration_datetime with notification status"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseStatus = 5  # Notification
        mock_product.GracePeriodRemaining = 1440  # 1 day in minutes

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.expiration_datetime(conn, output)

        result = output.get_output()
        assert "notification" in result.lower()

    def test_expiration_datetime_grace_period(self) -> None:
        """Test expiration_datetime with grace period"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseStatus = 2  # OOB Grace
        mock_product.GracePeriodRemaining = 2880  # 2 days in minutes

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.expiration_datetime(conn, output)

        result = output.get_output()
        assert "initial grace period ends" in result.lower()

    def test_expiration_datetime_oob_grace(self) -> None:
        """Test expiration_datetime with OOB grace"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseStatus = 3  # OOT Grace
        mock_product.GracePeriodRemaining = 4320  # 3 days in minutes

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.expiration_datetime(conn, output)

        result = output.get_output()
        assert "additional grace period ends" in result.lower()

    def test_expiration_datetime_extended_grace(self) -> None:
        """Test expiration_datetime with extended grace"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseStatus = 6  # Extended grace
        mock_product.GracePeriodRemaining = 5760  # 4 days in minutes

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.expiration_datetime(conn, output)

        result = output.get_output()
        assert "extended grace period ends" in result.lower()
