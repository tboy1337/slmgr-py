"""Tests for expiration datetime edge cases in slmgr.py"""

from unittest.mock import Mock

import slmgr


class TestExpirationDateTime:
    """Test expiration_datetime function edge cases"""

    def test_expiration_datetime_avma(self) -> None:
        """Test expiration_datetime with AVMA license"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseStatus = 1
        mock_product.GracePeriodRemaining = 1440
        mock_product.Description = "VIRTUAL_MACHINE_ACTIVATION"
        mock_product.Name = "Windows Test Edition"
        mock_product.LicenseIsAddon = False

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.expiration_datetime(conn, output)

        result = output.get_output()
        assert "Automatic VM activation" in result

    def test_expiration_datetime_additional_grace(self) -> None:
        """Test expiration_datetime with additional grace period"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseStatus = 3
        mock_product.GracePeriodRemaining = 1440
        mock_product.Name = "Windows Test Edition"
        mock_product.Description = "Test"
        mock_product.LicenseIsAddon = False

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.expiration_datetime(conn, output)

        result = output.get_output()
        assert "Additional grace period" in result

    def test_expiration_datetime_non_genuine_grace(self) -> None:
        """Test expiration_datetime with non-genuine grace period"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        mock_product.PartialProductKey = "XXXXX"
        mock_product.LicenseStatus = 4
        mock_product.GracePeriodRemaining = 1440
        mock_product.Name = "Windows Test Edition"
        mock_product.Description = "Test"
        mock_product.LicenseIsAddon = False

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.expiration_datetime(conn, output)

        result = output.get_output()
        assert "Non-genuine grace period" in result
