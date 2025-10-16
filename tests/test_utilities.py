"""Tests for utility functions in slmgr.py"""

from unittest.mock import Mock

import pytest

import slmgr


class TestUtilityFunctions:
    """Test utility functions"""

    def test_get_error_message_known_code(self) -> None:
        """Test get_error_message with known error code"""
        result = slmgr.get_error_message(0xC004C001)
        assert "invalid" in result.lower()

    def test_get_error_message_unknown_code(self) -> None:
        """Test get_error_message with unknown error code"""
        result = slmgr.get_error_message(0x99999999)
        assert "slui.exe" in result

    def test_show_error_with_positive_code(
        self, capsys: pytest.CaptureFixture[str]
    ) -> None:
        """Test show_error with positive error code"""
        slmgr.show_error("Test error", 5)
        captured = capsys.readouterr()
        assert "Test error" in captured.err

    def test_show_error_with_negative_code(
        self, capsys: pytest.CaptureFixture[str]
    ) -> None:
        """Test show_error with negative error code"""
        slmgr.show_error("Test error", -2147024809)
        captured = capsys.readouterr()
        assert "0x" in captured.err

    def test_show_error_with_description(
        self, capsys: pytest.CaptureFixture[str]
    ) -> None:
        """Test show_error with description"""
        slmgr.show_error("Test error", 0xC004C001, "Custom desc")
        captured = capsys.readouterr()
        assert "Custom desc" in captured.err

    def test_show_error_with_no_code(self, capsys: pytest.CaptureFixture[str]) -> None:
        """Test show_error with no error code"""
        slmgr.show_error("Test error", None, "Description")
        captured = capsys.readouterr()
        assert "Test error" in captured.err
        assert "Description" in captured.err

    def test_quit_with_error(self) -> None:
        """Test quit_with_error exits"""
        with pytest.raises(SystemExit):
            slmgr.quit_with_error("Error", 1)

    def test_is_kms_client_true(self) -> None:
        """Test is_kms_client returns True"""
        assert slmgr.is_kms_client("VOLUME_KMSCLIENT")

    def test_is_kms_client_false(self) -> None:
        """Test is_kms_client returns False"""
        assert not slmgr.is_kms_client("OTHER")

    def test_is_kms_server_true(self) -> None:
        """Test is_kms_server returns True"""
        assert slmgr.is_kms_server("VOLUME_KMS")

    def test_is_kms_server_false_for_client(self) -> None:
        """Test is_kms_server returns False for client"""
        assert not slmgr.is_kms_server("VOLUME_KMSCLIENT")

    def test_is_kms_server_false(self) -> None:
        """Test is_kms_server returns False"""
        assert not slmgr.is_kms_server("OTHER")

    def test_is_tbl_true(self) -> None:
        """Test is_tbl returns True"""
        assert slmgr.is_tbl("TIMEBASED_")

    def test_is_tbl_false(self) -> None:
        """Test is_tbl returns False"""
        assert not slmgr.is_tbl("OTHER")

    def test_is_avma_true(self) -> None:
        """Test is_avma returns True"""
        assert slmgr.is_avma("VIRTUAL_MACHINE_ACTIVATION")

    def test_is_avma_false(self) -> None:
        """Test is_avma returns False"""
        assert not slmgr.is_avma("OTHER")

    def test_is_mak_true(self) -> None:
        """Test is_mak returns True"""
        assert slmgr.is_mak("MAK")

    def test_is_mak_false(self) -> None:
        """Test is_mak returns False"""
        assert not slmgr.is_mak("OTHER")

    def test_is_token_activated_true(self) -> None:
        """Test is_token_activated returns True"""
        product = Mock()
        product.TokenActivationILVID = 123
        assert slmgr.is_token_activated(product)

    def test_is_token_activated_false_none(self) -> None:
        """Test is_token_activated returns False when None"""
        product = Mock()
        product.TokenActivationILVID = None
        assert not slmgr.is_token_activated(product)

    def test_is_token_activated_false_max(self) -> None:
        """Test is_token_activated returns False when 0xFFFFFFFF"""
        product = Mock()
        product.TokenActivationILVID = 0xFFFFFFFF
        assert not slmgr.is_token_activated(product)

    def test_is_token_activated_exception(self) -> None:
        """Test is_token_activated returns False on exception"""
        product = Mock()
        del product.TokenActivationILVID
        assert not slmgr.is_token_activated(product)

    def test_is_ad_activated_true(self) -> None:
        """Test is_ad_activated returns True"""
        product = Mock()
        product.VLActivationType = 1
        assert slmgr.is_ad_activated(product)

    def test_is_ad_activated_false(self) -> None:
        """Test is_ad_activated returns False"""
        product = Mock()
        product.VLActivationType = 0
        assert not slmgr.is_ad_activated(product)

    def test_is_ad_activated_exception(self) -> None:
        """Test is_ad_activated returns False on exception"""
        product = Mock()
        del product.VLActivationType
        assert not slmgr.is_ad_activated(product)

    def test_get_is_primary_windows_sku_not_windows(self) -> None:
        """Test get_is_primary_windows_sku returns 0 for non-Windows"""
        product = Mock()
        product.ApplicationId = "other-id"
        product.PartialProductKey = "12345"
        result = slmgr.get_is_primary_windows_sku(product)
        assert result == 0

    def test_get_is_primary_windows_sku_no_key(self) -> None:
        """Test get_is_primary_windows_sku returns 0 without key"""
        product = Mock()
        product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        product.PartialProductKey = None
        result = slmgr.get_is_primary_windows_sku(product)
        assert result == 0

    def test_get_is_primary_windows_sku_addon(self) -> None:
        """Test get_is_primary_windows_sku returns 0 for addon"""
        product = Mock()
        product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        product.PartialProductKey = "12345"
        product.LicenseIsAddon = True
        result = slmgr.get_is_primary_windows_sku(product)
        assert result == 0

    def test_get_is_primary_windows_sku_not_addon(self) -> None:
        """Test get_is_primary_windows_sku returns 1 for primary"""
        product = Mock()
        product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        product.PartialProductKey = "12345"
        product.LicenseIsAddon = False
        result = slmgr.get_is_primary_windows_sku(product)
        assert result == 1

    def test_get_is_primary_windows_sku_exception_kms_client(self) -> None:
        """Test get_is_primary_windows_sku returns 1 for KMS client on exception"""
        product = Mock()
        product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        product.PartialProductKey = "12345"
        product.Description = "VOLUME_KMSCLIENT"
        del product.LicenseIsAddon
        result = slmgr.get_is_primary_windows_sku(product)
        assert result == 1

    def test_get_is_primary_windows_sku_exception_kms_server(self) -> None:
        """Test get_is_primary_windows_sku returns 1 for KMS server on exception"""
        product = Mock()
        product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        product.PartialProductKey = "12345"
        product.Description = "VOLUME_KMS"
        del product.LicenseIsAddon
        result = slmgr.get_is_primary_windows_sku(product)
        assert result == 1

    def test_get_is_primary_windows_sku_exception_indeterminate(self) -> None:
        """Test get_is_primary_windows_sku returns 2 when indeterminate"""
        product = Mock()
        product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        product.PartialProductKey = "12345"
        product.Description = "OTHER"
        del product.LicenseIsAddon
        result = slmgr.get_is_primary_windows_sku(product)
        assert result == 2

    def test_check_product_for_command_no_id_windows_not_addon(self) -> None:
        """Test check_product_for_command matches Windows not addon"""
        product = Mock()
        product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        product.LicenseIsAddon = False
        assert slmgr.check_product_for_command(product, "")

    def test_check_product_for_command_no_id_addon(self) -> None:
        """Test check_product_for_command doesn't match addon"""
        product = Mock()
        product.ApplicationId = slmgr.WINDOWS_APP_ID.lower()
        product.LicenseIsAddon = True
        assert not slmgr.check_product_for_command(product, "")

    def test_check_product_for_command_matching_id(self) -> None:
        """Test check_product_for_command matches ID"""
        product = Mock()
        product.ID = "test-id"
        assert slmgr.check_product_for_command(product, "test-id")

    def test_check_product_for_command_case_insensitive(self) -> None:
        """Test check_product_for_command is case insensitive"""
        product = Mock()
        product.ID = "TEST-ID"
        assert slmgr.check_product_for_command(product, "test-id")

    def test_check_product_for_command_no_match(self) -> None:
        """Test check_product_for_command no match"""
        product = Mock()
        product.ID = "other-id"
        product.ApplicationId = "other-app"
        product.LicenseIsAddon = False
        assert not slmgr.check_product_for_command(product, "test-id")

    def test_get_days_from_mins_exact(self) -> None:
        """Test get_days_from_mins exact division"""
        assert slmgr.get_days_from_mins(1440) == 1  # 1 day

    def test_get_days_from_mins_ceiling(self) -> None:
        """Test get_days_from_mins rounds up"""
        assert slmgr.get_days_from_mins(1441) == 2

    def test_get_days_from_mins_zero(self) -> None:
        """Test get_days_from_mins with zero"""
        assert slmgr.get_days_from_mins(0) == 0

    def test_guid_to_string_valid(self) -> None:
        """Test guid_to_string with valid bytes"""
        guid_bytes = bytes(
            [
                0x34,
                0x27,
                0xC9,
                0x55,  # First 4 bytes (reversed)
                0x82,
                0xD6,  # Next 2 bytes (reversed)
                0x71,
                0x4D,  # Next 2 bytes (reversed)
                0x98,
                0x3E,  # Next 2 bytes
                0xD6,
                0xEC,
                0x3F,
                0x16,
                0x05,
                0x9F,  # Last 6 bytes
            ]
        )
        result = slmgr.guid_to_string(guid_bytes)
        assert result.startswith("{")
        assert result.endswith("}")
        assert "-" in result

    def test_guid_to_string_short(self) -> None:
        """Test guid_to_string with short bytes"""
        result = slmgr.guid_to_string(b"short")
        assert result == ""

    def test_wmi_date_to_datetime_valid(self) -> None:
        """Test wmi_date_to_datetime with valid date"""
        wmi_date = "20231225143000.000000+000"
        result = slmgr.wmi_date_to_datetime(wmi_date)
        assert result is not None
        assert result.year == 2023
        assert result.month == 12
        assert result.day == 25
        assert result.hour == 14
        assert result.minute == 30

    def test_wmi_date_to_datetime_zero(self) -> None:
        """Test wmi_date_to_datetime with zero date"""
        result = slmgr.wmi_date_to_datetime("00000000000000.000000+000")
        assert result is None

    def test_wmi_date_to_datetime_empty(self) -> None:
        """Test wmi_date_to_datetime with empty string"""
        result = slmgr.wmi_date_to_datetime("")
        assert result is None

    def test_wmi_date_to_datetime_invalid(self) -> None:
        """Test wmi_date_to_datetime with invalid date"""
        result = slmgr.wmi_date_to_datetime("invalid")
        assert result is None

    def test_output_indeterminate_operation_warning(self) -> None:
        """Test output_indeterminate_operation_warning"""
        product = Mock()
        product.Description = "Test Product"
        product.ID = "test-id"
        output = slmgr.OutputManager()
        slmgr.output_indeterminate_operation_warning(product, output)
        result = output.get_output()
        assert "SLMGR was not able to validate" in result
        assert "Test Product" in result
        assert "test-id" in result

    def test_fail_remote_exec_local(self) -> None:
        """Test fail_remote_exec doesn't raise for local"""
        slmgr.fail_remote_exec(False)

    def test_fail_remote_exec_remote(self) -> None:
        """Test fail_remote_exec raises for remote"""
        with pytest.raises(slmgr.SLMgrError):
            slmgr.fail_remote_exec(True)


class TestSLMgrError:
    """Test SLMgrError exception class"""

    def test_slmgr_error_with_code(self) -> None:
        """Test SLMgrError with error code"""
        error = slmgr.SLMgrError("Test message", 123)
        assert error.message == "Test message"
        assert error.error_code == 123
        assert str(error) == "Test message"

    def test_slmgr_error_without_code(self) -> None:
        """Test SLMgrError without error code"""
        error = slmgr.SLMgrError("Test message")
        assert error.message == "Test message"
        assert error.error_code is None


class TestOutputManager:
    """Test OutputManager class"""

    def test_output_manager_line_out(self) -> None:
        """Test OutputManager.line_out"""
        output = slmgr.OutputManager()
        output.line_out("Line 1")
        output.line_out("Line 2")
        assert output.get_output() == "Line 1\nLine 2"

    def test_output_manager_line_flush_no_text(
        self, capsys: pytest.CaptureFixture[str]
    ) -> None:
        """Test OutputManager.line_flush without text"""
        output = slmgr.OutputManager()
        output.line_out("Test")
        output.line_flush()
        captured = capsys.readouterr()
        assert "Test" in captured.out
        assert output.get_output() == ""

    def test_output_manager_line_flush_with_text(
        self, capsys: pytest.CaptureFixture[str]
    ) -> None:
        """Test OutputManager.line_flush with text"""
        output = slmgr.OutputManager()
        output.line_out("Line 1")
        output.line_flush("Line 2")
        captured = capsys.readouterr()
        assert "Line 1" in captured.out
        assert "Line 2" in captured.out

    def test_output_manager_empty_flush(
        self, capsys: pytest.CaptureFixture[str]
    ) -> None:
        """Test OutputManager.line_flush when empty"""
        output = slmgr.OutputManager()
        output.line_flush()
        captured = capsys.readouterr()
        assert captured.out == ""

    def test_output_manager_get_output_empty(self) -> None:
        """Test OutputManager.get_output when empty"""
        output = slmgr.OutputManager()
        assert output.get_output() == ""
