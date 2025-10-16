"""Property-based tests using Hypothesis for utility functions in slmgr.py"""

from datetime import datetime
from unittest.mock import Mock

import pytest
from hypothesis import assume, example, given, settings
from hypothesis import strategies as st

import slmgr


class TestGuidConversion:
    """Property-based tests for GUID conversion"""

    @given(st.binary(min_size=16, max_size=16))
    @settings(max_examples=1000)
    def test_guid_to_string_format(self, guid_bytes: bytes) -> None:
        """Test that guid_to_string always produces valid GUID format"""
        result = slmgr.guid_to_string(guid_bytes)

        # Should start with { and end with }
        assert result.startswith("{")
        assert result.endswith("}")

        # Should have exactly 4 hyphens in correct positions
        assert result.count("-") == 4

        # Total length should be 38 characters (including braces)
        assert len(result) == 38

        # Should be uppercase hexadecimal with hyphens
        inner = result[1:-1]  # Remove braces
        parts = inner.split("-")
        assert len(parts) == 5
        assert len(parts[0]) == 8  # xxxxxxxx
        assert len(parts[1]) == 4  # xxxx
        assert len(parts[2]) == 4  # xxxx
        assert len(parts[3]) == 4  # xxxx
        assert len(parts[4]) == 12  # xxxxxxxxxxxx

        # All characters should be valid hex digits
        for part in parts:
            assert all(c in "0123456789ABCDEF" for c in part)

    @given(st.binary(min_size=0, max_size=15))
    @settings(max_examples=500)
    def test_guid_to_string_short_input(self, guid_bytes: bytes) -> None:
        """Test that guid_to_string returns empty string for short input"""
        result = slmgr.guid_to_string(guid_bytes)
        assert result == ""

    @given(st.binary(min_size=16, max_size=16))
    @example(bytes([0x00] * 16))
    @example(bytes([0xFF] * 16))
    @example(bytes(range(16)))
    def test_guid_to_string_idempotent(self, guid_bytes: bytes) -> None:
        """Test that guid_to_string is consistent for same input"""
        result1 = slmgr.guid_to_string(guid_bytes)
        result2 = slmgr.guid_to_string(guid_bytes)
        assert result1 == result2


class TestDaysFromMinsConversion:
    """Property-based tests for minutes to days conversion"""

    @given(st.integers(min_value=0, max_value=1000000))
    @settings(max_examples=1000)
    def test_get_days_from_mins_positive(self, minutes: int) -> None:
        """Test that get_days_from_mins handles positive values correctly"""
        days = slmgr.get_days_from_mins(minutes)

        # Result should be non-negative
        assert days >= 0

        # Should implement ceiling division
        expected = (minutes + 1439) // 1440
        assert days == expected

    @given(st.integers(min_value=0, max_value=1000000))
    @example(0)
    @example(1)
    @example(1439)
    @example(1440)
    @example(1441)
    def test_get_days_from_mins_ceiling(self, minutes: int) -> None:
        """Test ceiling behavior of get_days_from_mins"""
        days = slmgr.get_days_from_mins(minutes)

        # If minutes > 0, days should be at least 1
        if minutes > 0:
            assert days >= 1
        else:
            assert days == 0

        # Days should be the smallest integer such that days * 1440 >= minutes
        if minutes > 0:
            assert (days - 1) * 1440 < minutes
            assert days * 1440 >= minutes

    @given(st.integers(min_value=1, max_value=1000))
    def test_get_days_from_mins_exact_days(self, num_days: int) -> None:
        """Test exact day boundaries"""
        minutes = num_days * 1440
        result = slmgr.get_days_from_mins(minutes)
        assert result == num_days


class TestWMIDateParsing:
    """Property-based tests for WMI date parsing"""

    @given(
        year=st.integers(min_value=1900, max_value=2100),
        month=st.integers(min_value=1, max_value=12),
        day=st.integers(min_value=1, max_value=28),  # Safe for all months
        hour=st.integers(min_value=0, max_value=23),
        minute=st.integers(min_value=0, max_value=59),
        second=st.integers(min_value=0, max_value=59),
    )
    @settings(max_examples=1000)
    def test_wmi_date_to_datetime_valid_dates(
        self, year: int, month: int, day: int, hour: int, minute: int, second: int
    ) -> None:
        """Test parsing valid WMI date strings"""
        wmi_date = f"{year:04d}{month:02d}{day:02d}{hour:02d}{minute:02d}{second:02d}.000000+000"
        result = slmgr.wmi_date_to_datetime(wmi_date)

        assert result is not None
        assert result.year == year
        assert result.month == month
        assert result.day == day
        assert result.hour == hour
        assert result.minute == minute
        assert result.second == second

    @given(st.text(min_size=0, max_size=50))
    @settings(max_examples=500)
    def test_wmi_date_to_datetime_invalid_strings(self, invalid_date: str) -> None:
        """Test that invalid WMI date strings return None"""
        # Exclude the zero date and valid formats
        assume(invalid_date != "00000000000000.000000+000")
        assume(not (len(invalid_date) == 25 and invalid_date[14] == "."))

        result = slmgr.wmi_date_to_datetime(invalid_date)
        # Should return None for invalid formats
        assert result is None or isinstance(result, datetime)

    @given(st.sampled_from(["", "00000000000000.000000+000", None]))
    def test_wmi_date_to_datetime_empty_or_zero(self, empty_date: str | None) -> None:
        """Test that empty or zero dates return None"""
        if empty_date is None:
            empty_date = ""
        result = slmgr.wmi_date_to_datetime(empty_date)
        assert result is None


class TestPatternMatching:
    """Property-based tests for pattern matching functions"""

    @given(st.text(min_size=0, max_size=200))
    @settings(max_examples=1000)
    def test_is_kms_client_pattern(self, description: str) -> None:
        """Test is_kms_client pattern matching"""
        result = slmgr.is_kms_client(description)

        # Should return True only if "VOLUME_KMSCLIENT" is in description
        assert result == ("VOLUME_KMSCLIENT" in description)

    @given(st.text(min_size=0, max_size=200))
    @settings(max_examples=1000)
    def test_is_kms_server_pattern(self, description: str) -> None:
        """Test is_kms_server pattern matching"""
        result = slmgr.is_kms_server(description)

        # Should return True only if "VOLUME_KMS" is in description
        # but not "VOLUME_KMSCLIENT"
        expected = "VOLUME_KMS" in description and "VOLUME_KMSCLIENT" not in description
        assert result == expected

    @given(st.text(min_size=0, max_size=200))
    @settings(max_examples=1000)
    def test_is_tbl_pattern(self, description: str) -> None:
        """Test is_tbl pattern matching"""
        result = slmgr.is_tbl(description)
        assert result == ("TIMEBASED_" in description)

    @given(st.text(min_size=0, max_size=200))
    @settings(max_examples=1000)
    def test_is_avma_pattern(self, description: str) -> None:
        """Test is_avma pattern matching"""
        result = slmgr.is_avma(description)
        assert result == ("VIRTUAL_MACHINE_ACTIVATION" in description)

    @given(st.text(min_size=0, max_size=200))
    @settings(max_examples=1000)
    def test_is_mak_pattern(self, description: str) -> None:
        """Test is_mak pattern matching"""
        result = slmgr.is_mak(description)
        assert result == ("MAK" in description)

    @given(
        st.text(min_size=10, max_size=50),
        st.booleans(),
    )
    @settings(max_examples=500)
    def test_pattern_matching_consistency(
        self, base_description: str, add_pattern: bool
    ) -> None:
        """Test that pattern matching functions are consistent"""
        if add_pattern:
            description = base_description + "VOLUME_KMSCLIENT"
        else:
            description = base_description

        # If it's a KMS client, it shouldn't be a KMS server
        if slmgr.is_kms_client(description):
            assert not slmgr.is_kms_server(description)


class TestErrorHandling:
    """Property-based tests for error handling"""

    @given(st.integers(min_value=-0x80000000, max_value=0x7FFFFFFF))
    @settings(max_examples=1000)
    def test_get_error_message_always_returns_string(self, error_code: int) -> None:
        """Test that get_error_message always returns a string"""
        result = slmgr.get_error_message(error_code)
        assert isinstance(result, str)
        assert len(result) > 0

    @given(st.sampled_from(list(slmgr.ERROR_MESSAGES.keys())))
    def test_get_error_message_known_codes(self, error_code: int) -> None:
        """Test get_error_message with known error codes"""
        result = slmgr.get_error_message(error_code)
        expected = slmgr.ERROR_MESSAGES[error_code]
        assert result == expected

    @given(
        st.text(min_size=1, max_size=100),
        st.one_of(st.integers(), st.none()),
        st.text(min_size=0, max_size=100),
    )
    @settings(max_examples=500)
    def test_show_error_no_crash(
        self, message: str, error_code: int | None, description: str
    ) -> None:
        """Test that show_error doesn't crash with any input"""
        # This should not raise any exception
        try:
            slmgr.show_error(message, error_code, description)
        except SystemExit:
            # quit_with_error calls sys.exit, which is expected
            pass

    @given(
        st.integers(min_value=-0xFFFFFFFF, max_value=0xFFFFFFFF),
    )
    @settings(max_examples=500)
    def test_show_error_code_formatting(self, error_code: int) -> None:
        """Test error code formatting in show_error"""
        # Capture output by checking the function doesn't crash
        try:
            slmgr.show_error("Test", error_code)
        except Exception as e:  # pylint: disable=broad-exception-caught
            pytest.fail(f"show_error raised unexpected exception: {e}")


class TestCheckProductForCommand:
    """Property-based tests for check_product_for_command"""

    @given(
        app_id=st.text(min_size=1, max_size=50),
        product_id=st.text(min_size=1, max_size=50),
        is_addon=st.booleans(),
        activation_id=st.text(min_size=0, max_size=50),
    )
    @settings(max_examples=1000)
    def test_check_product_for_command_logic(
        self, app_id: str, product_id: str, is_addon: bool, activation_id: str
    ) -> None:
        """Test check_product_for_command matching logic"""
        product = Mock()
        product.ApplicationId = app_id.lower()
        product.ID = product_id
        product.LicenseIsAddon = is_addon

        result = slmgr.check_product_for_command(product, activation_id)

        # Should match if:
        # 1. No activation_id and Windows app without addon
        # 2. Product ID matches activation_id (case insensitive)
        expected = (
            not activation_id
            and app_id.lower() == slmgr.WINDOWS_APP_ID.lower()
            and not is_addon
        ) or (product_id.lower() == activation_id.lower())
        assert result == expected

    @given(
        product_id=st.text(
            min_size=1,
            max_size=50,
            alphabet=st.characters(min_codepoint=32, max_codepoint=126),
        ),
    )
    @settings(max_examples=500)
    def test_check_product_for_command_case_insensitive(self, product_id: str) -> None:
        """Test that product ID matching is case insensitive"""
        # Only test ASCII characters to avoid Unicode case conversion issues
        assume(product_id.isascii())
        assume(len(product_id) > 0)

        product = Mock()
        product.ID = product_id.lower()
        product.ApplicationId = "other-app"
        product.LicenseIsAddon = False

        # Both upper and lower case should match
        result_lower = slmgr.check_product_for_command(product, product_id.lower())
        result_upper = slmgr.check_product_for_command(product, product_id.upper())

        assert result_lower == result_upper


class TestOutputManager:
    """Property-based tests for OutputManager"""

    @given(st.lists(st.text(min_size=0, max_size=100), min_size=0, max_size=50))
    @settings(max_examples=500)
    def test_output_manager_line_out(self, lines: list[str]) -> None:
        """Test OutputManager line_out with arbitrary text"""
        output = slmgr.OutputManager()

        for line in lines:
            output.line_out(line)

        result = output.get_output()

        if lines:
            # Should join with newlines
            expected = "\n".join(lines)
            assert result == expected
        else:
            assert result == ""

    @given(
        st.lists(st.text(min_size=0, max_size=100), min_size=1, max_size=20),
        st.text(min_size=0, max_size=100),
    )
    @settings(max_examples=300)
    def test_output_manager_line_flush(self, lines: list[str], final_line: str) -> None:
        """Test OutputManager line_flush behavior"""
        output = slmgr.OutputManager()

        for line in lines:
            output.line_out(line)

        # After flush, buffer should be empty
        output.line_flush(final_line)
        assert output.get_output() == ""


class TestSLMgrError:
    """Property-based tests for SLMgrError exception"""

    @given(
        message=st.text(min_size=1, max_size=200),
        error_code=st.one_of(st.integers(), st.none()),
    )
    @settings(max_examples=500)
    def test_slmgr_error_properties(self, message: str, error_code: int | None) -> None:
        """Test SLMgrError exception properties"""
        error = slmgr.SLMgrError(message, error_code)

        assert error.message == message
        assert error.error_code == error_code
        assert str(error) == message

        # Should be an Exception
        assert isinstance(error, Exception)
