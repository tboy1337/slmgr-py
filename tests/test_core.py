"""Tests for core classes in slmgr.py"""

from unittest.mock import Mock, patch

import pytest

import slmgr


class TestWMIConnection:
    """Test WMIConnection class"""

    @patch("slmgr.wmi.WMI")
    def test_wmi_connection_local_connect(self, mock_wmi: Mock) -> None:
        """Test local WMI connection"""
        output = slmgr.OutputManager()
        conn = slmgr.WMIConnection()

        mock_wmi_instance = Mock()
        mock_wmi.return_value = mock_wmi_instance
        mock_wmi_instance.StdRegProv = Mock()

        conn.connect(output)

        assert conn.wmi_service is not None
        assert conn.registry is not None
        assert not conn.is_remote

    @patch("slmgr.pythoncom")
    @patch("slmgr.wmi.WMI")
    def test_wmi_connection_remote_no_creds(
        self, mock_wmi: Mock, mock_com: Mock
    ) -> None:
        """Test remote WMI connection without credentials"""
        output = slmgr.OutputManager()
        conn = slmgr.WMIConnection("remote-pc")

        mock_wmi_instance = Mock()
        mock_wmi.return_value = mock_wmi_instance

        # Mock version check
        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_wmi_instance.query.return_value = [mock_service]
        mock_wmi_instance.StdRegProv = Mock()

        conn.connect(output)

        assert conn.is_remote
        mock_com.CoInitialize.assert_called_once()

    @patch("slmgr.pythoncom")
    @patch("slmgr.win32com.client.Dispatch")
    @patch("slmgr.wmi.WMI")
    def test_wmi_connection_remote_with_creds(
        self, mock_wmi: Mock, mock_dispatch: Mock, mock_com: Mock
    ) -> None:
        """Test remote WMI connection with credentials"""
        output = slmgr.OutputManager()
        conn = slmgr.WMIConnection("remote-pc", "user", "pass")

        mock_wmi_instance = Mock()
        mock_wmi.return_value = mock_wmi_instance

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_wmi_instance.query.return_value = [mock_service]

        # Mock registry connection
        mock_locator = Mock()
        mock_server = Mock()
        mock_server.Security_ = Mock()
        mock_server.Security_.ImpersonationLevel = 0
        mock_reg_prov = Mock()
        mock_server.Get.return_value = mock_reg_prov
        mock_locator.ConnectServer.return_value = mock_server
        mock_dispatch.return_value = mock_locator

        conn.connect(output)

        assert conn.is_remote
        assert conn.registry == mock_reg_prov

    @patch("slmgr.wmi.WMI")
    def test_wmi_connection_remote_version_error(self, mock_wmi: Mock) -> None:
        """Test remote WMI connection with unsupported version"""
        output = slmgr.OutputManager()
        conn = slmgr.WMIConnection("remote-pc")

        mock_wmi_instance = Mock()
        mock_wmi.return_value = mock_wmi_instance

        mock_service = Mock()
        mock_service.Version = "6.0"
        mock_wmi_instance.query.return_value = [mock_service]

        with pytest.raises(slmgr.SLMgrError):
            conn.connect(output)

    @patch("slmgr.wmi.WMI")
    def test_wmi_connection_error(self, mock_wmi: Mock) -> None:
        """Test WMI connection error"""
        output = slmgr.OutputManager()
        conn = slmgr.WMIConnection()

        mock_wmi.side_effect = Exception("Connection failed")

        with pytest.raises(slmgr.SLMgrError):
            conn.connect(output)

    @patch("slmgr.wmi.WMI")
    def test_get_service_object_success(self, mock_wmi: Mock) -> None:
        """Test get_service_object success"""
        conn = slmgr.WMIConnection()
        mock_service = Mock()
        mock_service.Version = "10.0"

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        result = conn.get_service_object("Version")
        assert result == mock_service

    def test_get_service_object_not_connected(self) -> None:
        """Test get_service_object when not connected"""
        conn = slmgr.WMIConnection()
        with pytest.raises(slmgr.SLMgrError):
            conn.get_service_object("Version")

    @patch("slmgr.wmi.WMI")
    def test_get_service_object_no_results(self, mock_wmi: Mock) -> None:
        """Test get_service_object with no results"""
        conn = slmgr.WMIConnection()
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = []

        with pytest.raises(slmgr.SLMgrError):
            conn.get_service_object("Version")

    def test_get_product_collection_with_where(self) -> None:
        """Test get_product_collection with where clause"""
        conn = slmgr.WMIConnection()
        mock_product = Mock()
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        result = conn.get_product_collection("ID", "ID = 'test'")
        assert len(result) == 1
        assert result[0] == mock_product

    def test_get_product_collection_without_where(self) -> None:
        """Test get_product_collection without where clause"""
        conn = slmgr.WMIConnection()
        mock_product = Mock()
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        result = conn.get_product_collection("ID", "")
        assert len(result) == 1

    def test_get_product_collection_not_connected(self) -> None:
        """Test get_product_collection when not connected"""
        conn = slmgr.WMIConnection()
        with pytest.raises(slmgr.SLMgrError):
            conn.get_product_collection("ID", "")

    def test_get_product_object_single(self) -> None:
        """Test get_product_object with single result"""
        conn = slmgr.WMIConnection()
        mock_product = Mock()
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        result = conn.get_product_object("ID", "ID = 'test'")
        assert result == mock_product

    def test_get_product_object_none(self) -> None:
        """Test get_product_object with no results"""
        conn = slmgr.WMIConnection()
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = []

        with pytest.raises(slmgr.SLMgrError) as exc_info:
            conn.get_product_object("ID", "ID = 'test'")
        assert exc_info.value.error_code == slmgr.HR_SL_E_PKEY_NOT_INSTALLED

    def test_get_product_object_multiple(self) -> None:
        """Test get_product_object with multiple results"""
        conn = slmgr.WMIConnection()
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [Mock(), Mock()]

        with pytest.raises(slmgr.SLMgrError) as exc_info:
            conn.get_product_object("ID", "ID = 'test'")
        assert exc_info.value.error_code == slmgr.HR_INVALID_ARG


class TestRegistryManager:
    """Test RegistryManager class"""

    @patch("slmgr.winreg.OpenKey")
    @patch("slmgr.winreg.SetValueEx")
    def test_set_string_value_local(self, mock_set: Mock, mock_open: Mock) -> None:
        """Test set_string_value locally"""
        conn = slmgr.WMIConnection()
        conn.is_remote = False
        reg = slmgr.RegistryManager(conn)

        mock_key = Mock()
        mock_open.return_value.__enter__.return_value = mock_key

        result = reg.set_string_value(
            slmgr.HKEY_LOCAL_MACHINE, "test\\path", "name", "value"
        )
        assert result == 0
        mock_set.assert_called_once()

    def test_set_string_value_remote(self) -> None:
        """Test set_string_value remotely"""
        conn = slmgr.WMIConnection("remote-pc")
        reg = slmgr.RegistryManager(conn)

        mock_registry = Mock()
        mock_registry.SetStringValue.return_value = 0
        conn.registry = mock_registry

        result = reg.set_string_value(
            slmgr.HKEY_LOCAL_MACHINE, "test\\path", "name", "value"
        )
        assert result == 0

    @patch("slmgr.winreg.OpenKey")
    @patch("slmgr.winreg.SetValueEx")
    def test_set_string_value_local_error(
        self, mock_set: Mock, mock_open: Mock
    ) -> None:
        """Test set_string_value local error"""
        conn = slmgr.WMIConnection()
        conn.is_remote = False
        reg = slmgr.RegistryManager(conn)

        mock_open.side_effect = Exception("Registry error")

        result = reg.set_string_value(
            slmgr.HKEY_LOCAL_MACHINE, "test\\path", "name", "value"
        )
        assert result == 1

    @patch("slmgr.winreg.OpenKey")
    @patch("slmgr.winreg.DeleteValue")
    def test_delete_value_local(self, mock_delete: Mock, mock_open: Mock) -> None:
        """Test delete_value locally"""
        conn = slmgr.WMIConnection()
        conn.is_remote = False
        reg = slmgr.RegistryManager(conn)

        mock_key = Mock()
        mock_open.return_value.__enter__.return_value = mock_key

        result = reg.delete_value(slmgr.HKEY_LOCAL_MACHINE, "test\\path", "name")
        assert result == 0
        mock_delete.assert_called_once()

    def test_delete_value_remote(self) -> None:
        """Test delete_value remotely"""
        conn = slmgr.WMIConnection("remote-pc")
        reg = slmgr.RegistryManager(conn)

        mock_registry = Mock()
        mock_registry.DeleteValue.return_value = 0
        conn.registry = mock_registry

        result = reg.delete_value(slmgr.HKEY_LOCAL_MACHINE, "test\\path", "name")
        assert result == 0

    @patch("slmgr.winreg.OpenKey")
    @patch("slmgr.winreg.DeleteValue")
    def test_delete_value_not_found(self, mock_delete: Mock, mock_open: Mock) -> None:
        """Test delete_value file not found"""
        conn = slmgr.WMIConnection()
        conn.is_remote = False
        reg = slmgr.RegistryManager(conn)

        mock_delete.side_effect = FileNotFoundError()

        result = reg.delete_value(slmgr.HKEY_LOCAL_MACHINE, "test\\path", "name")
        assert result == 2

    @patch("slmgr.winreg.OpenKey")
    @patch("slmgr.winreg.DeleteValue")
    def test_delete_value_error(self, mock_delete: Mock, mock_open: Mock) -> None:
        """Test delete_value error"""
        conn = slmgr.WMIConnection()
        conn.is_remote = False
        reg = slmgr.RegistryManager(conn)

        mock_delete.side_effect = Exception("Error")

        result = reg.delete_value(slmgr.HKEY_LOCAL_MACHINE, "test\\path", "name")
        assert result == 1

    @patch("slmgr.winreg.OpenKey")
    def test_key_exists_local_true(self, mock_open: Mock) -> None:
        """Test key_exists locally returns True"""
        conn = slmgr.WMIConnection()
        conn.is_remote = False
        reg = slmgr.RegistryManager(conn)

        mock_key = Mock()
        mock_open.return_value = mock_key

        result = reg.key_exists(slmgr.HKEY_LOCAL_MACHINE, "test\\path")
        assert result is True

    @patch("slmgr.winreg.OpenKey")
    def test_key_exists_local_false(self, mock_open: Mock) -> None:
        """Test key_exists locally returns False"""
        conn = slmgr.WMIConnection()
        conn.is_remote = False
        reg = slmgr.RegistryManager(conn)

        mock_open.side_effect = Exception("Not found")

        result = reg.key_exists(slmgr.HKEY_LOCAL_MACHINE, "test\\path")
        assert result is False

    def test_key_exists_remote_true(self) -> None:
        """Test key_exists remotely returns True"""
        conn = slmgr.WMIConnection("remote-pc")
        reg = slmgr.RegistryManager(conn)

        mock_registry = Mock()
        mock_registry.CheckAccess.return_value = [0]
        conn.registry = mock_registry

        result = reg.key_exists(slmgr.HKEY_LOCAL_MACHINE, "test\\path")
        assert result is True

    def test_key_exists_remote_false(self) -> None:
        """Test key_exists remotely returns False"""
        conn = slmgr.WMIConnection("remote-pc")
        reg = slmgr.RegistryManager(conn)

        mock_registry = Mock()
        mock_registry.CheckAccess.return_value = [2]
        conn.registry = mock_registry

        result = reg.key_exists(slmgr.HKEY_LOCAL_MACHINE, "test\\path")
        assert result is False
