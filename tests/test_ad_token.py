"""Tests for AD activation and Token-based activation functions"""

from unittest.mock import Mock, patch

import pytest

import slmgr


class TestTokenActivation:
    """Test Token-based Activation functions"""

    def test_tka_list_ils_empty(self) -> None:
        """Test tka_list_ils with no licenses"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = []

        slmgr.tka_list_ils(conn, output)

        result = output.get_output()
        assert "No licenses found" in result

    def test_tka_list_ils_with_licenses(self) -> None:
        """Test tka_list_ils with licenses"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_license = Mock()
        mock_license.ILID = "test-ilid"
        mock_license.ILVID = 1
        mock_license.ExpirationDate = "20241225143000.000000+000"
        mock_license.AdditionalInfo = "Test Info"
        mock_license.AuthorizationStatus = 0
        mock_license.Description = "Test License"

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_license]

        slmgr.tka_list_ils(conn, output)

        result = output.get_output()
        assert "test-ilid" in result
        assert "Test License" in result

    def test_tka_list_ils_with_error(self) -> None:
        """Test tka_list_ils with authorization error"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_license = Mock()
        mock_license.ILID = "test-ilid"
        mock_license.ILVID = 1
        mock_license.AuthorizationStatus = 0xC004C001

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_license]

        slmgr.tka_list_ils(conn, output)

        result = output.get_output()
        assert "Error: 0x" in result

    def test_tka_remove_il(self) -> None:
        """Test tka_remove_il"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_license = Mock()
        mock_license.ILID = "test-ilid"
        mock_license.ILVID = 1
        mock_license.ID = "slid-123"

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_license]

        slmgr.tka_remove_il(conn, output, "test-ilid", "1")

        mock_license.Uninstall.assert_called_once()
        result = output.get_output()
        assert "Removed" in result

    def test_tka_remove_il_not_found(self) -> None:
        """Test tka_remove_il when not found"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = []

        slmgr.tka_remove_il(conn, output, "test-ilid", "1")

        result = output.get_output()
        assert "No licenses found" in result

    @patch("slmgr.win32com.client.Dispatch")
    def test_tka_list_certs(self, mock_dispatch: Mock) -> None:
        """Test tka_list_certs"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_signer = Mock()
        mock_signer.GetCertificateThumbprints.return_value = [
            "thumbprint|subject|issuer|2023-01-01|2024-01-01"
        ]
        mock_dispatch.return_value = mock_signer

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.GetTokenActivationGrants.return_value = ["grant1"]

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.tka_list_certs(conn, output)

        result = output.get_output()
        assert "thumbprint" in result

    @patch("slmgr.win32com.client.Dispatch")
    def test_tka_list_certs_error(self, mock_dispatch: Mock) -> None:
        """Test tka_list_certs with error"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_dispatch.side_effect = Exception("Error creating signer")

        with pytest.raises(slmgr.SLMgrError):
            slmgr.tka_list_certs(conn, output)

    @patch("slmgr.win32com.client.Dispatch")
    def test_tka_activate(self, mock_dispatch: Mock) -> None:
        """Test tka_activate"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_signer = Mock()
        mock_signer.Sign.return_value = "auth-info"
        mock_dispatch.return_value = mock_signer

        mock_service = Mock()
        mock_service.Version = "10.0"

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.GenerateTokenActivationChallenge.return_value = "challenge"
        mock_product.LicenseStatus = 1

        conn.wmi_service = Mock()
        call_count = [0]

        def query_side_effect(q: str) -> list:
            call_count[0] += 1
            if "SoftwareLicensingService" in q:
                return [mock_service]
            return [mock_product]

        conn.wmi_service.query.side_effect = query_side_effect

        slmgr.tka_activate(conn, output, "thumbprint", "pin")

        result = output.get_output()
        assert "activated successfully" in result.lower()

    @patch("slmgr.win32com.client.Dispatch")
    def test_tka_activate_extended_grace(self, mock_dispatch: Mock) -> None:
        """Test tka_activate with extended grace period"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_signer = Mock()
        mock_signer.Sign.return_value = "auth-info"
        mock_dispatch.return_value = mock_signer

        mock_service = Mock()
        mock_service.Version = "10.0"

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.Name = "Test Product"
        mock_product.GenerateTokenActivationChallenge.return_value = "challenge"
        mock_product.LicenseStatus = 6

        conn.wmi_service = Mock()
        conn.wmi_service.query.side_effect = lambda q: (
            [mock_service] if "SoftwareLicensingService" in q else [mock_product]
        )

        slmgr.tka_activate(conn, output, "thumbprint", "")

        result = output.get_output()
        assert "Extended grace period" in result


class TestADActivation:
    """Test Active Directory activation functions"""

    def test_ad_activate_online(self) -> None:
        """Test ad_activate_online"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.ad_activate_online(
            conn, output, "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX", "AO-Name"
        )

        mock_service.DoActiveDirectoryOnlineActivation.assert_called_once()
        result = output.get_output()
        assert "activated successfully" in result.lower()

    def test_ad_activate_online_remote(self) -> None:
        """Test ad_activate_online on remote machine"""
        conn = slmgr.WMIConnection("remote-pc")
        output = slmgr.OutputManager()

        with pytest.raises(slmgr.SLMgrError):
            slmgr.ad_activate_online(conn, output, "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX")

    def test_ad_get_iid(self) -> None:
        """Test ad_get_iid"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.GenerateActiveDirectoryOfflineActivationId.return_value = (
            "123456789"
        )
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.ad_get_iid(conn, output, "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX")

        result = output.get_output()
        assert "123456789" in result
        assert "phone.inf" in result

    def test_ad_get_iid_remote(self) -> None:
        """Test ad_get_iid on remote machine"""
        conn = slmgr.WMIConnection("remote-pc")
        output = slmgr.OutputManager()

        with pytest.raises(slmgr.SLMgrError):
            slmgr.ad_get_iid(conn, output, "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX")

    def test_ad_activate_phone(self) -> None:
        """Test ad_activate_phone"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.ad_activate_phone(
            conn, output, "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX", "111111-222222", "AO-Name"
        )

        mock_service.DepositActiveDirectoryOfflineActivationConfirmation.assert_called_once()
        result = output.get_output()
        assert "activated successfully" in result.lower()

    @patch("slmgr.win32com.client.Dispatch")
    @patch("slmgr.win32com.client.GetObject")
    def test_ad_list_activation_objects(
        self, mock_get_object: Mock, mock_dispatch: Mock
    ) -> None:
        """Test ad_list_activation_objects"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_ad_sys_info = Mock()
        mock_ad_sys_info.DomainDNSName = "example.com"
        mock_dispatch.return_value = mock_ad_sys_info

        mock_root_dse = Mock()
        mock_root_dse.Get.return_value = "CN=Configuration,DC=example,DC=com"

        mock_container = Mock()
        mock_ao = Mock()
        mock_ao.Class = "msSPP-ActivationObject"
        guid_bytes = b"\x01\x02\x03\x04\x05\x06\x07\x08\x09\x0a\x0b\x0c\x0d\x0e\x0f\x10"
        mock_ao.Get.side_effect = lambda attr: {
            "displayName": "Test AO",
            "msSPP-CSVLKSkuId": guid_bytes,
            "msSPP-CSVLKPartialProductKey": "XXXXX",
            "msSPP-CSVLKPid": "extended-pid",
            "distinguishedName": "CN=TestAO,CN=Activation Objects",
        }[attr]
        mock_ao.GetInfoEx = Mock()

        mock_container.__iter__ = Mock(return_value=iter([mock_ao]))

        mock_namespace = Mock()
        mock_namespace.OpenDSObject.side_effect = [mock_root_dse, mock_container]
        mock_get_object.return_value = mock_namespace

        slmgr.ad_list_activation_objects(conn, output)

        result = output.get_output()
        assert "Activation Objects" in result
        assert "Test AO" in result

    @patch("slmgr.win32com.client.Dispatch")
    @patch("slmgr.win32com.client.GetObject")
    def test_ad_list_activation_objects_unsupported(
        self, mock_get_object: Mock, mock_dispatch: Mock
    ) -> None:
        """Test ad_list_activation_objects with unsupported schema"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_ad_sys_info = Mock()
        mock_ad_sys_info.DomainDNSName = "example.com"
        mock_dispatch.return_value = mock_ad_sys_info

        mock_root_dse = Mock()
        mock_root_dse.Get.return_value = "CN=Configuration,DC=example,DC=com"

        mock_namespace = Mock()
        mock_namespace.OpenDSObject.side_effect = [
            mock_root_dse,
            Exception("Container not found"),
        ]
        mock_get_object.return_value = mock_namespace

        slmgr.ad_list_activation_objects(conn, output)

        result = output.get_output()
        assert "not supported" in result

    @patch("slmgr.win32com.client.Dispatch")
    @patch("slmgr.win32com.client.GetObject")
    def test_ad_delete_activation_object(
        self, mock_get_object: Mock, mock_dispatch: Mock
    ) -> None:
        """Test ad_delete_activation_object"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_ad_sys_info = Mock()
        mock_ad_sys_info.DomainDNSName = "example.com"
        mock_dispatch.return_value = mock_ad_sys_info

        mock_root_dse = Mock()
        mock_root_dse.Get.return_value = "CN=Configuration,DC=example,DC=com"

        mock_container = Mock()

        mock_obj = Mock()
        mock_obj.Class = "msSPP-ActivationObject"
        mock_obj.Name = "CN=TestAO"
        mock_obj.Parent = "LDAP://CN=Activation Objects"

        mock_parent = Mock()
        mock_parent.Delete = Mock()

        def get_object_side_effect(path: str) -> Mock:
            if "rootDSE" in path:
                return mock_root_dse
            if "CN=TestAO" in path:
                return mock_obj
            if path == mock_obj.Parent:
                return mock_parent
            if "Activation Objects" in path:
                return mock_container
            return Mock()

        mock_namespace = Mock()
        mock_namespace.OpenDSObject.side_effect = [mock_root_dse, mock_container]

        mock_get_object.side_effect = get_object_side_effect

        slmgr.ad_delete_activation_object(conn, output, "TestAO")

        mock_parent.Delete.assert_called_once_with(
            "msSPP-ActivationObject", "CN=TestAO"
        )
        result = output.get_output()
        assert "Operation completed successfully" in result

    @patch("slmgr.win32com.client.Dispatch")
    @patch("slmgr.win32com.client.GetObject")
    def test_ad_delete_activation_object_with_cn(
        self, mock_get_object: Mock, mock_dispatch: Mock
    ) -> None:
        """Test ad_delete_activation_object with CN= prefix"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_ad_sys_info = Mock()
        mock_ad_sys_info.DomainDNSName = "example.com"
        mock_dispatch.return_value = mock_ad_sys_info

        mock_root_dse = Mock()
        mock_root_dse.Get.return_value = "CN=Configuration,DC=example,DC=com"

        mock_container = Mock()

        mock_obj = Mock()
        mock_obj.Class = "msSPP-ActivationObject"
        mock_obj.Name = "CN=TestAO"
        mock_obj.Parent = "LDAP://CN=Activation Objects"

        mock_parent = Mock()

        def get_object_side_effect(path: str) -> Mock:
            if "rootDSE" in path:
                return mock_root_dse
            if "Activation Objects" in path and "CN=TestAO" not in path:
                return mock_container
            if "CN=TestAO" in path:
                return mock_obj
            if path == mock_obj.Parent:
                return mock_parent
            return Mock()

        mock_namespace = Mock()
        mock_namespace.OpenDSObject.side_effect = [mock_root_dse, mock_container]

        mock_get_object.side_effect = get_object_side_effect

        slmgr.ad_delete_activation_object(conn, output, "CN=TestAO")

        result = output.get_output()
        assert "Operation completed successfully" in result

    @patch("slmgr.win32com.client.Dispatch")
    @patch("slmgr.win32com.client.GetObject")
    def test_ad_delete_activation_object_full_dn(
        self, mock_get_object: Mock, mock_dispatch: Mock
    ) -> None:
        """Test ad_delete_activation_object with full DN"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_ad_sys_info = Mock()
        mock_ad_sys_info.DomainDNSName = "example.com"
        mock_dispatch.return_value = mock_ad_sys_info

        mock_root_dse = Mock()
        mock_root_dse.Get.return_value = "CN=Configuration,DC=example,DC=com"

        mock_container = Mock()

        mock_obj = Mock()
        mock_obj.Class = "msSPP-ActivationObject"
        mock_obj.Name = "CN=TestAO"
        mock_obj.Parent = "LDAP://CN=Activation Objects"

        mock_parent = Mock()

        def get_object_side_effect(path: str) -> Mock:
            if "rootDSE" in path:
                return mock_root_dse
            if "Activation Objects" in path and "CN=TestAO" not in path:
                return mock_container
            if "CN=TestAO" in path:
                return mock_obj
            if path == mock_obj.Parent:
                return mock_parent
            return Mock()

        mock_namespace = Mock()
        mock_namespace.OpenDSObject.side_effect = [mock_root_dse, mock_container]

        mock_get_object.side_effect = get_object_side_effect

        slmgr.ad_delete_activation_object(
            conn, output, "CN=TestAO,CN=Activation Objects"
        )

        result = output.get_output()
        assert "Operation completed successfully" in result
