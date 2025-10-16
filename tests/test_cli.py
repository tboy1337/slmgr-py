"""Tests for CLI parsing and main entry point"""

from unittest.mock import Mock, patch

import pytest

import slmgr


class TestCLIParsing:
    """Test CLI argument parsing"""

    def test_parse_arguments_no_args(self) -> None:
        """Test parse_arguments with no arguments"""
        with patch("sys.argv", ["slmgr.py"]):
            with pytest.raises(SystemExit):
                slmgr.parse_arguments()

    def test_parse_arguments_local_command(self) -> None:
        """Test parse_arguments with local command"""
        with patch("sys.argv", ["slmgr.py", "/dli"]):
            computer, username, password, command_args = slmgr.parse_arguments()
            assert computer == "."
            assert username == ""
            assert password == ""
            assert command_args == ["/dli"]

    def test_parse_arguments_remote_no_creds(self) -> None:
        """Test parse_arguments with remote computer"""
        with patch("sys.argv", ["slmgr.py", "remote-pc", "/dli"]):
            computer, username, password, command_args = slmgr.parse_arguments()
            assert computer == "remote-pc"
            assert username == ""
            assert password == ""
            assert command_args == ["/dli"]

    def test_parse_arguments_remote_with_creds(self) -> None:
        """Test parse_arguments with remote computer and credentials"""
        with patch("sys.argv", ["slmgr.py", "remote-pc", "user", "pass", "/dli"]):
            computer, username, password, command_args = slmgr.parse_arguments()
            assert computer == "remote-pc"
            assert username == "user"
            assert password == "pass"
            assert command_args == ["/dli"]

    def test_parse_arguments_dash_prefix(self) -> None:
        """Test parse_arguments with dash prefix"""
        with patch("sys.argv", ["slmgr.py", "-dli"]):
            computer, username, password, command_args = slmgr.parse_arguments()
            assert command_args == ["-dli"]

    def test_display_usage(self, capsys: pytest.CaptureFixture[str]) -> None:
        """Test display_usage"""
        slmgr.display_usage()
        captured = capsys.readouterr()
        assert "Windows Software Licensing Management Tool" in captured.out
        assert "/ipk" in captured.out
        assert "/ato" in captured.out

    @patch("slmgr.WMIConnection")
    @patch("slmgr.RegistryManager")
    def test_execute_command_ipk(
        self, mock_reg_class: Mock, mock_conn_class: Mock
    ) -> None:
        """Test execute_command with /ipk"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_product = Mock()
        mock_product.Description = "Test Product"

        conn.wmi_service = Mock()
        conn.get_service_object = Mock(return_value=mock_service)
        conn.get_product_collection = Mock(return_value=[mock_product])
        reg.set_string_value = Mock(return_value=0)
        reg.delete_value = Mock(return_value=0)
        reg.key_exists = Mock(return_value=False)

        slmgr.execute_command(
            conn, reg, output, ["/ipk", "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"]
        )

        # Should not raise any error
        result = output.get_output()
        assert "successfully" in result.lower()

    def test_execute_command_ipk_missing_arg(self) -> None:
        """Test execute_command /ipk without product key"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        with pytest.raises(SystemExit):
            slmgr.execute_command(conn, reg, output, ["/ipk"])

    def test_execute_command_upk(self) -> None:
        """Test execute_command with /upk"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_product = Mock()
        mock_product.Uninstall = Mock()
        mock_product.Description = "Test Product"

        conn.wmi_service = Mock()
        conn.get_service_object = Mock(return_value=mock_service)
        conn.get_product_collection = Mock(return_value=[mock_product])
        reg.delete_value = Mock(return_value=0)

        slmgr.execute_command(conn, reg, output, ["/upk"])

    def test_execute_command_dti(self) -> None:
        """Test execute_command with /dti"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.OfflineInstallationId = "123456789"

        conn.wmi_service = Mock()
        conn.get_product_collection = Mock(return_value=[mock_product])

        slmgr.execute_command(conn, reg, output, ["/dti"])

    def test_execute_command_ato(self) -> None:
        """Test execute_command with /ato"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.Name = "Test Product"
        mock_product.Description = "Test Description"
        mock_product.LicenseStatus = 0
        mock_product.Activate = Mock()

        conn.wmi_service = Mock()
        conn.get_product_collection = Mock(return_value=[mock_product])

        slmgr.execute_command(conn, reg, output, ["/ato"])

    def test_execute_command_atp_missing_cid(self) -> None:
        """Test execute_command /atp without confirmation ID"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        with pytest.raises(SystemExit):
            slmgr.execute_command(conn, reg, output, ["/atp"])

    def test_execute_command_dli(self) -> None:
        """Test execute_command with /dli"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.KeyManagementServiceListeningPort = 1688
        mock_service.KeyManagementServiceDnsPublishing = True
        mock_service.KeyManagementServiceLowPriority = False
        mock_service.KeyManagementServiceHostCaching = True
        mock_service.ClientMachineId = "test-client-id"

        # Create a simple object with attributes instead of Mock
        class MockProduct:
            Name = "Test Product"
            Description = "Test Description"
            ApplicationId = slmgr.WINDOWS_APP_ID.lower()
            ApplicationID = slmgr.WINDOWS_APP_ID.lower()  # Note: capital D for verbose
            PartialProductKey = "XXXXX"
            LicenseIsAddon = False
            LicenseStatus = 1
            ID = "test-id"
            GracePeriodRemaining = 0
            ProductKeyID = "extended-pid"
            ProductKeyChannel = "Retail"
            OfflineInstallationId = "123456"
            IsKeyManagementServiceMachine = 0

        mock_product = MockProduct()

        # Mock the KMS product check
        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 0

        def get_product_side_effect(select: str, where: str = "") -> list:
            if "IsKeyManagementServiceMachine" in select:
                return [mock_kms_product]
            return [mock_product]

        conn.wmi_service = Mock()
        conn.get_service_object = Mock(return_value=mock_service)
        conn.get_product_collection = Mock(side_effect=get_product_side_effect)
        conn.get_product_object = Mock(return_value=mock_product)

        slmgr.execute_command(conn, reg, output, ["/dli"])

    def test_execute_command_dlv(self) -> None:
        """Test execute_command with /dlv"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        mock_service.KeyManagementServiceListeningPort = 1688
        mock_service.KeyManagementServiceDnsPublishing = True
        mock_service.KeyManagementServiceLowPriority = False
        mock_service.KeyManagementServiceHostCaching = True
        mock_service.ClientMachineId = "test-client-id"
        mock_service.RemainingWindowsReArmCount = 3

        # Create a simple object with attributes instead of Mock
        class MockProduct:
            Name = "Test Product"
            Description = "Test Description"
            ApplicationId = slmgr.WINDOWS_APP_ID.lower()
            ApplicationID = slmgr.WINDOWS_APP_ID.lower()  # Note: capital D for verbose
            PartialProductKey = "XXXXX"
            LicenseIsAddon = False
            LicenseStatus = 1
            ID = "test-id"
            GracePeriodRemaining = 0
            ProductKeyID = "extended-pid"
            ProductKeyChannel = "Retail"
            OfflineInstallationId = "123456"
            EvaluationEndDate = "20241225143000.000000+000"
            TrustedTime = "20231225143000.000000+000"
            RemainingAppReArmCount = 2
            RemainingSkuReArmCount = 1
            IsKeyManagementServiceMachine = 0

        mock_product = MockProduct()

        # Mock the KMS product check
        mock_kms_product = Mock()
        mock_kms_product.IsKeyManagementServiceMachine = 0

        def get_product_side_effect(select: str, where: str = "") -> list:
            if "IsKeyManagementServiceMachine" in select:
                return [mock_kms_product]
            return [mock_product]

        conn.wmi_service = Mock()
        conn.get_service_object = Mock(return_value=mock_service)
        conn.get_product_collection = Mock(side_effect=get_product_side_effect)
        conn.get_product_object = Mock(return_value=mock_product)

        slmgr.execute_command(conn, reg, output, ["/dlv"])

    def test_execute_command_xpr(self) -> None:
        """Test execute_command with /xpr"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.Name = "Test Product"
        mock_product.LicenseStatus = 1
        mock_product.GracePeriodRemaining = 0

        conn.wmi_service = Mock()
        conn.get_product_collection = Mock(return_value=[mock_product])

        slmgr.execute_command(conn, reg, output, ["/xpr"])

    def test_execute_command_cpky(self) -> None:
        """Test execute_command with /cpky"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_service = Mock()
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.execute_command(conn, reg, output, ["/cpky"])

    def test_execute_command_ilc_missing_file(self) -> None:
        """Test execute_command /ilc without license file"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        with pytest.raises(SystemExit):
            slmgr.execute_command(conn, reg, output, ["/ilc"])

    @patch("os.path.exists")
    @patch("os.walk")
    @patch("builtins.open")
    def test_execute_command_rilc(
        self, mock_open_file: Mock, mock_walk: Mock, mock_exists: Mock
    ) -> None:
        """Test execute_command with /rilc"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_service = Mock()
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]
        mock_exists.return_value = False

        slmgr.execute_command(conn, reg, output, ["/rilc"])

    def test_execute_command_rearm(self) -> None:
        """Test execute_command with /rearm"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_service = Mock()
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.execute_command(conn, reg, output, ["/rearm"])

    def test_execute_command_rearm_app_missing_id(self) -> None:
        """Test execute_command /rearm-app without app ID"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        with pytest.raises(SystemExit):
            slmgr.execute_command(conn, reg, output, ["/rearm-app"])

    def test_execute_command_rearm_sku_missing_id(self) -> None:
        """Test execute_command /rearm-sku without activation ID"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        with pytest.raises(SystemExit):
            slmgr.execute_command(conn, reg, output, ["/rearm-sku"])

    def test_execute_command_skms_missing_name(self) -> None:
        """Test execute_command /skms without name"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        with pytest.raises(SystemExit):
            slmgr.execute_command(conn, reg, output, ["/skms"])

    def test_execute_command_ckms(self) -> None:
        """Test execute_command with /ckms"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.KeyManagementServiceMachine = ""
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.execute_command(conn, reg, output, ["/ckms"])

    def test_execute_command_skhc(self) -> None:
        """Test execute_command with /skhc"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_service = Mock()
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.execute_command(conn, reg, output, ["/skhc"])

    def test_execute_command_ckhc(self) -> None:
        """Test execute_command with /ckhc"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_service = Mock()
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.execute_command(conn, reg, output, ["/ckhc"])

    def test_execute_command_sprt_missing_port(self) -> None:
        """Test execute_command /sprt without port"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        with pytest.raises(SystemExit):
            slmgr.execute_command(conn, reg, output, ["/sprt"])

    def test_execute_command_sai_missing_interval(self) -> None:
        """Test execute_command /sai without interval"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        with pytest.raises(SystemExit):
            slmgr.execute_command(conn, reg, output, ["/sai"])

    def test_execute_command_sri_missing_interval(self) -> None:
        """Test execute_command /sri without interval"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        with pytest.raises(SystemExit):
            slmgr.execute_command(conn, reg, output, ["/sri"])

    def test_execute_command_sdns(self) -> None:
        """Test execute_command with /sdns"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.DisableKeyManagementServiceDnsPublishing = Mock()

        mock_product = Mock()
        mock_product.IsKeyManagementServiceMachine = 1

        conn.wmi_service = Mock()
        conn.get_service_object = Mock(return_value=mock_service)
        conn.get_product_collection = Mock(return_value=[mock_product])

        slmgr.execute_command(conn, reg, output, ["/sdns"])

    def test_execute_command_cdns(self) -> None:
        """Test execute_command with /cdns"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.DisableKeyManagementServiceDnsPublishing = Mock()

        mock_product = Mock()
        mock_product.IsKeyManagementServiceMachine = 1

        conn.wmi_service = Mock()
        conn.get_service_object = Mock(return_value=mock_service)
        conn.get_product_collection = Mock(return_value=[mock_product])

        slmgr.execute_command(conn, reg, output, ["/cdns"])

    def test_execute_command_spri(self) -> None:
        """Test execute_command with /spri"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.SetKeyManagementServiceLowPriority = Mock()

        mock_product = Mock()
        mock_product.IsKeyManagementServiceMachine = 1

        conn.wmi_service = Mock()
        conn.get_service_object = Mock(return_value=mock_service)
        conn.get_product_collection = Mock(return_value=[mock_product])

        slmgr.execute_command(conn, reg, output, ["/spri"])

    def test_execute_command_cpri(self) -> None:
        """Test execute_command with /cpri"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.SetKeyManagementServiceLowPriority = Mock()

        mock_product = Mock()
        mock_product.IsKeyManagementServiceMachine = 1

        conn.wmi_service = Mock()
        conn.get_service_object = Mock(return_value=mock_service)
        conn.get_product_collection = Mock(return_value=[mock_product])

        slmgr.execute_command(conn, reg, output, ["/cpri"])

    def test_execute_command_act_type_no_params(self) -> None:
        """Test execute_command /act-type without params"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_service = Mock()
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.execute_command(conn, reg, output, ["/act-type"])

    def test_execute_command_act_type_with_type(self) -> None:
        """Test execute_command /act-type with type"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        mock_service = Mock()
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.execute_command(conn, reg, output, ["/act-type", "2"])

    def test_execute_command_lil(self) -> None:
        """Test execute_command with /lil"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = []

        slmgr.execute_command(conn, reg, output, ["/lil"])

    def test_execute_command_ril_missing_ilid(self) -> None:
        """Test execute_command /ril without ILID and ILvID"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        with pytest.raises(SystemExit):
            slmgr.execute_command(conn, reg, output, ["/ril", "ilid"])

    def test_execute_command_unrecognized(
        self, capsys: pytest.CaptureFixture[str]
    ) -> None:
        """Test execute_command with unrecognized option"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        with pytest.raises(SystemExit):
            slmgr.execute_command(conn, reg, output, ["/invalid"])

    def test_execute_command_no_args(self, capsys: pytest.CaptureFixture[str]) -> None:
        """Test execute_command with no arguments"""
        conn = Mock()
        reg = Mock()
        output = slmgr.OutputManager()

        slmgr.execute_command(conn, reg, output, [])

        captured = capsys.readouterr()
        assert "Windows Software Licensing Management Tool" in captured.out

    @patch("slmgr.parse_arguments")
    @patch("slmgr.WMIConnection")
    @patch("slmgr.RegistryManager")
    def test_main_help(
        self,
        mock_reg_class: Mock,
        mock_conn_class: Mock,
        mock_parse: Mock,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """Test main with help option"""
        mock_parse.return_value = (".", "", "", ["/?"])

        slmgr.main()

        captured = capsys.readouterr()
        assert "Windows Software Licensing Management Tool" in captured.out

    @patch("slmgr.parse_arguments")
    def test_main_keyboard_interrupt(self, mock_parse: Mock) -> None:
        """Test main with keyboard interrupt"""
        mock_parse.side_effect = KeyboardInterrupt()

        with pytest.raises(SystemExit):
            slmgr.main()

    @patch("slmgr.parse_arguments")
    @patch("slmgr.WMIConnection")
    def test_main_slmgr_error(self, mock_conn_class: Mock, mock_parse: Mock) -> None:
        """Test main with SLMgrError"""
        mock_parse.return_value = (".", "", "", ["/dli"])
        mock_conn = Mock()
        mock_conn_class.return_value = mock_conn
        mock_conn.connect.side_effect = slmgr.SLMgrError("Test error", 123)

        with pytest.raises(SystemExit):
            slmgr.main()

    @patch("slmgr.parse_arguments")
    @patch("slmgr.WMIConnection")
    def test_main_generic_error(self, mock_conn_class: Mock, mock_parse: Mock) -> None:
        """Test main with generic error"""
        mock_parse.return_value = (".", "", "", ["/dli"])
        mock_conn = Mock()
        mock_conn_class.return_value = mock_conn
        mock_conn.connect.side_effect = Exception("Generic error")

        with pytest.raises(SystemExit):
            slmgr.main()

    def test_execute_command_atp_with_activation_id(self) -> None:
        """Test execute_command /atp with activation ID"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
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
        mock_product.LicenseStatus = 1
        mock_product.DepositOfflineConfirmationId = Mock()

        conn.wmi_service = Mock()
        conn.get_service_object = Mock(return_value=mock_service)
        conn.get_product_collection = Mock(return_value=[mock_product])
        conn.get_product_object = Mock(return_value=mock_product)

        slmgr.execute_command(conn, reg, output, ["/atp", "123456", "test-id"])

        result = output.get_output()
        assert "successfully" in result.lower() or "deposited" in result.lower()

    def test_execute_command_skms_with_activation_id(self) -> None:
        """Test execute_command /skms with activation ID"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.SetKeyManagementServiceMachine = Mock()

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.execute_command(
            conn, reg, output, ["/skms", "kms.example.com:1688", "test-id"]
        )

        result = output.get_output()
        assert "key management service machine name" in result.lower()

    def test_execute_command_skms_domain_with_activation_id(self) -> None:
        """Test execute_command /skms-domain with activation ID"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.SetKeyManagementServiceLookupDomain = Mock()

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.execute_command(
            conn, reg, output, ["/skms-domain", "example.com", "test-id"]
        )

        result = output.get_output()
        assert "key management service lookup domain" in result.lower()

    def test_execute_command_ckms_domain_with_activation_id(self) -> None:
        """Test execute_command /ckms-domain with activation ID"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = "test-id"
        mock_product.ClearKeyManagementServiceLookupDomain = Mock()

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.execute_command(conn, reg, output, ["/ckms-domain", "test-id"])

        result = output.get_output()
        assert "cleared" in result.lower()

    def test_execute_command_sprt(self) -> None:
        """Test execute_command /sprt"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.SetKeyManagementServiceListeningPort = Mock()

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.execute_command(conn, reg, output, ["/sprt", "1688"])

        result = output.get_output()
        assert "port" in result.lower()

    def test_execute_command_sai(self) -> None:
        """Test execute_command /sai"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = slmgr.WINDOWS_APP_ID.lower()
        mock_product.SetVLActivationInterval = Mock()

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.execute_command(conn, reg, output, ["/sai", "120"])

        result = output.get_output()
        assert "activation interval" in result.lower()

    def test_execute_command_sri(self) -> None:
        """Test execute_command /sri"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = slmgr.WINDOWS_APP_ID.lower()
        mock_product.SetVLRenewalInterval = Mock()

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        slmgr.execute_command(conn, reg, output, ["/sri", "10080"])

        result = output.get_output()
        assert "renewal interval" in result.lower()

    def test_execute_command_fta_with_pin(self) -> None:
        """Test execute_command /fta with PIN"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_product = Mock()
        mock_product.ID = slmgr.WINDOWS_APP_ID.lower()
        mock_product.Name = "Test Product"
        mock_product.DepositTokenActivationResponse = Mock()

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_product]

        with patch("win32com.client.Dispatch") as mock_dispatch:
            mock_ils = Mock()
            mock_ils.GetInstalledLicenses = Mock(return_value=[])
            mock_dispatch.return_value = mock_ils

            slmgr.execute_command(conn, reg, output, ["/fta", "ABCD1234", "1234"])

            result = output.get_output()
            # The test covers the code path, just check it ran
            assert "activating" in result.lower() or "not found" in result.lower()

    def test_execute_command_ad_activation_online_with_ao_name(self) -> None:
        """Test execute_command /ad-activation-online with AO name"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.InstallProductKey = Mock()

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.execute_command(
            conn,
            reg,
            output,
            ["/ad-activation-online", "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX", "TestAO"],
        )

        result = output.get_output()
        # The test covers the code path with ao_name parameter
        assert "activated" in result.lower()

    def test_execute_command_ad_activation_apply_cid_with_ao_name(self) -> None:
        """Test execute_command /ad-activation-apply-cid with AO name"""
        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.InstallProductKey = Mock()

        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.execute_command(
            conn,
            reg,
            output,
            [
                "/ad-activation-apply-cid",
                "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX",
                "123456",
                "TestAO",
            ],
        )

        result = output.get_output()
        # The test covers the code path with ao_name parameter
        assert "activated" in result.lower()
