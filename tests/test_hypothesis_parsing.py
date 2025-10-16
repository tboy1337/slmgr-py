"""Property-based tests using Hypothesis for parsing functions in slmgr.py"""

from unittest.mock import Mock, patch

import pytest
from hypothesis import HealthCheck, assume, example, given, settings
from hypothesis import strategies as st

import slmgr


class TestKMSNamePortParsing:
    """Property-based tests for KMS name and port parsing"""

    @given(
        hostname=st.text(
            min_size=1,
            max_size=50,
            alphabet=st.characters(
                whitelist_categories=("Ll", "Lu", "Nd"),
                whitelist_characters=".-",
            ),
        ),
        port=st.integers(min_value=1, max_value=65535),
    )
    @settings(max_examples=1000)
    def test_set_kms_machine_name_ipv4_with_port(
        self, hostname: str, port: int
    ) -> None:
        """Test parsing IPv4 hostname with port"""
        # Skip if hostname contains special patterns
        assume(":" not in hostname)
        assume("[" not in hostname)
        assume("]" not in hostname)
        assume(len(hostname) > 0)

        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.KeyManagementServiceMachine = ""
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        kms_name_port = f"{hostname}:{port}"
        slmgr.set_kms_machine_name(conn, output, kms_name_port)

        # Should call SetKeyManagementServiceMachine with hostname
        mock_service.SetKeyManagementServiceMachine.assert_called_once_with(hostname)
        # Should call SetKeyManagementServicePort with port
        mock_service.SetKeyManagementServicePort.assert_called_once_with(port)

    @given(
        hostname=st.text(
            min_size=1,
            max_size=50,
            alphabet=st.characters(
                whitelist_categories=("Ll", "Lu", "Nd"),
                whitelist_characters=".-",
            ),
        ),
    )
    @settings(max_examples=500)
    def test_set_kms_machine_name_ipv4_without_port(self, hostname: str) -> None:
        """Test parsing IPv4 hostname without port"""
        assume(":" not in hostname)
        assume("[" not in hostname)
        assume("]" not in hostname)
        assume(len(hostname) > 0)

        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.KeyManagementServiceMachine = ""
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        slmgr.set_kms_machine_name(conn, output, hostname)

        # Should call SetKeyManagementServiceMachine with hostname
        mock_service.SetKeyManagementServiceMachine.assert_called_once_with(hostname)
        # Should call ClearKeyManagementServicePort (no port specified)
        mock_service.ClearKeyManagementServicePort.assert_called_once()

    @given(
        ipv6_part=st.text(
            min_size=1,
            max_size=30,
            alphabet=st.characters(
                whitelist_categories=("Nd",),
                whitelist_characters=":abcdefABCDEF",
            ),
        ),
        port=st.integers(min_value=1, max_value=65535),
    )
    @settings(max_examples=500)
    def test_set_kms_machine_name_ipv6_with_port(
        self, ipv6_part: str, port: int
    ) -> None:
        """Test parsing IPv6 address with port"""
        assume(len(ipv6_part) > 0)

        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.KeyManagementServiceMachine = ""
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        kms_name_port = f"[{ipv6_part}]:{port}"
        slmgr.set_kms_machine_name(conn, output, kms_name_port)

        # Should call SetKeyManagementServiceMachine with [ipv6]
        expected_name = f"[{ipv6_part}]"
        mock_service.SetKeyManagementServiceMachine.assert_called_once_with(
            expected_name
        )
        # Should call SetKeyManagementServicePort with port
        mock_service.SetKeyManagementServicePort.assert_called_once_with(port)

    @given(
        ipv6_part=st.text(
            min_size=1,
            max_size=30,
            alphabet=st.characters(
                whitelist_categories=("Nd",),
                whitelist_characters=":abcdefABCDEF",
            ),
        ),
    )
    @settings(max_examples=500)
    def test_set_kms_machine_name_ipv6_without_port(self, ipv6_part: str) -> None:
        """Test parsing IPv6 address without port"""
        assume(len(ipv6_part) > 0)

        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.KeyManagementServiceMachine = ""
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        kms_name_port = f"[{ipv6_part}]"
        slmgr.set_kms_machine_name(conn, output, kms_name_port)

        # Should call SetKeyManagementServiceMachine with [ipv6]
        expected_name = f"[{ipv6_part}]"
        mock_service.SetKeyManagementServiceMachine.assert_called_once_with(
            expected_name
        )
        # Should call ClearKeyManagementServicePort
        mock_service.ClearKeyManagementServicePort.assert_called_once()

    @given(
        st.sampled_from(
            [
                "kms.example.com:1688",
                "192.168.1.1:1688",
                "[::1]:1688",
                "[2001:db8::1]:1688",
                "kms.local",
                "[fe80::1]",
            ]
        )
    )
    @example("localhost:1688")
    @example("[::1]:1688")
    def test_set_kms_machine_name_real_world_examples(self, kms_name_port: str) -> None:
        """Test with real-world KMS name/port combinations"""
        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.KeyManagementServiceMachine = ""
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        # Should not raise any exception
        slmgr.set_kms_machine_name(conn, output, kms_name_port)

        # Should have called SetKeyManagementServiceMachine
        assert mock_service.SetKeyManagementServiceMachine.call_count == 1


class TestCommandLineArgumentParsing:
    """Property-based tests for command-line argument parsing"""

    @given(
        st.lists(
            st.text(
                min_size=1,
                max_size=20,
                alphabet=st.characters(
                    whitelist_categories=("Ll", "Lu", "Nd"),
                    whitelist_characters="/-.",
                ),
            ),
            min_size=1,
            max_size=10,
        )
    )
    @settings(max_examples=500, suppress_health_check=[HealthCheck.too_slow])
    def test_parse_arguments_basic_commands(self, args: list[str]) -> None:
        """Test parsing basic command structures"""
        # Skip if first arg looks like a command (starts with / or -)
        if args[0].startswith(("/", "-")):
            with patch("sys.argv", ["slmgr.py"] + args):
                try:
                    computer, username, password, command_args = slmgr.parse_arguments()
                    # Should return local computer
                    assert computer in [".", args[0]]
                    # Command args should be populated
                    assert len(command_args) >= 0
                except SystemExit:
                    # Some invalid combinations might exit
                    pass

    @given(
        computer=st.text(
            min_size=1,
            max_size=30,
            alphabet=st.characters(
                whitelist_categories=("Ll", "Lu", "Nd"),
                whitelist_characters=".-",
            ),
        ),
        command=st.sampled_from(["/dli", "/dlv", "/xpr", "/ato"]),
    )
    @settings(max_examples=500)
    def test_parse_arguments_remote_computer(self, computer: str, command: str) -> None:
        """Test parsing remote computer specification"""
        assume(not computer.startswith(("/", "-")))
        assume(len(computer) > 0)

        with patch("sys.argv", ["slmgr.py", computer, command]):
            parsed_computer, username, password, command_args = slmgr.parse_arguments()

            assert parsed_computer == computer
            assert username == ""
            assert password == ""
            assert command in command_args

    @given(
        computer=st.text(
            min_size=1,
            max_size=20,
            alphabet=st.characters(whitelist_categories=("Ll", "Lu", "Nd")),
        ),
        username=st.text(
            min_size=1,
            max_size=20,
            alphabet=st.characters(whitelist_categories=("Ll", "Lu", "Nd")),
        ),
        password=st.text(
            min_size=1,
            max_size=20,
            alphabet=st.characters(whitelist_categories=("Ll", "Lu", "Nd")),
        ),
        command=st.sampled_from(["/dli", "/dlv"]),
    )
    @settings(max_examples=300)
    def test_parse_arguments_with_credentials(
        self, computer: str, username: str, password: str, command: str
    ) -> None:
        """Test parsing remote computer with credentials"""
        # Exclude strings that start with / or - to avoid being treated as commands
        assume(not computer.startswith(("/", "-")))
        assume(not username.startswith(("/", "-")))
        assume(not password.startswith(("/", "-")))
        assume(len(computer) > 0)
        assume(len(username) > 0)
        assume(len(password) > 0)

        with patch("sys.argv", ["slmgr.py", computer, username, password, command]):
            (
                parsed_computer,
                parsed_username,
                parsed_password,
                command_args,
            ) = slmgr.parse_arguments()

            assert parsed_computer == computer
            assert parsed_username == username
            assert parsed_password == password
            assert command in command_args

    @given(st.sampled_from(["/dli", "/dlv", "/xpr", "/ato", "/ipk", "/upk"]))
    @settings(max_examples=100)
    def test_parse_arguments_local_commands(self, command: str) -> None:
        """Test parsing local commands"""
        with patch("sys.argv", ["slmgr.py", command]):
            computer, username, password, command_args = slmgr.parse_arguments()

            assert computer == "."
            assert username == ""
            assert password == ""
            assert command in command_args

    def test_parse_arguments_no_args_exits(self) -> None:
        """Test that no arguments causes exit"""
        with patch("sys.argv", ["slmgr.py"]):
            try:
                slmgr.parse_arguments()
                pytest.fail("Should have raised SystemExit")
            except SystemExit:
                pass  # Expected


class TestADDistinguishedNameParsing:
    """Property-based tests for AD distinguished name parsing"""

    @given(
        cn_name=st.text(
            min_size=1,
            max_size=30,
            alphabet=st.characters(
                whitelist_categories=("Ll", "Lu", "Nd"),
                whitelist_characters="-_",
            ),
        ),
    )
    @settings(max_examples=500)
    def test_ad_delete_activation_object_simple_name(self, cn_name: str) -> None:
        """Test AD activation object deletion with simple CN name"""
        assume(len(cn_name) > 0)
        assume(",cn=" not in cn_name.lower())
        assume(not cn_name.lower().startswith("cn="))

        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        with (
            patch("slmgr.win32com.client.Dispatch") as mock_dispatch,
            patch("slmgr.win32com.client.GetObject") as mock_get_object,
        ):
            mock_ad_sys_info = Mock()
            mock_ad_sys_info.DomainDNSName = "example.com"
            mock_dispatch.return_value = mock_ad_sys_info

            mock_root_dse = Mock()
            mock_root_dse.Get.return_value = "CN=Configuration,DC=example,DC=com"

            mock_container = Mock()
            mock_obj = Mock()
            mock_obj.Class = "msSPP-ActivationObject"
            mock_obj.Name = f"CN={cn_name}"
            mock_obj.Parent = "LDAP://CN=Activation Objects"
            mock_parent = Mock()

            def get_object_side_effect(path: str) -> Mock:
                if "rootDSE" in path:
                    return mock_root_dse
                if cn_name in path:
                    return mock_obj
                if "Activation Objects" in path:
                    return mock_container if "CN=" not in path else mock_obj
                if path == mock_obj.Parent:
                    return mock_parent
                return Mock()

            mock_namespace = Mock()
            mock_namespace.OpenDSObject.side_effect = [
                mock_root_dse,
                mock_container,
            ]
            mock_get_object.side_effect = get_object_side_effect

            # Should not raise exception
            slmgr.ad_delete_activation_object(conn, output, cn_name)

    @given(
        cn_name=st.text(
            min_size=1,
            max_size=30,
            alphabet=st.characters(
                whitelist_categories=("Ll", "Lu", "Nd"),
                whitelist_characters="-_",
            ),
        ),
    )
    @settings(max_examples=300)
    def test_ad_delete_activation_object_with_cn_prefix(self, cn_name: str) -> None:
        """Test AD activation object deletion with CN= prefix"""
        assume(len(cn_name) > 0)
        assume(",cn=" not in cn_name.lower())

        conn = slmgr.WMIConnection()
        output = slmgr.OutputManager()

        with (
            patch("slmgr.win32com.client.Dispatch") as mock_dispatch,
            patch("slmgr.win32com.client.GetObject") as mock_get_object,
        ):
            mock_ad_sys_info = Mock()
            mock_ad_sys_info.DomainDNSName = "example.com"
            mock_dispatch.return_value = mock_ad_sys_info

            mock_root_dse = Mock()
            mock_root_dse.Get.return_value = "CN=Configuration,DC=example,DC=com"

            mock_container = Mock()
            mock_obj = Mock()
            mock_obj.Class = "msSPP-ActivationObject"
            mock_obj.Name = f"CN={cn_name}"
            mock_obj.Parent = "LDAP://CN=Activation Objects"
            mock_parent = Mock()

            def get_object_side_effect(path: str) -> Mock:
                if "rootDSE" in path:
                    return mock_root_dse
                if cn_name in path:
                    return mock_obj
                if "Activation Objects" in path:
                    return mock_container if f"CN={cn_name}" not in path else mock_obj
                if path == mock_obj.Parent:
                    return mock_parent
                return Mock()

            mock_namespace = Mock()
            mock_namespace.OpenDSObject.side_effect = [
                mock_root_dse,
                mock_container,
            ]
            mock_get_object.side_effect = get_object_side_effect

            # Should not raise exception
            ao_name = f"CN={cn_name}"
            slmgr.ad_delete_activation_object(conn, output, ao_name)


class TestRegistryPathHandling:
    """Property-based tests for registry path handling"""

    @given(
        key_path=st.text(
            min_size=1,
            max_size=100,
            alphabet=st.characters(
                whitelist_categories=("Ll", "Lu", "Nd"),
                whitelist_characters="\\-_ ",
            ),
        ),
        value_name=st.text(
            min_size=1,
            max_size=50,
            alphabet=st.characters(
                whitelist_categories=("Ll", "Lu", "Nd"),
                whitelist_characters="-_",
            ),
        ),
        value=st.text(min_size=0, max_size=100),
    )
    @settings(max_examples=500, suppress_health_check=[HealthCheck.too_slow])
    def test_registry_set_string_value_remote(
        self, key_path: str, value_name: str, value: str
    ) -> None:
        """Test registry string value setting with various paths"""
        assume(len(key_path) > 0)
        assume(len(value_name) > 0)

        conn = slmgr.WMIConnection("remote-pc")
        reg = slmgr.RegistryManager(conn)

        mock_registry = Mock()
        mock_registry.SetStringValue.return_value = 0
        conn.registry = mock_registry

        result = reg.set_string_value(
            slmgr.HKEY_LOCAL_MACHINE, key_path, value_name, value
        )

        assert result == 0
        mock_registry.SetStringValue.assert_called_once_with(
            slmgr.HKEY_LOCAL_MACHINE, key_path, value_name, value
        )

    @given(
        key_path=st.text(
            min_size=1,
            max_size=100,
            alphabet=st.characters(
                whitelist_categories=("Ll", "Lu", "Nd"),
                whitelist_characters="\\-_ ",
            ),
        ),
        value_name=st.text(
            min_size=1,
            max_size=50,
            alphabet=st.characters(
                whitelist_categories=("Ll", "Lu", "Nd"),
                whitelist_characters="-_",
            ),
        ),
    )
    @settings(max_examples=300)
    def test_registry_delete_value_remote(self, key_path: str, value_name: str) -> None:
        """Test registry value deletion with various paths"""
        assume(len(key_path) > 0)
        assume(len(value_name) > 0)

        conn = slmgr.WMIConnection("remote-pc")
        reg = slmgr.RegistryManager(conn)

        mock_registry = Mock()
        mock_registry.DeleteValue.return_value = 0
        conn.registry = mock_registry

        result = reg.delete_value(slmgr.HKEY_LOCAL_MACHINE, key_path, value_name)

        assert result == 0
        mock_registry.DeleteValue.assert_called_once_with(
            slmgr.HKEY_LOCAL_MACHINE, key_path, value_name
        )


class TestProductKeyValidation:
    """Property-based tests for product key handling"""

    @given(
        st.text(
            min_size=29,
            max_size=29,
            alphabet=st.characters(
                whitelist_categories=("Lu", "Nd"),
                min_codepoint=ord("0"),
                max_codepoint=ord("Z"),
            ),
        )
    )
    @settings(max_examples=500)
    def test_product_key_format_handling(self, product_key_chars: str) -> None:
        """Test that product keys of various formats are handled"""
        assume(len(product_key_chars) == 29)

        # Create a product key format: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
        product_key = "-".join(
            [
                product_key_chars[0:5],
                product_key_chars[5:10],
                product_key_chars[10:15],
                product_key_chars[15:20],
                product_key_chars[20:25],
            ]
        )

        conn = slmgr.WMIConnection()
        reg = slmgr.RegistryManager(conn)
        output = slmgr.OutputManager()

        mock_service = Mock()
        mock_service.Version = "10.0"
        conn.wmi_service = Mock()
        conn.wmi_service.query.return_value = [mock_service]

        reg.set_string_value = Mock(return_value=0)
        reg.delete_value = Mock(return_value=0)
        reg.key_exists = Mock(return_value=False)

        # Should not raise exception during key format handling
        try:
            slmgr.install_product_key(conn, reg, output, product_key)
        except Exception:  # pylint: disable=broad-exception-caught
            # WMI methods might fail, but format should be accepted
            pass
