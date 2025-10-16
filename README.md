# slmgr-py: Windows Software Licensing Management Tool (Python)

A complete Python conversion of Microsoft's `slmgr.vbs` script for managing Windows licensing and activation. This tool provides command-line management of Windows Software Protection Platform (SPP) with full feature parity to the original VBScript version.

## Features

- **Full Feature Parity**: All 30+ commands from the original slmgr.vbs
- **License Management**: Install, uninstall, display, and activate Windows product keys
- **KMS Support**: Complete Key Management Service client and server configuration
- **Token-based Activation**: Manage and force token-based activation
- **Active Directory Activation**: Configure and manage AD-based activation
- **Remote Execution**: Manage licensing on remote computers
- **AVMA Support**: Automatic VM Activation for virtualized environments

## Requirements

- **Python**: 3.10 or higher
- **Operating System**: Windows 10/11 or Windows Server 2016+
- **Python Packages**:
  - pywin32
  - wmi
- **Permissions**: Administrator/elevated privileges required for most operations

## Installation

### From Source

1. Clone this repository:

```bash
git clone https://github.com/tboy1337/slmgr-py.git
cd slmgr-py
```

2. Install required packages:

```bash
pip install -r requirements.txt
```

## Usage

### Basic Syntax

```bash
python slmgr.py [MachineName [User Password]] [Option]
```

- `MachineName`: Name of remote machine (default is local machine)
- `User`: Account with required privilege on remote machine
- `Password`: Password for the account

### Common Commands

#### Display License Information

```bash
# Display current license
python slmgr.py /dli

# Display detailed license information
python slmgr.py /dlv

# Display all licenses
python slmgr.py /dli all
```

#### Product Key Management

```bash
# Install product key
python slmgr.py /ipk XXXXX-XXXXX-XXXXX-XXXXX-XXXXX

# Uninstall current product key
python slmgr.py /upk

# Clear product key from registry
python slmgr.py /cpky
```

#### Activation

```bash
# Activate Windows online
python slmgr.py /ato

# Get Installation ID for phone activation
python slmgr.py /dti

# Activate with Confirmation ID (phone activation)
python slmgr.py /atp <Confirmation-ID>

# Check activation expiration
python slmgr.py /xpr
```

#### KMS Client Configuration

```bash
# Set KMS server
python slmgr.py /skms kms.example.com:1688

# Clear KMS server
python slmgr.py /ckms

# Set KMS DNS lookup domain
python slmgr.py /skms-domain example.com

# Enable/disable KMS host caching
python slmgr.py /skhc  # Enable
python slmgr.py /ckhc  # Disable
```

#### KMS Server Configuration

```bash
# Set KMS listening port
python slmgr.py /sprt 1688

# Set activation/renewal intervals
python slmgr.py /sai 120  # 120 minutes
python slmgr.py /sri 10080  # 7 days

# Enable/disable DNS publishing
python slmgr.py /sdns  # Enable
python slmgr.py /cdns  # Disable

# Set KMS priority
python slmgr.py /spri  # Normal
python slmgr.py /cpri  # Low

# Set activation type
python slmgr.py /act-type 2  # 0=All, 1=AD, 2=KMS, 3=Token
```

#### Token-based Activation

```bash
# List installed issuance licenses
python slmgr.py /lil

# Remove issuance license
python slmgr.py /ril <ILID> <ILvID>

# List TKA certificates
python slmgr.py /ltc

# Force token activation
python slmgr.py /fta <Thumbprint> [PIN]
```

#### Active Directory Activation

```bash
# Activate AD forest online
python slmgr.py /ad-activation-online <Product-Key> [AO-Name]

# Get AD Installation ID
python slmgr.py /ad-activation-get-iid <Product-Key>

# Activate AD forest with CID
python slmgr.py /ad-activation-apply-cid <Product-Key> <CID> [AO-Name]

# List activation objects
python slmgr.py /ao-list

# Delete activation object
python slmgr.py /del-ao <AO-DN>
```

#### License File Management

```bash
# Install license file
python slmgr.py /ilc license.xrm-ms

# Reinstall all system licenses
python slmgr.py /rilc
```

#### Rearm Operations

```bash
# Rearm Windows (reset grace period)
python slmgr.py /rearm

# Rearm specific application
python slmgr.py /rearm-app <Application-ID>

# Rearm specific SKU
python slmgr.py /rearm-sku <Activation-ID>
```

### Remote Execution

Execute commands on remote computers:

```bash
# Local authentication
python slmgr.py COMPUTER /dli

# With credentials
python slmgr.py COMPUTER username password /dli
```

**Note**: Some commands (AD activation, token activation) do not support remote execution.

## Global Options

| Option | Parameters | Description |
|--------|-----------|-------------|
| `/ipk` | `<Product Key>` | Install product key |
| `/ato` | `[Activation ID]` | Activate Windows |
| `/dli` | `[Activation ID\|All]` | Display license information |
| `/dlv` | `[Activation ID\|All]` | Display detailed license information |
| `/xpr` | `[Activation ID]` | Display expiration date |

## Advanced Options

| Option | Parameters | Description |
|--------|-----------|-------------|
| `/cpky` | None | Clear product key from registry |
| `/ilc` | `<License file>` | Install license file |
| `/rilc` | None | Reinstall system license files |
| `/rearm` | None | Reset licensing status |
| `/rearm-app` | `<Application ID>` | Reset app licensing status |
| `/rearm-sku` | `<Activation ID>` | Reset SKU licensing status |
| `/upk` | `[Activation ID]` | Uninstall product key |
| `/dti` | `[Activation ID]` | Display Installation ID |
| `/atp` | `<Confirmation ID> [Activation ID]` | Phone activation |

## Migration from VBScript

This Python version maintains command-line compatibility with the original `slmgr.vbs`:

```bash
# VBScript version
cscript slmgr.vbs /dli

# Python version - same syntax
python slmgr.py /dli
```

### Key Differences

1. **Syntax**: Use `python slmgr.py` instead of `cscript slmgr.vbs`
2. **Performance**: Generally faster execution due to Python's performance
3. **Error Messages**: Enhanced error reporting with Python stack traces
4. **Type Safety**: Full type annotations for better code reliability
5. **Cross-platform Readability**: Python code is more maintainable

## Common Use Cases

### Scenario 1: Enterprise KMS Deployment

```bash
# Configure client to use KMS server
python slmgr.py /skms kms.corp.local:1688

# Activate against KMS
python slmgr.py /ato

# Verify activation
python slmgr.py /dli
```

### Scenario 2: Volume License Activation

```bash
# Install MAK key
python slmgr.py /ipk XXXXX-XXXXX-XXXXX-XXXXX-XXXXX

# Activate online
python slmgr.py /ato

# Check remaining activations (if MAK)
python slmgr.py /dlv
```

### Scenario 3: Offline Activation

```bash
# Get Installation ID
python slmgr.py /dti

# (Call Microsoft activation center with IID, receive Confirmation ID)

# Apply Confirmation ID
python slmgr.py /atp CONFIRMATION-ID
```

### Scenario 4: License Troubleshooting

```bash
# Display detailed license info
python slmgr.py /dlv

# Check expiration
python slmgr.py /xpr

# Reinstall licenses if corrupted
python slmgr.py /rilc

# Rearm if in grace period
python slmgr.py /rearm
```

## Security Considerations

- **Elevated Privileges**: Most operations require administrator rights
- **Product Keys**: Handle product keys securely, never log or display
- **Remote Access**: Use secure credentials for remote operations
- **Registry Access**: Be cautious with `/cpky` as it permanently removes keys
- **Network**: KMS traffic uses port 1688 (TCP)

## Troubleshooting

### Common Issues

**"Access denied" errors**
- Run PowerShell or Command Prompt as Administrator
- Ensure user has appropriate permissions

**"Cannot connect to WMI" errors**
- Check Windows Management Instrumentation service is running
- Verify firewall allows WMI traffic for remote operations

**"Product not found" errors**
- Ensure Windows is properly licensed
- Check that a product key is installed (`/dli`)

**KMS activation fails**
- Verify KMS server is reachable (`ping kms-server`)
- Check port 1688 is not blocked
- Ensure client count meets threshold (25 for client OS, 5 for server)

## Contributing

Contributions are welcome! Please:

1. Maintain Python 3.10+ compatibility
2. Follow existing code style (Black, isort)
3. Add type hints for all functions
4. Include tests for new features
5. Update documentation

## References

- [Microsoft Volume Activation](https://docs.microsoft.com/windows/deployment/volume-activation/)
- [Windows Software Protection Platform](https://docs.microsoft.com/previous-versions/windows/desktop/spp/software-protection-platform-portal)
- [KMS Activation](https://docs.microsoft.com/windows-server/get-started/kms-activation-planning)
