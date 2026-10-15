# N-sight MCP server (read-only)

Gives Claude desktop read-only access to N-able N-sight RMM device data for QBR hardware tables.

## Setup
1. In N-sight, generate an API key (Settings > General Settings > API). Note your region's server.
2. In this folder run:
   ```
   powershell -ExecutionPolicy Bypass -File .\setup.ps1
   ```
   It reuses the Halo connector's Python environment, stores the API key in Windows Credential Manager,
   tests the connection and prints the Claude config entry.
3. Claude desktop > Settings > Developer > Edit Config, add the `nsight` entry next to `halo`,
   then fully quit and reopen Claude.

## Tools
| Tool | Use |
|---|---|
| `test_connection` | Check the key and server |
| `list_clients`, `list_sites` | Find a clientid |
| `list_workstations`, `list_servers` | Devices per client |
| `hardware_inventory` | QBR rows: PC name, last user, model, processor, RAM, OS, custom fields. Pages 25 devices per call |
| `device_asset_details` | Every raw asset field for one device |

## Notes
- Nothing writes to N-sight.
- The API allows 60 calls a minute; the server spaces calls about 1 second apart and caches results for the session.
- To change the key: `..\.venv\Scripts\python.exe -m keyring set nsight-mcp nsight_api_key`
