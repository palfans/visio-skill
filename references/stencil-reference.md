# Visio Stencil Reference

This document lists commonly available Visio stencil files and their master shapes.
**Master names are taken from live COM enumeration of the English (1033) pack; the Chinese column shows the localized name on the Simplified Chinese (2052) pack.**

## Stencil File Locations

Visio stencil files (`.vssx`) live under the Office install:

- `C:\Program Files\Microsoft Office\root\Office16\Visio Content\1033\` (English)
- `C:\Program Files\Microsoft Office\root\Office16\Visio Content\2052\` (Chinese Simplified)

`New-VisioDiagram.py` scans every `Office*` folder and every language subfolder, so a plain file name such as `NETSYM_M.VSSX` resolves on any language pack.

## Network & Infrastructure

### NETSYM_M.VSSX — Network Symbols (24 masters)

| # | Master Name | Chinese (2052) |
|---|-------------|----------------|
| 1 | Router | 路由器 |
| 2 | ATM router | ATM 路由器 |
| 3 | ISDN switch | ISDN 交换机 |
| 4 | ATM switch | ATM 交换机 |
| 5 | ATM/FastGB etherswitch | ATM/FastGB 以太网交换机 |
| 6 | Workgroup switch | 工作组交换机 |
| 7 | Small hub | 小型集线器 |
| 8 | 100BaseT hub | 100BaseT 集线器 |
| 9 | CDDI/FDDI concentrator | CDDI/FDDI 集中器 |
| 10 | Terminal server | 终端服务器 |
| 11 | Comm server | 通信服务器 |
| 12 | Probe | 探测器 |
| 13 | Bridge | 网桥 |
| 14 | Gateway | 网关 |
| 15 | WAN | WAN |
| 16 | DSU/CSU | DSU/CSU |
| 17 | Host | 主机 |
| 18 | Relational database | 关系数据库 |
| 19 | Lock | 锁 |
| 20 | Lock and key | 锁和密钥 |
| 21 | Key | 密钥 |
| 22 | Public/private key server | 公钥/私钥服务器 |
| 23 | Certificate server | 证书服务器 |
| 24 | Dynamic connector | 动态连接线 |

### SERVER_M.VSSX — Servers (17 masters)

| # | Master Name | Chinese (2052) |
|---|-------------|----------------|
| 1 | Server | 服务器 |
| 2 | File server | 文件服务器 |
| 3 | Email server | 电子邮件服务器 |
| 4 | Web server | Web 服务器 |
| 5 | Proxy server | 代理服务器 |
| 6 | Real-time communications server | 实时通信服务器 |
| 7 | E-Commerce server | 电子商务服务器 |
| 8 | Database server | 数据库服务器 |
| 9 | Content management server | 内容管理服务器 |
| 10 | FTP server | FTP 服务器 |
| 11 | Streaming media server | 流化媒体服务器 |
| 12 | Management server | 管理服务器 |
| 13 | Directory server | 目录服务器 |
| 14 | Print server | 打印服务器 |
| 15 | Mobile information server | 移动信息服务器 |
| 16 | Application server | 应用程序服务器 |
| 17 | Dynamic connector | 动态连接线 |

### COMPS_M.VSSX — Computers and Monitors (10 masters)

| # | Master Name | Chinese (2052) |
|---|-------------|----------------|
| 1 | PC | PC |
| 2 | Laptop computer | 笔记本电脑 |
| 3 | LCD monitor | LCD 显示器 |
| 4 | Terminal | 终端 |
| 5 | Tablet computer | 平板电脑 |
| 6 | PDA | PDA |
| 7 | iMac | iMac |
| 8 | New iMac | 新式 iMac |
| 9 | CRT monitor | CRT 监视器 |
| 10 | Dynamic connector | 动态连接线 |

## Cloud Services (未逐一验证，使用前建议用发现脚本确认)

- `AZURECLOUD_M.VSSX` — Azure Cloud Services
- `AZURECOMPUTE_M.VSSX` — Azure Compute
- `AZURENETWORKING_M.VSSX` — Azure Networking
- `AZUREDATABASES_M.VSSX` — Azure Databases
- `AZURESTORAGE_M.VSSX` — Azure Storage
- `AWSCOMPUTE_M.VSSX` — AWS Compute
- `AWSSTORAGE_M.VSSX` — AWS Storage
- `AWSNETCONTENTDELIVERY_M.VSSX` — AWS Networking
- `KUBERNETESVISIOSTENCIL_M.VSSX` — Kubernetes

## How to Discover Master Names

Use Python to list all masters in any stencil file:

```python
import win32com.client

visio = win32com.client.gencache.EnsureDispatch("Visio.Application")
visio.Visible = 0
s = visio.Documents.OpenEx(r"C:\Path\To\Stencil.vssx", 64)
for i in range(1, s.Masters.Count + 1):
    print(f"{i}: {s.Masters.Item(i).Name} / {s.Masters.Item(i).NameU}")
s.Close()
visio.Quit()
```

## Important Notes

1. **Use the universal English name in `master`** — it resolves on every language pack. The script compares `master` against both `Name` (localized) and `NameU` (universal), ignoring case, so `路由器` still works on a Chinese install.
2. **Spaces matter** — `Web server` is not `Webserver`.
3. **`Dynamic connector` exists in every stencil** — ignore it, use `ConnectorToolDataObject` for connections instead.
4. **No firewall master** — NETSYM_M has no firewall shape; use `Gateway` as the closest alternative.
5. **No printer in COMPS_M** — use `Print server` from SERVER_M instead.
6. **Mixed usage works** — stencil masters and basic shapes (`rectangle`, `roundrect`, `ellipse`, `diamond`) can coexist in one diagram.
