# 世纪互联 (21Vianet) Intune 全量导出与 PAW 排除脚本使用指南

> ### ⚠️ DISCLAIMER / 免责声明 (USE AT YOUR OWN RISK)
> **English:** This sample script is not supported under any Microsoft standard support program or service. The sample script is provided **AS IS** without warranty of any kind. Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you.
>
> **中文：** 本仓库中的示例脚本按“原样 (AS IS)”提供，不提供任何明示或暗示的担保。对于第一步只读导出脚本 (`Export-IntuneAssignments_21V.ps1`)，它仅发送 HTTP GET 请求读取配置；对于第二步批量变更脚本 (`Add-IntunePAWExclusion_21V.ps1`)，**使用该脚本修改租户配置的一切风险由使用者自行承担 (USE AT YOUR OWN RISK)**。在附加 `-ExecuteChanges` 参数对生产环境执行实际写入前，请务必：
> 1. 人工逐行审核导出的 CSV 文件中标记为 `AddPAWExclusion = Yes` 的对象；
> 2. 先在默认的 **DryRun（只读预览）** 模式下运行并核对控制台输出；
> 3. 先在测试环境或少量非核心策略上验证，并妥善保管脚本自动生成的 `PreChange_Backup_<时间戳>` JSON 快照。

本仓库包含两套针对 **世纪互联 (Azure China / 21Vianet)** 环境的 Intune PowerShell 自动化脚本，用于解决原生示例脚本仅支持全球版 (`graph.microsoft.com`) 且仅能导出单一策略类别的问题。

---

## 一、脚本清单与作用

| 脚本文件名 | 模式 | 作用说明 |
| :--- | :--- | :--- |
| **`Export-IntuneAssignments_21V.ps1`** | **纯只读 (GET)** | 导出 Intune 中全部 **10 大类**对象（App、App Config、Settings Catalog、Device Config、Admin Templates、Compliance、Endpoint Security、PowerShell Script、Proactive Remediation）的配置 JSON 备份及所有 Assignment 分配明细；自动识别哪些是给 **All Windows** 的，生成 Excel 审核表。 |
| **`Add-IntunePAWExclusion_21V.ps1`** | **写入 (POST)** | 读取人工审核后的 CSV 表（将 `AddPAWExclusion` 标为 `Yes` 的行），在保留原有所有分配的前提下，批量追加 **PAW Exclusion Group**（或挂载 PAW Exclusion Filter）。默认开启 **DryRun 预览模式**，并在修改前自动备份原始分配快照。 |

---

## 二、Service Principal (App Registration) 权限要求

如果使用 Service Principal (`TenantId` + `ClientId` + `ClientSecret`) 连接世纪互联 Microsoft Graph (`https://microsoftgraph.chinacloudapi.cn`)，请确保在 **Entra ID (portal.partner.microsoftonline.cn) -> 应用注册 -> API 权限** 中授予以下 **Application (应用程序) 权限** 并点击 **代表租户授予管理员同意**：

### 1. 只读导出阶段 (`Export-IntuneAssignments_21V.ps1`)
- `DeviceManagementApps.Read.All`（读取应用、App Config）
- `DeviceManagementConfiguration.Read.All`（读取 Settings Catalog、Device Config、Admin Templates、Compliance、Endpoint Security）
- `DeviceManagementManagedDevices.Read.All`（读取受管设备相关策略）
- **`DeviceManagementScripts.Read.All`**（⚠️ **必选**：用于读取 `deviceManagementScripts` PowerShell 脚本与 `deviceHealthScripts` Proactive Remediations，缺少此权限这两类会返回 403 Forbidden）
- `Group.Read.All`（或 `Directory.Read.All`，用于把 Assignment 里的 Group GUID 翻译为可读的组名称）

### 2. 批量添加排除阶段 (`Add-IntunePAWExclusion_21V.ps1`)
- `DeviceManagementApps.ReadWrite.All`
- `DeviceManagementConfiguration.ReadWrite.All`
- `DeviceManagementManagedDevices.ReadWrite.All`
- **`DeviceManagementScripts.ReadWrite.All`**（用于修改 PowerShell 脚本与 Proactive Remediations 的 Assignment）
- `Group.Read.All`

> 💡 **自动检测与安装 `Microsoft.Graph.Authentication`**：这两份脚本在启动时会自动检查本机是否安装了 `Microsoft.Graph.Authentication` 模块；若未安装会自动以 `CurrentUser` 范围静默安装，随后调用 `Connect-MgGraph -Environment China` 与 `Invoke-MgGraphRequest`，并在连接后自动校验当前 Token 是否具备所需的 Intune 权限。

![Entra ID API 权限配置与管理员同意](images/01_entra_app_permissions.png)

---

## 三、分步操作流程（附可视化截图说明）

### 第一步：执行全量只读导出 (`Export-IntuneAssignments_21V.ps1`)
在 PowerShell 中运行以下命令（纯只读 `GET` 操作，绝对不会修改任何租户配置）：

```powershell
cd "C:\Scripts\Intune21V"

.\Export-IntuneAssignments_21V.ps1 `
    -TenantId "<您的21V租户ID或域名，如 contoso.partner.onmschina.cn>" `
    -ClientId "<App Registration 的 Client ID>" `
    -ClientSecret "<Client Secret>"
```

脚本运行结束后，会在当前目录下生成 `Export_Output_<时间戳>` 文件夹，内含：
1. **`Intune_Assignments_AllWindows_Candidates.csv`**：自动筛选出的疑似分配给 **All Windows**（包括 Windows 策略分配给 `All Devices`、`All Users`，或分配给名字含 `Windows`/`Win10`/`Win11` 的组）的重点审核表。
2. **`Intune_Assignments_Review_All.csv`**：租户内所有对象的全量分配表。
3. **`Intune_Objects_Summary.csv`**：每个策略/应用/脚本的汇总状态表（每个对象一行，便于全局盘点哪些对象尚未配置 PAW 排除）。
4. **`Backup_JSON\`**：按类别保存的每个策略/应用/脚本的完整配置与当前 Assignment JSON 备份。

![第一步只读导出控制台输出示例](images/02_step1_export_console.png)

---

### 第二步：在 Excel 中人工审核 CSV 并标记 `AddPAWExclusion = Yes`
1. 用 Excel 打开 `Intune_Assignments_AllWindows_Candidates.csv`（或 `Intune_Assignments_Review_All.csv`）。
2. 查看 `Category`、`ItemName`、`TargetType`、`TargetGroupName` 以及 `ExistingExcludedGroups`（是否已经排除过 PAW）。
3. 对于确定需要添加 PAW 排除的 Tier 2 策略/应用/脚本，将最后一列 **`AddPAWExclusion`** 的值从 `No` 改为 **`Yes`**，保存 CSV。

> ⚠️ **特别注意（Intune 平台核心限制）**：
> - 如果某条策略/应用的分配目标是 **`All Devices`** 或 **Device Group（设备组）**，直接添加 **PAW 设备组** 作为 Exclusion Group 完全生效。
> - 如果某条策略/应用的分配目标是 **`All Users`** 或 **User Group（用户组）**，Intune **不支持跨类型排除设备组**（排除不会生效）。针对 `All Users` 分配，要么排除 **PAW 用户组**，要么使用 **Assignment Filter (`-PAWAssignmentFilterId`)** 进行设备属性排除。

![第二步 Excel 审核与勾选示例](images/03_step2_csv_excel_review.png)

---

### 第三步：批量添加 PAW Exclusion（先 DryRun 预览，再正式执行）

#### 1. 先跑 DryRun 预览（默认不加 `-ExecuteChanges`，100% 只读预览，不会修改任何配置）
```powershell
.\Add-IntunePAWExclusion_21V.ps1 `
    -CsvPath ".\Export_Output_xxxx\Intune_Assignments_AllWindows_Candidates.csv" `
    -PAWExclusionGroupId "<PAW设备组的Object_ID_GUID>" `
    -TenantId "<21V_TenantId>" `
    -ClientId "<ClientId>" `
    -ClientSecret "<ClientSecret>"
```

#### 2. 确认预览输出无误后，加上 `-ExecuteChanges` 正式执行写入
```powershell
.\Add-IntunePAWExclusion_21V.ps1 `
    -CsvPath ".\Export_Output_xxxx\Intune_Assignments_AllWindows_Candidates.csv" `
    -PAWExclusionGroupId "<PAW设备组的Object_ID_GUID>" `
    -TenantId "<21V_TenantId>" `
    -ClientId "<ClientId>" `
    -ClientSecret "<ClientSecret>" `
    -ExecuteChanges
```
正式执行前，脚本会自动在 `PreChange_Backup_<时间戳>` 目录下保存每个待修改对象的修改前 JSON 快照。

![第三步 DryRun 预览与正式写入对比示例](images/04_step3_dryrun_preview.png)

---

## 四、进阶架构：配合 21V Enrollment Time Grouping 实现 PAW 零时间窗事前隔离

在世纪互联 (21Vianet) 环境中，若依赖传统的 Entra ID 动态安全组等待设备入网后再收敛进组，往往存在 **数分钟到数十分钟的动态组评估延迟**，导致新注册的 Tier 0 (PAW) 设备在刚入网时被 `All Devices` 的 Tier 2 普通办公软件/脚本污染。

推荐将本仓库脚本与 **Windows Autopilot 设备准备工作 (Device Preparation / Enrollment Time Grouping)** 结合，按以下 **四步先后顺序** 构建零时间窗防线：

1. **第 1 步（建组与授权）**：在 21V Entra ID 中创建静态安全组（如 `SG-Intune-PAW-Devices`），并将内置服务主体 **`Intune Provisioning Client`**（App ID: `f1346770-5b25-470b-88bd-d5744ab7952c`）添加为该组的 **所有者 (Owner)**。此时组内可以没有任何设备成员（空组）。
2. **第 2 步（先筑拦截网 — 运行本仓库脚本）**：在新 PAW 设备开机注册**之前**，先运行本仓库的 `Export-IntuneAssignments_21V.ps1` 与 `Add-IntunePAWExclusion_21V.ps1`，将上述静态安全组 `SG-Intune-PAW-Devices` 提前挂到所有现有 Tier 2 策略、应用和脚本的 **Exclude Group** 上。
3. **第 3 步（配置注册时间分组 Enrollment Time Grouping）**：在 21V Intune 管理中心（`intune.partner.microsoftonline.cn`）进入 **设备 > 设备载入 > 设备准备策略**，在 PAW 专用设备准备策略的 **注册时间分组 (Enrollment time grouping)** 中选择 `SG-Intune-PAW-Devices`。
4. **第 4 步（开机即隔离闭环）**：当管理员使用 PAW 专用账号在新机器 OOBE 界面登录注册时，Intune 会在下发任何策略**之前**，当场将该设备塞入 `SG-Intune-PAW-Devices`，**0 秒命中第 2 步提前布好的 Exclude 拦截网**，彻底杜绝 Tier 2 策略污染。

![世纪互联 21V PAW 事前隔离防线与 Enrollment Time Grouping 闭环架构](images/05_paw_isolation_workflow.png)
