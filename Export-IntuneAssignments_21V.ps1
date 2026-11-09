<#
.SYNOPSIS
    导出世纪互联 (21Vianet / Azure China) Intune 中所有 App、Script、Policy、Remediation 的配置与 Assignment 明细。
    专门用于排查哪些策略/应用/脚本分配给了 "All Windows" (All Devices / All Users / Windows 动态组)，并生成可供批量添加 PAW Exclusion 的审核表。

.DESCRIPTION
    本脚本为 **纯只读 (Read-Only)** 脚本，仅发送 HTTP GET 请求，不会修改任何 Intune 或 Entra ID 配置。
    兼容世纪互联 (21Vianet) 环境：
      - 登录端点: https://login.partner.microsoftonline.cn
      - Graph 端点: https://microsoftgraph.chinacloudapi.cn

    覆盖的 Intune 资源类别：
      1. Apps (mobileApps)
      2. App Configuration Policies - Managed Devices (mobileAppConfigurations)
      3. App Configuration Policies - Managed Apps (targetedManagedAppConfigurations)
      4. Settings Catalog Policies (configurationPolicies)
      5. Device Configuration Profiles (deviceConfigurations)
      6. Administrative Templates / GPO (groupPolicyConfigurations)
      7. Device Compliance Policies (deviceCompliancePolicies)
      8. Endpoint Security Policies (intents)
      9. PowerShell Scripts (deviceManagementScripts)
     10. Proactive Remediations (deviceHealthScripts)

.PARAMETER TenantId
    世纪互联租户 ID (GUID) 或租户域名 (例如: contoso.partner.onmschina.cn)

.PARAMETER ClientId
    Service Principal (App Registration) 的 Application (client) ID

.PARAMETER ClientSecret
    Service Principal 的客户端密钥 (明文字符串或从安全提示输入)

.PARAMETER UseInteractive
    如果未提供 Service Principal 密钥，可使用交互式浏览器登录 (需安装 Microsoft.Graph.Authentication 模块)

.PARAMETER OutputRoot
    导出结果的根目录，默认为当前脚本所在目录下的 Export_Output_<时间戳>

.NOTES
    DISCLAIMER / 免责声明 (USE AT YOUR OWN RISK):
    This sample script is not supported under any Microsoft standard support program or service.
    The sample script is provided AS IS without warranty of any kind. Microsoft further disclaims
    all implied warranties including, without limitation, any implied warranties of merchantability
    or of fitness for a particular purpose. The entire risk arising out of the use or performance
    of the sample scripts and documentation remains with you.
    本示例脚本按“原样 (AS IS)”提供，不提供任何明示或暗示的担保。使用本脚本的风险由使用者自行承担。
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,

    [Parameter(Mandatory = $false)]
    [switch]$UseInteractive,

    [Parameter(Mandatory = $false)]
    [Alias("OutputFolder", "OutputPath")]
    [string]$OutputRoot,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeSettingsCatalogSettings,

    [Parameter(Mandatory = $false)]
    [switch]$SkipSettingsCatalogDetail
)

$ErrorActionPreference = "Stop"

# ============================================================================
# 1. 世纪互联 (21Vianet) 端点常量定义
# ============================================================================
$GraphBaseUri = "https://microsoftgraph.chinacloudapi.cn"
$AuthorityBase = "https://login.partner.microsoftonline.cn"

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $OutputRoot = Join-Path $ScriptDir "Export_Output_$Timestamp"
}
$BackupJsonRoot = Join-Path $OutputRoot "Backup_JSON"
New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
New-Item -ItemType Directory -Path $BackupJsonRoot -Force | Out-Null

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Intune 21Vianet (Azure China) 全量配置与 Assignment 导出工具" -ForegroundColor Cyan
Write-Host "  模式: 只读 (Read-Only GET) | 输出目录: $OutputRoot" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan

# ============================================================================
# 2. 模块检查与身份认证：自动安装/导入 Microsoft.Graph.Authentication 并连接 21V
# ============================================================================
function Ensure-MgGraphAuthModule {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-Host "[Module] 未检测到 Microsoft.Graph.Authentication 模块，正在为您自动安装 (Scope: CurrentUser)..." -ForegroundColor Yellow
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction SilentlyContinue | Out-Null
        } catch {}
        Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Host "[Module] Microsoft.Graph.Authentication 安装完成！" -ForegroundColor Green
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
}

function Initialize-21VGraphAuth {
    Ensure-MgGraphAuthModule

    $requiredIntunePerms = @(
        "DeviceManagementApps.Read.All",
        "DeviceManagementConfiguration.Read.All",
        "DeviceManagementManagedDevices.Read.All",
        "DeviceManagementScripts.Read.All",
        "Group.Read.All"
    )

    if (-not $UseInteractive -and (-not [string]::IsNullOrWhiteSpace($ClientId))) {
        if ([string]::IsNullOrWhiteSpace($TenantId)) {
            $TenantId = Read-Host "请输入 21V TenantId (例如 xxx.partner.onmschina.cn 或 GUID)"
        }
        if ([string]::IsNullOrWhiteSpace($ClientSecret)) {
            $secSecret = Read-Host "请输入 Service Principal ClientSecret" -AsSecureString
        } else {
            $secSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
        }

        $cred = New-Object System.Management.Automation.PSCredential($ClientId, $secSecret)
        Write-Host "[Auth] 正在通过 Connect-MgGraph -Environment China 使用 Service Principal ($ClientId) 登录租户 ($TenantId)..." -ForegroundColor Yellow
        Connect-MgGraph -Environment China -TenantId $TenantId -ClientSecretCredential $cred -NoWelcome -ErrorAction Stop
    }
    else {
        $ctx = Get-MgContext
        if (-not $ctx -or $ctx.Environment -ne "China" -or $ctx.AuthType -ne "Delegated") {
            Write-Host "[Auth] 正在调用 Connect-MgGraph -Environment China 进行交互式登录..." -ForegroundColor Yellow
            if ($TenantId) {
                Connect-MgGraph -Environment China -TenantId $TenantId -Scopes $requiredIntunePerms -NoWelcome -ErrorAction Stop
            } else {
                Connect-MgGraph -Environment China -Scopes $requiredIntunePerms -NoWelcome -ErrorAction Stop
            }
        }
    }

    $ctx = Get-MgContext
    Write-Host "[Auth] MgGraph (China) 连接成功！AuthType: $($ctx.AuthType) | TenantId: $($ctx.TenantId) | Account/AppId: $($ctx.Account)$($ctx.ClientId)" -ForegroundColor Green

    # 检查当前 Token 中包含的权限 (Scopes / App Roles)
    $currentScopes = @($ctx.Scopes)
    $missingPerms = [System.Collections.Generic.List[string]]::new()
    foreach ($reqPerm in $requiredIntunePerms) {
        $basePrefix = $reqPerm -replace '\.Read\.All$', ''
        $matched = $currentScopes | Where-Object { $_ -eq $reqPerm -or $_ -eq "$basePrefix.ReadWrite.All" -or ($_ -eq "Directory.Read.All" -and $reqPerm -eq "Group.Read.All") -or ($_ -eq "Directory.ReadWrite.All" -and $reqPerm -eq "Group.Read.All") }
        if (-not $matched) {
            $missingPerms.Add($reqPerm)
        }
    }

    if ($missingPerms.Count -gt 0) {
        Write-Warning "================================================================"
        Write-Warning "检测到当前连接的账号/Service Principal 缺少以下 Intune Graph 权限："
        Write-Warning "  缺失权限: $($missingPerms -join ', ')"
        Write-Warning "  当前已拥有的权限: $(if ($currentScopes.Count -gt 0) { $currentScopes -join ', ' } else { '<无 Application Roles，可能未授予管理员同意>' })"
        Write-Warning "  提示: 若后续读取报 403 Forbidden，请在 Entra ID -> App Registrations 中为该 SP 添加上述 Application 权限并点击 [Grant admin consent]，或改用 -UseInteractive 以 Intune 管理员账号交互式登录。"
        Write-Warning "================================================================"
    }
}

# ============================================================================
# 3. 通用只读 Graph GET 封装（基于 Invoke-MgGraphRequest，自动处理 21V 域名替换、分页与 429 退避）
# ============================================================================
function Invoke-21VGraphGet {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RelativeOrFullUri,

        [switch]$SingleObject
    )

    # 强制修正任何硬编码的全球版域名到 21V 域名
    $uri = $RelativeOrFullUri
    if ($uri -match "^https://graph\.microsoft\.com") {
        $uri = $uri -replace "^https://graph\.microsoft\.com", $GraphBaseUri
    }
    elseif ($uri -notmatch "^https://") {
        if (-not $uri.StartsWith("/")) { $uri = "/$uri" }
        $uri = "$GraphBaseUri$uri"
    }

    $allItems = [System.Collections.Generic.List[object]]::new()
    $nextUri = $uri

    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        # 再次确保分页返回的 nextLink 指向 21V
        $nextUri = $nextUri -replace "^https://graph\.microsoft\.com", $GraphBaseUri

        $maxRetries = 3
        $attempt = 0
        $response = $null

        while ($true) {
            $attempt++
            try {
                $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject -ErrorAction Stop
                break
            }
            catch {
                $statusCode = $null
                if ($_.Exception.Response) {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                }
                if (($statusCode -eq 429 -or $statusCode -eq 503 -or $statusCode -eq 504) -and $attempt -le $maxRetries) {
                    $waitSec = 5 * $attempt
                    Write-Warning "请求被限流或暂时不可用 (HTTP $statusCode)，等待 $waitSec 秒后重试 ($attempt/$maxRetries)..."
                    Start-Sleep -Seconds $waitSec
                }
                else {
                    throw $_
                }
            }
        }

        if ($SingleObject) {
            return $response
        }

        if ($null -ne $response) {
            $hasValueProp = $false
            if ($response -is [System.Collections.IDictionary]) {
                $hasValueProp = $response.Contains("value")
            } else {
                $hasValueProp = ($null -ne $response.PSObject.Properties["value"])
            }

            if ($hasValueProp) {
                if ($null -ne $response.value) {
                    foreach ($item in $response.value) {
                        $allItems.Add($item)
                    }
                }
            }
            else {
                $allItems.Add($response)
            }

            $nextLinkProp = $null
            if ($response -is [System.Collections.IDictionary]) {
                $nextLinkProp = $response["@odata.nextLink"]
            } elseif ($response.PSObject.Properties["@odata.nextLink"]) {
                $nextLinkProp = $response."@odata.nextLink"
            }
            $nextUri = $nextLinkProp
        }
        else {
            $nextUri = $null
        }
    }

    return ,$allItems.ToArray()
}

# ============================================================================
# 4. 辅助函数：文件名清理、组名缓存解析、Assignment Filter 缓存、Windows 平台判定
# ============================================================================
function Remove-InvalidFileNameChars {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return "Unnamed" }
    $clean = $Name
    foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) {
        $clean = $clean.Replace($ch, [char]'_')
    }
    $clean = ($clean -replace '[\[\]\r\n\t]', '_').Trim()
    # 防止末尾是点或空格导致 Windows 文件系统异常
    $clean = $clean.TrimEnd('.', ' ')
    if ([string]::IsNullOrWhiteSpace($clean)) { $clean = "Unnamed" }
    if ($clean.Length -gt 80) { $clean = $clean.Substring(0, 80).TrimEnd('.', ' ') }
    return $clean
}

$Script:GroupCache = @{}
function Resolve-GroupName {
    param([string]$GroupId)
    if ([string]::IsNullOrWhiteSpace($GroupId)) { return "" }
    if ($Script:GroupCache.ContainsKey($GroupId)) {
        return $Script:GroupCache[$GroupId]
    }
    try {
        $grp = Invoke-21VGraphGet -RelativeOrFullUri "/v1.0/groups/$($GroupId)?`$select=id,displayName,groupTypes,membershipRule,mailEnabled,securityEnabled" -SingleObject
        $name = if ($grp.displayName) { $grp.displayName } else { "<Unnamed Group: $GroupId>" }
        $Script:GroupCache[$GroupId] = $name
        return $name
    }
    catch {
        $Script:GroupCache[$GroupId] = "<Unresolved/Deleted Group: $GroupId>"
        return $Script:GroupCache[$GroupId]
    }
}

$Script:FilterCache = @{}
function Initialize-AssignmentFilters {
    Write-Host "[Init] 正在加载租户内的 Assignment Filters..." -ForegroundColor Gray
    try {
        $filters = Invoke-21VGraphGet -RelativeOrFullUri "/beta/deviceManagement/assignmentFilters"
        foreach ($f in $filters) {
            if ($f.id) {
                $Script:FilterCache[$f.id] = [PSCustomObject]@{
                    DisplayName = $f.displayName
                    Platform    = $f.platform
                    Rule        = $f.rule
                }
            }
        }
        Write-Host "[Init] 已加载 $($Script:FilterCache.Count) 个 Assignment Filters。" -ForegroundColor Gray
    }
    catch {
        Write-Warning "[Init] 无法读取 Assignment Filters (可能无权限或未启用): $($_.Exception.Message)"
    }
}

function Get-ItemPlatformInfo {
    param(
        [string]$Category,
        [object]$Item
    )

    $odataType = ""
    if ($Item.PSObject.Properties["@odata.type"]) {
        $odataType = [string]$Item."@odata.type"
    }
    $platforms = ""
    if ($Item.PSObject.Properties["platforms"] -and $null -ne $Item.platforms) {
        $platforms = ($Item.platforms -join ",")
    }
    elseif ($Item.PSObject.Properties["platformType"] -and $null -ne $Item.platformType) {
        $platforms = [string]$Item.platformType
    }

    # 根据类别和属性推断是否适用于 Windows（注意避开 PowerShell 7 只读自动变量 $IsWindows）
    $isWinPlatform = $false
    $platformLabel = "Unknown"

    if ($Category -in @("PowerShellScript", "ProactiveRemediation", "AdminTemplate")) {
        $isWinPlatform = $true
        $platformLabel = "Windows"
    }
    elseif (-not [string]::IsNullOrWhiteSpace($platforms)) {
        $platformLabel = $platforms
        if ($platforms -match "windows") {
            $isWinPlatform = $true
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($odataType)) {
        $platformLabel = $odataType -replace "^#microsoft\.graph\.", ""
        if ($odataType -match "windows|win32|winGet|officeSuiteApp|microsoftStoreForBusiness|edge|msi") {
            $isWinPlatform = $true
        }
        elseif ($odataType -match "ios|macOS|android") {
            $isWinPlatform = $false
        }
        elseif ($Category -in @("App", "AppConfig_ManagedDevice", "AppConfig_ManagedApp", "EndpointSecurity")) {
            # WebApp 或通用策略可能跨平台，默认标记为需人工确认 (True)
            $isWinPlatform = $true
        }
    }
    else {
        $isWinPlatform = $true
    }

    return [PSCustomObject]@{
        ODataType           = $odataType
        Platform            = $platformLabel
        IsWindowsApplicable = $isWinPlatform
    }
}

# ============================================================================
# 5. 解析单个资源的 Assignments 并生成报告行
# ============================================================================
function ConvertTo-AssignmentReviewRows {
    param(
        [string]$Category,
        [object]$Item,
        [array]$Assignments
    )

    $itemId = [string]$Item.id
    $itemName = if ($Item.PSObject.Properties["displayName"] -and -not [string]::IsNullOrWhiteSpace($Item.displayName)) {
        [string]$Item.displayName
    } elseif ($Item.PSObject.Properties["name"] -and -not [string]::IsNullOrWhiteSpace($Item.name)) {
        [string]$Item.name
    } else {
        "<Unnamed: $itemId>"
    }

    $platInfo = Get-ItemPlatformInfo -Category $Category -Item $Item
    $rows = [System.Collections.Generic.List[object]]::new()

    if ($null -eq $Assignments -or $Assignments.Count -eq 0) {
        $rows.Add([PSCustomObject]@{
            Category                  = $Category
            ItemName                  = $itemName
            ItemId                    = $itemId
            Platform                  = $platInfo.Platform
            ODataType                 = $platInfo.ODataType
            IsWindowsApplicable       = $platInfo.IsWindowsApplicable
            IsAssigned                = $false
            AssignmentId              = ""
            AssignmentDirection       = "Unassigned"
            Intent                    = ""
            TargetType                = "None"
            TargetGroupId             = ""
            TargetGroupName           = ""
            FilterId                  = ""
            FilterName                = ""
            FilterMode                = ""
            ExistingExcludedGroupIds  = ""
            ExistingExcludedGroups    = ""
            IsAllWindowsCandidate     = $false
            CandidateReason           = "Not Assigned"
            AddPAWExclusion           = "No"
        })
        return ,$rows.ToArray()
    }

    # 先汇总该策略/应用下已经存在的所有 Exclusion Groups，方便审核时一眼看到是否已排除 PAW
    $existingExclIds = [System.Collections.Generic.List[string]]::new()
    $existingExclNames = [System.Collections.Generic.List[string]]::new()

    foreach ($a in $Assignments) {
        $t = $a.target
        if ($null -ne $t) {
            $tType = [string]$t."@odata.type"
            if ($tType -eq "#microsoft.graph.exclusionGroupAssignmentTarget" -and $t.groupId) {
                $gId = [string]$t.groupId
                $gName = Resolve-GroupName -GroupId $gId
                $existingExclIds.Add($gId)
                $existingExclNames.Add("$gName ($gId)")
            }
        }
    }

    $existingExclIdsStr = ($existingExclIds -join "; ")
    $existingExclNamesStr = ($existingExclNames -join "; ")

    foreach ($a in $Assignments) {
        $assignmentId = [string]$a.id
        $intent = if ($a.PSObject.Properties["intent"] -and $a.intent) { [string]$a.intent } else { "apply" }
        $t = $a.target
        $tODataType = if ($null -ne $t -and $t.PSObject.Properties["@odata.type"]) { [string]$t."@odata.type" } else { "" }

        $direction = "Include"
        $targetTypeFriendly = $tODataType -replace "^#microsoft\.graph\.", ""
        $targetGroupId = ""
        $targetGroupName = ""

        switch ($tODataType) {
            "#microsoft.graph.allDevicesAssignmentTarget" {
                $targetTypeFriendly = "All Devices"
            }
            "#microsoft.graph.allLicensedUsersAssignmentTarget" {
                $targetTypeFriendly = "All Users"
            }
            "#microsoft.graph.groupAssignmentTarget" {
                $targetTypeFriendly = "Include Group"
                $targetGroupId = [string]$t.groupId
                $targetGroupName = Resolve-GroupName -GroupId $targetGroupId
            }
            "#microsoft.graph.exclusionGroupAssignmentTarget" {
                $direction = "Exclude"
                $targetTypeFriendly = "Exclude Group"
                $targetGroupId = [string]$t.groupId
                $targetGroupName = Resolve-GroupName -GroupId $targetGroupId
            }
            default {
                if ($t -and $t.PSObject.Properties["groupId"] -and $t.groupId) {
                    $targetGroupId = [string]$t.groupId
                    $targetGroupName = Resolve-GroupName -GroupId $targetGroupId
                }
            }
        }

        $filterId = ""
        $filterMode = ""
        $filterName = ""
        if ($null -ne $t) {
            if ($t.PSObject.Properties["deviceAndAppManagementAssignmentFilterId"] -and $t.deviceAndAppManagementAssignmentFilterId) {
                $filterId = [string]$t.deviceAndAppManagementAssignmentFilterId
                if ($Script:FilterCache.ContainsKey($filterId)) {
                    $filterName = $Script:FilterCache[$filterId].DisplayName
                }
            }
            if ($t.PSObject.Properties["deviceAndAppManagementAssignmentFilterType"] -and $t.deviceAndAppManagementAssignmentFilterType) {
                $filterMode = [string]$t.deviceAndAppManagementAssignmentFilterType
            }
        }

        # 判断是否为 "All Windows" 候选对象
        $isCandidate = $false
        $reasonList = [System.Collections.Generic.List[string]]::new()

        if ($direction -eq "Include" -and $platInfo.IsWindowsApplicable) {
            if ($targetTypeFriendly -eq "All Devices") {
                $isCandidate = $true
                $reasonList.Add("Windows-applicable item assigned to [All Devices]")
            }
            elseif ($targetTypeFriendly -eq "All Users") {
                $isCandidate = $true
                $reasonList.Add("Windows-applicable item assigned to [All Users] (Note: requires User Group exclusion or Assignment Filter for PAW)")
            }
            elseif ($targetTypeFriendly -eq "Include Group" -and $targetGroupName -match "(?i)(all.*win|windows|win10|win11|workstation|client|pc)") {
                $isCandidate = $true
                $reasonList.Add("Assigned to Windows-related Group [$targetGroupName]")
            }
        }

        $rows.Add([PSCustomObject]@{
            Category                  = $Category
            ItemName                  = $itemName
            ItemId                    = $itemId
            Platform                  = $platInfo.Platform
            ODataType                 = $platInfo.ODataType
            IsWindowsApplicable       = $platInfo.IsWindowsApplicable
            IsAssigned                = $true
            AssignmentId              = $assignmentId
            AssignmentDirection       = $direction
            Intent                    = $intent
            TargetType                = $targetTypeFriendly
            TargetGroupId             = $targetGroupId
            TargetGroupName           = $targetGroupName
            FilterId                  = $filterId
            FilterName                = $filterName
            FilterMode                = $filterMode
            ExistingExcludedGroupIds  = $existingExclIdsStr
            ExistingExcludedGroups    = $existingExclNamesStr
            IsAllWindowsCandidate     = $isCandidate
            CandidateReason           = ($reasonList -join " | ")
            AddPAWExclusion           = "No"
        })
    }

    return ,$rows.ToArray()
}

# ============================================================================
# 6. 定义要导出的 10 大 Intune 类别及各自的查询方式
# ============================================================================
$Workloads = @(
    @{
        Category       = "App"
        Description    = "Intune 应用程序 (mobileApps)"
        ListUri        = "/beta/deviceAppManagement/mobileApps"
        AssignmentsUri = "/beta/deviceAppManagement/mobileApps/{0}/assignments"
        FilterClientApps = $true
    },
    @{
        Category       = "AppConfig_ManagedDevice"
        Description    = "应用配置策略 - 托管设备 (mobileAppConfigurations)"
        ListUri        = "/beta/deviceAppManagement/mobileAppConfigurations"
        AssignmentsUri = "/beta/deviceAppManagement/mobileAppConfigurations/{0}/assignments"
    },
    @{
        Category       = "AppConfig_ManagedApp"
        Description    = "应用配置策略 - 托管应用 (targetedManagedAppConfigurations)"
        ListUri        = "/beta/deviceAppManagement/targetedManagedAppConfigurations"
        AssignmentsUri = "/beta/deviceAppManagement/targetedManagedAppConfigurations/{0}/assignments"
    },
    @{
        Category       = "SettingsCatalog"
        Description    = "设置目录策略 (configurationPolicies)"
        ListUri        = "/beta/deviceManagement/configurationPolicies"
        AssignmentsUri = "/beta/deviceManagement/configurationPolicies/{0}/assignments"
        SettingsUri    = "/beta/deviceManagement/configurationPolicies/{0}/settings"
    },
    @{
        Category       = "DeviceConfig"
        Description    = "设备配置模板与自定义策略 (deviceConfigurations)"
        ListUri        = "/beta/deviceManagement/deviceConfigurations"
        AssignmentsUri = "/beta/deviceManagement/deviceConfigurations/{0}/assignments"
    },
    @{
        Category       = "AdminTemplate"
        Description    = "管理模板策略 / GPO (groupPolicyConfigurations)"
        ListUri        = "/beta/deviceManagement/groupPolicyConfigurations"
        AssignmentsUri = "/beta/deviceManagement/groupPolicyConfigurations/{0}/assignments"
    },
    @{
        Category       = "CompliancePolicy"
        Description    = "设备合规性策略 (deviceCompliancePolicies)"
        ListUri        = "/beta/deviceManagement/deviceCompliancePolicies"
        AssignmentsUri = "/beta/deviceManagement/deviceCompliancePolicies/{0}/assignments"
    },
    @{
        Category       = "EndpointSecurity"
        Description    = "终端安全策略 (intents - Antivirus/BitLocker/Firewall/ASR)"
        ListUri        = "/beta/deviceManagement/intents"
        AssignmentsUri = "/beta/deviceManagement/intents/{0}/assignments"
    },
    @{
        Category       = "PowerShellScript"
        Description    = "Windows PowerShell 脚本 (deviceManagementScripts)"
        ListUri        = "/beta/deviceManagement/deviceManagementScripts"
        DetailUri      = "/beta/deviceManagement/deviceManagementScripts/{0}"
        AssignmentsUri = "/beta/deviceManagement/deviceManagementScripts/{0}/assignments"
    },
    @{
        Category       = "ProactiveRemediation"
        Description    = "主动修复脚本 / Remediations (deviceHealthScripts)"
        ListUri        = "/beta/deviceManagement/deviceHealthScripts"
        DetailUri      = "/beta/deviceManagement/deviceHealthScripts/{0}"
        AssignmentsUri = "/beta/deviceManagement/deviceHealthScripts/{0}/assignments"
    }
)

# ============================================================================
# 7. 主执行逻辑
# ============================================================================
Initialize-21VGraphAuth
Initialize-AssignmentFilters

$AllReviewRows = [System.Collections.Generic.List[object]]::new()
$CategorySummary = [System.Collections.Generic.List[object]]::new()

foreach ($w in $Workloads) {
    $cat = $w.Category
    $desc = $w.Description
    Write-Host "`n----------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "[Workload] 正在读取: $desc ($cat)..." -ForegroundColor Cyan

    $catBackupDir = Join-Path $BackupJsonRoot $cat
    New-Item -ItemType Directory -Path $catBackupDir -Force | Out-Null

    $items = @()
    $usedExpandAssignments = $false
    try {
        # 优先尝试 $expand=assignments 批量拉取分配（可将大租户导出速度提升 20 倍以上）
        try {
            $expandUri = "$($w.ListUri)?`$expand=assignments"
            $items = Invoke-21VGraphGet -RelativeOrFullUri $expandUri
            $usedExpandAssignments = $true
        }
        catch {
            # 若当前端点不支持 $expand=assignments，自动无缝回退到普通列表拉取
            $items = Invoke-21VGraphGet -RelativeOrFullUri $w.ListUri
            $usedExpandAssignments = $false
        }
    }
    catch {
        Write-Warning "  -> 跳过 [$cat]: 读取列表失败 ($($_.Exception.Message))"
        $CategorySummary.Add([PSCustomObject]@{
            Category      = $cat
            Description   = $desc
            TotalItems    = 0
            AssignedItems = 0
            Status        = "Skipped/Error: $($_.Exception.Message)"
        })
        continue
    }

    # 对于 mobileApps，过滤掉内置/系统默认且不支持分配的 microsoftStoreForBusinessApp 系统占位项（可选保留所有用户创建应用）
    if ($w.FilterClientApps -and $items.Count -gt 0) {
        $items = @($items | Where-Object {
            $_.PSObject.Properties["isFeatured"] -ne $null -or $_.PSObject.Properties["displayName"] -ne $null
        })
    }

    Write-Host "  -> 共找到 $($items.Count) 个 [$cat] 对象$(if ($usedExpandAssignments) { ' (已通过 $expand=assignments 批量附带分配)' })。" -ForegroundColor Green
    $assignedCount = 0

    foreach ($item in $items) {
        # 每次循环重置变量，防止前一次循环残留脏数据
        $assignments = @()
        $detailObj = $item
        $settingsArr = $null

        $itemId = [string]$item.id
        $itemName = if ($item.PSObject.Properties["displayName"] -and $item.displayName) {
            [string]$item.displayName
        } elseif ($item.PSObject.Properties["name"] -and $item.name) {
            [string]$item.name
        } else {
            $itemId
        }

        # 读取详细对象 (例如 PowerShell 脚本内容)
        if ($w.ContainsKey("DetailUri")) {
            try {
                $detailObj = Invoke-21VGraphGet -RelativeOrFullUri ($w.DetailUri -f $itemId) -SingleObject
            }
            catch {
                $detailObj = $item
            }
        }

        # 读取 Settings Catalog 策略内部详细设置项（仅当显式指定 -IncludeSettingsCatalogSettings 时拉取，避免大租户产生数百次额外请求）
        if ($w.ContainsKey("SettingsUri") -and $IncludeSettingsCatalogSettings -and (-not $SkipSettingsCatalogDetail)) {
            try {
                $settingsArr = Invoke-21VGraphGet -RelativeOrFullUri ($w.SettingsUri -f $itemId)
            }
            catch {
                $settingsArr = @()
            }
        }

        # 读取 Assignments（如果 $expand=assignments 已经带回了 assignments，则直接使用，无需额外发起 N 次 HTTP GET）
        if ($usedExpandAssignments -and $null -ne $item.PSObject.Properties["assignments"] -and $null -ne $item.assignments) {
            $assignments = @($item.assignments)
        }
        else {
            try {
                $assignments = Invoke-21VGraphGet -RelativeOrFullUri ($w.AssignmentsUri -f $itemId)
            }
            catch {
                Write-Warning "  -> [$cat] '$itemName' ($itemId) 读取 Assignments 失败: $($_.Exception.Message)"
                $assignments = @()
            }
        }

        if ($assignments.Count -gt 0) {
            $assignedCount++
        }

        # 保存完整 JSON 备份 (含配置本体 + Assignments)，可用于审计和随时回滚
        $backupPayload = [ordered]@{
            ExportedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
            Category      = $cat
            ItemId        = $itemId
            ItemName      = $itemName
            Configuration = $detailObj
            Settings      = $settingsArr
            Assignments   = $assignments
        }
        $safeFileName = "{0}_{1}.json" -f (Remove-InvalidFileNameChars -Name $itemName), $itemId
        $backupFilePath = Join-Path $catBackupDir $safeFileName
        [System.IO.File]::WriteAllText($backupFilePath, ($backupPayload | ConvertTo-Json -Depth 20), [System.Text.Encoding]::UTF8)

        # 转换为 CSV 报告行
        $itemRows = ConvertTo-AssignmentReviewRows -Category $cat -Item $detailObj -Assignments $assignments
        foreach ($r in $itemRows) {
            $AllReviewRows.Add($r)
        }
    }

    $CategorySummary.Add([PSCustomObject]@{
        Category      = $cat
        Description   = $desc
        TotalItems    = $items.Count
        AssignedItems = $assignedCount
        Status        = "Success"
    })
}

# ============================================================================
# 8. 导出 CSV 报表与汇总统计
# ============================================================================
$FullCsvPath = Join-Path $OutputRoot "Intune_Assignments_Review_All.csv"
$CandidatesCsvPath = Join-Path $OutputRoot "Intune_Assignments_AllWindows_Candidates.csv"
$SummaryCsvPath = Join-Path $OutputRoot "Intune_Export_Category_Summary.csv"

# 兼容 Windows PowerShell 5.1 (UTF8 默认带 BOM) 与 PowerShell 7+ (utf8BOM)
$CsvEncoding = if ($PSVersionTable.PSVersion.Major -ge 6) { "utf8BOM" } else { "UTF8" }
$AllReviewRows | Export-Csv -Path $FullCsvPath -NoTypeInformation -Encoding $CsvEncoding
$CandidateRows = @($AllReviewRows | Where-Object { $_.IsAllWindowsCandidate -eq $true })
$CandidateRows | Export-Csv -Path $CandidatesCsvPath -NoTypeInformation -Encoding $CsvEncoding
$CategorySummary | Export-Csv -Path $SummaryCsvPath -NoTypeInformation -Encoding $CsvEncoding

Write-Host "`n================================================================" -ForegroundColor Cyan
Write-Host "  导出完成！汇总统计如下：" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
$CategorySummary | Format-Table -AutoSize | Out-String | Write-Host

Write-Host "1. 全量 Assignment 审核表 (含所有平台与未分配项):" -ForegroundColor Yellow
Write-Host "   $FullCsvPath"
Write-Host "2. 疑似针对 [All Windows] 的重点审核表 (共 $($CandidateRows.Count) 条分配记录):" -ForegroundColor Yellow
Write-Host "   $CandidatesCsvPath"
Write-Host "3. 原始配置与 Assignment JSON 完整备份目录:" -ForegroundColor Yellow
Write-Host "   $BackupJsonRoot"
Write-Host "`n下一步操作建议：" -ForegroundColor Cyan
Write-Host "  请用 Excel 打开上述 CSV 文件，检查需要排除 PAW 的行，将最后一列 [AddPAWExclusion] 改为 'Yes'，然后使用 Add-IntunePAWExclusion_21V.ps1 执行排除。"
