<#
.SYNOPSIS
    根据审核后的 CSV 表格，在世纪互联 (21Vianet) Intune 中批量为指定的 App、Script、Policy、Remediation 添加 PAW 排除 (Exclusion)。

.DESCRIPTION
    本脚本读取 Export-IntuneAssignments_21V.ps1 导出的 CSV 文件，筛选出 [AddPAWExclusion] 列标记为 'Yes' / 'True' / '1' 的对象，
    并在保留原有全部 Assignments 的基础上，追加 PAW Exclusion Group（或设置 PAW Assignment Filter Exclude）。

    内置多重安全保护：
      1. 默认支持 -WhatIf 与 -DryRun 预览模式，不加 -ExecuteChanges 不会真正写入任何修改。
      2. 修改任意策略/应用前，自动在 PreChange_Backup 目录保存一份实时 Assignments JSON 快照，方便随时回滚。
      3. 幂等去重：如果某个对象已经排除了目标 PAW Group，自动跳过，不会重复添加。
      4. 跨类型排除告警：如果某个对象的 Include 目标是 [All Users]，而排除的是 [PAW 设备组]，会自动发出风险提示
         （Intune 平台不支持在 All Users / 用户组分配下排除设备组，此类场景建议使用 PAW 用户组或 Assignment Filter）。

.PARAMETER CsvPath
    经过人工 Review 并把需要修改的行标记为 AddPAWExclusion = Yes 的 CSV 文件路径。

.PARAMETER PAWExclusionGroupId
    要添加为排除项 (Exclude Group) 的 PAW 组 Object ID (GUID)。

.PARAMETER PAWAssignmentFilterId
    （可选替代方案）如果部分策略分配给了 All Users，需要通过 Assignment Filter (Mode=exclude) 排除 PAW 设备，可指定 Filter GUID。

.PARAMETER TenantId
    世纪互联租户 ID (GUID) 或域名 (例如: contoso.partner.onmschina.cn)

.PARAMETER ClientId
    Service Principal 的 Application (client) ID（需具备对应 ReadWrite 权限）

.PARAMETER ClientSecret
    Service Principal 的客户端密钥

.PARAMETER UseInteractive
    使用 Connect-MgGraph -Environment China 交互式登录

.PARAMETER ExecuteChanges
    开关参数。如果不加此参数，脚本默认以 DryRun (只预览不修改) 模式运行，确保安全。

.NOTES
    ================================================================================
    ⚠️ DISCLAIMER / 免责声明 (USE AT YOUR OWN RISK)
    ================================================================================
    This sample script is not supported under any Microsoft standard support program or service.
    The sample script is provided AS IS without warranty of any kind. Microsoft further disclaims
    all implied warranties including, without limitation, any implied warranties of merchantability
    or of fitness for a particular purpose. The entire risk arising out of the use or performance
    of the sample scripts and documentation remains with you.

    本示例脚本按“原样 (AS IS)”提供，不提供任何明示或暗示的担保。
    使用本脚本对生产环境进行批量变更的风险完全由使用者自行承担 (USE AT YOUR OWN RISK)。
    在附加 -ExecuteChanges 参数正式写入生产环境前，请务必：
      1. 仔细核对 CSV 表格中标记为 AddPAWExclusion = Yes 的每一行记录；
      2. 先在默认 DryRun 预览模式下运行并检查预期的变更输出；
      3. 先在测试租户或小范围试点策略上验证效果，并妥善保留 PreChange_Backup 快照。
    ================================================================================
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ })]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [string]$PAWExclusionGroupId,

    [Parameter(Mandatory = $false)]
    [string]$PAWAssignmentFilterId,

    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,

    [Parameter(Mandatory = $false)]
    [switch]$UseInteractive,

    [Parameter(Mandatory = $false)]
    [switch]$ExecuteChanges
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($PAWExclusionGroupId) -and [string]::IsNullOrWhiteSpace($PAWAssignmentFilterId)) {
    throw "请至少提供 -PAWExclusionGroupId (PAW 排除组 GUID) 或 -PAWAssignmentFilterId (PAW 分配筛选器 GUID) 之一。"
}

$GraphBaseUri = "https://microsoftgraph.chinacloudapi.cn"
$AuthorityBase = "https://login.partner.microsoftonline.cn"

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$PreChangeBackupDir = Join-Path $ScriptDir "PreChange_Backup_$Timestamp"
New-Item -ItemType Directory -Path $PreChangeBackupDir -Force | Out-Null

$IsDryRun = -not $ExecuteChanges.IsPresent
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Intune 21Vianet PAW Exclusion 批量添加工具" -ForegroundColor Cyan
Write-Host "  [DISCLAIMER] Provided 'AS IS' without warranty. USE AT YOUR OWN RISK." -ForegroundColor DarkYellow
Write-Host "  [免责声明] 本脚本按原样提供，请先通过 DryRun 及测试验证，风险自担。" -ForegroundColor DarkYellow
if ($IsDryRun) {
    Write-Host "  当前模式: [DRY-RUN 预览模式] (不会对租户做任何实际修改)" -ForegroundColor Yellow
    Write-Host "  确认预览无误后，请附加 -ExecuteChanges 参数正式执行写入。" -ForegroundColor Yellow
} else {
    Write-Host "  当前模式: [EXECUTE 正式写入模式] (将修改租户 Assignment 配置)" -ForegroundColor Red
}
Write-Host "================================================================" -ForegroundColor Cyan

# ============================================================================
# 1. 模块检查与身份认证：自动安装/导入 Microsoft.Graph.Authentication 并连接 21V
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
        "DeviceManagementApps.ReadWrite.All",
        "DeviceManagementConfiguration.ReadWrite.All",
        "DeviceManagementManagedDevices.ReadWrite.All",
        "DeviceManagementScripts.ReadWrite.All",
        "Group.Read.All"
    )

    if (-not $UseInteractive -and (-not [string]::IsNullOrWhiteSpace($ClientId))) {
        if ([string]::IsNullOrWhiteSpace($TenantId)) {
            $TenantId = Read-Host "请输入 21V TenantId"
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
}

function Invoke-21VGraphRequest {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("GET", "POST")]
        [string]$Method,

        [Parameter(Mandatory = $true)]
        [string]$RelativeUri,

        [Parameter(Mandatory = $false)]
        [object]$BodyObject
    )

    $uri = if ($RelativeUri.StartsWith("/")) { "$GraphBaseUri$RelativeUri" } else { "$GraphBaseUri/$RelativeUri" }

    if ($Method -eq "GET") {
        return Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject -ErrorAction Stop
    } else {
        if ($IsDryRun) {
            throw "[安全拦截] 当前处于 DryRun 模式，底层已硬性阻断所有 POST/写操作 ($uri)！"
        }
        $jsonBody = $BodyObject | ConvertTo-Json -Depth 20 -Compress
        return Invoke-MgGraphRequest -Method POST -Uri $uri -Body $jsonBody -ContentType "application/json" -OutputType PSObject -ErrorAction Stop
    }
}

# ============================================================================
# 2. 各 Intune 类别的读取与 Assign 更新路由表
# ============================================================================
function Get-CategoryEndpointConfig {
    param([string]$Category)
    switch ($Category) {
        "App" {
            return @{
                GetAssignmentsUri = "/beta/deviceAppManagement/mobileApps/{0}/assignments"
                PostAssignUri     = "/beta/deviceAppManagement/mobileApps/{0}/assign"
                PayloadRootKey    = "mobileAppAssignments"
            }
        }
        "AppConfig_ManagedDevice" {
            return @{
                GetAssignmentsUri = "/beta/deviceAppManagement/mobileAppConfigurations/{0}/assignments"
                PostAssignUri     = "/beta/deviceAppManagement/mobileAppConfigurations/{0}/assign"
                PayloadRootKey    = "assignments"
            }
        }
        "AppConfig_ManagedApp" {
            return @{
                GetAssignmentsUri = "/beta/deviceAppManagement/targetedManagedAppConfigurations/{0}/assignments"
                PostAssignUri     = "/beta/deviceAppManagement/targetedManagedAppConfigurations/{0}/assign"
                PayloadRootKey    = "assignments"
            }
        }
        "SettingsCatalog" {
            return @{
                GetAssignmentsUri = "/beta/deviceManagement/configurationPolicies/{0}/assignments"
                PostAssignUri     = "/beta/deviceManagement/configurationPolicies/{0}/assign"
                PayloadRootKey    = "assignments"
            }
        }
        "DeviceConfig" {
            return @{
                GetAssignmentsUri = "/beta/deviceManagement/deviceConfigurations/{0}/assignments"
                PostAssignUri     = "/beta/deviceManagement/deviceConfigurations/{0}/assign"
                PayloadRootKey    = "assignments"
            }
        }
        "AdminTemplate" {
            return @{
                GetAssignmentsUri = "/beta/deviceManagement/groupPolicyConfigurations/{0}/assignments"
                PostAssignUri     = "/beta/deviceManagement/groupPolicyConfigurations/{0}/assign"
                PayloadRootKey    = "assignments"
            }
        }
        "CompliancePolicy" {
            return @{
                GetAssignmentsUri = "/beta/deviceManagement/deviceCompliancePolicies/{0}/assignments"
                PostAssignUri     = "/beta/deviceManagement/deviceCompliancePolicies/{0}/assign"
                PayloadRootKey    = "assignments"
            }
        }
        "EndpointSecurity" {
            return @{
                GetAssignmentsUri = "/beta/deviceManagement/intents/{0}/assignments"
                PostAssignUri     = "/beta/deviceManagement/intents/{0}/assign"
                PayloadRootKey    = "assignments"
            }
        }
        "PowerShellScript" {
            return @{
                GetAssignmentsUri = "/beta/deviceManagement/deviceManagementScripts/{0}/assignments"
                PostAssignUri     = "/beta/deviceManagement/deviceManagementScripts/{0}/assign"
                PayloadRootKey    = "deviceManagementScriptAssignments"
            }
        }
        "ProactiveRemediation" {
            return @{
                GetAssignmentsUri = "/beta/deviceManagement/deviceHealthScripts/{0}/assignments"
                PostAssignUri     = "/beta/deviceManagement/deviceHealthScripts/{0}/assign"
                PayloadRootKey    = "deviceHealthScriptAssignments"
            }
        }
        default {
            throw "不支持的 Category 类型: $Category"
        }
    }
}

# ============================================================================
# 3. 读取 CSV 并按 (Category + ItemId) 去重聚合待处理项
# ============================================================================
$csvRows = Import-Csv -Path $CsvPath
$selectedRows = @($csvRows | Where-Object {
    $_.AddPAWExclusion -match "^(?i)(yes|true|1|y)$" -and (-not [string]::IsNullOrWhiteSpace($_.ItemId))
})

if ($selectedRows.Count -eq 0) {
    Write-Warning "在 CSV 文件 [$CsvPath] 中没有找到任何 [AddPAWExclusion] = 'Yes' 的行，脚本退出。"
    return
}

# 按 Category + ItemId 分组，防止同一个策略多条分配记录触发多次重复更新
$groupedItems = @($selectedRows | Group-Object -Property Category, ItemId)
Write-Host "[Init] 从 CSV 中共筛选出 $($groupedItems.Count) 个唯一策略/应用/脚本需要添加 PAW Exclusion。" -ForegroundColor Green

Initialize-21VGraphAuth

$ExecutionLog = [System.Collections.Generic.List[object]]::new()

foreach ($grp in $groupedItems) {
    $firstRow = $grp.Group[0]
    $cat = [string]$firstRow.Category
    $itemId = [string]$firstRow.ItemId
    $itemName = [string]$firstRow.ItemName

    Write-Host "`n----------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "[Processing] [$cat] $itemName (ID: $itemId)" -ForegroundColor Cyan

    $ep = Get-CategoryEndpointConfig -Category $cat

    # 1) 实时读取当前最新的 Assignments
    $currentAssignmentsResp = Invoke-21VGraphRequest -Method GET -RelativeUri ($ep.GetAssignmentsUri -f $itemId)
    $currentAssignments = @()
    if ($null -ne $currentAssignmentsResp.value) {
        $currentAssignments = @($currentAssignmentsResp.value)
    }

    # 保存修改前快照
    $preBackupFile = Join-Path $PreChangeBackupDir ("{0}_{1}_Before.json" -f $cat, $itemId)
    $preBackupPayload = @{
        Category    = $cat
        ItemId      = $itemId
        ItemName    = $itemName
        Timestamp   = (Get-Date).ToUniversalTime().ToString("o")
        Assignments = $currentAssignments
    } | ConvertTo-Json -Depth 20
    [System.IO.File]::WriteAllText($preBackupFile, $preBackupPayload, [System.Text.Encoding]::UTF8)

    # 2) 检查是否包含 All Users 分配，若是且只提供了 Group Exclusion，给出醒目提示
    $hasAllUsers = $false
    foreach ($a in $currentAssignments) {
        if ($a.target -and $a.target."@odata.type" -eq "#microsoft.graph.allLicensedUsersAssignmentTarget") {
            $hasAllUsers = $true
        }
    }
    if ($hasAllUsers -and (-not [string]::IsNullOrWhiteSpace($PAWExclusionGroupId)) -and [string]::IsNullOrWhiteSpace($PAWAssignmentFilterId)) {
        Write-Warning "  [重要提示] '$itemName' 包含 [All Users] 分配！在 Intune 中，All Users 分配不支持排除 Device Group（只能排除 User Group 或使用 Assignment Filter）。请确认您的 PAW 组类型或改用 -PAWAssignmentFilterId。"
    }

    # 3) 构造更新后的 Assignments 数组（完整保留现有分配）
    $updatedAssignments = [System.Collections.Generic.List[object]]::new()
    $alreadyExcluded = $false

    foreach ($a in $currentAssignments) {
        $t = $a.target
        if ($t -and $t."@odata.type" -eq "#microsoft.graph.exclusionGroupAssignmentTarget" -and $t.groupId -eq $PAWExclusionGroupId) {
            $alreadyExcluded = $true
        }

        # 如果指定了 PAWAssignmentFilterId，则在 Include 分配上挂载 Exclude Filter
        if (-not [string]::IsNullOrWhiteSpace($PAWAssignmentFilterId) -and $t -and $t."@odata.type" -ne "#microsoft.graph.exclusionGroupAssignmentTarget") {
            $t | Add-Member -NotePropertyName "deviceAndAppManagementAssignmentFilterId" -NotePropertyValue $PAWAssignmentFilterId -Force
            $t | Add-Member -NotePropertyName "deviceAndAppManagementAssignmentFilterType" -NotePropertyValue "exclude" -Force
        }

        # 清理只读或无需回传的底层元数据以保证 /assign 接口兼容
        $cleanAssignment = [ordered]@{}
        if ($cat -eq "App") {
            $cleanAssignment["@odata.type"] = "#microsoft.graph.mobileAppAssignment"
            $cleanAssignment["intent"] = if ($a.intent) { $a.intent } else { "required" }
            $cleanAssignment["target"] = $a.target
            if ($null -ne $a.settings) {
                $cleanAssignment["settings"] = $a.settings
            }
        }
        elseif ($cat -eq "ProactiveRemediation") {
            $cleanAssignment["@odata.type"] = "#microsoft.graph.deviceHealthScriptAssignment"
            $cleanAssignment["target"] = $a.target
            $cleanAssignment["runRemediationScript"] = [bool]$a.runRemediationScript
            if ($null -ne $a.runSchedule) {
                $cleanAssignment["runSchedule"] = $a.runSchedule
            }
        }
        elseif ($cat -eq "PowerShellScript") {
            $cleanAssignment["@odata.type"] = "#microsoft.graph.deviceManagementScriptAssignment"
            $cleanAssignment["target"] = $a.target
        }
        else {
            $cleanAssignment["target"] = $a.target
        }

        $updatedAssignments.Add($cleanAssignment)
    }

    if ($alreadyExcluded -and [string]::IsNullOrWhiteSpace($PAWAssignmentFilterId)) {
        Write-Host "  -> [Skipped] PAW Group ($PAWExclusionGroupId) 已经存在于 '$itemName' 的 Exclusion 列表中，无需重复添加。" -ForegroundColor Green
        $ExecutionLog.Add([PSCustomObject]@{
            Category = $cat
            ItemName = $itemName
            ItemId   = $itemId
            Status   = "Skipped (Already Excluded)"
        })
        continue
    }

    # 4) 追加新的 PAW Exclusion Group Assignment
    if (-not [string]::IsNullOrWhiteSpace($PAWExclusionGroupId) -and (-not $alreadyExcluded)) {
        $exclTarget = [ordered]@{
            "@odata.type" = "#microsoft.graph.exclusionGroupAssignmentTarget"
            "groupId"     = $PAWExclusionGroupId
        }

        $newExclAssignment = [ordered]@{}
        if ($cat -eq "App") {
            # 取该 App 现有 Include 分配的第一个 intent（如 required 或 available）
            $primaryIntent = "required"
            $inc = @($currentAssignments | Where-Object { $_.target."@odata.type" -ne "#microsoft.graph.exclusionGroupAssignmentTarget" })
            if ($inc.Count -gt 0 -and $inc[0].intent) {
                $primaryIntent = $inc[0].intent
            }
            $newExclAssignment["@odata.type"] = "#microsoft.graph.mobileAppAssignment"
            $newExclAssignment["intent"] = $primaryIntent
            $newExclAssignment["target"] = $exclTarget
        }
        elseif ($cat -eq "ProactiveRemediation") {
            $newExclAssignment["@odata.type"] = "#microsoft.graph.deviceHealthScriptAssignment"
            $newExclAssignment["target"] = $exclTarget
            $newExclAssignment["runRemediationScript"] = $false
        }
        elseif ($cat -eq "PowerShellScript") {
            $newExclAssignment["@odata.type"] = "#microsoft.graph.deviceManagementScriptAssignment"
            $newExclAssignment["target"] = $exclTarget
        }
        else {
            $newExclAssignment["target"] = $exclTarget
        }

        $updatedAssignments.Add($newExclAssignment)
    }

    $postBody = [ordered]@{
        $ep.PayloadRootKey = $updatedAssignments.ToArray()
    }

    $postUri = $ep.PostAssignUri -f $itemId

    if ($IsDryRun) {
        $plannedFile = Join-Path $PreChangeBackupDir ("{0}_{1}_DryRun_PlannedPostBody.json" -f $cat, $itemId)
        [System.IO.File]::WriteAllText($plannedFile, ($postBody | ConvertTo-Json -Depth 20), [System.Text.Encoding]::UTF8)
        Write-Host "  -> [DRY-RUN] 将调用 POST $postUri" -ForegroundColor Yellow
        Write-Host "  -> [DRY-RUN] 原有分配数: $($currentAssignments.Count) -> 更新后分配数: $($updatedAssignments.Count)" -ForegroundColor Yellow
        Write-Host "  -> [DRY-RUN] 拟发送的 JSON 请求体已保存至: $(Split-Path $plannedFile -Leaf)" -ForegroundColor DarkYellow
        $ExecutionLog.Add([PSCustomObject]@{
            Category = $cat
            ItemName = $itemName
            ItemId   = $itemId
            Status   = "DryRun Preview Ready"
        })
    }
    else {
        if ($PSCmdlet.ShouldProcess("[$cat] $itemName ($itemId)", "POST $postUri 添加 PAW Exclusion")) {
            try {
                Invoke-21VGraphRequest -Method POST -RelativeUri $postUri -BodyObject $postBody | Out-Null
                Write-Host "  -> [SUCCESS] 已成功为 '$itemName' 添加 PAW Exclusion！" -ForegroundColor Green
                $ExecutionLog.Add([PSCustomObject]@{
                    Category = $cat
                    ItemName = $itemName
                    ItemId   = $itemId
                    Status   = "Updated Successfully"
                })
            }
            catch {
                Write-Warning "  -> [FAILED] 更新 '$itemName' 失败: $($_.Exception.Message)"
                $ExecutionLog.Add([PSCustomObject]@{
                    Category = $cat
                    ItemName = $itemName
                    ItemId   = $itemId
                    Status   = "Failed: $($_.Exception.Message)"
                })
            }
        }
    }
}

$LogCsvPath = Join-Path $PreChangeBackupDir "PAW_Exclusion_Execution_Result.csv"
$CsvEncoding = if ($PSVersionTable.PSVersion.Major -ge 6) { "utf8BOM" } else { "UTF8" }
$ExecutionLog | Export-Csv -Path $LogCsvPath -NoTypeInformation -Encoding $CsvEncoding

Write-Host "`n================================================================" -ForegroundColor Cyan
Write-Host "  处理完成！执行结果汇总：" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
$ExecutionLog | Format-Table -AutoSize | Out-String | Write-Host
Write-Host "修改前备份与执行日志目录: $PreChangeBackupDir" -ForegroundColor Yellow
