# Thin direct-COM wrapper for create_visio_from_plan.ps1.
# Use when the proxy queue is stale or the user asks to output immediately.

param(
    [Parameter(Mandatory = $true)]
    [string]$PlanPath,

    [Parameter(Mandatory = $true)]
    [string]$OutVsdx,

    [Alias("OutPreview")]
    [string]$OutEmf = "",

    [switch]$Visible,

    [switch]$KeepVisioOpen
)

$ErrorActionPreference = "Stop"

$mainScript = Join-Path $PSScriptRoot "create_visio_from_plan.ps1"
if (-not (Test-Path -LiteralPath $mainScript)) {
    throw "Main Visio generator not found: $mainScript"
}

$params = @{
    PlanPath = $PlanPath
    OutVsdx = $OutVsdx
    PreferDirect = $true
}
if ($OutEmf) { $params.OutEmf = $OutEmf }
if ($Visible) { $params.Visible = $true }
if ($KeepVisioOpen) { $params.KeepVisioOpen = $true }

& $mainScript @params
exit $LASTEXITCODE
