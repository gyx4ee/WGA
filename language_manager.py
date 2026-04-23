from __future__ import annotations

import json
import subprocess
from dataclasses import dataclass


BDS_TIP = "0402:00000402"
PHONETIC_TIP = "0402:00020402"
TRADITIONAL_TIP = "0402:00040402"


@dataclass
class LanguageStatus:
    has_bulgarian: bool
    has_language_pack: bool
    has_bds: bool
    has_phonetic: bool
    has_traditional: bool
    capability_name: str
    summary: str


def _run_powershell(script: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        check=False,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


def get_language_status() -> LanguageStatus:
    script = r"""
$list = Get-WinUserLanguageList
$bg = $list | Where-Object { $_.LanguageTag -eq 'bg-BG' } | Select-Object -First 1
$caps = @()
try {
    $caps = @(Get-WindowsCapability -Online -ErrorAction Stop | Where-Object { $_.Name -like 'Language.*bg-BG*' })
} catch {
    $caps = @()
}
$cap = $caps | Where-Object { $_.Name -like 'Language.Basic*bg-BG*' } | Select-Object -First 1
$installedLanguage = $null
try {
    $installedLanguage = Get-InstalledLanguage -ErrorAction Stop | Where-Object { $_.LanguageId -eq 'bg-BG' -or $_.Language -eq 'bg-BG' } | Select-Object -First 1
} catch {
    $installedLanguage = $null
}
$dismText = ''
try {
    $dismText = (dism /Online /Get-Capabilities /Format:Table | Out-String)
} catch {
    $dismText = ''
}
$capInstalled = [bool]($cap -and $cap.State -eq 'Installed')
$anyBgCapabilityInstalled = [bool]($caps | Where-Object { $_.State -eq 'Installed' } | Select-Object -First 1)
$installedLanguageDetected = [bool]$installedLanguage
$dismBasicInstalled = [bool]($dismText -match 'Language\.Basic~~~bg-BG' -and $dismText -match 'Installed')
$tips = @()
if ($bg) { $tips = @($bg.InputMethodTips) }
[pscustomobject]@{
    has_bulgarian = [bool]$bg
    has_language_pack = [bool]($capInstalled -or $installedLanguageDetected -or $dismBasicInstalled -or $anyBgCapabilityInstalled)
    has_bds = [bool]($tips -contains '0402:00000402')
    has_phonetic = [bool]($tips -contains '0402:00020402')
    has_traditional = [bool]($tips -contains '0402:00040402')
    capability_name = if ($cap) { $cap.Name } elseif ($installedLanguageDetected) { 'Get-InstalledLanguage:bg-BG' } else { '' }
    summary = if ($bg) { ($tips -join ', ') } elseif ($installedLanguageDetected) { 'Bulgarian language pack detected, but bg-BG is not in the user language list' } else { 'Bulgarian language entry not found' }
} | ConvertTo-Json -Compress
"""
    result = _run_powershell(script)
    if result.returncode != 0:
        error_text = (result.stderr or result.stdout or "Language status check failed.").strip()
        raise RuntimeError(error_text)

    payload = json.loads((result.stdout or "").strip() or "{}")
    return LanguageStatus(
        has_bulgarian=bool(payload.get("has_bulgarian")),
        has_language_pack=bool(payload.get("has_language_pack")),
        has_bds=bool(payload.get("has_bds")),
        has_phonetic=bool(payload.get("has_phonetic")),
        has_traditional=bool(payload.get("has_traditional")),
        capability_name=str(payload.get("capability_name", "")),
        summary=str(payload.get("summary", "")),
    )


def build_language_action(action_id: str, status: LanguageStatus) -> tuple[str, str]:
    if action_id == "language_refresh":
        return ("Refresh Language Status", "No command is required.")

    if action_id == "toggle_bulgarian_bds":
        if status.has_bds:
            return (
                "Remove Bulgarian BDS",
                "$list=Get-WinUserLanguageList; $bg=$list|Where-Object {$_.LanguageTag -eq 'bg-BG'}|Select-Object -First 1; "
                "if($bg -and $bg.InputMethodTips -contains '0402:00000402'){[void]$bg.InputMethodTips.Remove('0402:00000402'); Set-WinUserLanguageList $list -Force}",
            )
        return (
            "Add Bulgarian BDS",
            "$list=Get-WinUserLanguageList; $bg=$list|Where-Object {$_.LanguageTag -eq 'bg-BG'}|Select-Object -First 1; "
            "if(-not $bg){$list.Add('bg-BG'); $bg=$list|Where-Object {$_.LanguageTag -eq 'bg-BG'}|Select-Object -First 1}; "
            "if($bg.InputMethodTips -notcontains '0402:00000402'){[void]$bg.InputMethodTips.Add('0402:00000402')}; Set-WinUserLanguageList $list -Force",
        )

    if action_id == "toggle_bulgarian_phonetic":
        if status.has_phonetic:
            return (
                "Remove Bulgarian Phonetic",
                "$list=Get-WinUserLanguageList; $bg=$list|Where-Object {$_.LanguageTag -eq 'bg-BG'}|Select-Object -First 1; "
                "if($bg -and $bg.InputMethodTips -contains '0402:00020402'){[void]$bg.InputMethodTips.Remove('0402:00020402'); Set-WinUserLanguageList $list -Force}",
            )
        return (
            "Add Bulgarian Phonetic",
            "$list=Get-WinUserLanguageList; $bg=$list|Where-Object {$_.LanguageTag -eq 'bg-BG'}|Select-Object -First 1; "
            "if(-not $bg){$list.Add('bg-BG'); $bg=$list|Where-Object {$_.LanguageTag -eq 'bg-BG'}|Select-Object -First 1}; "
            "if($bg.InputMethodTips -notcontains '0402:00020402'){[void]$bg.InputMethodTips.Add('0402:00020402')}; Set-WinUserLanguageList $list -Force",
        )

    if action_id == "toggle_bulgarian_traditional":
        if status.has_traditional:
            return (
                "Remove Bulgarian Traditional Phonetic",
                "$list=Get-WinUserLanguageList; $bg=$list|Where-Object {$_.LanguageTag -eq 'bg-BG'}|Select-Object -First 1; "
                "if($bg -and $bg.InputMethodTips -contains '0402:00040402'){[void]$bg.InputMethodTips.Remove('0402:00040402'); Set-WinUserLanguageList $list -Force}",
            )
        return (
            "Add Bulgarian Traditional Phonetic",
            "$list=Get-WinUserLanguageList; $bg=$list|Where-Object {$_.LanguageTag -eq 'bg-BG'}|Select-Object -First 1; "
            "if(-not $bg){$list.Add('bg-BG'); $bg=$list|Where-Object {$_.LanguageTag -eq 'bg-BG'}|Select-Object -First 1}; "
            "if($bg.InputMethodTips -notcontains '0402:00040402'){[void]$bg.InputMethodTips.Add('0402:00040402')}; Set-WinUserLanguageList $list -Force",
        )

    if action_id == "toggle_bulgarian_language_pack":
        if status.has_language_pack:
            return (
                "Remove Bulgarian Language Pack",
                "$n=(Get-WindowsCapability -Online|Where-Object {$_.Name -like 'Language.Basic*bg-BG*'}|Select-Object -First 1).Name; "
                "if($n){Remove-WindowsCapability -Online -Name $n}",
            )
        return (
            "Install Bulgarian Language Pack",
            "$n=(Get-WindowsCapability -Online|Where-Object {$_.Name -like 'Language.Basic*bg-BG*'}|Select-Object -First 1).Name; "
            "if($n){Add-WindowsCapability -Online -Name $n}",
        )

    if action_id == "remove_bulgarian_language":
        return (
            "Remove Bulgarian Language Entry",
            "$list=Get-WinUserLanguageList; $bg=$list|Where-Object {$_.LanguageTag -eq 'bg-BG'}|Select-Object -First 1; "
            "if($bg){[void]$list.Remove($bg); Set-WinUserLanguageList $list -Force}",
        )

    raise KeyError(action_id)
