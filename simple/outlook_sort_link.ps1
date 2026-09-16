#Requires -Version 5.1
<#
.SYNOPSIS
  Сортирует почту классического Outlook и создаёт связанные черновики.

.DESCRIPTION
  Берёт выделенные письма или непрочитанные из Входящих, собирает тред
  через GetConversation(), ставит категории и сохраняет ответ через Reply().
  Так черновик остаётся в той же переписке. Send() скрипт не вызывает.

.EXAMPLE
  powershell -File simple/outlook_sort_link.ps1
  powershell -File simple/outlook_sort_link.ps1 -Inbox
  powershell -File simple/outlook_sort_link.ps1 -Inbox -DryRun
#>
param(
    [switch]$Inbox,
    [switch]$All,
    [int]$Limit = 200,
    [switch]$DryRun,
    [switch]$DraftTriage
)

$ErrorActionPreference = "Stop"
$olMail = 43
$olFolderInbox = 6

function Get-NormalizedSubject([string]$Subject) {
    $text = ($Subject -replace "\s+", " ").Trim()
    do {
        $prev = $text
        $text = [regex]::Replace($text, "^(?i)((re|fw|fwd|aw|sv|vs|отв|на|пересл)\s*[:：]\s*)+", "")
        $text = $text.Trim()
    } while ($text -ne $prev)
    return $text.ToLowerInvariant()
}

function Get-SmtpAddress($Item) {
    $addr = [string]$Item.SenderEmailAddress
    if ($addr -match "@") { return $addr }
    try {
        $ex = $Item.Sender.GetExchangeUser()
        if ($ex -and $ex.PrimarySmtpAddress) { return [string]$ex.PrimarySmtpAddress }
    } catch {}
    return $addr
}

function Get-CurrentUserSmtp($Namespace) {
    try {
        $smtp = $Namespace.CurrentUser.PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
        )
        if ($smtp) { return [string]$smtp }
    } catch {}
    try { return [string]$Namespace.CurrentUser.Address } catch { return "" }
}

function Get-ThreadItems($Namespace, $Item) {
    $found = @()
    try {
        $conv = $Item.GetConversation()
        if ($null -ne $conv) {
            $table = $conv.GetTable()
            while (-not $table.EndOfTable) {
                $row = $table.GetNextRow()
                try {
                    $entryId = $row.Item("EntryID")
                    $found += $Namespace.GetItemFromID($entryId)
                } catch {}
            }
        }
    } catch {}
    if ($found.Count -eq 0) { return @($Item) }
    return $found | Sort-Object ReceivedTime
}

function Test-IsFyi($Mail) {
    $blob = "$(Get-SmtpAddress $Mail) $($Mail.Subject) $($Mail.Body)"
    return $blob -match "no-?reply|noreply|mailer-daemon|notifications?@|unsubscribe|отписаться|newsletter|рассылк"
}

function Test-NeedsReply($Mail) {
    $body = [string]$Mail.Body
    return ($body -match "\?") -or ($body -match "(?i)\b(прошу|нужно|необходим|подтверд|уточн|пожалуйста|please|asap|срочно)\b")
}

function Get-LastInbound($ThreadItems, [string]$Me) {
    $inbound = @($ThreadItems | Where-Object {
        (Get-SmtpAddress $_) -and ((Get-SmtpAddress $_).ToLower() -ne $Me.ToLower())
    })
    if ($inbound.Count -gt 0) { return $inbound[-1] }
    return $ThreadItems[-1]
}

function New-ThreadDigest($ThreadItems, [string]$Category, [string]$Me) {
    $lines = @(
        "Категория: $Category",
        "Связанные письма этого треда ($($ThreadItems.Count)):",
        ""
    )
    $i = 1
    foreach ($mail in $ThreadItems) {
        $from = Get-SmtpAddress $mail
        $mine = ""
        if ($Me -and $from.ToLower() -eq $Me.ToLower()) { $mine = " (я)" }
        $stamp = Get-Date $mail.ReceivedTime -Format "yyyy-MM-dd HH:mm"
        $lines += "$i. $stamp $from$mine`: $($mail.Subject)"
        $i++
    }
    $lines += @("", "--- Черновик ответа ---", "", "")
    return ($lines -join "`r`n")
}

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$me = Get-CurrentUserSmtp $namespace

$seed = @()
if (-not $Inbox) {
    $selection = $outlook.ActiveExplorer().Selection
    if ($selection.Count -lt 1) {
        throw "В Outlook ничего не выделено. Выделите письмо или укажите -Inbox."
    }
    for ($i = 1; $i -le $selection.Count; $i++) {
        $item = $selection.Item($i)
        if ([int]$item.Class -ne $olMail) { continue }
        $seed += $item
    }
} else {
    $folder = $namespace.GetDefaultFolder($olFolderInbox)
    $items = $folder.Items
    $items.Sort("[ReceivedTime]", $true)
    foreach ($item in $items) {
        if ([int]$item.Class -ne $olMail) { continue }
        if (-not $All -and -not $item.UnRead) { continue }
        $seed += $item
        if ($seed.Count -ge $Limit) { break }
    }
}

$groups = @{}
foreach ($item in $seed) {
    foreach ($mail in (Get-ThreadItems $namespace $item)) {
        if ([int]$mail.Class -ne $olMail) { continue }
        $cid = [string]$mail.ConversationID
        if ([string]::IsNullOrWhiteSpace($cid)) {
            $cid = "subj:" + (Get-NormalizedSubject $mail.Subject)
        }
        if (-not $groups.ContainsKey($cid)) { $groups[$cid] = @{} }
        $groups[$cid][$mail.EntryID] = $mail
    }
}

$drafts = 0
$threadCount = 0
foreach ($cid in $groups.Keys) {
    $threadCount++
    $threadItems = @($groups[$cid].Values | Sort-Object ReceivedTime)
    $lastInbound = Get-LastInbound $threadItems $me
    if (Test-IsFyi $lastInbound) {
        $category = "FYI"
        $needDraft = $false
    } elseif (Test-NeedsReply $lastInbound) {
        $category = "Reply"
        $needDraft = $true
    } else {
        $category = "Triage"
        $needDraft = [bool]$DraftTriage
    }
    $short = (Get-NormalizedSubject $lastInbound.Subject)
    if ($short.Length -gt 40) { $short = $short.Substring(0, 40) }
    if ([string]::IsNullOrWhiteSpace($short)) { $short = "без темы" }
    $label = "$category, Связка:$short"
    Write-Host ("{0} | {1} писем | {2}" -f $category, $threadItems.Count, $lastInbound.Subject)

    if ($DryRun) { continue }

    foreach ($mail in $threadItems) {
        try {
            $mail.Categories = $label
            $mail.Save()
        } catch {}
    }
    if ($needDraft) {
        $reply = $lastInbound.Reply()
        $digest = New-ThreadDigest $threadItems $category $me
        $reply.Body = $digest + "`r`n" + [string]$reply.Body
        $reply.Categories = $label
        $reply.Save()
        $drafts++
    }
}

Write-Host "Тредов: $threadCount. Связанных черновиков: $drafts. Отправку делает человек."
