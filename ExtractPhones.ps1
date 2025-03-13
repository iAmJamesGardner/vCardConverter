params($contact)

$contacts = $contact -split "`r`n"
$phones = @{
    Home = @()
    Business = @()
    Mobile = @()
    Uncategorized = @()
}

$matches = ($contacts | Select-String -Pattern "^TEL(;[Tt][Yy][Pp][Ee]=.*?)?:(.+)").Matches | ? { $_ -notmatch 'FAX' }

foreach ($match in $matches) {
    $typeInfo = $match.Groups[1].Value -replace "^;type=",''
    $number = $match.Groups[2].Value.Trim()
    $isPref = $typeInfo -imatch 'PREF'

    if ($typeInfo -imatch 'HOME') {
        if ($isPref) { $phones['HOME'] = @($number) + $phones['HOME'] } else { $phones['Home'] += $number}
    } elseif ($typeInfo -imatch 'WORK|TELEX') {
        if ($isPref) { $phones['Business'] = @($number) + $phones['Business'] } else { $phones['Business'] += $number}
    } elseif ($typeInfo -imatch 'CELL|IPHONE' -or -not $typeInfo) {
        if ($ispref) { $phones['Mobile'] = @($number) + $phones['Mobile'] } else { $phones['Mobile'] += $number }
    } else {
        $phones['Uncategorized'] += $number
    }
}

$data = @{}
if ($phones['Home'].Count -gt 0) { $data['Home Phone'] = $phones['Home'][0] }
if ($phones['Home'].Count -gt 1) { $data['Home Phone 2'] = $phones['Home'][1] }
if ($phones['Business'].Count -gt 0) { $data['Business Phone'] = $phones['Business'][0] }
if ($phones['Business'].Count -gt 1) { $data['Business Phone 2'] = $phones['Business'][1] }
if ($phones['Mobile'].Count -gt 0) { $data['Mobile Phone'] = $phones['Mobile'][0] }
if ($phones['Mobile'].Count -gt 1) { $data['Mobile Phone 2'] = $phones['Mobile'][1] }

if ($phones['Mobile'].Count -gt 2) {
    if (-not $data['Home Phone']) { $data['Home Phone'] = $phones['Mobile'][2] } elseif (-not $data['Home Phone 2']) { $data['Home Phone 2'] = $phones['Mobile'][2] }
}

foreach ($tel in $phones['Uncategorized']) {
    if (-not $data['Mobile Phone']) { $data['Mobile Phone'] = $tel }
    elseif (-not $data['Mobile Phone 2']) { $data['Mobile Phone 2'] = $tel }
    elseif (-not $data['Home Phone']) { $data['Home Phone'] = $tel }
    elseif (-not $data['Home Phone 2']) { $data['Home Phone 2'] = $tel }
    elseif (-not $data['Other Phone']) { $data['Other Phone'] = $tel }
}

return $data