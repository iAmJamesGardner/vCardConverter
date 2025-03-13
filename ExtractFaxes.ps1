params($contact)

$contacts = $contact -split "`r`n"
$phones = @{
    Home = @()
    Business = @()
    Mobile = @()
    Uncategorized = @()
}

$matches = ($contacts | Select-String -Pattern "^TEL(;[Tt][Yy][Pp][Ee]=.*?)?:(.+)").Matches | ? { $_ -match 'FAX' }

foreach ($match in $matches) {
    $typeInfo = $match.Groups[1].Value -replace "^;type=",''
    $number = $match.Groups[2].Value.Trim()
    $isPref = $typeInfo -imatch 'PREF'

    if ($typeInfo -imatch 'HOME') {
        if ($isPref) { $phones['HOME'] = @($number) + $phones['HOME'] } else { $phones['Home'] += $number}
    } elseif ($typeInfo -imatch 'WORK') {
        if ($isPref) { $phones['Business'] = @($number) + $phones['Business'] } else { $phones['Business'] += $number}
    } else {
        $phones['Uncategorized'] += $number
    }
}

$data = @{}
if ($phones['Home'].Count -gt 0) { $data['Home Fax'] = $phones['Home'][0] }
if ($phones['Home'].Count -gt 1) { $data['Home Fax 2'] = $phones['Home'][1] }
if ($phones['Business'].Count -gt 0) { $data['Business Fax'] = $phones['Business'][0] }
if ($phones['Business'].Count -gt 1) { $data['Business Fax 2'] = $phones['Business'][1] }
if ($phones['Mobile'].Count -gt 0) { $data['Mobile Fax'] = $phones['Mobile'][0] }
if ($phones['Mobile'].Count -gt 1) { $data['Mobile Fax 2'] = $phones['Mobile'][1] }

if ($phones['Mobile'].Count -gt 2) {
    if (-not $data['Home Fax']) { $data['Home Fax'] = $phones['Mobile'][2] } elseif (-not $data['Home Fax 2']) { $data['Home Fax 2'] = $phones['Mobile'][2] }
}

foreach ($tel in $phones['Uncategorized']) {
    if (-not $data['Mobile Fax']) { $data['Mobile Fax'] = $tel }
    elseif (-not $data['Mobile Fax 2']) { $data['Mobile Fax 2'] = $tel }
    elseif (-not $data['Home Fax']) { $data['Home Fax'] = $tel }
    elseif (-not $data['Home Fax 2']) { $data['Home Fax 2'] = $tel }
    elseif (-not $data['Other Fax']) { $data['Other Fax'] = $tel }
}

return $data