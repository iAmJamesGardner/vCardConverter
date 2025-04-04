param($contact)

$contacts = $contact -split "`r`n"
$data = @{}
$match = ($contacts | Select-String -Pattern "^N:(.+)").Matches

if ($match.Count -gt 0) {
    $components = $match[0].Groups[1].Value -split ';'
    while ($components.count -lt 5) { $components += "" }

    $family = $components[0].Trim()
    $given = $components[1].Trim()
    $additional = $components[2].Trim()
    $prefix = $components[3].Trim()
    $suffix = $components[4].Trim()

    $fullNameParts = @()
    if ($prefix) {
        $fullNameParts += $prefix
        $data['Title'] = $prefix
    }
    if ($given) {
        $fullNameParts += $given
        $data['First Name'] = $given
    }
    if ($additional) {
        $fullNameParts += $additional
        $data['Middle Name'] = $additional
    }
    if ($family) {
        $fullNameParts += $family
        $data['Last Name'] = $family
    }

    $fullName = $fullNameParts -join " "
    if ($suffix) {
        $fullName = "$fullName, $suffix"
        $data['Suffix'] = $suffix
    }

    $data['Full Name'] = $fullName
}

return $data