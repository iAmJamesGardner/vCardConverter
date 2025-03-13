param($contact)

$contacts = $contact -split "`r`n"
$data = @{}
$match = ($contacts | Select-String -Pattern "^N:(.+)").Matches

if ($match.Count -gt 0) {
    $parts = $match[0].Groups[1].Value -split ';'
    $data['Last Name'] = $parts[0]
    $data['First Name'] = $parts[1]
    $data['Middle Name'] = $parts[2]
    $data['Suffix'] = $parts[3]
}

return $data