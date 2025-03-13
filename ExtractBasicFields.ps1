params($contact)

$contacts = $contact -split "`r`n"
$data = @{}

$fieldMap = @{
    'FN' = 'Name'
    'ORG' = 'Company'
    'TITLE' = 'Job Title'
    'BDAY' = 'Birthday'
    'NOTE' = 'Notes'
}

foreach ($field in $fieldMap.Keys) {
    $pattern = "^$($field)(;.*?)?:(.+)"
    $match = ($contacts | Select-String -pattern $pattern).Matches
    if ($match.Count -gt 0) { $data[$fieldMap[$field]] = $match[0].Groups[2].Value.Trim() }
}


return $data