# Function to parse the ADR field into its components
function Parse-AddressField {
    param([string]$adrField)

    # Split the ADR field into components by semicolon
    $components = $adrField -split ";"
    while ($components.Count -lt 7) { $components += "" }

    # Map the components to Outlook Webmail fields
    $addressData = @{
        "PO Box"          = $components[0].Trim()
        "Extended Address"= $components[1].Trim()
        "Street"          = $components[2].Trim()
        "City"            = $components[3].Trim()
        "State"           = $components[4].Trim()
        "Postal Code"     = $components[5].Trim()
        "Country"         = $components[6].Trim()
    }

    # Format the address to match Outlook Webmail import fields
    $formattedAddress = @{
        "Street"      = $addressData["Street"]
        "City"        = $addressData["City"]
        "State"       = $addressData["State"]
        "Postal Code" = $addressData["Postal Code"]
        "Country"     = $addressData["Country"]
    }

    # Include PO Box and Extended Address if available
    if ($addressData["PO Box"]) {
        $formattedAddress["Street"] = "PO Box " + $addressData["PO Box"] + " " + $formattedAddress["Street"]
    }
    if ($addressData["Extended Address"]) {
        $formattedAddress["Street"] = $formattedAddress["Street"] + " " + $addressData["Extended Address"]
    }

    return $formattedAddress
}

# Function to extract addresses from the vCard entry
function Get-Addresses {
    param([string]$entry)

    # Regex pattern to match address fields (ADR)
    $addressPattern = "ADR;TYPE=([^:]+):(.*)"
    $matches = [regex]::Matches($entry, $addressPattern)

    $addresses = @{
        "Home" = @()
        "Work" = @()
    }

    foreach ($match in $matches) {
        $type = $match.Groups[1].Value.Trim()
        $address = $match.Groups[2].Value.Trim()

        # Parse the address
        $parsedAddress = Parse-AddressField -adrField $address
        
        if ($type -match "home") {
            $addresses["Home"] += $parsedAddress
        }
        elseif ($type -match "work") {
            $addresses["Work"] += $parsedAddress
        }
    }

    return $addresses
}

# Function to create a formatted address string for CSV
function Format-AddressForCSV {
    param([hashtable]$address)

    if ($address) {
        return "$($address["Street"]), $($address["City"]), $($address["State"]), $($address["Postal Code"]), $($address["Country"])"
    } else {
        return ""
    }
}

# Function to split contacts with multiple addresses
function Split-MultiAddressEntries {
    param([hashtable]$contactData)

    # List to store the resulting contact entries
    $splitContacts = @()

    # Check for multiple addresses
    $multiAddressKeys = @("Home", "Work")
    $excessAddresses = @()

    foreach ($key in $multiAddressKeys) {
        if ($contactData[$key] -is [array] -and $contactData[$key].Count -gt 1) {
            foreach ($address in $contactData[$key]) {
                # Create a copy of the original contact data
                $newEntry = $contactData.Clone()
                # Replace the address field with a single address
                $newEntry[$key] = $address
                $excessAddresses += $newEntry
            }
        }
    }

    # Return the split contacts
    if ($excessAddresses.Count -eq 0) {
        return @($contactData) # Return the original contact as a single element array
    } else {
        return $excessAddresses
    }
}

# Example Usage
$entry = @"
BEGIN:VCARD
FN:John Doe
ADR;TYPE=home:;;123 Main St;Springfield;IL;62701;USA
ADR;TYPE=home:;;456 Another St;Springfield;IL;62702;USA
ADR;TYPE=work:;;789 Corporate Ave;Metropolis;NY;10001;USA
ADR;TYPE=work:;;1010 Office Rd;Gotham;NY;10002;USA
END:VCARD
"@

# Extract addresses from the vCard entry
$parsedAddresses = Get-Addresses -entry $entry

# Example contact data structure
$contactData = @{
    "FN" = "John Doe"
    "Home" = $parsedAddresses["Home"]
    "Work" = $parsedAddresses["Work"]
}

# Split contacts with multiple addresses
$splitResults = Split-MultiAddressEntries -contactData $contactData

# Print the results
foreach ($contact in $splitResults) {
    Write-Host "Contact:"
    foreach ($key in $contact.Keys) {
        Write-Host "  $key: $($contact[$key])"
    }
    Write-Host ""
}
