# Function to split contact entries based on multiple phone numbers
function Split-MultiPhoneEntries {
    param(
        [hashtable]$contactData
    )

    # List to store the resulting contact entries
    $splitContacts = @()

    # Find keys with multiple phone numbers
    $multiPhoneKeys = @()
    foreach ($key in $contactData.Keys) {
        if ($contactData[$key] -is [array] -and $contactData[$key].Count -gt 1) {
            $multiPhoneKeys += $key
        }
    }

    # If no multi-number keys, return the original entry as a single element array
    if ($multiPhoneKeys.Count -eq 0) {
        return @($contactData)
    }

    # Create a list of split contacts based on combinations of multi-number keys
    $combinations = @()

    foreach ($key in $multiPhoneKeys) {
        # Create individual contacts for each phone number
        foreach ($number in $contactData[$key]) {
            # Create a copy of the original contact data
            $newEntry = $contactData.Clone()

            # Replace the array field with a single number
            $newEntry[$key] = $number
            $combinations += $newEntry
        }
    }

    # Combine split contacts with fields that don't have multiple numbers
    foreach ($combination in $combinations) {
        foreach ($key in $contactData.Keys) {
            if (-not $multiPhoneKeys.Contains($key)) {
                $combination[$key] = $contactData[$key]
            }
        }
        # Add the processed combination to the result list
        $splitContacts += $combination
    }

    return $splitContacts
}

# Example usage:
# Simulating a contact with multiple phone numbers in a hashtable
$contactData = @{
    "FN" = "John Doe"
    "TEL;CELL" = @("123456789", "987654321")
    "TEL;WORK" = @("5551234")
    "EMAIL" = "john.doe@example.com"
}

# Process the contact data to split multi-number fields
$splitResults = Split-MultiPhoneEntries -contactData $contactData

# Print the results to visualize the split contacts
foreach ($contact in $splitResults) {
    Write-Host "Contact:"
    foreach ($key in $contact.Keys) {
        Write-Host "  $key: $($contact[$key])"
    }
    Write-Host ""
}
