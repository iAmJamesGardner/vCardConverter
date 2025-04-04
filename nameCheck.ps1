# Function to generate an N field based on a full name (FN field)
function Get-NFieldFromFullName {
    param(
        [string]$fullName
    )

    # Trim any extra spaces
    $fullName = $fullName.Trim()
    # Split the full name by whitespace
    $parts = $fullName -split "\s+"
    
    # Handle different cases
    if ($parts.Count -eq 0) {
        return ";;;;"
    }
    elseif ($parts.Count -eq 1) {
        # Only one part found; assume it's the given name
        return ";$($parts[0]);;;"
    }
    else {
        # A common assumption is that the last word is the family name
        $lastName = $parts[-1]
        $firstName = $parts[0]
        # If there are words in between, assume they are additional names
        if ($parts.Count -gt 2) {
            $middleNames = $parts[1..($parts.Count - 2)] -join " "
        }
        else {
            $middleNames = ""
        }
        # Construct the N field as "Family;Given;Additional;;" (empty for prefix and suffix)
        return "$lastName;$firstName;$middleNames;;"
    }
}

# Example usage in your vCard converter script:
# Assume you have extracted the FN and N fields from the vCard file
# $vcardFullName holds the full name from the FN field
# $vcardNField holds the existing N field from the vCard

if ($vcardFullName) {
    $expectedNField = Get-NFieldFromFullName -fullName $vcardFullName
    if ($vcardNField -ne $expectedNField) {
        Write-Host "The N field does not match the expected format based on the FN field."
        Write-Host "Updating N field to: $expectedNField"
        # Update the N field accordingly
        $vcardNField = $expectedNField
    }
    else {
        Write-Host "The N field matches the full name format."
    }
}
