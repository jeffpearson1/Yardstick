using namespace System.Collections.Generic

class VersionPro : System.IComparable {
    [string]$OriginalString
    [List[Object]]$Segments

    # Constructor from string
    VersionPro([string]$versionString) {
        $this.OriginalString = $versionString
        $this.Segments = [List[Object]]::new()
        $this.ParseVersion($versionString)
    }

    # Parse version string into segments
    hidden [void]ParseVersion([string]$versionString) {
        if ([string]::IsNullOrWhiteSpace($versionString)) {
            return
        }

        # Split on any non-alphanumeric character
        $parts = $versionString -split '[^a-zA-Z0-9]+'
        
        foreach ($part in $parts) {
            if ([string]::IsNullOrWhiteSpace($part)) {
                continue
            }
            
            # Try to parse as integer first
            $intValue = 0
            if ([int]::TryParse($part, [ref]$intValue)) {
                $this.Segments.Add($intValue)
            }
            else {
                # Store as lowercase string for case-insensitive comparison
                $this.Segments.Add($part.ToLower())
            }
        }
    }

    # IComparable implementation
    [int]CompareTo([Object]$obj) {
        if ($null -eq $obj) {
            return 1
        }

        if ($obj -isnot [VersionPro]) {
            throw [System.ArgumentException]::new("Object must be of type VersionPro")
        }

        $other = [VersionPro]$obj
        $maxLength = [Math]::Max($this.Segments.Count, $other.Segments.Count)

        for ($i = 0; $i -lt $maxLength; $i++) {
            $thisSegment = if ($i -lt $this.Segments.Count) { $this.Segments[$i] } else { 0 }
            $otherSegment = if ($i -lt $other.Segments.Count) { $other.Segments[$i] } else { 0 }

            # Compare based on types
            $thisIsInt = $thisSegment -is [int]
            $otherIsInt = $otherSegment -is [int]

            if ($thisIsInt -and $otherIsInt) {
                # Both are integers - numeric comparison
                $comparison = $thisSegment.CompareTo($otherSegment)
                if ($comparison -ne 0) {
                    return $comparison
                }
            }
            elseif ($thisIsInt -and -not $otherIsInt) {
                # Integer comes before string
                return -1
            }
            elseif (-not $thisIsInt -and $otherIsInt) {
                # String comes after integer
                return 1
            }
            else {
                # Both are strings - lexicographic comparison
                $comparison = [string]::Compare($thisSegment, $otherSegment, [System.StringComparison]::Ordinal)
                if ($comparison -ne 0) {
                    return $comparison
                }
            }
        }

        return 0
    }

    # IEquatable implementation
    [bool]Equals([Object]$obj) {
        if ($null -eq $obj) {
            return $false
        }

        if ($obj -isnot [VersionPro]) {
            return $false
        }

        return $this.CompareTo($obj) -eq 0
    }

    # Override GetHashCode
    [int]GetHashCode() {
        $hash = 17
        foreach ($segment in $this.Segments) {
            $hash = $hash * 31 + $segment.GetHashCode()
        }
        return $hash
    }

    # Override ToString
    [string]ToString() {
        return $this.OriginalString
    }

    # Overload for ToString with segment number
    [string]ToString([int]$numSegments) {
        if ($numSegments -lt 0) {
            throw [System.ArgumentOutOfRangeException]::new("numSegments", "Number of segments cannot be negative")
        }
        else {
            #return the segments and pad additional ones with .0
            $segmentsToReturn = @()
            for ($i = 0; $i -lt $numSegments; $i++) {
                if ($i -lt $this.Segments.Count) {
                    $segmentsToReturn += $this.Segments[$i]
                }
                else {
                    $segmentsToReturn += 0
                }
            }
            return ($segmentsToReturn -join ".")
        }
    }

    # Comparison operators
    static [bool]op_Equality([VersionPro]$left, [VersionPro]$right) {
        if ($null -eq $left) { return $null -eq $right }
        return $left.Equals($right)
    }

    static [bool]op_Inequality([VersionPro]$left, [VersionPro]$right) {
        return -not [VersionPro]::op_Equality($left, $right)
    }

    static [bool]op_LessThan([VersionPro]$left, [VersionPro]$right) {
        if ($null -eq $left) { return $null -ne $right }
        return $left.CompareTo($right) -lt 0
    }

    static [bool]op_LessThanOrEqual([VersionPro]$left, [VersionPro]$right) {
        if ($null -eq $left) { return $true }
        return $left.CompareTo($right) -le 0
    }

    static [bool]op_GreaterThan([VersionPro]$left, [VersionPro]$right) {
        if ($null -eq $left) { return $false }
        return $left.CompareTo($right) -gt 0
    }

    static [bool]op_GreaterThanOrEqual([VersionPro]$left, [VersionPro]$right) {
        if ($null -eq $left) { return $null -eq $right }
        return $left.CompareTo($right) -ge 0
    }
}

# Export the class
Export-ModuleMember -Function * -Cmdlet * -Variable * -Alias *
