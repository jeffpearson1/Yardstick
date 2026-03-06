using module ..\Modules\VersionPro.psm1

Describe "VersionPro Constructor and ParseVersion" {
    It "Parses simple dotted version" {
        $v = [VersionPro]::new("1.2.3")
        $v.Segments.Count | Should -Be 3
        $v.Segments[0] | Should -Be 1
        $v.Segments[1] | Should -Be 2
        $v.Segments[2] | Should -Be 3
    }

    It "Parses single segment" {
        $v = [VersionPro]::new("5")
        $v.Segments.Count | Should -Be 1
        $v.Segments[0] | Should -Be 5
    }

    It "Parses version with alpha segments" {
        $v = [VersionPro]::new("1.2.3-beta")
        $v.Segments.Count | Should -Be 4
        $v.Segments[3] | Should -Be "beta"
    }

    It "Parses dash-separated version" {
        $v = [VersionPro]::new("2024-01-15")
        $v.Segments.Count | Should -Be 3
        $v.Segments[0] | Should -Be 2024
        $v.Segments[1] | Should -Be 1
        $v.Segments[2] | Should -Be 15
    }

    It "Preserves OriginalString" {
        $v = [VersionPro]::new("1.2.3")
        $v.OriginalString | Should -Be "1.2.3"
    }

    It "Handles empty string" {
        $v = [VersionPro]::new("")
        $v.Segments.Count | Should -Be 0
    }

    It "Handles whitespace string" {
        $v = [VersionPro]::new("  ")
        $v.Segments.Count | Should -Be 0
    }

    It "Lowercases alpha segments" {
        $v = [VersionPro]::new("1.2.BETA")
        $v.Segments[2] | Should -Be "beta"
    }
}

Describe "VersionPro.CompareTo" {
    It "Equal versions return 0" {
        $v1 = [VersionPro]::new("1.2.3")
        $v2 = [VersionPro]::new("1.2.3")
        $v1.CompareTo($v2) | Should -Be 0
    }

    It "Greater version returns positive" {
        $v1 = [VersionPro]::new("2.0.0")
        $v2 = [VersionPro]::new("1.9.9")
        $v1.CompareTo($v2) | Should -BeGreaterThan 0
    }

    It "Lesser version returns negative" {
        $v1 = [VersionPro]::new("1.0.0")
        $v2 = [VersionPro]::new("1.0.1")
        $v1.CompareTo($v2) | Should -BeLessThan 0
    }

    It "Shorter version with implicit zeros is equal" {
        $v1 = [VersionPro]::new("1.2")
        $v2 = [VersionPro]::new("1.2.0")
        $v1.CompareTo($v2) | Should -Be 0
    }

    It "Shorter version less than longer" {
        $v1 = [VersionPro]::new("1.2")
        $v2 = [VersionPro]::new("1.2.1")
        $v1.CompareTo($v2) | Should -BeLessThan 0
    }

    It "Null comparison returns positive" {
        $v1 = [VersionPro]::new("1.0")
        $v1.CompareTo($null) | Should -BeGreaterThan 0
    }

    It "Throws on non-VersionPro object" {
        $v1 = [VersionPro]::new("1.0")
        { $v1.CompareTo("string") } | Should -Throw "*VersionPro*"
    }

    It "Integer segment before string segment" {
        $v1 = [VersionPro]::new("1.2.3")
        $v2 = [VersionPro]::new("1.2.beta")
        $v1.CompareTo($v2) | Should -BeLessThan 0
    }

    It "String segment after integer segment" {
        $v1 = [VersionPro]::new("1.2.beta")
        $v2 = [VersionPro]::new("1.2.3")
        $v1.CompareTo($v2) | Should -BeGreaterThan 0
    }

    It "Alpha segments compared ordinally" {
        $v1 = [VersionPro]::new("1.0.alpha")
        $v2 = [VersionPro]::new("1.0.beta")
        $v1.CompareTo($v2) | Should -BeLessThan 0
    }
}

Describe "VersionPro.Equals" {
    It "Equal versions return true" {
        $v1 = [VersionPro]::new("1.2.3")
        $v2 = [VersionPro]::new("1.2.3")
        $v1.Equals($v2) | Should -BeTrue
    }

    It "Unequal versions return false" {
        $v1 = [VersionPro]::new("1.2.3")
        $v2 = [VersionPro]::new("1.2.4")
        $v1.Equals($v2) | Should -BeFalse
    }

    It "Null returns false" {
        $v1 = [VersionPro]::new("1.0")
        $v1.Equals($null) | Should -BeFalse
    }

    It "Non-VersionPro returns false" {
        $v1 = [VersionPro]::new("1.0")
        $v1.Equals(42) | Should -BeFalse
    }
}

Describe "VersionPro.GetHashCode" {
    It "Equal versions produce the same hash code" {
        $v1 = [VersionPro]::new("1.2.3")
        $v2 = [VersionPro]::new("1.2.3")
        $v1.GetHashCode() | Should -Be $v2.GetHashCode()
    }

    It "Different versions produce different hash codes" {
        $v1 = [VersionPro]::new("1.0")
        $v2 = [VersionPro]::new("2.0")
        $v1.GetHashCode() | Should -Not -Be $v2.GetHashCode()
    }
}

Describe "VersionPro.ToString" {
    It "Default returns OriginalString" {
        $v = [VersionPro]::new("1.2.3")
        $v.ToString() | Should -Be "1.2.3"
    }

    It "4 segments pads with zeros" {
        $v = [VersionPro]::new("1.2")
        $v.ToString(4) | Should -Be "1.2.0.0"
    }

    It "Fewer segments truncates" {
        $v = [VersionPro]::new("1.2.3.4")
        $v.ToString(2) | Should -Be "1.2"
    }

    It "Zero segments returns empty string" {
        $v = [VersionPro]::new("1.2")
        $v.ToString(0) | Should -Be ""
    }

    It "Negative segments throws" {
        $v = [VersionPro]::new("1.2")
        { $v.ToString(-1) } | Should -Throw "*negative*"
    }
}

Describe "VersionPro Operators" {
    It "-eq for equal versions" {
        $v1 = [VersionPro]::new("1.0")
        $v2 = [VersionPro]::new("1.0")
        [VersionPro]::op_Equality($v1, $v2) | Should -BeTrue
    }

    It "-ne for unequal versions" {
        $v1 = [VersionPro]::new("1.0")
        $v2 = [VersionPro]::new("2.0")
        [VersionPro]::op_Inequality($v1, $v2) | Should -BeTrue
    }

    It "-lt for lesser version" {
        $v1 = [VersionPro]::new("1.0")
        $v2 = [VersionPro]::new("2.0")
        [VersionPro]::op_LessThan($v1, $v2) | Should -BeTrue
    }

    It "-le for equal versions" {
        $v1 = [VersionPro]::new("1.0")
        $v2 = [VersionPro]::new("1.0")
        [VersionPro]::op_LessThanOrEqual($v1, $v2) | Should -BeTrue
    }

    It "-gt for greater version" {
        $v1 = [VersionPro]::new("2.0")
        $v2 = [VersionPro]::new("1.0")
        [VersionPro]::op_GreaterThan($v1, $v2) | Should -BeTrue
    }

    It "-ge for equal versions" {
        $v1 = [VersionPro]::new("2.0")
        $v2 = [VersionPro]::new("2.0")
        [VersionPro]::op_GreaterThanOrEqual($v1, $v2) | Should -BeTrue
    }

    It "-lt with null left returns true" {
        $v2 = [VersionPro]::new("1.0")
        [VersionPro]::op_LessThan($null, $v2) | Should -BeTrue
    }

    It "-gt with null left returns false" {
        $v2 = [VersionPro]::new("1.0")
        [VersionPro]::op_GreaterThan($null, $v2) | Should -BeFalse
    }
}
