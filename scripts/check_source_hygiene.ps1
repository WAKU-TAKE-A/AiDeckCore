$ErrorActionPreference = "Stop"

$expected_source = @(
    "deck2pptx/",
    "tests/",
    "examples/",
    "docs/",
    "scripts/",
    "LICENSE",
    "README.md",
    "README_AI.md",
    "pyproject.toml",
    ".gitignore"
)

$expected_generated = @(
    ".venv/",
    ".venv-release/",
    "__pycache__/",
    "deck2pptx/__pycache__/",
    "tests/__pycache__/",
    ".pytest_cache/",
    "deck2pptx.egg-info/",
    "outputs/",
    "dist/",
    "build/"
)

$operational_reference = @(
    "_sample/",
    "dual-model-operation-kit/",
    "Inputs/"
)

$statusOutput = git status --short --ignored
$hasUnexpected = $false
$untrackedSource = 0
$operationalUntracked = 0
$generatedArtifacts = 0

Write-Host "`n--- Source Hygiene Report ---"

if ($null -eq $statusOutput -or $statusOutput.Length -eq 0) {
    Write-Host "[OK] Git tree is completely clean and committed."
    exit 0
}

foreach ($line in $statusOutput) {
    # Line format: "XY filename"
    $status = $line.Substring(0, 2)
    $file = $line.Substring(3)
    
    $isSource = $false
    foreach ($src in $expected_source) {
        if ($file.StartsWith($src) -or $file -eq $src) { $isSource = $true; break }
    }
    
    $isGenerated = $false
    foreach ($gen in $expected_generated) {
        if ($file.StartsWith($gen) -or $file -eq $gen) { $isGenerated = $true; break }
    }
    
    $isOperational = $false
    foreach ($op in $operational_reference) {
        if ($file.StartsWith($op) -or $file -eq $op) { $isOperational = $true; break }
    }

    if ($status -eq "!!") {
        if ($isGenerated -or $isOperational) {
            # Expected ignored generated or operational file.
            $generatedArtifacts++
        } else {
            Write-Warning "Unexpected ignored file: $file"
            $hasUnexpected = $true
        }
    } elseif ($status -eq "??") {
        if ($isSource) {
            Write-Host "[Info] Untracked intended source: $file"
            $untrackedSource++
        } elseif ($isOperational) {
            Write-Host "[Info] Untracked operational/reference: $file"
            $operationalUntracked++
        } else {
            Write-Error "Unexpected untracked file: $file"
            $hasUnexpected = $true
        }
    } else {
        # Modified/Staged files
        if ($isSource -or $isOperational) {
            Write-Host "[Info] Modified/Staged source: $file"
        } else {
            Write-Error "Unexpected modified file: $file"
            $hasUnexpected = $true
        }
    }
}

Write-Host "---------------------------"
if ($hasUnexpected) {
    Write-Error "[FAIL] Repository contains unexpected untracked or modified artifacts."
    exit 1
} else {
    if ($untrackedSource -gt 0) {
        Write-Host "[OK] Tree has expected untracked source ($untrackedSource items). Commit-ready."
    } else {
        Write-Host "[OK] Tree is commit-ready."
    }
    exit 0
}
