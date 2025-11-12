# Define variables
$url = "https://codeload.github.com/revolsc15/companythemes/zip/refs/heads/main"
$destination = "$env:APPDATA\CompanyThemes"
$extractedPath = "$destination\companythemes-main"
$tempFile = "$env:TEMP\companythemes.zip"

# Microsoft Templates paths
$templatesPath = "$env:APPDATA\Microsoft\Templates"
$documentThemesPath = "$templatesPath\Document Themes"
$oldNormalDotm = "$templatesPath\Normal.dotm"
$newNormalDotm = "$extractedPath\Normal.dotm"
$backupNormalDotm = "$templatesPath\Normal.dotm.backup"

# Excel paths
$excelXLSTARTPath = "$env:APPDATA\Microsoft\Excel\XLSTART"
$excelTemplateSource = "$extractedPath\WCPL_Excel Theme.xltx"
$excelTemplateDest = "$excelXLSTARTPath\WCPL_Excel Theme.xltx"
$excelBook1Template = "$excelXLSTARTPath\Book.xltx"

# Files to check for existence
$filesToCheck = @(
    "Default Theme.thmx",
    "Normal.dotm",
    "themes",
    "WCPL_Excel Theme.xltx",
    "WCPL_PPT Theme.thmx",
    "WCPL_WordTheme.thmx"
)

function Process-DocumentThemes {
    Write-Host "`nProcessing Document Themes folder and theme files..."
    
    # Create Document Themes folder if it doesn't exist
    if (!(Test-Path $documentThemesPath)) {
        try {
            New-Item -ItemType Directory -Path $documentThemesPath -Force
            Write-Host "✓ Created Document Themes directory: $documentThemesPath"
        }
        catch {
            Write-Error "Failed to create Document Themes directory: $_"
            return
        }
    } else {
        Write-Host "Document Themes directory already exists: $documentThemesPath"
    }
    
    # Define theme files to copy
    $themeFiles = @(
        @{
            Source = "$extractedPath\Default Theme.thmx"
            Destination = "$documentThemesPath\Default Theme.thmx"
            Name = "Default Theme.thmx"
        },
        @{
            Source = "$extractedPath\WCPL_PPT Theme.thmx"
            Destination = "$documentThemesPath\WCPL_PPT Theme.thmx"
            Name = "WCPL_PPT Theme.thmx"
        },
        @{
            Source = "$extractedPath\WCPL_WordTheme.thmx"
            Destination = "$documentThemesPath\WCPL_WordTheme.thmx"
            Name = "WCPL_WordTheme.thmx"
        }
    )
    
    # Copy each theme file
    foreach ($themeFile in $themeFiles) {
        if (Test-Path $themeFile.Source) {
            try {
                Copy-Item -Path $themeFile.Source -Destination $themeFile.Destination -Force
                Write-Host "✓ Copied $($themeFile.Name) to Document Themes"
            }
            catch {
                Write-Warning "Failed to copy $($themeFile.Name): $_"
            }
        } else {
            Write-Warning "Source theme file not found: $($themeFile.Source)"
        }
    }
    
    Write-Host "Document Themes processing completed!"
}

function Process-NormalDotm {
    Write-Host "`nProcessing Normal.dotm file..."

    # Ensure Microsoft Templates directory exists
    if (!(Test-Path $templatesPath)) {
        New-Item -ItemType Directory -Path $templatesPath -Force
        Write-Host "Created Microsoft Templates directory: $templatesPath"
    }

    # Check if source Normal.dotm exists
    if (!(Test-Path $newNormalDotm)) {
        Write-Warning "Source Normal.dotm not found at: $newNormalDotm"
        Write-Host "Skipping Normal.dotm processing."
        return
    }

    # Remove existing backup if it exists
    if (Test-Path $backupNormalDotm) {
        try {
            Remove-Item -Path $backupNormalDotm -Force
            Write-Host "✓ Removed existing backup file"
        }
        catch {
            Write-Warning "Failed to remove existing backup: $_"
        }
    }

    # Rename existing Normal.dotm if it exists
    if (Test-Path $oldNormalDotm) {
        try {
            Rename-Item -Path $oldNormalDotm -NewName "Normal.dotm.backup" -Force
            Write-Host "✓ Renamed existing Normal.dotm to Normal.dotm.backup"
        }
        catch {
            Write-Error "Failed to rename existing Normal.dotm: $_"
            exit 1
        }
    } else {
        Write-Host "No existing Normal.dotm found to rename."
    }

    # Copy new Normal.dotm to Microsoft Templates
    try {
        Copy-Item -Path $newNormalDotm -Destination $templatesPath -Force
        Write-Host "✓ Copied new Normal.dotm to Microsoft Templates directory"
    }
    catch {
        Write-Error "Failed to copy new Normal.dotm: $_"
        
        # Try to restore backup if copy failed
        if (Test-Path $backupNormalDotm) {
            try {
                Rename-Item -Path $backupNormalDotm -NewName "Normal.dotm" -Force
                Write-Host "✓ Restored backup Normal.dotm"
            }
            catch {
                Write-Warning "Failed to restore backup Normal.dotm"
            }
        }
        exit 1
    }

    Write-Host "Normal.dotm has been successfully updated!"
}

function Process-ExcelTemplate {
    Write-Host "`nProcessing Excel template..."

    # Check if Book.xltx already exists
    if (Test-Path $excelBook1Template) {
        Write-Host "Book.xltx already exists in XLSTART directory. Skipping Excel template processing."
        return
    }

    # Check if source Excel template exists
    if (!(Test-Path $excelTemplateSource)) {
        Write-Warning "Source Excel template not found at: $excelTemplateSource"
        Write-Host "Skipping Excel template processing."
        return
    }

    # Ensure Excel XLSTART directory exists
    if (!(Test-Path $excelXLSTARTPath)) {
        try {
            New-Item -ItemType Directory -Path $excelXLSTARTPath -Force
            Write-Host "Created Excel XLSTART directory: $excelXLSTARTPath"
        }
        catch {
            Write-Error "Failed to create Excel XLSTART directory: $_"
            return
        }
    }

    # Copy Excel template to XLSTART directory
    try {
        Copy-Item -Path $excelTemplateSource -Destination $excelXLSTARTPath -Force
        Write-Host "✓ Copied WCPL_Excel Theme.xltx to Excel XLSTART directory"
    }
    catch {
        Write-Error "Failed to copy Excel template: $_"
        return
    }

    # Rename the template to Book.xltx
    try {
        Rename-Item -Path $excelTemplateDest -NewName "Book.xltx" -Force
        Write-Host "✓ Renamed WCPL_Excel Theme.xltx to Book.xltx"
    }
    catch {
        Write-Error "Failed to rename Excel template: $_"
        
        # Clean up the copied file if rename failed
        if (Test-Path $excelTemplateDest) {
            try {
                Remove-Item -Path $excelTemplateDest -Force
                Write-Host "✓ Cleaned up temporary Excel template file"
            }
            catch {
                Write-Warning "Failed to clean up temporary Excel template file"
            }
        }
        return
    }

    Write-Host "Excel template has been successfully installed!"
}

# Create destination directory if it doesn't exist
if (!(Test-Path $destination)) {
    New-Item -ItemType Directory -Path $destination -Force
    Write-Host "Created directory: $destination"
}

# Check if all files already exist
$allFilesExist = $true
if (Test-Path $extractedPath) {
    Write-Host "Scanning for existing files..."
    foreach ($file in $filesToCheck) {
        $fullPath = Join-Path $extractedPath $file
        if (Test-Path $fullPath) {
            Write-Host "✓ Found: $file"
        } else {
            Write-Host "✗ Missing: $file"
            $allFilesExist = $false
        }
    }
} else {
    Write-Host "Extracted directory not found: $extractedPath"
    $allFilesExist = $false
}

# If all files exist, process templates and exit
if ($allFilesExist) {
    Write-Host "All required files are already present. Skipping download."
    Process-DocumentThemes
    Process-NormalDotm
    Process-ExcelTemplate
    exit 0
}

# Download the file
try {
    Write-Host "Downloading company themes..."
    Invoke-WebRequest -Uri $url -OutFile $tempFile
    Write-Host "Download completed successfully."
}
catch {
    Write-Error "Failed to download the file: $_"
    exit 1
}

# Extract the zip file
try {
    Write-Host "Extracting files to $destination..."
    Expand-Archive -Path $tempFile -DestinationPath $destination -Force
    Write-Host "Extraction completed successfully."
}
catch {
    Write-Error "Failed to extract the archive: $_"
    exit 1
}

# Verify extraction was successful
Write-Host "Verifying extracted files..."
$verificationPassed = $true
foreach ($file in $filesToCheck) {
    $fullPath = Join-Path $extractedPath $file
    if (Test-Path $fullPath) {
        Write-Host "✓ Verified: $file"
    } else {
        Write-Host "✗ Not found: $file"
        $verificationPassed = $false
    }
}

# Clean up temporary file
Remove-Item $tempFile -Force
Write-Host "Temporary file cleaned up."

if ($verificationPassed) {
    Write-Host "Company themes have been successfully installed to: $destination"
} else {
    Write-Warning "Some files may be missing. Please check the extraction."
}

# Process templates
Process-DocumentThemes
Process-NormalDotm
Process-ExcelTemplate