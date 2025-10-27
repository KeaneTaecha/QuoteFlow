# Building Windows .exe from macOS

This guide explains how to automatically build Windows .exe files from your macOS development environment using GitHub Actions.

## Prerequisites

1. **GitHub account** (free)
2. **Git installed** on your Mac
3. **Your code** ready to push to GitHub

## Setup Steps

### 1. Initialize Git Repository (if not already done)

```bash
cd /Users/kunanonttaechaaukarkaul/Documents/CMKL/Komfortflow/QuoteFlow
git init
git add .
git commit -m "Initial commit"
```

### 2. Create GitHub Repository

1. Go to [GitHub.com](https://github.com)
2. Click "New repository"
3. Name it `QuoteFlow` (or any name you prefer)
4. Make it **Public** (required for free GitHub Actions)
5. **Don't** initialize with README (you already have files)
6. Click "Create repository"

### 3. Connect Local Repository to GitHub

```bash
git remote add origin https://github.com/YOUR_USERNAME/QuoteFlow.git
git branch -M main
git push -u origin main
```

Replace `YOUR_USERNAME` with your actual GitHub username.

### 4. Trigger the Build

The GitHub Actions workflow will automatically run when you push code. You can also trigger it manually:

1. Go to your repository on GitHub
2. Click "Actions" tab
3. Click "Build Windows Executable"
4. Click "Run workflow" button

### 5. Download the .exe File

After the build completes (usually 2-5 minutes):

1. Go to "Actions" tab
2. Click on the latest successful build
3. Scroll down to "Artifacts"
4. Download "QuoteFlow-Windows.zip"
5. Extract to get `QuoteFlow.exe`

## What the Workflow Does

- **Runs on Windows** - Uses GitHub's Windows runners
- **Installs Python 3.10** - Same version as your Mac
- **Installs dependencies** - From your requirements.txt
- **Builds executable** - Using PyInstaller with optimized settings
- **Creates artifacts** - Uploads the .exe file for download
- **Auto-releases** - Creates GitHub releases for main branch builds

## File Structure

The workflow expects this structure:
```
QuoteFlow/
├── .github/workflows/build-windows.yml  # GitHub Actions workflow
├── src/                                 # Your source code
├── data/quotation_template.xlsx         # Template file
├── prices.db                           # Database file
├── requirements.txt                    # Python dependencies
└── QuoteFlow.spec                      # PyInstaller spec (optional)
```

## Troubleshooting

### Build Fails
- Check the "Actions" tab for error details
- Ensure all required files are committed
- Verify `requirements.txt` has all dependencies

### Missing Files
- Make sure `data/quotation_template.xlsx` exists
- Ensure `prices.db` is present
- Check that all source files are committed

### Large File Size
- The workflow uses the same optimizations as your Mac build
- Should produce a ~30MB .exe file

## Benefits

✅ **Free** - No cost for public repositories
✅ **Automatic** - Builds on every push
✅ **Cross-platform** - Build Windows .exe from Mac
✅ **Reliable** - Uses clean Windows environment
✅ **Versioned** - Each build is tagged and downloadable

## Next Steps

1. Push your code to GitHub
2. Wait for the build to complete
3. Download your Windows .exe file
4. Test it on a Windows machine
5. Share with Windows users!

The .exe file will work on any Windows 10/11 machine without requiring Python or any dependencies to be installed.
