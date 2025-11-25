# Security Warnings - How to Handle

When you first run the QuoteFlow application, you may see security warnings from Windows Defender or macOS Gatekeeper. This is normal for unsigned applications.

## Windows - "Windows protected your PC" Warning

### Option 1: Allow the app (Recommended)
1. Click **"More info"** on the warning dialog
2. Click **"Run anyway"**
3. The app will launch normally

### Option 2: Add to Windows Defender exclusions
1. Open **Windows Security** (Windows Defender)
2. Go to **Virus & threat protection**
3. Click **"Manage settings"** under Virus & threat protection settings
4. Scroll down to **Exclusions** and click **"Add or remove exclusions"**
5. Click **"Add an exclusion"** → **"File"**
6. Select the `QuoteFlow.exe` file

### Option 3: Unblock the file
1. Right-click on `QuoteFlow.exe`
2. Select **"Properties"**
3. Check **"Unblock"** at the bottom
4. Click **"OK"**

## macOS - "App can't be opened because it is from an unidentified developer"

### Option 1: Allow the app (Recommended)
1. Right-click on `QuoteFlow.app`
2. Select **"Open"** from the context menu
3. Click **"Open"** in the dialog that appears
4. The app will launch normally

### Option 2: Allow from System Preferences
1. Go to **System Preferences** → **Security & Privacy**
2. Click the **"General"** tab
3. Look for a message about QuoteFlow being blocked
4. Click **"Open Anyway"**

### Option 3: Remove quarantine attribute (Terminal)
```bash
sudo xattr -r -d com.apple.quarantine /path/to/QuoteFlow.app
```

## Why This Happens

- The application is **not code-signed** (no digital certificate)
- This is **normal for open-source applications**
- The app is **completely safe** - it's just a quotation management tool
- **No personal data is collected** or sent anywhere

## Verification

You can verify the app is safe by:
1. **Checking the source code** on GitHub
2. **Running antivirus scans** (they should come back clean)
3. **Using the app in a sandboxed environment** if you're concerned

## For Developers

To eliminate these warnings, you would need to:
1. **Purchase a code signing certificate** (Windows: ~$200/year, macOS: $99/year)
2. **Sign the executables** during the build process
3. **Notarize the macOS app** (additional step for macOS)

For most users, the warnings are just a one-time inconvenience and can be safely ignored.
