# How to Sideload Word Add-ins

If you don't see the "Upload My Add-in" option, here are several methods to enable it:

## Method 1: Enable Developer Mode (Windows)

### For Word Desktop (Windows)

1. **Open Registry Editor** (may require admin):
   - Press `Win + R`, type `regedit`, press Enter
   - Navigate to: `HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Developer`
   - Create a new DWORD value: `EnableDevMode`
   - Set value to `1`
   - Restart Word

2. **Or use Group Policy** (if available):
   - Enable "Allow sideloading of Office Add-ins" policy

### For Word Desktop (Mac)

1. **Enable Developer Mode**:
   - Open Terminal
   - Run: `defaults write com.microsoft.Word OfficeDevModeEnabled -bool true`
   - Restart Word

2. **Alternative**: Use Office.js Dev Tools
   - Install Office Add-in Dev Tools: `npm install -g office-addin-dev-certs`
   - This enables localhost development

## Method 2: Use Office Add-in Dev Tools (Recommended)

This is the easiest method and doesn't require admin rights:

### Install Office Add-in Dev Tools

```bash
npm install -g office-addin-dev-certs
npm install -g office-addin-manifest
```

### Sideload via Command Line

```bash
# Navigate to TaskPaneMod folder
cd TaskPaneMod

# Sideload the manifest
office-addin-manifest validate manifest.xml
office-addin-manifest sideload manifest.xml
```

### Or Use Yeoman Generator

```bash
npm install -g yo generator-office
yo office
# Follow prompts to create a basic add-in, then replace files
```

## Method 3: Manual Registry/Configuration (Windows)

If you have admin access:

1. **Enable Trusted Catalogs**:
   - Registry path: `HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs`
   - Add your localhost or domain URL

2. **Enable Developer Mode**:
   - Registry path: `HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Developer`
   - Set `EnableDevMode` = 1

## Method 4: Use Word on the Web

Word on the web sometimes has different sideloading options:

1. Go to [Office.com](https://office.com)
2. Open Word online
3. Go to Insert > Add-ins
4. Look for "Upload My Add-in" or "My Add-ins" > "Upload from file"

## Method 5: Network Share Deployment (Enterprise)

If you're in an organization:

1. **Place manifest.xml on network share**:
   - Upload to SharePoint or network drive
   - Get HTTPS URL to manifest.xml

2. **Deploy via Group Policy** (requires IT admin):
   - Configure Office Add-ins via Group Policy
   - Point to manifest.xml URL

3. **Use Centralized Deployment** (requires admin):
   - Office 365 Admin Center > Settings > Integrated apps
   - Upload manifest.xml
   - Assign to users

## Method 6: Direct URL Loading (Temporary)

Some versions allow loading directly:

1. **Get HTTPS URL to manifest.xml** (from AWS)
2. **In Word**, go to Insert > Add-ins > My Add-ins
3. **Look for "Add from URL"** or similar option
4. **Paste the manifest.xml URL**

## Troubleshooting

### "Upload My Add-in" Option Missing

**Possible causes:**
- IT policy blocking sideloading
- Developer mode not enabled
- Using Office 365 web version (limited support)
- Office version doesn't support sideloading

**Solutions:**
1. Check with IT if sideloading is allowed
2. Try enabling developer mode (Method 1)
3. Use Office Add-in Dev Tools (Method 2)
4. Deploy via network share (Method 5)

### "Add-in failed to load"

**Check:**
- Manifest.xml is accessible via HTTPS
- All URLs in manifest.xml use HTTPS (not HTTP)
- CORS headers are configured on server
- Office.js loads successfully (check browser console)

### Registry Changes Not Working

- Ensure you're editing the correct Office version (16.0 = Office 2016/365)
- Try both `HKEY_CURRENT_USER` and `HKEY_LOCAL_MACHINE`
- Restart Word completely after changes

## Alternative: Test Without Sideloading

You can test the add-in functionality without Word:

1. **Open `taskpane.html` directly in browser** (after deploying to HTTPS)
2. **Remove Office.js dependency** temporarily
3. **Test Pyodide and DuckDB** functionality
4. **Add Word integration** once basic functionality works

## Quick Test Script

Create a test HTML file to verify WASM works:

```html
<!DOCTYPE html>
<html>
<head>
    <script src="https://cdn.jsdelivr.net/pyodide/v0.25.0/full/pyodide.js"></script>
</head>
<body>
    <h1>Pyodide Test</h1>
    <div id="output"></div>
    <script>
        async function test() {
            const pyodide = await loadPyodide();
            const result = pyodide.runPython("2 + 2");
            document.getElementById('output').textContent = result;
        }
        test();
    </script>
</body>
</html>
```

If this works, Pyodide is loading correctly.

## For IT Administrators

To enable sideloading for users:

1. **Group Policy**:
   - Enable "Allow sideloading of Office Add-ins"
   - Path: User Configuration > Policies > Administrative Templates > Microsoft Office 2016 > Security Settings

2. **Registry** (for all users):
   - Set `HKEY_LOCAL_MACHINE\Software\Microsoft\Office\16.0\WEF\Developer\EnableDevMode` = 1

3. **Centralized Deployment**:
   - Deploy via Office 365 Admin Center
   - No sideloading needed for end users

## Next Steps

1. **Try Method 2** (Office Add-in Dev Tools) - easiest, no admin needed
2. **If that fails**, check with IT about sideloading policies
3. **Deploy to AWS first** - get HTTPS URL, then sideload
4. **Test WASM separately** - verify Pyodide/DuckDB work before Word integration

