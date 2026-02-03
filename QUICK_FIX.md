# Quick Fix for "Add-in Error" in Word

## If Server is Running But Word Still Shows Error

### Step 1: Test Server in Browser First

1. Open browser and go to: `http://localhost:3000/taskpane.html`
2. Check browser console (F12) for errors
3. If it doesn't load, your server isn't working correctly

### Step 2: Test Debug Page

1. Go to: `http://localhost:3000/debug.html`
2. This will show what's loading and what's failing
3. Check the output for errors

### Step 3: Word-Specific Issues

**Word often requires HTTPS even for localhost.** Try:

#### Option A: Use HTTPS with Self-Signed Certificate

```bash
# Install office-addin-dev-certs
npm install -g office-addin-dev-certs

# Generate certificates
office-addin-dev-certs install --machine

# This creates certificates in:
# ~/.office-addin-dev-certs/
```

Then use a server that supports HTTPS (like `npx http-server -S`)

#### Option B: Enable Developer Mode in Word

**Mac:**
```bash
defaults write com.microsoft.Word OfficeDevModeEnabled -bool true
```

**Windows:**
- Registry: `HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Developer`
- Add DWORD: `EnableDevMode` = 1

Then restart Word.

#### Option C: Check Word's Browser Console

1. In Word, press **F12** (or right-click add-in > Inspect)
2. Check **Console** tab for errors
3. Check **Network** tab for failed requests
4. Look for CORS errors or 404s

### Step 4: Common Fixes

#### Fix 1: Remove and Re-add Add-in

1. In Word: **Insert** > **Add-ins** > **My Add-ins**
2. Right-click your add-in > **Remove**
3. Close Word completely
4. Re-open Word
5. Re-upload manifest.xml

#### Fix 2: Clear Office Cache

**Mac:**
```bash
rm -rf ~/Library/Containers/com.microsoft.Word/Data/Library/Caches/
```

**Windows:**
- Close Word
- Delete: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`

#### Fix 3: Check Firewall

- Ensure port 3000 isn't blocked
- Try a different port (8080, 8000)

#### Fix 4: Use Different Port

Update manifest.xml to use port 8080:
- Change `localhost:3000` to `localhost:8080`
- Start server on port 8080: `python3 -m http.server 8080`

### Step 5: Verify Manifest URLs

Open manifest.xml and verify:
- All URLs use the same protocol (http or https)
- All URLs use the same port
- All URLs are accessible (test in browser)

### Step 6: Check CORS Headers

Your server needs to send CORS headers. Create a simple server with CORS:

**Python with CORS:**
```python
#!/usr/bin/env python3
from http.server import HTTPServer, SimpleHTTPRequestHandler

class CORSRequestHandler(SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()

if __name__ == '__main__':
    server = HTTPServer(('localhost', 3000), CORSRequestHandler)
    print('Server running on http://localhost:3000')
    server.serve_forever()
```

Save as `server-cors.py` and run: `python3 server-cors.py`

### Step 7: Test with Minimal Manifest

Create a minimal test to isolate the issue:

1. Create `test.html` with just:
```html
<!DOCTYPE html>
<html>
<head>
    <script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
</head>
<body>
    <h1>Test</h1>
    <script>
        Office.onReady(() => {
            document.body.innerHTML += '<p>Office.js loaded!</p>';
        });
    </script>
</body>
</html>
```

2. Create minimal `test-manifest.xml` pointing to `test.html`
3. If this works, the issue is with your main files
4. If this doesn't work, the issue is with Word/server setup

### Most Likely Issues

1. **Word requires HTTPS** - Even for localhost, Word prefers HTTPS
2. **CORS blocking** - Server needs CORS headers
3. **Office.js not loading** - Check network tab in Word's console
4. **Wrong manifest** - Using HTTPS manifest but HTTP server (or vice versa)

### Recommended Solution

**Use office-addin-dev-certs for proper HTTPS:**

```bash
npm install -g office-addin-dev-certs
cd TaskPaneMod
office-addin-dev-certs install --machine

# Then use a server that supports the certificates
# Or deploy to AWS which has HTTPS built-in
```

This is the most reliable way to test locally.

