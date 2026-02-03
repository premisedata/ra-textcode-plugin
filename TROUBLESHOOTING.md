# Troubleshooting "Add-in Error" - Network Connectivity

## Error: "Sorry, we can't load the add-in. Please make sure you have network and/or Internet connectivity."

This error means Word can't reach the URLs specified in `manifest.xml`.

## Quick Fixes

### Option 1: Start Local Server (For Testing)

The manifest currently points to `https://localhost:3000`. You need to:

1. **Start a local HTTPS server:**
   ```bash
   cd TaskPaneMod
   
   # Install office-addin-dev-certs for HTTPS
   npm install -g office-addin-dev-certs
   office-addin-dev-certs install --machine
   
   # Start server with HTTPS
   npx office-addin-dev-certs verify
   # Then use a server that supports HTTPS, or:
   npx http-server -p 3000 -S -C cert.pem -K key.pem
   ```

2. **Or use HTTP (may not work, but worth trying):**
   - Update manifest.xml: change `https://localhost:3000` to `http://localhost:3000`
   - Start server: `python -m http.server 3000`
   - Note: Word may reject HTTP, but some versions allow localhost HTTP

### Option 2: Deploy to AWS First (Recommended)

1. Deploy files to AWS (S3 + CloudFront or Amplify)
2. Get your HTTPS URL (e.g., `https://your-domain.cloudfront.net`)
3. Update all URLs in `manifest.xml` to use your AWS domain
4. Re-upload manifest.xml to Word

### Option 3: Use Office Add-in Dev Tools

This automatically sets up localhost with HTTPS:

```bash
npm install -g office-addin-dev-certs
npm install -g office-addin-manifest

cd TaskPaneMod
office-addin-dev-certs install
office-addin-manifest validate manifest.xml
office-addin-manifest sideload manifest.xml
```

## Common Issues

### Issue 1: localhost:3000 Not Accessible

**Symptoms:**
- Error about network connectivity
- Add-in fails to load

**Solutions:**
- Start a web server on port 3000
- Or update manifest.xml to point to your AWS deployment
- Check firewall isn't blocking port 3000

### Issue 2: HTTPS Certificate Errors

**Symptoms:**
- Browser/Word shows certificate warnings
- Add-in still fails to load

**Solutions:**
- Use `office-addin-dev-certs` to generate trusted certificates
- Or deploy to AWS (they handle SSL automatically)

### Issue 3: CORS Errors

**Symptoms:**
- Add-in loads but shows errors in console
- Resources fail to load

**Solutions:**
- Add CORS headers to your server
- Ensure all resources are on same domain or have proper CORS

### Issue 4: Wrong URLs in Manifest

**Symptoms:**
- 404 errors in browser console
- Resources not found

**Solutions:**
- Verify all URLs in manifest.xml are correct
- Check that files exist at those paths
- Ensure paths are relative to server root

## Step-by-Step Debugging

1. **Check Browser Console:**
   - In Word, press F12 to open developer tools
   - Look for errors in Console tab
   - Check Network tab for failed requests

2. **Verify Server is Running:**
   ```bash
   # Test if server responds
   curl http://localhost:3000/taskpane.html
   # Or open in browser
   open http://localhost:3000/taskpane.html
   ```

3. **Verify Manifest URLs:**
   - Open manifest.xml
   - Check all URLs point to accessible locations
   - Ensure HTTPS is used (or HTTP for localhost if allowed)

4. **Test Files Directly:**
   - Open `taskpane.html` directly in browser
   - Check if Pyodide/DuckDB load correctly
   - Verify no JavaScript errors

5. **Check Office.js Loading:**
   - Open browser console in Word (F12)
   - Look for Office.js errors
   - Verify `Office.onReady` fires

## Quick Test Script

Create `test-server.html` to verify your setup:

```html
<!DOCTYPE html>
<html>
<head>
    <title>Server Test</title>
</head>
<body>
    <h1>Server is Running!</h1>
    <p>If you see this, your server is working.</p>
    <script>
        console.log('Server test page loaded');
    </script>
</body>
</html>
```

Place in TaskPaneMod folder and access via `http://localhost:3000/test-server.html`

## Recommended Workflow

1. **First**: Deploy to AWS (get HTTPS URL)
2. **Then**: Update manifest.xml with AWS URLs
3. **Finally**: Sideload manifest.xml in Word

This avoids localhost issues entirely.

## If Nothing Works

1. **Check Word Version:**
   - Some older versions have limited add-in support
   - Ensure you're using Office 365 or Office 2016+

2. **Check IT Policies:**
   - Some organizations block localhost add-ins
   - May need to use network deployment

3. **Try Word Online:**
   - Sometimes Word Online has better error messages
   - Go to office.com and open Word there

4. **Check Firewall/Antivirus:**
   - May be blocking localhost connections
   - Temporarily disable to test

