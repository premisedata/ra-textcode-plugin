# How to Debug Word Add-in Issues

## Step 1: Open Word's Developer Console

**Mac:**
- In Word, go to **View** > **Developer** (if available)
- Or press **Option + Command + I**
- Or right-click the add-in task pane > **Inspect**

**Windows:**
- Press **F12** in Word
- Or right-click add-in > **Inspect Element**

## Step 2: Check Console Tab

Look for errors like:
- `Failed to load resource`
- `CORS policy blocked`
- `Office.js failed to load`
- `404 Not Found`

## Step 3: Check Network Tab

1. Go to **Network** tab
2. Reload the add-in
3. Look for failed requests (red entries)
4. Check which resources are failing:
   - `taskpane.html` - Main page
   - `office.js` - Office JavaScript API
   - `pyodide.js` - Python WASM
   - `duckdb-browser.js` - Database WASM
   - `taskpane.css` - Stylesheet

## Step 4: Common Errors and Fixes

### Error: "Failed to load resource: net::ERR_CONNECTION_REFUSED"
**Fix:** Server isn't running or wrong port

### Error: "CORS policy blocked"
**Fix:** Use `server-cors.py` instead of basic HTTP server

### Error: "Office.js failed to load"
**Fix:** Check internet connection, Office.js loads from CDN

### Error: "404 Not Found"
**Fix:** Check file paths in manifest.xml match actual files

### Error: "Mixed Content" (HTTP/HTTPS mix)
**Fix:** Word requires HTTPS. Use `office-addin-dev-certs` or deploy to AWS

## Step 5: Test Individual Components

1. **Test server:** Open `http://localhost:3000/taskpane.html` in browser
2. **Test debug page:** Open `http://localhost:3000/debug.html` in browser
3. **Check Office.js:** Look for Office.js errors in console
4. **Check WASM:** Look for Pyodide/DuckDB loading errors

## What to Look For

When you open the console, you should see:
- ✓ Office.js loading successfully
- ✓ taskpane.html loading
- ✓ Pyodide initializing
- ✓ DuckDB initializing

If any of these fail, that's your issue.

