# How to See Latest Changes in Word Add-in

## Server Behavior

**Yes, the server automatically serves the latest files** - Python's `http.server` and similar servers read files fresh on each request. However, **Word and browsers cache files**, so you need to force a refresh.

## How to See Changes

### Method 1: Hard Refresh in Word (Recommended)

1. **Remove and re-add the add-in:**
   - In Word: **Insert** > **Add-ins** > **My Add-ins**
   - Right-click your add-in > **Remove**
   - Re-upload `manifest.xml` (or `manifest-http.xml`)

2. **Or restart Word completely:**
   - Close Word entirely
   - Re-open Word
   - The add-in should reload with latest files

### Method 2: Clear Browser Cache (Word uses embedded browser)

**Mac:**
```bash
# Clear Word's cache
rm -rf ~/Library/Containers/com.microsoft.Word/Data/Library/Caches/
```

**Windows:**
- Close Word
- Delete: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`

### Method 3: Add Cache-Busting to URLs

Add version parameter to force reload:

In `taskpane.html`, change:
```html
<script src="taskpane.js"></script>
```

To:
```html
<script src="taskpane.js?v=1"></script>
```

Then increment the version number each time you make changes.

### Method 4: Use Browser Dev Tools in Word

1. In Word, press **F12** (or Option+Command+I on Mac)
2. Go to **Network** tab
3. Check "Disable cache" checkbox
4. Reload the add-in

### Method 5: Test in Regular Browser First

Before testing in Word:
1. Open `http://localhost:3000/taskpane.html` in browser
2. Press **Ctrl+Shift+R** (or Cmd+Shift+R on Mac) for hard refresh
3. Verify changes appear
4. Then test in Word

## Quick Refresh Workflow

1. **Make code changes** to `taskpane.js` or `taskpane.html`
2. **Save files**
3. **In Word:** Remove add-in and re-add it (fastest)
4. **Or:** Restart Word completely

## Server Status

The server (`server-cors.py` or `python -m http.server`) **does NOT need to be restarted** - it automatically serves the latest file contents.

## Troubleshooting

**Changes not appearing?**
- ✅ Server is running (check terminal)
- ✅ Files are saved
- ✅ Hard refresh in Word (remove/re-add add-in)
- ✅ Check browser console for errors
- ✅ Verify file paths are correct

**Still not working?**
- Check file permissions
- Verify server is serving from correct directory
- Check for JavaScript errors in console
- Try accessing file directly in browser: `http://localhost:3000/taskpane.js`

