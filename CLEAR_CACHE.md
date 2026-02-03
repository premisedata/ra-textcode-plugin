# How to Clear Word Add-in Cache

Word aggressively caches add-in files. Here's how to force it to load the latest version:

## Method 1: Remove and Re-add Add-in (Recommended)

1. **In Word**: Go to **Insert** > **Add-ins** > **My Add-ins**
2. **Right-click** on "CodeText" add-in
3. **Click "Remove"**
4. **Close Word completely**
5. **Re-open Word**
6. **Re-add the add-in**: Insert > Add-ins > Upload My Add-in > Select `manifest.xml`

## Method 2: Clear Word's Cache (Mac)

1. **Close Word completely**
2. **Open Terminal** and run:
   ```bash
   rm -rf ~/Library/Containers/com.microsoft.Word/Data/Library/Caches/
   ```
3. **Re-open Word**
4. **Re-add the add-in** if needed

## Method 3: Clear Word's Cache (Windows)

1. **Close Word completely**
2. **Open File Explorer**
3. **Navigate to**: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
4. **Delete** the cache folders
5. **Re-open Word**

## Method 4: Hard Refresh in Browser Console

1. **In Word**, press **F12** (or Option+Command+I on Mac) to open developer tools
2. **Right-click** the refresh button in the console
3. **Select "Empty Cache and Hard Reload"**
4. **Or** press **Ctrl+Shift+R** (Windows) or **Cmd+Shift+R** (Mac)

## Method 5: Check Version in Console

1. **In Word**, press **F12** to open developer tools
2. **Go to Console tab**
3. **Look for** the script tag - it should show `taskpane.js?v=11`
4. **If it shows an older version**, the cache hasn't cleared

## Verify Latest Version

After clearing cache, check:
1. **Title bar** should show "CodeText v11"
2. **Console** should show `taskpane.js?v=11`
3. **Categories should show expand icons** (▶) for categories with children

## If Still Not Working

1. **Check server is running**: Make sure `server-cors.py` is running
2. **Check file timestamps**: Make sure `taskpane.js` was recently modified
3. **Try a different browser**: Word uses an embedded browser - try Word Online
4. **Check manifest.xml**: Make sure URLs point to the correct server

