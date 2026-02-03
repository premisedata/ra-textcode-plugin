# How to Open Console in Word

## Method 1: Keyboard Shortcut (Easiest)

**On Mac:**
- Press **Option + Command + I** (⌥⌘I)
- Or press **F12** (if your Mac supports it)

**On Windows:**
- Press **F12**
- Or press **Ctrl + Shift + I**

## Method 2: Right-Click Menu

1. **Right-click** anywhere in the add-in task pane
2. Look for **"Inspect"** or **"Inspect Element"** in the context menu
3. Click it to open developer tools

## Method 3: Word Menu (if available)

1. In Word, go to **View** menu
2. Look for **Developer** or **Developer Tools** option
3. Click to enable developer mode

## What You'll See

Once the console opens, you'll see:
- **Console tab** - Shows all the debug messages
- **Network tab** - Shows file loading
- **Elements tab** - Shows HTML structure

## What to Look For

In the Console tab, look for messages like:
- "Root categories for DIANA: 7"
- "Political Domain (ID: 1) has X children"
- "Rendering Political Domain: hasChildren=true/false"

These messages will help diagnose why the tree isn't showing.

## If Console Won't Open

1. **Try Word Online** - Sometimes easier to access console
2. **Check Word version** - Older versions might not support it
3. **Use the debug panel** - I've added a visual debug panel in the UI that shows the same info

