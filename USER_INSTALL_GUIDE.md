# Installing CodeText Add-in for Microsoft Word

Simple instructions for end users to install the CodeText qualitative data coding tool.

---

## What is CodeText?

CodeText is a Microsoft Word add-in that helps you code qualitative data directly in your Word documents. It allows you to:

- Highlight and code text passages
- Organize codes in hierarchical categories
- Generate formatted reports
- Export your coded data

---

## Before You Start

You'll need:
- Microsoft Word (2016 or later, Office 365)
- The `manifest.xml` file (provided by your administrator)
- Internet connection (for initial installation)

---

## Installation Steps

### For Windows:

1. **Open Microsoft Word**
   - Start a new blank document or open an existing one

2. **Access Add-ins**
   - Click the **Insert** tab in the ribbon
   - Click **Add-ins** (or **Get Add-ins** in some versions)
   - Click **My Add-ins** on the left sidebar

3. **Upload the Add-in**
   - At the bottom of the dialog, you'll see **Upload My Add-in**
   - Click **Browse**
   - Select the `manifest.xml` file you downloaded
   - Click **Upload**

4. **Verify Installation**
   - A panel should open on the right side of Word
   - You'll see "CodeText" at the top of the panel
   - If successful, you're ready to start coding!

### For Mac:

1. **Open Microsoft Word**
   - Start a new blank document or open an existing one

2. **Access Add-ins**
   - Click the **Insert** tab in the ribbon
   - Click **Add-ins**

3. **Upload the Add-in**
   - Click **My Add-ins**
   - At the bottom, click **+ Add a custom add-in**
   - Select **Add from file...**
   - Select the `manifest.xml` file you downloaded
   - Click **Add**

4. **Verify Installation**
   - A panel should open on the right side of Word
   - You'll see "CodeText" at the top of the panel
   - If successful, you're ready to start coding!

---

## Opening CodeText After Installation

Once installed, you can open CodeText in two ways:

### Method 1: Ribbon Button
- Go to the **Home** tab
- Look for the **CodeText** group in the ribbon
- Click **Show CodeText**

### Method 2: Add-ins Menu
- Click **Insert** tab → **Add-ins**
- Click **My Add-ins**
- Select **CodeText**

---

## Quick Start Tutorial

### 1. Select a Project
- At the top of the CodeText panel, choose your project from the dropdown
- Examples: DIANA, JOHNSON, RIDE, etc.

### 2. View Categories
- You'll see a hierarchical list of categories
- Click the **▶** arrow to expand categories with subcategories
- Click any category name to select it

### 3. Code Some Text
- In your Word document, select (highlight) the text you want to code
- Click **Get Selected Text** in the CodeText panel
- The selected text will appear in the panel
- Make sure you have a category selected (highlighted in blue)
- Click **Code Selected Text**
- The text will be highlighted in your document with the category's color!

### 4. Generate a Report
- After coding several passages, click **Generate Report**
- A formatted report will be inserted at the end of your document
- The report groups all your coded text by category

### 5. Export Your Data
- Click **Export Data** to download your coded data as JSON
- This creates a backup you can import later

---

## Features Overview

### Category Tree
- Hierarchical categories organized by domain
- Color-coded for easy identification
- Expandable/collapsible subcategories
- Search box to filter categories

### Text Coding
- Select text in Word → Click button → Text is coded!
- Each category has a unique highlight color
- Code the same text with multiple categories
- Add optional comments to coded passages

### Reports
- Generate formatted HTML reports
- Organized by category
- Shows all coded passages
- Includes participant information (if applicable)

### Data Management
- Export coded data as JSON
- Import categories from files
- Clear all data (with confirmation)

---

## Troubleshooting

### The add-in panel won't open

**Try these steps:**
1. Close and reopen Microsoft Word
2. Go to Insert → Add-ins → My Add-ins
3. Make sure CodeText is listed
4. If not, reinstall using the installation steps above

### "This add-in is not trusted" error

**Solution:**
- This means your organization hasn't approved the add-in
- Contact your IT administrator
- They may need to add it to the trusted catalog

### The add-in loads but shows an error

**Check:**
- Do you have an internet connection?
- Try closing and reopening the panel
- Check if there are any Word updates available

### Categories aren't loading

**Solution:**
1. Click **Import Categories** button
2. Your administrator should provide a `test-categories.json` file
3. Select that file when prompted

### Text isn't being highlighted

**Make sure:**
- You've selected text first (it should be highlighted blue in Word)
- You clicked "Get Selected Text"
- You've selected a category (it should be highlighted blue in the panel)
- Then click "Code Selected Text"

### Changes aren't saving

**Remember:**
- All data is stored locally in your browser
- Export your data regularly as backup
- Coded text is saved in your Word document (via highlighting)
- The report captures all your codes

---

## Tips & Best Practices

### 1. Save Often
- Word autosave works as normal
- Export your data periodically as backup
- Your highlights are saved with the document

### 2. Use Search
- Type in the category search box to quickly find categories
- Especially useful with large category lists

### 3. Organize Your Work
- Use one Word document per interview/transcript
- Code as you review the transcript
- Generate reports after each document or at the end

### 4. Color Coding
- Each category has a unique color
- This makes it easy to see patterns at a glance
- Colors are automatically assigned from your category list

### 5. Multiple Coders
- Each person can install CodeText independently
- Share the same category file for consistency
- Compare coded documents to check inter-rater reliability

---

## Keyboard Shortcuts

Currently, CodeText doesn't have keyboard shortcuts, but you can:
- Use Word's normal text selection shortcuts
- Click buttons in the panel with your mouse
- Use Tab to navigate between panel elements

---

## Privacy & Data

### What data is stored?
- All your coded data is stored **locally** in your browser
- No data is sent to external servers
- Your Word document contains the highlighted text
- Reports are generated locally in your document

### What data is sent over the internet?
- Only the add-in code and category definitions are loaded from the server
- Your actual coded text never leaves your computer

### Can others see my work?
- Only if you share your Word document with them
- The add-in itself doesn't share any data

---

## Getting Help

### Need assistance?

**Contact your project administrator:**
- They provided the manifest file
- They have the category definitions
- They can help with technical issues

**Word Add-in General Help:**
- Microsoft Support: https://support.microsoft.com/office
- Search for "Word Add-ins" in Microsoft's help center

---

## Uninstalling

If you need to remove CodeText:

### Windows:
1. Go to Insert → Add-ins → My Add-ins
2. Right-click on CodeText
3. Select **Remove**

### Mac:
1. Go to Insert → Add-ins → My Add-ins
2. Click the **...** menu next to CodeText
3. Select **Remove**

**Note:** Uninstalling the add-in does NOT remove your highlighted text or reports from your Word documents.

---

## Version Information

- Current Version: 1.33
- Last Updated: January 2025
- Requires: Microsoft Word 2016 or later

---

**Happy Coding! 📊**

For questions or issues, contact your project administrator.

