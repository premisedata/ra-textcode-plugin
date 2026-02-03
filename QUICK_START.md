# Quick Start: Import Test Data

## Step-by-Step Instructions

### 1. Make Sure Server is Running

```bash
cd TaskPaneMod
python3 server-cors.py
```

Or if using basic server:
```bash
python3 -m http.server 3000
```

### 2. Open Word Add-in in Word

1. Open Microsoft Word
2. The add-in should appear in the ribbon (Home tab > CodeText button)
3. Click the button to open the task pane

### 3. Import Categories

1. **Click "Import Categories"** button in the task pane
2. **File picker will open** - navigate to `TaskPaneMod` folder
3. **Select `test-categories.json`**
4. **Wait for import** - you'll see "Importing categories..." status
5. **Success message** will show how many categories were imported

### 4. View Categories

1. **Select a project** from the dropdown (DIANA, JOHNSON, JOHNSON2, or RIDE)
2. **Category tree will populate** with imported categories
3. **Click on categories** to see the hierarchical structure

### 5. Test Coding

1. **Select some text** in your Word document
2. **Click "Get Selected Text"** in the task pane
3. **Click a category** from the tree
4. **Click "Code Selected Text"**
5. **Text should be highlighted** in yellow in Word
6. **Status shows** "Coded text with category: [category name]"

## Troubleshooting

**"Import Categories" button doesn't work:**
- Check browser console (F12 in Word) for errors
- Make sure file is valid JSON
- Try refreshing the add-in (remove/re-add)

**Categories don't appear:**
- Make sure you selected a project from dropdown
- Check that import completed successfully
- Try importing again

**File not found:**
- Make sure `test-categories.json` is in the `TaskPaneMod` folder
- Check file path is correct

## File Location

The test data file is at:
```
TaskPaneMod/test-categories.json
```

Make sure this file exists before trying to import.

