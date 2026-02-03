# RA Doc Tagging Add-in

A Microsoft Word add-in for research analysis document tagging with hierarchical categories, automatic participant parsing, and report generation.

## Features

- 🏷️ Code text directly in Word with hierarchical categories
- 📊 Automatic participant and demographic parsing
- 📝 Generate formatted summary reports
- 💾 Export coded data to JSON
- 🔄 Multi-project support (DIANA, JOHNSON, JOHNSON2, RIDE)

## Installation

1. **Download the manifest file:**
   - [Download manifest.xml](https://github.com/premisedata/ra-textcode-plugin/raw/refs/heads/gh-pages/manifest.xml)

2. **Install in Word (Mac):**
   ```bash
   # Copy to Word's add-in folder
   cp manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml
   ```

3. **Install in Word (Windows):**
   - Save `manifest.xml` to: `%USERPROFILE%\AppData\Local\Microsoft\Office\16.0\Wef\`
   - Or use Word's **Insert > Add-ins > My Add-ins > Upload My Add-in**

4. **Restart Word** and open "RA Doc Tagging" from the Home tab

## Usage

1. **Select text** in your Word document
2. **Choose a category** from the tree (categories auto-load by project)
3. **Click "Code Selected Text"** - text is highlighted and tagged with a comment
4. **Generate Report** to create a formatted summary at the end of your document

The add-in automatically:
- Detects participant IDs (P1:, P2:, etc.)
- Parses demographics from `<(gender, age, occupation, location)>` format
- Identifies project and country from filename

## Technical Details

- **Built with:** Office.js, vanilla JavaScript
- **Hosted on:** GitHub Pages
- **Storage:** Client-side only (in-memory during session)
- **Deployment:** Automatic via GitHub Actions on push to `gh-pages`

## Project Structure

- `manifest.xml` - Office Add-in manifest
- `taskpane.html` - Main UI
- `taskpane-simple.js` - Application logic
- `taskpane.css` - Styling
- `test-categories.json` - Category hierarchy data
- `assets/` - Icon files

## License

Internal use only.
