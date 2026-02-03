# CodeText Word Add-in (WASM Proof of Concept)

This is a proof of concept Word Add-in that replicates the functionality of the original c0d3t3xt tool using Pyodide (Python WASM) and DuckDB WASM.

## Features

- ✅ Text selection and coding directly in Word
- ✅ Category tree for organizing codes
- ✅ Python functions ported from db-api.py
- ✅ Local database using DuckDB WASM (no VPN needed)
- ✅ Report generation
- ✅ Data export to JSON

## Architecture

- **Pyodide**: Runs Python code in the browser (replicates db-api.py functions)
- **DuckDB WASM**: Local SQL database (replaces PostgreSQL)
- **Office.js**: Word Add-in API for document interaction

## Setup

### 1. Deploy to AWS (or other HTTPS host)

Word Add-ins **require HTTPS** (or localhost for development). Deploy the `TaskPaneMod` folder to:

- **AWS S3 + CloudFront** (recommended for static hosting)
- **AWS Amplify** (simpler setup)
- **Any HTTPS web server** (Apache, Nginx, etc.)

### 2. Update manifest.xml

Update all URLs in `manifest.xml` to point to your AWS domain:

```xml
<!-- Replace these URLs with your AWS domain -->
<SourceLocation DefaultValue="https://your-domain.com/taskpane.html"/>
<AppDomain>https://your-domain.com</AppDomain>
```

### 3. Local Development (Optional)

For local testing, you can use:

**Python HTTP Server:**
```bash
cd TaskPaneMod
python -m http.server 3000
```

Then update manifest.xml to use `https://localhost:3000` (Word allows localhost for development)

### 4. Sideload the add-in in Word

**If you see "Upload My Add-in" option:**
1. Open Microsoft Word
2. Go to **Insert** > **Add-ins** > **My Add-ins**
3. Click **Upload My Add-in**
4. Upload the `manifest.xml` file (can be from local file or hosted URL)
5. The add-in should appear in the ribbon under the Home tab

**If "Upload My Add-in" is missing:**
- See `SIDELOADING_GUIDE.md` for detailed instructions
- May require enabling Developer Mode or using Office Add-in Dev Tools
- IT policies may block sideloading (contact your IT department)
- Alternative: Use `office-addin-manifest sideload` command-line tool

**Note:** For production, you may want to publish the add-in through the Office Store or your organization's add-in catalog.

## Usage

1. **Select text** in your Word document
2. Click **Get Selected Text** in the task pane
3. **Choose a category** from the category tree
4. (Optional) Add a comment
5. Click **Code Selected Text** to save the coded text
6. Use **Generate Report** to create an HTML report at the end of the document
7. Use **Export Data** to download all coded data as JSON

## Files

- `manifest.xml` - Word Add-in manifest (defines the add-in)
- `taskpane.html` - Main UI interface
- `taskpane.js` - Application logic (Pyodide + DuckDB integration)
- `taskpane.css` - Styling
- `python/functions.py` - Python functions reference (loaded via Pyodide in JS)

## Differences from Original

- ✅ **No VPN required** - Everything runs locally in the browser
- ✅ **No backend server** - WASM runs entirely client-side
- ✅ **No Word macros** - Uses Office.js API instead
- ✅ **Local storage only** - Data stored in DuckDB WASM (can export/import)

## Ported Functions

The following functions from `db-api.py` have been ported to Python/Pyodide:

- `parseParticipantNumber()` - Extracts participant numbers (P1:, P2:, etc.)
- `getProjectName()` - Determines project from filename
- `parseCountryName()` - Parses country from DIANA filenames
- `cleanCategory()` - Removes display_order suffixes
- `createCitationJOHNSON()` - Creates citations for JOHNSON project
- `createCitationJOHNSON2()` - Creates citations for JOHNSON2 project
- `createCitationRIDE()` - Creates citations for RIDE project

## Database Schema

The DuckDB schema matches the PostgreSQL schema:

- `c0d3t3xt` - Stores coded text excerpts
- `participants` - Participant demographics
- `categories` - Category hierarchy
- `category_str` - Flattened category strings

## Next Steps

- [ ] Load categories from CSV file (similar to `create-categories.py`)
- [ ] Add participant table management UI
- [ ] Implement full report generation with summaries (port `create_html_report_files()`)
- [ ] Add data persistence to IndexedDB (save state between sessions)
- [ ] Port remaining validation functions
- [ ] Add category import/export functionality
- [ ] Support for multiple projects simultaneously

## AWS Deployment Options

### Option 1: S3 + CloudFront (Static Hosting)

1. Upload all files to an S3 bucket
2. Enable static website hosting
3. Create CloudFront distribution pointing to S3
4. Update manifest.xml with CloudFront URL
5. Ensure CloudFront uses HTTPS

### Option 2: AWS Amplify

1. Connect your repository to AWS Amplify
2. Point to `TaskPaneMod` folder
3. Amplify automatically provides HTTPS
4. Update manifest.xml with Amplify URL

### Option 3: EC2/Elastic Beanstalk

1. Deploy files to web server (Apache/Nginx)
2. Configure SSL certificate
3. Update manifest.xml with domain

## Troubleshooting

**Add-in doesn't load:**
- **Must use HTTPS** (not HTTP) - Word Add-ins require secure connections
- Check browser console for errors (F12 in Word)
- Verify manifest.xml URLs match your AWS domain exactly
- Ensure CORS headers allow Office.js to load resources
- Check that all CDN resources (Pyodide, DuckDB) are accessible

**Pyodide fails to load:**
- Check internet connection (Pyodide loads from CDN)
- Try a different Pyodide version in `taskpane.js`

**DuckDB errors:**
- Check browser console for SQL errors
- Verify table schema matches

## Notes

- This is a proof of concept - production use would need additional features
- Categories are currently hardcoded - should load from CSV/API
- Database is in-memory - data is lost on page refresh (can add IndexedDB persistence)
- Report generation is simplified - full version would match original HTML output

