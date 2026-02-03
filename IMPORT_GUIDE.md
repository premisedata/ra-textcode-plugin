# How to Import Database Data into Word Add-in

## Step 1: Export Data from PostgreSQL Database

### Option A: Use the Export Script

1. **Update database credentials** in `export-db-data.py`:
   ```python
   DB_CONFIG = {
       'host': '10.0.4.134',
       'port': '5432',
       'dbname': 'msi',
       'user': 'your_username',  # Add your username
       'password': 'your_password'  # Add your password
   }
   ```

2. **Run the export script**:
   ```bash
   cd TaskPaneMod
   python3 export-db-data.py DIANA db-export.json
   ```
   
   Or export all projects:
   ```bash
   python3 export-db-data.py
   ```

3. **The script will create** `db-export.json` with:
   - All categories for the project
   - Category strings (flattened paths)
   - Sample coded data (for reference)

### Option B: Manual SQL Export

If you have direct database access, you can run:

```sql
-- Export categories
SELECT id, name, country, parent_id, project, display_order
FROM rna.categories
WHERE project = 'DIANA'
ORDER BY display_order, name;

-- Export category strings
SELECT category_str, project, display_order
FROM rna.category_str
WHERE project = 'DIANA'
ORDER BY display_order, category_str;
```

Save results as JSON with this structure:
```json
{
  "categories": [
    {
      "id": 1,
      "name": "Political Domain",
      "country": null,
      "parent_id": null,
      "project": "DIANA",
      "display_order": 1
    }
  ],
  "category_strings": [
    {
      "category_str": "Political Domain > Political Indicators",
      "project": "DIANA",
      "display_order": 1
    }
  ]
}
```

## Step 2: Import into Word Add-in

1. **Open the Word Add-in** in Word
2. **Click "Import Categories"** button
3. **Select the JSON file** you exported
4. **Categories will be loaded** and appear in the category tree

## Step 3: Verify Import

1. **Select a project** from the dropdown (DIANA, JOHNSON, etc.)
2. **Check the category tree** - you should see all imported categories
3. **Try coding some text** to verify it works

## File Format

The JSON file should have this structure:

```json
{
  "categories": [
    {
      "id": 1,
      "name": "Category Name",
      "country": "armenia" or null,
      "parent_id": null or parent_id,
      "project": "DIANA",
      "display_order": 1
    }
  ],
  "category_strings": [
    {
      "category_str": "Parent > Child > Grandchild",
      "project": "DIANA",
      "display_order": 1
    }
  ],
  "sample_coded_data": [
    {
      "id": 1,
      "tagged_text": "Sample text",
      "category_str": "Category > Subcategory",
      "doc_filename": "DIANA_Ukraine_FG1.docx",
      "participant_number": 1,
      "user_comment": "Optional comment"
    }
  ]
}
```

## Troubleshooting

**Import fails:**
- Check JSON file is valid (use a JSON validator)
- Ensure "categories" array exists
- Check console for specific error messages

**Categories don't appear:**
- Select the correct project from dropdown
- Check that categories have matching project value
- Try refreshing the add-in

**Missing parent categories:**
- Ensure parent categories are imported before children
- Check parent_id values match actual category IDs

## Next Steps

Once categories are imported:
1. Test coding text with different categories
2. Generate reports to verify data structure
3. Export coded data to verify it matches original format

