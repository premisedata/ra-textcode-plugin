# Category Management in Legacy System

## How Categories Work

### Categories are **FIXED/PREDEFINED** - Users Cannot Add Categories

Categories come from **CSV files** that are imported into the database by administrators.

## Category Source: CSV Files

Categories are imported from CSV files using `helpers/create-categories.py`:

**Example CSV files:**
- `DIANA_Code List_20250226.csv`
- `DIANA_Code List_20250325.csv`
- `JOHNSON-Qual Coding Tool Framework.csv`
- `RIDE-categories.csv`
- `JOHNSON_PH2_Coding-Tool-Framework_V1.1-CS.csv`

**CSV Format:**
```
Political Domain, , , 
, Political Indicators of Vulnerability to FMI, , 
, , Political division is a serious problem, 
```

- Column 0 = Level 1 category (e.g., "Political Domain")
- Column 1 = Level 2 category (e.g., "Political Indicators...")
- Column 2 = Level 3 category (e.g., "Political division...")
- Column 3 = Level 4 category (if any)

## How Categories are Loaded

1. **Admin runs `create-categories.py`**:
   - Script reads CSV file
   - **Deletes all existing categories** for that project
   - Inserts new categories from CSV
   - Creates hierarchical structure (parent_id relationships)
   - Generates `category_str` table (flattened paths like "Political > Indicators > ...")

2. **Word macro queries database**:
   - `getcat1` - Gets top-level categories (parent_id is null)
   - `getcat2` - Gets child categories (by parent_id)
   - Categories are **read-only** - no way to add/edit from Word

## Category Structure

**Database Schema:**
- `rna.categories` - Hierarchical categories
  - `id` - Primary key
  - `name` - Category name
  - `country` - Country-specific categories (e.g., "Armenia-Specific Issues")
  - `parent_id` - Parent category ID (null for top-level)
  - `project` - Project name (DIANA, JOHNSON, RIDE, JOHNSON2)
  - `display_order` - Order for display

- `rna.category_str` - Flattened category paths
  - `category_str` - Full path like "Political Domain > Indicators > ..."
  - `project` - Project name
  - `display_order` - Order for display

## User Experience

**Users can:**
- ✅ Select from predefined categories
- ✅ Code text with existing categories
- ✅ See hierarchical category tree

**Users CANNOT:**
- ❌ Add new categories
- ❌ Edit category names
- ❌ Delete categories
- ❌ Reorder categories

## Category Updates

Categories are updated by:
1. **Admin creates/updates CSV file** with new category structure
2. **Admin runs `create-categories.py`** to import
3. **All existing categories for that project are deleted** and replaced
4. **Word users see new categories** on next use

## Country-Specific Categories

For DIANA project, categories can be country-specific:
- Format: `"Armenia-Specific Issues"` in CSV
- These appear in addition to general categories
- Only shown for that specific country

## Implications for Word Add-in

Since categories are **fixed and managed centrally**, the Word Add-in should:

1. **Load categories from database export** (via import function)
2. **Display categories in tree** (read-only)
3. **Allow coding with existing categories**
4. **NOT allow users to add/edit categories** (matches legacy behavior)

If you want to allow users to add categories, that would be a **new feature** not in the legacy system.

