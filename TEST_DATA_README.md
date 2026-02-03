# Test Categories Data

## File: `test-categories.json`

This file contains a realistic test dataset based on the structure found in the legacy codebase.

## What's Included

### Categories by Project

**DIANA Project:**
- Political Domain (with 3 levels of subcategories)
- Economic Domain (with 4 levels including "Greatest Negative Influence > NATO/EU")
- Social Domain (with 3 levels)
- Security Domain (with 3 levels)
- Media Domain (with 3 levels)
- Country-specific categories:
  - Armenia-Specific Issues (Nagorno-Karabakh conflict)
  - Ukraine-Specific Issues (War impact, Refugee crisis)

**JOHNSON Project:**
- Key Influencers > Economic > Greatest Negative Influence > NATO
- Media Consumption (Traditional media, Social media)

**JOHNSON2 Project:**
- Key Influencers (Economic, Political)
- Media Consumption

**RIDE Project:**
- Main Themes > Recruitment (methods, channels)
- Main Themes > Member Experience (current, former)

### Category Strings

Flattened category paths (e.g., "Political Domain > Political Indicators > Political division...") for all categories.

### Sample Coded Data

3 sample coded text entries to demonstrate the data structure.

## How to Use

1. **Open Word Add-in**
2. **Click "Import Categories"**
3. **Select `test-categories.json`**
4. **Select a project** from the dropdown (DIANA, JOHNSON, JOHNSON2, RIDE)
5. **See categories appear** in the tree

## Category Structure

The categories follow the same hierarchical structure as the legacy system:
- Up to 4 levels deep
- Parent-child relationships via `parent_id`
- `display_order` for sorting
- Country-specific categories for DIANA project
- Project-specific categories

## Testing Scenarios

1. **Test hierarchical display**: Select DIANA project, see nested categories
2. **Test project filtering**: Switch between projects, see different categories
3. **Test coding**: Select text, choose a category, code it
4. **Test country-specific**: For DIANA, country-specific categories appear separately
5. **Test deep nesting**: Some categories go 4 levels deep (e.g., Economic > Indicators > Greatest Negative > NATO)

## Notes

- IDs are sequential but don't match database IDs (this is test data)
- Category names are based on examples found in the codebase
- Structure matches the CSV import format from `create-categories.py`
- When you get the real CSV/data, you can replace this file

