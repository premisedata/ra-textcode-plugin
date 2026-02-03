#!/usr/bin/env python3
"""
Export data from PostgreSQL database to JSON for import into Word Add-in
Usage: python3 export-db-data.py [project] [output_file.json]
"""

import psycopg2
from psycopg2.extras import RealDictCursor
import json
import sys
import os

# Database connection (update with your credentials)
DB_CONFIG = {
    'host': '10.0.4.134',
    'port': '5432',
    'dbname': 'msi',
    'user': '',  # Add your username
    'password': ''  # Add your password
}

def export_categories(project=None):
    """Export categories from database"""
    conn = psycopg2.connect(**DB_CONFIG)
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    
    if project:
        query = """
            SELECT id, name, country, parent_id, project, display_order
            FROM rna.categories
            WHERE project = %s
            ORDER BY display_order, name
        """
        cursor.execute(query, (project,))
    else:
        query = """
            SELECT id, name, country, parent_id, project, display_order
            FROM rna.categories
            ORDER BY project, display_order, name
        """
        cursor.execute(query)
    
    categories = cursor.fetchall()
    return [dict(cat) for cat in categories]

def export_category_strings(project=None):
    """Export category_str (flattened category paths)"""
    conn = psycopg2.connect(**DB_CONFIG)
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    
    if project:
        query = """
            SELECT category_str, project, display_order
            FROM rna.category_str
            WHERE project = %s
            ORDER BY display_order, category_str
        """
        cursor.execute(query, (project,))
    else:
        query = """
            SELECT category_str, project, display_order
            FROM rna.category_str
            ORDER BY project, display_order, category_str
        """
        cursor.execute(query)
    
    category_strings = cursor.fetchall()
    return [dict(cs) for cs in category_strings]

def export_sample_coded_data(project=None, limit=100):
    """Export sample coded data (for testing)"""
    conn = psycopg2.connect(**DB_CONFIG)
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    
    if project:
        query = """
            SELECT id, tagged_text, category_str, dtg, username, doc_filename, 
                   participant_number, user_comment, citation
            FROM rna.c0d3t3xt
            WHERE doc_filename LIKE %s
            ORDER BY dtg DESC
            LIMIT %s
        """
        cursor.execute(query, (f'%{project}%', limit))
    else:
        query = """
            SELECT id, tagged_text, category_str, dtg, username, doc_filename, 
                   participant_number, user_comment, citation
            FROM rna.c0d3t3xt
            ORDER BY dtg DESC
            LIMIT %s
        """
        cursor.execute(query, (limit,))
    
    coded_data = cursor.fetchall()
    # Convert datetime to string for JSON
    result = []
    for row in coded_data:
        r = dict(row)
        if r.get('dtg'):
            r['dtg'] = r['dtg'].isoformat()
        result.append(r)
    return result

def main():
    project = sys.argv[1] if len(sys.argv) > 1 else None
    output_file = sys.argv[2] if len(sys.argv) > 2 else 'db-export.json'
    
    print(f"Exporting data from database...")
    if project:
        print(f"Project filter: {project}")
    
    data = {
        'categories': export_categories(project),
        'category_strings': export_category_strings(project),
        'sample_coded_data': export_sample_coded_data(project, limit=50),
        'export_info': {
            'project': project,
            'timestamp': None  # Will be set by json.dumps
        }
    }
    
    # Add timestamp
    import datetime
    data['export_info']['timestamp'] = datetime.datetime.now().isoformat()
    
    # Write to JSON file
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    
    print(f"\nExport complete!")
    print(f"Categories: {len(data['categories'])}")
    print(f"Category strings: {len(data['category_strings'])}")
    print(f"Sample coded data: {len(data['sample_coded_data'])}")
    print(f"\nSaved to: {output_file}")
    print(f"\nTo import into Word Add-in:")
    print(f"  1. Open the add-in")
    print(f"  2. Use 'Import Categories' button")
    print(f"  3. Select {output_file}")

if __name__ == '__main__':
    main()

