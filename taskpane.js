/* global Office, duckdb */

let pyodide = null;
let db = null;
let conn = null;
let selectedText = '';
let selectedCategory = '';
let selectedCategoryColor = '#FFFF00'; // Default yellow
let currentDocFilename = '';

// Initialize when Office.js is ready
Office.onReady((info) => {
    console.log('[OFFICE] Office.js ready, host:', info.host);
    if (info.host === Office.HostType.Word) {
        console.log('[OFFICE] Starting initialization...');
        initializeApp();
    } else {
        showStatus('This add-in only works in Microsoft Word.', 'error');
    }
}).catch((error) => {
    console.error('[OFFICE] Office.onReady failed:', error);
    showStatus('Office.js initialization failed', 'error');
    document.getElementById('main-content').style.display = 'block';
});

async function initializeApp() {
    console.log('[INIT] initializeApp called');
    
    // Show UI immediately - don't wait for anything
    console.log('[INIT] Showing UI...');
    document.getElementById('main-content').style.display = 'block';
    setupEventHandlers();
    showStatus('Initializing...', 'info');
    console.log('[INIT] UI shown');
    
    // Skip Pyodide and DuckDB entirely for now - just get IndexedDB working
    try {
        console.log('[INIT] Initializing IndexedDB...');
        showStatus('Initializing storage...', 'info');
        await initializeIndexedDB();
        console.log('[INIT] IndexedDB ready');
        
        showStatus('Ready!', 'success');
        console.log('[INIT] Loading categories...');
        await loadCategories();
        console.log('[INIT] Complete');
    } catch (error) {
        showStatus(`Error: ${error.message}`, 'error');
        console.error('[INIT] Error:', error);
    }
}

async function loadPythonFunctions() {
    if (!pyodide) {
        throw new Error('Pyodide not initialized');
    }
    
    // Load Python code that replicates db-api.py functions
    const pythonCode = `
import re
import json
from datetime import datetime

def parseParticipantNumber(tagged_text, paragraph):
    ret = None
    match = re.search(r"([P_]\\d+):", tagged_text)
    if match:
        ret = match.group(1).replace('_','P')
    else:
        match = re.search(r"([P_]\\d+):", paragraph)
        if match:
            ret = match.group(1).replace('_','P')
    return ret

def getProjectName(doc_filename):
    ret = ""
    if "_FG" in doc_filename:
        ret = "DIANA"
    if "_LVA_" in doc_filename:
        ret = "JOHNSON"
    if "RIDE_" in doc_filename:
        ret = "RIDE"
    if "JOHNSON2_" in doc_filename:
        ret = "JOHNSON2"
    return ret

def parseCountryName(doc_filename):
    ret = '<NO-COUNTRY-FOUND>'
    a = doc_filename.split('_FG')
    if len(a) >= 2:
        b = a[0].split('_')
        if len(b) == 2:
            ret = b[1]
    return ret

def cleanCategory(c):
    ret = c.replace(' -- display_order=None','')
    ret = re.sub(r" -- display_order=\\d+", "", ret)
    return ret

def createCitationJOHNSON(citation):
    gender = ""
    age = ""
    occupation = ""
    city = ""
    res = re.search(r"<gender>(.*?)</gender>", citation)
    if res:
        gender = res.group(1)
    res = re.search(r"<age>(.*?)</age>", citation)
    if res:
        age = res.group(1)
    res = re.search(r"<occupation>(.*?)</occupation>", citation)
    if res:
        occupation = res.group(1)
    res = re.search(r"<city>(.*?)</city>", citation)
    if res:
        city = res.group(1)
    return f"({gender}, {age}, {occupation}, {city})"

def createCitationJOHNSON2(citation, participant_number):
    gender = ""
    age = ""
    occupation = ""
    city = ""
    res = re.search(r"<gender>(.*?)</gender>", citation)
    if res:
        gender = res.group(1)
    res = re.search(r"<age>(.*?)</age>", citation)
    if res:
        age = res.group(1)
    res = re.search(r"<occupation>(.*?)</occupation>", citation)
    if res:
        occupation = res.group(1)
    res = re.search(r"<city>(.*?)</city>", citation)
    if res:
        city = res.group(1)
    return f"({gender}, {age}, {occupation}, {city})"

def createCitationRIDE(citation, doc_filename):
    gender = ""
    age = ""
    occupation = ""
    city = ""
    res = re.search(r"<gender>(.*?)</gender>", citation)
    if res:
        gender = res.group(1)
    res = re.search(r"<age>(.*?)</age>", citation)
    if res:
        age = res.group(1)
    
    occupation = 'Unk'
    if 'former_member' in doc_filename.lower():
        occupation = 'Former Member'
    elif 'recruited_non_member' in doc_filename.lower():
        occupation = 'Recruited Non-Member'
    elif 'current_member' in doc_filename.lower():
        occupation = 'Current Member'
    
    res = re.search(r"<city>(.*?)</city>", citation)
    if res:
        city = res.group(1)
    
    return f"({gender}, {age}, {occupation}, {city})"
`;
    
    await pyodide.runPython(pythonCode);
}

async function waitForDuckDB() {
    console.log('[DuckDB] Waiting for script to load...');
    let attempts = 0;
    const maxAttempts = 50;
    
    while (typeof duckdb === 'undefined' && attempts < maxAttempts) {
        await new Promise(resolve => setTimeout(resolve, 100));
        attempts++;
        if (attempts % 10 === 0) {
            console.log(`[DuckDB] Still waiting... (${attempts}/${maxAttempts})`);
        }
    }
    
    if (typeof duckdb === 'undefined') {
        console.error('[DuckDB] Script never loaded');
        throw new Error('DuckDB failed to load from CDN');
    }
    console.log('[DuckDB] Script loaded');
}

async function initializeDuckDB() {
    console.log('[DuckDB] Getting module reference...');
    const DuckDBModule = typeof duckdb !== 'undefined' ? duckdb : 
                         typeof window.duckdb !== 'undefined' ? window.duckdb :
                         typeof window.DuckDB !== 'undefined' ? window.DuckDB :
                         null;
    
    if (!DuckDBModule) {
        throw new Error('DuckDB WASM is not available.');
    }
    console.log('[DuckDB] Module found, initializing...');
    
    // Simple initialization - try the most common pattern first
    if (typeof DuckDBModule.createAsyncDuckDB === 'function') {
        console.log('[DuckDB] Using createAsyncDuckDB...');
        db = await DuckDBModule.createAsyncDuckDB();
        console.log('[DuckDB] Database created, connecting...');
        conn = await db.connect();
        console.log('[DuckDB] Connected');
    } else if (typeof DuckDBModule.Database !== 'undefined') {
        console.log('[DuckDB] Using sync Database API...');
        db = new DuckDBModule.Database(':memory:');
        conn = db.connect();
        console.log('[DuckDB] Connected (sync)');
    } else {
        console.error('[DuckDB] Unknown API, available keys:', Object.keys(DuckDBModule));
        throw new Error('Unsupported DuckDB API');
    }
    
    // Helper to execute queries (handles both sync and async)
    const execQuery = async (sql) => {
        console.log('[DuckDB] Executing:', sql.substring(0, 50) + '...');
        const result = conn.query(sql);
        const finalResult = (result && typeof result.then === 'function') ? await result : result;
        console.log('[DuckDB] Query completed');
        return finalResult;
    };
    
    // Create tables - ensure each completes before next
    console.log('[DuckDB] Creating tables...');
    await execQuery(`CREATE TABLE IF NOT EXISTS c0d3t3xt (
        id INTEGER PRIMARY KEY,
        tagged_text TEXT,
        category_str TEXT,
        dtg TIMESTAMP,
        username TEXT,
        doc_filename TEXT,
        participant_number INTEGER,
        user_comment TEXT,
        citation TEXT
    )`);
    
    await execQuery(`CREATE TABLE IF NOT EXISTS participants (
        number TEXT,
        education TEXT,
        employment TEXT,
        ethnicity TEXT,
        age TEXT,
        gender TEXT,
        location TEXT,
        doc_number_unique_id TEXT PRIMARY KEY
    )`);
    
    await execQuery(`CREATE TABLE IF NOT EXISTS categories (
        id INTEGER PRIMARY KEY,
        name TEXT,
        country TEXT,
        parent_id INTEGER,
        project TEXT,
        display_order INTEGER,
        color TEXT
    )`);
    
    await execQuery(`CREATE TABLE IF NOT EXISTS category_str (
        category_str TEXT PRIMARY KEY,
        project TEXT,
        display_order INTEGER
    )`);
    console.log('[DuckDB] All tables created');
    
    // Create simple wrapper for consistent async interface
    const originalConn = conn;
    conn = {
        type: 'duckdb',
        db: db,
        conn: originalConn,
        query: async (sql, params) => {
            const result = params ? originalConn.query(sql, params) : originalConn.query(sql);
            return (result && typeof result.then === 'function') ? await result : result;
        }
    };
    
    // Load sample categories
    console.log('[DuckDB] Loading sample categories...');
    await loadSampleCategories();
    console.log('[DuckDB] Sample categories loaded');
}

// Fallback: Use IndexedDB instead of DuckDB
async function initializeIndexedDB() {
    // Create a simple IndexedDB-based storage
    return new Promise((resolve, reject) => {
        const request = indexedDB.open('codetext', 1);
        
        request.onerror = () => {
            console.error('IndexedDB failed, using in-memory storage');
            // Ultimate fallback: in-memory array
            initializeInMemoryStorage();
            resolve();
        };
        
        request.onsuccess = () => {
            db = request.result;
            conn = {
                type: 'indexeddb',
                db: db,
                query: async (sql, params) => {
                    return await queryIndexedDB(sql, params);
                }
            };
            console.log('IndexedDB initialized');
            loadSampleCategories();
            resolve();
        };
        
        request.onupgradeneeded = (event) => {
            const db = event.target.result;
            
            // Create object stores
            if (!db.objectStoreNames.contains('coded_text')) {
                db.createObjectStore('coded_text', { keyPath: 'id', autoIncrement: true });
            }
            if (!db.objectStoreNames.contains('participants')) {
                db.createObjectStore('participants', { keyPath: 'doc_number_unique_id' });
            }
            if (!db.objectStoreNames.contains('categories')) {
                db.createObjectStore('categories', { keyPath: 'id', autoIncrement: true });
            }
            if (!db.objectStoreNames.contains('category_str')) {
                db.createObjectStore('category_str', { keyPath: 'category_str' });
            }
        };
    });
}

// Ultimate fallback: in-memory storage
function initializeInMemoryStorage() {
    console.log('Using in-memory storage (no persistence)');
    db = {
        type: 'memory',
        data: {
            c0d3t3xt: [],
            participants: [],
            categories: [],
            category_str: []
        },
        nextId: 1
    };
    
    conn = {
        type: 'memory',
        query: (sql, params) => {
            return queryInMemory(sql, params);
        }
    };
    
    loadSampleCategories();
}

// Simple query interface for in-memory storage
function queryInMemory(sql, params) {
    const upperSQL = sql.toUpperCase().trim();
    
    // INSERT INTO c0d3t3xt
    if (upperSQL.includes('INSERT INTO C0D3T3XT')) {
        const item = {
            id: db.nextId++,
            tagged_text: params[0],
            category_str: params[1],
            dtg: new Date().toISOString(),
            username: params[2] || 'user',
            doc_filename: params[3],
            participant_number: params[4],
            user_comment: params[5] || '',
            citation: params[6] || ''
        };
        db.data.c0d3t3xt.push(item);
        return [item];
    }
    
    // SELECT FROM c0d3t3xt
    if (upperSQL.includes('SELECT') && upperSQL.includes('FROM C0D3T3XT')) {
        return db.data.c0d3t3xt;
    }
    
    // SELECT FROM categories
    if (upperSQL.includes('SELECT') && upperSQL.includes('FROM CATEGORIES')) {
        return db.data.categories;
    }
    
    // DELETE FROM
    if (upperSQL.includes('DELETE FROM C0D3T3XT')) {
        db.data.c0d3t3xt = [];
        return [];
    }
    
    // INSERT INTO categories
    if (upperSQL.includes('INSERT') && upperSQL.includes('INTO CATEGORIES')) {
        // Ensure parent_id is a number or null
        const parentId = params[2] !== null && params[2] !== undefined ? Number(params[2]) : null;
        const item = {
            id: db.nextId++,
            name: params[0],
            country: params[1] || null,
            parent_id: parentId,
            project: params[3],
            display_order: params[4] || 0,
            color: params[5] || null // Add color support
        };
        db.data.categories.push(item);
        return [item];
    }
    
    return [];
}

// Query interface for IndexedDB
async function queryIndexedDB(sql, params) {
    const upperSQL = sql.toUpperCase().trim();
    const transaction = db.transaction(['coded_text', 'participants', 'categories', 'category_str'], 'readwrite');
    
    // INSERT INTO c0d3t3xt
    if (upperSQL.includes('INSERT INTO C0D3T3XT')) {
        const store = transaction.objectStore('coded_text');
        const item = {
            tagged_text: params[0],
            category_str: params[1],
            dtg: new Date().toISOString(),
            username: params[2] || 'user',
            doc_filename: params[3],
            participant_number: params[4],
            user_comment: params[5] || '',
            citation: params[6] || ''
        };
        return new Promise((resolve) => {
            const request = store.add(item);
            request.onsuccess = () => {
                item.id = request.result;
                resolve([item]);
            };
        });
    }
    
    // SELECT FROM c0d3t3xt
    if (upperSQL.includes('SELECT') && upperSQL.includes('FROM C0D3T3XT')) {
        const store = transaction.objectStore('coded_text');
        return new Promise((resolve) => {
            const request = store.getAll();
            request.onsuccess = () => resolve(request.result);
        });
    }
    
    // SELECT FROM categories
    if (upperSQL.includes('SELECT') && upperSQL.includes('FROM CATEGORIES')) {
        const store = transaction.objectStore('categories');
        return new Promise((resolve) => {
            const request = store.getAll();
            request.onsuccess = () => resolve(request.result);
        });
    }
    
    // DELETE FROM
    if (upperSQL.includes('DELETE FROM C0D3T3XT')) {
        const store = transaction.objectStore('coded_text');
        return new Promise((resolve) => {
            const request = store.clear();
            request.onsuccess = () => resolve([]);
        });
    }
    
    // INSERT INTO categories
    if (upperSQL.includes('INSERT') && upperSQL.includes('INTO CATEGORIES')) {
        const store = transaction.objectStore('categories');
        // Ensure parent_id is a number or null
        const parentId = params[2] !== null && params[2] !== undefined ? Number(params[2]) : null;
        const item = {
            name: params[0],
            country: params[1] || null,
            parent_id: parentId,
            project: params[3],
            display_order: params[4] || 0,
            color: params[5] || null // Add color support
        };
        return new Promise((resolve) => {
            const request = store.add(item);
            request.onsuccess = () => {
                item.id = request.result;
                resolve([item]);
            };
        });
    }
    
    return [];
}

async function loadSampleCategories() {
    // Sample categories - in production, load from database export
    // These are just placeholders until you import real categories
    const sampleCategories = [
        { name: 'Political Domain', country: null, parent_id: null, project: 'DIANA', display_order: 1, color: '#FF6B6B' },
        { name: 'Economic Domain', country: null, parent_id: null, project: 'DIANA', display_order: 2, color: '#4ECDC4' },
        { name: 'Social Domain', country: null, parent_id: null, project: 'DIANA', display_order: 3, color: '#95E1D3' },
        { name: 'Security Domain', country: null, parent_id: null, project: 'DIANA', display_order: 4, color: '#F38181' },
        { name: 'Media Domain', country: null, parent_id: null, project: 'DIANA', display_order: 5, color: '#AA96DA' },
    ];
    
    for (const cat of sampleCategories) {
        // Check if category already exists (for IndexedDB/memory)
        let existing = [];
        if (conn.type === 'indexeddb' || conn.type === 'memory') {
            let allCats = conn.query('SELECT * FROM categories');
            if (allCats && typeof allCats.then === 'function') {
                allCats = await allCats;
            }
            existing = allCats.filter(c => c.name === cat.name && c.project === cat.project);
        }
        
        if (existing.length === 0) {
            let result = conn.query(`
                INSERT OR IGNORE INTO categories (name, country, parent_id, project, display_order, color)
                VALUES (?, ?, ?, ?, ?, ?)
            `, [cat.name, cat.country || null, cat.parent_id || null, cat.project, cat.display_order || 0, cat.color || null]);
            if (result && typeof result.then === 'function') {
                await result;
            }
        }
    }
}

function setupEventHandlers() {
    document.getElementById('get-selection-btn').addEventListener('click', getSelectedText);
    document.getElementById('code-text-btn').addEventListener('click', codeSelectedText);
    document.getElementById('generate-report-btn').addEventListener('click', generateReport);
    document.getElementById('export-data-btn').addEventListener('click', exportData);
    document.getElementById('clear-db-btn').addEventListener('click', clearDatabase);
    document.getElementById('import-categories-btn').addEventListener('click', importCategories);
    document.getElementById('category-search').addEventListener('input', filterCategories);
    document.getElementById('project-select').addEventListener('change', async () => {
        await loadCategories();
    });
}

async function getSelectedText() {
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load("text");
            
            const doc = context.document;
            doc.load("name");
            
            await context.sync();
            
            selectedText = selection.text.trim();
            // Get document name - use filename or fallback
            currentDocFilename = doc.name || 'Untitled.docx';
            
            // If document is unsaved, try to get a better name
            if (!currentDocFilename || currentDocFilename === '') {
                try {
                    // Try to get the document title or use a default
                    currentDocFilename = 'Untitled.docx';
                } catch (e) {
                    currentDocFilename = 'Untitled.docx';
                }
            }
            
            document.getElementById('selected-text-display').textContent = 
                selectedText || '(No text selected)';
            
            if (selectedText) {
                document.getElementById('code-text-btn').disabled = false;
                showStatus(`Selected text from: ${currentDocFilename}`, 'success');
            } else {
                document.getElementById('code-text-btn').disabled = true;
                showStatus('Please select text in the document', 'info');
            }
        });
    } catch (error) {
        showStatus(`Error getting selection: ${error.message}`, 'error');
        console.error('Error in getSelectedText:', error);
        // Set fallback filename
        currentDocFilename = 'Untitled.docx';
    }
}

async function codeSelectedText() {
    if (!selectedText || !selectedCategory) {
        showStatus('Please select text and a category', 'error');
        return;
    }
    
    try {
        // Ensure we have document filename
        if (!currentDocFilename) {
            // Try to get it again
            await Word.run(async (context) => {
                const doc = context.document;
                doc.load("name");
                await context.sync();
                currentDocFilename = doc.name || 'Unknown.docx';
            });
        }
        
        // Escape quotes for Python
        const escapedText = selectedText.replace(/"/g, '\\"').replace(/'/g, "\\'");
        const escapedFilename = (currentDocFilename || 'Unknown.docx').replace(/"/g, '\\"').replace(/'/g, "\\'");
        
        // Get project name using Python
        const project = pyodide.runPython(`getProjectName("${escapedFilename}")`);
        
        // Parse participant number using Python
        const participantNumberResult = pyodide.runPython(
            `parseParticipantNumber("${escapedText}", "")`
        );
        const participantNumber = participantNumberResult === 'None' || participantNumberResult === null 
            ? null 
            : parseInt(participantNumberResult) || null;
        
        const userComment = document.getElementById('user-comment').value || '';
        
        // Insert into database
        let queryResult = conn.query(`
            INSERT INTO c0d3t3xt 
            (tagged_text, category_str, dtg, username, doc_filename, participant_number, user_comment)
            VALUES (?, datetime('now'), 'user', ?, ?, ?)
        `, [selectedText, selectedCategory, currentDocFilename || 'Unknown.docx', participantNumber, userComment]);
        
        // Handle async queries
        if (queryResult && typeof queryResult.then === 'function') {
            await queryResult;
        }
        
        // Highlight text in Word using category color
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            // Word expects hex color format (e.g., "#FFFF00")
            // Ensure color is in correct format
            let highlightColor = selectedCategoryColor;
            if (!highlightColor.startsWith('#')) {
                highlightColor = '#' + highlightColor;
            }
            // If color is invalid, default to yellow
            if (!/^#[0-9A-Fa-f]{6}$/.test(highlightColor)) {
                highlightColor = '#FFFF00';
            }
            selection.font.highlightColor = highlightColor;
            await context.sync();
        });
        
        showStatus(`Coded text with category: ${selectedCategory}`, 'success');
        
        // Clear selection
        selectedText = '';
        selectedCategory = '';
        selectedCategoryColor = '#FFFF00'; // Reset to default
        document.getElementById('selected-text-display').textContent = '';
        document.getElementById('user-comment').value = '';
        document.getElementById('code-text-btn').disabled = true;
        
        // Clear category selection
        document.querySelectorAll('.category-item').forEach(item => {
            item.classList.remove('selected');
        });
        
        // Update output
        let countResult = conn.query('SELECT COUNT(*) as count FROM c0d3t3xt');
        if (countResult && typeof countResult.then === 'function') {
            countResult = await countResult;
        }
        const count = Array.isArray(countResult) ? countResult.length : (countResult[0]?.count || 0);
        document.getElementById('output').textContent = `Total coded items: ${count}`;
        
    } catch (error) {
        showStatus(`Error coding text: ${error.message}`, 'error');
        console.error(error);
    }
}

async function loadCategories() {
    try {
        const project = document.getElementById('project-select').value;
        
        // Handle both sync and async query returns
        // Build query based on storage type
        let categories;
        if (conn.type === 'indexeddb' || conn.type === 'memory') {
            // For IndexedDB/memory, get all and filter
            let allCats = conn.query('SELECT id, name, parent_id, display_order, project, color FROM categories');
            if (allCats && typeof allCats.then === 'function') {
                allCats = await allCats;
            }
            categories = allCats.filter(cat => cat.project === project);
            // Sort
            categories.sort((a, b) => {
                if (a.display_order !== b.display_order) {
                    return (a.display_order || 0) - (b.display_order || 0);
                }
                return (a.name || '').localeCompare(b.name || '');
            });
        } else {
            // For DuckDB, use SQL with parameter
            categories = conn.query(`
                SELECT id, name, parent_id, display_order, project, color
                FROM categories
                WHERE project = ?
                ORDER BY display_order, name
            `, [project]);
        }
        
        // If it's a promise, await it
        if (categories && typeof categories.then === 'function') {
            categories = await categories;
        }
        
        // Debug: Log what we got from database
        const debugInfo = document.getElementById('debug-info');
        const debugText = document.getElementById('debug-text');
        
        let debugMsg = `Loaded ${categories.length} categories. `;
        
        if (categories.length > 0) {
            // Show sample categories
            const sample = categories.slice(0, 3).map(c => 
                `${c.name} (id=${c.id}, parent_id=${c.parent_id || 'null'})`
            ).join('; ');
            debugMsg += `Sample: ${sample}. `;
            
            // Show categories with parent_id
            const withParents = categories.filter(c => {
                const pid = c.parent_id;
                return pid != null && pid !== '' && pid !== 'null' && pid !== 'undefined';
            });
            debugMsg += `Categories with parent_id: ${withParents.length}. `;
            
            if (withParents.length > 0) {
                const sampleWithParents = withParents.slice(0, 3).map(c => 
                    `${c.name} (parent_id=${c.parent_id})`
                ).join('; ');
                debugMsg += `Examples: ${sampleWithParents}. `;
            } else {
                debugMsg += `⚠️ NO categories have parent_id set! `;
            }
        }
        
        debugText.textContent = debugMsg;
        debugInfo.style.display = 'block';
        
        console.log(`Loaded ${categories.length} categories for ${project}`);
        if (categories.length > 0) {
            console.log('Sample categories:', categories.slice(0, 5).map(c => ({
                id: c.id,
                name: c.name,
                parent_id: c.parent_id,
                idType: typeof c.id,
                parentIdType: typeof c.parent_id,
                project: c.project
            })));
            
            // Show categories with parent_id
            const withParents = categories.filter(c => {
                const pid = c.parent_id;
                return pid != null && pid !== '' && pid !== 'null' && pid !== 'undefined';
            });
            console.log(`Categories with parent_id: ${withParents.length}`);
            if (withParents.length > 0) {
                console.log('Categories with parents:', withParents.slice(0, 5).map(c => 
                    `${c.name} (id=${c.id}, parent_id=${c.parent_id})`
                ));
            } else {
                console.error('⚠️ NO categories have parent_id! This means the import didn\'t save parent_id correctly.');
            }
        }
        
        // If no categories for this project, show message
        if (!categories || categories.length === 0) {
            const treeDiv = document.getElementById('category-tree');
            treeDiv.innerHTML = '<p style="padding: 20px; text-align: center; color: #666;">No categories loaded. Click "Import Categories" to load from database export.</p>';
            return;
        }
        
        const treeDiv = document.getElementById('category-tree');
        treeDiv.innerHTML = '';
        
        // Debug: Log categories to see what we have
        console.log('Loaded categories:', categories.length);
        console.log('Sample categories:', categories.slice(0, 5));
        
        // Build category tree - SIMPLIFIED APPROACH
        // Create objects with children arrays
        const categoryMap = new Map();
        
        // First pass: create all category objects with empty children arrays
        categories.forEach(cat => {
            // Normalize IDs - handle both string and number, and handle null/undefined
            let catId = null;
            if (cat.id != null && cat.id !== '') {
                catId = typeof cat.id === 'string' ? parseInt(cat.id, 10) : Number(cat.id);
                if (isNaN(catId)) catId = null;
            }
            
            let parentId = null;
            if (cat.parent_id != null && cat.parent_id !== '' && cat.parent_id !== 'null' && cat.parent_id !== 'undefined') {
                parentId = typeof cat.parent_id === 'string' ? parseInt(cat.parent_id, 10) : Number(cat.parent_id);
                if (isNaN(parentId)) parentId = null;
            }
            
            if (catId == null) {
                console.error('Category has null/undefined ID:', cat);
                return;
            }
            
            categoryMap.set(catId, {
                id: catId,
                name: cat.name,
                parent_id: parentId,
                display_order: cat.display_order || 0,
                project: cat.project,
                color: cat.color || null,
                children: []
            });
        });
        
        console.log(`Created ${categoryMap.size} category objects in map`);
        
        // Second pass: link children to parents
        let linkedCount = 0;
        let notFoundCount = 0;
        categoryMap.forEach((cat, catId) => {
            if (cat.parent_id != null) {
                const parent = categoryMap.get(cat.parent_id);
                if (parent) {
                    parent.children.push(cat);
                    linkedCount++;
                } else {
                    notFoundCount++;
                    console.warn(`Parent ${cat.parent_id} not found in map for category "${cat.name}" (ID: ${catId}). Available IDs: ${Array.from(categoryMap.keys()).slice(0, 10).join(', ')}`);
                }
            }
        });
        
        console.log(`Linked ${linkedCount} children to parents. ${notFoundCount} parents not found.`);
        
        // Debug: Show what root categories have
        const testRoot = Array.from(categoryMap.values()).find(c => c.name === 'Political Domain');
        if (testRoot) {
            console.log(`Political Domain (ID: ${testRoot.id}) has ${testRoot.children.length} children`);
            if (testRoot.children.length === 0) {
                // Find what should be its children
                const shouldBeChildren = Array.from(categoryMap.values()).filter(c => c.parent_id === testRoot.id);
                console.log(`But ${shouldBeChildren.length} categories have parent_id=${testRoot.id}:`, shouldBeChildren.map(c => `${c.name} (id=${c.id}, parent_id=${c.parent_id})`));
            }
        }
        
        // Filter root categories (no parent_id) for this project
        const rootCategories = Array.from(categoryMap.values())
            .filter(cat => (cat.parent_id === null || cat.parent_id === undefined || cat.parent_id === '') && cat.project === project);
        
        // Sort root categories by display_order
        rootCategories.sort((a, b) => (a.display_order || 0) - (b.display_order || 0));
        
        // Debug: Log root categories and their children
        const debugInfo = document.getElementById('debug-info');
        const debugText = document.getElementById('debug-text');
        
        let debugMsg = `Found ${rootCategories.length} root categories, linked ${linkedCount} children. `;
        
        rootCategories.forEach(root => {
            debugMsg += `${root.name}: ${root.children.length} children. `;
            if (root.children.length > 0) {
                console.log(`✓ ${root.name} has ${root.children.length} children:`, root.children.map(c => c.name));
            }
        });
        
        debugText.textContent = debugMsg;
        debugInfo.style.display = 'block';
        
        function renderCategory(cat, level = 1, parentPath = '', parentContainer = null) {
            const container = document.createElement('div');
            container.className = 'category-container';
            
            // Build full path for coding (but show just name in UI)
            const fullPath = parentPath ? `${parentPath} > ${cat.name}` : cat.name;
            
            const div = document.createElement('div');
            div.className = `category-item level-${level}`;
            
            // Show category name (not full path) for better hierarchy visibility
            div.textContent = cat.name;
            div.dataset.categoryId = cat.id;
            div.dataset.categoryName = fullPath; // Store full path for coding
            div.dataset.categoryColor = cat.color || '#FFFF00'; // Store color, default yellow
            div.dataset.hasChildren = (cat.children && cat.children.length > 0) ? 'true' : 'false';
            
            // Apply category color as left border if available
            if (cat.color) {
                div.style.borderLeftColor = cat.color;
                div.style.borderLeftWidth = '3px';
                div.style.borderLeftStyle = 'solid';
            }
            
            // Add expand/collapse indicator for categories with children
            // Debug: Check if category has children
            const hasChildren = cat.children && Array.isArray(cat.children) && cat.children.length > 0;
            
            if (hasChildren) {
                console.log(`Rendering ${cat.name} WITH ${cat.children.length} children`);
                // Update debug info
                const debugText = document.getElementById('debug-text');
                if (debugText) {
                    const current = debugText.textContent;
                    debugText.textContent = current + ` [${cat.name} has ${cat.children.length} children]`;
                }
                const expandIcon = document.createElement('span');
                expandIcon.className = 'expand-icon';
                expandIcon.textContent = '▶';
                expandIcon.style.marginRight = '4px';
                expandIcon.style.fontSize = '10px';
                expandIcon.style.display = 'inline-block';
                expandIcon.style.cursor = 'pointer';
                div.insertBefore(expandIcon, div.firstChild);
                
                // Create container for children (initially hidden)
                const childrenContainer = document.createElement('div');
                childrenContainer.className = 'category-children';
                childrenContainer.style.display = 'none';
                childrenContainer.style.marginLeft = '16px';
                
                // Sort children by display_order
                cat.children.sort((a, b) => (a.display_order || 0) - (b.display_order || 0));
                cat.children.forEach(child => {
                    renderCategory(child, level + 1, fullPath, childrenContainer);
                });
                
                // Toggle expand/collapse on click
                expandIcon.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const isExpanded = childrenContainer.style.display !== 'none';
                    childrenContainer.style.display = isExpanded ? 'none' : 'block';
                    expandIcon.textContent = isExpanded ? '▶' : '▼';
                });
                
                container.appendChild(div);
                container.appendChild(childrenContainer);
            } else {
                if (cat.children && cat.children.length === 0) {
                    console.log(`Rendering ${cat.name} with EMPTY children array`);
                } else {
                    console.log(`Rendering ${cat.name} with NO children (children=${cat.children})`);
                }
                container.appendChild(div);
            }
            
            // Handle category selection
            div.addEventListener('click', (e) => {
                // Don't trigger selection if clicking expand icon
                if (e.target.classList.contains('expand-icon')) {
                    return;
                }
                
                // Remove previous selection
                document.querySelectorAll('.category-item').forEach(item => {
                    item.classList.remove('selected');
                });
                div.classList.add('selected');
                selectedCategory = fullPath; // Use full path for coding
                selectedCategoryColor = div.dataset.categoryColor || '#FFFF00'; // Store color
                document.getElementById('code-text-btn').disabled = !selectedText;
            });
            
            // Append to parent container or treeDiv
            if (parentContainer) {
                parentContainer.appendChild(container);
            } else {
                treeDiv.appendChild(container);
            }
        }
        
        rootCategories.forEach(cat => renderCategory(cat, 1, ''));
        
    } catch (error) {
        showStatus(`Error loading categories: ${error.message}`, 'error');
    }
}

function filterCategories() {
    const searchTerm = document.getElementById('category-search').value.toLowerCase();
    const items = document.querySelectorAll('.category-item');
    
    items.forEach(item => {
        const text = item.textContent.toLowerCase();
        if (text.includes(searchTerm)) {
            item.style.display = 'block';
        } else {
            item.style.display = 'none';
        }
    });
}

async function generateReport() {
    try {
        showStatus('Generating report...', 'info');
        
        let results = conn.query(`
            SELECT category_str, tagged_text, participant_number, doc_filename, user_comment
            FROM c0d3t3xt
            ORDER BY category_str, id
        `);
        
        // Handle async queries
        if (results && typeof results.then === 'function') {
            results = await results;
        }
        
        if (results.length === 0) {
            showStatus('No coded text found', 'info');
            return;
        }
        
        // Group by category
        const grouped = {};
        results.forEach(row => {
            if (!grouped[row.category_str]) {
                grouped[row.category_str] = [];
            }
            grouped[row.category_str].push(row);
        });
        
        // Generate HTML report
        let html = '<html><head><meta charset="UTF-8"><style>body{font-family:Arial;padding:20px;}h3{margin-top:30px;}blockquote{border-left:3px solid #ccc;padding-left:10px;margin:10px 0;}</style></head><body>';
        html += '<h1>CodeText Report</h1>';
        
        for (const [category, items] of Object.entries(grouped)) {
            html += `<h3>${category}</h3>`;
            items.forEach(item => {
                const prefix = item.participant_number ? `[P${item.participant_number}] ` : '';
                html += `<blockquote>${prefix}"${item.tagged_text}"`;
                if (item.user_comment) {
                    html += `<br><small>Comment: ${item.user_comment}</small>`;
                }
                html += '</blockquote>';
            });
        }
        
        html += '</body></html>';
        
        // Insert at end of document
        await Word.run(async (context) => {
            const body = context.document.body;
            body.insertBreak(Word.BreakType.pageBreak, Word.InsertLocation.end);
            body.insertHtml(html, Word.InsertLocation.end);
            await context.sync();
        });
        
        showStatus(`Report generated with ${results.length} coded items`, 'success');
        
    } catch (error) {
        showStatus(`Error generating report: ${error.message}`, 'error');
        console.error(error);
    }
}

async function exportData() {
    try {
        let results = conn.query('SELECT * FROM c0d3t3xt');
        if (results && typeof results.then === 'function') {
            results = await results;
        }
        const json = JSON.stringify(results, null, 2);
        const blob = new Blob([json], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `codetext-export-${Date.now()}.json`;
        a.click();
        URL.revokeObjectURL(url);
        showStatus('Data exported successfully', 'success');
    } catch (error) {
        showStatus(`Error exporting data: ${error.message}`, 'error');
    }
}

async function importCategories() {
    // Create file input
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';
    input.onchange = async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        try {
            showStatus('Importing categories...', 'info');
            const text = await file.text();
            const data = JSON.parse(text);
            
            if (!data.categories || !Array.isArray(data.categories)) {
                throw new Error('Invalid file format. Expected categories array.');
            }
            
            // Import categories
            let imported = 0;
            for (const cat of data.categories) {
                // Check if exists
                let existing = [];
                if (conn.type === 'indexeddb' || conn.type === 'memory') {
                    let allCats = conn.query('SELECT * FROM categories WHERE name = ? AND project = ?', [cat.name, cat.project]);
                    if (allCats && typeof allCats.then === 'function') {
                        allCats = await allCats;
                    }
                    existing = allCats;
                }
                
                if (existing.length === 0) {
                    // Ensure parent_id is a number or null - handle 0 as valid parent_id
                    let parentId = null;
                    if (cat.parent_id !== null && cat.parent_id !== undefined && cat.parent_id !== '') {
                        parentId = typeof cat.parent_id === 'string' ? parseInt(cat.parent_id, 10) : Number(cat.parent_id);
                        if (isNaN(parentId)) parentId = null;
                    }
                    
                    console.log(`Importing: ${cat.name} with parent_id=${parentId} (original: ${cat.parent_id})`);
                    
                    let result = conn.query(`
                        INSERT INTO categories (name, country, parent_id, project, display_order, color)
                        VALUES (?, ?, ?, ?, ?, ?)
                    `, [cat.name, cat.country || null, parentId, cat.project, cat.display_order || 0, cat.color || null]);
                    if (result && typeof result.then === 'function') {
                        await result;
                    }
                    imported++;
                }
            }
            
            showStatus(`Imported ${imported} categories`, 'success');
            await loadCategories();
            
        } catch (error) {
            showStatus(`Import error: ${error.message}`, 'error');
            console.error(error);
        }
    };
    input.click();
}

async function clearDatabase() {
    if (!confirm('Are you sure you want to clear all coded data?')) {
        return;
    }
    
    try {
        let result1 = conn.query('DELETE FROM c0d3t3xt');
        if (result1 && typeof result1.then === 'function') {
            await result1;
        }
        let result2 = conn.query('DELETE FROM participants');
        if (result2 && typeof result2.then === 'function') {
            await result2;
        }
        showStatus('Database cleared', 'success');
        await loadCategories();
        document.getElementById('output').textContent = '';
    } catch (error) {
        showStatus(`Error clearing database: ${error.message}`, 'error');
    }
}

function showStatus(message, type = 'info') {
    const statusDiv = document.getElementById('status');
    statusDiv.textContent = message;
    // Map 'warning' to 'info' if warning class doesn't exist
    statusDiv.className = type === 'warning' ? 'status-info' : `status-${type}`;
    statusDiv.className = `status-${type}`;
    
    if (type === 'success') {
        setTimeout(() => {
            statusDiv.textContent = 'Ready';
            statusDiv.className = 'status-info';
        }, 3000);
    }
}

