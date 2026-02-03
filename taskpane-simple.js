/* global Office */

// Simple working version with hierarchical categories
let conn = null;
let selectedText = '';
let selectedCategory = '';
let selectedCategoryColor = '#FFFF00';
let allCategories = [];
let currentDocFilename = '';
let currentParagraphText = '';

Office.onReady((info) => {
    console.log('Office ready');
    if (info.host === Office.HostType.Word) {
        initApp();
    }
});

// ===== Participant Parsing Functions (ported from Python) =====

function parseParticipantNumber(taggedText, paragraphText) {
    // Parse participant number from text (P1:, P2:, R446:, _1:, etc.)
    let match = taggedText.match(/([PR_]\d+):/);
    if (match) {
        return match[1].replace('_', 'P');
    }
    
    match = paragraphText.match(/([PR_]\d+):/);
    if (match) {
        return match[1].replace('_', 'P');
    }
    
    return null;
}

function getProjectName(filename) {
    // Determine project name from filename
    if (filename.includes('_FG')) return 'DIANA';
    if (filename.includes('_LVA_')) return 'JOHNSON';
    if (filename.includes('JOHNSON2_')) return 'JOHNSON2';
    if (filename.includes('RIDE_')) return 'RIDE';
    return '';
}

function parseCountryName(filename) {
    // Parse country name from DIANA filename format (e.g., DIANA_Armenia_FG1.docx)
    if (!filename.includes('_FG')) return null;
    
    const parts = filename.split('_FG');
    if (parts.length >= 2) {
        const beforeFG = parts[0].split('_');
        if (beforeFG.length >= 2) {
            return beforeFG[beforeFG.length - 1]; // Last part before _FG
        }
    }
    return null;
}

function parseDemographics(text) {
    // Parse demographics from angle bracket format
    // Format: <(Female /7/, 18-34, Not Available, Agmola)>
    const demographics = {};
    
    // Try angle bracket with parentheses format: <(...)>
    let angleBracketMatch = text.match(/<\((.*?)\)>/);
    
    // Also try pipe format: <|...|>
    if (!angleBracketMatch) {
        angleBracketMatch = text.match(/<\|(.*?)\|>/);
    }
    
    if (angleBracketMatch) {
        const parts = angleBracketMatch[1].split(',').map(p => p.trim());
        
        console.log('Found demographics string:', angleBracketMatch[1]);
        console.log('Split into parts:', parts);
        
        // Parse the parts
        if (parts.length >= 1) {
            // First part is gender (may include /number/ notation)
            const genderPart = parts[0].trim();
            // Extract just the gender word (Female, Male, etc.)
            const genderMatch = genderPart.match(/^(Female|Male|Other)/i);
            if (genderMatch) {
                demographics.gender = genderMatch[1];
            }
        }
        
        if (parts.length >= 2) {
            // Second part is age range
            demographics.age = parts[1].trim();
        }
        
        if (parts.length >= 3) {
            // Third part is occupation
            const occupation = parts[2].trim();
            if (occupation && occupation !== 'Not Available' && occupation !== 'N/A') {
                demographics.occupation = occupation;
            }
        }
        
        if (parts.length >= 4) {
            // Fourth part is location
            demographics.location = parts[3].trim();
        }
        
        console.log('Parsed demographics:', demographics);
        return demographics;
    }
    
    // Try XML format as fallback
    const genderMatch = text.match(/<gender>(.*?)<\/gender>/i);
    if (genderMatch) demographics.gender = genderMatch[1];
    
    const ageMatch = text.match(/<age>(.*?)<\/age>/i);
    if (ageMatch) demographics.age = ageMatch[1];
    
    const occupationMatch = text.match(/<occupation>(.*?)<\/occupation>/i);
    if (occupationMatch) demographics.occupation = occupationMatch[1];
    
    const cityMatch = text.match(/<city>(.*?)<\/city>/i);
    if (cityMatch) demographics.city = cityMatch[1];
    
    const locationMatch = text.match(/<location>(.*?)<\/location>/i);
    if (locationMatch) demographics.location = locationMatch[1];
    
    return demographics;
}

function createCitation(participantNumber, demographics, project) {
    // Create citation string based on project type
    if (!participantNumber && Object.keys(demographics).length === 0) {
        return null;
    }
    
    const parts = [];
    
    if (participantNumber) {
        parts.push(participantNumber);
    }
    
    if (demographics.gender) parts.push(demographics.gender);
    if (demographics.age) parts.push(demographics.age);
    if (demographics.occupation) parts.push(demographics.occupation);
    if (demographics.city) parts.push(demographics.city);
    if (demographics.location) parts.push(demographics.location);
    
    return parts.length > 0 ? `(${parts.join(', ')})` : null;
}

function initApp() {
    console.log('Init starting');
    document.getElementById('main-content').style.display = 'block';
    document.getElementById('status').textContent = 'Loading categories...';
    
    // Simple in-memory storage
    conn = {
        data: {
            categories: [],
            coded: []
        }
    };
    
    setupHandlers();
    loadCategoriesFromJSON();
    
    // Set up auto-capture of selected text
    setupAutoSelection();
}

function setupAutoSelection() {
    // Debounce timer to prevent too many rapid calls
    let selectionTimeout = null;
    
    // Automatically capture text when selection changes in Word
    Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        function() {
            // Show spinner immediately
            document.getElementById('text-loading-spinner').style.display = 'flex';
            
            // Clear any pending timeout
            if (selectionTimeout) {
                clearTimeout(selectionTimeout);
            }
            
            // Wait 200ms after user stops selecting before processing
            // This reduces lag by batching rapid selection changes
            selectionTimeout = setTimeout(() => {
                getSelection();
            }, 200);
        }
    );
    console.log('Auto-selection enabled with debouncing and loading indicator');
}

function setupHandlers() {
    document.getElementById('get-selection-btn').onclick = getSelection;
    document.getElementById('code-text-btn').onclick = codeText;
    document.getElementById('project-select').onchange = () => renderCategories();
    document.getElementById('generate-report-btn').onclick = generateReport;
    document.getElementById('export-data-btn').onclick = exportData;
    document.getElementById('clear-db-btn').onclick = clearData;
    document.getElementById('import-categories-btn').onclick = loadCategoriesFromJSON;
    
    // Search handler
    const searchInput = document.getElementById('category-search');
    if (searchInput) {
        searchInput.addEventListener('input', (e) => {
            const term = e.target.value.toLowerCase();
            filterCategories(term);
        });
    }
}

function getSelection() {
    // PHASE 1: Get text FAST and update immediately
    Word.run(async (context) => {
        const range = context.document.getSelection();
        range.load('text');
        
        await context.sync();
        
        selectedText = range.text;
        
        // Hide spinner and show text
        document.getElementById('text-loading-spinner').style.display = 'none';
        
        // INSTANT display - this Word.run completes here, UI updates immediately
        let display = selectedText.substring(0, 200) + (selectedText.length > 200 ? '...' : '');
        document.getElementById('selected-text-display').textContent = display;
        document.getElementById('output').textContent = 'Parsing demographics...';
        document.getElementById('code-text-btn').disabled = !selectedText || !selectedCategory;
    }).then(() => {
        // PHASE 2: After UI updates, run demographics in separate Word.run (truly async)
        parseDemographicsAsync();
    }).catch(() => {
        // Hide spinner on error too
        document.getElementById('text-loading-spinner').style.display = 'none';
    });
}

async function parseDemographicsAsync() {
    // Separate Word.run - doesn't block the first one
    Word.run(async (context) => {
        try {
            const range = context.document.getSelection();
            const doc = context.document;
            const paragraph = range.paragraphs.getFirst();
            const nextPara1 = paragraph.getNext();
            const nextPara2 = nextPara1.getNext();
            
            paragraph.load('text');
            nextPara1.load('text');
            nextPara2.load('text');
            doc.load('properties');
            
            await context.sync();
            
            // Combine paragraphs to find demographics
            let combinedText = paragraph.text;
            try {
                if (nextPara1.text) combinedText += ' ' + nextPara1.text;
            } catch (e) {}
            try {
                if (nextPara2.text) combinedText += ' ' + nextPara2.text;
            } catch (e) {}
            
            currentParagraphText = combinedText;
            
            // Get document filename
            try {
                currentDocFilename = doc.properties.title || 'Untitled';
                if (!currentDocFilename || currentDocFilename.trim() === '') {
                    currentDocFilename = 'Document';
                }
            } catch (e) {
                currentDocFilename = 'Document';
            }
            
            // Parse demographics
            const participantId = parseParticipantNumber(selectedText, currentParagraphText);
            const demographics = parseDemographics(currentParagraphText);
            
            // Update output with demographics info
            let debugMsg = '';
            if (participantId) {
                debugMsg += `✓ ${participantId}`;
            }
            
            if (Object.keys(demographics).length > 0) {
                const demoList = Object.entries(demographics).map(([k, v]) => v).join(', ');
                debugMsg += debugMsg ? ` | ${demoList}` : demoList;
            }
            
            if (!debugMsg) {
                debugMsg = 'No participant info found';
            }
            
            document.getElementById('output').textContent = debugMsg;
            
        } catch (error) {
            console.error('Error parsing demographics:', error);
            document.getElementById('output').textContent = 'Ready';
        }
    });
}

function codeText() {
    if (!selectedText || !selectedCategory) return;
    
    Word.run(async (context) => {
        const range = context.document.getSelection();
        
        // Parse participant information
        const participantNumber = parseParticipantNumber(selectedText, currentParagraphText);
        const demographics = parseDemographics(currentParagraphText);
        const project = getProjectName(currentDocFilename);
        const country = parseCountryName(currentDocFilename);
        const citation = createCitation(participantNumber, demographics, project);
        
        // Build comment text with category and metadata
        let commentText = `Category: ${selectedCategory}`;
        if (citation) {
            commentText += `\n${citation}`;
        }
        if (project) {
            commentText += `\nProject: ${project}`;
        }
        commentText += `\nCoded: ${new Date().toLocaleString()}`;
        
        // Add pale yellow highlighting to all coded text FIRST (before comment)
        range.font.highlightColor = '#FFFF00';
        
        // Then add comment to the selected text
        const comment = range.insertComment(commentText);
        
        await context.sync();
        
        console.log(`Added comment: ${commentText}`);
        
        // Store coded text with participant info
        conn.data.coded.push({
            text: selectedText,
            category: selectedCategory,
            color: selectedCategoryColor,
            wordColor: 'Yellow',
            timestamp: new Date(),
            participantNumber: participantNumber,
            demographics: demographics,
            citation: citation,
            project: project,
            country: country,
            docFilename: currentDocFilename
        });
        
        let output = `Coded: "${selectedText.substring(0, 50)}..." as "${selectedCategory}"`;
        if (participantNumber) {
            output += ` [${participantNumber}]`;
        }
        if (citation) {
            output += ` ${citation}`;
        }
        output += ' (Comment added)';
        
        document.getElementById('output').textContent = output;
        document.getElementById('status').textContent = `Coded ${conn.data.coded.length} items`;
    });
}

// Map hex colors to Word's supported highlight colors
function mapHexToWordColor(hexColor) {
    // Word supports these highlight colors:
    // Yellow, Turquoise, Pink, Green, Blue, Red, DarkBlue, Teal, 
    // DarkYellow, Gray, DarkRed, DarkGray, Black, White
    
    if (!hexColor) return 'Yellow';
    
    const hex = hexColor.toUpperCase();
    
    // Map common colors from our palette to Word colors
    const colorMap = {
        '#FF6B6B': 'Red',        // Red tones -> Red
        '#F38181': 'Pink',       // Pink tones -> Pink
        '#E74C3C': 'Red',        // Dark red -> Red
        '#4ECDC4': 'Turquoise',  // Turquoise -> Turquoise
        '#95E1D3': 'Turquoise',  // Light turquoise -> Turquoise
        '#4D96FF': 'Blue',       // Blue -> Blue
        '#9B59B6': 'Pink',       // Purple -> Pink (closest)
        '#AA96DA': 'Pink',       // Light purple -> Pink
        '#FFD93D': 'Yellow',     // Yellow -> Yellow
        '#6BCF7F': 'Green',      // Green -> Green
        '#FFFF00': 'Yellow'      // Yellow -> Yellow
    };
    
    // Check exact match first
    if (colorMap[hex]) {
        return colorMap[hex];
    }
    
    // Fallback: analyze RGB to pick closest color
    const r = parseInt(hex.slice(1, 3), 16);
    const g = parseInt(hex.slice(3, 5), 16);
    const b = parseInt(hex.slice(5, 7), 16);
    
    // Simple heuristic based on dominant channel
    if (r > 200 && g > 200 && b < 100) return 'Yellow';
    if (r > 200 && g < 100 && b < 100) return 'Red';
    if (r < 100 && g > 150 && b > 150) return 'Turquoise';
    if (r < 100 && g < 100 && b > 200) return 'Blue';
    if (r < 150 && g > 200 && b < 150) return 'Green';
    if (r > 200 && g < 150 && b > 150) return 'Pink';
    
    return 'Yellow'; // Default fallback
}

async function loadCategoriesFromJSON() {
    try {
        document.getElementById('status').textContent = 'Loading categories...';
        const response = await fetch('test-categories.json');
        const data = await response.json();
        
        if (data.categories && Array.isArray(data.categories)) {
            allCategories = data.categories;
            console.log(`Loaded ${allCategories.length} categories`);
            document.getElementById('status').textContent = 'Ready';
            renderCategories();
        } else {
            throw new Error('Invalid category data format');
        }
    } catch (error) {
        console.error('Error loading categories:', error);
        document.getElementById('status').textContent = 'Error loading categories';
        // Fallback to simple categories
        allCategories = [
            {id: 1, name: 'Political Domain', parent_id: null, project: 'DIANA', display_order: 1, color: '#FFE0B2'},
            {id: 2, name: 'Economic Domain', parent_id: null, project: 'DIANA', display_order: 2, color: '#C8E6C9'},
            {id: 3, name: 'Social Domain', parent_id: null, project: 'DIANA', display_order: 3, color: '#B3E5FC'}
        ];
        renderCategories();
    }
}

function renderCategories() {
    const project = document.getElementById('project-select').value;
    const treeDiv = document.getElementById('category-tree');
    treeDiv.innerHTML = '';
    
    // Filter by project
    const projectCategories = allCategories.filter(cat => cat.project === project);
    console.log(`Rendering ${projectCategories.length} categories for ${project}`);
    
    if (projectCategories.length === 0) {
        treeDiv.innerHTML = '<p style="padding:10px;color:#666;">No categories for this project. Click "Import Categories" to load them.</p>';
        return;
    }
    
    // Build tree structure
    const categoryMap = new Map();
    projectCategories.forEach(cat => {
        categoryMap.set(cat.id, { ...cat, children: [] });
    });
    
    // Link children to parents
    const rootCategories = [];
    categoryMap.forEach(cat => {
        if (cat.parent_id && categoryMap.has(cat.parent_id)) {
            categoryMap.get(cat.parent_id).children.push(cat);
        } else {
            rootCategories.push(cat);
        }
    });
    
    // Sort by display_order
    const sortCategories = (cats) => {
        cats.sort((a, b) => (a.display_order || 0) - (b.display_order || 0));
        cats.forEach(cat => {
            if (cat.children.length > 0) {
                sortCategories(cat.children);
            }
        });
    };
    sortCategories(rootCategories);
    
    console.log(`Found ${rootCategories.length} root categories`);
    
    // Render tree
    rootCategories.forEach(cat => {
        renderCategoryNode(cat, treeDiv, 0);
    });
}

function renderCategoryNode(category, parentElement, level) {
    const container = document.createElement('div');
    container.className = 'category-container';
    container.style.marginLeft = (level * 20) + 'px';
    
    const item = document.createElement('div');
    item.className = 'category-item';
    item.style.padding = '8px';
    item.style.cursor = 'pointer';
    item.style.borderBottom = '1px solid #eee';
    item.style.display = 'flex';
    item.style.alignItems = 'center';
    
    // Add expand/collapse icon if has children
    if (category.children && category.children.length > 0) {
        const icon = document.createElement('span');
        icon.className = 'expand-icon';
        icon.textContent = '▶';
        icon.style.marginRight = '8px';
        icon.style.fontSize = '10px';
        icon.style.transition = 'transform 0.2s';
        item.appendChild(icon);
    } else {
        const spacer = document.createElement('span');
        spacer.style.width = '18px';
        spacer.style.display = 'inline-block';
        item.appendChild(spacer);
    }
    
    // Category name with color indicator
    const nameSpan = document.createElement('span');
    nameSpan.textContent = category.name;
    if (category.color) {
        const colorDot = document.createElement('span');
        colorDot.style.display = 'inline-block';
        colorDot.style.width = '12px';
        colorDot.style.height = '12px';
        colorDot.style.borderRadius = '50%';
        colorDot.style.backgroundColor = category.color;
        colorDot.style.marginRight = '6px';
        item.appendChild(colorDot);
    }
    item.appendChild(nameSpan);
    
    // Click handler for selection
    item.onclick = (e) => {
        e.stopPropagation();
        
        // If has children, toggle expand/collapse
        if (category.children && category.children.length > 0) {
            const icon = item.querySelector('.expand-icon');
            const childrenDiv = container.querySelector('.category-children');
            if (childrenDiv.style.display === 'none') {
                childrenDiv.style.display = 'block';
                icon.style.transform = 'rotate(90deg)';
            } else {
                childrenDiv.style.display = 'none';
                icon.style.transform = 'rotate(0deg)';
            }
        }
        
        // Select this category
        document.querySelectorAll('.category-item').forEach(el => el.style.background = '');
        item.style.background = '#e3f2fd';
        selectedCategory = category.name;
        selectedCategoryColor = category.color || '#FFFF00';
        document.getElementById('code-text-btn').disabled = !selectedText;
    };
    
    container.appendChild(item);
    
    // Render children
    if (category.children && category.children.length > 0) {
        const childrenDiv = document.createElement('div');
        childrenDiv.className = 'category-children';
        childrenDiv.style.display = 'none'; // Start collapsed
        
        category.children.forEach(child => {
            renderCategoryNode(child, childrenDiv, level + 1);
        });
        
        container.appendChild(childrenDiv);
    }
    
    parentElement.appendChild(container);
}

function filterCategories(searchTerm) {
    if (!searchTerm) {
        // Show all
        document.querySelectorAll('.category-container').forEach(el => el.style.display = 'block');
        return;
    }
    
    // Filter categories
    document.querySelectorAll('.category-container').forEach(container => {
        const itemText = container.querySelector('.category-item')?.textContent.toLowerCase() || '';
        if (itemText.includes(searchTerm)) {
            container.style.display = 'block';
            // Expand parent
            const childrenDiv = container.querySelector('.category-children');
            if (childrenDiv) {
                childrenDiv.style.display = 'block';
            }
        } else {
            container.style.display = 'none';
        }
    });
}

function generateReport() {
    console.log('Generating report...');
    document.getElementById('status').textContent = 'Generating report...';
    
    if (conn.data.coded.length === 0) {
        document.getElementById('status').textContent = 'No coded text found';
        document.getElementById('output').textContent = 'No coded text to report. Code some text first.';
        return;
    }
    
    // Group by category
    const grouped = {};
    conn.data.coded.forEach(item => {
        if (!grouped[item.category]) {
            grouped[item.category] = [];
        }
        grouped[item.category].push(item);
    });
    
    // Generate HTML report with participant information
    let html = '<html><head><meta charset="UTF-8"><style>';
    html += 'body{font-family:Arial,sans-serif;padding:20px;line-height:1.6;}';
    html += 'h1{color:#0078d4;border-bottom:2px solid #0078d4;padding-bottom:10px;}';
    html += 'h3{color:#333;margin-top:30px;margin-bottom:15px;}';
    html += 'blockquote{border-left:3px solid #0078d4;padding-left:15px;margin:10px 0 15px 0;color:#333;}';
    html += '.citation{color:#0078d4;font-weight:bold;margin-right:5px;}';
    html += '.demographics{color:#666;font-size:0.95em;margin-top:3px;}';
    html += '.metadata{color:#999;font-size:0.85em;font-style:italic;margin-top:5px;}';
    html += '</style></head><body>';
    html += '<h1>CodeText Report</h1>';
    html += `<p><strong>Generated:</strong> ${new Date().toLocaleString()}</p>`;
    html += `<p><strong>Total coded items:</strong> ${conn.data.coded.length}</p>`;
    
    // Sort categories alphabetically
    const sortedCategories = Object.keys(grouped).sort();
    
    for (const category of sortedCategories) {
        const items = grouped[category];
        html += `<h3>${category} (${items.length} items)</h3>`;
        
        items.forEach((item, index) => {
            html += '<blockquote>';
            
            // Add citation if available
            if (item.citation) {
                html += `<span class="citation">${item.citation}</span>`;
            } else if (item.participantNumber) {
                html += `<span class="citation">[${item.participantNumber}]</span>`;
            }
            
            // Add quoted text
            html += `"${item.text}"`;
            
            // Add demographics if available
            if (item.demographics && Object.keys(item.demographics).length > 0) {
                const demoStr = Object.entries(item.demographics)
                    .map(([key, value]) => `${key}: ${value}`)
                    .join(', ');
                html += `<div class="demographics">${demoStr}</div>`;
            }
            
            // Add metadata (timestamp, document)
            let metadata = [];
            if (item.docFilename) {
                metadata.push(`Document: ${item.docFilename}`);
            }
            if (item.project) {
                metadata.push(`Project: ${item.project}`);
            }
            if (item.country) {
                metadata.push(`Country: ${item.country}`);
            }
            metadata.push(`Coded: ${item.timestamp.toLocaleString()}`);
            
            html += `<div class="metadata">${metadata.join(' | ')}</div>`;
            html += '</blockquote>';
        });
    }
    
    html += '</body></html>';
    
    // Insert at end of document
    Word.run(async (context) => {
        const body = context.document.body;
        body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
        body.insertHtml(html, Word.InsertLocation.end);
        await context.sync();
        
        document.getElementById('status').textContent = `Report generated with ${conn.data.coded.length} items`;
        document.getElementById('output').textContent = `Report inserted into document with ${Object.keys(grouped).length} categories.`;
    }).catch(error => {
        console.error('Error generating report:', error);
        document.getElementById('status').textContent = 'Error generating report';
        document.getElementById('output').textContent = `Error: ${error.message}`;
    });
}

function exportData() {
    const json = JSON.stringify(conn.data.coded, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `codetext-export-${Date.now()}.json`;
    a.click();
    URL.revokeObjectURL(url);
    document.getElementById('status').textContent = 'Data exported';
    document.getElementById('output').textContent = `Exported ${conn.data.coded.length} coded items`;
}

function clearData() {
    if (confirm('Clear all coded data?')) {
        conn.data.coded = [];
        document.getElementById('status').textContent = 'Data cleared';
        document.getElementById('output').textContent = 'All coded text has been cleared.';
    }
}

