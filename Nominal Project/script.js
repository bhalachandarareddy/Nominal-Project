let allCadetsData = [];
let sortedCadetsForDisplay = [];
let headers = null;

// Rank order for sorting
const RANK_ORDER = ['SUO', 'JUO', 'CSM', 'SGT', 'CPL', 'LCPL', 'CDT'];

// Button handlers
document.getElementById('load-files-btn').addEventListener('click', loadFilesFromFolder);
document.getElementById('select-all-btn').addEventListener('click', selectAllCadets);
document.getElementById('deselect-all-btn').addEventListener('click', deselectAllCadets);
document.getElementById('generate-btn').addEventListener('click', generateNominal);
document.getElementById('search-input').addEventListener('input', handleSearch);

async function loadFilesFromFolder() {
    try {
        // Fetch the list of files from the 'files' folder
        const response = await fetch('files/');
        const html = await response.text();
        
        // Parse HTML to extract Excel file names
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');
        const links = doc.querySelectorAll('a');
        
        const excelFiles = [];
        links.forEach(link => {
            const href = link.getAttribute('href');
            if (href && (href.endsWith('.xlsx') || href.endsWith('.xls'))) {
                excelFiles.push('files/' + href);
            }
        });
        
        if (excelFiles.length === 0) {
            alert('No Excel files found in the "files" folder. Please add Excel files to the folder.');
            return;
        }
        
        document.getElementById('file-status').textContent = `Loading ${excelFiles.length} file(s)...`;
        
        // Load all Excel files
        allCadetsData = [];
        for (const filePath of excelFiles) {
            await loadExcelFile(filePath);
        }
        
        document.getElementById('file-status').textContent = `Loaded ${excelFiles.length} file(s) with ${allCadetsData.length} cadets`;
        
        if (allCadetsData.length > 0) {
            document.getElementById('heading-section').style.display = 'block';
            displayCadets();
        } else {
            alert('No cadet data found in the Excel files.');
        }
        
    } catch (error) {
        console.error('Error loading files:', error);
        alert('Error loading files from folder. Make sure the server is running and files exist in the "files" folder.');
    }
}

async function loadExcelFile(filePath) {
    try {
        const response = await fetch(filePath);
        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true, dateNF: 'm/d/yyyy' });
        
        // Get the first sheet
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        
        // Convert to JSON
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, dateNF: 'm/d/yyyy' });
        
        if (jsonData.length < 2) return;
        
        // Set headers from the first file
        if (!headers) {
            headers = jsonData[0];
        }
        
        // Find column indices
        const nameIndex = headers.findIndex(h => h && h.toString().toLowerCase().includes('name'));
        const rankIndex = headers.findIndex(h => h && h.toString().toLowerCase().includes('rank'));
        const regNoIndex = headers.findIndex(h => h && (h.toString().toLowerCase().includes('regimental') || h.toString().toLowerCase().includes('reg') || h.toString().toLowerCase().includes('regt')));
        
        // Extract cadet data
        for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            const name = row[nameIndex];
            
            if (name && name.toString().trim() !== '') {
                // Determine batch from regimental number
                let batch = 'C';
                if (regNoIndex >= 0 && row[regNoIndex]) {
                    const regNo = row[regNoIndex].toString().trim();
                    batch = getBatchFromRegNo(regNo);
                }
                
                // Debug log
                const regNoForDebug = regNoIndex >= 0 ? row[regNoIndex] : 'N/A';
                console.log(`Name: ${name.toString().trim()}, RegNo: ${regNoForDebug}, Detected Batch: ${batch}`);
                
                allCadetsData.push({
                    rowData: row,
                    sheet: sheet,
                    name: name.toString().trim(),
                    rank: rankIndex >= 0 ? (row[rankIndex] || 'CDT').toString().trim().toUpperCase() : 'CDT',
                    batch: batch,
                    originalRowIndex: i
                });
            }
        }
        
    } catch (error) {
        console.error(`Error loading file ${filePath}:`, error);
    }
}

function displayCadets() {
    if (allCadetsData.length === 0) {
        alert('No cadet data found.');
        return;
    }

    // Sort cadets: First by batch (C, B2, B1), then by rank order
    sortedCadetsForDisplay = sortCadets([...allCadetsData]);
    
    // Group by batch and rank
    const groupedData = groupCadetsByBatchAndRank(sortedCadetsForDisplay);
    
    // Display cadets
    const cadetList = document.getElementById('cadet-list');
    cadetList.innerHTML = '';

    let cadetIndex = 0;
    
    // Display in order: C, B2, B1
    ['C', 'B2', 'B1'].forEach(batch => {
        if (groupedData[batch]) {
            // Create batch group
            const batchGroup = document.createElement('div');
            batchGroup.className = 'batch-group';
            
            const batchHeader = document.createElement('div');
            batchHeader.className = 'batch-header';
            batchHeader.textContent = `Batch ${batch}`;
            batchGroup.appendChild(batchHeader);
            
            // Display ranks in order
            RANK_ORDER.forEach(rank => {
                if (groupedData[batch][rank] && groupedData[batch][rank].length > 0) {
                    const rankGroup = document.createElement('div');
                    rankGroup.className = 'rank-group';
                    
                    const rankHeader = document.createElement('div');
                    rankHeader.className = 'rank-header';
                    rankHeader.textContent = rank;
                    rankGroup.appendChild(rankHeader);
                    
                    groupedData[batch][rank].forEach(cadet => {
                        const cadetItem = document.createElement('div');
                        cadetItem.className = 'cadet-item';
                        cadetItem.dataset.name = cadet.name.toLowerCase();
                        cadetItem.dataset.batch = cadet.batch;
                        cadetItem.dataset.rank = cadet.rank;
                        
                        const checkbox = document.createElement('input');
                        checkbox.type = 'checkbox';
                        checkbox.id = `cadet-${cadetIndex}`;
                        checkbox.dataset.index = cadetIndex;
                        
                        const nameSpan = document.createElement('span');
                        nameSpan.className = 'cadet-name';
                        nameSpan.textContent = cadet.name;
                        
                        const rankSpan = document.createElement('span');
                        rankSpan.className = 'cadet-rank';
                        rankSpan.textContent = cadet.rank;
                        
                        const batchSpan = document.createElement('span');
                        batchSpan.className = 'cadet-batch';
                        batchSpan.textContent = cadet.batch;
                        
                        cadetItem.appendChild(checkbox);
                        cadetItem.appendChild(nameSpan);
                        cadetItem.appendChild(rankSpan);
                        cadetItem.appendChild(batchSpan);
                        
                        rankGroup.appendChild(cadetItem);
                        cadetIndex++;
                    });
                    
                    batchGroup.appendChild(rankGroup);
                }
            });
            
            cadetList.appendChild(batchGroup);
        }
    });

    // Show the cadet section
    document.getElementById('cadet-section').style.display = 'block';
}

function sortCadets(cadets) {
    return cadets.sort((a, b) => {
        // Normalize batch names
        const batchA = normalizeBatch(a.batch);
        const batchB = normalizeBatch(b.batch);
        
        // Sort by batch first (C, B2, B1)
        const batchOrder = { 'C': 0, 'B2': 1, 'B1': 2 };
        const batchComparison = (batchOrder[batchA] || 999) - (batchOrder[batchB] || 999);
        
        if (batchComparison !== 0) return batchComparison;
        
        // Then sort by rank
        const rankA = RANK_ORDER.indexOf(a.rank);
        const rankB = RANK_ORDER.indexOf(b.rank);
        const rankComparison = (rankA === -1 ? 999 : rankA) - (rankB === -1 ? 999 : rankB);
        
        if (rankComparison !== 0) return rankComparison;
        
        // Finally sort by name
        return a.name.localeCompare(b.name);
    });
}

function getBatchFromRegNo(regNo) {
    // Remove any spaces and convert to uppercase
    const cleanRegNo = regNo.toString().trim().toUpperCase();
    
    // Check for different regimental number patterns
    // Pattern 1: AP2023SWA... or AP21SWA... (C batch)
    if (cleanRegNo.includes('2023') || cleanRegNo.match(/AP2[01]SWA/)) {
        return 'C';
    }
    
    // Pattern 2: AP2024SDIA... (B2 batch)
    if (cleanRegNo.includes('2024')) {
        return 'B2';
    }
    
    // Pattern 3: AP2025SDIA... (B1 batch)
    if (cleanRegNo.includes('2025')) {
        return 'B1';
    }
    
    // Fallback: Try to extract year from anywhere in the string
    const yearMatch = cleanRegNo.match(/20(23|24|25)/);
    if (yearMatch) {
        const year = yearMatch[0];
        if (year === '2023') return 'C';
        if (year === '2024') return 'B2';
        if (year === '2025') return 'B1';
    }
    
    // Default to C if unable to determine
    return 'C';
}

function normalizeBatch(batch) {
    const b = batch.toUpperCase();
    if (b.includes('C') && !b.includes('B')) return 'C';
    if (b.includes('B2') || b.includes('B-2')) return 'B2';
    if (b.includes('B1') || b.includes('B-1')) return 'B1';
    return b;
}

function groupCadetsByBatchAndRank(cadets) {
    const grouped = {};
    
    cadets.forEach(cadet => {
        const batch = normalizeBatch(cadet.batch);
        const rank = cadet.rank;
        
        if (!grouped[batch]) {
            grouped[batch] = {};
        }
        
        if (!grouped[batch][rank]) {
            grouped[batch][rank] = [];
        }
        
        grouped[batch][rank].push(cadet);
    });
    
    return grouped;
}

function handleSearch(event) {
    const searchTerm = event.target.value.toLowerCase().trim();
    const cadetItems = document.querySelectorAll('.cadet-item');
    
    cadetItems.forEach(item => {
        const name = item.dataset.name;
        if (name && name.includes(searchTerm)) {
            item.classList.remove('hidden');
        } else {
            item.classList.add('hidden');
        }
    });
}

function selectAllCadets() {
    const checkboxes = document.querySelectorAll('.cadet-item:not(.hidden) input[type="checkbox"]');
    checkboxes.forEach(checkbox => {
        checkbox.checked = true;
    });
}

function deselectAllCadets() {
    const checkboxes = document.querySelectorAll('.cadet-item:not(.hidden) input[type="checkbox"]');
    checkboxes.forEach(checkbox => {
        checkbox.checked = false;
    });
}

function generateNominal() {
    // Get selected cadets
    const checkboxes = document.querySelectorAll('.cadet-item input[type="checkbox"]:checked');
    
    if (checkboxes.length === 0) {
        alert('Please select at least one cadet to generate the nominal.');
        return;
    }

    // Get heading
    const heading = document.getElementById('heading-input').value.trim();
    
    // Get selected cadet data in sorted display order
    const selectedCadets = [];
    
    checkboxes.forEach(cb => {
        const index = parseInt(cb.dataset.index);
        if (index >= 0 && index < sortedCadetsForDisplay.length) {
            selectedCadets.push(sortedCadetsForDisplay[index]);
        }
    });
    
    // Cadets are already in the correct order from sortedCadetsForDisplay
    const sortedSelectedCadets = selectedCadets;
    
    // Create new workbook
    const newWorkbook = XLSX.utils.book_new();
    const wsData = [];
    
    // Add heading row if provided (merged cell)
    let startRow = 0;
    if (heading) {
        wsData.push([heading]);
        startRow = 1;
    }
    
    // Find SL NO column index
    const slNoIndex = headers.findIndex(h => h && (h.toString().toLowerCase().includes('sl') || h.toString().toLowerCase().includes('s.no') || h.toString().toLowerCase().includes('serial')));
    
    // Add header row
    wsData.push(headers);
    
    // Add selected cadet rows with updated SL NO
    sortedSelectedCadets.forEach((cadet, index) => {
        const rowData = [...cadet.rowData]; // Create a copy
        
        // Update SL NO to sequential number
        if (slNoIndex >= 0) {
            rowData[slNoIndex] = index + 1;
        }
        
        wsData.push(rowData);
    });
    
    // Create worksheet
    const newWorksheet = XLSX.utils.aoa_to_sheet(wsData);
    
    // If heading exists, merge cells in first row
    if (heading) {
        const range = XLSX.utils.decode_range(newWorksheet['!ref']);
        const numCols = range.e.c + 1;
        
        newWorksheet['!merges'] = [{
            s: { r: 0, c: 0 },
            e: { r: 0, c: numCols - 1 }
        }];
        
        // Style the heading cell
        const headingCell = newWorksheet['A1'];
        if (headingCell) {
            headingCell.s = {
                font: { bold: true, sz: 16 },
                alignment: { horizontal: 'center', vertical: 'center' },
                fill: { fgColor: { rgb: 'E0E0E0' } }
            };
        }
    }
    
    // Copy formatting from original data
    const range = XLSX.utils.decode_range(newWorksheet['!ref']);
    const headerRow = heading ? 1 : 0;
    const dataStartRow = heading ? 2 : 1;
    
    // Apply formatting to data cells
    for (let R = dataStartRow; R <= range.e.r; R++) {
        const cadetIndex = R - dataStartRow;
        if (cadetIndex < sortedSelectedCadets.length) {
            const cadet = sortedSelectedCadets[cadetIndex];
            
            for (let C = 0; C <= range.e.c; C++) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                const originalRowIndex = cadet.originalRowIndex;
                const originalSheet = cadet.sheet;
                const originalCellAddress = XLSX.utils.encode_cell({ r: originalRowIndex, c: C });
                
                if (originalSheet[originalCellAddress]) {
                    const originalCell = originalSheet[originalCellAddress];
                    
                    if (!newWorksheet[cellAddress]) {
                        newWorksheet[cellAddress] = {};
                    }
                    
                    // Copy cell properties
                    newWorksheet[cellAddress].t = originalCell.t;
                    newWorksheet[cellAddress].v = originalCell.v;
                    
                    if (originalCell.z) {
                        newWorksheet[cellAddress].z = originalCell.z;
                    }
                    if (originalCell.w) {
                        newWorksheet[cellAddress].w = originalCell.w;
                    }
                    if (originalCell.s) {
                        newWorksheet[cellAddress].s = originalCell.s;
                    }
                }
            }
        }
    }
    
    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Nominal');

    // Generate filename with current date and time
    const now = new Date();
    const dateStr = now.toISOString().slice(0, 10);
    const timeStr = now.toTimeString().slice(0, 8).replace(/:/g, '-');
    const fileName = `Nominal_${dateStr}_${timeStr}.xlsx`;

    // Download the file
    XLSX.writeFile(newWorkbook, fileName, { bookType: 'xlsx', cellDates: false });
    
    // Show success message
    alert(`Nominal generated successfully!\nFile: ${fileName}\nCadets included: ${checkboxes.length}`);
}