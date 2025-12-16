let allCadets = [];
let displayedCadets = [];

const RANK_ORDER = ['SUO', 'JUO', 'CSM', 'SGT', 'CPL', 'LCPL', 'CDT'];
const BATCH_ORDER = ['C', 'B2', 'B1']; // C is senior

document.getElementById('file-upload').addEventListener('change', handleFileUpload);
document.getElementById('select-all').addEventListener('click', selectAll);
document.getElementById('deselect-all').addEventListener('click', deselectAll);
document.getElementById('search').addEventListener('input', handleSearch);
document.getElementById('generate-btn').addEventListener('click', generateNominal);

function updateCount() {
    const count = document.querySelectorAll('.cadet-item input:checked').length;
    document.getElementById('selected-count').textContent = `Selected: ${count}`;
}

async function handleFileUpload(event) {
    const files = event.target.files;
    
    if (files.length === 0) {
        return;
    }
    
    document.getElementById('status').textContent = 'Loading files...';
    allCadets = [];
    
    try {
        for (const file of files) {
            await loadExcelFile(file);
        }
        
        document.getElementById('status').textContent = `Loaded ${allCadets.length} cadets from ${files.length} file(s)`;
        document.getElementById('main-section').style.display = 'block';
        displayCadets();
    } catch (error) {
        alert('Error loading files: ' + error.message);
        document.getElementById('status').textContent = 'Error loading files';
    }
}

async function loadExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
                
                if (jsonData.length < 2) {
                    resolve();
                    return;
                }
                
                const headers = jsonData[0];
                const nameIdx = headers.findIndex(h => h && h.toLowerCase().includes('name'));
                const rankIdx = headers.findIndex(h => h && h.toLowerCase().includes('rank'));
                const regIdx = headers.findIndex(h => h && (h.toLowerCase().includes('reg') || h.toLowerCase().includes('regt')));
                
                console.log(`Loading file: ${file.name}`);
                console.log(`Found ${jsonData.length - 1} rows`);
                
                let loadedCount = 0;
                for (let i = 1; i < jsonData.length; i++) {
                    const row = jsonData[i];
                    const name = row[nameIdx];
                    
                    const hasData = row.some(cell => cell !== undefined && cell !== null && cell !== '');
                    if (!hasData) continue;
                    
                    if (name && name.trim()) {
                        const regNo = regIdx >= 0 ? (row[regIdx] || '').toString() : '';
                        const batch = getBatch(regNo);
                        const rank = rankIdx >= 0 ? (row[rankIdx] || 'CDT').toString().trim().toUpperCase() : 'CDT';
                        
                        allCadets.push({
                            name: name.trim(),
                            rank: rank,
                            batch: batch,
                            rowData: row,
                            sheet: sheet,
                            rowIndex: i,
                            headers: headers
                        });
                        loadedCount++;
                    }
                }
                console.log(`Loaded ${loadedCount} cadets from ${file.name}`);
                resolve();
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = () => reject(new Error('Failed to read file'));
        reader.readAsArrayBuffer(file);
    });
}

function getBatch(regNo) {
    const reg = regNo.toUpperCase();
    if (reg.includes('2023') || reg.includes('AP21')) return 'C';
    if (reg.includes('2024')) return 'B2';
    if (reg.includes('2025')) return 'B1';
    return 'C'; // default
}

function displayCadets() {
    // Sort: Batch (C, B2, B1) then Rank
    displayedCadets = [...allCadets].sort((a, b) => {
        const batchCmp = BATCH_ORDER.indexOf(a.batch) - BATCH_ORDER.indexOf(b.batch);
        if (batchCmp !== 0) return batchCmp;
        
        const rankCmp = RANK_ORDER.indexOf(a.rank) - RANK_ORDER.indexOf(b.rank);
        if (rankCmp !== 0) return rankCmp;
        
        return a.name.localeCompare(b.name);
    });
    
    const container = document.getElementById('cadets-list');
    container.innerHTML = '';
    
    BATCH_ORDER.forEach(batch => {
        const batchCadets = displayedCadets.filter(c => c.batch === batch);
        if (batchCadets.length === 0) return;
        
        const batchDiv = document.createElement('div');
        batchDiv.className = 'batch-group';
        
        const batchHeader = document.createElement('div');
        batchHeader.className = 'batch-header';
        batchHeader.textContent = `Batch ${batch}`;
        batchDiv.appendChild(batchHeader);
        
        RANK_ORDER.forEach(rank => {
            const rankCadets = batchCadets.filter(c => c.rank === rank);
            if (rankCadets.length === 0) return;
            
            const rankDiv = document.createElement('div');
            rankDiv.className = 'rank-group';
            
            const rankHeader = document.createElement('div');
            rankHeader.className = 'rank-header';
            rankHeader.textContent = rank;
            rankDiv.appendChild(rankHeader);
            
            rankCadets.forEach((cadet, idx) => {
                const cadetIdx = displayedCadets.indexOf(cadet);
                const item = document.createElement('div');
                item.className = 'cadet-item';
                item.dataset.name = cadet.name.toLowerCase();
                
                item.innerHTML = `
                    <input type="checkbox" data-idx="${cadetIdx}">
                    <span class="cadet-name">${cadet.name}</span>
                    <span class="cadet-rank">${cadet.rank}</span>
                    <span class="cadet-batch">${cadet.batch}</span>
                `;
                
                // Make entire row clickable
                item.addEventListener('click', function(e) {
                    const checkbox = this.querySelector('input[type="checkbox"]');
                    if (e.target !== checkbox) {
                        checkbox.checked = !checkbox.checked;
                    }
                    updateCount();
                });
                
                // Update count when checkbox is directly clicked
                const checkbox = item.querySelector('input[type="checkbox"]');
                checkbox.addEventListener('change', updateCount);
                
                rankDiv.appendChild(item);
            });
            
            batchDiv.appendChild(rankDiv);
        });
        
        container.appendChild(batchDiv);
    });
}

function handleSearch(e) {
    const term = e.target.value.toLowerCase();
    document.querySelectorAll('.cadet-item').forEach(item => {
        const name = item.dataset.name;
        item.classList.toggle('hidden', !name.includes(term));
    });
}

function selectAll() {
    document.querySelectorAll('.cadet-item:not(.hidden) input').forEach(cb => cb.checked = true);
    updateCount();
}

function deselectAll() {
    document.querySelectorAll('.cadet-item:not(.hidden) input').forEach(cb => cb.checked = false);
    updateCount();
}

function generateNominal() {
    const selected = [];
    document.querySelectorAll('.cadet-item input:checked').forEach(cb => {
        const idx = parseInt(cb.dataset.idx);
        selected.push(displayedCadets[idx]);
    });
    
    if (selected.length === 0) {
        alert('Please select at least one cadet!');
        return;
    }
    
    const heading = document.getElementById('heading-input').value.trim();
    const headers = selected[0].headers;
    const slIdx = headers.findIndex(h => h && (h.toLowerCase().includes('sl') || h.toLowerCase().includes('s.no')));
    const dobIdx = headers.findIndex(h => h && (h.toLowerCase().includes('dob') || h.toLowerCase().includes('birth')));
    const enrollIdx = headers.findIndex(h => h && (h.toLowerCase().includes('enroll') || h.toLowerCase().includes('admission')));
    
    const wsData = [];
    const startRow = heading ? 1 : 0;
    
    if (heading) {
        wsData.push([heading]);
    }
    
    wsData.push(headers);
    
    selected.forEach((cadet, idx) => {
        const row = [...cadet.rowData];
        if (slIdx >= 0) row[slIdx] = idx + 1;
        wsData.push(row);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    
    // Merge heading cells
    if (heading) {
        if (!ws['!merges']) ws['!merges'] = [];
        ws['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } });
    }
    
    // Set column widths for better visibility
    ws['!cols'] = headers.map(() => ({ wch: 15 }));
    
    XLSX.utils.book_append_sheet(wb, ws, 'Nominal');
    
    const date = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
    XLSX.writeFile(wb, `Nominal_${date}.xlsx`);
    
    alert(`Generated nominal with ${selected.length} cadets!`);
}
