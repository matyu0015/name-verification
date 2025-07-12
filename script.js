let file1Data = null;
let file2Data = null;

document.getElementById('file1').addEventListener('change', handleFile1);
document.getElementById('file2').addEventListener('change', handleFile2);
document.getElementById('compareBtn').addEventListener('click', compareNames);

function handleFile1(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        file1Data = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        checkFilesLoaded();
    };
    reader.readAsArrayBuffer(file);
}

function handleFile2(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        file2Data = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        checkFilesLoaded();
    };
    reader.readAsArrayBuffer(file);
}

function checkFilesLoaded() {
    if (file1Data && file2Data) {
        document.getElementById('compareBtn').disabled = false;
    }
}

function normalizeString(str) {
    if (!str) return '';
    return str.toString().trim().replace(/\s+/g, '').toLowerCase();
}

function compareNames() {
    const results = [];
    const file1Names = new Map();
    
    // ファイル1から名前のマップを作成（姓と名を結合）
    for (let i = 0; i < file1Data.length; i++) {
        const row = file1Data[i];
        if (row.length >= 2) {
            const lastName = normalizeString(row[0]);
            const firstName = normalizeString(row[1]);
            const fullName = lastName + firstName;
            if (fullName) {
                file1Names.set(fullName, {
                    lastName: row[0],
                    firstName: row[1],
                    row: i + 1
                });
            }
        }
    }
    
    // ファイル2の各名前を確認
    for (let i = 0; i < file2Data.length; i++) {
        const row = file2Data[i];
        if (row.length >= 1 && row[0]) {
            const fullName = row[0].toString();
            const normalizedFullName = normalizeString(fullName);
            
            const found = file1Names.has(normalizedFullName);
            const matchInfo = found ? file1Names.get(normalizedFullName) : null;
            
            results.push({
                name: fullName,
                found: found,
                matchInfo: matchInfo,
                row: i + 1
            });
        }
    }
    
    displayResults(results);
}

function displayResults(results) {
    const resultsDiv = document.getElementById('results');
    
    let html = '<h2>照合結果</h2>';
    html += '<div class="summary">';
    
    const foundCount = results.filter(r => r.found).length;
    const notFoundCount = results.filter(r => !r.found).length;
    
    html += `<p>照合完了: 全${results.length}件</p>`;
    html += `<p class="found">見つかった: ${foundCount}件</p>`;
    html += `<p class="not-found">見つからなかった: ${notFoundCount}件</p>`;
    html += '</div>';
    
    html += '<table>';
    html += '<thead><tr><th>ファイル②の名前</th><th>行番号</th><th>結果</th><th>ファイル①での位置</th></tr></thead>';
    html += '<tbody>';
    
    results.forEach(result => {
        const rowClass = result.found ? 'found-row' : 'not-found-row';
        html += `<tr class="${rowClass}">`;
        html += `<td>${result.name}</td>`;
        html += `<td>${result.row}</td>`;
        html += `<td>${result.found ? '○' : '×'}</td>`;
        html += `<td>${result.found ? `${result.matchInfo.row}行目（${result.matchInfo.lastName} ${result.matchInfo.firstName}）` : '-'}</td>`;
        html += '</tr>';
    });
    
    html += '</tbody></table>';
    
    resultsDiv.innerHTML = html;
}