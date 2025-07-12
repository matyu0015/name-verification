let file1Data = null;
let file2Data = null;
let columnMapping = {
    file1: { lastName: 0, firstName: 1 },
    file2: { fullName: 0 }
};

document.getElementById('file1').addEventListener('change', handleFile1);
document.getElementById('file2').addEventListener('change', handleFile2);
document.getElementById('compareBtn').addEventListener('click', compareNames);

// チェックボックスのイベントリスナー
['matchSchool', 'matchBirthdate', 'matchMobile', 'matchPhone'].forEach(id => {
    document.getElementById(id).addEventListener('change', updateColumnMapping);
});

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
        updateColumnMapping();
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
        updateColumnMapping();
    };
    reader.readAsArrayBuffer(file);
}

function checkFilesLoaded() {
    if (file1Data && file2Data) {
        document.getElementById('compareBtn').disabled = false;
    }
}

function updateColumnMapping() {
    const needsMapping = 
        document.getElementById('matchSchool').checked ||
        document.getElementById('matchBirthdate').checked ||
        document.getElementById('matchMobile').checked ||
        document.getElementById('matchPhone').checked;
    
    const mappingDiv = document.getElementById('columnMapping');
    const mappingInputs = document.getElementById('mappingInputs');
    
    if (needsMapping && file1Data && file2Data) {
        mappingDiv.style.display = 'block';
        
        let html = '<div class="mapping-section">';
        html += '<h4>ファイル①の列番号（A=1, B=2, C=3...）</h4>';
        html += '<div class="mapping-grid">';
        html += '<label>姓の列: <input type="number" id="file1LastName" value="1" min="1"></label>';
        html += '<label>名の列: <input type="number" id="file1FirstName" value="2" min="1"></label>';
        
        if (document.getElementById('matchSchool').checked) {
            html += '<label>学校名の列: <input type="number" id="file1School" min="1"></label>';
        }
        if (document.getElementById('matchBirthdate').checked) {
            html += '<label>生年月日の列: <input type="number" id="file1Birthdate" min="1"></label>';
        }
        if (document.getElementById('matchMobile').checked) {
            html += '<label>携帯番号の列: <input type="number" id="file1Mobile" min="1"></label>';
        }
        if (document.getElementById('matchPhone').checked) {
            html += '<label>固定番号の列: <input type="number" id="file1Phone" min="1"></label>';
        }
        
        html += '</div></div>';
        
        html += '<div class="mapping-section">';
        html += '<h4>ファイル②の列番号</h4>';
        html += '<div class="mapping-grid">';
        html += '<label>姓名の列: <input type="number" id="file2FullName" value="1" min="1"></label>';
        
        if (document.getElementById('matchSchool').checked) {
            html += '<label>学校名の列: <input type="number" id="file2School" min="1"></label>';
        }
        if (document.getElementById('matchBirthdate').checked) {
            html += '<label>生年月日の列: <input type="number" id="file2Birthdate" min="1"></label>';
        }
        if (document.getElementById('matchMobile').checked) {
            html += '<label>携帯番号の列: <input type="number" id="file2Mobile" min="1"></label>';
        }
        if (document.getElementById('matchPhone').checked) {
            html += '<label>固定番号の列: <input type="number" id="file2Phone" min="1"></label>';
        }
        
        html += '</div></div>';
        
        mappingInputs.innerHTML = html;
    } else {
        mappingDiv.style.display = 'none';
    }
}

function getColumnMappings() {
    const mapping = {
        file1: {
            lastName: parseInt(document.getElementById('file1LastName')?.value || 1) - 1,
            firstName: parseInt(document.getElementById('file1FirstName')?.value || 2) - 1
        },
        file2: {
            fullName: parseInt(document.getElementById('file2FullName')?.value || 1) - 1
        }
    };
    
    if (document.getElementById('matchSchool').checked) {
        mapping.file1.school = parseInt(document.getElementById('file1School').value) - 1;
        mapping.file2.school = parseInt(document.getElementById('file2School').value) - 1;
    }
    if (document.getElementById('matchBirthdate').checked) {
        mapping.file1.birthdate = parseInt(document.getElementById('file1Birthdate').value) - 1;
        mapping.file2.birthdate = parseInt(document.getElementById('file2Birthdate').value) - 1;
    }
    if (document.getElementById('matchMobile').checked) {
        mapping.file1.mobile = parseInt(document.getElementById('file1Mobile').value) - 1;
        mapping.file2.mobile = parseInt(document.getElementById('file2Mobile').value) - 1;
    }
    if (document.getElementById('matchPhone').checked) {
        mapping.file1.phone = parseInt(document.getElementById('file1Phone').value) - 1;
        mapping.file2.phone = parseInt(document.getElementById('file2Phone').value) - 1;
    }
    
    return mapping;
}

function normalizeString(str) {
    if (!str) return '';
    return str.toString().trim().replace(/\s+/g, '').toLowerCase();
}

function normalizePhone(phone) {
    if (!phone) return '';
    return phone.toString().replace(/[-\s()]/g, '');
}

function normalizeBirthdate(date) {
    if (!date) return '';
    // 様々な日付形式に対応
    const dateStr = date.toString();
    // スラッシュ、ハイフン、ドットを統一
    return dateStr.replace(/[\/\-\.年月日]/g, '');
}

function compareNames() {
    const mapping = getColumnMappings();
    const results = [];
    const file1Records = new Map();
    
    // 選択された照合項目を取得
    const matchOptions = {
        school: document.getElementById('matchSchool').checked,
        birthdate: document.getElementById('matchBirthdate').checked,
        mobile: document.getElementById('matchMobile').checked,
        phone: document.getElementById('matchPhone').checked
    };
    
    // ファイル1からレコードのマップを作成
    for (let i = 1; i < file1Data.length; i++) { // ヘッダー行をスキップ
        const row = file1Data[i];
        if (row.length > mapping.file1.firstName) {
            const record = {
                lastName: row[mapping.file1.lastName] || '',
                firstName: row[mapping.file1.firstName] || '',
                row: i + 1,
                originalRow: row
            };
            
            const fullName = normalizeString(record.lastName + record.firstName);
            if (fullName) {
                // 追加項目のデータを取得
                if (matchOptions.school && mapping.file1.school !== undefined) {
                    record.school = normalizeString(row[mapping.file1.school] || '');
                }
                if (matchOptions.birthdate && mapping.file1.birthdate !== undefined) {
                    record.birthdate = normalizeBirthdate(row[mapping.file1.birthdate] || '');
                }
                if (matchOptions.mobile && mapping.file1.mobile !== undefined) {
                    record.mobile = normalizePhone(row[mapping.file1.mobile] || '');
                }
                if (matchOptions.phone && mapping.file1.phone !== undefined) {
                    record.phone = normalizePhone(row[mapping.file1.phone] || '');
                }
                
                // 複合キーを作成
                const key = createCompositeKey(fullName, record, matchOptions);
                file1Records.set(key, record);
            }
        }
    }
    
    // ファイル2の各レコードを確認
    for (let i = 1; i < file2Data.length; i++) { // ヘッダー行をスキップ
        const row = file2Data[i];
        if (row.length > mapping.file2.fullName && row[mapping.file2.fullName]) {
            const fullName = row[mapping.file2.fullName].toString();
            const normalizedFullName = normalizeString(fullName);
            
            const record2 = {
                fullName: normalizedFullName
            };
            
            // 追加項目のデータを取得
            if (matchOptions.school && mapping.file2.school !== undefined) {
                record2.school = normalizeString(row[mapping.file2.school] || '');
            }
            if (matchOptions.birthdate && mapping.file2.birthdate !== undefined) {
                record2.birthdate = normalizeBirthdate(row[mapping.file2.birthdate] || '');
            }
            if (matchOptions.mobile && mapping.file2.mobile !== undefined) {
                record2.mobile = normalizePhone(row[mapping.file2.mobile] || '');
            }
            if (matchOptions.phone && mapping.file2.phone !== undefined) {
                record2.phone = normalizePhone(row[mapping.file2.phone] || '');
            }
            
            // 複合キーを作成して検索
            const key = createCompositeKey(normalizedFullName, record2, matchOptions);
            const found = file1Records.has(key);
            const matchInfo = found ? file1Records.get(key) : null;
            
            results.push({
                name: fullName,
                found: found,
                matchInfo: matchInfo,
                row: i + 1,
                record2: record2,
                matchedFields: getMatchedFields(record2, matchInfo, matchOptions)
            });
        }
    }
    
    displayResults(results, matchOptions);
}

function createCompositeKey(name, record, matchOptions) {
    let key = name;
    if (matchOptions.school && record.school) key += '|' + record.school;
    if (matchOptions.birthdate && record.birthdate) key += '|' + record.birthdate;
    if (matchOptions.mobile && record.mobile) key += '|' + record.mobile;
    if (matchOptions.phone && record.phone) key += '|' + record.phone;
    return key;
}

function getMatchedFields(record2, matchInfo, matchOptions) {
    if (!matchInfo) return {};
    
    const matched = {
        name: true // 名前は常に一致している
    };
    
    if (matchOptions.school) {
        matched.school = record2.school === matchInfo.school;
    }
    if (matchOptions.birthdate) {
        matched.birthdate = record2.birthdate === matchInfo.birthdate;
    }
    if (matchOptions.mobile) {
        matched.mobile = record2.mobile === matchInfo.mobile;
    }
    if (matchOptions.phone) {
        matched.phone = record2.phone === matchInfo.phone;
    }
    
    return matched;
}

function displayResults(results, matchOptions) {
    const resultsDiv = document.getElementById('results');
    
    let html = '<h2>照合結果</h2>';
    html += '<div class="summary">';
    
    const foundCount = results.filter(r => r.found).length;
    const notFoundCount = results.filter(r => !r.found).length;
    
    html += `<p>照合完了: 全${results.length}件</p>`;
    html += `<p class="found">見つかった: ${foundCount}件</p>`;
    html += `<p class="not-found">見つからなかった: ${notFoundCount}件</p>`;
    
    // 照合条件の表示
    const conditions = ['氏名（必須）'];
    if (matchOptions.school) conditions.push('学校名');
    if (matchOptions.birthdate) conditions.push('生年月日');
    if (matchOptions.mobile) conditions.push('携帯番号');
    if (matchOptions.phone) conditions.push('固定番号');
    html += `<p>照合条件: ${conditions.join('、')}</p>`;
    
    html += '</div>';
    
    html += '<table>';
    html += '<thead><tr>';
    html += '<th>ファイル②の名前</th>';
    html += '<th>行番号</th>';
    html += '<th>結果</th>';
    html += '<th>照合詳細</th>';
    html += '<th>ファイル①での位置</th>';
    html += '</tr></thead>';
    html += '<tbody>';
    
    results.forEach(result => {
        const rowClass = result.found ? 'found-row' : 'not-found-row';
        html += `<tr class="${rowClass}">`;
        html += `<td>${result.name}</td>`;
        html += `<td>${result.row}</td>`;
        html += `<td>${result.found ? '○' : '×'}</td>`;
        
        // 照合詳細
        html += '<td>';
        if (result.found) {
            const details = [];
            if (result.matchedFields.name) details.push('氏名: ○');
            if (matchOptions.school) details.push(`学校名: ${result.matchedFields.school ? '○' : '×'}`);
            if (matchOptions.birthdate) details.push(`生年月日: ${result.matchedFields.birthdate ? '○' : '×'}`);
            if (matchOptions.mobile) details.push(`携帯: ${result.matchedFields.mobile ? '○' : '×'}`);
            if (matchOptions.phone) details.push(`固定: ${result.matchedFields.phone ? '○' : '×'}`);
            html += details.join(' / ');
        } else {
            html += '-';
        }
        html += '</td>';
        
        html += `<td>${result.found ? `${result.matchInfo.row}行目（${result.matchInfo.lastName} ${result.matchInfo.firstName}）` : '-'}</td>`;
        html += '</tr>';
    });
    
    html += '</tbody></table>';
    
    resultsDiv.innerHTML = html;
}