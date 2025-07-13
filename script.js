let file1Data = null;
let file2Data = null;
let columnMapping = {
    file1: { lastName: 0, firstName: 1 },
    file2: { fullName: 0 }
};
let extractColumns = [];
let lastResults = [];
let lastValidExtractColumns = [];

// 列番号をアルファベットに変換
function numberToColumn(num) {
    let column = '';
    while (num > 0) {
        let remainder = (num - 1) % 26;
        column = String.fromCharCode(65 + remainder) + column;
        num = Math.floor((num - 1) / 26);
    }
    return column;
}

// アルファベットを列番号に変換
function columnToNumber(column) {
    let num = 0;
    for (let i = 0; i < column.length; i++) {
        num = num * 26 + (column.charCodeAt(i) - 64);
    }
    return num;
}

document.getElementById('file1').addEventListener('change', handleFile1);
document.getElementById('file2').addEventListener('change', handleFile2);
document.getElementById('compareBtn').addEventListener('click', compareNames);

// チェックボックスのイベントリスナー
['matchKana', 'matchSchool', 'matchBirthdate', 'matchMobile', 'matchEmail'].forEach(id => {
    document.getElementById(id).addEventListener('change', updateColumnMapping);
});

// ラジオボタンのイベントリスナー
document.querySelectorAll('input[name="file2Format"]').forEach(radio => {
    radio.addEventListener('change', updateColumnMapping);
});

// 抽出列の追加ボタン
document.getElementById('addExtractColumn').addEventListener('click', addExtractColumn);

// 抽出列管理関数
function initializeExtractColumns() {
    extractColumns = [{ column: '', name: '' }];
    renderExtractColumns();
}

function addExtractColumn() {
    if (extractColumns.length < 5) {
        extractColumns.push({ column: '', name: '' });
        renderExtractColumns();
    }
}

function removeExtractColumn(index) {
    if (extractColumns.length > 1) {
        extractColumns.splice(index, 1);
        renderExtractColumns();
    }
}

function renderExtractColumns() {
    const container = document.getElementById('extractInputs');
    container.innerHTML = extractColumns.map((col, index) => `
        <div class="extract-item">
            <label>ファイル①の列:</label>
            <input type="text" 
                   value="${col.column}" 
                   placeholder="例: F" 
                   maxlength="3"
                   onchange="updateExtractColumn(${index}, 'column', this.value)">
            <label>項目名:</label>
            <input type="text" 
                   class="column-name"
                   value="${col.name}" 
                   placeholder="例: 住所" 
                   onchange="updateExtractColumn(${index}, 'name', this.value)">
            ${extractColumns.length > 1 ? 
                `<button type="button" onclick="removeExtractColumn(${index})">削除</button>` : 
                ''}
        </div>
    `).join('');
    
    // 最大5列まで
    document.getElementById('addExtractColumn').disabled = extractColumns.length >= 5;
}

function updateExtractColumn(index, field, value) {
    extractColumns[index][field] = value;
}

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
        updateColumnMapping();
        // 抽出列セクションを表示
        document.getElementById('extractColumns').style.display = 'block';
        initializeExtractColumns();
    }
}

function updateColumnMapping() {
    try {
        const needsMapping = 
            document.getElementById('matchKana').checked ||
            document.getElementById('matchSchool').checked ||
            document.getElementById('matchBirthdate').checked ||
            document.getElementById('matchMobile').checked ||
            document.getElementById('matchEmail').checked;
        
        const mappingDiv = document.getElementById('columnMapping');
        const mappingInputs = document.getElementById('mappingInputs');
        
        if (!mappingDiv || !mappingInputs) return;
    
    if (needsMapping && file1Data && file2Data) {
        mappingDiv.style.display = 'block';
        
        let html = '<div class="mapping-section">';
        html += '<h4>ファイル①の列を指定（A, B, C...）</h4>';
        html += '<div class="mapping-grid">';
        html += '<label>姓の列 (例: A): <input type="text" id="file1LastName" value="A" placeholder="A"></label>';
        html += '<label>名の列 (例: B): <input type="text" id="file1FirstName" value="B" placeholder="B"></label>';
        
        if (document.getElementById('matchKana').checked) {
            html += '<label>カナ姓の列 (例: C): <input type="text" id="file1KanaLastName" placeholder="C"></label>';
            html += '<label>カナ名の列 (例: D): <input type="text" id="file1KanaFirstName" placeholder="D"></label>';
        }
        if (document.getElementById('matchSchool').checked) {
            html += '<label>学校名の列 (例: E): <input type="text" id="file1School" placeholder="E"></label>';
        }
        if (document.getElementById('matchBirthdate').checked) {
            html += '<label>生年月日の列 (例: F): <input type="text" id="file1Birthdate" placeholder="F"></label>';
        }
        if (document.getElementById('matchMobile').checked) {
            html += '<label>携帯番号の列 (例: G): <input type="text" id="file1Mobile" placeholder="G"></label>';
        }
        if (document.getElementById('matchEmail').checked) {
            html += '<label>メールアドレスの列 (例: H): <input type="text" id="file1Email" placeholder="H"></label>';
        }
        
        html += '</div></div>';
        
        html += '<div class="mapping-section">';
        html += '<h4>ファイル②の列を指定（A, B, C...）</h4>';
        html += '<div class="mapping-grid">';
        
        const file2Format = document.querySelector('input[name="file2Format"]:checked').value;
        if (file2Format === 'combined') {
            html += '<label>姓名の列 (例: A): <input type="text" id="file2FullName" value="A" placeholder="A"></label>';
        } else {
            html += '<label>姓の列 (例: A): <input type="text" id="file2LastName" value="A" placeholder="A"></label>';
            html += '<label>名の列 (例: B): <input type="text" id="file2FirstName" value="B" placeholder="B"></label>';
        }
        
        if (document.getElementById('matchKana').checked) {
            if (file2Format === 'combined') {
                html += '<label>カナ姓名の列 (例: B): <input type="text" id="file2KanaFullName" placeholder="B"></label>';
            } else {
                html += '<label>カナ姓の列 (例: C): <input type="text" id="file2KanaLastName" placeholder="C"></label>';
                html += '<label>カナ名の列 (例: D): <input type="text" id="file2KanaFirstName" placeholder="D"></label>';
            }
        }
        if (document.getElementById('matchSchool').checked) {
            html += '<label>学校名の列 (例: E): <input type="text" id="file2School" placeholder="E"></label>';
        }
        if (document.getElementById('matchBirthdate').checked) {
            html += '<label>生年月日の列 (例: F): <input type="text" id="file2Birthdate" placeholder="F"></label>';
        }
        if (document.getElementById('matchMobile').checked) {
            html += '<label>携帯番号の列 (例: G): <input type="text" id="file2Mobile" placeholder="G"></label>';
        }
        if (document.getElementById('matchEmail').checked) {
            html += '<label>メールアドレスの列 (例: H): <input type="text" id="file2Email" placeholder="H"></label>';
        }
        
        html += '</div></div>';
        
        mappingInputs.innerHTML = html;
    } else {
        mappingDiv.style.display = 'none';
    }
    } catch (error) {
        console.error('Error in updateColumnMapping:', error);
    }
}

function getColumnMappings() {
    const file2Format = document.querySelector('input[name="file2Format"]:checked').value;
    
    const mapping = {
        file1: {
            lastName: columnToNumber(document.getElementById('file1LastName')?.value || 'A') - 1,
            firstName: columnToNumber(document.getElementById('file1FirstName')?.value || 'B') - 1
        },
        file2: {}
    };
    
    if (file2Format === 'combined') {
        mapping.file2.fullName = columnToNumber(document.getElementById('file2FullName')?.value || 'A') - 1;
    } else {
        mapping.file2.lastName = columnToNumber(document.getElementById('file2LastName')?.value || 'A') - 1;
        mapping.file2.firstName = columnToNumber(document.getElementById('file2FirstName')?.value || 'B') - 1;
    }
    
    if (document.getElementById('matchKana').checked) {
        mapping.file1.kanaLastName = columnToNumber(document.getElementById('file1KanaLastName')?.value || 'C') - 1;
        mapping.file1.kanaFirstName = columnToNumber(document.getElementById('file1KanaFirstName')?.value || 'D') - 1;
        
        if (file2Format === 'combined') {
            mapping.file2.kanaFullName = columnToNumber(document.getElementById('file2KanaFullName')?.value || 'B') - 1;
        } else {
            mapping.file2.kanaLastName = columnToNumber(document.getElementById('file2KanaLastName')?.value || 'C') - 1;
            mapping.file2.kanaFirstName = columnToNumber(document.getElementById('file2KanaFirstName')?.value || 'D') - 1;
        }
    }
    if (document.getElementById('matchSchool').checked) {
        mapping.file1.school = columnToNumber(document.getElementById('file1School').value) - 1;
        mapping.file2.school = columnToNumber(document.getElementById('file2School').value) - 1;
    }
    if (document.getElementById('matchBirthdate').checked) {
        mapping.file1.birthdate = columnToNumber(document.getElementById('file1Birthdate').value) - 1;
        mapping.file2.birthdate = columnToNumber(document.getElementById('file2Birthdate').value) - 1;
    }
    if (document.getElementById('matchMobile').checked) {
        mapping.file1.mobile = columnToNumber(document.getElementById('file1Mobile').value) - 1;
        mapping.file2.mobile = columnToNumber(document.getElementById('file2Mobile').value) - 1;
    }
    if (document.getElementById('matchEmail').checked) {
        mapping.file1.email = columnToNumber(document.getElementById('file1Email').value) - 1;
        mapping.file2.email = columnToNumber(document.getElementById('file2Email').value) - 1;
    }
    
    return mapping;
}

function normalizeString(str) {
    if (!str) return '';
    return str.toString().trim().replace(/\s+/g, '').toLowerCase();
}

function normalizeKana(str) {
    if (!str) return '';
    // カタカナをヒラガナに統一して正規化
    return str.toString().trim().replace(/\s+/g, '').toLowerCase()
        .replace(/[ァ-ヶ]/g, function(match) {
            const chr = match.charCodeAt(0) - 0x60;
            return String.fromCharCode(chr);
        });
}

function normalizePhone(phone) {
    if (!phone) return '';
    return phone.toString().replace(/[-\s()]/g, '');
}

function normalizeEmail(email) {
    if (!email) return '';
    return email.toString().trim().toLowerCase();
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
        kana: document.getElementById('matchKana').checked,
        school: document.getElementById('matchSchool').checked,
        birthdate: document.getElementById('matchBirthdate').checked,
        mobile: document.getElementById('matchMobile').checked,
        email: document.getElementById('matchEmail').checked
    };
    
    const file2Format = document.querySelector('input[name="file2Format"]:checked').value;
    
    // 有効な抽出列を取得
    const validExtractColumns = extractColumns.filter(col => 
        col.column && col.name && columnToNumber(col.column) > 0
    );
    
    // ファイル1からレコードのマップを作成（完全一致用と部分一致用）
    const file1RecordsByName = new Map(); // 名前のみでの検索用
    
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
                if (matchOptions.kana && mapping.file1.kanaLastName !== undefined && mapping.file1.kanaFirstName !== undefined) {
                    record.kanaLastName = row[mapping.file1.kanaLastName] || '';
                    record.kanaFirstName = row[mapping.file1.kanaFirstName] || '';
                    record.kanaFullName = normalizeKana(record.kanaLastName + record.kanaFirstName);
                }
                if (matchOptions.school && mapping.file1.school !== undefined) {
                    record.school = normalizeString(row[mapping.file1.school] || '');
                }
                if (matchOptions.birthdate && mapping.file1.birthdate !== undefined) {
                    record.birthdate = normalizeBirthdate(row[mapping.file1.birthdate] || '');
                }
                if (matchOptions.mobile && mapping.file1.mobile !== undefined) {
                    record.mobile = normalizePhone(row[mapping.file1.mobile] || '');
                }
                if (matchOptions.email && mapping.file1.email !== undefined) {
                    record.email = normalizeEmail(row[mapping.file1.email] || '');
                }
                
                // 抽出列のデータを取得
                record.extractData = {};
                validExtractColumns.forEach(extractCol => {
                    const colIndex = columnToNumber(extractCol.column) - 1;
                    if (colIndex >= 0 && colIndex < row.length) {
                        record.extractData[extractCol.name] = row[colIndex] || '';
                    }
                });
                
                // 複合キーを作成（完全一致用）
                const key = createCompositeKey(fullName, record, matchOptions);
                file1Records.set(key, record);
                
                // 名前のみでも検索できるように保存
                if (!file1RecordsByName.has(fullName)) {
                    file1RecordsByName.set(fullName, []);
                }
                file1RecordsByName.get(fullName).push(record);
            }
        }
    }
    
    // ファイル2の各レコードを確認
    for (let i = 1; i < file2Data.length; i++) { // ヘッダー行をスキップ
        const row = file2Data[i];
        let fullName, normalizedFullName;
        
        if (file2Format === 'combined') {
            if (row.length > mapping.file2.fullName && row[mapping.file2.fullName]) {
                fullName = row[mapping.file2.fullName].toString();
                normalizedFullName = normalizeString(fullName);
            } else {
                continue;
            }
        } else {
            if (row.length > mapping.file2.firstName) {
                const lastName = row[mapping.file2.lastName] || '';
                const firstName = row[mapping.file2.firstName] || '';
                fullName = lastName + ' ' + firstName;
                normalizedFullName = normalizeString(lastName + firstName);
            } else {
                continue;
            }
        }
        
        const record2 = {
            fullName: normalizedFullName
        };
        
        // 追加項目のデータを取得
        if (matchOptions.kana) {
            if (file2Format === 'combined' && mapping.file2.kanaFullName !== undefined) {
                record2.kanaFullName = normalizeKana(row[mapping.file2.kanaFullName] || '');
            } else if (file2Format === 'separated' && mapping.file2.kanaLastName !== undefined && mapping.file2.kanaFirstName !== undefined) {
                const kanaLastName = row[mapping.file2.kanaLastName] || '';
                const kanaFirstName = row[mapping.file2.kanaFirstName] || '';
                record2.kanaFullName = normalizeKana(kanaLastName + kanaFirstName);
            }
        }
        if (matchOptions.school && mapping.file2.school !== undefined) {
            record2.school = normalizeString(row[mapping.file2.school] || '');
        }
        if (matchOptions.birthdate && mapping.file2.birthdate !== undefined) {
            record2.birthdate = normalizeBirthdate(row[mapping.file2.birthdate] || '');
        }
        if (matchOptions.mobile && mapping.file2.mobile !== undefined) {
            record2.mobile = normalizePhone(row[mapping.file2.mobile] || '');
        }
        if (matchOptions.email && mapping.file2.email !== undefined) {
            record2.email = normalizeEmail(row[mapping.file2.email] || '');
        }
        
        // 複合キーを作成して完全一致を検索
        const key = createCompositeKey(normalizedFullName, record2, matchOptions);
        const perfectMatch = file1Records.has(key);
        let matchInfo = perfectMatch ? file1Records.get(key) : null;
        let matchedFields = {};
        let matchScore = 0;
        
        // 完全一致が見つからない場合、部分一致を検索
        if (!perfectMatch && file1RecordsByName.has(normalizedFullName)) {
            const candidates = file1RecordsByName.get(normalizedFullName);
            let bestMatch = null;
            let bestScore = 0;
            
            for (const candidate of candidates) {
                const fields = getMatchedFields(record2, candidate, matchOptions);
                const score = calculateMatchScore(fields, matchOptions);
                
                if (score > bestScore) {
                    bestScore = score;
                    bestMatch = candidate;
                    matchedFields = fields;
                }
            }
            
            if (bestMatch) {
                matchInfo = bestMatch;
                matchScore = bestScore;
            }
        } else if (perfectMatch) {
            matchedFields = getMatchedFields(record2, matchInfo, matchOptions);
            matchScore = calculateMatchScore(matchedFields, matchOptions);
        }
        
        results.push({
            name: fullName,
            found: perfectMatch,
            partialMatch: !perfectMatch && matchInfo !== null,
            matchInfo: matchInfo,
            row: i + 1,
            record2: record2,
            matchedFields: matchedFields,
            matchScore: matchScore,
            extractData: matchInfo ? matchInfo.extractData : {}
        });
    }
    
    displayResults(results, matchOptions, validExtractColumns);
}

function createCompositeKey(name, record, matchOptions) {
    let key = name;
    if (matchOptions.kana && record.kanaFullName) key += '|' + record.kanaFullName;
    if (matchOptions.school && record.school) key += '|' + record.school;
    if (matchOptions.birthdate && record.birthdate) key += '|' + record.birthdate;
    if (matchOptions.mobile && record.mobile) key += '|' + record.mobile;
    if (matchOptions.email && record.email) key += '|' + record.email;
    return key;
}

function getMatchedFields(record2, matchInfo, matchOptions) {
    if (!matchInfo) return {};
    
    const matched = {
        name: true // 名前は常に一致している
    };
    
    if (matchOptions.kana) {
        matched.kana = record2.kanaFullName === matchInfo.kanaFullName;
    }
    if (matchOptions.school) {
        matched.school = record2.school === matchInfo.school;
    }
    if (matchOptions.birthdate) {
        matched.birthdate = record2.birthdate === matchInfo.birthdate;
    }
    if (matchOptions.mobile) {
        matched.mobile = record2.mobile === matchInfo.mobile;
    }
    if (matchOptions.email) {
        matched.email = record2.email === matchInfo.email;
    }
    
    return matched;
}

function calculateMatchScore(matchedFields, matchOptions) {
    let totalFields = 1; // 名前は必須
    let matchedCount = matchedFields.name ? 1 : 0;
    
    if (matchOptions.kana) {
        totalFields++;
        if (matchedFields.kana) matchedCount++;
    }
    if (matchOptions.school) {
        totalFields++;
        if (matchedFields.school) matchedCount++;
    }
    if (matchOptions.birthdate) {
        totalFields++;
        if (matchedFields.birthdate) matchedCount++;
    }
    if (matchOptions.mobile) {
        totalFields++;
        if (matchedFields.mobile) matchedCount++;
    }
    if (matchOptions.email) {
        totalFields++;
        if (matchedFields.email) matchedCount++;
    }
    
    return (matchedCount / totalFields) * 100;
}

function displayResults(results, matchOptions, validExtractColumns) {
    // 結果を保存（エクスポート用）
    lastResults = results;
    lastValidExtractColumns = validExtractColumns;
    
    const resultsDiv = document.getElementById('results');
    
    let html = '<h2>照合結果</h2>';
    html += '<div class="summary">';
    
    const foundCount = results.filter(r => r.found).length;
    const partialCount = results.filter(r => r.partialMatch).length;
    const notFoundCount = results.filter(r => !r.found && !r.partialMatch).length;
    
    html += `<p>照合完了: 全${results.length}件</p>`;
    html += `<p class="found">完全一致: ${foundCount}件</p>`;
    html += `<p class="partial">部分一致: ${partialCount}件</p>`;
    html += `<p class="not-found">見つからなかった: ${notFoundCount}件</p>`;
    
    // 照合条件の表示
    const conditions = ['氏名（必須）'];
    if (matchOptions.kana) conditions.push('カナ氏名');
    if (matchOptions.school) conditions.push('学校名');
    if (matchOptions.birthdate) conditions.push('生年月日');
    if (matchOptions.mobile) conditions.push('携帯番号');
    if (matchOptions.email) conditions.push('メールアドレス');
    html += `<p>照合条件: ${conditions.join('、')}</p>`;
    
    html += '</div>';
    
    // エクスポートボタンを追加
    html += '<div class="export-section">';
    html += '<button type="button" onclick="exportResults()">結果をCSVでエクスポート</button>';
    html += '</div>';
    
    html += '<table>';
    html += '<thead><tr>';
    html += '<th>ファイル②の名前</th>';
    html += '<th>行番号</th>';
    html += '<th>結果</th>';
    html += '<th>一致度</th>';
    html += '<th>照合詳細</th>';
    html += '<th>ファイル①での位置</th>';
    
    // 抽出列のヘッダーを追加
    validExtractColumns.forEach(col => {
        html += `<th>${col.name}</th>`;
    });
    
    html += '</tr></thead>';
    html += '<tbody>';
    
    results.forEach(result => {
        let rowClass = 'not-found-row';
        if (result.found) {
            rowClass = 'found-row';
        } else if (result.partialMatch) {
            // 一致度に応じてクラスを変更
            if (result.matchScore >= 75) {
                rowClass = 'partial-high-row';
            } else if (result.matchScore >= 50) {
                rowClass = 'partial-medium-row';
            } else {
                rowClass = 'partial-low-row';
            }
        }
        
        html += `<tr class="${rowClass}">`;
        html += `<td>${result.name}</td>`;
        html += `<td>${result.row}</td>`;
        
        // 結果列
        html += '<td>';
        if (result.found) {
            html += '<span class="result-icon perfect">◎</span>';
        } else if (result.partialMatch) {
            html += '<span class="result-icon partial">△</span>';
        } else {
            html += '<span class="result-icon none">×</span>';
        }
        html += '</td>';
        
        // 一致度
        html += '<td>';
        if (result.found || result.partialMatch) {
            const scoreClass = result.matchScore >= 75 ? 'high' : result.matchScore >= 50 ? 'medium' : 'low';
            html += `<div class="match-score ${scoreClass}">`;
            html += `<div class="score-bar" style="width: ${result.matchScore}%"></div>`;
            html += `<span class="score-text">${Math.round(result.matchScore)}%</span>`;
            html += '</div>';
        } else {
            html += '-';
        }
        html += '</td>';
        
        // 照合詳細
        html += '<td>';
        if (result.found || result.partialMatch) {
            const details = [];
            if (result.matchedFields.name) {
                details.push('<span class="field-match">氏名: ○</span>');
            }
            if (matchOptions.kana) {
                details.push(result.matchedFields.kana ? 
                    '<span class="field-match">カナ: ○</span>' : 
                    '<span class="field-nomatch">カナ: ×</span>');
            }
            if (matchOptions.school) {
                details.push(result.matchedFields.school ? 
                    '<span class="field-match">学校名: ○</span>' : 
                    '<span class="field-nomatch">学校名: ×</span>');
            }
            if (matchOptions.birthdate) {
                details.push(result.matchedFields.birthdate ? 
                    '<span class="field-match">生年月日: ○</span>' : 
                    '<span class="field-nomatch">生年月日: ×</span>');
            }
            if (matchOptions.mobile) {
                details.push(result.matchedFields.mobile ? 
                    '<span class="field-match">携帯: ○</span>' : 
                    '<span class="field-nomatch">携帯: ×</span>');
            }
            if (matchOptions.email) {
                details.push(result.matchedFields.email ? 
                    '<span class="field-match">メール: ○</span>' : 
                    '<span class="field-nomatch">メール: ×</span>');
            }
            html += details.join(' / ');
        } else {
            html += '-';
        }
        html += '</td>';
        
        html += `<td>${result.matchInfo ? `${result.matchInfo.row}行目（${result.matchInfo.lastName} ${result.matchInfo.firstName}）` : '-'}</td>`;
        
        // 抽出列のデータを追加
        validExtractColumns.forEach(col => {
            if (result.found || result.partialMatch) {
                const extractValue = result.extractData[col.name] || '-';
                html += `<td>${extractValue}</td>`;
            } else {
                html += '<td>-</td>';
            }
        });
        
        html += '</tr>';
    });
    
    html += '</tbody></table>';
    
    resultsDiv.innerHTML = html;
}

// CSVエクスポート関数
function exportResults() {
    if (!lastResults || lastResults.length === 0) {
        alert('エクスポートする結果がありません。');
        return;
    }
    
    // CSVヘッダーを作成
    const headers = ['ファイル②の名前', '行番号', '結果', '一致度(%)', '照合詳細', 'ファイル①での位置'];
    
    // 抽出列のヘッダーを追加
    lastValidExtractColumns.forEach(col => {
        headers.push(col.name);
    });
    
    // CSVデータを作成
    const csvData = [headers];
    
    lastResults.forEach(result => {
        const row = [];
        
        // 基本情報
        row.push(result.name);
        row.push(result.row);
        
        // 結果
        if (result.found) {
            row.push('完全一致');
        } else if (result.partialMatch) {
            row.push('部分一致');
        } else {
            row.push('不一致');
        }
        
        // 一致度
        if (result.found || result.partialMatch) {
            row.push(Math.round(result.matchScore));
        } else {
            row.push('-');
        }
        
        // 照合詳細
        if (result.found || result.partialMatch) {
            const details = [];
            if (result.matchedFields.name) details.push('氏名:○');
            if (result.matchedFields.kana !== undefined) {
                details.push(result.matchedFields.kana ? 'カナ:○' : 'カナ:×');
            }
            if (result.matchedFields.school !== undefined) {
                details.push(result.matchedFields.school ? '学校名:○' : '学校名:×');
            }
            if (result.matchedFields.birthdate !== undefined) {
                details.push(result.matchedFields.birthdate ? '生年月日:○' : '生年月日:×');
            }
            if (result.matchedFields.mobile !== undefined) {
                details.push(result.matchedFields.mobile ? '携帯:○' : '携帯:×');
            }
            if (result.matchedFields.email !== undefined) {
                details.push(result.matchedFields.email ? 'メール:○' : 'メール:×');
            }
            row.push(details.join(' / '));
        } else {
            row.push('-');
        }
        
        // ファイル①での位置
        if (result.matchInfo) {
            row.push(`${result.matchInfo.row}行目（${result.matchInfo.lastName} ${result.matchInfo.firstName}）`);
        } else {
            row.push('-');
        }
        
        // 抽出列のデータ
        lastValidExtractColumns.forEach(col => {
            if (result.found || result.partialMatch) {
                row.push(result.extractData[col.name] || '-');
            } else {
                row.push('-');
            }
        });
        
        csvData.push(row);
    });
    
    // CSVファイルを作成
    const csvContent = csvData.map(row => 
        row.map(cell => {
            // セル内の値にカンマ、改行、ダブルクォートが含まれる場合の処理
            const cellStr = String(cell);
            if (cellStr.includes(',') || cellStr.includes('\n') || cellStr.includes('"')) {
                return '"' + cellStr.replace(/"/g, '""') + '"';
            }
            return cellStr;
        }).join(',')
    ).join('\n');
    
    // BOMを付けてUTF-8として保存（Excelで文字化けを防ぐ）
    const bom = '\uFEFF';
    const blob = new Blob([bom + csvContent], { type: 'text/csv;charset=utf-8;' });
    
    // ダウンロード
    const link = document.createElement('a');
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
    link.download = `照合結果_${timestamp}.csv`;
    link.href = URL.createObjectURL(blob);
    link.click();
    
    // メモリ解放
    URL.revokeObjectURL(link.href);
}