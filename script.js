/**
 * 水质类别自动判定系统 v2.3 (排序优化版)
 * 更新内容：超标因子排序逻辑（pH/DO置顶 + 倍数降序）
 */

// 1. 配置
const HEADER_KEYWORDS = {
    site: ['监测断面', '断面名称', '断面', '监测点位', '点位名称', '点位', '测站', '名称'],
    time: ['采样时间', '监测时间', '日期', '时间', '采样日期'],
    year: ['年', '年份'],
    month: ['月', '月份'],
    day: ['日', '日期'], 
    ph: ['ph', 'ph值', '酸碱度'],
    do: ['溶解氧', 'do'],
    cod_mn: ['高锰酸盐指数', 'codmn', 'imn', '高锰酸盐'],
    cod: ['化学需氧量', 'cod', 'codcr'],
    bod5: ['五日生化需氧量', 'bod5', 'bod'],
    nh3_n: ['氨氮', 'nh3-n', 'nh3n', 'nh3'],
    tp: ['总磷', 'tp']
};

const STANDARDS = {
    ph: { name: 'pH值', limits: [6, 9], type: 'range' },
    do: { name: '溶解氧', limits: [7.5, 6, 5, 3, 2], type: 'desc' },
    cod_mn: { name: '高锰酸盐指数', limits: [2, 4, 6, 10, 15], type: 'asc' },
    cod: { name: '化学需氧量', limits: [15, 15, 20, 30, 40], type: 'asc' },
    bod5: { name: '五日生化需氧量', limits: [3, 3, 4, 6, 10], type: 'asc' },
    nh3_n: { name: '氨氮', limits: [0.15, 0.5, 1.0, 1.5, 2.0], type: 'asc' },
    tp: { name: '总磷', limits: { river: [0.02, 0.1, 0.2, 0.3, 0.4], lake: [0.005, 0.025, 0.05, 0.1, 0.2] }, type: 'asc' }
};

const GRADES = ['Ⅰ', 'Ⅱ', 'Ⅲ', 'Ⅳ', 'Ⅴ', '劣Ⅴ'];
const GRADE_CLASSES = ['grade-I', 'grade-II', 'grade-III', 'grade-IV', 'grade-V', 'grade-VI'];

let currentData = [];

// 2. 初始化
document.addEventListener('DOMContentLoaded', () => {
    initStandardTable();
    setupEventListeners();
});

function setupEventListeners() {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');

    dropZone.onclick = () => fileInput.click();
    dropZone.ondragover = (e) => { e.preventDefault(); dropZone.classList.add('dragover'); };
    dropZone.ondragleave = () => dropZone.classList.remove('dragover');
    dropZone.ondrop = (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        if(e.dataTransfer.files.length) readExcel(e.dataTransfer.files[0]);
    };
    fileInput.onchange = (e) => { if(e.target.files.length) readExcel(e.target.files[0]); };

    document.getElementById('btnAnalyze').onclick = startAnalysis;
    document.getElementById('btnReset').onclick = () => location.reload();
    document.getElementById('btnTemplate').onclick = downloadTemplate;
    document.getElementById('btnExport').onclick = exportResults;
}

// 3. 数据处理逻辑
function readExcel(file) {
    document.getElementById('fileInfo').innerText = `已选择: ${file.name}`;
    document.getElementById('fileInfo').classList.remove('hidden');
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            processRawDataFuzzy(json);
        } catch (err) {
            showError('文件解析失败');
        }
    };
    reader.readAsArrayBuffer(file);
}

function processRawDataFuzzy(rows) {
    if (!rows || rows.length < 1) return showError('文件无内容');
    
    let headerIndex = 0;
    let colIdx = {};
    const indicatorKeys = Object.keys(STANDARDS);

    const isRowHeader = (row) => {
        if (!row) return false;
        const rowStr = row.join('').toLowerCase();
        let matchCount = 0;
        indicatorKeys.forEach(k => {
            if (HEADER_KEYWORDS[k].some(key => rowStr.includes(key))) matchCount++;
        });
        return matchCount >= 3; 
    };

    if (!isRowHeader(rows[0]) && rows.length > 1 && isRowHeader(rows[1])) {
        headerIndex = 1;
        console.log("检测到第一行可能为标题，已自动跳过，将第二行识别为表头");
    }

    const headers = rows[headerIndex].map(h => (h || '').toString().trim().toLowerCase());

    colIdx.site = findColIndex(headers, HEADER_KEYWORDS.site, ['河流名称']);
    indicatorKeys.forEach(k => colIdx[k] = findColIndex(headers, HEADER_KEYWORDS[k]));
    ['time', 'year', 'month', 'day'].forEach(k => colIdx[k] = findColIndex(headers, HEADER_KEYWORDS[k]));

    currentData = [];
    for (let i = headerIndex + 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0 || row.every(c => c === "")) continue;

        let rowObj = { raw: {} };
        rowObj.site = (colIdx.site !== -1 && row[colIdx.site]) ? row[colIdx.site] : `未知点位-${i}`;
        rowObj.time = parseRowTime(row, colIdx);

        let hasValue = false;
        indicatorKeys.forEach(key => {
            if (colIdx[key] !== -1) {
                let val = parseFloat(row[colIdx[key]]);
                if (val === -1) val = NaN;
                rowObj[key] = isNaN(val) ? undefined : val;
                rowObj.raw[key] = row[colIdx[key]];
                if (rowObj[key] !== undefined) hasValue = true;
            }
        });

        if (hasValue) currentData.push(rowObj);
    }

    if (currentData.length === 0) return showError('未找到有效数据行');
    
    const valEl = document.getElementById('validationMsg');
    valEl.innerText = `✅ 识别表头在第 ${headerIndex+1} 行，成功提取 ${currentData.length} 条数据`;
    valEl.style.color = 'var(--success)';
    valEl.classList.remove('hidden');
    document.getElementById('btnAnalyze').disabled = false;
}

function parseRowTime(row, colIdx) {
    if (colIdx.year !== -1 && colIdx.month !== -1) {
        const y = row[colIdx.year];
        const m = row[colIdx.month];
        if (y && m) return `${y}-${m.toString().padStart(2, '0')}-01`;
    }
    if (colIdx.time !== -1 && row[colIdx.time]) {
        return formatDate(row[colIdx.time]);
    }
    return '-';
}

function findColIndex(headers, keywords, excludes = []) {
    return headers.findIndex(h => {
        const match = keywords.some(k => h.includes(k));
        const exc = excludes.some(e => h.includes(e));
        return match && !exc;
    });
}

// 4. 判定逻辑
function startAnalysis() {
    const waterType = document.querySelector('input[name="waterType"]:checked').value;
    const tbody = document.getElementById('resultTableBody');
    tbody.innerHTML = '';

    currentData.forEach((row, index) => {
        const res = analyzeRow(row, waterType);
        const tr = document.createElement('tr');
        
        let excHtml = res.isExceeded ? res.exceededFactors.map(f => {
            const isSpec = (f.key === 'ph' || f.key === 'do');
            // 如果是pH/DO，不显示倍数；其他显示(倍数)
            return `<span class="exceeded-tag ${isSpec?'priority':''}">${f.name}${isSpec?'':`(${f.multiple.toFixed(2)})`}</span>`;
        }).join('') : '<span class="text-normal">优良/达标</span>';

        tr.innerHTML = `
            <td>${index + 1}</td>
            <td><strong>${row.site}</strong></td>
            <td>${row.time}</td>
            <td><span class="grade-badge ${res.gradeClass}">${res.finalGrade}类</span></td>
            <td>${excHtml}</td>
            <td>
                <div class="tooltip-wrapper">
                    <span class="data-icon">📊</span>
                    <div class="tooltip-content">
                        <strong>原始数据明细：</strong>
                        <div class="raw-data-list">${Object.keys(STANDARDS).map(k => `
                            <div class="raw-item"><span class="raw-label">${STANDARDS[k].name}</span><span class="raw-val">${row.raw[k]||'-'}</span></div>
                        `).join('')}</div>
                    </div>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });

    document.getElementById('resultSection').classList.remove('hidden');
    document.getElementById('recordStats').innerText = `共 ${currentData.length} 条记录`;
}

/**
 * 核心修改：判定单行数据
 * 包含排序逻辑：pH/DO优先，其余按倍数降序
 */
function analyzeRow(row, waterType) {
    let maxG = 0;
    let factors = [];
    
    Object.keys(STANDARDS).forEach(key => {
        const val = row[key];
        if (val === undefined) return;
        const std = STANDARDS[key];
        const limits = (key === 'tp') ? std.limits[waterType] : std.limits;
        
        let g = 0;
        // pH 修约逻辑: 5.5-9.5 算达标（不严格为劣V）
        // 这里严格按GB标准III类判定是否列入超标因子
        if (key === 'ph') {
            g = (Math.round(val) < 6 || Math.round(val) > 9) ? 5 : 0;
        } else if (std.type === 'desc') {
            g = 5; for(let i=0; i<5; i++) if(val >= limits[i]) { g = i; break; }
        } else {
            g = 5; for(let i=0; i<5; i++) if(val <= limits[i]) { g = i; break; }
        }

        if (g > maxG) maxG = g;
        
        // 如果该项劣于 III 类，则加入超标因子列表
        if (g > 2) {
            // pH 和 溶解氧 的倍数暂设为0 (仅作为占位，不参与常规倍数排序)
            let mult = (key === 'ph' || key === 'do') ? 0 : (val - limits[2]) / limits[2];
            factors.push({ key, name: std.name, multiple: mult });
        }
    });

    // === 新增：超标因子排序逻辑 ===
    factors.sort((a, b) => {
        const priorityKeys = ['ph', 'do'];
        const aIsPriority = priorityKeys.includes(a.key);
        const bIsPriority = priorityKeys.includes(b.key);

        // 1. pH和溶解氧优先
        if (aIsPriority && !bIsPriority) return -1;
        if (!aIsPriority && bIsPriority) return 1;
        
        // 2. 如果都不是优先项，按超标倍数降序排列 (倍数大的排前面)
        return b.multiple - a.multiple;
    });

    return {
        finalGrade: GRADES[maxG], gradeClass: GRADE_CLASSES[maxG],
        isExceeded: maxG > 2, exceededFactors: factors
    };
}

// 5. 其他工具
function formatDate(val) {
    if (!val) return '-';
    if (typeof val === 'number') {
        const date = new Date((val - (25567 + 2)) * 86400 * 1000);
        return date.toISOString().split('T')[0];
    }
    return val.toString().replace(/\//g, '-');
}

function initStandardTable() {
    const tbody = document.getElementById('stdTableBody');
    const data = [
        {k:'do', l:'溶解氧'}, {k:'cod_mn', l:'高锰酸盐指数'}, {k:'cod', l:'COD'},
        {k:'bod5', l:'BOD5'}, {k:'nh3_n', l:'氨氮'}, 
        {k:'tp', l:'总磷(河)', t:'river'}, {k:'tp', l:'总磷(湖)', t:'lake'}
    ];
    let html = '';
    data.forEach(d => {
        const limits = (d.k==='tp') ? STANDARDS.tp.limits[d.t] : STANDARDS[d.k].limits;
        html += `<tr><td>${d.l}</td>${limits.map(v=>`<td>${v}</td>`).join('')}</tr>`;
    });
    tbody.innerHTML = html;
}
function showError(m) {
    const el = document.getElementById('validationMsg');
    el.innerText = `❌ ${m}`;
    el.classList.remove('hidden');
    document.getElementById('btnAnalyze').disabled = true;
}

function downloadTemplate() {
    const data = [
        ['监测数据导入表 (第一行标题可自动跳过)'],
        ['断面名称','采样时间','pH值','溶解氧','高锰酸盐指数','化学需氧量','五日生化需氧量','氨氮','总磷'],
        ['示例断面1','2024-01-01',7.5, 6.8, 3.2, 18, 2.4, 0.45, 0.12],
        ['示例断面2(严重超标)','2024-01-02',8.0, 1.5, 15, 60, 12, 2.5, 0.6]
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    XLSX.writeFile(wb, "水质数据模板.xlsx");
}

function exportResults() {
    const waterType = document.querySelector('input[name="waterType"]:checked').value;
    const header = ['序号','断面名称','时间','类别','超标因子(倍数)'];
    const data = currentData.map((r, i) => {
        const res = analyzeRow(r, waterType);
        // 导出时也按照排序后的顺序生成字符串
        const factorsStr = res.exceededFactors.map(f => {
            const isSpec = (f.key === 'ph' || f.key === 'do');
            return `${f.name}${isSpec ? '' : `(${f.multiple.toFixed(2)})`}`;
        }).join(', ');
        
        return [i+1, r.site, r.time, res.finalGrade, factorsStr];
    });
    const ws = XLSX.utils.aoa_to_sheet([header, ...data]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Result");
    XLSX.writeFile(wb, "判定结果导出.xlsx");
}
