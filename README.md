<!DOCTYPE html>
<html lang="hi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes">
    <title>गाँव-वार खाता बही (JSON फाइल शेयर)</title>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
    <style>
        * {
            box-sizing: border-box;
        }
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            background: #f0f7ea;
            margin: 0;
            padding: 20px;
        }
        .container {
            max-width: 1600px;
            margin: auto;
            background: white;
            padding: 20px;
            border-radius: 28px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.1);
        }
        h1, h2 {
            color: #1f4f1a;
            text-align: center;
            margin: 0.2em 0;
        }
        .sub {
            text-align: center;
            color: #3b6e35;
            margin-bottom: 20px;
            font-size: 14px;
        }
        form {
            background: #eaf5e5;
            padding: 20px;
            border-radius: 24px;
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
            align-items: end;
            margin-bottom: 25px;
        }
        .input-group {
            display: flex;
            flex-direction: column;
            gap: 6px;
            font-weight: bold;
            font-size: 14px;
            min-width: 100px;
        }
        input, select {
            padding: 8px 12px;
            border-radius: 14px;
            border: 1px solid #bdd6af;
            font-size: 14px;
            background: white;
        }
        input[type=number]::-webkit-inner-spin-button, 
        input[type=number]::-webkit-outer-spin-button {
            opacity: 0.5;
        }
        button {
            background: #2b6b2a;
            color: white;
            border: none;
            padding: 8px 20px;
            border-radius: 40px;
            cursor: pointer;
            font-weight: bold;
            transition: 0.2s;
        }
        button:hover {
            background: #1d471c;
        }
        .action-buttons {
            display: flex;
            gap: 12px;
            justify-content: center;
            margin: 15px 0 20px;
            flex-wrap: wrap;
        }
        .export-btn {
            background: #0f5c3e;
        }
        .pdf-btn {
            background: #aa2e1c;
        }
        .json-export {
            background: #1e6f9f;
        }
        .village-card {
            border: 1px solid #c9e0bb;
            border-radius: 24px;
            background: #fffef7;
            padding: 12px 18px;
            margin-bottom: 30px;
            box-shadow: 0 4px 10px rgba(0,0,0,0.05);
        }
        .village-title {
            font-size: 1.8rem;
            font-weight: bold;
            background: #d9e8cf;
            display: inline-block;
            padding: 4px 24px;
            border-radius: 40px;
            margin-bottom: 18px;
        }
        .table-wrapper {
            overflow-x: auto;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 18px;
            min-width: 1000px;
        }
        th, td {
            border: 1px solid #c0d8b0;
            padding: 10px 6px;
            text-align: center;
            font-size: 13px;
        }
        th {
            background: #d4e8c7;
            font-weight: bold;
        }
        .total-row {
            background: #fff1b5;
            font-weight: bold;
        }
        .edit-btn {
            background: #f5a623;
            color: white;
            padding: 4px 10px;
            font-size: 12px;
            border-radius: 30px;
            margin-right: 5px;
        }
        .delete-btn {
            background: #c7372f;
            padding: 4px 10px;
            font-size: 12px;
            border-radius: 30px;
        }
        .action-cell {
            display: flex;
            gap: 6px;
            justify-content: center;
            flex-wrap: wrap;
        }
        footer {
            text-align: center;
            margin-top: 35px;
            font-size: 13px;
            color: #4b6043;
        }
        .import-section {
            background: #fef3c7;
            padding: 12px 18px;
            border-radius: 20px;
            margin-bottom: 20px;
            display: flex;
            flex-wrap: wrap;
            gap: 15px;
            align-items: center;
            justify-content: center;
        }
        .file-input-label {
            background: #9b6b1e;
            color: white;
            padding: 8px 20px;
            border-radius: 40px;
            cursor: pointer;
            font-weight: bold;
            display: inline-block;
        }
        .file-input-label:hover {
            background: #7a5217;
        }
        #fileInput {
            display: none;
        }
        .status-msg {
            font-size: 12px;
            color: #2c5e2e;
            background: #e0f0dc;
            padding: 6px 12px;
            border-radius: 20px;
        }
        @media (max-width: 850px) {
            .input-group { width: 100%; }
            form { flex-direction: column; }
            th, td { font-size: 11px; padding: 6px 3px; }
            .action-cell { flex-direction: column; align-items: center; }
        }
    </style>
</head>
<body>
<div class="container">
    <h1>📒 गाँव-वार हिसाब बही</h1>
    <div class="sub">🌾 चावल = <strong>तमी</strong> | दाल, आलू, गोभी, सब्जी = <strong>kg</strong> | 👕 कपड़ा | + पैसा + 📅 दिनांक</div>

    <!-- JSON इम्पोर्ट सेक्शन -->
    <div class="import-section">
        <label class="file-input-label">
            📂 JSON फाइल से इम्पोर्ट करें
            <input type="file" id="fileInput" accept=".json" style="display:none">
        </label>
        <span id="importStatus" class="status-msg">कोई फाइल सेलेक्ट नहीं</span>
        <button id="clearAllBtn" style="background:#c7372f;">🗑️ सारा डाटा हटाएँ</button>
    </div>

    <form id="entryForm">
        <div class="input-group"><span>देने वाला नाम</span><input type="text" id="name" placeholder="राम सिंह" required></div>
        <div class="input-group"><span>गाँव का नाम</span><input type="text" id="village" placeholder="सोहावल" required></div>
        <div class="input-group"><span>📅 दिनांक</span><input type="date" id="date" required></div>
        
        <div class="input-group"><span>🌾 चावल (तमी)</span><input type="number" id="rice" step="any" value="0" inputmode="numeric"></div>
        <div class="input-group"><span>🫘 दाल (kg)</span><input type="number" id="dal" step="any" value="0" inputmode="numeric"></div>
        <div class="input-group"><span>🥔 आलू (kg)</span><input type="number" id="aloo" step="any" value="0" inputmode="numeric"></div>
        <div class="input-group"><span>🥬 गोभी (kg)</span><input type="number" id="gobhi" step="any" value="0" inputmode="numeric"></div>
        <div class="input-group"><span>🍆 सब्जी (kg)</span><input type="number" id="sabji" step="any" value="0" inputmode="numeric"></div>
        <div class="input-group"><span>👕 कपड़ा</span><input type="number" id="kapda" step="any" value="0" inputmode="numeric" placeholder="मात्रा"></div>
        <div class="input-group"><span>💰 दिया गया पैसा (₹)</span><input type="number" id="amount" step="any" required inputmode="numeric"></div>
        <button type="submit" id="submitBtn">✅ जोड़ें / अपडेट करें</button>
        <button type="button" id="cancelEditBtn" style="background:#888;">❌ एडिट रद्द करें</button>
    </form>

    <div class="action-buttons">
        <button id="exportExcelBtn" class="export-btn">📎 Excel डाउनलोड करें</button>
        <button id="exportPdfBtn" class="export-btn pdf-btn">🖨️ PDF सेव करें</button>
        <button id="exportJsonBtn" class="json-export">📤 JSON (शेयर करें)</button>
    </div>

    <div id="ledgerContainer"></div>
    <footer>✅ डाटा रिफ्रेश पर सुरक्षित | ✏️ एडिट बटन से सुधार करें | 📤 JSON फाइल से किसी भी फोन में शेयर करें</footer>
</div>

<script>
    let allEntries = [];
    let editModeId = null;

    function loadData() {
        const stored = localStorage.getItem("villageLedgerShareable");
        if(stored) {
            allEntries = JSON.parse(stored);
        } else {
            const today = new Date().toISOString().split('T')[0];
            allEntries = [
                { id: "1", name: "राम सिंह", village: "सोहावल", date: today, rice: 12, dal: 6, aloo: 9, gobhi: 4, sabji: 7, kapda: 25, amount: 1450 },
                { id: "2", name: "श्याम पाल", village: "सोहावल", date: today, rice: 7, dal: 3, aloo: 5, gobhi: 2, sabji: 4, kapda: 15, amount: 680 },
                { id: "3", name: "गीता देवी", village: "रामपुर", date: today, rice: 15, dal: 8, aloo: 11, gobhi: 6, sabji: 9, kapda: 30, amount: 2100 }
            ];
        }
        renderLedger();
    }

    function saveToLocal() {
        localStorage.setItem("villageLedgerShareable", JSON.stringify(allEntries));
    }

    function fillFormForEdit(entry) {
        document.getElementById('name').value = entry.name;
        document.getElementById('village').value = entry.village;
        document.getElementById('date').value = entry.date;
        document.getElementById('rice').value = entry.rice;
        document.getElementById('dal').value = entry.dal;
        document.getElementById('aloo').value = entry.aloo;
        document.getElementById('gobhi').value = entry.gobhi;
        document.getElementById('sabji').value = entry.sabji;
        document.getElementById('kapda').value = entry.kapda || 0;
        document.getElementById('amount').value = entry.amount;
        document.getElementById('submitBtn').textContent = '✏️ एडिट करें (अपडेट)';
        document.getElementById('submitBtn').style.background = '#f5a623';
        editModeId = entry.id;
        document.getElementById('entryForm').style.border = '2px solid #f5a623';
    }

    function cancelEdit() {
        document.getElementById('entryForm').reset();
        document.getElementById('submitBtn').textContent = '✅ जोड़ें / अपडेट करें';
        document.getElementById('submitBtn').style.background = '#2b6b2a';
        document.getElementById('entryForm').style.border = 'none';
        editModeId = null;
        setDefaultDate();
    }

    function addOrUpdateEntry(e) {
        e.preventDefault();
        const name = document.getElementById('name').value.trim();
        const village = document.getElementById('village').value.trim();
        const date = document.getElementById('date').value;
        const rice = parseFloat(document.getElementById('rice').value) || 0;
        const dal = parseFloat(document.getElementById('dal').value) || 0;
        const aloo = parseFloat(document.getElementById('aloo').value) || 0;
        const gobhi = parseFloat(document.getElementById('gobhi').value) || 0;
        const sabji = parseFloat(document.getElementById('sabji').value) || 0;
        const kapda = parseFloat(document.getElementById('kapda').value) || 0;
        const amount = parseFloat(document.getElementById('amount').value) || 0;

        if(!name || !village || !date) { alert("नाम, गाँव और दिनांक जरूर भरें"); return; }

        if(editModeId !== null) {
            const index = allEntries.findIndex(entry => entry.id === editModeId);
            if(index !== -1) {
                allEntries[index] = { id: editModeId, name, village, date, rice, dal, aloo, gobhi, sabji, kapda, amount };
                saveToLocal();
                renderLedger();
                cancelEdit();
            } else {
                alert("एंट्री नहीं मिली");
                cancelEdit();
            }
        } else {
            const existingIndex = allEntries.findIndex(entry => entry.name === name && entry.village === village && entry.date === date);
            if(existingIndex !== -1) {
                if(confirm("इस तारीख को इस व्यक्ति की एंट्री पहले से है। क्या आप अपडेट करना चाहते हैं?")) {
                    allEntries[existingIndex] = { ...allEntries[existingIndex], rice, dal, aloo, gobhi, sabji, kapda, amount };
                    saveToLocal();
                    renderLedger();
                }
            } else {
                const newId = Date.now().toString();
                allEntries.push({ id: newId, name, village, date, rice, dal, aloo, gobhi, sabji, kapda, amount });
                saveToLocal();
                renderLedger();
            }
            document.getElementById('entryForm').reset();
            setDefaultDate();
        }
    }

    function deleteEntry(id) {
        if(confirm("क्या आप यह एंट्री हटाना चाहते हैं?")) {
            allEntries = allEntries.filter(entry => entry.id !== id);
            saveToLocal();
            renderLedger();
        }
    }

    function clearAllData() {
        if(confirm("⚠️ क्या आप सारा डाटा हटाना चाहते हैं? यह वापस नहीं आएगा!")) {
            allEntries = [];
            saveToLocal();
            renderLedger();
            document.getElementById('importStatus').innerHTML = "✅ सारा डाटा हटा दिया गया";
            setTimeout(() => {
                document.getElementById('importStatus').innerHTML = "कोई फाइल सेलेक्ट नहीं";
            }, 2000);
        }
    }

    // JSON Export (शेयर करने के लिए)
    function exportToJSON() {
        if(allEntries.length === 0) { alert("कोई डेटा नहीं है"); return; }
        
        const exportData = {
            version: "1.0",
            exportedOn: new Date().toISOString(),
            data: allEntries
        };
        
        const jsonStr = JSON.stringify(exportData, null, 2);
        const blob = new Blob([jsonStr], {type: "application/json"});
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `khata_data_${new Date().toISOString().slice(0,19)}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        alert("📤 JSON फाइल बन गई! अब आप इसे WhatsApp, Email या किसी भी तरह से शेयर कर सकते हैं।");
    }

    // JSON Import (दूसरे फोन से लोड करने के लिए)
    function importFromJSON(file) {
        const reader = new FileReader();
        reader.onload = function(event) {
            try {
                const jsonData = JSON.parse(event.target.result);
                let importedData;
                
                // चेक करें कि फाइल हमारे फॉर्मेट में है या सिर्फ array
                if(jsonData.data && Array.isArray(jsonData.data)) {
                    importedData = jsonData.data;
                } else if(Array.isArray(jsonData)) {
                    importedData = jsonData;
                } else {
                    throw new Error("फाइल का फॉर्मेट सही नहीं है");
                }
                
                // हर एंट्री में id होनी चाहिए
                importedData = importedData.map(entry => {
                    if(!entry.id) {
                        entry.id = Date.now().toString() + Math.random();
                    }
                    // पुराने डेटा में kapda न हो तो 0 करें
                    if(entry.kapda === undefined) entry.kapda = 0;
                    return entry;
                });
                
                if(confirm(`${importedData.length} एंट्रियाँ मिलीं। क्या आप मौजूदा डाटा को इससे बदलना चाहते हैं?`)) {
                    allEntries = importedData;
                    saveToLocal();
                    renderLedger();
                    document.getElementById('importStatus').innerHTML = `✅ ${importedData.length} एंट्रियाँ लोड हो गईं`;
                    setTimeout(() => {
                        document.getElementById('importStatus').innerHTML = "कोई फाइल सेलेक्ट नहीं";
                    }, 3000);
                }
            } catch(error) {
                alert("फाइल पढ़ने में त्रुटि: " + error.message);
                document.getElementById('importStatus').innerHTML = "❌ इम्पोर्ट फेल";
            }
        };
        reader.onerror = function() {
            alert("फाइल पढ़ने में त्रुटि");
        };
        reader.readAsText(file);
    }

    function renderLedger() {
        const container = document.getElementById('ledgerContainer');
        if(!container) return;

        const villageMap = new Map();
        allEntries.forEach(entry => {
            if(!villageMap.has(entry.village)) villageMap.set(entry.village, []);
            villageMap.get(entry.village).push(entry);
        });

        if(villageMap.size === 0) {
            container.innerHTML = `<div style="text-align:center; padding: 30px;">✨ कोई डेटा नहीं, ऊपर जोड़ें ✨</div>`;
            return;
        }

        let html = '';
        for(let [village, entries] of villageMap.entries()) {
            let totalRice=0, totalDal=0, totalAloo=0, totalGobhi=0, totalSabji=0, totalKapda=0, totalAmount=0;
            entries.forEach(e => {
                totalRice += e.rice; totalDal += e.dal; totalAloo += e.aloo;
                totalGobhi += e.gobhi; totalSabji += e.sabji; totalKapda += e.kapda || 0;
                totalAmount += e.amount;
            });

            html += `<div class="village-card">
                        <div class="village-title">🏡 गाँव : ${escapeHtml(village)}</div>
                        <div class="table-wrapper">
                        <table>
                            <thead>
                                <tr>
                                    <th>देने वाला नाम</th>
                                    <th>गाँव</th>
                                    <th>📅 दिनांक</th>
                                    <th>🌾 चावल (तमी)</th>
                                    <th>🫘 दाल (kg)</th>
                                    <th>🥔 आलू (kg)</th>
                                    <th>🥬 गोभी (kg)</th>
                                    <th>🍆 सब्जी (kg)</th>
                                    <th>👕 कपड़ा</th>
                                    <th>💰 पैसा (₹)</th>
                                    <th>क्रिया</th>
                                </tr>
                            </thead>
                            <tbody>`;
            
            entries.forEach(entry => {
                const formattedDate = entry.date ? entry.date.split('-').reverse().join('-') : '';
                html += `<tr>
                            <td>${escapeHtml(entry.name)}</td>
                            <td>${escapeHtml(entry.village)}</td>
                            <td>${formattedDate}</td>
                            <td>${entry.rice}</td>
                            <td>${entry.dal}</td>
                            <td>${entry.aloo}</td>
                            <td>${entry.gobhi}</td>
                            <td>${entry.sabji}</td>
                            <td>${entry.kapda || 0}</td>
                            <td>${entry.amount}</td>
                            <td class="action-cell">
                                <button class="edit-btn" data-id="${entry.id}">✏️ एडिट</button>
                                <button class="delete-btn" data-id="${entry.id}">❌ हटाएँ</button>
                            </td>
                         <tr>`;
            });
            html += `<tr class="total-row">
                        <td colspan="3"><strong>📦 गाँव कुल योग</strong></td>
                        <td><strong>${totalRice}</strong></td>
                        <td><strong>${totalDal}</strong></td>
                        <td><strong>${totalAloo}</strong></td>
                        <td><strong>${totalGobhi}</strong></td>
                        <td><strong>${totalSabji}</strong></td>
                        <td><strong>${totalKapda}</strong></td>
                        <td><strong>₹ ${totalAmount}</strong></td>
                        <td></td>
                    </tr>
                            </tbody>
                        </table>
                        </div>
                    </div><br>`;
        }
        container.innerHTML = html;
        
        document.querySelectorAll('.edit-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                const id = btn.getAttribute('data-id');
                const entry = allEntries.find(e => e.id === id);
                if(entry) fillFormForEdit(entry);
            });
        });
        
        document.querySelectorAll('.delete-btn').forEach(btn => {
            btn.addEventListener('click', () => deleteEntry(btn.getAttribute('data-id')));
        });
    }

    function escapeHtml(str) {
        return str.replace(/[&<>]/g, function(m) {
            if(m === '&') return '&amp;';
            if(m === '<') return '&lt;';
            if(m === '>') return '&gt;';
            return m;
        });
    }

    function exportToExcel() {
        if(allEntries.length === 0) { alert("कोई डेटा नहीं"); return; }
        const excelData = allEntries.map(entry => ({
            "देने वाला नाम": entry.name,
            "गाँव का नाम": entry.village,
            "दिनांक": entry.date,
            "चावल (तमी)": entry.rice,
            "दाल (kg)": entry.dal,
            "आलू (kg)": entry.aloo,
            "गोभी (kg)": entry.gobhi,
            "सब्जी (kg)": entry.sabji,
            "कपड़ा": entry.kapda || 0,
            "दिया गया पैसा (₹)": entry.amount
        }));
        const ws = XLSX.utils.json_to_sheet(excelData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "गाँव_खाता");
        XLSX.writeFile(wb, `khata_${new Date().toISOString().slice(0,19)}.xlsx`);
    }

    async function exportToPDF() {
        const element = document.getElementById('ledgerContainer');
        if(!element || element.innerText.trim() === "") { alert("पहले कोई डेटा जोड़ें"); return; }
        const title = document.querySelector('h1')?.innerText || "गाँव खाता";
        const clone = element.cloneNode(true);
        const pdfDiv = document.createElement('div');
        pdfDiv.style.padding = "15px";
        pdfDiv.style.fontFamily = "Arial";
        pdfDiv.innerHTML = `<h2 style="text-align:center;">${title}</h2>
                            <p style="text-align:center">चावल (तमी) · दाल,आलू,गोभी,सब्जी (kg) · कपड़ा · पैसा (₹) | दिनांक: ${new Date().toLocaleDateString()}</p><hr/>`;
        pdfDiv.appendChild(clone);
        const opt = { margin: [0.5,0.5,0.5,0.5], filename: `hisab_${Date.now()}.pdf`, image: { type: 'jpeg', quality: 0.98 }, html2canvas: { scale: 2 }, jsPDF: { unit: 'in', format: 'a4', orientation: 'landscape' } };
        html2pdf().set(opt).from(pdfDiv).save();
    }

    function setDefaultDate() {
        const dateInput = document.getElementById('date');
        if(dateInput && !dateInput.value) {
            dateInput.valueAsDate = new Date();
        }
    }

    // फाइल सेलेक्ट इवेंट
    document.getElementById('fileInput').addEventListener('change', function(e) {
        if(this.files && this.files[0]) {
            document.getElementById('importStatus').innerHTML = "⏳ लोड हो रहा...";
            importFromJSON(this.files[0]);
            this.value = ''; // रीसेट करें ताकि दोबारा वही फाइल सेलेक्ट कर सकें
        }
    });

    document.getElementById('entryForm').addEventListener('submit', addOrUpdateEntry);
    document.getElementById('cancelEditBtn').addEventListener('click', cancelEdit);
    document.getElementById('exportExcelBtn').addEventListener('click', exportToExcel);
    document.getElementById('exportPdfBtn').addEventListener('click', exportToPDF);
    document.getElementById('exportJsonBtn').addEventListener('click', exportToJSON);
    document.getElementById('clearAllBtn').addEventListener('click', clearAllData);
    
    loadData();
    setDefaultDate();
</script>
</body>
</html>
