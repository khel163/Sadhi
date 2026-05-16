<!DOCTYPE html>
<html lang="hi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes">
    <title>गाँव-वार खाता बही (नंबर कीपैड + नाम पहले)</title>
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
            max-width: 1400px;
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
        }
        form {
            background: #eaf5e5;
            padding: 20px;
            border-radius: 24px;
            display: flex;
            flex-wrap: wrap;
            gap: 15px;
            align-items: end;
            margin-bottom: 35px;
        }
        .input-group {
            display: flex;
            flex-direction: column;
            gap: 6px;
            font-weight: bold;
            font-size: 14px;
            min-width: 110px;
        }
        input, select {
            padding: 8px 12px;
            border-radius: 14px;
            border: 1px solid #bdd6af;
            font-size: 14px;
            background: white;
        }
        /* नंबर इनपुट के लिए स्पिनर हटाएँ (optional) */
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
            gap: 20px;
            justify-content: center;
            margin: 15px 0 25px;
            flex-wrap: wrap;
        }
        .export-btn {
            background: #0f5c3e;
        }
        .pdf-btn {
            background: #aa2e1c;
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
<!DOCTYPE html>
<html lang="hi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes">
    <title>गाँव-वार खाता बही (नंबर कीपैड + नाम पहले)</title>
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
            max-width: 1400px;
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
        }
        form {
            background: #eaf5e5;
            padding: 20px;
            border-radius: 24px;
            display: flex;
            flex-wrap: wrap;
            gap: 15px;
            align-items: end;
            margin-bottom: 35px;
        }
        .input-group {
            display: flex;
            flex-direction: column;
            gap: 6px;
            font-weight: bold;
            font-size: 14px;
            min-width: 110px;
        }
        input, select {
            padding: 8px 12px;
            border-radius: 14px;
            border: 1px solid #bdd6af;
            font-size: 14px;
            background: white;
        }
        /* नंबर इनपुट के लिए स्पिनर हटाएँ (optional) */
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
            gap: 20px;
            justify-content: center;
            margin: 15px 0 25px;
            flex-wrap: wrap;
        }
        .export-btn {
            background: #0f5c3e;
        }
        .pdf-btn {
            background: #aa2e1c;
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
        table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 18px;
            overflow: hidden;
        }
        th, td {
            border: 1px solid #c0d8b0;
            padding: 10px 6px;
            text-align: center;
            font-size: 14px;
        }
        th {
            background: #d4e8c7;
            font-weight: bold;
        }
        .total-row {
            background: #fff1b5;
            font-weight: bold;
        }
        .delete-btn {
            background: #c7372f;
            padding: 4px 12px;
            font-size: 12px;
            border-radius: 30px;
        }
        footer {
            text-align: center;
            margin-top: 35px;
            font-size: 13px;
            color: #4b6043;
        }
        @media (max-width: 850px) {
            .input-group { width: 100%; }
            form { flex-direction: column; }
        }
    </style>
</head>
<body>
<div class="container">
    <h1>📒 गाँव-वार हिसाब बही</h1>
    <div class="sub">🌾 चावल · दाल · आलू · गोभी · सब्जी (तमी) + पैसा (₹) | <strong>पहले नाम फिर गाँव</strong></div>

    <form id="entryForm">
        <div class="input-group"><span>देने वाला नाम</span><input type="text" id="name" placeholder="राम सिंह" required></div>
        <div class="input-group"><span>गाँव का नाम</span><input type="text" id="village" placeholder="सोहावल" required></div>
        
        <!-- सभी number inputs में inputmode="numeric" और type="number" -->
        <div class="input-group"><span>चावल (तमी)</span><input type="number" id="rice" step="any" value="0" inputmode="numeric"></div>
        <div class="input-group"><span>दाल (तमी)</span><input type="number" id="dal" step="any" value="0" inputmode="numeric"></div>
        <div class="input-group"><span>आलू (तमी)</span><input type="number" id="aloo" step="any" value="0" inputmode="numeric"></div>
        <div class="input-group"><span>गोभी (तमी)</span><input type="number" id="gobhi" step="any" value="0" inputmode="numeric"></div>
        <div class="input-group"><span>सब्जी (तमी)</span><input type="number" id="sabji" step="any" value="0" inputmode="numeric"></div>

        <div class="input-group"><span>दिया गया पैसा (₹)</span><input type="number" id="amount" step="any" required inputmode="numeric"></div>
        <button type="submit">✅ जोड़ें / अपडेट करें</button>
    </form>

    <div class="action-buttons">
        <button id="exportExcelBtn" class="export-btn">📎 Excel डाउनलोड करें</button>
        <button id="exportPdfBtn" class="export-btn pdf-btn">🖨️ PDF सेव करें</button>
    </div>

    <div id="ledgerContainer"></div>
    <footer>✅ डाटा रिफ्रेश पर सुरक्षित | 📱 नंबर वाले बॉक्स पर कीपैड (मोबाइल) | नाम → गाँव क्रम</footer>
</div>

<script>
    // ---------- डाटा ----------
    let allEntries = [];

    function loadData() {
        const stored = localStorage.getItem("villageLedgerNamedFirst");
        if(stored) {
            allEntries = JSON.parse(stored);
        } else {
            // डेमो डेटा (नाम पहले दिखाने के लिए)
            allEntries = [
                { id: "1", name: "राम सिंह", village: "सोहावल", rice: 12, dal: 6, aloo: 9, gobhi: 4, sabji: 7, amount: 1450 },
                { id: "2", name: "श्याम पाल", village: "सोहावल", rice: 7, dal: 3, aloo: 5, gobhi: 2, sabji: 4, amount: 680 },
                { id: "3", name: "गीता देवी", village: "रामपुर", rice: 15, dal: 8, aloo: 11, gobhi: 6, sabji: 9, amount: 2100 }
            ];
        }
        renderLedger();
    }

    function saveToLocal() {
        localStorage.setItem("villageLedgerNamedFirst", JSON.stringify(allEntries));
    }

    function addOrUpdateEntry(e) {
        e.preventDefault();
        const name = document.getElementById('name').value.trim();
        const village = document.getElementById('village').value.trim();
        const rice = parseFloat(document.getElementById('rice').value) || 0;
        const dal = parseFloat(document.getElementById('dal').value) || 0;
        const aloo = parseFloat(document.getElementById('aloo').value) || 0;
        const gobhi = parseFloat(document.getElementById('gobhi').value) || 0;
        const sabji = parseFloat(document.getElementById('sabji').value) || 0;
        const amount = parseFloat(document.getElementById('amount').value) || 0;

        if(!name || !village) { alert("नाम और गाँव जरूर भरें"); return; }

        const existingIndex = allEntries.findIndex(entry => entry.name === name && entry.village === village);
        if(existingIndex !== -1) {
            allEntries[existingIndex] = { ...allEntries[existingIndex], rice, dal, aloo, gobhi, sabji, amount };
        } else {
            const newId = Date.now().toString();
            allEntries.push({ id: newId, name, village, rice, dal, aloo, gobhi, sabji, amount });
        }
        saveToLocal();
        renderLedger();
        document.getElementById('entryForm').reset();
    }

    function deleteEntry(id) {
        allEntries = allEntries.filter(entry => entry.id !== id);
        saveToLocal();
        renderLedger();
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
            let totalRice=0, totalDal=0, totalAloo=0, totalGobhi=0, totalSabji=0, totalAmount=0;
            entries.forEach(e => {
                totalRice += e.rice; totalDal += e.dal; totalAloo += e.aloo;
                totalGobhi += e.gobhi; totalSabji += e.sabji; totalAmount += e.amount;
            });

            html += `<div class="village-card">
                        <div class="village-title">🏡 गाँव : ${escapeHtml(village)}</div>
                        <table>
                            <thead>
                                <tr>
                                    <th>देने वाला नाम</th>
                                    <th>गाँव</th>
                                    <th>चावल (तमी)</th><th>दाल (तमी)</th>
                                    <th>आलू (तमी)</th><th>गोभी (तमी)</th><th>सब्जी (तमी)</th>
                                    <th>दिया पैसा (₹)</th><th>हटाएँ</th>
                                </tr>
                            </thead>
                            <tbody>`;
            // यहाँ हर पंक्ति में नाम के बाद गाँव दिखेगा
            entries.forEach(entry => {
                html += `<tr>
                            <td>${escapeHtml(entry.name)}</td>
                            <td>${escapeHtml(entry.village)}</td>
                            <td>${entry.rice}</td><td>${entry.dal}</td>
                            <td>${entry.aloo}</td><td>${entry.gobhi}</td>
                            <td>${entry.sabji}</td>
                            <td>${entry.amount}</td>
                            <td><button class="delete-btn" data-id="${entry.id}">❌ हटाएँ</button>
                         </tr>`;
            });
            html += `<tr class="total-row">
                        <td colspan="2"><strong>📦 गाँव कुल योग</strong></td>
                        <td><strong>${totalRice}</strong></td><td><strong>${totalDal}</strong></td>
                        <td><strong>${totalAloo}</strong></td><td><strong>${totalGobhi}</strong></td>
                        <td><strong>${totalSabji}</strong></td>
                        <td><strong>₹ ${totalAmount}</strong></td><td></td>
                    </tr>
                            </tbody>
                        </table>
                    </div><br>`;
        }
        container.innerHTML = html;
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

    // Excel Export (नाम के बाद गाँव वाला क्रम)
    function exportToExcel() {
        if(allEntries.length === 0) { alert("कोई डेटा नहीं"); return; }
        const excelData = allEntries.map(entry => ({
            "देने वाला नाम": entry.name,
            "गाँव का नाम": entry.village,
            "चावल (तमी)": entry.rice,
            "दाल (तमी)": entry.dal,
            "आलू (तमी)": entry.aloo,
            "गोभी (तमी)": entry.gobhi,
            "सब्जी (तमी)": entry.sabji,
            "दिया गया पैसा (₹)": entry.amount
        }));
        const ws = XLSX.utils.json_to_sheet(excelData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "गाँव_खाता");
        XLSX.writeFile(wb, `khata_${new Date().toISOString().slice(0,19)}.xlsx`);
    }

    // PDF Export
    async function exportToPDF() {
        const element = document.getElementById('ledgerContainer');
        if(!element || element.innerText.trim() === "") { alert("पहले कोई डेटा जोड़ें"); return; }
        const title = document.querySelector('h1')?.innerText || "गाँव खाता";
        const clone = element.cloneNode(true);
        const pdfDiv = document.createElement('div');
        pdfDiv.style.padding = "15px";
        pdfDiv.style.fontFamily = "Arial";
        pdfDiv.innerHTML = `<h2 style="text-align:center;">${title}</h2>
                            <p style="text-align:center">तमी (मात्रा) + पैसा | दिनांक: ${new Date().toLocaleDateString()}</p><hr/>`;
        pdfDiv.appendChild(clone);
        const opt = { margin: [0.5,0.5,0.5,0.5], filename: `hisab_${Date.now()}.pdf`, image: { type: 'jpeg', quality: 0.98 }, html2canvas: { scale: 2 }, jsPDF: { unit: 'in', format: 'a4', orientation: 'landscape' } };
        html2pdf().set(opt).from(pdfDiv).save();
    }

    document.getElementById('entryForm').addEventListener('submit', addOrUpdateEntry);
    document.getElementById('exportExcelBtn').addEventListener('click', exportToExcel);
    document.getElementById('exportPdfBtn').addEventListener('click', exportToPDF);
    loadData();
</script>
</body>
</html>￼Enter            background: #d9e8cf;
            display: inline-block;
            padding: 4px 24px;
            border-radius: 40px;
            margin-bottom: 18px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 18px;
            overflow: hidden;
        }
        th, td {
            border: 1px solid #c0d8b0;
            padding: 10px 6px;
            text-align: center;
            font-size: 14px;
        }
        th {
            background: #d4e8c7;
            font-weight: bold;
        }
        .total-row {
            background: #fff1b5;
            font-weight: bold;
        }
        .delete-btn {
            background: #c7372f;
            padding: 4px 12px;
            font-size: 12px;
            border-radius: 30px;
        }
        footer {
            text-align: center;
            margin-top: 35px;
            font-size: 13px;
            color: #4b6043;
        }
        @media (max-width: 850px) {
            .input-group { width: 100%; }
            form { flex-direction: column; }
        }
    </style>
</head>
<body>
<div class="container">
    <h1>📒 गाँव-वार हिसाब बही</h1>
    <div class="sub">🌾 चावल · दाल · आलू · गोभी · सब्जी (तमी) + पैसा (₹) | <strong>पहले नाम फिर गाँव</strong></div>

    <form id="entryForm">
        <div class="input-group"><span>देने वाला नाम</span><input type="text" id="name" placeholder="राम सिंह" required></div>
        <div class="input-group"><span>गाँव का नाम</span><input type="text" id="village" placeholder="सोहावल" required></div>
        
        <!-- सभी number inputs में inputmode="numeric" और type="number" -->
        <div class="input-group"><span>चावल (तमी)</span><input type="number" id="rice" step="any" value="0" inputmode="numeric"></div>
        <div class="input-group"><span>दाल (तमी)</span><input type="number" id="dal" step="any" value="0" inputmode="numeric"></div>
        <div class="input-group"><span>आलू (तमी)</span><input type="number" id="aloo" step="any" value="0" inputmode="numeric"></div>
        <div class="input-group"><span>गोभी (तमी)</span><input type="number" id="gobhi" step="any" value="0" inputmode="numeric"></div>
        <div class="input-group"><span>सब्जी (तमी)</span><input type="number" id="sabji" step="any" value="0" inputmode="numeric"></div>

        <div class="input-group"><span>दिया गया पैसा (₹)</span><input type="number" id="amount" step="any" required inputmode="numeric"></div>
        <button type="submit">✅ जोड़ें / अपडेट करें</button>
    </form>

    <div class="action-buttons">
        <button id="exportExcelBtn" class="export-btn">📎 Excel डाउनलोड करें</button>
        <button id="exportPdfBtn" class="export-btn pdf-btn">🖨️ PDF सेव करें</button>
    </div>

    <div id="ledgerContainer"></div>
    <footer>✅ डाटा रिफ्रेश पर सुरक्षित | 📱 नंबर वाले बॉक्स पर कीपैड (मोबाइल) | नाम → गाँव क्रम</footer>
</div>

<script>
    // ---------- डाटा ----------
    let allEntries = [];

nction loadData() {
        const stored = localStorage.getItem("villageLedgerNamedFirst");
        if(stored) {
            allEntries = JSON.parse(stored);
        } else {
            // डेमो डेटा (नाम पहले दिखाने के लिए)
            allEntries = [
                { id: "1", name: "राम सिंह", village: "सोहावल", rice: 12, dal: 6, aloo: 9, gobhi: 4, sabji: 7, amount: 1450 },
                { id: "2", name: "श्याम पाल", village: "सोहावल", rice: 7, dal: 3, aloo: 5, gobhi: 2, sabji: 4, amount: 680 },
                { id: "3", name: "गीता देवी", village: "रामपुर", rice: 15, dal: 8, aloo: 11, gobhi: 6, sabji: 9, amount: 2100 }
            ];
        }
        renderLedger();
    }

    function saveToLocal() {
        localStorage.setItem("villageLedgerNamedFirst", JSON.stringify(allEntries));
    }

    function addOrUpdateEntry(e) {
        e.preventDefault();
        const name = document.getElementById('name').value.trim();
        const village = document.getElementById('village').value.trim();
        const rice = parseFloat(document.getElementById('rice').value) || 0;
        const dal = parseFloat(document.getElementById('dal').value) || 0;
        const aloo = parseFloat(document.getElementById('aloo').value) || 0;
        const gobhi = parseFloat(document.getElementById('gobhi').value) || 0;
        const sabji = parseFloat(document.getElementById('sabji').value) || 0;
        const amount = parseFloat(document.getElementById('amount').value) || 0;

        if(!name || !village) { alert("नाम और गाँव जरूर भरें"); return; }

        const existingIndex = allEntries.findIndex(entry => entry.name === name && entry.village === village);
        if(existingIndex !== -1) {
            allEntries[existingIndex] = { ...allEntries[existingIndex], rice, dal, aloo, gobhi, sabji, amount };
        } else {
            const newId = Date.now().toString();
            allEntries.push({ id: newId, name, village, rice, dal, aloo, gobhi, sabji, amount });
        }
        saveToLocal();
        renderLedger();
        document.getElementById('entryForm').reset();
    }

    function deleteEntry(id) {
        allEntries = allEntries.filter(entry => entry.id !== id);
        saveToLocal();
        renderLedger();
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
            let totalRice=0, totalDal=0, totalAloo=0, totalGobhi=0, totalSabji=0, totalAmount=0;
            entries.forEach(e => {
                totalRice += e.rice; totalDal += e.dal; totalAloo += e.aloo;
                totalGobhi += e.gobhi; totalSabji += e.sabji; totalAmount += e.amount;
            });

            html += `<div class="village-card">
                        <div class="village-title">🏡 गाँव : ${escapeHtml(village)}</div>
                        <table>
                            <thead>
                                <tr>
                                    <th>देने वाला नाम</th>
                                    <th>गाँव</th>
                                    <th>चावल (तमी)</th><th>दाल (तमी)</th>
                                    <th>आलू (तमी)</th><th>गोभी (तमी)</th><th>सब्जी (तमी)</th>
