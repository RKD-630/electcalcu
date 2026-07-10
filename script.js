// ================= FULLSCREEN & GLOBAL UTILS =================
        const fsBtn = document.getElementById('fsToggle');
        if(fsBtn) {
            fsBtn.addEventListener('click', () => {
                if (!document.fullscreenElement) {
                    document.documentElement.requestFullscreen().catch(err => {
                        console.log(`Error attempting to enable full-screen mode: ${err.message}`);
                    });
                } else {
                    if (document.exitFullscreen) {
                        document.exitFullscreen();
                    }
                }
            });
            document.addEventListener('fullscreenchange', () => {
                fsBtn.querySelector('.icon').textContent = document.fullscreenElement ? '✖' : '⛶';
            });
        }

        let dbData = [], filteredData = [], debounceTimer = null;
        let html5QrcodeScanner = null;
        let currentScanType = 'ic';
        let currentScanMode = 'barcode';
        
        const showMsg = (el, text, type) => { el.textContent = text; el.className = `msg ${type}`; el.style.display = 'block'; };
        
        // Tab Switching
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
                document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
                btn.classList.add('active');
                document.getElementById(btn.dataset.tab).classList.add('active');
            });
        });

        // ================= SCANNER LOGIC =================
        function openScanner(type) {
            currentScanType = type;
            const titles = {'ic':'📷 Scan IC','resistor':'📷 Scan Resistor','smd':'📷 Scan SMD','capacitor':'📷 Scan Capacitor'};
            document.getElementById('scannerTitle').textContent = titles[type] || '📷 Scanner';
            document.getElementById('scannerModal').classList.add('active');
            setTimeout(initScanner, 300);
        }
        function closeScanner() { stopScanner(); document.getElementById('scannerModal').classList.remove('active'); }
        function switchScanMode(mode) { currentScanMode = mode; stopScanner(); setTimeout(initScanner, 300); }
        function initScanner() {
            if(currentScanMode === 'barcode') initBarcodeScanner();
            else initOCRScanner();
        }
        function initBarcodeScanner() {
            document.getElementById('reader').innerHTML = '';
            html5QrcodeScanner = new Html5Qrcode("reader");
            html5QrcodeScanner.start({ facingMode: "environment" }, { fps: 10, qrbox: { width: 250, height: 250 } }, onScanSuccess, onScanFailure).catch(() => alert("Camera access denied."));
        }
        function initOCRScanner() {
            document.getElementById('reader').innerHTML = `
                <div style="padding:40px; text-align:center;">
                    <p style="margin-bottom:20px; color:var(--text-light);">Take a photo or upload an image</p>
                    <input type="file" id="ocrInput" accept="image/*" capture="environment" style="display:none;" onchange="processOCRImage(this)">
                    <button onclick="document.getElementById('ocrInput').click()" class="scan-btn" style="font-size:1.1rem; padding:15px 30px;">📸 Capture / Upload</button>
                    <p style="margin-top:15px; font-size:0.85rem; color:var(--text-light);">Ensure good lighting & clear text</p>
                </div>`;
        }
        function onScanSuccess(decodedText) { handleScannedData(decodedText); }
        function onScanFailure() {}
        function stopScanner() { if(html5QrcodeScanner) html5QrcodeScanner.stop().catch(()=>{}); }

        async function processOCRImage(input) {
            const file = input.files[0]; if(!file) return;
            const el = document.getElementById(`${currentScanType}OcrStatus`);
            el.textContent = '⏳ Processing OCR...'; el.classList.add('active');
            try {
                const { data: { text } } = await Tesseract.recognize(file, 'eng', { logger: m => m.status==='recognizing text' && (el.textContent=`⏳ Progress: ${Math.round(m.progress*100)}%`) });
                const clean = text.trim().replace(/\s+/g, '');
                if(clean) { el.textContent = `✅ Found: ${clean}`; setTimeout(() => handleScannedData(clean), 800); }
                else el.textContent = '⚠️ No text detected. Try again.';
            } catch { el.textContent = '❌ OCR failed.'; }
        }

        function handleScannedData(data) {
            switch(currentScanType) {
                case 'ic': document.getElementById('icInput').value = data; performSearch(); break;
                case 'resistor': parseAndSetResistorCode(data); break;
                case 'smd': document.getElementById('smdInput').value = data; decodeSMD(); break;
                case 'capacitor': document.getElementById('capInput').value = data; decodeCap(); break;
            }
        }
        function parseAndSetResistorCode(code) {
            const match = code.match(/\d{3,4}/);
            if(match) { document.getElementById('smdInput').value = match[0]; decodeSMD(); document.querySelector('[data-tab="smd"]').click(); }
            else alert('Numeric SMD code not detected. Please use manual color selection.');
        }

        // ================= DATABASE LOGIC =================
        document.getElementById('fileInput').addEventListener('change', (e) => {
            const file = e.target.files[0]; if(!file) return;
            if(!/\.xls|xlsx$/i.test(file.name)) { showMsg(document.getElementById('uploadMsg'), 'Invalid file.', 'error'); return; }
            document.getElementById('fileName').textContent = file.name;
            document.getElementById('fileName').classList.add('valid');
            document.getElementById('loader').classList.add('active');
            const reader = new FileReader();
            reader.onload = (evt) => {
                try {
                    const wb = XLSX.read(evt.target.result, {type:'array'});
                    dbData = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {defval:''});
                    document.getElementById('loader').classList.remove('active');
                    showMsg(document.getElementById('uploadMsg'), `✅ Loaded ${dbData.length} records.`, 'info');
                } catch(err) { document.getElementById('loader').classList.remove('active'); showMsg(document.getElementById('uploadMsg'), `❌ ${err.message}`, 'error'); }
            };
            reader.readAsArrayBuffer(file);
        });

        function performSearch() {
            const q = document.getElementById('icInput').value.trim().toUpperCase().replace(/[^A-Z0-9]/g,'');
            if(!q) return showMsg(document.getElementById('searchMsg'), 'Enter an IC number.', 'error');
            if(!dbData.length) return showMsg(document.getElementById('searchMsg'), 'No data uploaded.', 'error');
            filteredData = dbData.filter(r => Object.values(r).some(v => String(v).toUpperCase().replace(/[^A-Z0-9]/g,'').includes(q)));
            renderTable(filteredData, q);
            document.getElementById('downloadBtn').disabled = filteredData.length === 0;
            showMsg(document.getElementById('searchMsg'), filteredData.length ? `Found ${filteredData.length}.` : 'No Record Found', filteredData.length?'info':'error');
        }
        document.getElementById('searchBtn').addEventListener('click', performSearch);
        document.getElementById('icInput').addEventListener('input', () => clearTimeout(debounceTimer) || (debounceTimer = setTimeout(() => document.getElementById('icInput').value.trim() && performSearch(), 400)));
        document.getElementById('clearBtn').addEventListener('click', () => { document.getElementById('icInput').value=''; document.getElementById('tableContainer').innerHTML='<p class="no-data">Search cleared.</p>'; document.getElementById('searchMsg').style.display='none'; document.getElementById('downloadBtn').disabled=true; });
        document.getElementById('downloadBtn').addEventListener('click', () => { const ws=XLSX.utils.json_to_sheet(filteredData); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Results"); XLSX.writeFile(wb,"ic_results.csv"); });
        function renderTable(data, q) {
            if(!data.length) return document.getElementById('tableContainer').innerHTML='<p class="no-data">No Record Found</p>';
            const h = Object.keys(data[0]); let html='<table><thead><tr>'; h.forEach(k=>html+=`<th>${k}</th>`); html+='</tr></thead><tbody>';
            data.forEach(r=>{ html+='<tr>'; h.forEach(k=>{ let c=String(r[k]); if(q && c.toUpperCase().replace(/[^A-Z0-9]/g,'').includes(q)) c=c.replace(new RegExp(`(${q})`,'gi'),'<mark class="highlight">$1</mark>'); html+=`<td>${c}</td>`; }); html+='</tr>'; });
            document.getElementById('tableContainer').innerHTML=html+'</tbody></table>';
        }

        // ================= RESISTANCE & DECODERS =================
        const cMap=[{n:'Black',v:0,m:1,h:'#000'},{n:'Brown',v:1,m:10,h:'#8B4513'},{n:'Red',v:2,m:100,h:'#F00'},{n:'Orange',v:3,m:1e3,h:'#FFA500'},{n:'Yellow',v:4,m:1e4,h:'#FF0'},{n:'Green',v:5,m:1e5,h:'#080'},{n:'Blue',v:6,m:1e6,h:'#00F'},{n:'Violet',v:7,m:1e7,h:'#8A2BE2'},{n:'Grey',v:8,m:1e8,h:'#808080'},{n:'White',v:9,m:1e9,h:'#FFF'}];
        ['band1','band2'].forEach(id=>{
            const el=document.getElementById(id);
            cMap.forEach(c=>{
                let opt = new Option(c.n,c.v);
                opt.style.backgroundColor = c.h;
                opt.style.color = ['#000','#8B4513','#F00','#080','#00F','#8A2BE2','#808080'].includes(c.h) ? '#FFF' : '#000';
                el.add(opt);
            });
        });
        document.getElementById('band3').innerHTML='<option value="1" style="background:#000;color:#FFF">Black ×1</option><option value="10" style="background:#8B4513;color:#FFF">Brown ×10</option><option value="100" style="background:#F00;color:#FFF">Red ×100</option><option value="1000" style="background:#FFA500;color:#000">Orange ×1k</option><option value="10000" style="background:#FF0;color:#000">Yellow ×10k</option><option value="100000" style="background:#080;color:#FFF">Green ×100k</option><option value="1000000" style="background:#00F;color:#FFF">Blue ×1M</option><option value="10000000" style="background:#8A2BE2;color:#FFF">Violet ×10M</option>';
        function updateRes(){
            const b1=+document.getElementById('band1').value,b2=+document.getElementById('band2').value,mult=+document.getElementById('band3').value;
            const is4=document.querySelector('input[name="bandCount"]:checked').value==='4';
            document.getElementById('tolRow').style.display=is4?'flex':'none';
            const val=((b1*10)+b2)*mult;
            document.getElementById('resDisplay').textContent=(val>=1e6?+(val/1e6).toFixed(2)+'MΩ':val>=1e3?+(val/1e3).toFixed(2)+'kΩ':+(val).toFixed(2)+'Ω');
            document.getElementById('resTol').textContent=is4?`±${document.getElementById('band4').value}%`:'';
            document.querySelector('.c1').style.backgroundColor=cMap.find(c=>c.v===b1).h;
            document.querySelector('.c2').style.backgroundColor=cMap.find(c=>c.v===b2).h;
            document.querySelector('.c3').style.backgroundColor=cMap.find(c=>c.m===mult).h;
        }
        document.getElementById('band1').addEventListener('change',updateRes);
        document.getElementById('band2').addEventListener('change',updateRes);
        document.getElementById('band3').addEventListener('change',updateRes);
        document.getElementById('band4').addEventListener('change',updateRes);
        document.querySelectorAll('input[name="bandCount"]').forEach(r=>r.addEventListener('change',updateRes));
        updateRes();

        function decodeSMD(){
            const c=document.getElementById('smdInput').value.trim().toUpperCase(), res=document.getElementById('smdResult');
            if(!c){res.innerHTML='<p style="color:var(--error);">Enter a code</p>';return;}
            let info='',val=0,sys='';
            if(c.includes('R')){val=parseFloat(c.replace('R','.'));sys='R Notation';info=`R→decimal. ${c}→${val}Ω`;}
            else if(/^\d{3}$/.test(c)){val=parseInt(c.substring(0,2))*Math.pow(10,+c[2]);sys='3-Digit';info=`${c.substring(0,2)}×10^${c[2]}=${val}Ω`;}
            else if(/^\d{4}$/.test(c)){val=parseInt(c.substring(0,3))*Math.pow(10,+c[3]);sys='4-Digit';info=`${c.substring(0,3)}×10^${c[3]}=${val}Ω`;}
            else{res.innerHTML='<p style="color:var(--error);">Invalid format. Use 3/4 digits or R notation.</p>';return;}
            const fv=val>=1e6?+(val/1e6).toFixed(2)+'MΩ':val>=1e3?+(val/1e3).toFixed(2)+'kΩ':+(val).toFixed(2)+'Ω';
            res.innerHTML=`<div class="result-grid"><div class="result-item"><h4>Value</h4><span>${fv}</span></div><div class="result-item"><h4>System</h4><span>${sys}</span></div></div><div class="info-badge">💡 ${info}</div>`;
        }
        document.getElementById('smdDecode').addEventListener('click',decodeSMD);
        document.getElementById('smdInput').addEventListener('keypress',e=>e.key==='Enter'&&decodeSMD());

        function decodeCap(){
            const c=document.getElementById('capInput').value.trim().toUpperCase(), res=document.getElementById('capResult');
            if(!c){res.innerHTML='<p style="color:var(--error);">Enter a code</p>';return;}
            const m=c.match(/^(\d{1,3})([JKMZY])?$/);
            if(!m){res.innerHTML='<p style="color:var(--error);">Invalid. Use 1-3 digits + optional letter.</p>';return;}
            let pf = 0;
            if(m[1].length <= 2) {
                pf = parseInt(m[1]);
            } else {
                pf = parseInt(m[1].substring(0,m[1].length-1)) * Math.pow(10, parseInt(m[1][m[1].length-1]));
            }
            const nf=pf/1000, uf=nf/1000;
            const tMap={J:'±5%',K:'±10%',M:'±20%',Z:'+80/-20%','':'Standard'};
            const unit=pf<1000?+(pf).toFixed(2)+' pF':nf<1?+(nf).toFixed(2)+' nF':+(nf).toFixed(2)+' nF';
            res.innerHTML=`<div class="result-grid"><div class="result-item"><h4>Capacitance</h4><span>${unit}${uf>=0.001?` (${+(uf).toFixed(2)} µF)`:''}</span></div><div class="result-item"><h4>Tolerance</h4><span>${tMap[m[2] || '']}</span></div></div><div class="info-badge">💡 ${+(pf).toFixed(2)} pF total</div>`;
        }
        document.getElementById('capDecode').addEventListener('click',decodeCap);
        document.getElementById('capInput').addEventListener('keypress',e=>e.key==='Enter'&&decodeCap());
        document.getElementById('scannerModal').addEventListener('click',e=>e.target.id==='scannerModal'&&closeScanner());

        // ================= ELECTRICITY CALCULATOR =================
        let elecApps = [];
        let elecEditIdx = -1;

        document.getElementById('elecForm').addEventListener('submit', (e) => {
            e.preventDefault();
            const app = {
                name: document.getElementById('elecName').value.trim(),
                watts: +document.getElementById('elecWatts').value,
                qty: +document.getElementById('elecQty').value,
                hours: +document.getElementById('elecHours').value,
                days: +document.getElementById('elecDays').value
            };
            if(elecEditIdx > -1) {
                elecApps[elecEditIdx] = app;
                elecEditIdx = -1;
                document.getElementById('elecAddBtn').textContent = '➕ Add Appliance';
                document.getElementById('elecCancelEdit').style.display = 'none';
            } else {
                elecApps.push(app);
            }
            document.getElementById('elecForm').reset();
            document.getElementById('elecQty').value = 1;
            document.getElementById('elecDays').value = 30;
            renderElecTable();
            calcElecTotals();
        });

        document.getElementById('elecCostPerUnit').addEventListener('input', calcElecTotals);

        function renderElecTable() {
            const tb = document.getElementById('elecTableBody');
            if(!elecApps.length) return tb.innerHTML='<tr><td colspan="8" class="no-data">No appliances added yet.</td></tr>';
            let h = '';
            elecApps.forEach((a,i) => {
                const dw = a.watts * a.qty * a.hours;
                const mk = (dw * a.days) / 1000;
                h += `<tr>
                    <td>${a.name}</td><td>${a.watts}</td><td>${a.qty}</td><td>${a.hours}</td><td>${a.days}</td>
                    <td>${+(dw).toFixed(2)} W·h</td><td>${+(mk).toFixed(2)} kWh</td>
                    <td><button class="secondary" style="padding:6px 10px;font-size:0.8rem;" onclick="elecEdit(${i})">✏️</button> 
                        <button class="delete-btn" onclick="elecRemove(${i})">🗑️</button></td>
                </tr>`;
            });
            tb.innerHTML = h;
        }

        function calcElecTotals() {
            let dwTotal=0, mkTotal=0;
            elecApps.forEach(a => {
                const d = a.watts*a.qty*a.hours; dwTotal+=d; mkTotal+=(d*a.days)/1000;
            });
            const yk = mkTotal*12;
            const cost = +document.getElementById('elecCostPerUnit').value || 0;
            document.getElementById('elecSumDaily').textContent = `${+(dwTotal).toFixed(2)} W·h`;
            document.getElementById('elecSumMonth').textContent = `${+(mkTotal).toFixed(2)} kWh`;
            document.getElementById('elecSumYear').textContent = `${+(yk).toFixed(2)} kWh`;
            document.getElementById('elecMonthCost').textContent = `₹${+(mkTotal*cost).toFixed(2)}`;
            document.getElementById('elecYearCost').textContent = `₹${+(yk*cost).toFixed(2)}`;
        }

        function elecRemove(i) {
            if(!confirm('Remove this appliance?')) return;
            elecApps.splice(i,1);
            if(elecEditIdx===i) { elecEditIdx=-1; document.getElementById('elecForm').reset(); document.getElementById('elecQty').value=1; document.getElementById('elecDays').value=30; document.getElementById('elecAddBtn').textContent='➕ Add Appliance'; document.getElementById('elecCancelEdit').style.display='none'; }
            else if(elecEditIdx>i) elecEditIdx--;
            renderElecTable(); calcElecTotals();
        }

        function elecEdit(i) {
            const a = elecApps[i];
            document.getElementById('elecName').value = a.name;
            document.getElementById('elecWatts').value = a.watts;
            document.getElementById('elecQty').value = a.qty;
            document.getElementById('elecHours').value = a.hours;
            document.getElementById('elecDays').value = a.days;
            elecEditIdx = i;
            document.getElementById('elecAddBtn').textContent = '💾 Update';
            document.getElementById('elecCancelEdit').style.display = 'inline-block';
            document.getElementById('elecName').focus();
        }
        document.getElementById('elecCancelEdit').addEventListener('click', () => {
            elecEditIdx = -1; document.getElementById('elecForm').reset(); document.getElementById('elecQty').value=1; document.getElementById('elecDays').value=30;
            document.getElementById('elecAddBtn').textContent='➕ Add Appliance'; document.getElementById('elecCancelEdit').style.display='none';
        });

        function elecResetAll() {
            if(!confirm('Reset all electricity data?')) return;
            elecApps = []; elecEditIdx = -1;
            document.getElementById('elecForm').reset(); document.getElementById('elecQty').value=1; document.getElementById('elecDays').value=30;
            document.getElementById('elecCostPerUnit').value=8; document.getElementById('elecAddBtn').textContent='➕ Add Appliance';
            document.getElementById('elecCancelEdit').style.display='none';
            renderElecTable(); calcElecTotals();
        }

        function elecExportCSV() {
            if(!elecApps.length) return alert('No data to export.');
            let csv='Appliance,Watts,Qty,Hours/Day,Days/Month,Total Watts/Day,kWh/Month\n';
            let tDW=0, tMK=0;
            elecApps.forEach(a => {
                const d=a.watts*a.qty*a.hours, k=(d*a.days)/1000; tDW+=d; tMK+=k;
                csv+=`${a.name},${a.watts},${a.qty},${a.hours},${a.days},${+(d).toFixed(2)},${+(k).toFixed(2)}\n`;
            });
            const c=+document.getElementById('elecCostPerUnit').value||0;
            csv+=`\nTOTALS,,,,,${+(tDW).toFixed(2)},${+(tMK).toFixed(2)}\n`;
            csv+=`Cost per Unit (₹/kWh),${+(c).toFixed(2)}\nMonthly Bill,₹${+(tMK*c).toFixed(2)}\nYearly Bill,₹${+(tMK*12*c).toFixed(2)}\n`;
            const blob=new Blob([csv],{type:'text/csv'}); const u=URL.createObjectURL(blob); const a=document.createElement('a');
            a.href=u; a.download='electricity_usage.csv'; a.click(); URL.revokeObjectURL(u);
        }

        // ================= W/Ve+/A CALCULATOR =================
        function calculateWVA() {
            const wInput = document.getElementById('calcW');
            const vInput = document.getElementById('calcV');
            const aInput = document.getElementById('calcA');
            const res = document.getElementById('wvaResult');
            
            let w = parseFloat(wInput.value);
            let v = parseFloat(vInput.value);
            let a = parseFloat(aInput.value);
            
            let filled = 0;
            if(!isNaN(w)) filled++;
            if(!isNaN(v)) filled++;
            if(!isNaN(a)) filled++;
            
            if(filled < 2) {
                res.innerHTML = '<span style="color:var(--error);">Please enter exactly two values.</span>';
                res.style.display = 'flex';
                return;
            }
            if(filled === 3) {
                res.innerHTML = '<span style="color:var(--error);">Please leave one field empty to calculate.</span>';
                res.style.display = 'flex';
                return;
            }
            
            let resultText = '';
            if(isNaN(w)) {
                w = v * a;
                wInput.value = +(w).toFixed(2);
                resultText = `Calculated Power: ${+(w).toFixed(2)} Watts`;
            } else if(isNaN(v)) {
                if(a === 0) { res.innerHTML = '<span style="color:var(--error);">Amperes cannot be zero.</span>'; res.style.display = 'flex'; return; }
                v = w / a;
                vInput.value = +(v).toFixed(2);
                resultText = `Calculated Voltage: ${+(v).toFixed(2)} Volts`;
            } else if(isNaN(a)) {
                if(v === 0) { res.innerHTML = '<span style="color:var(--error);">Volts cannot be zero.</span>'; res.style.display = 'flex'; return; }
                a = w / v;
                aInput.value = +(a).toFixed(2);
                resultText = `Calculated Current: ${+(a).toFixed(2)} Amperes`;
            }
            
            res.innerHTML = `<span>${resultText}</span>`;
            res.style.display = 'flex';
        }
        
        function clearWVA() {
            document.getElementById('calcW').value = '';
            document.getElementById('calcV').value = '';
            document.getElementById('calcA').value = '';
            document.getElementById('wvaResult').style.display = 'none';
        }