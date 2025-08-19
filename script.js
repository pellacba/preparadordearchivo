(function(){
  let workbook, sheetName, originalAOA, resultAOA, filenameBase = "salida";
  const $ = sel => document.querySelector(sel);
  const $status = $('#status');
  const $process = $('#process');
  const $reset = $('#reset');
  const $previews = $('#previews');
  const $fname = $('#filename');

  function setStatus(msg, cls){
    $status.className = cls || ""; $status.textContent = msg || "";
  }

  function parseFile(file){
    filenameBase = file.name.replace(/\.[^.]+$/, '') || 'salida';
    $fname.value = filenameBase;
    const reader = new FileReader();
    reader.onload = e => {
      try{
        const data = new Uint8Array(e.target.result);
        workbook = XLSX.read(data, {type:'array', cellDates:true});
        sheetName = workbook.SheetNames[0];
        const ws = workbook.Sheets[sheetName];
        originalAOA = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:"", blankrows:false});
        if(!originalAOA || !originalAOA.length){
          setStatus('El archivo está vacío.', 'err');
          return;
        }
        $process.disabled = false; $reset.disabled = false;
        setStatus(`Archivo cargado: ${file.name} ✓`, 'ok');
        renderPreview('#preview-before', originalAOA);
        $previews.classList.remove('hidden');
        resultAOA = null; // limpio resultado previo
      }catch(err){
        console.error(err);
        setStatus('No pude leer el archivo. ¿Es un Excel válido?', 'err');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function detectVendedorIndex(aoa){
    const candRows = aoa.slice(0,2);
    for(const row of candRows){
      for(let i=0;i<row.length;i++){
        const v = String(row[i]).trim().toLowerCase();
        if(v === 'vendedor'){ return i; }
      }
    }
    let maxIdx = -1, maxCount=-1;
    const cols = Math.max(...aoa.map(r=>r.length));
    for(let c=0;c<cols;c++){
      let count=0; for(let r=0;r<aoa.length;r++){ if(aoa[r][c]!==undefined && aoa[r][c]!=="" ) count++; }
      if(count>maxCount){ maxCount=count; maxIdx=c; }
    }
    return maxIdx < 0 ? 0 : maxIdx;
  }

  function processAOA(){
    if(!originalAOA) return;
    let aoa = originalAOA.map(row => row.slice());

    const vIdx = detectVendedorIndex(aoa);
    for(let r=1;r<aoa.length;r++){
      if(!aoa[r]) aoa[r]=[];
      aoa[r][vIdx] = 221;
    }

    aoa = aoa.map(row => row.slice(1));

    if(aoa.length>0) aoa.shift();

    resultAOA = aoa;
    renderPreview('#preview-after', resultAOA);
  }

  function renderPreview(sel, aoa){
    const cont = document.querySelector(sel);
    if(!aoa || !aoa.length){ cont.innerHTML = '<small class="warn">(sin datos)</small>'; return; }
    const head = aoa[0] || [];
    const sample = aoa.slice(0, Math.min(aoa.length, 12));
    let html = '<div style="overflow:auto; max-height:360px">\n<table>\n<thead><tr>' + head.map(h=>`<th>${escapeHtml(String(h))}</th>`).join('') + '</tr></thead>\n<tbody>';
    for(let r=1; r<sample.length; r++){
      html += '<tr>' + (sample[r]||[]).map(c=>`<td>${escapeHtml(c==null?"":String(c))}</td>`).join('') + '</tr>';
    }
    html += '</tbody></table></div>';
    cont.innerHTML = html;
  }

  function escapeHtml(s){
    return s.replace(/[&<>"]/g, c=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c]));
  }

  function downloadCSV(){
    if(!resultAOA){ processAOA(); }
    if(!resultAOA || !resultAOA.length){ setStatus('No hay datos para exportar.', 'err'); return; }
    const ws = XLSX.utils.aoa_to_sheet(resultAOA);
    const csv = XLSX.utils.sheet_to_csv(ws, { FS: ',', RS: '\n' });
    const blob = new Blob([csv], {type:'text/csv;charset=utf-8;'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const fname = ($fname.value && $fname.value.trim()) || filenameBase || 'salida';
    a.href = url; a.download = `${fname}.csv`;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
    setStatus('CSV descargado (delimitado por comas).', 'ok');
  }

  const input = document.getElementById('file');
  input.addEventListener('change', e => { const f=e.target.files[0]; if(f) parseFile(f); });

  const dz = document.getElementById('drop');
  ;['dragenter','dragover'].forEach(ev => dz.addEventListener(ev, e=>{e.preventDefault(); dz.style.borderColor = 'var(--accent)';}));
  ;['dragleave','drop'].forEach(ev => dz.addEventListener(ev, e=>{e.preventDefault(); dz.style.borderColor = '#374151';}));
  dz.addEventListener('drop', e=>{ const f = e.dataTransfer.files && e.dataTransfer.files[0]; if(f) parseFile(f); });

  $process.addEventListener('click', ()=>{ try{ processAOA(); downloadCSV(); } catch(err){ console.error(err); setStatus('Error al procesar / exportar.', 'err'); } });
  $reset.addEventListener('click', ()=>{ workbook=null; sheetName=null; originalAOA=null; resultAOA=null; input.value=''; setStatus('Listo para un nuevo archivo.'); $('#preview-before').innerHTML=''; $('#preview-after').innerHTML=''; $previews.classList.add('hidden'); $process.disabled=true; $reset.disabled=true; $fname.value=""; });
})();

