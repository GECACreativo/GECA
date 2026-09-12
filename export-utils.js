// ── GECA Creativo · Utilidades compartidas de exportación (Excel, CSV, PDF) ──
// Usado por admin.html y dashboard.html. Requiere que la página haya cargado
// xlsx.full.min.js, html2canvas.min.js y jspdf.umd.min.js antes de este script.

var GECA_ACENTOS = {'á':'a','é':'e','í':'i','ó':'o','ú':'u','ü':'u','ñ':'n'};

function gecaSlug(s){
  var texto = (s || 'archivo').toString().toLowerCase();
  texto = texto.replace(/[áéíóúüñ]/g, function(ch){ return GECA_ACENTOS[ch] || ch; });
  return texto.replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '') || 'archivo';
}

function gecaDescargarBlob(blob, filename){
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(function(){ URL.revokeObjectURL(url); }, 1000);
}

function gecaExportCSV(filename, headers, rows){
  function esc(val){
    val = (val === null || val === undefined) ? '' : String(val);
    if(/[",\n]/.test(val)) val = '"' + val.replace(/"/g, '""') + '"';
    return val;
  }
  var csv = headers.map(esc).join(',') + '\n';
  rows.forEach(function(row){ csv += row.map(esc).join(',') + '\n'; });
  var blob = new Blob(['﻿' + csv], {type:'text/csv;charset=utf-8;'});
  gecaDescargarBlob(blob, filename);
}

function gecaSheetName(name){
  // Excel prohíbe : \ / ? * [ ] en el nombre de una hoja, y máximo 31 caracteres.
  var safe = (name || 'Datos').toString().replace(/[:\\\/\?\*\[\]]/g, ' ').replace(/\s+/g, ' ').trim();
  return safe.substring(0, 31) || 'Datos';
}

function gecaExportExcel(filename, sheetName, headers, rows){
  try{
    var data = [headers].concat(rows);
    var ws = XLSX.utils.aoa_to_sheet(data);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, gecaSheetName(sheetName));
    XLSX.writeFile(wb, filename);
  } catch(err){
    console.error('Error al exportar a Excel:', err);
    alert('No se pudo generar el archivo de Excel: ' + err.message);
  }
}

var GECA_LOGO_URL = 'Logo%20GECA.png';
var GECA_LOGO_CACHE = null;

function gecaCargarLogo(){
  if(GECA_LOGO_CACHE) return Promise.resolve(GECA_LOGO_CACHE);
  return new Promise(function(resolve){
    var img = new Image();
    img.onload = function(){
      try{
        var canvas = document.createElement('canvas');
        canvas.width = img.naturalWidth; canvas.height = img.naturalHeight;
        canvas.getContext('2d').drawImage(img, 0, 0);
        GECA_LOGO_CACHE = {dataUrl: canvas.toDataURL('image/png'), ratio: img.naturalWidth / img.naturalHeight};
      } catch(e){ GECA_LOGO_CACHE = null; }
      resolve(GECA_LOGO_CACHE);
    };
    img.onerror = function(){ resolve(null); };
    img.src = GECA_LOGO_URL;
  });
}

function gecaDibujarMembrete(pdf, pageWidth, pageHeight, logo, margenSup, margenInf, numeroPagina){
  // Fondo del encabezado y pie (cubre cualquier desborde del contenido)
  pdf.setFillColor(250, 248, 244);
  pdf.rect(0, 0, pageWidth, margenSup, 'F');
  pdf.rect(0, pageHeight - margenInf, pageWidth, margenInf, 'F');

  // Encabezado: logo + título + fecha
  if(logo && logo.dataUrl){
    var logoH = 12, logoW = logoH * logo.ratio;
    try{ pdf.addImage(logo.dataUrl, 'PNG', 10, 7, logoW, logoH); } catch(e){}
  }
  pdf.setFont(undefined, 'bold');
  pdf.setFontSize(13);
  pdf.setTextColor(15, 14, 13);
  pdf.text('GECA Creativo', pageWidth - 10, 12, {align:'right'});
  pdf.setFont(undefined, 'normal');
  pdf.setFontSize(8.5);
  pdf.setTextColor(90, 85, 80);
  pdf.text('Reporte de evaluación · generado el ' + new Date().toLocaleDateString('es-CO'), pageWidth - 10, 17, {align:'right'});
  pdf.setDrawColor(255, 136, 143);
  pdf.setLineWidth(0.7);
  pdf.line(10, margenSup - 3, pageWidth - 10, margenSup - 3);

  // Pie de página: datos de la empresa
  pdf.setDrawColor(230, 225, 220);
  pdf.setLineWidth(0.3);
  pdf.line(10, pageHeight - margenInf + 3, pageWidth - 10, pageHeight - margenInf + 3);
  pdf.setFontSize(7.5);
  pdf.setTextColor(90, 85, 80);
  pdf.text('GECA Creativo  ·  Medellín, Colombia  ·  +57 300 183 4330  ·  mentescreativas@gecacreativo.com  ·  www.gecacreativo.com', 10, pageHeight - margenInf + 8);
  pdf.text('Página ' + numeroPagina, pageWidth - 10, pageHeight - margenInf + 8, {align:'right'});
}

async function gecaExportPDF(filename, elementId, btn){
  var el = document.getElementById(elementId);
  if(!el){ return; }
  var textoOriginal = btn ? btn.textContent : null;
  if(btn){ btn.disabled = true; btn.textContent = 'Generando PDF...'; }
  try{
    var logo = await gecaCargarLogo();
    var canvas = await html2canvas(el, {scale:2, backgroundColor:'#faf8f4', useCORS:true});
    var imgData = canvas.toDataURL('image/png');
    var jsPDFLib = window.jspdf.jsPDF;
    var pdf = new jsPDFLib('p', 'mm', 'a4');
    var pageWidth = pdf.internal.pageSize.getWidth();
    var pageHeight = pdf.internal.pageSize.getHeight();
    var margenLat = 10, margenSup = 24, margenInf = 20;
    var altoUtil = pageHeight - margenSup - margenInf;
    var imgWidth = pageWidth - margenLat * 2;
    var imgHeight = canvas.height * imgWidth / canvas.width;

    var mostrado = 0;
    var pagina = 1;
    pdf.addImage(imgData, 'PNG', margenLat, margenSup - mostrado, imgWidth, imgHeight);
    gecaDibujarMembrete(pdf, pageWidth, pageHeight, logo, margenSup, margenInf, pagina);
    mostrado += altoUtil;

    while(mostrado < imgHeight){
      pdf.addPage();
      pagina++;
      pdf.addImage(imgData, 'PNG', margenLat, margenSup - mostrado, imgWidth, imgHeight);
      gecaDibujarMembrete(pdf, pageWidth, pageHeight, logo, margenSup, margenInf, pagina);
      mostrado += altoUtil;
    }
    pdf.save(filename);
  } finally {
    if(btn){ btn.disabled = false; btn.textContent = textoOriginal; }
  }
}
