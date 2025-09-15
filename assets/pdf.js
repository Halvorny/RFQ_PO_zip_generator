'use strict';
// Build RFQ/PO Document List PDFs using jsPDF (no autotable; simple manual table).

(function(){
  const { jsPDF } = window.jspdf || {};

  function extractIds(title){
    const rfq = /RFQ\s*([0-9]+)/i.exec(title||''); 
    const po  = /PO\s*([0-9]+)/i.exec(title||''); 
    return { rfq: rfq ? rfq[1] : '', po: po ? po[1] : '' };
  }

  function header(doc, mode, titleText, refText){
    doc.setFont('helvetica','bold'); doc.setFontSize(14);
    doc.text(titleText, 20, 20);
    doc.setFont('helvetica','normal'); doc.setFontSize(10);
    if(refText) doc.text(refText, 20, 28);
    doc.setDrawColor(200); doc.line(20, 32, 190, 32);
  }

  function footer(doc, mode, page, pages){
    doc.setFont('helvetica','normal'); doc.setFontSize(9);
    const label = mode==='po' ? 'PO Document list' : 'RFQ Document list';
    doc.text(`${label} · Page ${page} of ${pages}`, 20, 286);
  }

  function tableHeader(doc){
    doc.setFont('helvetica','bold'); doc.setFontSize(10);
    doc.text('Part number', 20, 40);
    doc.text('Drawing folder', 70, 40);
    doc.text('Drawing ref.', 140, 40);
    doc.setDrawColor(220); doc.line(20, 43, 190, 43);
  }

  function emitRow(doc, y, pn, drw, ref){
    doc.setFont('helvetica','normal'); doc.setFontSize(10);
    doc.text(pn || '', 20, y);
    doc.text(drw || '', 70, y);
    doc.text(ref || '', 140, y);
  }

  function buildRFQ(manifest, { titleFrom }={}){
    if(!jsPDF) return alert('jsPDF not loaded.');
    const doc = new jsPDF({unit:'mm', format:'a4'});
    const ids = extractIds(titleFrom || manifest.root);
    const title = ids.rfq ? `Request document: RFQ${ids.rfq}` : 'Request document';
    header(doc, 'rfq', title, manifest.root);

    tableHeader(doc);
    let y = 50, page = 1, rows = [];
    // Only PDFs, one row per PN/DRW (first PDF found)
    for(const grp of manifest.grouped){
      const pdf = grp.files.find(f => f.ext === 'pdf');
      if(!pdf) continue;
      rows.push([grp.pn, grp.drw, pdf.name]);
    }
    // paginate
    const maxY = 275, step = 6;
    rows.forEach((r,i)=>{
      if(y > maxY){
        footer(doc,'rfq',page, ''); 
        doc.addPage(); page++; header(doc,'rfq', title, manifest.root); tableHeader(doc); y = 50;
      }
      emitRow(doc, y, r[0], r[1], r[2]); y += step;
    });
    const pages = doc.getNumberOfPages();
    for(let p=1;p<=pages;p++){ doc.setPage(p); footer(doc,'rfq',p,pages); }
    const fname = ids.rfq ? `RFQ${ids.rfq} Document list.pdf` : `RFQ Document list.pdf`;
    doc.save(fname);
  }

  function buildPO(manifest, { titleFrom }={}){
    if(!jsPDF) return alert('jsPDF not loaded.');
    const doc = new jsPDF({unit:'mm', format:'a4'});
    const ids = extractIds(titleFrom || manifest.root);
    const title = ids.po ? `Purchase order: PO${ids.po}` : 'Purchase order';
    const ref = ids.rfq ? `Reference document: RFQ${ids.rfq}` : '';
    header(doc, 'po', title, ref || manifest.root);

    tableHeader(doc);
    let y = 50, page = 1;
    const maxY = 275, step = 6;

    // Rows: every included file for each PN/DRW
    for(const grp of manifest.grouped){
      // Sort by ext then name for readability
      const sorted = grp.files.slice().sort((a,b)=> (a.ext+a.name).localeCompare(b.ext+b.name));
      for(const f of sorted){
        if(y > maxY){
          footer(doc,'po',page,''); doc.addPage(); page++;
          header(doc,'po', title, ref || manifest.root); tableHeader(doc); y = 50;
        }
        emitRow(doc, y, grp.pn, grp.drw, f.name);
        y += step;
      }
    }
    const pages = doc.getNumberOfPages();
    for(let p=1;p<=pages;p++){ doc.setPage(p); footer(doc,'po',p,pages); }
    const fname = ids.po ? `PO${ids.po} Document list.pdf` : `PO Document list.pdf`;
    doc.save(fname);
  }

  window.pdfBuilder = { buildRFQ, buildPO };
})();
