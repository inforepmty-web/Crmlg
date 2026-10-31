const PDFJS='https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.mjs';
const PDFLIB='https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/+esm';
const XLSXURL='https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm';
const clean=v=>String(v??'').replace(/\s+/g,' ').trim();
const html=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const blank=()=>({suscriptor:'',rfc:'',email:'',accountNo:'',tipoContratacion:'',authorizedPerson:'',mesesPlazo:'',telefono:'',representante:'',domicilioFiscal:'',domicilioEntrega:''});
const match=(t,re)=>{const m=t.match(re);return clean(m?.[1]??m?.[0]??'')};

export function renderContractWorkspace({templateName='',sourceName='',fields={},seller={}}={}){
 const f={...blank(),...fields},ok=clean(seller.name)&&clean(seller.rfc);
 const input=(label,name,value='',type='text')=>`<label>${html(label)}<input name="${name}" type="${type}" value="${html(value)}"/></label>`;
 return `<section class="panel"><div class="toolbar"><div class="grow"><h2>Generar contrato de servicios</h2><p class="muted">Carga la propuesta del cliente y revisa los datos antes de generar el contrato.</p></div></div>
 <div class="formgrid"><label>Plantilla PDF del contrato<input id="contractTemplateFile" type="file" accept="application/pdf,.pdf"/></label><label>Propuesta o archivo fuente (PDF o Excel)<input id="contractSourceFile" type="file" accept="application/pdf,.pdf,.xlsx,.xls,.csv"/></label></div>
 <p class="hint">Plantilla: <strong>${html(templateName||'No seleccionada')}</strong> · Fuente: <strong>${html(sourceName||'No seleccionada')}</strong></p>
 ${ok?`<div class="success">Vendedor: <strong>${html(seller.name)}</strong> · RFC ${html(seller.rfc)}</div>`:`<div class="warning">Configura el nombre y RFC del vendedor en Vendedores para poder generar el contrato.</div>`}
 <form id="contractReviewForm"><h3>Datos reconocidos</h3><p class="hint">Corrige lo necesario. Los campos sin información en la fuente quedan vacíos.</p><div class="formgrid">
 ${input('Nombre o razón social','suscriptor',f.suscriptor)}${input('RFC del cliente','rfc',f.rfc)}${input('Correo electrónico','email',f.email,'email')}${input('No. de cuenta','accountNo',f.accountNo)}
 <label>Tipo de contratación<select name="tipoContratacion"><option value="">Seleccionar</option>${['Renovación','Suscripción Nueva','Adición','Creación Subcuenta','Responsabilidad de pago','Eq. Prop. Cliente'].map(x=>`<option ${f.tipoContratacion===x?'selected':''}>${html(x)}</option>`).join('')}</select></label>
 ${input('Persona autorizada para recibir','authorizedPerson',f.authorizedPerson)}${input('Plazo mínimo (meses)','mesesPlazo',f.mesesPlazo,'number')}${input('Teléfono de contacto','telefono',f.telefono)}${input('Representante legal','representante',f.representante)}${input('Domicilio fiscal','domicilioFiscal',f.domicilioFiscal)}${input('Domicilio de entrega','domicilioEntrega',f.domicilioEntrega)}</div>
 <p class="hint">Equipos, Servicios Contratados y Servicios Adicionales conservarán “VER ANEXO”. La propuesta PDF se agrega al final; si subes Excel, se crea una hoja de anexo.</p>
 <button type="button" class="primary" id="generateContract" ${!templateName||!sourceName||!ok?'disabled':''}>Generar y descargar contrato</button></form></section>`;
}

async function pdfText(file){
 const pdfjs=await import(PDFJS);pdfjs.GlobalWorkerOptions.workerSrc='https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.mjs';
 const doc=await pdfjs.getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise,all=[];
 for(let n=1;n<=doc.numPages;n++){const p=await doc.getPage(n),c=await p.getTextContent(),items=c.items.filter(x=>x.str?.trim()).map(x=>({x:x.transform[4],y:x.transform[5],s:x.str.trim()})).sort((a,b)=>b.y-a.y||a.x-b.x),lines=[];for(const it of items){let l=lines.find(z=>Math.abs(z.y-it.y)<2.5);if(!l){l={y:it.y,a:[]};lines.push(l)}l.a.push(it)}all.push(lines.sort((a,b)=>b.y-a.y).map(l=>l.a.sort((a,b)=>a.x-b.x).map(x=>x.s).join(' ')).join('\n'))}
 return all.join('\n\n');
}
async function spreadsheetText(file){if(/\.csv$/i.test(file.name))return file.text();const X=await import(XLSXURL),w=X.read(await file.arrayBuffer(),{type:'array'});return w.SheetNames.map(n=>'HOJA: '+n+'\n'+X.utils.sheet_to_json(w.Sheets[n],{header:1,defval:''}).map(r=>r.map(x=>String(x??'').trim()).filter(Boolean).join(' | ')).filter(Boolean).join('\n')).join('\n\n')}
export async function parseContractSource(file){const isPdf=/\.pdf$/i.test(file.name)||file.type==='application/pdf',text=isPdf?await pdfText(file):await spreadsheetText(file),f=blank();
 f.suscriptor=match(text,/DATOS DEL CLIENTE\s+([^\n]+?)(?=\s+INFORMACIÓN DE LA OFERTA)/i)||match(text,/(?:Nombre o razón social|Razón social|Nombre del cliente|Cliente|Empresa)\s*[:|]\s*([^\n|]+)/i);
 f.rfc=match(text,/\bRFC\s*[:|]\s*([A-Z&Ñ]{3,4}\s*\d{6}\s*[A-Z0-9]{3})/i).replace(/\s/g,'').toUpperCase();
 f.email=match(text,/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
 f.accountNo=match(text,/(?:FOLIO DE COTIZACIÓN|No\.?\s*Cuenta|Cuenta AT&T|Cuenta)\s*[:|]?\s*(\d{6,12})/i);
 f.tipoContratacion=match(text,/(?:Trámite|Tipo de contratación)\s*[:|]\s*(Renovación|Suscripción Nueva|Adición|Creación Subcuenta|Responsabilidad de pago|Eq\.?\s*Prop\.?\s*Cliente)/i);
 f.authorizedPerson=match(text,/Persona autorizada para recibir\s*:\s*([^\n|]+)/i);
 f.mesesPlazo=match(text,/(?:Plazo de contratación|Plazo)\s*:\s*(\d{1,2})\s*(?:Meses|meses|m)/i);
 return {fields:f,text,isPdf};}

function cover(page,x,top,w,h,color,rgb,H){page.drawRectangle({x,y:H-top-h,width:w,height:h,color:rgb(...color),borderWidth:0})}
function text(page,font,value,x,top,w,size,H){const v=clean(value);if(!v)return;let s=size;while(font.widthOfTextAtSize(v,s)>w&&s>5.5)s-=.25;page.drawText(v,{x,y:H-top-s,size:s,font,color:rgb(0.03,0.03,0.03),maxWidth:w})}
function mark(page,font,x,top,H){page.drawText('X',{x,y:H-top-8,size:8,font,color:rgb(0.03,0.03,0.03)})}
function clearSample(page,rgb,H){const white=[1,1,1],blue=[232/255,243/255,250/255],grey=[220/255,220/255,220/255];[
 [80,50,30,14,white],[203,50,31,14,white],[484,49,58,15,white],[415,82,10,14,blue],[27,117,98,16,blue],[134,117,112,16,blue],[283,103,137,16,blue],[369,119,77,16,blue],[527,105,64,17,blue],[487,134,10,16,blue],
 [22,199,188,16,white],[444,198,112,16,white],[481,217,70,17,white],[22,240,430,16,white],[462,239,86,17,white],[22,260,336,18,white],[361,260,185,18,white],[22,282,327,15,white],[44,303,26,16,white],[176,304,70,16,white],[379,303,90,17,white],[512,297,55,17,white],[115,337,10,16,grey],[197,337,19,16,white],[193,231,10,15,grey],[199,253,10,15,grey],[55,507,10,16,grey],[221,505,151,17,white],[485,627,10,16,grey],[486,659,10,16,grey]
 ].forEach(b=>cover(page,...b,rgb,H))}
function checkType(page,font,type,H){const key=clean(type).toLowerCase(),x={'suscripción nueva':27,'adición':128,'creación subcuenta':190,'responsabilidad de pago':291,'renovación':416,'eq. prop. cliente':495}[key];if(x!==undefined)mark(page,font,x,83,H)}
function lines(text,font,size,width){const out=[];let line='';for(const word of String(text||'').split(/\s+/)){const test=line?line+' '+word:word;if(font.widthOfTextAtSize(test,size)>width&&line){out.push(line);line=word}else line=test}if(line)out.push(line);return out}
async function excelAnnex(pdf,raw,f,lib,fonts){let p=pdf.addPage([612,792]),y=54;const put=(value,size=8,bold=false)=>{const font=bold?fonts.bold:fonts.regular;for(const ln of lines(value,font,size,520)){if(y>745){p=pdf.addPage([612,792]);y=54}p.drawText(ln,{x:46,y:792-y-size,size,font,color:lib.rgb(.05,.12,.22)});y+=size+5}};put('ANEXO DE SERVICIOS Y EQUIPOS',15,true);y+=8;put('Cliente: '+f.suscriptor,10,true);put('RFC: '+f.rfc,9);put('Cuenta: '+f.accountNo,9);put('Contratación: '+f.tipoContratacion,9);put('Plazo: '+(f.mesesPlazo?f.mesesPlazo+' meses':''),9);y+=8;put('INFORMACIÓN DEL ARCHIVO FUENTE',11,true);String(raw||'').split(/\n/).filter(Boolean).forEach(x=>put(x,8))}
export async function generateContractPdf({templateFile,sourceFile,fields,seller}){if(!templateFile||!sourceFile)throw Error('Selecciona la plantilla PDF y la propuesta fuente.');if(!clean(seller?.name)||!clean(seller?.rfc))throw Error('Configura nombre y RFC del vendedor en Vendedores.');
 const lib=await import(PDFLIB),pdf=await lib.PDFDocument.load(await templateFile.arrayBuffer()),page=pdf.getPage(0),{width,height}=page.getSize(),regular=await pdf.embedFont(lib.StandardFonts.Helvetica),bold=await pdf.embedFont(lib.StandardFonts.HelveticaBold),f={...blank(),...fields};
 clearSample(page,lib.rgb,height);text(page,regular,f.accountNo,530,107,58,9,height,lib.rgb);text(page,regular,seller.name,285,105,184,9.5,height,lib.rgb);text(page,regular,String(seller.rfc).replace(/\s/g,'').toUpperCase(),371,121,86,8.5,height,lib.rgb);text(page,regular,f.suscriptor,22,200,410,9.5,height,lib.rgb);text(page,regular,f.rfc,445,200,108,8.5,height,lib.rgb);text(page,regular,f.telefono,482,219,68,8.5,height,lib.rgb);text(page,regular,f.domicilioFiscal,22,242,425,7.5,height,lib.rgb);text(page,regular,f.domicilioEntrega,22,263,334,7.5,height,lib.rgb);text(page,regular,f.authorizedPerson,22,284,330,8.5,height,lib.rgb);text(page,regular,f.representante,22,221,390,8.5,height,lib.rgb);text(page,regular,f.email,222,507,150,8.5,height,lib.rgb);if(f.tipoContratacion)checkType(page,bold,f.tipoContratacion,height);const term=Number(f.mesesPlazo);if(Number.isInteger(term)&&term>0&&term<=72){mark(page,bold,116,339,height);text(page,regular,String(term),198,339,35,8.5,height,lib.rgb)}
 const isPdf=/\.pdf$/i.test(sourceFile.name)||sourceFile.type==='application/pdf';if(isPdf){const annex=await lib.PDFDocument.load(await sourceFile.arrayBuffer());(await pdf.copyPages(annex,annex.getPageIndices())).forEach(p=>pdf.addPage(p))}else{const parsed=await parseContractSource(sourceFile);await excelAnnex(pdf,parsed.text,f,lib,{regular,bold})}
 const bytes=await pdf.save(),url=URL.createObjectURL(new Blob([bytes],{type:'application/pdf'})),a=document.createElement('a');a.href=url;a.download='Contrato_'+(f.suscriptor||'cliente').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Za-z0-9_-]+/g,'_')+'.pdf';document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url);return {pages:pdf.getPageCount()}}
