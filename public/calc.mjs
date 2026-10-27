export const money = (n) => new Intl.NumberFormat('es-MX', { style:'currency', currency:'MXN', maximumFractionDigits:2 }).format(Number(n)||0);
export const nval = (v) => Number.isFinite(Number(v)) ? Number(v) : 0;
export const validHttps = (s) => { try {const u=new URL(s);return u.protocol==='https:' && Boolean(u.hostname);} catch {return false;} };
export const parseSafePrice = (v) => { const n=Number(v); if(!Number.isFinite(n)||n<0||n>10000000) throw new Error('Precio inválido'); return Math.round(n*100)/100; };
export const priceForTerm = (device,months) => Number(device?.prices?.[String(months)]??0);
export function calculateQuote({plan,device,termMonths,quantity=1,controlMonthly=0}) {
  const months=Number(termMonths), qty=Number(quantity);
  if(!plan||!device) throw new Error('Selecciona plan y equipo');
  if(!Number.isInteger(months)||months<1||months>72)throw new Error('Plazo inválido');
  if(!Number.isInteger(qty)||qty<1||qty>1000)throw new Error('Cantidad inválida');
  const rent=parseSafePrice(plan.monthly), cost=parseSafePrice(priceForTerm(device,months)), control=parseSafePrice(controlMonthly);
  if(!Object.prototype.hasOwnProperty.call(device.prices||{},String(months))) throw new Error('Este equipo no tiene precio configurado para el plazo seleccionado');
  const eqMonthly=Math.round(cost/months*100)/100;
  return {months,qty,planMonthly:rent,deviceTotal:cost,deviceMonthly:eqMonthly,controlMonthly:control,unitMonthly:Math.round((rent+eqMonthly+control)*100)/100,totalMonthly:Math.round((rent+eqMonthly+control)*qty*100)/100};
}
export const escapeHtml = (v='') => String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
export const stageOptions = ['Nuevo','Contactado','Interesado','Cotización enviada','Negociación','Ganado','Perdido'];
