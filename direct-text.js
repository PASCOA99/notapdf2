/** Independent, low-overhead direct-text overlay for NotaPDF Studio.
 * Original pop-up notes, highlights and text-selection code are not replaced.
 * Direct-text boxes are temporary editing controls; only PDF content is exported.
 */
const clamp=(value,min,max)=>Math.max(min,Math.min(max,value));
const fontSizeOf=b=>clamp(Number(b.fontSize)||14,6,72);
const spacingOf=b=>[0.95,1,1.05,1.15].includes(Number(b.lineHeight))?Number(b.lineHeight):1;
const alignmentOf=b=>['left','center','right'].includes(b.align)?b.align:'left';
const colorOf=b=>/^#[0-9a-f]{6}$/i.test(b.color||'')?b.color:'#161616';

/** Word-wrap with the same PDF font used to draw selectable PDF text. */
export function wrapPdfText(text,font,fontSize,maxWidth){
 const out=[];
 const fits=s=>font.widthOfTextAtSize(s,fontSize)<=maxWidth+0.01;
 for(const paragraph of String(text).replace(/\r\n?/g,'\n').replace(/\t/g,'    ').split('\n')){
  if(!paragraph){out.push('');continue;}
  let line='';
  // Keep whitespace after words predictable. Long tokens fall back to per-character wrapping.
  const tokens=paragraph.match(/\S+\s*|\s+/gu)||[];
  for(const token of tokens){
   const candidate=line+token;
   if(fits(candidate)){line=candidate;continue;}
   if(line){out.push(line.replace(/\s+$/u,''));line='';}
   if(fits(token.trimStart())){line=token.trimStart();continue;}
   for(const char of Array.from(token.trimStart())){
    if(line&&!fits(line+char)){out.push(line);line='';}
    line+=char;
   }
  }
  out.push(line.replace(/\s+$/u,''));
 }
 return out;
}

/** Invert PDF.js's 0/90/180/270-degree screen rotation, including CropBox offsets. */
export function visualPointToPdf(x,y,crop,angle){
 const r=((angle%360)+360)%360;
 switch(r){
  case 0:return {x:crop.x+x,y:crop.y+crop.height-y};
  case 90:return {x:crop.x+y,y:crop.y+x};
  case 180:return {x:crop.x+crop.width-x,y:crop.y+y};
  case 270:return {x:crop.x+crop.width-y,y:crop.y+crop.height-x};
  default:throw new Error('Rotação de página não suportada para texto incorporado: '+angle+'°.');
 }
}

function canvasWrapText(ctx,text,maxWidth){
 const out=[];
 for(const paragraph of String(text).replace(/\r\n?/g,'\n').replace(/\t/g,'    ').split('\n')){
  if(!paragraph){out.push('');continue;}
  let line='';
  const tokens=paragraph.match(/\S+\s*|\s+/gu)||[];
  for(const token of tokens){
   if(ctx.measureText(line+token).width<=maxWidth+.01){line+=token;continue;}
   if(line){out.push(line.trimEnd());line='';}
   if(ctx.measureText(token.trimStart()).width<=maxWidth+.01){line=token.trimStart();continue;}
   for(const char of Array.from(token.trimStart())){
    if(line&&ctx.measureText(line+char).width>maxWidth){out.push(line);line='';}
    line+=char;
   }
  }
  out.push(line.trimEnd());
 }
 return out;
}

/** Add selectable/searchable PDF text, with transparent PNG fallback for unsupported Unicode. */
export async function burnDirectTextBoxes(doc,pages,boxes){
 if(!boxes?.length)return 0;
 const {StandardFonts,rgb,degrees}=window.PDFLib;
 const font=await doc.embedFont(StandardFonts.Helvetica);
 let completed=0;
 for(const box of boxes){
  if(!box.text?.trim())continue;
  const page=pages[box.page-1];
  if(!page)throw new Error(`A caixa de texto está vinculada a uma página inexistente (${box.page}).`);
  const angle=((page.getRotation()?.angle||0)%360+360)%360;
  const crop=page.getCropBox?.()||{x:0,y:0,...page.getSize()};
  const displayWidth=angle%180===0?crop.width:crop.height;
  const displayHeight=angle%180===0?crop.height:crop.width;
  const left=box.x*displayWidth,top=box.y*displayHeight;
  const width=box.width*displayWidth,height=box.height*displayHeight;
  const size=fontSizeOf(box),lineHeight=size*spacingOf(box),align=alignmentOf(box);
  if(width<8||height<8)throw new Error(`A caixa de texto da página ${box.page} é muito pequena.`);
  const color=colorOf(box);
  let native=true,lines;
  try{
   lines=wrapPdfText(box.text,font,size,width-1.5);
   for(const line of lines)font.encodeText(line);
  }catch(err){native=false;}
  if(native){
   if(lines.length*lineHeight>height+0.5)throw new Error(`Texto da página ${box.page} não cabe na altura da caixa. Aumente-a pela alça inferior direita ou reduza a fonte.`);
   const channels=[1,3,5].map(i=>parseInt(color.slice(i,i+2),16)/255);
   for(let i=0;i<lines.length;i++){
    const line=lines[i];if(!line)continue;
    const widthOfLine=font.widthOfTextAtSize(line,size);
    const inset=align==='center'?(width-widthOfLine)/2:align==='right'?width-widthOfLine:0;
    // 0.84em approximates the first-line baseline of the CSS Arial preview.
    const pos=visualPointToPdf(left+Math.max(0,inset),top+size*.84+i*lineHeight,crop,angle);
    page.drawText(line,{x:pos.x,y:pos.y,rotate:degrees(angle),font,size,color:rgb(...channels)});
   }
  }else{
   // Standard Helvetica's WinAnsi coverage includes Portuguese accents. For other
   // scripts/emoji, preserve actual glyphs as a transparent, permanently burnt image.
   const scale=2;
   const canvas=document.createElement('canvas');canvas.width=Math.ceil(width*scale);canvas.height=Math.ceil(height*scale);
   const ctx=canvas.getContext('2d');ctx.scale(scale,scale);ctx.font=`${size}px Arial, sans-serif`;
   const wrapped=canvasWrapText(ctx,box.text,width-1.5);
   if(wrapped.length*lineHeight>height+0.5)throw new Error(`Texto da página ${box.page} não cabe na altura da caixa. Aumente a caixa.`);
   ctx.fillStyle=color;ctx.textBaseline='alphabetic';ctx.textAlign=align;
   const anchor=align==='center'?width/2:align==='right'?width:0;
   for(let i=0;i<wrapped.length;i++)ctx.fillText(wrapped[i],anchor,size*.84+i*lineHeight);
   const image=await doc.embedPng(canvas.toDataURL('image/png'));
   // Origin is the displayed bottom-left corner; rotation keeps glyphs upright.
   const pos=visualPointToPdf(left,top+height,crop,angle);
   page.drawImage(image,{x:pos.x,y:pos.y,width,height,rotate:degrees(angle)});
  }
  completed++;
 }
 return completed;
}

export function createDirectTextController({state,pageContainer,layer,notify,commit}){
 let activeId=null;
 const boxes=()=>state.textBoxes||[];
 const view=()=>state.currentViewport;
 function refreshSelection(){
  for(const node of layer.querySelectorAll('.direct-text-box'))node.classList.toggle('active',node.dataset.textId===activeId);
 }
 function activate(id){activeId=id;refreshSelection();}
 function addBox(e){
  const viewport=view();if(!viewport||state.locked||state.placing)return;
  const rect=pageContainer.getBoundingClientRect();
  const atX=clamp((e.clientX-rect.left)/rect.width,0,1);
  const atY=clamp((e.clientY-rect.top)/rect.height,0,1);
  const width=clamp(260*state.zoom/viewport.width,.1,.8),height=clamp(90*state.zoom/viewport.height,.04,.5);
  const box={id:'t'+Array.from(crypto.getRandomValues(new Uint8Array(8)),v=>v.toString(16).padStart(2,'0')).join(''),
   page:state.page,x:Math.min(atX,1-width),y:Math.min(atY,1-height),width,height,
   text:'',fontSize:14,color:'#161616',align:'left',lineHeight:1};
  state.textBoxes.push(box);commit();activeId=box.id;
  draw(viewport);
  const field=[...layer.querySelectorAll('.direct-text-box')].find(n=>n.dataset.textId===box.id)?.querySelector('textarea');
  field?.focus();
  notify('Caixa de texto criada. Digite, arraste pela alça superior e redimensione no canto inferior direito.',5500);
 }
 function draw(viewport){
  layer.replaceChildren();if(!viewport)return;
  for(const box of boxes().filter(b=>b.page===state.page)){
   const node=document.createElement('div');node.className='direct-text-box';node.dataset.textId=box.id;
   function position(){
    node.style.left=(box.x*viewport.width)+'px';node.style.top=(box.y*viewport.height)+'px';
    node.style.width=(box.width*viewport.width)+'px';node.style.height=(box.height*viewport.height)+'px';
   }
   position();
   const textarea=document.createElement('textarea');textarea.className='direct-text-editor';textarea.value=box.text;
   textarea.setAttribute('aria-label','Texto incorporado na página '+box.page);textarea.spellcheck=true;
   function repaintText(){
    textarea.style.fontSize=fontSizeOf(box)*state.zoom+'px';
    textarea.style.lineHeight=String(spacingOf(box));
    textarea.style.color=colorOf(box);textarea.style.textAlign=alignmentOf(box);
   }
   repaintText();
   textarea.addEventListener('input',()=>{box.text=textarea.value;node.classList.toggle('text-overflow',textarea.scrollHeight>textarea.clientHeight+3)});
   textarea.addEventListener('blur',()=>commit());
   textarea.addEventListener('pointerdown',()=>activate(box.id));
   textarea.addEventListener('keydown',e=>{
    if(['ArrowLeft','ArrowRight','ArrowUp','ArrowDown','Escape'].includes(e.key)||((e.ctrlKey||e.metaKey)&&['z','y'].includes(e.key.toLowerCase())))e.stopPropagation();
    if(e.key==='Escape')textarea.blur();
   });
   const toolbar=document.createElement('div');toolbar.className='direct-text-toolbar';toolbar.setAttribute('aria-label','Formatar e mover caixa de texto');
   function toolTip(element,title){element.title=title;return element;}
   const grip=toolTip(document.createElement('button'),'Arraste para mover a caixa');grip.type='button';grip.className='direct-text-drag';grip.textContent='⠿';grip.setAttribute('aria-label','Arrastar a caixa de texto');
   toolbar.append(grip);
   function selector(title,values,current,handler){
    const select=document.createElement('select');select.title=title;select.setAttribute('aria-label',title);
    for(const [value,label] of values){const option=document.createElement('option');option.value=value;option.textContent=label;select.append(option)}
    select.value=String(current);select.addEventListener('change',()=>{handler(select.value);commit();repaintText()});toolbar.append(select);return select;
   }
   const size=document.createElement('input');size.type='number';size.min='6';size.max='72';size.step='1';size.value=String(fontSizeOf(box));size.title='Tamanho da fonte em pontos';size.setAttribute('aria-label','Tamanho da fonte em pontos');size.className='direct-text-size';
   size.addEventListener('change',()=>{box.fontSize=clamp(Number(size.value)||14,6,72);size.value=String(box.fontSize);repaintText();commit()});toolbar.append(size);
   const color=document.createElement('input');color.type='color';color.value=colorOf(box);color.title='Cor do texto';color.setAttribute('aria-label','Cor do texto');
   color.addEventListener('input',()=>{box.color=color.value;repaintText()});color.addEventListener('change',commit);toolbar.append(color);
   selector('Alinhamento do texto',[['left','Esq.'],['center','Centro'],['right','Dir.']],alignmentOf(box),value=>box.align=value);
   selector('Espaçamento entre linhas',[['0.95','0,95×'],['1','1,0×'],['1.05','1,05×'],['1.15','1,15×']],spacingOf(box),value=>box.lineHeight=Number(value));
   const remove=toolTip(document.createElement('button'),'Excluir caixa (desfazer disponível)');remove.type='button';remove.textContent='×';remove.setAttribute('aria-label','Excluir caixa de texto');remove.className='direct-text-remove';
   remove.addEventListener('click',()=>{state.textBoxes=boxes().filter(b=>b.id!==box.id);activeId=null;commit();draw(view());notify('Caixa excluída. Você pode desfazer a exclusão.',2400)});toolbar.append(remove);
   const resize=toolTip(document.createElement('button'),'Arraste para alterar largura e altura');resize.className='direct-text-resize';resize.type='button';resize.textContent='◢';resize.setAttribute('aria-label','Redimensionar caixa de texto');
   node.append(textarea,toolbar,resize);layer.append(node);
   node.addEventListener('pointerdown',()=>activate(box.id));
   function pointerGesture(handle,kind){
    handle.addEventListener('pointerdown',e=>{
     if(e.button!==0)return;e.preventDefault();e.stopPropagation();activate(box.id);
     const rect=pageContainer.getBoundingClientRect(),startX=e.clientX,startY=e.clientY;
     const start={x:box.x,y:box.y,width:box.width,height:box.height};
     handle.setPointerCapture(e.pointerId);
     const move=event=>{
      const dx=(event.clientX-startX)/rect.width,dy=(event.clientY-startY)/rect.height;
      if(kind==='move'){
       box.x=clamp(start.x+dx,0,1-box.width);box.y=clamp(start.y+dy,0,1-box.height);
      }else{
       const minW=Math.min(.98,Math.max(50*state.zoom/rect.width,.035));
       const minH=Math.min(.98,Math.max(22*state.zoom/rect.height,.02));
       box.width=clamp(start.width+dx,minW,1-box.x);
       box.height=clamp(start.height+dy,minH,1-box.y);
      }
      position();
     };
     const finish=()=>{handle.removeEventListener('pointermove',move);handle.removeEventListener('pointerup',finish);handle.removeEventListener('pointercancel',finish);commit();node.classList.toggle('text-overflow',textarea.scrollHeight>textarea.clientHeight+3)};
     handle.addEventListener('pointermove',move);handle.addEventListener('pointerup',finish);handle.addEventListener('pointercancel',finish);
    });
   }
   pointerGesture(grip,'move');pointerGesture(resize,'resize');
  }
  refreshSelection();
 }
 pageContainer.addEventListener('contextmenu',e=>{
  if(!state.pdf||state.locked||state.placing||e.target.closest('.note-marker,.note-popup,.old-annotation,.direct-text-box'))return;
  e.preventDefault();e.stopPropagation();addBox(e);
 });
 pageContainer.addEventListener('pointerdown',e=>{
  if(!e.target.closest('.direct-text-box')){activeId=null;refreshSelection();}
 },true);
 return {draw,close:()=>{activeId=null;layer.replaceChildren()},burn:burnDirectTextBoxes};
}
