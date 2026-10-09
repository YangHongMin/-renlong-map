const OWNER=(()=>{try{return localStorage.getItem("is_owner")==="1"}catch(e){return false}})(); // 管理者本人的瀏覽不計入統計（在管理頁登入時設定）
const GA_ID="G-5YXFVRJ2JM",PIXEL_ID="996673046774269";
const esc=s=>String(s==null?"":s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const okId=s=>typeof s==="string"&&/^[A-Za-z0-9_-]{1,40}$/.test(s);
const okImg=s=>typeof s==="string"&&/^img\/[A-Za-z0-9_\-\/.]+$/.test(s)&&s.indexOf("..")<0;
const sanitize=d=>(Array.isArray(d)?d:[]).filter(l=>l&&typeof l==="object"&&okId(l.id)).map(l=>({...l,photos:(Array.isArray(l.photos)?l.photos:[]).filter(okImg)}));
window.dataLayer=window.dataLayer||[];
function gtag(){dataLayer.push(arguments)}
if(!OWNER){
  const gs=document.createElement("script");gs.async=true;gs.src="https://www.googletagmanager.com/gtag/js?id="+GA_ID;document.head.appendChild(gs);
  gtag("js",new Date());gtag("config",GA_ID,{send_page_view:false});
  !function(f,b,e,v,n,t,s){if(f.fbq)return;n=f.fbq=function(){n.callMethod?n.callMethod.apply(n,arguments):n.queue.push(arguments)};if(!f._fbq)f._fbq=n;n.push=n;n.loaded=!0;n.version="2.0";n.queue=[];t=b.createElement(e);t.async=!0;t.src=v;s=b.getElementsByTagName(e)[0];s.parentNode.insertBefore(t,s)}(window,document,"script","https://connect.facebook.net/en_US/fbevents.js");
  fbq("init",PIXEL_ID);fbq("track","PageView");
}
const STD=["ViewContent","Lead","Contact","PageView"];
function ev(ga,gp,fb,fp){if(OWNER)return;try{if(ga)gtag("event",ga,gp||{});if(fb)fbq(STD.includes(fb)?"track":"trackCustom",fb,fp||{})}catch(e){}}
let firstView=true;
function pageview(l){
  ev("page_view",{page_title:document.title,page_location:location.href});
  if(!firstView)ev(null,null,"PageView"); firstView=false;  // 第一次的 PageView 已由 Pixel 載入時送出
  if(l)ev("view_item",{items:[{item_id:l.id,item_name:l.title}]},"ViewContent",{content_ids:[l.id],content_name:l.title,content_type:"product"});
}
const LINE_URL="https://lin.ee/S6hfHqge",FB_URL="https://www.facebook.com/profile.php?id=61573837941258";
const joinBlock=()=>`<div class="join"><div><h3>新物件・降價通知</h3><p>加入官方 LINE，第一時間收到新上架與行情；也可以直接告訴我您的需求（區域、坪數、預算），我幫您留意。</p><div class="btns"><a class="btn line" href="${LINE_URL}" target="_blank" rel="noopener" onclick="ev('join_line_click',{item_id:curId},'Lead',{content_name:'join'})">加入 LINE 官方帳號</a><a class="btn fb" href="${FB_URL}" target="_blank" rel="noopener">追蹤 Facebook 粉絲專頁</a></div></div><img class="qr" src="line_qr.png" alt="LINE 官方帳號 QR Code"></div>`;
const needBlock=()=>`<div class="need" id="need"><h3>找不到合適的？告訴我您的需求</h3><p>填寫後會整理成一則訊息，由您確認後傳到 LINE。有符合的物件，我會第一時間通知您。不需要留電話。</p>
 <div class="ng"><label>我想<select id="n_t"><option>買廠房／土地</option><option>租廠房／土地</option><option>委託出售</option><option>委託出租</option></select></label>
 <label>區域<input id="n_a" maxlength="40" placeholder="例如：仁武、大寮、岡山" autocomplete="off"></label>
 <label>坪數<input id="n_s" maxlength="30" placeholder="例如：土地300坪左右" autocomplete="off"></label>
 <label>預算／價格<input id="n_b" maxlength="30" placeholder="例如：月租50萬內" autocomplete="off"></label>
 <label>用途<input id="n_u" maxlength="30" placeholder="例如：製造、倉儲、貨櫃場" autocomplete="off"></label>
 <label>其他需求<input id="n_m" maxlength="100" placeholder="例如：要天車、貨櫃車可進出" autocomplete="off"></label></div>
 <div class="btns"><button class="btn line" type="button" id="n_go">整理成訊息並開啟 LINE</button></div><div class="nout" id="n_out" style="display:none"></div></div>`;
const CATS=["大型廠房","小型廠房","土地","其他"];
// 顯示用價格：1 億以上的「XXXXX萬」換成「約X.XX億元」（與 scripts/build.py 的 price_disp 相同規則）
const priceOf=l=>{const raw=String(l.price||"").trim(),m=raw.match(/^(?:總價)?\s*([\d,]+(?:\.\d+)?)\s*萬/);if(!m)return raw;const v=parseFloat(m[1].replace(/,/g,""));if(v<10000)return raw;const rest=raw.slice(m[0].length).trim();return "約"+(Math.round(v/100)/100)+"億元"+(rest?" "+rest:"")};
const catOf=l=>{if(CATS.includes(l.category))return l.category;const b=parseFloat(String(l.build_ping||"").replace(/,/g,""))||0,d=parseFloat(String(l.land_ping||"").replace(/,/g,""))||0,t=String(l.title||"")+String(l.zoning||"");
  if(!b)return d?"土地":"其他";if(!/廠|倉|工業|丁種|乙種|甲種/.test(t))return "其他";return b>=500?"大型廠房":"小型廠房"};
let all=[],F={type:"",area:"",kw:"",cat:""},msg="我想詢問工業物件";
const vis=()=>all.filter(l=>l.status!=="待確認");
const $=document.getElementById("app");
const dist=l=>(String(l.area||"").match(/[市縣](.+?[區鄉鎮市])/)||[])[1]||"";
let curId="";
function toast(x){const t=document.getElementById("toast");t.textContent=x;t.style.display="block";clearTimeout(toast.t);toast.t=setTimeout(()=>t.style.display="none",3500)}
const copy=t=>{try{return navigator.clipboard.writeText(t).then(()=>true,()=>false)}catch(e){return Promise.resolve(false)}};
function lineGo(w){copy(msg);toast("已複製物件編號，請在 LINE 貼上傳送");ev("line_click",{where:w||"page",item_id:curId},"Lead",{content_name:msg})}
const shareUrl=()=>location.origin+(curId?"/p/"+curId+".html":"/");
function share(e){e.preventDefault();const m=document.getElementById("smenu");m.style.display=m.style.display==="block"?"none":"block";document.getElementById("smore").style.display=navigator.share?"block":"none"}
function sfb(){ev("share",{method:"facebook",item_id:curId});window.open("https://www.facebook.com/sharer/sharer.php?u="+encodeURIComponent(shareUrl()),"_blank","noopener");document.getElementById("smenu").style.display="none"}
function sln(){ev("share",{method:"line",item_id:curId});window.open("https://social-plugins.line.me/lineit/share?url="+encodeURIComponent(shareUrl()),"_blank","noopener");document.getElementById("smenu").style.display="none"}
function scp(){ev("share",{method:"copy",item_id:curId});copy(shareUrl());document.getElementById("smenu").style.display="none";toast("已複製連結")}
function smore(){ev("share",{method:"native",item_id:curId});navigator.share({title:document.title,url:shareUrl()}).catch(()=>{});document.getElementById("smenu").style.display="none"}
function needGo(){
  const g=id=>document.getElementById(id).value.trim();
  const t=g("n_t"),a=g("n_a"),s=g("n_s"),b=g("n_b"),u=g("n_u"),m=g("n_m");
  if(!(a||s||b||u||m)){toast("請至少填寫一項需求");return}
  const text=["【需求登記】","我想："+t,a&&"區域："+a,s&&"坪數："+s,b&&"預算／價格："+b,u&&"用途："+u,m&&"其他需求："+m,"（來自網站 fulllife.blog）"].filter(Boolean).join("\n");
  copy(text).then(ok=>toast(ok?"已複製需求內容，請在 LINE 貼上傳送":"請手動複製下方內容，再到 LINE 貼上"));
  window.open(LINE_URL,"_blank","noopener");
  const o=document.getElementById("n_out");o.style.display="block";
  o.innerHTML='<b>需求已整理好</b><div>請到 LINE 的對話框貼上並傳送。若沒有自動複製，請手動複製：</div><textarea id="n_txt" readonly rows="7"></textarea>';
  const tx=document.getElementById("n_txt");tx.value=text;tx.onclick=()=>tx.select();
  ev("need_form_submit",{need_type:t},"NeedForm",{need_type:t}); // 只送需求類別，不送內容
}
function toNeed(e){e.preventDefault();history.pushState({},"","./#need");route()}
fetch("data/listings.json?v="+Math.floor(Date.now()/60000)).then(r=>r.json()).then(d=>{all=sanitize(d);route()}).catch(()=>{$.textContent="資料載入失敗，請稍後再試。"});
window.addEventListener("popstate",route);
function route(){const q=new URLSearchParams(location.search),id=q.get("id"),pv=q.get("preview")==="1";const l=all.find(x=>x.id===id&&(x.status!=="待確認"||pv));l?detail(l):list();pageview(l)}
function go(e,id){e.preventDefault();history.pushState({},"","?id="+id);route();scrollTo(0,0)}
function home(e){e.preventDefault();history.pushState({},"","./");route()}
function list(){
  document.title="富住通大型工業地產｜高雄工業廠房";msg="我想詢問工業物件";curId="";
  const areas=[...new Set(vis().map(dist).filter(Boolean))];
  $.innerHTML=`<h1>工業廠房・工業用地</h1><div class="chips" id="chips"></div><div class="panel">
   <label>類別：</label><select id="ft"><option value="">不限</option><option value="售">出售</option><option value="租">出租</option></select>
   <label>區域：</label><select id="fa"><option value="">不限</option>${areas.map(a=>`<option>${esc(a)}</option>`).join("")}</select>
   <label>關鍵字：</label><input id="fk" placeholder="輸入關鍵字"><button onclick="draw()">我要查詢</button></div><div class="grid" id="grid"></div>${needBlock()}${joinBlock()}`;
  ["ft","fa"].forEach(i=>{const e=document.getElementById(i);e.value=F[i==="ft"?"type":"area"];e.onchange=draw});
  const k=document.getElementById("fk");k.value=F.kw;k.oninput=draw;draw();
  document.getElementById("n_go").onclick=needGo;
  if(location.hash==="#need")document.getElementById("need").scrollIntoView();
}
function chips(){
  const v=vis(),n=c=>c?v.filter(l=>catOf(l)===c).length:v.length;
  document.getElementById("chips").innerHTML=["",...CATS].map(c=>`<button type="button" class="${F.cat===c?"on":""}" data-c="${esc(c)}">${c||"全部"}（${n(c)}）</button>`).join("");
  document.querySelectorAll("#chips button").forEach(b=>b.onclick=()=>{F.cat=b.dataset.c;draw()});
}
function draw(){
  chips();F.type=document.getElementById("ft").value;F.area=document.getElementById("fa").value;F.kw=document.getElementById("fk").value.trim();
  const rows=vis().filter(l=>(!F.type||l.type===F.type)&&(!F.cat||catOf(l)===F.cat)&&(!F.area||dist(l)===F.area)&&(!F.kw||[l.title,l.area,l.zoning,l.desc,l.note].join(" ").includes(F.kw)));
  document.getElementById("grid").innerHTML=rows.map(l=>{const p=(l.photos||[])[0];const done=l.status==="已成交";
   return `<a class="card" href="?id=${esc(l.id)}" onclick="go(event,'${esc(l.id)}')"><div class="ph" style="${p?`background-image:url('${esc(p)}')`:""}">${p?"":"🏭"}<span class="tag ${done?"done":l.type==="租"?"rent":""}">${done?"已成交":"出"+esc(l.type)}</span></div>
   <div class="info"><h3>${esc(l.title)}</h3><div class="meta"><span class="cat">${esc(catOf(l))}</span>${esc([l.area,l.zoning,l.land_ping?"土地 "+l.land_ping+" 坪":"",l.build_ping?"建坪 "+l.build_ping+" 坪":""].filter(Boolean).join("｜"))}</div><div class="price">${esc(priceOf(l))}</div></div></a>`}).join("")||"目前沒有符合的物件";
}
function detail(l){
  document.title=String(l.title||"")+"｜富住通大型工業地產";msg="我想詢問物件 "+l.id+"｜"+l.title;curId=l.id;
  const rows=[["編號",l.id],["類別",catOf(l)],["區域",l.area],["價格",priceOf(l)],["使用分區",l.zoning],["基地面積",l.land_ping?l.land_ping+" 坪":""],["建物面積",l.build_ping?l.build_ping+" 坪":""],["備註",l.note]].filter(r=>r[1]);
  $.innerHTML=`<a class="back" href="./" onclick="home(event)">← 回物件列表</a><h1>${esc(l.title)}</h1>
  <div class="gal">${(l.photos||[]).map(p=>`<img src="${esc(p)}" loading="lazy" alt="${esc(l.title)}">`).join("")}</div>
  <table>${rows.map(r=>`<tr><th>${esc(r[0])}</th><td>${esc(r[1])}</td></tr>`).join("")}</table>
  <p>${String(l.desc||"").split("；").filter(Boolean).map(x=>"・"+esc(x)).join("<br>")}</p>
  ${l.status==="已成交"?"":`<div class="contact"><img src="avatar.jpg" alt="楊紘珉"><div><b>楊紘珉</b><div class="sub">富住通商用不動產｜大型工業地產</div>
   <div>0905-858-141｜0978-133-561</div><div class="btns"><a class="btn line" href="${LINE_URL}" target="_blank" rel="noopener" onclick="lineGo('detail')">LINE 詢問這個物件</a><a class="btn tel" href="tel:0905858141" onclick="ev('call_click',{item_id:curId},'Contact')">撥打電話</a></div>
   <div class="sub" style="margin-top:6px">按 LINE 會自動複製物件編號，到 LINE 貼上傳送即可。</div></div></div>`}
  <p><a href="./#need" onclick="toNeed(event)" style="color:var(--blue)">找不到合適的？告訴我您的需求</a></p>${joinBlock()}`;
}
