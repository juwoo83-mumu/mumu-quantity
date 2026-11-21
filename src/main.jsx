
import React,{useEffect,useMemo,useState} from 'react';
import{createRoot}from'react-dom/client';
import{createClient}from'@supabase/supabase-js';
import{Camera,LayoutDashboard,ClipboardList,History,Users,LogOut,Plus,Minus,Download,RefreshCw}from'lucide-react';
import{createWorker}from'tesseract.js';
import*as XLSX from'xlsx';
import './style.css';
const url=import.meta.env.VITE_SUPABASE_URL,key=import.meta.env.VITE_SUPABASE_ANON_KEY;
const db=url&&key?createClient(url,key):null;
const empty=()=>({address:'',duvet:0,small:0,large:0});
const time=e=>new Date(e.created_at).toLocaleString('ko-KR',{timeZone:'Asia/Seoul',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit'});
const period=e=>new Date(e.created_at).toLocaleString('en-US',{timeZone:'Asia/Seoul',hour:'2-digit',hourCycle:'h23'}).split(':')[0]<12?'오전':'오후';
const dateKST=()=>new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Seoul',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
const shiftKST=(day,offset)=>{const d=new Date(day+'T12:00:00+09:00');d.setUTCDate(d.getUTCDate()+offset);return new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Seoul',year:'numeric',month:'2-digit',day:'2-digit'}).format(d)};
const monthAgoKST=day=>{const [y,m,d]=day.split('-').map(Number);const first=new Date(Date.UTC(y,m-2,1));const maxDay=new Date(Date.UTC(first.getUTCFullYear(),first.getUTCMonth()+1,0)).getUTCDate();first.setUTCDate(Math.min(d,maxDay));return first.toISOString().slice(0,10)};
const sum=rows=>rows.reduce((a,r)=>({duvet:a.duvet+r.duvet,small:a.small+r.small,large:a.large+r.large}),{duvet:0,small:0,large:0});
const coverPrice=n=>n<=0?0:n===1?12000:15000+Math.floor((n-2)/2)*7000+((n-2)%2)*3000;
const won=n=>n.toLocaleString('ko-KR')+'원';
const people=rows=>new Set(rows.map(x=>x.employee_id)).size;
function App(){
const[session,setSession]=useState(null),[ready,setReady]=useState(false),[profile,setProfile]=useState(null),[tab,setTab]=useState('entry');
const[rows,setRows]=useState([]),[changes,setChanges]=useState([]),[staff,setStaff]=useState([]);
const[startDate,setStartDate]=useState(dateKST()),[endDate,setEndDate]=useState(dateKST()),[preset,setPreset]=useState('today'),[person,setPerson]=useState('all'),[shift,setShift]=useState('전체'),[addressSearch,setAddressSearch]=useState('');
const[form,setForm]=useState(empty()),[editing,setEditing]=useState(null),[reason,setReason]=useState('주소지 수정'),[note,setNote]=useState('');
const[selectedIds,setSelectedIds]=useState([]),[copyNotice,setCopyNotice]=useState('');
const[busy,setBusy]=useState(false),[notice,setNotice]=useState(''),[authMode,setAuthMode]=useState('login'),[email,setEmail]=useState(''),[pass,setPass]=useState(''),[confirmPass,setConfirmPass]=useState(''),[name,setName]=useState('');
useEffect(()=>{if(!db){setReady(true);return}db.auth.getSession().then(({data})=>{setSession(data.session);setReady(true)});const{data:{subscription}}=db.auth.onAuthStateChange((_event,s)=>setSession(s));return()=>subscription.unsubscribe()},[]);
useEffect(()=>{if(!session){setProfile(null);return}db.from('profiles').select('*').eq('id',session.user.id).single().then(({data,error})=>{setProfile(data);if(error)setNotice(error.message)})},[session]);
const owner=profile?.role==='owner';
async function load(){if(!profile?.approved)return;setBusy(true);
if(!startDate||!endDate||startDate>endDate){setNotice('조회 시작일은 종료일보다 늦을 수 없습니다.');setBusy(false);return}
const from=startDate+'T00:00:00+09:00',to=new Date(new Date(endDate+'T00:00:00+09:00').getTime()+86400000).toISOString();
const[a,b,c]=await Promise.all([
 fetchEntries(from,to),
 db.from('changes').select('*').order('created_at',{ascending:false}).limit(1000),
 owner?db.from('profiles').select('*').order('name'):Promise.resolve({data:[]})
]);
if(a.error)setNotice(a.error.message);else setRows(a.data||[]);
if(b.error)setNotice(b.error.message);else setChanges(b.data||[]);
if(c.error)setNotice(c.error.message);else setStaff(c.data||[]);
setBusy(false)}
useEffect(()=>{setSelectedIds([]);setCopyNotice('');if(profile?.approved)load()},[profile,startDate,endDate]);
function choosePreset(type){
 const today=dateKST();setPreset(type);setEndDate(today);
 setStartDate(type==='today'?today:type==='twoWeeks'?shiftKST(today,-13):monthAgoKST(today));
}
async function fetchEntries(from,to){
 let all=[],offset=0;
 while(true){
  const res=await db.from('entries').select('*').gte('created_at',from).lt('created_at',to).order('created_at',{ascending:false}).order('id',{ascending:false}).range(offset,offset+999);
  if(res.error)return res;
  all.push(...(res.data||[]));
  if(!res.data||res.data.length<1000)break;
  offset+=1000;
 }
 return {data:all,error:null};
}
const names=Object.fromEntries(staff.map(x=>[x.id,x.name]));if(profile)names[profile.id]=profile.name;
const filtered=rows.filter(r=>(person==='all'||r.employee_id===person)&&(shift==='전체'||period(r)===shift)&&r.address.toLocaleLowerCase().includes(addressSearch.trim().toLocaleLowerCase()));
const visibleIds=new Set(filtered.map(r=>String(r.id)));
const selectedRows=filtered.filter(r=>selectedIds.includes(String(r.id)));
const selectedTotal=sum(selectedRows),selectedCoverPrice=coverPrice(selectedTotal.duvet),selectedSmallPrice=selectedTotal.small*700,selectedLargePrice=selectedTotal.large*900,selectedAmount=selectedCoverPrice+selectedSmallPrice+selectedLargePrice;
const selectedCopyText=[selectedTotal.duvet>0?`이불 ${selectedTotal.duvet/2}세트`:null,selectedTotal.small>0?`수건 소 ${selectedTotal.small}`:null,selectedTotal.large>0?`${selectedTotal.small>0?"대":"수건 대"} ${selectedTotal.large}`:null].filter(Boolean).join(', ');
const selectedSummary=`무무베딩 작업내역 정산\n선택 작업 ${selectedRows.length}건\n이불커버 ${selectedTotal.duvet}개 (${selectedTotal.duvet/2}세트): ${won(selectedCoverPrice)}\n수건 소 ${selectedTotal.small}개 × 700원: ${won(selectedSmallPrice)}\n수건 대 ${selectedTotal.large}개 × 900원: ${won(selectedLargePrice)}\n총 금액: ${won(selectedAmount)}`;
function toggleSelected(id){setSelectedIds(prev=>prev.includes(String(id))?prev.filter(x=>x!==String(id)):[...prev,String(id)]);setCopyNotice('')}
function toggleAll(list){const ids=list.map(r=>String(r.id));setSelectedIds(prev=>ids.every(id=>prev.includes(id))?prev.filter(id=>!ids.includes(id)):[...new Set([...prev,...ids])]);setCopyNotice('')}
async function copySummary(text){try{await navigator.clipboard.writeText(text);setCopyNotice('복사했습니다.')}catch{setCopyNotice('복사가 차단되었습니다. 아래 정산 내용을 직접 선택해 복사해 주세요.')}}
const total=sum(filtered),am=filtered.filter(r=>period(r)==='오전'),pm=filtered.filter(r=>period(r)==='오후');
async function login(e){e.preventDefault();setNotice('');if(authMode==='signup'&&(pass.length<8||pass!==confirmPass)){setNotice(pass.length<8?'비밀번호는 8자 이상 입력해 주세요.':'비밀번호 확인이 일치하지 않습니다.');return}setBusy(true);const{error}=authMode==='signup'?await db.auth.signUp({email,password:pass,options:{data:{name}}}):await db.auth.signInWithPassword({email,password:pass});setBusy(false);setNotice(error?error.message:authMode==='signup'?'가입 신청이 접수되었습니다. 이메일 인증 후 사업주 승인을 기다려 주세요.':'')}
async function logout(){await db.auth.signOut();setTab('entry')}
async function photo(file){
 if(!file)return;
 setBusy(true);setNotice('사진을 분석 중입니다. 주소가 있는 부분을 크게 촬영하면 인식률이 높아집니다.');
 let worker;
 try{
  const image=await createImageBitmap(file);
  const scale=Math.min(3,Math.max(1,1600/image.width));
  const canvas=document.createElement('canvas');canvas.width=Math.round(image.width*scale);canvas.height=Math.round(image.height*scale);
  const ctx=canvas.getContext('2d',{willReadFrequently:true});ctx.drawImage(image,0,0,canvas.width,canvas.height);
  const pixels=ctx.getImageData(0,0,canvas.width,canvas.height);
  for(let i=0;i<pixels.data.length;i+=4){
   const g=.299*pixels.data[i]+.587*pixels.data[i+1]+.114*pixels.data[i+2];
   const adjusted=Math.max(0,Math.min(255,(g-128)*1.35+128));
   pixels.data[i]=pixels.data[i+1]=pixels.data[i+2]=adjusted;
  }
  ctx.putImageData(pixels,0,0);image.close();
  worker=await createWorker('kor+eng');
  const {data}=await worker.recognize(canvas);
  const lines=data.text.split(/\r?\n/).map(x=>x.trim()).filter(Boolean);
  const addressPattern=/(서울|경기|인천|부산|대구|광주|대전|울산|세종|제주|강원|충북|충남|전북|전남|경북|경남|특별시|광역시|특별자치|[가-힣]+[시군구]|[가-힣]+(?:로|길|동|읍|면|리)\s*\d)/;
  const matched=lines.filter(x=>addressPattern.test(x));
  const text=(matched.length?matched:lines).join(' ').replace(/\s+/g,' ').trim();
  setForm(v=>({...v,address:text}));
  setNotice(text?'인식된 주소를 반드시 확인·수정한 후 등록해 주세요.':'주소를 인식하지 못했습니다. 주소 부분만 가까이 촬영하거나 직접 입력해 주세요.');
 }catch(err){setNotice('사진 인식에 실패했습니다. 주소를 직접 입력해 주세요.');}
 finally{if(worker)await worker.terminate();setBusy(false);}
}
async function save(e){e.preventDefault();if([...form.address.trim()].length<2){setNotice('주소지는 2글자 이상 입력해 주세요.');return}setBusy(true);
const clean={address:form.address.trim(),duvet:Number(form.duvet),small:Number(form.small),large:Number(form.large)};
let error;if(editing){({error}=await db.rpc('revise_entry',{p_id:editing.id,p_address:clean.address,p_duvet:clean.duvet,p_small:clean.small,p_large:clean.large,p_reason:reason,p_note:note}))}
else{({error}=await db.from('entries').insert({...clean,employee_id:session.user.id}))}
setBusy(false);if(error){setNotice(error.message);return}setNotice(editing?'수정 내용이 반영되고 이전 수량이 이력에 보존되었습니다.':'수량이 등록되었습니다.');setEditing(null);setForm(empty());setNote('');await load()}
function startEdit(r){setEditing(r);setForm({address:r.address,duvet:r.duvet,small:r.small,large:r.large});setReason('주소지 수정');setNote('');setTab('entry');window.scrollTo({top:0,behavior:'smooth'})}
async function approve(p){const{error}=await db.from('profiles').update({approved:!p.approved}).eq('id',p.id);if(error)setNotice(error.message);else load()}
function exportXlsx(){const records=filtered.map(r=>({등록시간:time(r),시간대:period(r),직원:names[r.employee_id]||r.employee_id,주소:r.address,이불커버:r.duvet,'수건 소':r.small,'수건 대':r.large}));const sheet=XLSX.utils.json_to_sheet(records),book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,sheet,'작업내역');XLSX.writeFile(book,`무무베딩_${startDate}_${endDate}.xlsx`)}
const entryTable=(list,allowEdit=false,selectable=false)=><div className="scroll"><table><thead><tr>{selectable&&<th><input type="checkbox" aria-label="현재 검색 결과 전체 선택" checked={list.length>0&&list.every(r=>selectedIds.includes(String(r.id)))} onChange={()=>toggleAll(list)}/></th>}<th>등록시간</th><th>직원</th><th>주소지</th><th>이불커버</th><th>수건 소</th><th>수건 대</th>{allowEdit&&<th>수정</th>}</tr></thead><tbody>{list.map(r=><tr key={r.id} className={selectable&&selectedIds.includes(String(r.id))?'selected-entry':''}>{selectable&&<td><input type="checkbox" aria-label={`${r.address} 작업 선택`} checked={selectedIds.includes(String(r.id))} onChange={()=>toggleSelected(r.id)}/></td>}<td>{time(r)}</td><td>{names[r.employee_id]||'직원'}</td><td>{r.address}</td><td>{r.duvet}</td><td>{r.small}</td><td>{r.large}</td>{allowEdit&&<td><button onClick={()=>startEdit(r)}>수정</button></td>}</tr>)}</tbody></table>{!list.length&&<p className="muted">등록된 내역이 없습니다.</p>}</div>;
const billingPanel=<section className="panel billing-panel"><div className="billing-heading"><div><h2>선택 작업내역 정산</h2><p className="muted">아래 작업내역에서 체크한 항목만 합산합니다. 기간이나 검색 조건을 변경하면 선택이 초기화됩니다.</p></div><button type="button" onClick={()=>{setSelectedIds([]);setCopyNotice('')}} disabled={!selectedIds.length}>선택 해제</button></div><div className="billing-grid"><div><span>선택 작업</span><strong>{selectedRows.length}건</strong></div><div><span>이불커버</span><strong>{selectedTotal.duvet}개 · {selectedTotal.duvet/2}세트</strong><small>{won(selectedCoverPrice)}</small></div><div><span>수건 소</span><strong>{selectedTotal.small}개</strong><small>{won(selectedSmallPrice)}</small></div><div><span>수건 대</span><strong>{selectedTotal.large}개</strong><small>{won(selectedLargePrice)}</small></div></div><div className="billing-bottom"><div><span>총 청구금액</span><strong>{won(selectedAmount)}</strong></div><div className="billing-actions"><button type="button" className="billing-copy" disabled={!selectedRows.length} onClick={()=>copySummary(selectedCopyText)}>수량 복사</button></div></div>{copyNotice&&<p role="status" className="muted">{copyNotice}</p>}<textarea className="billing-preview" readOnly value={selectedCopyText} aria-label="복사용 수량"/></section>;
if(!db)return <div className="center"><div className="login"><h1>무무베딩 수량관리</h1><p>클라우드 연결 전입니다. 프로젝트의 README 안내에 따라 Supabase 환경변수를 설정해 주세요.</p></div></div>;
if(!ready)return <div className="center">불러오는 중...</div>;
if(!session)return <div className="center"><form className="login" onSubmit={login}><div className="brand">MUMU BEDDING</div><h1>{authMode==='login'?'수량관리 로그인':'직원 회원가입'}</h1><p>현장 수량 기록과 실시간 운영 현황</p>{authMode==='signup'&&<label>직원 이름<input required value={name} onChange={e=>setName(e.target.value)} placeholder="이름"/></label>}<label>이메일<input type="email" required value={email} onChange={e=>setEmail(e.target.value)} placeholder="name@example.com"/></label><label>비밀번호<input type="password" minLength={authMode==='signup'?8:6} required value={pass} onChange={e=>setPass(e.target.value)} placeholder={authMode==='signup'?'8자 이상 입력':'비밀번호 입력'}/></label>{authMode==='signup'&&<><label>비밀번호 확인<input type="password" minLength="8" required value={confirmPass} onChange={e=>setConfirmPass(e.target.value)} placeholder="비밀번호 다시 입력"/></label>{confirmPass&&<p className={pass===confirmPass?'password-ok':'password-error'}>{pass===confirmPass?'✓ 비밀번호가 일치합니다.':'비밀번호가 일치하지 않습니다.'}</p>}</>}<button className="primary" disabled={busy||(authMode==='signup'&&(pass.length<8||pass!==confirmPass))}>{authMode==='login'?'로그인':'가입 신청'}</button><button type="button" className="plain" onClick={()=>{setAuthMode(authMode==='login'?'signup':'login');setPass('');setConfirmPass('');setNotice('')}}>{authMode==='login'?'신규 직원 회원가입':'로그인으로 돌아가기'}</button>{notice&&<div className="notice">{notice}</div>}</form></div>;
if(!profile)return <div className="center">계정 정보를 불러오는 중...</div>;
if(!profile.approved)return <div className="center"><div className="login"><div className="brand">MUMU BEDDING</div><h2>사업주 승인 대기 중</h2><p>{profile.name}님의 계정이 생성되었습니다. 사업주 승인 후 수량을 등록할 수 있습니다.</p><button onClick={()=>location.reload()}>승인 상태 새로고침</button><button className="plain" onClick={logout}>로그아웃</button></div></div>;
const tabs=owner?[['dashboard','대시보드',LayoutDashboard],['entries','작업내역',ClipboardList],['stats','오전·오후 통계',LayoutDashboard],['history','수정이력',History],['staff','직원관리',Users]]:[['entry','수량 등록',Plus],['mine','내 작업내역',ClipboardList],['history','수정이력',History]];
const shownChanges=changes.filter(x=>owner||x.changed_by===session.user.id);
return <div className="shell"><header><div><div className="brand">MUMU BEDDING</div><h1>수량관리 시스템</h1></div><div className="user"><span>{profile.name} {owner?'· 사업주':''}</span><button title="로그아웃" onClick={logout}><LogOut size={17}/></button></div></header><nav>{tabs.map(([id,label,Icon])=><button key={id} className={tab===id?'selected':''} onClick={()=>setTab(id)}><Icon size={17}/>{label}</button>)}</nav><main>{notice&&<div className="notice">{notice}<button onClick={()=>setNotice('')}>닫기</button></div>}
{tab==='entry'&&<section className="panel editor"><h2>{editing?'작업내역 수정':'작업 수량 등록'}</h2><form onSubmit={save}><label className="camera"><Camera size={24}/> 주소 사진 촬영 · 업로드<input type="file" accept="image/*" capture="environment" onChange={e=>photo(e.target.files?.[0])}/></label><label>주소지 (2글자 이상)<input required minLength={2} value={form.address} onChange={e=>setForm({...form,address:e.target.value})} placeholder="사진 인식 결과를 확인하거나 직접 입력"/></label><div className="quantities">{[['duvet','이불커버'],['small','수건 소'],['large','수건 대']].map(([k,label])=><div className="quantity" key={k}><span>{label}</span><div><button type="button" onClick={()=>setForm({...form,[k]:Math.max(0,form[k]-1)})}><Minus size={18}/></button><input className="quantity-input" type="number" min="0" step="1" inputMode="numeric" aria-label={`${label} 수량 직접 입력`} value={form[k]} onChange={e=>setForm(v=>({...v,[k]:e.target.value===""?0:Math.max(0,Math.floor(Number(e.target.value)||0))}))}/><button type="button" onClick={()=>setForm({...form,[k]:form[k]+1})}><Plus size={18}/></button></div></div>)}</div>{editing&&<><label>수정 사유<select value={reason} onChange={e=>setReason(e.target.value)}>{['주소지 수정','이불커버 수량 수정','수건 소 수량 수정','수건 대 수량 수정','전체 수정'].map(x=><option key={x}>{x}</option>)}</select></label><label>상세 수정 사유<textarea value={note} onChange={e=>setNote(e.target.value)} placeholder="수정 이유를 입력하세요"/></label></>}<button className="primary" disabled={busy}>{busy?'처리 중...':editing?'수정 즉시 반영':'작업 수량 등록'}</button>{editing&&<button type="button" onClick={()=>{setEditing(null);setForm(empty())}}>수정 취소</button>}</form></section>}
{['dashboard','entries','stats','mine'].includes(tab)&&<><div className="toolbar"><div className="date-range"><div className="preset-buttons">{[['today','오늘'],['twoWeeks','최근 2주'],['month','최근 한 달']].map(([id,label])=><button key={id} type="button" className={preset===id?'selected-preset':''} onClick={()=>choosePreset(id)}>{label}</button>)}</div><label>시작일 <input type="date" value={startDate} max={endDate} onChange={e=>{setStartDate(e.target.value);setPreset('custom')}}/></label><label>종료일 <input type="date" value={endDate} min={startDate} onChange={e=>{setEndDate(e.target.value);setPreset('custom')}}/></label></div>{owner&&<><label>직원<select value={person} onChange={e=>setPerson(e.target.value)}><option value="all">전체 직원</option>{staff.filter(p=>p.approved).map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select></label><label>시간대<select value={shift} onChange={e=>setShift(e.target.value)}><option>전체</option><option>오전</option><option>오후</option></select></label><button onClick={exportXlsx}><Download size={16}/> 엑셀</button></>}<label>주소 검색<input type="search" value={addressSearch} onChange={e=>setAddressSearch(e.target.value)} placeholder="주소 일부 입력"/></label><button onClick={load}><RefreshCw size={16}/> 새로고침</button></div><div className="cards">{[['활동 직원',people(filtered)+'명'],['이불커버',total.duvet+'개'],['수건 소',total.small+'개'],['수건 대',total.large+'개'],['작업 건수',filtered.length+'건']].map(([k,v])=><div className="metric" key={k}><span>{k}</span><strong>{v}</strong></div>)}</div></>}
{tab==='dashboard'&&<section className="panel"><h2>선택 기간 작업 내역</h2>{entryTable(filtered)}</section>}
{tab==='entries'&&<>{billingPanel}<section className="panel"><h2>전체 작업 내역 <small>체크박스로 정산할 작업을 선택하세요.</small></h2>{entryTable(filtered,false,true)}</section></>}
{tab==='mine'&&<section className="panel"><h2>내 작업 내역 <small>수정 버튼으로 즉시 변경할 수 있습니다.</small></h2>{entryTable(filtered,true)}</section>}
{tab==='stats'&&<><div className="cards"><div className="metric"><span>오전 활동 직원</span><strong>{people(am)}명</strong></div><div className="metric"><span>오후 활동 직원</span><strong>{people(pm)}명</strong></div><div className="metric"><span>오전 작업 건수</span><strong>{am.length}건</strong></div><div className="metric"><span>오후 작업 건수</span><strong>{pm.length}건</strong></div></div><div className="columns">{[['오전 00:00–11:59',am],['오후 12:00–23:59',pm]].map(([title,list])=><section className="panel" key={title}><h2>{title}</h2><div className="mini"><span>이불커버 <b>{sum(list).duvet}</b></span><span>수건 소 <b>{sum(list).small}</b></span><span>수건 대 <b>{sum(list).large}</b></span></div>{entryTable(list)}</section>)}</div><section className="panel"><h2>직원별 활동 현황</h2><div className="scroll"><table><thead><tr><th>직원</th><th>오전</th><th>오후</th><th>총 작업 건수</th></tr></thead><tbody>{[...new Set(filtered.map(x=>x.employee_id))].map(id=>{const mine=filtered.filter(x=>x.employee_id===id);return <tr key={id}><td>{names[id]||'직원'}</td><td>{mine.some(x=>period(x)==='오전')?'활동':'-'}</td><td>{mine.some(x=>period(x)==='오후')?'활동':'-'}</td><td>{mine.length}</td></tr>})}</tbody></table></div></section></>}
{tab==='history'&&<section className="panel"><h2>수정 이력 <small>수정 전후 수량은 삭제되지 않습니다.</small></h2><div className="history">{shownChanges.map(c=><div className="change" key={c.id}><div className="changehead"><b>{c.reason}</b><span>{time(c)} · {names[c.changed_by]||'직원'}</span></div><div className="diff"><div><small>수정 전</small><p>{c.before_values.address}</p><p>이불 {c.before_values.duvet} · 소 {c.before_values.small} · 대 {c.before_values.large}</p></div><div><small>수정 후</small><p>{c.after_values.address}</p><p>이불 {c.after_values.duvet} · 소 {c.after_values.small} · 대 {c.after_values.large}</p></div></div>{c.note&&<p className="muted">사유: {c.note}</p>}</div>)}{!shownChanges.length&&<p className="muted">수정 이력이 없습니다.</p>}</div></section>}
{tab==='staff'&&<section className="panel"><h2>직원 관리</h2><p className="muted">신규 직원은 회원가입 후 사업주 승인 시 작업을 시작할 수 있습니다.</p>{staff.map(p=><div className="staff" key={p.id}><div><b>{p.name}</b><small>{p.role==='owner'?'사업주':p.approved?'활동 가능':'승인 대기 / 정지'}</small></div>{p.role!=='owner'&&<button onClick={()=>approve(p)}>{p.approved?'계정 정지':'승인 / 활성화'}</button>}</div>)}</section>}
</main><footer>© 무무베딩 · 현장 수량관리</footer></div>}
createRoot(document.getElementById('root')).render(<App/>);
