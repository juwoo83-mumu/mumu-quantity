
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
const sum=rows=>rows.reduce((a,r)=>({duvet:a.duvet+r.duvet,small:a.small+r.small,large:a.large+r.large}),{duvet:0,small:0,large:0});
const people=rows=>new Set(rows.map(x=>x.employee_id)).size;
function App(){
const[session,setSession]=useState(null),[ready,setReady]=useState(false),[profile,setProfile]=useState(null),[tab,setTab]=useState('entry');
const[rows,setRows]=useState([]),[changes,setChanges]=useState([]),[staff,setStaff]=useState([]);
const[date,setDate]=useState(dateKST()),[person,setPerson]=useState('all'),[shift,setShift]=useState('전체');
const[form,setForm]=useState(empty()),[editing,setEditing]=useState(null),[reason,setReason]=useState('주소지 수정'),[note,setNote]=useState('');
const[busy,setBusy]=useState(false),[notice,setNotice]=useState(''),[authMode,setAuthMode]=useState('login'),[email,setEmail]=useState(''),[pass,setPass]=useState(''),[name,setName]=useState('');
useEffect(()=>{if(!db){setReady(true);return}db.auth.getSession().then(({data})=>{setSession(data.session);setReady(true)});const{data:{subscription}}=db.auth.onAuthStateChange((_event,s)=>setSession(s));return()=>subscription.unsubscribe()},[]);
useEffect(()=>{if(!session){setProfile(null);return}db.from('profiles').select('*').eq('id',session.user.id).single().then(({data,error})=>{setProfile(data);if(error)setNotice(error.message)})},[session]);
const owner=profile?.role==='owner';
async function load(){if(!profile?.approved)return;setBusy(true);
const from=date+'T00:00:00+09:00',to=new Date(new Date(from).getTime()+86400000).toISOString();
const[a,b,c]=await Promise.all([
 db.from('entries').select('*').gte('created_at',from).lt('created_at',to).order('created_at',{ascending:false}),
 db.from('changes').select('*').order('created_at',{ascending:false}).limit(1000),
 owner?db.from('profiles').select('*').order('name'):Promise.resolve({data:[]})
]);
if(a.error)setNotice(a.error.message);else setRows(a.data||[]);
if(b.error)setNotice(b.error.message);else setChanges(b.data||[]);
if(c.error)setNotice(c.error.message);else setStaff(c.data||[]);
setBusy(false)}
useEffect(()=>{if(profile?.approved)load()},[profile,date]);
const names=Object.fromEntries(staff.map(x=>[x.id,x.name]));if(profile)names[profile.id]=profile.name;
const filtered=rows.filter(r=>(person==='all'||r.employee_id===person)&&(shift==='전체'||period(r)===shift));
const total=sum(filtered),am=filtered.filter(r=>period(r)==='오전'),pm=filtered.filter(r=>period(r)==='오후');
async function login(e){e.preventDefault();setNotice('');setBusy(true);const{error}=authMode==='signup'?await db.auth.signUp({email,password:pass,options:{data:{name}}}):await db.auth.signInWithPassword({email,password:pass});setBusy(false);setNotice(error?error.message:authMode==='signup'?'가입 신청이 접수되었습니다. 이메일 인증 후 사업주 승인을 기다려 주세요.':'')}
async function logout(){await db.auth.signOut();setTab('entry')}
async function photo(file){if(!file)return;setBusy(true);setNotice('사진에서 한글 주소를 인식하고 있습니다.');try{const w=await createWorker('kor+eng');const{data}=await w.recognize(file);await w.terminate();const txt=data.text.split('\n').map(x=>x.trim()).filter(Boolean).join(' ');setForm(v=>({...v,address:txt}));setNotice(txt?'주소를 확인한 후 등록해 주세요.':'주소를 찾지 못했습니다. 직접 입력해 주세요.')}catch(err){setNotice('사진 인식에 실패했습니다. 주소를 직접 입력해 주세요.')}setBusy(false)}
async function save(e){e.preventDefault();if(!form.address.trim()){setNotice('주소지를 입력해 주세요.');return}setBusy(true);
const clean={address:form.address.trim(),duvet:Number(form.duvet),small:Number(form.small),large:Number(form.large)};
let error;if(editing){({error}=await db.rpc('revise_entry',{p_id:editing.id,p_address:clean.address,p_duvet:clean.duvet,p_small:clean.small,p_large:clean.large,p_reason:reason,p_note:note}))}
else{({error}=await db.from('entries').insert({...clean,employee_id:session.user.id}))}
setBusy(false);if(error){setNotice(error.message);return}setNotice(editing?'수정 내용이 반영되고 이전 수량이 이력에 보존되었습니다.':'수량이 등록되었습니다.');setEditing(null);setForm(empty());setNote('');await load()}
function startEdit(r){setEditing(r);setForm({address:r.address,duvet:r.duvet,small:r.small,large:r.large});setReason('주소지 수정');setNote('');setTab('entry');window.scrollTo({top:0,behavior:'smooth'})}
async function approve(p){const{error}=await db.from('profiles').update({approved:!p.approved}).eq('id',p.id);if(error)setNotice(error.message);else load()}
function exportXlsx(){const records=filtered.map(r=>({등록시간:time(r),시간대:period(r),직원:names[r.employee_id]||r.employee_id,주소:r.address,이불커버:r.duvet,'수건 소':r.small,'수건 대':r.large}));const sheet=XLSX.utils.json_to_sheet(records),book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,sheet,'작업내역');XLSX.writeFile(book,`무무베딩_${date}.xlsx`)}
const entryTable=(list,allowEdit=false)=><div className="scroll"><table><thead><tr><th>등록시간</th><th>직원</th><th>주소지</th><th>이불커버</th><th>수건 소</th><th>수건 대</th>{allowEdit&&<th>수정</th>}</tr></thead><tbody>{list.map(r=><tr key={r.id}><td>{time(r)}</td><td>{names[r.employee_id]||'직원'}</td><td>{r.address}</td><td>{r.duvet}</td><td>{r.small}</td><td>{r.large}</td>{allowEdit&&<td><button onClick={()=>startEdit(r)}>수정</button></td>}</tr>)}</tbody></table>{!list.length&&<p className="muted">등록된 내역이 없습니다.</p>}</div>;
if(!db)return <div className="center"><div className="login"><h1>무무베딩 수량관리</h1><p>클라우드 연결 전입니다. 프로젝트의 README 안내에 따라 Supabase 환경변수를 설정해 주세요.</p></div></div>;
if(!ready)return <div className="center">불러오는 중...</div>;
if(!session)return <div className="center"><form className="login" onSubmit={login}><div className="brand">MUMU BEDDING</div><h1>{authMode==='login'?'수량관리 로그인':'직원 회원가입'}</h1><p>현장 수량 기록과 실시간 운영 현황</p>{authMode==='signup'&&<label>직원 이름<input required value={name} onChange={e=>setName(e.target.value)} placeholder="이름"/></label>}<label>이메일<input type="email" required value={email} onChange={e=>setEmail(e.target.value)} placeholder="name@example.com"/></label><label>비밀번호<input type="password" minLength="6" required value={pass} onChange={e=>setPass(e.target.value)}/></label><button className="primary" disabled={busy}>{authMode==='login'?'로그인':'가입 신청'}</button><button type="button" className="plain" onClick={()=>{setAuthMode(authMode==='login'?'signup':'login');setNotice('')}}>{authMode==='login'?'신규 직원 회원가입':'로그인으로 돌아가기'}</button>{notice&&<div className="notice">{notice}</div>}</form></div>;
if(!profile)return <div className="center">계정 정보를 불러오는 중...</div>;
if(!profile.approved)return <div className="center"><div className="login"><div className="brand">MUMU BEDDING</div><h2>사업주 승인 대기 중</h2><p>{profile.name}님의 계정이 생성되었습니다. 사업주 승인 후 수량을 등록할 수 있습니다.</p><button onClick={()=>location.reload()}>승인 상태 새로고침</button><button className="plain" onClick={logout}>로그아웃</button></div></div>;
const tabs=owner?[['dashboard','대시보드',LayoutDashboard],['entries','작업내역',ClipboardList],['stats','오전·오후 통계',LayoutDashboard],['history','수정이력',History],['staff','직원관리',Users]]:[['entry','수량 등록',Plus],['mine','내 작업내역',ClipboardList],['history','수정이력',History]];
const shownChanges=changes.filter(x=>owner||x.changed_by===session.user.id);
return <div className="shell"><header><div><div className="brand">MUMU BEDDING</div><h1>수량관리 시스템</h1></div><div className="user"><span>{profile.name} {owner?'· 사업주':''}</span><button title="로그아웃" onClick={logout}><LogOut size={17}/></button></div></header><nav>{tabs.map(([id,label,Icon])=><button key={id} className={tab===id?'selected':''} onClick={()=>setTab(id)}><Icon size={17}/>{label}</button>)}</nav><main>{notice&&<div className="notice">{notice}<button onClick={()=>setNotice('')}>닫기</button></div>}
{tab==='entry'&&<section className="panel editor"><h2>{editing?'작업내역 수정':'작업 수량 등록'}</h2><form onSubmit={save}><label className="camera"><Camera size={24}/> 주소 사진 촬영 · 업로드<input type="file" accept="image/*" capture="environment" onChange={e=>photo(e.target.files?.[0])}/></label><label>주소지<input required value={form.address} onChange={e=>setForm({...form,address:e.target.value})} placeholder="사진 인식 결과를 확인하거나 직접 입력"/></label><div className="quantities">{[['duvet','이불커버'],['small','수건 소'],['large','수건 대']].map(([k,label])=><div className="quantity" key={k}><span>{label}</span><div><button type="button" onClick={()=>setForm({...form,[k]:Math.max(0,form[k]-1)})}><Minus size={18}/></button><strong>{form[k]}</strong><button type="button" onClick={()=>setForm({...form,[k]:form[k]+1})}><Plus size={18}/></button></div></div>)}</div>{editing&&<><label>수정 사유<select value={reason} onChange={e=>setReason(e.target.value)}>{['주소지 수정','이불커버 수량 수정','수건 소 수량 수정','수건 대 수량 수정','전체 수정'].map(x=><option key={x}>{x}</option>)}</select></label><label>상세 수정 사유<textarea value={note} onChange={e=>setNote(e.target.value)} placeholder="수정 이유를 입력하세요"/></label></>}<button className="primary" disabled={busy}>{busy?'처리 중...':editing?'수정 즉시 반영':'작업 수량 등록'}</button>{editing&&<button type="button" onClick={()=>{setEditing(null);setForm(empty())}}>수정 취소</button>}</form></section>}
{['dashboard','entries','stats','mine'].includes(tab)&&<><div className="toolbar"><label>날짜 <input type="date" value={date} onChange={e=>setDate(e.target.value)}/></label>{owner&&<><label>직원<select value={person} onChange={e=>setPerson(e.target.value)}><option value="all">전체 직원</option>{staff.filter(p=>p.approved).map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select></label><label>시간대<select value={shift} onChange={e=>setShift(e.target.value)}><option>전체</option><option>오전</option><option>오후</option></select></label><button onClick={exportXlsx}><Download size={16}/> 엑셀</button></>}<button onClick={load}><RefreshCw size={16}/> 새로고침</button></div><div className="cards">{[['활동 직원',people(filtered)+'명'],['이불커버',total.duvet+'개'],['수건 소',total.small+'개'],['수건 대',total.large+'개'],['작업 건수',filtered.length+'건']].map(([k,v])=><div className="metric" key={k}><span>{k}</span><strong>{v}</strong></div>)}</div></>}
{tab==='dashboard'&&<section className="panel"><h2>오늘의 작업 내역</h2>{entryTable(filtered)}</section>}
{tab==='entries'&&<section className="panel"><h2>전체 작업 내역</h2>{entryTable(filtered)}</section>}
{tab==='mine'&&<section className="panel"><h2>내 작업 내역 <small>수정 버튼으로 즉시 변경할 수 있습니다.</small></h2>{entryTable(filtered,true)}</section>}
{tab==='stats'&&<><div className="cards"><div className="metric"><span>오전 활동 직원</span><strong>{people(am)}명</strong></div><div className="metric"><span>오후 활동 직원</span><strong>{people(pm)}명</strong></div><div className="metric"><span>오전 작업 건수</span><strong>{am.length}건</strong></div><div className="metric"><span>오후 작업 건수</span><strong>{pm.length}건</strong></div></div><div className="columns">{[['오전 00:00–11:59',am],['오후 12:00–23:59',pm]].map(([title,list])=><section className="panel" key={title}><h2>{title}</h2><div className="mini"><span>이불커버 <b>{sum(list).duvet}</b></span><span>수건 소 <b>{sum(list).small}</b></span><span>수건 대 <b>{sum(list).large}</b></span></div>{entryTable(list)}</section>)}</div><section className="panel"><h2>직원별 활동 현황</h2><div className="scroll"><table><thead><tr><th>직원</th><th>오전</th><th>오후</th><th>총 작업 건수</th></tr></thead><tbody>{[...new Set(filtered.map(x=>x.employee_id))].map(id=>{const mine=filtered.filter(x=>x.employee_id===id);return <tr key={id}><td>{names[id]||'직원'}</td><td>{mine.some(x=>period(x)==='오전')?'활동':'-'}</td><td>{mine.some(x=>period(x)==='오후')?'활동':'-'}</td><td>{mine.length}</td></tr>})}</tbody></table></div></section></>}
{tab==='history'&&<section className="panel"><h2>수정 이력 <small>수정 전후 수량은 삭제되지 않습니다.</small></h2><div className="history">{shownChanges.map(c=><div className="change" key={c.id}><div className="changehead"><b>{c.reason}</b><span>{time(c)} · {names[c.changed_by]||'직원'}</span></div><div className="diff"><div><small>수정 전</small><p>{c.before_values.address}</p><p>이불 {c.before_values.duvet} · 소 {c.before_values.small} · 대 {c.before_values.large}</p></div><div><small>수정 후</small><p>{c.after_values.address}</p><p>이불 {c.after_values.duvet} · 소 {c.after_values.small} · 대 {c.after_values.large}</p></div></div>{c.note&&<p className="muted">사유: {c.note}</p>}</div>)}{!shownChanges.length&&<p className="muted">수정 이력이 없습니다.</p>}</div></section>}
{tab==='staff'&&<section className="panel"><h2>직원 관리</h2><p className="muted">신규 직원은 회원가입 후 사업주 승인 시 작업을 시작할 수 있습니다.</p>{staff.map(p=><div className="staff" key={p.id}><div><b>{p.name}</b><small>{p.role==='owner'?'사업주':p.approved?'활동 가능':'승인 대기 / 정지'}</small></div>{p.role!=='owner'&&<button onClick={()=>approve(p)}>{p.approved?'계정 정지':'승인 / 활성화'}</button>}</div>)}</section>}
</main><footer>© 무무베딩 · 현장 수량관리</footer></div>}
createRoot(document.getElementById('root')).render(<App/>);
