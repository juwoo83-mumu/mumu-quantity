
import React,{useEffect,useMemo,useState,useRef} from 'react';
import{createRoot}from'react-dom/client';
import{createClient}from'@supabase/supabase-js';
import{Camera,LayoutDashboard,ClipboardList,History,Users,LogOut,Plus,Minus,Download,RefreshCw,Search,MapPin,CalendarDays,Clock3}from'lucide-react';
import*as XLSX from'xlsx';
import './style.css';
const url=import.meta.env.VITE_SUPABASE_URL,key=import.meta.env.VITE_SUPABASE_ANON_KEY;
const db=url&&key?createClient(url,key):null;
const empty=()=>({address:'',duvet:0,small:0,large:0});
// 주소 사진을 꼭 찍을 필요 없이, 2글자 이상 입력하면 기존 주소를 기기에서 즉시 검색합니다.
// 공백·특수기호가 달라도 실제 글자와 숫자 순서가 같으면 동일한 주소로 판단합니다.
// 예: '서울 강남구 123-4', '서울강남구123_4' -> 같은 주소
const normalizedAddress=s=>String(s||'').normalize('NFC').toLowerCase().replace(/[^가-힣a-z0-9]/g,'');
const spaceFreeAddress=normalizedAddress;
// 공백으로 구분된 단어를 특수기호 없이 비교합니다. 순서는 무시하지만 중복 단어의 개수까지 유지합니다.
// 예: '서울 강남구 역삼동 123-4' === '123_4 역삼동 강남구 서울'
const addressWords=address=>String(address||'').normalize('NFC').toLowerCase().trim()
 .split(/\s+/).map(word=>word.replace(/[^가-힣a-z0-9]/g,'')).filter(Boolean);
const addressWordsKey=address=>addressWords(address).sort().join('\u001f');
const sameAddress=(a,b)=>{
 const compactA=normalizedAddress(a),compactB=normalizedAddress(b);
 return compactA.length>=2&&compactB.length>=2&&
  (compactA===compactB||addressWordsKey(a)===addressWordsKey(b));
};
// 일부 단어만 입력해도 입력 순서와 관계없이 후보를 찾습니다. 이 경우 중복 판단은 하지 않습니다.
function hasWordsInAnyOrder(query,known){
 const words=addressWords(query),target=addressWords(known);
 if(words.length<2||words.length>target.length)return false;
 const remaining=[...target];
 for(const word of words){
  const idx=remaining.findIndex(item=>item.includes(word));
  if(idx<0)return false;
  remaining.splice(idx,1);
 }
 return true;
}
// 주소 앞이나 뒤에 관리용 표기 '수정'이 붙어 있는 과거 주소는 후보에서 제외합니다.
// 앞에 띄어쓰기/특수기호가 있거나 '수정서울...'처럼 붙어 있어도 제외합니다.
// 단, '수정구', '수정동' 등 지명으로 시작하는 주소와 '성남시 수정구'는 유지합니다.
const hasEditMarker=address=>{
 const value=String(address||'').normalize('NFC').trim();
 const beginning=value.replace(/^[^가-힣a-z0-9]*/i,'');
 const prefix=/^수정(?:(?!구|동|로|길|리|면|읍|군|시)[가-힣a-z0-9]|[^가-힣a-z0-9])/i.test(beginning);
 const suffix=/수정[^가-힣a-z0-9]*$/i.test(value);
 return prefix||suffix;
};
function uniqueAddresses(addresses){
 const seenCompact=new Set(),seenWords=new Set();
 return addresses.filter(address=>{
  const compact=spaceFreeAddress(address),words=addressWordsKey(address);
  if(compact.length<2||hasEditMarker(address))return false;
  // 공백·기호만 달라도 같고, 단어의 종류·개수가 같은 채 순서만 달라도 같습니다.
  if(seenCompact.has(compact)||seenWords.has(words))return false;
  seenCompact.add(compact);
  seenWords.add(words);
  return true;
 });
}
const addressBigrams=s=>{const out=new Set();for(let i=0;i<s.length-1;i++)out.add(s.slice(i,i+2));return out};
function photoAddressScore(ocrLine,known){
 const a=normalizedAddress(ocrLine),b=normalizedAddress(known);
 if(a.length<3||b.length<3)return 0;
 if(sameAddress(ocrLine,known))return 1000;
 if(b.includes(a))return 850+Math.min(a.length,80);
 if(a.includes(b))return 830+Math.min(b.length,80);
 const A=addressBigrams(a),B=addressBigrams(b);let common=0;
 for(const n of A)if(B.has(n))common++;
 const dice=2*common/Math.max(1,A.size+B.size);
 return dice>=.4?Math.round(dice*700):0;
}
function findPhotoMatches(raw,addresses){
 const lines=String(raw||'').split(/\r?\n/).map(x=>x.trim()).filter(x=>normalizedAddress(x).length>=3).slice(0,18);
 if(!lines.length)return [];
 return addresses.map(address=>({address,score:Math.max(...lines.map(line=>photoAddressScore(line,address)))}))
 .filter(x=>x.score>=280).sort((a,b)=>b.score-a.score||a.address.localeCompare(b.address,'ko')).slice(0,8);
}
// OCR 라이브러리와 한글 데이터는 사진 버튼을 눌렀을 때만 내려받고, 같은 화면에서 재사용합니다.
let ocrWorkerPromise;
function getOcrWorker(){
 if(!ocrWorkerPromise)ocrWorkerPromise=import('tesseract.js').then(({createWorker})=>createWorker('kor+eng')).catch(error=>{ocrWorkerPromise=null;throw error});
 return ocrWorkerPromise;
}
const time=e=>new Date(e.created_at).toLocaleString('ko-KR',{timeZone:'Asia/Seoul',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit'});
const period=e=>new Date(e.created_at).toLocaleString('en-US',{timeZone:'Asia/Seoul',hour:'2-digit',hourCycle:'h23'}).split(':')[0]<12?'오전':'오후';
const dateKST=()=>new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Seoul',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
const shiftKST=(day,offset)=>{const d=new Date(day+'T12:00:00+09:00');d.setUTCDate(d.getUTCDate()+offset);return new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Seoul',year:'numeric',month:'2-digit',day:'2-digit'}).format(d)};
const monthAgoKST=day=>{const [y,m,d]=day.split('-').map(Number);const first=new Date(Date.UTC(y,m-2,1));const maxDay=new Date(Date.UTC(first.getUTCFullYear(),first.getUTCMonth()+1,0)).getUTCDate();first.setUTCDate(Math.min(d,maxDay));return first.toISOString().slice(0,10)};
const sum=rows=>rows.reduce((a,r)=>({duvet:a.duvet+r.duvet,small:a.small+r.small,large:a.large+r.large}),{duvet:0,small:0,large:0});
const coverPrice=n=>n<=0?0:n===1?12000:15000+Math.floor((n-2)/2)*7000+((n-2)%2)*3000;
const won=n=>n.toLocaleString('ko-KR')+'원';
const people=rows=>new Set(rows.map(x=>x.employee_id)).size;
const WORKPLACE_ADDRESS='경기도 고양시 일산동구 중앙로1261번길 77 (장항동)';
const INITIAL_WORKPLACE={latitude:37.65713466016,longitude:126.76958232013,radius_m:250,address:WORKPLACE_ADDRESS};
const asKoreanDateTime=value=>value?new Date(value).toLocaleString('ko-KR',{timeZone:'Asia/Seoul',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit'}):'—';
const metersBetween=(a,b,c,d)=>{const r=Math.PI/180;const x=Math.sin((c-a)*r/2)**2+Math.cos(a*r)*Math.cos(c*r)*Math.sin((d-b)*r/2)**2;return 6371000*2*Math.atan2(Math.sqrt(x),Math.sqrt(Math.max(0,1-x)))};
const workDateKST=r=>new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Seoul',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date(r.created_at));
function App(){
const[session,setSession]=useState(null),[ready,setReady]=useState(false),[profile,setProfile]=useState(null),[tab,setTab]=useState('entry');
const[rows,setRows]=useState([]),[changes,setChanges]=useState([]),[staff,setStaff]=useState([]),[staffEmails,setStaffEmails]=useState({}),[staffEmailError,setStaffEmailError]=useState('');
const[startDate,setStartDate]=useState(dateKST()),[endDate,setEndDate]=useState(dateKST()),[preset,setPreset]=useState('today'),[person,setPerson]=useState('all'),[shift,setShift]=useState('전체'),[addressSearch,setAddressSearch]=useState('');
const[addressMethod,setAddressMethod]=useState('search'),[form,setForm]=useState(empty()),[editing,setEditing]=useState(null),[reason,setReason]=useState('주소지 수정'),[note,setNote]=useState('');
const[knownAddresses,setKnownAddresses]=useState([]),[addressLookupState,setAddressLookupState]=useState('idle'),[chosenExistingAddress,setChosenExistingAddress]=useState(''),[photoMatches,setPhotoMatches]=useState([]),[photoBusy,setPhotoBusy]=useState(false);
const[selectedIds,setSelectedIds]=useState([]),[copyNotice,setCopyNotice]=useState(''),[existingBalanceInput,setExistingBalanceInput]=useState('');
const[ownerEditing,setOwnerEditing]=useState(null),[ownerEditForm,setOwnerEditForm]=useState({address:'',duvet:0,small:0,large:0}),[ownerEditNote,setOwnerEditNote]=useState(''),[ownerEditError,setOwnerEditError]=useState(''),[ownerEditSaving,setOwnerEditSaving]=useState(false);
const[pendingEntry,setPendingEntry]=useState(null),[confirmError,setConfirmError]=useState('');
const confirmDialogRef=useRef(null);
const registerInFlightRef=useRef(false);
const[busy,setBusy]=useState(false),[notice,setNotice]=useState(''),[authMode,setAuthMode]=useState('login'),[email,setEmail]=useState(''),[pass,setPass]=useState(''),[confirmPass,setConfirmPass]=useState(''),[name,setName]=useState('');
const[recoveryRoute,setRecoveryRoute]=useState(()=>window.location.pathname==='/reset-password');
const[recoveryPass,setRecoveryPass]=useState(''),[recoveryConfirm,setRecoveryConfirm]=useState(''),[recoveryError,setRecoveryError]=useState('');
const recoveryUrlError=new URLSearchParams(window.location.hash.replace(/^#/, '')).get('error_description')||new URLSearchParams(window.location.search).get('error_description');
// HR data stays separate from existing laundry quantity records.
const[workplace,setWorkplace]=useState(INITIAL_WORKPLACE),[attendance,setAttendance]=useState([]),[leaves,setLeaves]=useState([]),[hrError,setHrError]=useState('');
// 한국 날짜가 자정에 바뀌면 열린 전날 출근을 현재 출근 상태로 보지 않습니다.
const[todayKst,setTodayKst]=useState(dateKST());
const[hrBusy,setHrBusy]=useState(false),[geoStatus,setGeoStatus]=useState(''),[lastDistance,setLastDistance]=useState(null);
const[leaveStart,setLeaveStart]=useState(dateKST()),[leaveEnd,setLeaveEnd]=useState(dateKST()),[leaveType,setLeaveType]=useState('연차'),[leaveNote,setLeaveNote]=useState('');
const[radiusInput,setRadiusInput]=useState('250');
const geoWatchRef=useRef(null),outsideRef=useRef({at:0,count:0}),autoOutRef=useRef(false);
const hrReadingRef=useRef(false);

useEffect(()=>{
 if(!db){setReady(true);return}
 let active=true;
 // Subscribe before reading the initial session: recovery redirects may emit PASSWORD_RECOVERY.
 const{data:{subscription}}=db.auth.onAuthStateChange((event,s)=>{
  if(!active)return;
  setSession(s);
  if(event==='PASSWORD_RECOVERY')setRecoveryRoute(true);
 });
 db.auth.getSession().then(({data,error})=>{
  if(!active)return;
  if(!error)setSession(data.session);
  setReady(true);
 }).catch(()=>{if(active)setReady(true)});
 return()=>{active=false;subscription.unsubscribe()};
},[]);
useEffect(()=>{if(!session){setProfile(null);return}db.from('profiles').select('*').eq('id',session.user.id).single().then(({data,error})=>{setProfile(data);if(error)setNotice(error.message)})},[session]);
const owner=profile?.role==='owner';
useEffect(()=>{if(profile?.role==='owner'&&tab==='entry')setTab('dashboard')},[profile?.role]);
const attendanceWorkDay=a=>a.work_date||workDateKST({created_at:a.check_in_at});
const currentAttendance=attendance.find(a=>a.employee_id===session?.user?.id&&!a.check_out_at&&attendanceWorkDay(a)===todayKst);
const attendanceToday=attendance.filter(a=>attendanceWorkDay(a)===todayKst);
const missedAttendance=attendance.filter(a=>a.employee_id===session?.user?.id&&!a.check_out_at&&attendanceWorkDay(a)<todayKst);
const attendanceStateLabel=a=>a.check_out_method==='geofence'?'지역 이탈 자동':a.check_out_method==='manual'?'수동':a.check_out_at?'기타':attendanceWorkDay(a)<todayKst?'미퇴근':'근무 중';
// 모바일 브라우저가 백그라운드에서 멈췄다가 재개되어도 날짜를 재확인합니다.
useEffect(()=>{
 const updateDay=()=>setTodayKst(dateKST());
 const timer=window.setInterval(updateDay,30000);
 window.addEventListener('focus',updateDay);
 document.addEventListener('visibilitychange',updateDay);
 return()=>{window.clearInterval(timer);window.removeEventListener('focus',updateDay);document.removeEventListener('visibilitychange',updateDay)};
},[]);
const pendingLeaves=leaves.filter(l=>l.status==='pending');
const ownerAttendByDate=attendance.filter(a=>workDateKST({created_at:a.check_in_at})>=startDate&&workDateKST({created_at:a.check_in_at})<=endDate);
async function loadHr(){
 if(!db||!profile?.approved||hrReadingRef.current)return;
 hrReadingRef.current=true;
 try{
  const [office,turns,requests]=await Promise.all([
   db.from('workplace_settings').select('*').eq('id',1).maybeSingle(),
   db.from('attendance_records').select('*').order('check_in_at',{ascending:false}).limit(1000),
   db.from('leave_requests').select('*').order('created_at',{ascending:false}).limit(1000)
  ]);
  const err=office.error||turns.error||requests.error;
  if(err){setHrError('근태·휴가 설정을 불러오지 못했습니다. HR_ATTENDANCE_LEAVE_SETUP.sql을 Supabase SQL Editor에서 한 번 실행해 주세요. ('+err.message+')');return}
  if(office.data){setWorkplace(office.data);setRadiusInput(String(office.data.radius_m));}
  setAttendance(turns.data||[]);setLeaves(requests.data||[]);setHrError('');
 }catch(e){setHrError(e.message||'근태 정보를 불러오지 못했습니다.');}
 finally{hrReadingRef.current=false}
}
useEffect(()=>{if(profile?.approved)loadHr()},[profile?.id,profile?.approved,todayKst]);
useEffect(()=>{if(profile?.approved&&tab==='dashboard')loadHr()},[tab,startDate,endDate]);
function readLocation(){return new Promise((resolve,reject)=>{
 if(!navigator.geolocation){reject(new Error('이 기기에서는 위치 서비스를 지원하지 않습니다.'));return}
 navigator.geolocation.getCurrentPosition(p=>resolve(p.coords),e=>reject(new Error(e.code===1?'위치 사용 권한이 거부되었습니다. 브라우저에서 위치를 허용해 주세요.':e.code===2?'GPS 위치를 확인하지 못했습니다. 실외에서 재시도해 주세요.':'위치 확인 시간이 초과되었습니다. 다시 시도해 주세요.')),{enableHighAccuracy:true,timeout:18000,maximumAge:0});
})}
async function clockIn(){
 if(hrBusy||currentAttendance)return;setHrBusy(true);setHrError('');setGeoStatus('근무지 위치 확인 중…');
 try{
  const pos=await readLocation();
  if(pos.accuracy>120)throw new Error('위치 정확도가 낮습니다 ('+Math.round(pos.accuracy)+'m). 위치 서비스를 켜고 다시 시도해 주세요.');
  const dist=metersBetween(pos.latitude,pos.longitude,workplace.latitude,workplace.longitude);
  setLastDistance(Math.round(dist));
  if(dist>workplace.radius_m+Math.min(pos.accuracy,30))throw new Error('현재 근무지로부터 약 '+Math.round(dist)+'m 떨어져 있습니다. 출근은 지정된 근무지에서만 가능합니다.');
  const {error}=await db.rpc('clock_in_attendance',{p_lat:pos.latitude,p_lon:pos.longitude,p_accuracy:pos.accuracy});
  if(error)throw error;
  setGeoStatus('출근 완료. 이 화면이 열려 있는 동안 위치를 확인하여 자동 퇴근합니다.');
  outsideRef.current={at:0,count:0};await loadHr();
 }catch(e){setGeoStatus(e.message||'출근 처리에 실패했습니다.');}
 finally{setHrBusy(false)}
}
async function clockOut(method='manual',pos){
 if(autoOutRef.current)return;
 autoOutRef.current=true;setHrBusy(true);
 try{
  const args={p_method:method};
  if(method==='geofence'&&pos){args.p_lat=pos.latitude;args.p_lon=pos.longitude;args.p_accuracy=pos.accuracy}
  const {error}=await db.rpc('clock_out_attendance',args);
  if(error)throw error;
  setGeoStatus(method==='geofence'?'근무지역 이탈이 확인되어 자동 퇴근되었습니다.':'수동 퇴근 처리되었습니다.');
  if(geoWatchRef.current!==null){navigator.geolocation.clearWatch(geoWatchRef.current);geoWatchRef.current=null}
  outsideRef.current={at:0,count:0};await loadHr();
 }catch(e){setGeoStatus('퇴근 처리 오류: '+(e.message||'다시 시도해 주세요.'));}
 finally{autoOutRef.current=false;setHrBusy(false)}
}
useEffect(()=>{
 // Foreground web pages only: browsers may suspend or end monitoring in the background.
 if(!profile?.approved||owner||!currentAttendance||!navigator.geolocation)return;
 let disposed=false;
 const id=navigator.geolocation.watchPosition(async p=>{
  if(disposed||autoOutRef.current||dateKST()!==todayKst)return;
  const coords=p.coords;
  if(coords.accuracy>120){setGeoStatus('GPS 오차가 커서 자동 퇴근을 보류하고 있습니다. 수동 퇴근도 가능합니다.');outsideRef.current={at:0,count:0};return}
  const dist=metersBetween(coords.latitude,coords.longitude,workplace.latitude,workplace.longitude);
  setLastDistance(Math.round(dist));
  // Count only reliable observations beyond the radius and reported GPS accuracy.
  const outside=dist>workplace.radius_m+Math.max(30,coords.accuracy);
  if(!outside){outsideRef.current={at:0,count:0};setGeoStatus('근무지역 안입니다. 위치 확인 중');return}
  const now=Date.now(),previous=outsideRef.current;
  if(!previous.at){outsideRef.current={at:now,count:1};setGeoStatus('근무지역 이탈 의심: GPS를 한 번 더 확인합니다.');return}
  const count=previous.count+1;outsideRef.current={at:previous.at,count};
  if(count>=2&&now-previous.at>=10000)await clockOut('geofence',coords);
  else setGeoStatus('근무지역 이탈 확인 중…');
 },e=>{if(!disposed)setGeoStatus(e.code===1?'위치 권한이 없어 자동 퇴근을 할 수 없습니다. 수동 퇴근을 사용하세요.':'위치 신호를 확인할 수 없습니다. 수동 퇴근이 가능합니다.')},{enableHighAccuracy:true,maximumAge:0,timeout:20000});
 geoWatchRef.current=id;
 return()=>{disposed=true;navigator.geolocation.clearWatch(id);if(geoWatchRef.current===id)geoWatchRef.current=null;outsideRef.current={at:0,count:0}};
},[currentAttendance?.id,todayKst,profile?.approved,owner,workplace.latitude,workplace.longitude,workplace.radius_m]);
async function submitLeave(event){
 event.preventDefault();if(hrBusy)return;setHrBusy(true);setHrError('');
 try{
  const {error}=await db.rpc('submit_leave_request',{p_start_date:leaveStart,p_end_date:leaveEnd,p_leave_type:leaveType,p_note:leaveNote.trim()});
  if(error)throw error;setLeaveNote('');setGeoStatus('휴가 신청이 접수되었습니다. 사업주 승인을 기다려 주세요.');await loadHr();
 }catch(e){setHrError('휴가 신청 실패: '+e.message)}finally{setHrBusy(false)}
}
async function cancelLeave(id){if(hrBusy)return;setHrBusy(true);setHrError('');try{const{error}=await db.rpc('cancel_leave_request',{p_id:id});if(error)throw error;await loadHr()}catch(e){setHrError('휴가 신청 취소 실패: '+e.message)}finally{setHrBusy(false)}}
async function reviewLeave(id,status){if(!owner||hrBusy)return;setHrBusy(true);setHrError('');try{const{error}=await db.rpc('review_leave_request',{p_id:id,p_status:status});if(error)throw error;await loadHr()}catch(e){setHrError('휴가 승인/반려 실패: '+e.message)}finally{setHrBusy(false)}}
async function updateRadius(event){event.preventDefault();if(!owner)return;
 const radius=Number(radiusInput);if(!Number.isInteger(radius)||radius<100||radius>2000){setHrError('근무 반경은 100~2000m 정수로 지정해 주세요.');return}
 setHrBusy(true);try{
 const{error}=await db.from('workplace_settings').update({radius_m:radius,updated_by:session.user.id,updated_at:new Date().toISOString()}).eq('id',1);
 if(error)throw error;await loadHr();setGeoStatus('근무지역 반경을 변경했습니다.');
 }catch(e){setHrError('근무지 반경 저장 실패: '+e.message)}finally{setHrBusy(false)}
}

// Only show the review dialog after the form has passed validation. Nothing is saved until confirmation.
useEffect(()=>{
 const dialog=confirmDialogRef.current;
 if(!dialog)return;
 if(pendingEntry&&!dialog.open)dialog.showModal();
 if(!pendingEntry&&dialog.open)dialog.close();
},[pendingEntry]);
useEffect(()=>{if(tab!=='entry')setPendingEntry(null)},[tab]);

// 기존 업무자료 전체에서 '주소 텍스트만' 조회합니다. 직원은 다른 사람의 수량/이력은 볼 수 없습니다.
useEffect(()=>{
 if(!db||!profile?.approved||tab!=='entry')return;
 let stopped=false;
 setAddressLookupState('loading');
 db.rpc('list_existing_addresses',{p_limit:10000}).then(({data,error})=>{
  if(stopped)return;
  if(error){setAddressLookupState('error');return}
  setKnownAddresses(uniqueAddresses((data||[]).map(x=>x.address)));
  setAddressLookupState('ready');
 });
 return()=>{stopped=true};
},[profile?.id,profile?.approved,tab]);
const addressInputNormalized=normalizedAddress(form.address);
const matchingAddresses=useMemo(()=>{
 if(addressInputNormalized.length<2||chosenExistingAddress===form.address)return [];
 return knownAddresses.map(address=>{
  const known=normalizedAddress(address);
  let rank=0;
  if(sameAddress(address,form.address))rank=1000;
  else if(known.startsWith(addressInputNormalized))rank=900;
  else if(known.includes(addressInputNormalized))rank=800;
  else if(hasWordsInAnyOrder(form.address,address))rank=750;
  else if(addressInputNormalized.length>=5&&addressInputNormalized.includes(known))rank=700;
  return {address,rank};
 }).filter(x=>x.rank>0).sort((a,b)=>b.rank-a.rank||a.address.localeCompare(b.address,'ko')).slice(0,10);
},[addressInputNormalized,knownAddresses,chosenExistingAddress,form.address]);
function chooseAddress(address){
 setForm(previous=>({...previous,address}));
 setChosenExistingAddress(address);
 setPhotoMatches([]);
 setNotice('기존 작업내역의 주소를 선택했습니다. 수량을 입력해 주세요.');
}
function editAddress(value){
 setForm(previous=>({...previous,address:value}));
 setChosenExistingAddress('');
 setPhotoMatches([]);
}
async function load(){if(!profile?.approved)return;setBusy(true);
if(!startDate||!endDate||startDate>endDate){setNotice('조회 시작일은 종료일보다 늦을 수 없습니다.');setBusy(false);return}
const from=startDate+'T00:00:00+09:00',to=new Date(new Date(endDate+'T00:00:00+09:00').getTime()+86400000).toISOString();
const[a,b,c,d]=await Promise.all([
 fetchEntries(from,to),
 db.from('changes').select('*').order('created_at',{ascending:false}).limit(1000),
 owner?db.from('profiles').select('*').order('name'):Promise.resolve({data:[]}),
 owner?db.rpc('owner_list_staff_emails'):Promise.resolve({data:[]})
]);
if(a.error)setNotice(a.error.message);else setRows(a.data||[]);
if(b.error)setNotice(b.error.message);else setChanges(b.data||[]);
if(c.error)setNotice(c.error.message);else setStaff(c.data||[]);
if(d.error){
 setStaffEmails({});
 setStaffEmailError('이메일 목록을 불러오지 못했습니다. STAFF_EMAIL_SETUP.sql을 Supabase SQL Editor에서 한 번 실행해 주세요.');
}else{
 setStaffEmails(Object.fromEntries((d.data||[]).map(u=>[u.profile_id,u.email])));
 setStaffEmailError('');
}
setBusy(false)}
useEffect(()=>{setSelectedIds([]);setExistingBalanceInput('');setCopyNotice('');setOwnerEditing(null);if(profile?.approved)load()},[profile,startDate,endDate]);
useEffect(()=>{setOwnerEditing(null)},[person,shift,addressSearch]);
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
const filtered=rows.filter(r=>(person==='all'||r.employee_id===person)&&(shift==='전체'||period(r)===shift)&&(tab!=='entries'||r.address.toLocaleLowerCase().includes(addressSearch.trim().toLocaleLowerCase())));
const visibleIds=new Set(filtered.map(r=>String(r.id)));
const selectedRows=filtered.filter(r=>selectedIds.includes(String(r.id)));
const selectedTotal=sum(selectedRows),selectedCoverPrice=coverPrice(selectedTotal.duvet),selectedSmallPrice=selectedTotal.small*700,selectedLargePrice=selectedTotal.large*900,selectedAmount=selectedCoverPrice+selectedSmallPrice+selectedLargePrice;
// '-' is a temporary editing state, not a complete amount.
const hasExistingBalance=existingBalanceInput!==''&&existingBalanceInput!=='-';
const existingBalance=hasExistingBalance?Number(existingBalanceInput):0;
const formattedExistingBalance=existingBalanceInput===''?'':`${existingBalanceInput.startsWith('-')?'-':''}${existingBalanceInput==='-'?'':Number(existingBalanceInput.replace('-','')).toLocaleString('ko-KR')}`;
const remainingBalance=existingBalance-selectedAmount;
const selectedQuantityText=[selectedTotal.duvet>0?`이불 ${selectedTotal.duvet/2}세트`:null,selectedTotal.small>0?`수건 소 ${selectedTotal.small}`:null,selectedTotal.large>0?`${selectedTotal.small>0?"대":"수건 대"} ${selectedTotal.large}`:null].filter(Boolean).join(', ');
const showDepositNotice=selectedRows.length>0&&hasExistingBalance&&remainingBalance<20000;
const selectedCopyText=selectedRows.length&&hasExistingBalance?[selectedQuantityText,`잔액 ${won(remainingBalance)}`].filter(Boolean).join(', ')+(showDepositNotice?'\n입금 확인 부탁드려요':''):'';
const selectedSummary=`무무베딩 작업내역 정산\n선택 작업 ${selectedRows.length}건\n이불커버 ${selectedTotal.duvet}개 (${selectedTotal.duvet/2}세트): ${won(selectedCoverPrice)}\n수건 소 ${selectedTotal.small}개 × 700원: ${won(selectedSmallPrice)}\n수건 대 ${selectedTotal.large}개 × 900원: ${won(selectedLargePrice)}\n총 금액: ${won(selectedAmount)}`;
function toggleSelected(id){setSelectedIds(prev=>prev.includes(String(id))?prev.filter(x=>x!==String(id)):[...prev,String(id)]);setCopyNotice('')}
function toggleAll(list){const ids=list.map(r=>String(r.id));setSelectedIds(prev=>ids.every(id=>prev.includes(id))?prev.filter(id=>!ids.includes(id)):[...new Set([...prev,...ids])]);setCopyNotice('')}
function openOwnerEdit(row){
 if(!owner)return;
 setOwnerEditing({...row});
 setOwnerEditForm({address:row.address,duvet:row.duvet,small:row.small,large:row.large});
 setOwnerEditNote('');setOwnerEditError('');
 requestAnimationFrame(()=>document.getElementById('owner-entry-editor')?.scrollIntoView({behavior:'smooth',block:'start'}));
}
function editSelectedOwnerEntry(){
 if(selectedRows.length!==1){setOwnerEditError('수정할 작업내역을 한 건만 선택해 주세요.');return}
 openOwnerEdit(selectedRows[0]);
}
async function submitOwnerEdit(event){
 event.preventDefault();
 if(!owner||!ownerEditing)return;
 const newAddress=ownerEditForm.address.trim();
 if(normalizedAddress(newAddress).length<2){setOwnerEditError('주소는 공백과 특수기호를 제외하고 2글자 이상 입력해 주세요.');return}
 const inputKeys=['duvet','small','large'];
 const values={};
 for(const key of inputKeys){
  const value=ownerEditForm[key];
  if(value===''||!Number.isInteger(Number(value))||Number(value)<0||Number(value)>2147483647){
   setOwnerEditError('수량은 0 이상의 정수로 입력해 주세요.');return;
  }
  values[key]=Number(value);
 }
 const changed=inputKeys.filter(key=>values[key]!==ownerEditing[key]);
 const addressChanged=newAddress!==ownerEditing.address;
 if(!changed.length&&!addressChanged){setOwnerEditError('변경된 주소나 수량이 없습니다.');return}
 const labels={duvet:'이불커버 수량 수정',small:'수건 소 수량 수정',large:'수건 대 수량 수정'};
 const editReason=addressChanged&&changed.length===0?'주소지 수정':!addressChanged&&changed.length===1?labels[changed[0]]:'전체 수정';
 setOwnerEditSaving(true);setOwnerEditError('');
 try{
  // Other staff may have revised the same record after this screen was loaded.
  const {data:current,error:readError}=await db.from('entries').select('id,address,duvet,small,large,updated_at').eq('id',ownerEditing.id).single();
  if(readError)throw readError;
  if(current.updated_at!==ownerEditing.updated_at||inputKeys.some(k=>current[k]!==ownerEditing[k])||current.address!==ownerEditing.address){
   setOwnerEditing(null);setNotice('수정 중 다른 사용자가 이 작업내역을 변경했습니다. 최신 수량을 다시 확인해 주세요.');await load();return;
  }
  const {error}=await db.rpc('revise_entry',{
   p_id:ownerEditing.id,p_address:newAddress,
   p_duvet:values.duvet,p_small:values.small,p_large:values.large,
   p_reason:editReason,p_note:ownerEditNote.trim()
  });
  if(error)throw error;
  setOwnerEditing(null);
  if(addressChanged)setKnownAddresses(old=>uniqueAddresses([newAddress,...old]).slice(0,10000));
  setNotice('선택한 작업내역의 주소 및 수량 변경이 저장되었습니다. 수정 전후 기록은 수정이력에 보존됩니다.');
  await load();
 }catch(error){setOwnerEditError(error.message||'수정 중 오류가 발생했습니다.')}
 finally{setOwnerEditSaving(false)}
}
async function copySummary(text){try{await navigator.clipboard.writeText(text);setCopyNotice('복사했습니다.')}catch{setCopyNotice('복사가 차단되었습니다. 아래 정산 내용을 직접 선택해 복사해 주세요.')}}
const total=sum(filtered);
const staffSummary=Object.values(filtered.reduce((acc,r)=>{
 const id=r.employee_id;
 if(!acc[id])acc[id]={id,name:names[id]||'이름 미확인',duvet:0,days:new Set(),count:0};
 acc[id].duvet+=Number(r.duvet)||0;
 acc[id].days.add(workDateKST(r));
 acc[id].count++;
 return acc;
},{})).sort((a,b)=>b.duvet-a.duvet||a.name.localeCompare(b.name,'ko'));
async function login(e){e.preventDefault();setNotice('');if(authMode==='signup'&&(pass.length<8||pass!==confirmPass)){setNotice(pass.length<8?'비밀번호는 8자 이상 입력해 주세요.':'비밀번호 확인이 일치하지 않습니다.');return}setBusy(true);const{error}=authMode==='signup'?await db.auth.signUp({email,password:pass,options:{data:{name}}}):await db.auth.signInWithPassword({email,password:pass});setBusy(false);setNotice(error?error.message:authMode==='signup'?'가입 신청이 접수되었습니다. 이메일 인증 후 사업주 승인을 기다려 주세요.':'')}
async function requestPasswordReset(event){
 event.preventDefault();
 if(busy)return;
 setNotice('');setBusy(true);
 try{
  const {error}=await db.auth.resetPasswordForEmail(email.trim(),{
   redirectTo:`${window.location.origin}/reset-password`
  });
  if(error)throw error;
  setNotice('가입된 이메일이라면 비밀번호 재설정 링크가 발송됩니다. 받은편지함과 스팸함을 확인해 주세요.');
 }catch(error){
  const msg=String(error?.message||'');
  setNotice(/rate limit|too many|429/i.test(msg)?'이메일 발송 한도를 초과했습니다. 잠시 후 다시 시도하거나 관리자에게 이메일 발송 설정을 확인해 달라고 요청해 주세요.':`재설정 이메일 발송에 실패했습니다: ${msg}`);
 }finally{setBusy(false)}
}
function leaveRecovery(){
 window.history.replaceState({},'', '/');
 setRecoveryRoute(false);setRecoveryPass('');setRecoveryConfirm('');setRecoveryError('');
 setAuthMode('login');
}
async function submitNewPassword(event){
 event.preventDefault();
 if(busy)return;
 if(!session){setRecoveryError('재설정 링크가 만료되었거나 유효하지 않습니다. 로그인 화면에서 재설정 메일을 다시 요청해 주세요.');return}
 if(recoveryPass.length<8){setRecoveryError('새 비밀번호는 8자 이상 입력해 주세요.');return}
 if(recoveryPass!==recoveryConfirm){setRecoveryError('새 비밀번호와 비밀번호 확인이 일치하지 않습니다.');return}
 setRecoveryError('');setBusy(true);
 try{
  const{error}=await db.auth.updateUser({password:recoveryPass});
  if(error)throw error;
  await db.auth.signOut();
  leaveRecovery();setPass('');setConfirmPass('');
  setNotice('비밀번호가 변경되었습니다. 새 비밀번호로 로그인해 주세요.');
 }catch(error){setRecoveryError(`비밀번호 변경에 실패했습니다: ${error?.message||'다시 시도해 주세요.'}`)}
 finally{setBusy(false)}
}
async function logout(){if(geoWatchRef.current!==null){navigator.geolocation.clearWatch(geoWatchRef.current);geoWatchRef.current=null}await db.auth.signOut();setTab('entry');setAttendance([]);setLeaves([]);setHrError('');setGeoStatus('')}
async function photo(file){
 if(!file)return;
 setPhotoBusy(true);setPhotoMatches([]);
 setNotice('사진에서 글자를 읽는 중입니다. 한글 OCR은 시간이 걸릴 수 있습니다. 주소를 직접 입력하면 바로 검색할 수 있습니다.');
 try{
  const image=await createImageBitmap(file);
  const scale=Math.min(2,Math.min(1,1200/image.width));
  const canvas=document.createElement('canvas');canvas.width=Math.max(1,Math.round(image.width*scale));canvas.height=Math.max(1,Math.round(image.height*scale));
  const ctx=canvas.getContext('2d',{willReadFrequently:true});ctx.drawImage(image,0,0,canvas.width,canvas.height);image.close();
  const worker=await getOcrWorker();
  const {data}=await worker.recognize(canvas);
  const lines=data.text.split(/\r?\n/).map(x=>x.trim()).filter(Boolean);
  const addressPattern=/(서울|경기|인천|부산|대구|광주|대전|울산|세종|제주|강원|충북|충남|전북|전남|경북|경남|특별시|광역시|특별자치|[가-힣]+[시군구]|[가-힣]+(?:로|길|동|읍|면|리)\s*\d)/;
  const likely=lines.filter(x=>addressPattern.test(x));
  const suggestion=(likely.length?likely:lines).join(' ').replace(/\s+/g,' ').trim();
  const candidates=findPhotoMatches(data.text,knownAddresses);
  setPhotoMatches(candidates);
  setChosenExistingAddress('');
  setForm(previous=>({...previous,address:suggestion}));
  if(candidates.length)setNotice('과거 작업내역에서 비슷한 주소를 찾았습니다. 아래 후보 중 정확한 주소를 선택해 주세요.');
  else if(suggestion)setNotice('같은 주소를 찾지 못했습니다. 주소를 확인·수정한 후 신규 주소로 등록해 주세요.');
  else setNotice('사진에서 주소를 읽지 못했습니다. 주소를 2글자 이상 직접 입력해 주세요.');
 }catch(error){setNotice('사진 인식에 실패했습니다. 위의 주소 검색·직접 입력을 선택해 기존 주소를 검색해 주세요.');}
 finally{setPhotoBusy(false)}
}
async function save(e){
 e.preventDefault();
 if(busy||pendingEntry)return;
 const address=String(form.address||'').trim();
 if(normalizedAddress(address).length<2){setNotice('주소지는 공백과 특수기호를 제외하고 2글자 이상 입력해 주세요.');return}
 const clean={address};
 for(const key of ['duvet','small','large']){
  const raw=form[key];
  if(raw===''||!Number.isInteger(Number(raw))||Number(raw)<0||Number(raw)>2147483647){
   setNotice('수량은 0 이상의 정수로 입력해 주세요.');return;
  }
  clean[key]=Number(raw);
 }
 setNotice('');
 if(!editing){
  setConfirmError('');
  setPendingEntry(clean);
  return;
 }
 setBusy(true);
 const {error}=await db.rpc('revise_entry',{p_id:editing.id,p_address:clean.address,p_duvet:clean.duvet,p_small:clean.small,p_large:clean.large,p_reason:reason,p_note:note});
 setBusy(false);
 if(error){setNotice(error.message);return}
 setNotice('수정 내용이 반영되고 이전 수량이 이력에 보존되었습니다.');
 setEditing(null);setForm(empty());setNote('');setChosenExistingAddress('');setPhotoMatches([]);setAddressMethod('search');
 setKnownAddresses(old=>uniqueAddresses([clean.address,...old]).slice(0,10000));await load();
}
async function confirmNewEntry(){
 if(!pendingEntry||busy||registerInFlightRef.current||!session?.user?.id)return;
 registerInFlightRef.current=true;
 // Use the reviewed snapshot rather than the editable form to prevent mismatched submissions.
 const clean={...pendingEntry};
 setBusy(true);setConfirmError('');
 try{
  const {error}=await db.from('entries').insert({...clean,employee_id:session.user.id});
  if(error)throw error;
  setPendingEntry(null);
  setNotice('확인하신 주소와 수량으로 작업내역이 등록되었습니다.');
  setForm(empty());setNote('');setChosenExistingAddress('');setPhotoMatches([]);setAddressMethod('search');
  setKnownAddresses(old=>uniqueAddresses([clean.address,...old]).slice(0,10000));
  await load();
 }catch(error){setConfirmError(error.message||'등록에 실패했습니다. 다시 확인해 주세요.');}
 finally{registerInFlightRef.current=false;setBusy(false);}
}
function startEdit(r){setAddressMethod('search');setChosenExistingAddress(r.address);setPhotoMatches([]);setEditing(r);setForm({address:r.address,duvet:r.duvet,small:r.small,large:r.large});setReason('주소지 수정');setNote('');setTab('entry');window.scrollTo({top:0,behavior:'smooth'})}
async function approve(p){const{error}=await db.from('profiles').update({approved:!p.approved}).eq('id',p.id);if(error)setNotice(error.message);else load()}
function exportXlsx(){const records=filtered.map(r=>({등록시간:time(r),시간대:period(r),직원:names[r.employee_id]||r.employee_id,주소:r.address,이불커버:r.duvet,'수건 소':r.small,'수건 대':r.large}));const sheet=XLSX.utils.json_to_sheet(records),book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,sheet,'작업내역');XLSX.writeFile(book,`무무베딩_${startDate}_${endDate}.xlsx`)}
const entryTable=(list,allowEdit=false,selectable=false)=><div className="scroll"><table><thead><tr>{allowEdit&&!selectable&&<th>수정</th>}{selectable&&<th><input type="checkbox" aria-label="현재 검색 결과 전체 선택" checked={list.length>0&&list.every(r=>selectedIds.includes(String(r.id)))} onChange={()=>toggleAll(list)}/></th>}<th>등록시간</th><th>주소지</th><th>이불커버</th><th>수건 소</th><th>수건 대</th>{allowEdit&&selectable&&<th>수정</th>}<th>직원</th></tr></thead><tbody>{list.map(r=><tr key={r.id} className={selectable&&selectedIds.includes(String(r.id))?'selected-entry':''}>{allowEdit&&!selectable&&<td><button type="button" onClick={()=>startEdit(r)}>수정</button></td>}{selectable&&<td><input type="checkbox" aria-label={`${r.address} 작업 선택`} checked={selectedIds.includes(String(r.id))} onChange={()=>toggleSelected(r.id)}/></td>}<td>{time(r)}</td><td>{r.address}</td><td>{r.duvet}</td><td>{r.small}</td><td>{r.large}</td>{allowEdit&&selectable&&<td><button type="button" disabled={ownerEditSaving} onClick={()=>openOwnerEdit(r)}>수정</button></td>}<td>{names[r.employee_id]||'직원'}</td></tr>)}</tbody></table>{!list.length&&<p className="muted">등록된 내역이 없습니다.</p>}</div>;
const ownerEditPanel=owner&&ownerEditing&&<section className="panel owner-edit-panel" id="owner-entry-editor"><div className="owner-edit-heading"><div><h2>선택한 작업내역 수정</h2><p className="muted">한 건씩 주소지와 이불커버·수건 수량을 수정할 수 있습니다. 등록 직원과 시간은 유지됩니다.</p></div><button type="button" onClick={()=>{setOwnerEditing(null);setOwnerEditError('')}} disabled={ownerEditSaving}>닫기</button></div><div className="owner-edit-context"><div><span>등록시간</span><strong>{time(ownerEditing)}</strong></div><div><span>주소</span><strong>{ownerEditing.address}</strong></div><div><span>직원</span><strong>{names[ownerEditing.employee_id]||'직원'}</strong></div></div><form onSubmit={submitOwnerEdit}><label className="owner-edit-address">주소지 변경 (2글자 이상)<input type="text" autoComplete="off" required minLength={2} disabled={ownerEditSaving} value={ownerEditForm.address} onChange={e=>{setOwnerEditForm(v=>({...v,address:e.target.value}));if(ownerEditError)setOwnerEditError('')}} placeholder="변경할 주소를 입력해 주세요"/><small>변경 전: {ownerEditing.address}</small></label><div className="owner-edit-quantities">{[['duvet','이불커버'],['small','수건 소'],['large','수건 대']].map(([key,label])=><label key={key}>{label}<div className="owner-edit-counter"><button type="button" aria-label={`${label} 1개 감소`} disabled={ownerEditSaving} onClick={()=>setOwnerEditForm(v=>({...v,[key]:Math.max(0,(Number(v[key])||0)-1)}))}><Minus size={17}/></button><input aria-label={`${label} 수량 수정`} type="text" inputMode="numeric" pattern="[0-9]*" autoComplete="off" required disabled={ownerEditSaving} onFocus={e=>{if(e.currentTarget.value==='0')e.currentTarget.select()}} onClick={e=>{if(e.currentTarget.value==='0')e.currentTarget.select()}} value={ownerEditForm[key]} onChange={e=>setOwnerEditForm(v=>({...v,[key]:e.target.value.replace(/\D/g,'')}))} onBlur={()=>setOwnerEditForm(v=>({...v,[key]:v[key]===''?0:v[key]}))}/><button type="button" aria-label={`${label} 1개 증가`} disabled={ownerEditSaving} onClick={()=>setOwnerEditForm(v=>({...v,[key]:(Number(v[key])||0)+1}))}><Plus size={17}/></button></div><small>변경 전 {ownerEditing[key]}개</small></label>)}</div><label className="owner-edit-note">수정 메모 (선택)<textarea disabled={ownerEditSaving} value={ownerEditNote} onChange={e=>setOwnerEditNote(e.target.value)} placeholder="예: 현장 수량 재확인 후 변경" rows={2}/></label>{ownerEditError&&<p role="alert" className="owner-edit-error">{ownerEditError}</p>}<div className="owner-edit-actions"><button type="submit" className="owner-edit-save" disabled={ownerEditSaving}>{ownerEditSaving?'수정 저장 중...':'주소·수량 수정 저장'}</button><button type="button" disabled={ownerEditSaving} onClick={()=>setOwnerEditing(null)}>취소</button></div><p className="muted">저장하면 작업내역이 갱신되고, 주소 및 수량의 수정 전후 내용이 수정이력에 기록됩니다. 금액은 수량이 바뀐 경우에만 달라집니다.</p></form></section>;
const billingPanel=<section className="panel billing-panel"><div className="billing-heading"><div><h2>선택 작업내역 정산</h2><p className="muted">아래 작업내역에서 체크한 항목만 합산합니다. 기간이나 검색 조건을 변경하면 선택이 초기화됩니다.</p></div><button type="button" onClick={()=>{setSelectedIds([]);setExistingBalanceInput('');setCopyNotice('')}} disabled={!selectedIds.length}>선택 해제</button></div><div className="billing-grid"><div><span>선택 작업</span><strong>{selectedRows.length}건</strong></div><div><span>이불커버</span><strong>{selectedTotal.duvet}개 · {selectedTotal.duvet/2}세트</strong><small>{won(selectedCoverPrice)}</small></div><div><span>수건 소</span><strong>{selectedTotal.small}개</strong><small>{won(selectedSmallPrice)}</small></div><div><span>수건 대</span><strong>{selectedTotal.large}개</strong><small>{won(selectedLargePrice)}</small></div></div><div className="billing-bottom"><div><span>총 청구금액</span><strong>{won(selectedAmount)}</strong></div></div><div className="billing-balance"><label htmlFor="existing-balance">기존 잔액 (원)<span className="balance-input-row"><input id="existing-balance" type="text" inputMode="decimal" autoComplete="off" value={formattedExistingBalance} onChange={e=>{const raw=e.target.value.trim().replace(/^[−–]/,'-');const negative=raw.startsWith('-');const digits=raw.replace(/[^0-9]/g,'').slice(0,12);setExistingBalanceInput(`${negative?'-':''}${digits}`);setCopyNotice('')}} placeholder="예: -100,000" aria-describedby="balance-help"/><button type="button" className="balance-sign-toggle" aria-label="기존 잔액 부호 변경" title="기존 잔액 부호 변경" onClick={()=>{setExistingBalanceInput(v=>v.startsWith('-')?v.slice(1):`-${v}`);setCopyNotice('')}}>+/−</button></span></label><p id="balance-help" className="muted">기존 잔액에서 선택한 작업의 총 청구금액을 차감합니다. 입력한 잔액은 저장되지 않습니다.</p><div className="billing-balance-calculation"><div><span>기존 잔액</span><b>{hasExistingBalance?won(existingBalance):'미입력'}</b></div><div><span>선택한 총 청구금액</span><b>− {won(selectedAmount)}</b></div><div className="billing-remaining"><span>잔액</span><strong className={hasExistingBalance&&remainingBalance<0?'negative-balance':''}>{hasExistingBalance?won(remainingBalance):'입력 대기'}</strong></div></div>{showDepositNotice&&<p className="billing-low-notice" role="status">입금 확인 부탁드려요</p>}</div><div className="billing-actions"><button type="button" className="billing-copy" disabled={!selectedRows.length||!hasExistingBalance} onClick={()=>copySummary(selectedCopyText)}>수량·잔액 복사</button></div>{!hasExistingBalance&&<p className="muted">수량과 잔액을 복사하려면 기존 잔액을 입력해 주세요.</p>}{copyNotice&&<p role="status" className="muted">{copyNotice}</p>}<textarea className="billing-preview" readOnly value={selectedCopyText} placeholder="작업내역을 선택하고 기존 잔액을 입력하면 복사할 내용이 표시됩니다." aria-label="복사용 수량 및 잔액"/></section>;
if(!db)return <div className="center"><div className="login"><h1>무무베딩 수량관리</h1><p>클라우드 연결 전입니다. 프로젝트의 README 안내에 따라 Supabase 환경변수를 설정해 주세요.</p></div></div>;
if(!ready)return <div className="center">불러오는 중...</div>;
if(recoveryRoute)return <div className="center"><form className="login password-reset-form" onSubmit={submitNewPassword}>
 <div className="brand">MUMU BEDDING</div><h1>새 비밀번호 설정</h1>
 {recoveryUrlError?<p className="password-error" role="alert">재설정 링크가 유효하지 않거나 만료되었습니다. 새 링크를 요청해 주세요.</p>:session?<>
 <p>사용할 새 비밀번호를 입력해 주세요. 기존 작업내역과 계정은 유지됩니다.</p>
 <label>새 비밀번호<input type="password" autoComplete="new-password" minLength={8} required value={recoveryPass} onChange={e=>setRecoveryPass(e.target.value)} placeholder="8자 이상 입력"/></label>
 <label>새 비밀번호 확인<input type="password" autoComplete="new-password" minLength={8} required value={recoveryConfirm} onChange={e=>setRecoveryConfirm(e.target.value)} placeholder="새 비밀번호 다시 입력"/></label>
 {recoveryConfirm&&<p className={recoveryPass===recoveryConfirm?'password-ok':'password-error'}>{recoveryPass===recoveryConfirm?'✓ 비밀번호가 일치합니다.':'비밀번호가 일치하지 않습니다.'}</p>}
 <button className="primary" disabled={busy||recoveryPass.length<8||recoveryPass!==recoveryConfirm}>{busy?'변경 중...':'비밀번호 변경하기'}</button>
 </>:<p className="password-error" role="alert">재설정 인증을 확인할 수 없습니다. 링크가 만료되었다면 새 링크를 요청해 주세요.</p>}
 {recoveryError&&<div className="notice" role="alert">{recoveryError}</div>}
 <button type="button" className="plain" disabled={busy} onClick={()=>{leaveRecovery();setAuthMode('forgot');setNotice('')}}>재설정 링크 다시 요청</button>
 </form></div>;
if(!session)return <div className="center"><form className="login" onSubmit={authMode==='forgot'?requestPasswordReset:login}>
 <div className="brand">MUMU BEDDING</div>
 <h1>{authMode==='login'?'수량관리 로그인':authMode==='signup'?'직원 회원가입':'비밀번호 찾기'}</h1>
 <p>{authMode==='forgot'?'가입할 때 사용한 이메일을 입력하면 비밀번호 재설정 링크를 보내드립니다.':'현장 수량 기록과 실시간 운영 현황'}</p>
 {authMode==='signup'&&<label>직원 이름<input required value={name} onChange={e=>setName(e.target.value)} placeholder="이름" autoComplete="name"/></label>}
 <label>이메일<input type="email" required value={email} onChange={e=>setEmail(e.target.value)} placeholder="name@example.com" autoComplete="email"/></label>
 {authMode!=='forgot'&&<label>비밀번호<input type="password" autoComplete={authMode==='signup'?'new-password':'current-password'} minLength={authMode==='signup'?8:6} required value={pass} onChange={e=>setPass(e.target.value)} placeholder={authMode==='signup'?'8자 이상 입력':'비밀번호 입력'}/></label>}
 {authMode==='signup'&&<><label>비밀번호 확인<input type="password" autoComplete="new-password" minLength={8} required value={confirmPass} onChange={e=>setConfirmPass(e.target.value)} placeholder="비밀번호 다시 입력"/></label>{confirmPass&&<p className={pass===confirmPass?'password-ok':'password-error'}>{pass===confirmPass?'✓ 비밀번호가 일치합니다.':'비밀번호가 일치하지 않습니다.'}</p>}</>}
 <button className="primary" disabled={busy||(authMode==='signup'&&(pass.length<8||pass!==confirmPass))}>{busy?'처리 중...':authMode==='login'?'로그인':authMode==='signup'?'가입 신청':'재설정 이메일 보내기'}</button>
 {authMode==='login'&&<><button type="button" className="plain" onClick={()=>{setAuthMode('forgot');setPass('');setConfirmPass('');setNotice('')}}>비밀번호를 잊으셨나요?</button><button type="button" className="plain" onClick={()=>{setAuthMode('signup');setPass('');setConfirmPass('');setNotice('')}}>신규 직원 회원가입</button></>}
 {authMode!=='login'&&<button type="button" className="plain" onClick={()=>{setAuthMode('login');setPass('');setConfirmPass('');setNotice('')}}>로그인으로 돌아가기</button>}
 {notice&&<div className="notice" role="status">{notice}</div>}
 </form></div>;
if(!profile)return <div className="center">계정 정보를 불러오는 중...</div>;
if(!profile.approved)return <div className="center"><div className="login"><div className="brand">MUMU BEDDING</div><h2>사업주 승인 대기 중</h2><p>{profile.name}님의 계정이 생성되었습니다. 사업주 승인 후 수량을 등록할 수 있습니다.</p><button onClick={()=>location.reload()}>승인 상태 새로고침</button><button className="plain" onClick={logout}>로그아웃</button></div></div>;
const tabs=owner?[['dashboard','대시보드',LayoutDashboard],['entries','작업내역',ClipboardList],['history','수정이력',History],['staff','직원관리',Users]]:[['entry','수량 등록',Plus],['clock',currentAttendance?'퇴근':'출근',Clock3],['leave','휴가 신청',CalendarDays],['mine','내 기록 수정',ClipboardList],['history','수정이력',History]];
const shownChanges=changes.filter(x=>owner||x.changed_by===session.user.id);
return <div className="shell"><header><div><div className="brand">MUMU BEDDING</div><h1>수량관리 시스템</h1></div><div className="user"><span>{profile.name} {owner?'· 사업주':''}</span><button title="로그아웃" onClick={logout}><LogOut size={17}/></button></div></header><nav>{tabs.map(([id,label,Icon])=><button key={id} className={tab===id?'selected':''} onClick={()=>setTab(id)}><Icon size={17}/>{label}</button>)}</nav><main>{notice&&<div className="notice">{notice}<button onClick={()=>setNotice('')}>닫기</button></div>}
{tab==='entry'&&<section className="panel editor"><h2>{editing?'작업내역 수정':'작업 수량 등록'}</h2><form onSubmit={save}><div className="address-entry-block">
 <div className="address-method-switch" role="tablist" aria-label="주소 입력 방식">
  <button type="button" role="tab" aria-selected={addressMethod==='search'} className={addressMethod==='search'?'active':''} onClick={()=>setAddressMethod('search')}>주소 검색 · 직접 입력</button>
  <button type="button" role="tab" aria-selected={addressMethod==='photo'} className={addressMethod==='photo'?'active':''} onClick={()=>setAddressMethod('photo')}><Camera size={17}/> 사진으로 주소 읽기</button>
 </div>
 {addressMethod==='search'?<div className="address-method-content">
  <div className="address-instruction"><strong>기존 주소 검색 또는 신규 주소 입력</strong><p>주소 일부를 2글자 이상 입력하면 과거 작업내역에서 자동 검색합니다. 일치하는 주소가 없으면 새 주소를 직접 입력하세요.</p></div>
  <label>주소지 (2글자 이상)<input autoComplete="off" required minLength={2} value={form.address} onChange={e=>editAddress(e.target.value)} placeholder="예: 역삼동, 홍대입구, ○○로 12"/></label>
  {chosenExistingAddress===form.address&&form.address&&<p className="address-picked">✓ 기존 작업내역의 주소를 선택했습니다.</p>}
  {addressLookupState==='loading'&&<p className="muted">기존 주소 목록을 불러오는 중입니다...</p>}
  {addressLookupState==='error'&&<p className="address-error">주소 검색 준비가 필요합니다. Supabase에 기존 주소 조회 SQL을 먼저 실행해 주세요. 지금은 직접 입력할 수 있습니다.</p>}
  {addressInputNormalized.length>=2&&addressLookupState==='ready'&&chosenExistingAddress!==form.address&&<div className="address-candidates"><strong>기존 작업내역 주소 {matchingAddresses.length?'후보':'검색 결과'}</strong>{matchingAddresses.length?matchingAddresses.map(item=><button className="address-candidate" type="button" key={item.address} onClick={()=>chooseAddress(item.address)}>{item.address}</button>):<p className="address-empty">일치하는 주소가 없습니다. 신규 작업이라면 위 칸에 정확한 주소를 입력한 뒤 등록하세요.</p>}</div>}
 </div>:<div className="address-method-content photo-within-entry">
  <p className="muted">필요할 때만 사진을 촬영해 주소를 인식합니다. 한글 OCR은 시간이 걸릴 수 있으므로 빠르게 검색하려면 위의 ‘주소 검색·직접 입력’을 선택하세요.</p>
  <div className="optional-photo"><label className="camera"><Camera size={22}/>{photoBusy?'사진 분석 중...':'주소 사진 촬영 · 업로드'}<input type="file" accept="image/*" capture="environment" disabled={photoBusy} onChange={e=>{photo(e.target.files?.[0]);e.target.value=''}}/></label></div>
  {addressLookupState==='loading'&&<p className="muted">기존 주소 목록을 불러오는 중입니다...</p>}
  {addressLookupState==='error'&&<p className="address-error">기존 주소 조회 설정을 확인해 주세요. 인식한 주소는 직접 수정할 수 있습니다.</p>}
  {photoMatches.length>0&&<div className="address-candidates photo-candidates"><strong>사진에서 찾은 유사 주소 (정확한 주소를 선택하세요)</strong>{photoMatches.map(item=><button type="button" className="address-candidate" key={item.address} onClick={()=>{chooseAddress(item.address);setAddressMethod('search')}}>{item.address}</button>)}<button type="button" onClick={()=>setPhotoMatches([])}>일치하는 주소 없음 · 새 주소 직접 입력</button></div>}
  <label>사진 인식 결과 또는 신규 주소 (2글자 이상)<input autoComplete="off" required minLength={2} value={form.address} onChange={e=>editAddress(e.target.value)} placeholder="인식한 주소를 확인하거나 직접 입력하세요"/></label>
  <button type="button" className="photo-confirm" disabled={photoBusy||normalizedAddress(form.address).length<2} onClick={()=>{setAddressMethod('search');setNotice('주소를 확인한 뒤 수량을 입력해 주세요.')}}>주소 확인 · 수량 입력하기</button>
 </div>}
 </div><div className="quantities">{[['duvet','이불커버'],['small','수건 소'],['large','수건 대']].map(([k,label])=><div className="quantity" key={k}><span>{label}</span>{k==='duvet'&&<small className="duvet-description">큰 것은 모두 이불로 세며, 다리지 않은 것도 포함합니다.</small>}<div><button type="button" onClick={()=>setForm(v=>({...v,[k]:Math.max(0,(Number(v[k])||0)-1)}))}><Minus size={18}/></button><input className="quantity-input" type="text" inputMode="numeric" pattern="[0-9]*" autoComplete="off" onFocus={e=>{if(e.currentTarget.value==='0')e.currentTarget.select()}} onClick={e=>{if(e.currentTarget.value==='0')e.currentTarget.select()}} aria-label={`${label} 수량 직접 입력`} value={form[k]} onChange={e=>setForm(v=>({...v,[k]:e.target.value===""?"":e.target.value.replace(/\D/g,"")}))} onBlur={()=>setForm(v=>({...v,[k]:v[k]===""?0:v[k]}))}/><button type="button" onClick={()=>setForm(v=>({...v,[k]:(Number(v[k])||0)+1}))}><Plus size={18}/></button></div></div>)}</div>{editing&&<><label>수정 사유<select value={reason} onChange={e=>setReason(e.target.value)}>{['주소지 수정','이불커버 수량 수정','수건 소 수량 수정','수건 대 수량 수정','전체 수정'].map(x=><option key={x}>{x}</option>)}</select></label><label>상세 수정 사유<textarea value={note} onChange={e=>setNote(e.target.value)} placeholder="수정 이유를 입력하세요"/></label></>}<button className="primary" disabled={busy||photoBusy}>{busy||photoBusy?'처리 중...':editing?'수정 즉시 반영':'주소·수량 확인하기'}</button>{editing&&<button type="button" onClick={()=>{setEditing(null);setForm(empty());setChosenExistingAddress('');setPhotoMatches([]);setAddressMethod('search')}}>수정 취소</button>}</form><dialog ref={confirmDialogRef} className="entry-confirm-dialog" aria-labelledby="entry-confirm-title" onCancel={e=>{if(busy)e.preventDefault();else setPendingEntry(null)}} onClose={()=>{if(!busy)setPendingEntry(null)}}>{pendingEntry&&<div className="entry-confirm-content"><div className="entry-confirm-badge">등록 전 확인</div><h2 id="entry-confirm-title">주소와 수량이 맞나요?</h2><p className="entry-confirm-explanation">다시 한번 확인한 뒤 <b>확인 후 등록</b>을 눌러야 작업내역이 저장됩니다.</p><dl className="entry-confirm-details"><div className="entry-confirm-address"><dt>주소지</dt><dd>{pendingEntry.address}</dd></div><div><dt>이불커버</dt><dd>{pendingEntry.duvet.toLocaleString('ko-KR')}개</dd></div><div><dt>수건 소</dt><dd>{pendingEntry.small.toLocaleString('ko-KR')}개</dd></div><div><dt>수건 대</dt><dd>{pendingEntry.large.toLocaleString('ko-KR')}개</dd></div></dl>{confirmError&&<p className="entry-confirm-error" role="alert">{confirmError}</p>}<div className="entry-confirm-actions"><button type="button" autoFocus disabled={busy} onClick={()=>{setPendingEntry(null);setConfirmError('')}}>돌아가서 수정</button><button type="button" className="entry-confirm-submit" disabled={busy} onClick={confirmNewEntry}>{busy?'등록 중...':'확인 후 등록'}</button></div></div>}</dialog></section>}
{['dashboard','entries','mine'].includes(tab)&&<><div className="toolbar"><div className="date-range"><div className="preset-buttons">{[['today','오늘'],['twoWeeks','최근 2주'],['month','최근 한 달']].map(([id,label])=><button key={id} type="button" className={preset===id?'selected-preset':''} onClick={()=>choosePreset(id)}>{label}</button>)}</div><label>시작일 <input type="date" value={startDate} max={endDate} onChange={e=>{setStartDate(e.target.value);setPreset('custom')}}/></label><label>종료일 <input type="date" value={endDate} min={startDate} onChange={e=>{setEndDate(e.target.value);setPreset('custom')}}/></label></div>{owner&&<><label>직원<select value={person} onChange={e=>setPerson(e.target.value)}><option value="all">전체 직원</option>{staff.filter(p=>p.approved).map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select></label><label>시간대<select value={shift} onChange={e=>setShift(e.target.value)}><option>전체</option><option>오전</option><option>오후</option></select></label><button onClick={exportXlsx}><Download size={16}/> 엑셀</button></>}<button onClick={load}><RefreshCw size={16}/> 새로고침</button></div><div className="cards">{[['활동 직원',people(filtered)+'명'],['이불커버',total.duvet+'개'],['수건 소',total.small+'개'],['수건 대',total.large+'개'],['작업 건수',filtered.length+'건']].map(([k,v])=><div className="metric" key={k}><span>{k}</span><strong>{v}</strong></div>)}</div></>}

{hrError&&<p role="alert" className="hr-error">{hrError}</p>}
{tab==='clock'&&!owner&&<section className="panel hr-panel"><h2><Clock3 size={21}/> 출퇴근 기록</h2>
 <div className="hr-office"><MapPin size={18}/><div><strong>근무지 · 메트로폴리스</strong><p>{workplace.address||WORKPLACE_ADDRESS}</p><small>기준 반경 {workplace.radius_m}m · 위치 권한 필요</small></div></div>
 <div className="hr-attendance-status"><div><span>현재 상태</span><strong>{currentAttendance?'출근 중':'퇴근 상태'}</strong></div><div><span>출근 시각</span><strong>{currentAttendance?asKoreanDateTime(currentAttendance.check_in_at):'—'}</strong></div><div><span>현재 위치 거리</span><strong>{lastDistance===null?'위치 확인 전':lastDistance+'m'}</strong></div></div>
 <div className="hr-actions"><button className="primary" type="button" disabled={hrBusy||(!currentAttendance&&!!hrError)} onClick={()=>currentAttendance?clockOut('manual'):clockIn()}>{hrBusy?'처리 중…':currentAttendance?'퇴근하기':'출근하기'}</button></div>
 {missedAttendance.length>0&&<p className="hr-caution">퇴근 기록이 없는 지난 출근 {missedAttendance.length}건이 있습니다. 전날 기록은 ‘미퇴근’으로 남으며 새 날짜의 출근은 가능합니다. 필요하면 사업주에게 기록 확인을 요청하세요.</p>}
 {geoStatus&&<p role="status" className="muted">{geoStatus}</p>}
 <p className="hr-caution">근무지를 벗어나면 웹앱이 실행 중일 때 자동 퇴근할 수 있습니다. 직접 퇴근하려면 ‘퇴근하기’를 누르세요. 자정이 지나면 다시 ‘출근하기’로 바뀌며, 퇴근하지 않은 전날 기록은 ‘미퇴근’으로 남습니다. 앱을 닫거나 휴대전화가 절전 상태이면 자동 퇴근이 보장되지 않습니다. 업무상 배송을 위해 지역을 벗어나도 자동 퇴근될 수 있습니다.</p>
 <h3>내 최근 출퇴근</h3><div className="scroll"><table><thead><tr><th>출근</th><th>퇴근</th><th>퇴근 방식</th></tr></thead><tbody>{attendance.filter(a=>a.employee_id===session.user.id).slice(0,30).map(a=><tr key={a.id}><td>{asKoreanDateTime(a.check_in_at)}</td><td>{asKoreanDateTime(a.check_out_at)}</td><td>{attendanceStateLabel(a)}</td></tr>)}{!attendance.some(a=>a.employee_id===session.user.id)&&<tr><td colSpan={3}>출퇴근 기록이 없습니다.</td></tr>}</tbody></table></div>
 </section>}
{tab==='leave'&&!owner&&<section className="panel hr-panel"><h2><CalendarDays size={21}/> 휴가 신청</h2>
 <form className="hr-leave-form" onSubmit={submitLeave}><label>휴가 유형<select value={leaveType} onChange={e=>setLeaveType(e.target.value)}><option>연차</option><option>오전 반차</option><option>오후 반차</option><option>기타</option></select></label><label>시작일<input type="date" required value={leaveStart} onChange={e=>{setLeaveStart(e.target.value);if(e.target.value>leaveEnd)setLeaveEnd(e.target.value)}}/></label><label>종료일<input type="date" required min={leaveStart} value={leaveEnd} onChange={e=>setLeaveEnd(e.target.value)}/></label><label className="hr-note">사유 (선택)<textarea rows={2} maxLength={1000} placeholder="전달할 내용이 있으면 입력하세요" value={leaveNote} onChange={e=>setLeaveNote(e.target.value)}/></label><button className="primary" disabled={hrBusy||!!hrError}>휴가 신청하기</button></form>
 <h3>내 휴가 신청 내역</h3><div className="scroll"><table><thead><tr><th>기간</th><th>유형</th><th>상태</th><th>신청 취소</th></tr></thead><tbody>{leaves.filter(l=>l.employee_id===session.user.id).map(l=><tr key={l.id}><td>{l.start_date}{l.end_date!==l.start_date?' ~ '+l.end_date:''}</td><td>{l.leave_type}</td><td>{({pending:'승인 대기',approved:'승인',rejected:'반려',cancelled:'취소'})[l.status]||l.status}</td><td>{l.status==='pending'?<button disabled={hrBusy} onClick={()=>cancelLeave(l.id)}>취소</button>:'—'}</td></tr>)}{!leaves.some(l=>l.employee_id===session.user.id)&&<tr><td colSpan={4}>휴가 신청 내역이 없습니다.</td></tr>}</tbody></table></div>
 </section>}
{tab==='dashboard'&&<section className="panel hr-panel"><div className="hr-dashboard-title"><h2>출퇴근 · 휴가 현황</h2><button onClick={loadHr} disabled={hrBusy}><RefreshCw size={15}/> 새로고침</button></div>
 <div className="hr-stats"><div><span>오늘 출근한 인원</span><strong>{new Set(attendanceToday.map(a=>a.employee_id)).size}명</strong></div><div><span>현재 출근 중</span><strong>{attendance.filter(a=>!a.check_out_at&&attendanceWorkDay(a)===todayKst).length}명</strong></div><div><span>휴가 승인 대기</span><strong>{pendingLeaves.length}건</strong></div><div><span>오늘 승인 휴가</span><strong>{leaves.filter(l=>l.status==='approved'&&l.start_date<=dateKST()&&l.end_date>=dateKST()).length}건</strong></div></div>
 <h3>출퇴근 내역 · 선택 기간</h3><div className="scroll"><table><thead><tr><th>직원</th><th>출근 시간</th><th>퇴근 시간</th><th>퇴근 방식</th></tr></thead><tbody>{ownerAttendByDate.map(a=><tr key={a.id}><td>{names[a.employee_id]||'직원'}</td><td>{asKoreanDateTime(a.check_in_at)}</td><td>{asKoreanDateTime(a.check_out_at)}</td><td>{attendanceStateLabel(a)}</td></tr>)}{!ownerAttendByDate.length&&<tr><td colSpan={4}>선택 기간의 출퇴근 기록이 없습니다.</td></tr>}</tbody></table></div>
 <h3>휴가 신청 · 승인/반려</h3><div className="scroll"><table><thead><tr><th>직원</th><th>휴가 기간</th><th>유형</th><th>상태</th><th>처리</th></tr></thead><tbody>{leaves.map(l=><tr key={l.id}><td>{names[l.employee_id]||'직원'}</td><td>{l.start_date}{l.start_date!==l.end_date?' ~ '+l.end_date:''}</td><td>{l.leave_type}</td><td>{({pending:'승인 대기',approved:'승인',rejected:'반려',cancelled:'취소'})[l.status]||l.status}</td><td>{l.status==='pending'?<span className="hr-actions-inline"><button disabled={hrBusy} onClick={()=>reviewLeave(l.id,'approved')}>승인</button><button disabled={hrBusy} onClick={()=>reviewLeave(l.id,'rejected')}>반려</button></span>:'—'}</td></tr>)}{!leaves.length&&<tr><td colSpan={5}>휴가 신청 내역이 없습니다.</td></tr>}</tbody></table></div>
 <form className="hr-radius-settings" onSubmit={updateRadius}><strong><MapPin size={16}/> 출퇴근 기준 근무지</strong><p>{workplace.address||WORKPLACE_ADDRESS}</p><label>근무지 반경 (100~2,000m)<input inputMode="numeric" type="number" min="100" max="2000" step="10" value={radiusInput} onChange={e=>setRadiusInput(e.target.value)}/></label><button disabled={hrBusy||!!hrError} type="submit">반경 저장</button><p className="muted">위치는 중앙로1261번길 77 기준입니다. 반경 안에서 출근할 수 있고, 밖으로 나가면 위치가 안정적으로 확인되는 경우 자동 퇴근합니다.</p></form>
 </section>}
{tab==='dashboard'&&<section className="panel"><h2>직원별 이불커버 · 근무일수</h2><p className="muted">선택한 기간의 등록 기록을 기준으로 집계합니다. 근무일수는 기록이 1건 이상 있는 날짜를 하루로 계산하며, 실제 출퇴근 기록과는 다를 수 있습니다.</p><div className="scroll"><table><thead><tr><th>직원</th><th>이불커버 합계</th><th>근무일수</th><th>작업 건수</th></tr></thead><tbody>{staffSummary.map(p=><tr key={p.id}><td>{p.name}</td><td>{p.duvet.toLocaleString('ko-KR')}개</td><td>{p.days.size}일</td><td>{p.count.toLocaleString('ko-KR')}건</td></tr>)}{!staffSummary.length&&<tr><td colSpan={4}>해당 기간의 등록 기록이 없습니다.</td></tr>}</tbody></table></div></section>}
{tab==='entries'&&<>{billingPanel}{ownerEditPanel}<section className="panel"><div className="owner-list-heading"><div><h2>전체 작업 내역</h2><p className="muted">체크박스로 정산할 작업을 선택하세요. 주소·수량 수정은 한 건만 선택하거나 해당 행의 ‘수정’ 버튼을 누르세요.</p></div><button type="button" className="owner-edit-selected" disabled={selectedRows.length!==1||ownerEditSaving} onClick={editSelectedOwnerEntry}>선택한 1건 주소·수량 수정</button></div><div className="owner-address-search"><label htmlFor="owner-entry-address-search"><Search size={16}/> 주소 검색</label><input id="owner-entry-address-search" type="search" value={addressSearch} onChange={e=>setAddressSearch(e.target.value)} placeholder="주소 일부를 입력해 작업내역 검색"/><span className="muted">{filtered.length.toLocaleString('ko-KR')}건 표시</span></div>{entryTable(filtered,true,true)}</section></>}
{tab==='mine'&&<section className="panel"><h2>내 기록 수정 <small>작업내역을 확인하고 수정할 수 있습니다.</small></h2>{entryTable(filtered,true)}</section>}

{tab==='history'&&<section className="panel"><h2>수정 이력 <small>수정 전후 수량은 삭제되지 않습니다.</small></h2><div className="history">{shownChanges.map(c=><div className="change" key={c.id}><div className="changehead"><b>{c.reason}</b><span>{time(c)} · {names[c.changed_by]||'직원'}</span></div><div className="diff"><div><small>수정 전</small><p>{c.before_values.address}</p><p>이불 {c.before_values.duvet} · 소 {c.before_values.small} · 대 {c.before_values.large}</p></div><div><small>수정 후</small><p>{c.after_values.address}</p><p>이불 {c.after_values.duvet} · 소 {c.after_values.small} · 대 {c.after_values.large}</p></div></div>{c.note&&<p className="muted">사유: {c.note}</p>}</div>)}{!shownChanges.length&&<p className="muted">수정 이력이 없습니다.</p>}</div></section>}
{tab==='staff'&&<section className="panel"><h2>직원 관리</h2><p className="muted">직원 이름과 가입 이메일을 확인하고, 가입 승인 상태를 관리할 수 있습니다.</p>{staffEmailError&&<p className="staff-email-warning" role="alert">{staffEmailError}</p>}{staff.map(p=><div className="staff" key={p.id}><div className="staff-identity"><b>{p.name}</b><small className="staff-email">{staffEmails[p.id]|| (staffEmailError?'이메일 조회 설정 필요':'등록 이메일 없음')}</small><small>{p.role==='owner'?'사업주':p.approved?'활동 가능':'승인 대기 / 정지'}</small></div>{p.role!=='owner'&&<button onClick={()=>approve(p)}>{p.approved?'계정 정지':'승인 / 활성화'}</button>}</div>)}</section>}
</main><footer>© 무무베딩 · 현장 수량관리</footer></div>}
createRoot(document.getElementById('root')).render(<App/>);
