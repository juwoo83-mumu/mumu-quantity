# 무무베딩 Vercel 배포 안내

## 1. GitHub에 코드 올리기
1. 이 ZIP 압축을 풉니다.
2. GitHub에서 새 **Private** repository `mumu-quantity`를 생성합니다.
3. 압축 해제한 폴더의 파일들을 GitHub 저장소에 업로드합니다. `.env` 파일은 올리지 마세요 (`.gitignore`에 제외되어 있음).

## 2. Vercel에서 배포
1. https://vercel.com/new 에서 GitHub 계정 연결
2. `mumu-quantity` 저장소 Import
3. Framework Preset: **Vite**
4. Root Directory: `./`
5. Build Command: `npm run build`
6. Output Directory: `dist`
7. Environment Variables에 아래 두 개 입력:
   - `VITE_SUPABASE_URL` = `https://kaebuplwtzpfajqpazid.supabase.co`
   - `VITE_SUPABASE_ANON_KEY` = Supabase의 `sb_publishable_`로 시작하는 **공개용 키**
8. Deploy 클릭. 생성된 `https://...vercel.app` 주소로 접속.

## 3. Supabase 인증 URL
Supabase Dashboard → Authentication → URL Configuration:
- Site URL: 실제 배포된 `https://...vercel.app` 주소
- Redirect URLs: `https://...vercel.app/**` 추가
이메일 확인 기능이 활성화된 경우 인증 메일을 확인하세요.

## 4. 최초 사업주 계정
1. 배포된 웹앱에서 사업주 이메일로 회원가입
2. Supabase SQL Editor에서 다음 쿼리 실행. 이메일을 실제 가입 이메일로 변경:

```sql
update public.profiles p
set role='owner', approved=true
from auth.users u
where p.id=u.id and u.email='YOUR_OWNER_EMAIL@example.com';
```

3. 웹앱 로그아웃 후 다시 로그인하면 관리자 화면이 열립니다.

## 5. 직원 계정
직원은 회원가입 → 사업주 승인 대기 → 사업주가 직원관리에서 승인 → 사용 시작.

## 참고
- `supabase.sql`은 이미 성공적으로 실행했다면 다시 실행하지 마세요. 기존 테이블 생성 오류가 발생할 수 있습니다.
- `service_role` 또는 `sb_secret_` 키를 웹앱 환경변수로 사용하지 마세요.
- 이 ZIP은 Vercel용 소스 프로젝트입니다. Vercel에 직접 배포된 상태는 아닙니다.
- 현재 환경에서 npm 의존성 다운로드가 시간 초과되어 프로덕션 빌드 성공 여부는 아직 검증하지 못했습니다. Vercel의 빌드 로그로 확인하세요.
