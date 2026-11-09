# AutoPatcher 모니터링 대시보드 디자인 스펙 (Figma Design Spec)

## 1. 개요
* **대상 프로젝트**: AutoPatcher (HDS Auto Patcher)
* **대상 화면**: 버전 모니터링 탭 (Dashboard View)
* **프레임 규격**: 1440 x 900 (Desktop HD 기준)
* **디자인 프리뷰 파일**: [Monitoring_Dashboard_Sample.html](./Monitoring_Dashboard_Sample.html)

---

## 2. 디자인 토큰 (Design Tokens)

### 색상 (Color Palette)
| 토큰명 | Hex 코드 | 용도 |
| :--- | :--- | :--- |
| `bg/canvas` | `#111822` | 전체 배경 캔버스 |
| `bg/panel` | `#1B2533` | 윈도우 메인 탭 및 패널 배경 |
| `bg/card` | `#24303F` | 서버 버전 및 상태 집계 카드 배경 |
| `bg/card-hover` | `#2D3C4E` | 카드 호버 상태 배경 |
| `border/subtle` | `#2E3A4D` | 구분선 및 헤더 경계선 |
| `border/strong` | `#3B4A60` | 버튼 테두리 및 카드 윤곽선 |
| `text/primary` | `#F5EBDC` | 헤더 및 주요 텍스트 (아이보리 화이트) |
| `text/secondary` | `#9FB0C4` | 보조 설명 및 레이블 텍스트 (라이트 슬레이트) |
| `text/muted` | `#627285` | 갱신 시간 및 비활성 텍스트 |
| `status/uptodate` | `#2E8B57` | **일치** (Sea Green) |
| `status/mismatch` | `#CD5C5C` | **불일치** (Indian Red) |
| `status/pending` | `#E08A00` | **적용대기** (Amber Orange) |
| `status/failed` | `#8B0000` | **실패** (Dark Red) |
| `status/noresponse` | `#6B6B6B` | **무응답** (Gray) |
| `status/notinstalled` | `#3A3A3A` | **미설치** (Dark Charcoal) |
| `accent/green` | `#7CFC00` | 중앙 서버 버전 강조 (Lawn Green) |
| `accent/blue` | `#2563EB` | 주요 액션 버튼 (Royal Blue) |
| `accent/emerald` | `#059669` | 에이전트 배포 액션 (Emerald) |

---

## 3. 컴포넌트 구조 (Figma Frame Hierarchy)

```text
[Frame: AutoPatcher_Monitoring_View (1440 x 900)]
 ├── [Header_Window_Bar (H: 38px, Bg: #141C27)]
 │    ├── Title & Icon ("HDS Auto Patcher [ Version 2.0.4 ]")
 │    ├── Site Badge ("SITE: BGA")
 │    └── Window Controls (Min, Max, Close)
 ├── [Header_Action_Row (H: 56px)]
 │    ├── TabBar ("모니터링" [Active], "패치")
 │    └── ActionButtonGroup (Horizontal AutoLayout, Gap: 10px)
 │         ├── Button_Secondary ("새로고침")
 │         ├── Button_Primary ("배포 등록")
 │         └── Button_Success ("에이전트 배포")
 ├── [Header_Status_Row (H: 64px, Bg: #1B2533)]
 │    ├── TitleLabel ("BGA 버전 모니터링", 20pt Bold)
 │    ├── ServerCardGroup (Gap: 10px)
 │    │    ├── Card ("중앙 Main", Version: v2.1.0)
 │    │    └── Card ("중앙 Vision", Version: v2.1.0)
 │    └── StatusGrid (6 Columns Equal Width, Gap: 8px)
 │         ├── StatusCard ("일치", Count: 48, Dot: #2E8B57)
 │         ├── StatusCard ("불일치", Count: 6, Dot: #CD5C5C)
 │         ├── StatusCard ("적용대기", Count: 2, Dot: #E08A00)
 │         ├── StatusCard ("실패", Count: 1, Dot: #8B0000)
 │         ├── StatusCard ("무응답", Count: 3, Dot: #6B6B6B)
 │         └── StatusCard ("미설치", Count: 0, Dot: #3A3A3A)
 ├── [Table_Container (Flex: 1, AutoLayout)]
 │    └── DataGrid_Table
 │         ├── TableHeader (설비, Main, Vision1, Vision2, Vision3, Vision4, Vision5)
 │         └── TableRows (Equipment Rows with Status Badges)
 └── [Footer_Bar (H: 34px, Bg: #141C27)]
      ├── CentralServerPath & ExcelPath
      └── AutoRefreshIndicator ("자동 갱신 60초 · 마지막 갱신 10:48:20")
```

---

## 4. Figma로 가져오기 (Import into Figma)
1. **HTML to Figma 플러그인 이용 시**:
   - Figma 데스크톱 또는 웹 앱에서 빈 디자인 파일을 엽니다.
   - `Plugins` 검색창에서 **HTML to Design** 또는 **html.to.design** 플러그인을 실행합니다.
   - 본 프로젝트의 `Design/Monitoring_Dashboard_Sample.html` 파일을 브라우저로 연 뒤 해당 로컬 주소를 붙여넣거나 파일 업로드하면, 모든 레이어/텍스트/벡터가 온전한 Figma 프레임 및 AutoLayout으로 변환됩니다.
2. **Figma REST API 연동 시**:
   - 발급해주신 Personal Access Token을 통해 향후 프로젝트/파일 조회가 가능합니다.
