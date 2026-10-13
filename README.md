# QuickLaunchOnTray

### 소개
QuickLaunchOnTray는 Windows 알림 영역(시스템 트레이)에 사용자가 지정한 프로그램과 폴더의 아이콘을 등록해
한 번의 클릭으로 실행할 수 있게 해 주는 유틸리티입니다.

- 실행 파일과 같은 폴더의 `config.ini`만 읽습니다.
- 등록할 아이콘 수에는 제한이 없습니다(항목당 아이콘 1개).
- 잘못된 항목이 있어도 나머지 항목은 정상 등록되고, 건너뛴 항목은 시작할 때 경고로 알려 줍니다.

### 요구사항

| 항목 | 값 |
|---|---|
| OS | Windows 7 SP1 이상 (Windows 10/11 권장) |
| 런타임 | .NET Framework 4.8 |
| DPI | PerMonitorV2 인식(매니페스트 선언). 배율에 맞는 트레이 아이콘 크기를 사용합니다. |

### 설치

1. `QuickLaunchOnTray.exe`(와 한국어 UI용 `ko\QuickLaunchOnTray.resources.dll`)를 원하는 폴더에 둡니다.
2. 같은 폴더에 `config.ini`를 만듭니다. 형식은 `QuickLaunchOnTray/config.ini.sample`을 참고하세요.
3. `QuickLaunchOnTray.exe`를 실행하면 `config.ini`의 항목마다 트레이 아이콘이 등록됩니다.

### `config.ini` 규칙

```ini
; 주석은 ; 또는 #으로 시작합니다.
[Programs]
; 1) 이름=경로 : 이름이 트레이 아이콘 툴팁으로 표시됩니다.
메모장=C:\Windows\System32\notepad.exe
; 2) 경로만 : 파일명(확장자 제외) 또는 폴더명이 이름이 됩니다.
C:\Windows\System32\calc.exe
; 3) 폴더도 등록할 수 있습니다.
다운로드=C:\Users\Public\Downloads
```

| 규칙 | 동작 |
|---|---|
| 인코딩 | UTF-8 권장(다른 인코딩은 한글 경로가 깨질 수 있음) |
| 섹션 | `[Programs]` 섹션의 항목만 등록합니다(섹션 이름 대소문자 무시). 다른 섹션과 섹션 앞의 항목은 무시합니다. |
| 주석/빈 줄 | `;` 또는 `#`으로 시작하는 줄과 빈 줄은 무시합니다. |
| 이름 생략 | `경로만` 형식이거나 `=경로`처럼 이름이 비면 경로에서 이름을 만듭니다. |
| 이름 길이 | 툴팁은 최대 63자이며 초과분은 잘립니다(Windows 트레이 툴팁 상한). |
| 경로 없음/빈 경로 | 그 줄만 건너뛰고, 시작할 때 경고 대화상자에 줄 번호와 사유를 보여 줍니다. |
| 유효 항목 0개 | 오류 대화상자를 보여 주고 종료합니다(등록할 아이콘이 없으면 조작할 UI가 없으므로). |
| `config.ini` 없음 | 오류 대화상자를 보여 주고 종료합니다. |
| `[Programs]` 섹션 없음 | 오류 대화상자를 보여 주고 종료합니다. |
| 같은 경로 중복 | 중복 등록됩니다(아이콘 2개). |

### 사용법

| 조작 | 프로그램 항목 | 폴더 항목 |
|---|---|---|
| 좌클릭 | 실행 | 폴더 내용 메뉴 열기 |
| 우클릭 | 컨텍스트 메뉴(프로그램 실행 / 종료) | 컨텍스트 메뉴(폴더 열기 / 종료) |

- 폴더 내용 메뉴: 하위 폴더는 펼칠 때 읽어 오며(지연 로딩), 한 번 읽은 목록은 30초 동안 재사용합니다.
  하위 폴더 메뉴와 루트 메뉴 모두 `새로 고침`(캐시를 버리고 다시 읽기)과 `탐색기에서 열기`를 제공하고, 루트 메뉴에는 `닫기`가 있습니다.
- 파일 항목을 클릭하면 해당 파일을 기본 연결 프로그램으로 엽니다.
- 프로그램 실행에 실패하면 오류 대화상자를 보여 줍니다.
- 트레이 아이콘이 모두 사라지는 유일한 경로는 우클릭 메뉴의 `'QuickLaunchOnTray' 종료`(확인 후)입니다.
- 이미 실행 중이면 두 번째 인스턴스는 안내 메시지를 보여 주고 스스로 종료합니다(중복 아이콘 방지).

### 캐시(첫 클릭 지연 줄이기)

폴더 목록과 아이콘은 같은 내용을 다시 읽지 않도록 디스크에 캐시합니다.

| 항목 | 내용 |
|---|---|
| 위치 | `%LOCALAPPDATA%\QuickLaunchOnTray\cache\v1` (`folders\`, `icons\`) |
| 목록 캐시 | 폴더별 파일 1개. 폴더의 마지막 수정 시각(`LastWriteTimeUtc`)이 같으면 지난 실행에서 만든 목록을 그대로 사용합니다. |
| 아이콘 캐시 | 메뉴용 16×16 PNG. 실행 파일/바로가기는 경로+수정 시각+크기로 검증하고, 그 밖의 파일은 확장자 단위로 하루에 한 번 갱신합니다. |
| 예열 | 시작 2초 뒤 백그라운드에서 `config.ini`의 폴더 항목을 미리 읽습니다(로그온 직후 디스크 경합 완화). |
| 정리 | 오래된 항목(목록 90일/아이콘 30일)과 상한 초과분을 백그라운드에서 정리합니다. |
| 무효화 | 폴더 내용이 바뀌면 폴더 수정 시각이 달라져 자동으로 다시 읽습니다. 즉시 반영하려면 메뉴의 `새로 고침`을 사용합니다. |
| 초기화 | 캐시 폴더를 지우면 초기화됩니다(다음 실행에서 다시 만듭니다). 캐시를 못 쓰는 환경이면 캐시 없이 동작합니다. |

실측(등록 폴더를 처음 여는 시점): `C:\Windows\System32`(5,029항목) 1,390ms → 115ms, `C:\Windows\WinSxS`(28,206항목) 1,247ms → 34ms. 비용의 대부분은 목록 열거가 아니라 아이콘 추출이었습니다.

### 아이콘 출처

| 대상 | 사용하는 아이콘 |
|---|---|
| `.exe`/`.dll` 등 아이콘을 가진 파일 | **그 파일 자체의 아이콘**. 탐색기가 보여주는 아이콘과 16×16에서 픽셀 단위로 동일함을 실측했습니다. |
| `.lnk` 바로가기 | **대상 프로그램의 아이콘을 직접 참조**(바로가기 화살표가 없습니다). 바로가기에 아이콘이 명시되어 있으면 그 아이콘을, 대상이 폴더면 폴더 아이콘을 씁니다. 대상이 없거나 해석할 수 없는 바로가기만 셸 아이콘(화살표 포함)으로 폴백합니다. |
| 그 밖의 파일 | 확장자 연결 프로그램 아이콘(확장자당 1개) |
| 하위 폴더 | 폴더 아이콘. 메뉴와 트레이 아이콘은 DPI 배율에 맞는 트레이 크기(100%에서 16×16)로 그립니다. |

- 실행 파일과 바로가기는 **경로별로** 캐시하므로 프로그램마다 아이콘이 따로 유지됩니다. 실행 파일은 경로 + 수정 시각 + 크기가 바뀌면 즉시 다시 추출하고, 바로가기는 아이콘이 대상에서 오므로 **하루 단위**로 갱신합니다(대상 프로그램을 업데이트한 경우 최대 하루 뒤 반영).
- 바로가기 해석은 개당 약 0.16ms이며, 아이콘 추출 핸들은 사용 직후 해제합니다(바로가기 300회 반복 로드에서 GDI 핸들 증가 0).

### 로컬라이제이션

- 문자열 원본은 리소스 파일 하나이며, 언어별로 키를 중복 정의하지 않습니다.
  - `Resources/Strings.resx` : 중립(영어)
  - `Resources/Strings.<culture>.resx` : 예) `Strings.ko.resx` → 빌드 시 `ko\QuickLaunchOnTray.resources.dll` 위성 어셈블리 생성
- 조회 문화권은 프로세스 시작 시점의 `CurrentUICulture` 하나만 사용합니다. 지원하지 않는 문화권은 영어로 폴백하고, 키가 없으면 키 문자열을 그대로 반환합니다.
- 언어 추가: `Resources/Strings.ja.resx`처럼 파일을 만들고 `.csproj`의 `EmbeddedResource`에 추가하면 됩니다(코드 수정 불필요).

### 빌드

- Visual Studio 2017 이상 또는 MSBuild로 빌드합니다(`.NET Framework 4.8` 대상, C# WinForms).
- `Release` 구성 산출물: `QuickLaunchOnTray/bin/Release/QuickLaunchOnTray.exe`, 한국어 UI용 `.../ko/QuickLaunchOnTray.resources.dll`
- `config.ini`는 개인 경로 정보를 담으므로 `.gitignore` 대상입니다. 예시는 `QuickLaunchOnTray/config.ini.sample`을 사용하세요.

### 소스 구성

| 경로 | 역할 |
|---|---|
| `Program.cs` | 진입점. 단일 인스턴스 뮤텍스, 시작 오류 처리 |
| `TrayApplicationContext.cs` | 트레이 아이콘/메뉴, 폴더 메뉴 지연 로딩, 백그라운드 예열, 자원 정리 |
| `Services/ConfigurationService.cs` | `config.ini` 파싱, 항목 검증(건너뛴 줄은 경고 수집) |
| `Services/FolderIndexCache.cs` | 폴더 목록 캐시(메모리+디스크, 수정 시각 검증, 낡은 상태 재확인) |
| `Services/IconImageCache.cs` | 아이콘 이미지 캐시(메모리+디스크 PNG, 자체 무효화) |
| `Services/CacheStorage.cs` | 캐시 위치, 원자적 쓰기, 오래된 항목 정리 |
| `Services/IconLoader.cs` | 트레이 크기 아이콘 추출(DPI 배율 반영) |
| `Services/LocalizationService.cs` | 리소스 문자열 조회 |
| `Models/` | 등록 항목(`ProgramItem`), 폴더 항목(`FolderEntry`), 파싱 결과(`ConfigurationLoadResult`) |
| `Resources/` | 문자열 리소스(`Strings.resx`, `Strings.ko.resx`) |

### 라이선스
이 프로젝트는 MIT 라이선스 하에 배포됩니다.

---

## English Description

### Introduction
QuickLaunchOnTray registers icons for user-specified programs and folders in the Windows notification area,
so they can be launched with a single click.

- It reads only `config.ini`, located next to the executable.
- There is no limit on the number of registered icons (one icon per entry).
- Invalid entries never block the rest: they are skipped, and the application reports them once at startup.

### Requirements

| Item | Value |
|---|---|
| OS | Windows 7 SP1 or later (Windows 10/11 recommended) |
| Runtime | .NET Framework 4.8 |
| DPI | Per-monitor V2 aware (declared in the manifest); tray icons use the DPI-scaled small icon size. |

### Installation

1. Put `QuickLaunchOnTray.exe` (plus `ko\QuickLaunchOnTray.resources.dll` for the Korean UI) in a folder.
2. Create `config.ini` in the same folder. See `QuickLaunchOnTray/config.ini.sample` for the format.
3. Run `QuickLaunchOnTray.exe`; one tray icon is registered per entry.

### `config.ini` rules

```ini
; comments start with ; or #
[Programs]
; 1) name=path : the name is used as the tray icon tooltip
Notepad=C:\Windows\System32\notepad.exe
; 2) path only : the file name (without extension) or folder name becomes the name
C:\Windows\System32\calc.exe
; 3) folders can be registered as well
Downloads=C:\Users\Public\Downloads
```

| Rule | Behavior |
|---|---|
| Encoding | UTF-8 recommended (other encodings may corrupt non-ASCII paths) |
| Section | Only entries under `[Programs]` are registered (name match is case-insensitive). Other sections and entries before it are ignored. |
| Comments/blank lines | Lines starting with `;` or `#`, and blank lines, are ignored. |
| Missing name | For the path-only form, or when the name is empty (`=path`), the name is derived from the path. |
| Name length | Tooltips are limited to 63 characters; longer names are truncated (Windows tray tooltip limit). |
| Empty/missing path | The line is skipped and reported with its line number in a startup warning dialog. |
| No valid entry | An error dialog is shown and the application exits (with no icons there would be no UI to operate). |
| Missing `config.ini` | An error dialog is shown and the application exits. |
| Missing `[Programs]` section | An error dialog is shown and the application exits. |
| Duplicate path | Registered twice (two icons). |

### Usage

| Action | Program entry | Folder entry |
|---|---|---|
| Left click | Run the program | Open the folder contents menu |
| Right click | Context menu (run program / terminate) | Context menu (open folder / terminate) |

- Folder contents menu: subfolders are read lazily when expanded and the listing is cached for 30 seconds.
  Both subfolder menus and the root menu offer `Refresh` (drop the cache and read again) and `Open in Explorer`; the root menu also has `Close`.
- Clicking a file entry opens it with its default association.
- A failed launch shows an error dialog.
- The only way to remove all tray icons is `Terminate 'QuickLaunchOnTray'` from the right-click menu (after confirmation).
- If an instance is already running, a second instance shows a notice and exits by itself (no duplicate icons).

### Cache (reducing first-click latency)

Folder listings and icons are cached on disk so the same content is not read twice.

| Item | Detail |
|---|---|
| Location | `%LOCALAPPDATA%\QuickLaunchOnTray\cache\v1` (`folders\`, `icons\`) |
| Listing cache | One file per folder. If the folder's `LastWriteTimeUtc` is unchanged, the listing from a previous run is reused. |
| Icon cache | 16×16 PNGs for menu images. Executables/shortcuts are validated by path + write time + size; other files are refreshed per extension once a day. |
| Warm-up | Two seconds after startup, configured folder entries are read in the background (avoids competing with other logon activity). |
| Cleanup | Stale entries (listings 90 days / icons 30 days) and overflow beyond the caps are pruned in the background. |
| Invalidation | Changed folder contents change the folder timestamp, so the listing is read again automatically. Use `Refresh` in the menu to force it immediately. |
| Reset | Delete the cache folder to start over; the app works without the cache if it cannot be written. |

Measured (first open of a registered folder): `C:\Windows\System32` (5,029 entries) 1,390 ms → 115 ms, `C:\Windows\WinSxS` (28,206 entries) 1,247 ms → 34 ms. Most of the cost was icon extraction, not listing.

### Icon sources

| Target | Icon used |
|---|---|
| `.exe`/`.dll` and other files that carry icons | **The file's own icon.** Measured pixel-identical at 16×16 to the icon Explorer shows. |
| `.lnk` shortcuts | **The target program's icon, referenced directly** (no shortcut arrow overlay). An icon set explicitly on the shortcut wins, and a shortcut to a folder uses the folder icon; only shortcuts whose target is missing or unresolvable fall back to the shell icon (with overlay). |
| Other files | The associated program's icon (one per extension) |
| Subfolders | Folder icon. Both menus and tray icons are drawn at the DPI-scaled small icon size (16×16 at 100%). |

- Executables and shortcuts are cached **per file path**, so icons stay distinct per program. Executables are re-extracted as soon as path + write time + size change; shortcuts take their icon from the target, so they refresh **once a day** (a target program update appears within a day).
- Resolving a shortcut costs about 0.16 ms each, and icon handles are released right after use (300 repeated shortcut loads increased GDI handles by 0).

### Localization

- Strings live in resource files with a single key set per language:
  - `Resources/Strings.resx` : neutral (English)
  - `Resources/Strings.<culture>.resx` : e.g. `Strings.ko.resx` produces the `ko\QuickLaunchOnTray.resources.dll` satellite assembly
- A single culture source is used: `CurrentUICulture` at process start. Unsupported cultures fall back to English, and a missing key returns the key itself.
- To add a language, add e.g. `Resources/Strings.ja.resx` and register it as an `EmbeddedResource` in `.csproj` (no code change needed).

### Build and run

- Build with Visual Studio 2017+ or MSBuild (C# WinForms targeting .NET Framework 4.8).
- `Release` output: `QuickLaunchOnTray/bin/Release/QuickLaunchOnTray.exe` and the Korean satellite `.../ko/QuickLaunchOnTray.resources.dll`.
- `config.ini` holds machine-specific paths and is git-ignored; start from `QuickLaunchOnTray/config.ini.sample`.

### License
This project is licensed under the MIT License.
