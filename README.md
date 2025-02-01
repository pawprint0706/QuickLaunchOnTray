# QuickLaunchOnTray

### 소개  
QuickLaunchOnTray는 Windows 시스템 트레이에 사용자가 지정한 프로그램들의 아이콘을 등록하여 손쉽게 실행할 수 있도록 도와주는 유틸리티입니다. 이 프로그램은 실행 파일과 동일한 폴더에 위치한 `config.ini` 파일의 `[Programs]` 섹션에 기재된 정보를 읽어, 각 프로그램에 대한 트레이 아이콘을 생성합니다. 아이콘은 해당 프로그램 실행 파일에서 추출되며(추출에 실패하면 기본 아이콘 사용), 트레이 아이콘을 더블 클릭하거나 우클릭 메뉴의 **"Run this program"** 항목을 선택하면 지정된 프로그램이 실행됩니다. 또한, 우클릭 메뉴의 **"Terminate 'QuickLaunchOnTray'"** 항목을 통해 확인 메시지 후 프로그램을 안전하게 종료할 수 있습니다.

### 사용법

#### 1. `config.ini` 파일 설정  
- 실행 파일과 동일한 폴더에 `config.ini` 파일을 생성합니다.  
- `config.ini` 파일 내 `[Programs]` 섹션에 프로그램 정보를 추가합니다. 두 가지 형식을 지원합니다:

  - **Key=Value 형식 (툴팁에 표시될 이름 지정):**  
    ```ini
    [Programs]
    MyApp=C:\Path\To\MyApp.exe
    ```  
    여기서 `MyApp`은 트레이 아이콘의 툴팁에 표시될 이름입니다.

  - **단순 경로 형식 (파일명(확장자 제외)을 이름으로 사용):**  
    ```ini
    [Programs]
    C:\Path\To\MyApp.exe
    ```

#### 2. QuickLaunchOnTray 실행  
- 프로그램 실행 시, `config.ini` 파일에 지정된 각 프로그램에 대해 시스템 트레이에 아이콘이 등록됩니다.
- **프로그램 실행:**  
  - 트레이 아이콘을 더블 클릭하거나 우클릭 메뉴의 **"Run this program"** 항목을 선택하면 해당 프로그램이 실행됩니다.
- **프로그램 종료:**  
  - 우클릭 메뉴의 **"Terminate 'QuickLaunchOnTray'"** 항목을 선택하면 확인 메시지가 나타나고, 'Yes'를 선택할 경우 프로그램이 종료됩니다.

#### 3. 빌드 및 실행  
- Visual Studio를 사용하여 C# WinForms 프로젝트로 빌드합니다.  
- 빌드된 실행 파일과 동일한 폴더에 `config.ini` 파일을 위치시킵니다.  
- 실행 파일을 실행하면, 지정된 프로그램들의 트레이 아이콘이 생성됩니다.

### 라이선스  
이 프로젝트는 MIT 라이선스 하에 배포됩니다.

## English Description

### Introduction  
QuickLaunchOnTray is a utility that registers icons for user-specified programs in the Windows system tray, making it easy to launch them. The application reads the information from the `[Programs]` section in a `config.ini` file located in the same directory as the executable, and creates a tray icon for each program. The icon is extracted from the program’s executable (or a default icon is used if extraction fails). You can launch a program by double-clicking its tray icon or by selecting the **"Run this program"** option from the context menu. Additionally, you can safely terminate the application using the **"Terminate 'QuickLaunchOnTray'"** option, which prompts you for confirmation before exiting.

### Usage

#### 1. Setting up the `config.ini` File  
- Create a `config.ini` file in the same directory as the executable.  
- Add the program information under the `[Programs]` section in the `config.ini` file. Two formats are supported:

  - **Key=Value Format (specify a name for the tooltip):**  
    ```ini
    [Programs]
    MyApp=C:\Path\To\MyApp.exe
    ```  
    Here, `MyApp` will be used as the tooltip text for the tray icon.

  - **Simple Path Format (uses the file name without extension as the name):**  
    ```ini
    [Programs]
    C:\Path\To\MyApp.exe
    ```

#### 2. Running QuickLaunchOnTray  
- When the application runs, it registers a tray icon for each program specified in the `config.ini` file.  
- **Running a Program:**  
  - Double-click a tray icon or select the **"Run this program"** option from the context menu to launch the associated program.
- **Terminating the Application:**  
  - Right-click a tray icon and select **"Terminate 'QuickLaunchOnTray'"**. A confirmation prompt will appear, and selecting 'Yes' will terminate the application.

#### 3. Build and Run  
- Build the project as a C# WinForms application using Visual Studio.  
- Place the `config.ini` file in the same directory as the built executable.  
- Run the executable to have tray icons created for the specified programs.

### License  
This project is licensed under the MIT License.
