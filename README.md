# WinSys Guardian Advanced

WinSys Guardian Advanced (WGA) е настолно административно приложение за Windows, разработено с Python и `tkinter`. То обединява в един графичен интерфейс инструменти за подготовка, поддръжка, диагностика, инсталация на софтуер, работа с Microsoft Office, езикови настройки, архивиране на драйвери, локални потребители и обновяване на самото приложение.

Проектът е създаден като практичен помощник за техник, системен администратор или малък екип, който често подготвя нови Windows компютри, обслужва преинсталирани машини или изпълнява повтарящи се административни задачи.

GitHub repository:

```text
https://github.com/gyx4ee/WGA
```

## Текуща версия

Текущата версия е описана във `version.json`.

```text
0.2.47
```

Версия `0.2.47` добавя обратим профил за оптимизация на Windows 10 с 2/4 GB RAM, локално включен Open-Shell 4.4.198 и Windows XP Luna оформление на Start менюто без отделно изтегляне.

## Основна идея

При стандартна подготовка на Windows компютър обикновено се изпълняват много отделни действия:

- инсталиране на основен софтуер;
- подготовка или поправка на Microsoft Office;
- проверка на системното състояние;
- настройка на български език и клавиатурни подредби;
- архивиране на драйвери преди преинсталация;
- управление на локални потребители;
- активационни операции;
- проверка за нова версия на инструмента.

WGA събира тези действия в контролен dashboard и отделни менюта. Целта не е да замести enterprise платформи като Intune или Configuration Manager, а да даде лек, portable и локален инструмент за конкретната машина пред техника.

## Стартиране

В режим на разработка:

```bash
python app.py
```

При компилиран build входната точка е:

```text
WGA.exe
```

За част от функциите са нужни администраторски права, защото приложението използва Windows инструменти като `PowerShell`, `slmgr.vbs`, `pnputil`, `DISM`, `net.exe`, `winget`, Office Deployment Tool и Office Click-to-Run компоненти.

## Основни технологии

- Python
- `tkinter` за desktop интерфейс
- PowerShell за системни проверки и езикови настройки
- `pnputil` и `DISM` за драйвери
- `net.exe` и PowerShell за локални потребители
- Microsoft Office Deployment Tool за online Office сценарии
- `winget` за част от проверките и maintenance операциите
- JSON файлове за настройки, версии и ресурсен manifest
- PyInstaller за portable build
- Inno Setup за USB installer пакет

## Главен dashboard

След стартиране приложението показва loading екран с реални startup проверки. Зареждат се или се проверяват:

- конфигурационни файлове;
- системна информация;
- инсталационни ресурси;
- езиков статус;
- наличен софтуер;
- update статус;
- вътрешни модули и визуални ресурси.

Основният екран е dashboard с тъмна визуална тема, странична навигация, статус banner, системни панели, component status, automatic installer preview и бърз достъп до основните модули.

## Основни менюта

Главната навигация включва:

- Dashboard
- Activation
- Software
- Language
- Auto Installer
- Archive
- Nexus Admin
- Reset Console
- Exit

Вътрешно приложението съдържа и допълнителни подменюта за Windows activation, Office activation, Office install center, Office online deployment, Secret Install Interface, OneDrive reset и driver backup действия.

## Activation

Activation менюто обединява Windows и Office активационни сценарии.

### Windows 10 Key Manager

- Save or Replace Product Key
- Show Saved Product Key
- Clear Saved Product Key
- Run Windows 10 Activation

### Windows 11 Key Manager

- Save or Replace Product Key
- Show Saved Product Key
- Clear Saved Product Key
- Run Windows 11 Activation

### Office Activation Center

- Save or Replace Office Key
- Show Saved Office Key
- Clear Saved Office Key
- Office 2016 activation workflow
- Office 2019 activation workflow
- Office 2021 activation workflow

Записаните ключове се пазят локално в скрит secure store файл:

```text
.wga_secure_store.json
```

Този файл не трябва да се качва в GitHub. Той се създава локално до приложението и съдържа кодиран payload с чувствителните стойности.

## Software

Software секцията събира инсталационни и Office функции.

### Install Software

Менюто включва:

- Office Install Center
- Office Online God Mode
- Install Ninite
- Install Adobe Reader
- Secret Install Interface

### Office Install Center

Поддържани offline Office сценарии:

- Office 2016 Offline
- Office 2019 Offline
- Office 2021 Offline
- Office Professional 2021 Offline
- Office Professional 2024 Offline
- Office Standard 2024 Offline
- Office Standard 2021 Offline

Offline инсталациите очакват локални файлове в `Installers`. За всяка версия се проверяват `setup.exe` и съответният `Configuration.xml`.

### Office Online God Mode

Online Office менюто използва Microsoft Office Deployment Tool и поддържа сценарии за:

- Microsoft 365
- Office Professional Plus 2024
- Office Home & Business 2024
- Office Professional Plus 2021
- Office Home & Student 2021
- Office Professional Plus 2019
- Visio Professional
- Project Professional

Допълнителни Office maintenance функции:

- Check Activation Status
- Quick Repair Office
- Force Uninstall All Office Versions

## Secret Install Interface

Secret Install Interface групира инсталационни категории:

- System Runtimes
- Browsers & Comms
- Development
- Languages & DB
- Cybersecurity & Net
- Virtualization
- Multimedia & Design
- Gaming & Tools
- Utilities & Office
- Advanced Tools
- Update All Apps

Идеята е техникът да избира софтуер по работен сценарий, а не да търси отделни installer файлове ръчно.

## Automatic Installer

Automatic Installer е отделен екран за последователна подготовка на компютър. Той позволява избор на няколко програми или компоненти, проверява дали част от тях вече са налични и показва progress прозорци по време на изпълнение.

Функцията е полезна при:

- подготовка на нов компютър;
- работа след преинсталация;
- повтарящи се сервизни задачи;
- еднаква подготовка на няколко машини.

## Resource Manager

`resource_manager.py` и `installers_manifest.json` управляват локални и online ресурси.

Manifest файлът описва:

- идентификатор на ресурса;
- категория;
- очаквани файлове;
- URL за изтегляне;
- име на целевия архив;
- дали архивът се разархивира;
- размер;
- SHA256 checksum.

Resource Manager може да:

- провери наличните файлове в `Installers`;
- отчете липсващи ресурси;
- изтегли online пакети;
- покаже прогрес;
- провери SHA256;
- разархивира ZIP пакети.

## Language Manager

Language Manager управлява български език и клавиатурни подредби в Windows 11.

Функции:

- Refresh Language Status
- Toggle Bulgarian BDS
- Toggle Bulgarian Phonetic
- Toggle Bulgarian Traditional Phonetic
- Toggle Bulgarian Language Pack
- Remove Bulgarian Language

Модулът използва PowerShell команди като `Get-WinUserLanguageList`, `Set-WinUserLanguageList`, `Get-WindowsCapability`, `Add-WindowsCapability` и `Remove-WindowsCapability`.

## Driver Backup God Mode

Driver Backup менюто е предназначено за архивиране, възстановяване и хардуерен отчет.

Функции:

- Backup Drivers (Clean)
- Backup Drivers (Full)
- Create Recovery USB + RESTORE.bat
- Generate PC Report
- Driver Backup Tool v0.1
- Restore Drivers From Last Backup

При backup се създават:

- папка `DriversBackup_YYYY-MM-DD_HH-MM`;
- `backup_log.txt`;
- `drivers_list.txt`;
- `RESTORE_DRIVERS.bat`;
- ZIP архив при избрана такава опция.

Хардуерният отчет включва информация за CPU, RAM, GPU, BIOS, motherboard, дискове, операционна система и мрежови адаптери.

## Nexus Admin

Nexus Admin е модул за локална потребителска администрация.

Функции:

- List All Users
- Change Password
- Create New User
- Delete User
- User Details
- Toggle Administrator Rights

Модулът използва `net.exe` и PowerShell. За част от действията са нужни администраторски права.

## Reset OneDrive

Reset OneDrive предлага няколко начина за възстановяване на OneDrive клиента:

- стандартен reset;
- спиране и повторно стартиране на процеса;
- по-дълбоко почистване на локални OneDrive файлове в потребителския профил.

Третият метод е по-краен и трябва да се използва внимателно.

## Add Desktop Icons

Функцията създава support shortcuts и може да активира полезни Windows desktop икони като:

- This PC
- Network
- Control Panel
- User Files

## System Health

`system_health.py` събира информация за текущата машина:

- Windows версия и build;
- име на компютър;
- активен потребител;
- CPU;
- RAM;
- GPU;
- BIOS;
- motherboard;
- дискове и свободно място;
- локален IP адрес;
- uptime;
- battery status при лаптопи;
- наличие на важни системни компоненти.

Dashboard екранът използва тези данни, за да даде бърза първоначална ориентация.

## Update система

WGA има вграден update механизъм чрез `version.json`.

Локалният `version.json` съдържа:

- текуща версия;
- URL към online metadata;
- URL към latest release;
- URL към portable ZIP package;
- notes;
- changelog.

Текущите update адреси са:

```text
version_info_url: https://raw.githubusercontent.com/gyx4ee/WGA/refs/heads/main/version.json
download_url:     https://github.com/gyx4ee/WGA/releases/latest
package_url:      https://raw.githubusercontent.com/gyx4ee/WGA/refs/heads/main/WGA-portable.zip
```

Update процесът може да:

- провери online metadata;
- сравни локалната и remote версията;
- покаже notes и changelog;
- изтегли portable update package;
- разархивира файловете във временна папка;
- създаде helper `.cmd`;
- копира новите файлове чрез `robocopy`;
- запази `settings.json`, `.wga_secure_store.json` и други локални данни;
- рестартира приложението.

## Portable и USB режим

Проектът е подготвен за portable употреба. Приложението може да се стартира от USB носител заедно с `Installers` папката и готовите ресурси.

Portable логиката помага за:

- намиране на `Installers` спрямо мястото на стартиране;
- работа без отделна Python среда;
- пренасяне на инструмента между различни машини;
- update без класическа инсталационна платформа.

## Build и release

Основни build файлове:

- `WGA.spec` - PyInstaller конфигурация;
- `build_release.ps1` - build helper;
- `WGAInstaller.iss` - Inno Setup конфигурация;
- `version.json` - release metadata;
- `WGA-portable.zip` - portable package.

Типичният release процес включва:

1. обновяване на версията във `version.json`;
2. build чрез PyInstaller;
3. пакетиране на `dist/WGA`;
4. създаване или обновяване на `WGA-portable.zip`;
5. качване на новите файлове в GitHub;
6. commit и push към `main`.

## Основни файлове

- `app.py` - основно приложение, UI, менюта и workflow логика;
- `system_health.py` - системна диагностика;
- `resource_manager.py` - проверка и изтегляне на ресурси;
- `self_updater.py` - подготовка и стартиране на update helper;
- `update_checker.py` - online проверка за нова версия;
- `driver_backup.py` - driver backup, restore и hardware report;
- `language_manager.py` - език и клавиатурни подредби;
- `nexus_admin.py` - локални потребители и administrator права;
- `office_activation.py` - Office activation команди;
- `office_installers.py` - offline Office конфигурации;
- `office_inventory.py` - проверка за инсталиран Office;
- `office_maintenance.py` - Office repair, status и cleanup;
- `office_online.py` - online Office продукти през ODT;
- `adobe_reader.py` - Adobe Reader проверка и инсталация;
- `path_utils.py` - portable paths и runtime storage;
- `installers_manifest.json` - online resource manifest;
- `version.json` - версия, update адреси и changelog.

## Папки

- `assets` - икони, dashboard ресурси и визуални файлове;
- `Installers` - локални installer ресурси, ако се използват;
- `dist` - PyInstaller output;
- `build` - build cache;
- `installer-output` - release архиви;
- `Backups` / `DriversBackup_*` - driver backup резултати при работа на приложението.

В GitHub обикновено се качват кодът, assets, release metadata и готовият portable пакет. Локални чувствителни файлове като `.wga_secure_store.json` не трябва да се качват.

## Практическо предназначение

WGA е полезно при:

- подготовка на нов Windows компютър;
- работа след преинсталация;
- инсталиране на базов софтуер;
- подготовка и поддръжка на Microsoft Office;
- архивиране на драйвери преди рискови промени;
- създаване на recovery USB за драйвери;
- проверка на системна и хардуерна информация;
- управление на локални потребители;
- настройка на български език и клавиатура;
- работа от USB носител;
- бърза локална поддръжка в малки организации, сервизи и учебни среди.

## Важно

Част от функциите изпълняват системни команди и могат да променят Windows настройки. Приложението трябва да се използва внимателно, особено при:

- изтриване на потребители;
- промяна на administrator права;
- force uninstall на Office;
- промяна на езикови пакети;
- активационни операции;
- възстановяване на драйвери;
- update на portable build.

WGA е помощен административен инструмент. Потребителят трябва да разбира действието, което стартира.
