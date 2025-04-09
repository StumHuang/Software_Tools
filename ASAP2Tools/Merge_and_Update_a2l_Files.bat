@ECHO off

:: Set directories and paths
set MasterDir=01_Master
set SlaveDir=02_Slaves
set MergedDir=03_Merged
set VstPath=C:\Users\shuang\Documents\001_Git_Repository\Software_Tools\ASAP2Tools

:: Set ASAP2 Tool-Set Installation Directory
set ASAP2ToolInstallDir=C:\Program Files\Vector\ASAP2 Tool-Set 15.0\Bin
set ASAP2ToolInstallDirString="%ASAP2ToolInstallDir%"
set MergerExe=ASAP2Merger.exe
set UpdaterExe=ASAP2Updater.exe

:: Validate ASAP2 Tool-Set installation
if not exist %ASAP2ToolInstallDirString%\%MergerExe% (
    @ECHO !!!!!!!!!!ERROR!!!!!!!!!!
    @ECHO %MergerExe% not found at %ASAP2ToolInstallDirString%\! Please check for valid Installation of ASAP2 Tool-Set!
    @pause
    exit
)
if not exist %ASAP2ToolInstallDirString%\%UpdaterExe% (
    @ECHO !!!!!!!!!!ERROR!!!!!!!!!!
    @ECHO %UpdaterExe% not found at %ASAP2ToolInstallDirString%\! Please check for valid Installation of ASAP2 Tool-Set!
    @pause
    exit
)

:: Validate directories and input files
if not exist "%VstPath%" (
    @ECHO !!!!!!!!!!ERROR!!!!!!!!!!
    @ECHO VstPath directory "%VstPath%" does not exist! Please create it or update the script.
    @pause
    exit
)
if not exist %MasterDir%\*.a2l (
    @ECHO !!!!!!!!!!ERROR!!!!!!!!!!
    @ECHO No A2L file found at .\%MasterDir%\! Please insert supplier A2L file!
    @pause
    exit
)
set MasterA2LCount=0
for %%x in (%MasterDir%\*.a2l) do set /a MasterA2LCount+=1
if %MasterA2LCount% GTR 1 (
    @ECHO !!!!!!!!!!ERROR!!!!!!!!!!
    @ECHO Too many A2L files found at .\%MasterDir%\! Please insert only one supplier A2L file!
    @pause
    exit
)
if not exist %MasterDir%\*.out (
    @ECHO !!!!!!!!!!ERROR!!!!!!!!!!
    @ECHO No out file found at .\%MasterDir%\! Please insert supplier out file!
    @pause
    exit
)
set MasterMapCount=0
for %%x in (%MasterDir%\*.out) do set /a MasterMapCount+=1
if %MasterMapCount% GTR 1 (
    @ECHO !!!!!!!!!!ERROR!!!!!!!!!!
    @ECHO Too many out files found at .\%MasterDir%\! Please insert only one supplier out file!
    @pause
    exit
)
if not exist %SlaveDir%\*.a2l (
    @ECHO !!!!!!!!!!ERROR!!!!!!!!!!
    @ECHO No A2L file found at .\%SlaveDir%\! Please insert at least one FIH A2L file!
    @pause
    exit
)

:: Print input files
@ECHO.
@ECHO Supplier A2L file:
for /r %%f in (%MasterDir%\*.a2l) do @ECHO %%~nxf
@ECHO.
@ECHO Supplier out file:
for /r %%f in (%MasterDir%\*.out) do @ECHO %%~nxf
@ECHO.
@ECHO FIH A2L file(s):
for /r %%f in (%SlaveDir%\*.a2l) do @ECHO %%~nxf
@ECHO.

:: Prepare output folder
erase /Q .\%MergedDir%\*
copy .\%MasterDir%\*.a2l .\%MergedDir%\
cd %MergedDir%
for /f "delims=." %%i in ('dir /b *.a2l') do @rename "%%i.a2l" "%%i_merged.a2l"
cd ..

:: Validate Merger.ini file
if not exist "%~dp0Merger.ini" (
    @ECHO !!!!!!!!!!ERROR!!!!!!!!!!
    @ECHO Merger.ini file not found in the script directory! Please ensure it exists.
    @pause
    exit
)

:: Merge A2L files
for %%a in (.\%MergedDir%\*) do set FihMerged=%%~nxa
for /r .\%SlaveDir%\ %%a in (*.a2l) do ("%ASAP2ToolInstallDir%\%MergerExe%" -M .\%MergedDir%\%FihMerged% -S %%a -O .\%MergedDir%\%FihMerged% -P "%~dp0Merger.ini")

:: Update A2L file
for %%a in (.\%MasterDir%\*.out) do set MapFile=%%~nxa
if not exist "%~dp0Updater.ini" (
    @ECHO !!!!!!!!!!ERROR!!!!!!!!!!
    @ECHO Updater.ini file not found in the script directory! Please ensure it exists.
    @pause
    exit
)
"%ASAP2ToolInstallDir%\%UpdaterExe%" -I .\%MergedDir%\%FihMerged% -O .\%MergedDir%\%FihMerged% -A .\%MasterDir%\%MapFile% -L .\%MergedDir%\log1.txt -T "%~dp0Updater.ini"

:: Convert to UTF-8
if exist .\%MergedDir%\%FihMerged% (
    sed -i "/ASAP2_VERSION/,$!d" .\%MergedDir%\%FihMerged%
) else (
    @ECHO !!!!!!!!!!ERROR!!!!!!!!!!
    @ECHO Merged A2L file not found! Conversion to UTF-8 failed.
    @pause
    exit
)

:: Build VST file
@ECHO.
@ECHO Build VST File:
@ECHO.
for %%a in (.\%MasterDir%\*.h32) do set H32File=%%~nxa
@ECHO ELF H32 File: %H32File%
for /r %%f in (..\"VST_Tool"\*.xlsm) do set VST_Tool=%%~nxf
@ECHO ELF VST Tool File: %VST_Tool%
for /r %%f in (..\"VST_Tool"\*.vbs) do set VST_VBS=%%~nxf
@ECHO VBS File: %VST_VBS%
for %%a in (.\%MergedDir%\*.a2l) do set A2LFile=%%~nxa
@ECHO A2L File: %A2LFile%
cd %MergedDir%
for /f "delims=." %%i in ('dir /b *.a2l') do set VSTFile=%VstPath%\%%i.vst
@ECHO VST File: %VSTFile%
cd ..
cscript "%~dp0\..\VST_Tool\%VST_VBS%" "%~dp0\..\VST_Tool\%VST_Tool%" "%~dp0\.\%MergedDir%\%A2LFile%" "%~dp0\.\%MasterDir%\%H32File%" "%VSTFile%"

:: End
@ECHO.
@ECHO %VSTFile% created.
@pause
