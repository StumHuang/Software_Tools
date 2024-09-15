@ECHO off

::set ASAP2 Tool-Set Installation Directory
set ASAP2ToolInstallDir=C:\Program Files\Vector\ASAP2 Tool-Set 15.0\Bin
set ASAP2ToolInstallDirString="%ASAP2ToolInstallDir%"
set MergerExe=ASAP2Merger.exe
set UpdaterExe=ASAP2Updater.exe

::check Vector merger installation
if not exist %ASAP2ToolInstallDirString%\%MergerExe% (
@ECHO.
@ECHO !!!!!!!!!!ERROR!!!!!!!!!!
@ECHO.
@ECHO %MergerExe% not found at %ASAP2ToolInstallDirString%\! Please check for valid Installation of ASAP2 Tool-Set!
@ECHO.
@pause
exit
)

::check Vector updater installation
if not exist %ASAP2ToolInstallDirString%\%UpdaterExe% (
@ECHO.
@ECHO !!!!!!!!!!ERROR!!!!!!!!!!
@ECHO.
@ECHO %UpdaterExe% not found at %ASAP2ToolInstallDirString%\! Please check for valid Installation of ASAP2 Tool-Set!
@ECHO.
@pause
exit
)

::set environment variables for directories
set MasterDir=01_Master
set SlaveDir=02_Slaves
set MergedDir=03_Merged

::check if master A2L file exists
if not exist %MasterDir%\*.a2l (
@ECHO.
@ECHO !!!!!!!!!!ERROR!!!!!!!!!!
@ECHO.
@ECHO No A2L file found at .\%MasterDir%\! Please insert supplier A2L file!
@ECHO.
@pause
exit
)

::check that max one master A2L file exists
set MasterA2LCount=0
for %%x in (%MasterDir%\*.a2l) do set /a MasterA2LCount+=1
if %MasterA2LCount% GTR 1 (
@ECHO.
@ECHO !!!!!!!!!!ERROR!!!!!!!!!!
@ECHO.
@ECHO Too many A2L files found at .\%MasterDir%\! Please insert only one supplier A2L file!
@ECHO.
@pause
exit
)

::check if master out file exists
if not exist %MasterDir%\*.out (
@ECHO.
@ECHO !!!!!!!!!!ERROR!!!!!!!!!!
@ECHO.
@ECHO No out file found at .\%MasterDir%\! Please insert supplier out file!
@ECHO.
@pause
exit
)

::check that max one master out file exists
set MasterMapCount=0
for %%x in (%MasterDir%\*.out) do set /a MasterMapCount+=1
if %MasterMapCount% GTR 1 (
@ECHO.
@ECHO !!!!!!!!!!ERROR!!!!!!!!!!
@ECHO.
@ECHO Too many out files found at .\%MasterDir%\! Please insert only one supplier out file!
@ECHO.
@pause
exit
)

::check if at least one slave A2L file exists
if not exist %SlaveDir%\*.a2l (
@ECHO.
@ECHO !!!!!!!!!!ERROR!!!!!!!!!!
@ECHO.
@ECHO No A2L file found at .\%SlaveDir%\! Please insert at least one FIH A2L file!
@ECHO.
@pause
exit
)


::print input files
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


::clean output folder
erase /Q .\%MergedDir%\*

::initialize output folder with renamed supplier a2l file
copy .\%MasterDir%\*.a2l .\%MergedDir%\
cd %MergedDir%
for /f "delims=." %%i in ('dir /b *.a2l') do @rename "%%i.a2l" "%%i_FIH_merged.a2l"
cd ..

::get renamed output a2l file name
for %%a in (.\%MergedDir%\*) do set FihMerged=%%~nxa

::merge a2l files
for /r .\%SlaveDir%\ %%a in (*.a2l) do ("%ASAP2ToolInstallDir%\%MergerExe%" -M .\%MergedDir%\%FihMerged% -S %%a -O .\%MergedDir%\%FihMerged%)

::get out file name
for %%a in (.\%MasterDir%\*.out) do set MapFile=%%~nxa

::update a2l file
"%ASAP2ToolInstallDir%\%UpdaterExe%" -I .\%MergedDir%\%FihMerged% -O .\%MergedDir%\%FihMerged% -A .\%MasterDir%\%MapFile% -L .\%MergedDir%\log1.txt -T Updater.ini

:: convert to UTF-8
sed -i "/ASAP2_VERSION/,$!d" .\%MergedDir%\%FihMerged%

::end

::vst build
@ECHO.
@ECHO Build Vst File:
@ECHO.
@ECHO ELF H32 File:
for %%a in (.\%MasterDir%\*.h32) do set H32File=%%~nxa
@ECHO %H32File%

@ECHO.
@ECHO ELF VST Tool File:
for /r %%f in (..\"VST_Tool"\*.xlsm) do set VST_Tool=%%~nxf
@ECHO %VST_Tool%

@ECHO.
@ECHO VBS File:
for /r %%f in (..\"VST_Tool"\*.vbs) do set VST_VBS=%%~nxf
@ECHO %VST_VBS%

@ECHO.
for %%a in (.\%MergedDir%\*.a2l) do set A2LFile=%%~nxa
cd %MergedDir%
for /f "delims=." %%i in ('dir /b *.a2l') do set VSTFile=%%i.vst
cd ..

cscript "%~dp0\..\VST_Tool\%VST_VBS%" "%~dp0\..\VST_Tool\%VST_Tool%" "%~dp0\.\%MergedDir%\%A2LFile%" "%~dp0\.\%MasterDir%\%H32File%" "%~dp0\.\%MergedDir%\%VSTFile%"

::end
@ECHO.
@ECHO.
@ECHO %~dp0\%VSTFile% created.
@ECHO.
@pause
