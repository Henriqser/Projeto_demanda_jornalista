@echo off
chcp 65001 > nul
cls
title Automação para Jornalistas - Menu Principal
color 0A

:: ==============================================
:: CONFIGURAÇÕES GLOBAIS
:: ==============================================
set "SCRIPT_DIR=%~dp0"
set "PYTHON_SCRIPT=script.py"
set "BIBLIOTECAS=psycopg2 pandas pyarrow python-dotenv"
set "PROXY_SERVER=10.10.111.11:111"
set "USUARIO=%USERNAME%"

:: ==============================================
:: MENU PRINCIPAL
:: ==============================================
:menu_principal
cls
echo.
echo  ============================================
echo    AUTOMAÇÃO PARA DEMANDAS DE JORNALISTAS
echo  ============================================
echo.
echo  1. Instalar dependências (Python e bibliotecas)
echo  2. Executar programa principal
echo  3. Sair
echo.
choice /c 123 /n /m "Selecione uma opção: "

if errorlevel 3 goto :eof
if errorlevel 2 goto executar_programa
if errorlevel 1 goto instalar_dependencias

:: ==============================================
:: INSTALAR DEPENDÊNCIAS (COM PROXY)
:: ==============================================
:instalar_dependencias
cls
title Instalador Python Automático (Proxy Corporativo) - Modo Seguro
color 0A

:: ==============================================
:: CONFIGURAÇÃO INICIAL
:: ==============================================
set "PROXY_SERVER=10.10.111.11:1111"
set "USUARIO=%USERNAME%"

:: ==============================================
:: CAPTURA E VALIDAÇÃO DA SENHA
:: ==============================================
:captura_senha
call :getPassword "Digite a senha do proxy corporativo: " SENHA1
call :getPassword "Confirme a senha do proxy corporativo: " SENHA2

if not "%SENHA1%"=="%SENHA2%" (
    cls
    echo.
    echo [ERRO] As senhas não coincidem! Por favor, tente novamente.
    echo.
    set "SENHA1="
    set "SENHA2="
    goto captura_senha
)

set "SENHA=%SENHA1%"
set "SENHA1="
set "SENHA2="

:: ==============================================
:: TRATAMENTO DA SENHA
:: ==============================================
setlocal enabledelayedexpansion
set "SENHA_TRATADA="
if defined SENHA (
    call :encodeURL "!SENHA!" SENHA_TRATADA
)
set "PROXY_COMPLETO=http://%USUARIO%:!SENHA_TRATADA!@%PROXY_SERVER%"

:: ==============================================
:: VERIFICAÇÕES INTELIGENTES
:: ==============================================
echo.
echo [VERIFICANDO PRÉ-REQUISITOS...]

:: 1. Verifica o Python
where python >nul 2>&1
if %errorlevel% equ 0 (
    python --version
    echo ✓ Python já instalado
    set "PYTHON_INSTALADO=true"
) else (
    echo × Python não encontrado
    set "PYTHON_INSTALADO=false"
)

:: 2. Verifica o Pip
if "%PYTHON_INSTALADO%"=="true" (
    python -m pip --version >nul 2>&1
    if %errorlevel% equ 0 (
        echo ✓ Pip já instalado
        set "PIP_INSTALADO=true"
    ) else (
        echo × Pip não encontrado
        set "PIP_INSTALADO=false"
    )
) else (
    set "PIP_INSTALADO=false"
)

:: 3. Verifica as Bibliotecas
if "%PIP_INSTALADO%"=="true" (
    for %%b in (%BIBLIOTECAS%) do (
        pip show %%b >nul 2>&1
        if !errorlevel! equ 0 (
            echo ✓ %%b já instalado
        ) else (
            echo × %%b não encontrado
            set "INSTALAR_%%b=true"
        )
    )
)

:: ==============================================
:: INSTALAÇÕES CONDICIONAIS
:: ==============================================
echo.
echo [INICIANDO INSTALAÇÕES NECESSÁRIAS...]

:: 1. Instala Python
if "%PYTHON_INSTALADO%"=="false" (
    echo [1/3] Baixando Python...
    set "HTTP_PROXY=%PROXY_COMPLETO%"
    set "HTTPS_PROXY=%PROXY_COMPLETO%"
    curl --proxy "%PROXY_COMPLETO%" --fail -L -o python_installer.exe https://www.python.org/ftp/python/3.11.4/python-3.11.4-amd64.exe || (
        echo ERRO: Falha no download
        goto clean_up
    )
    
    echo [2/3] Instalando Python...
    start /wait python_installer.exe /quiet InstallAllUsers=1 PrependPath=1 Include_launcher=0 || (
        echo ERRO: Falha na instalação
        goto clean_up
    )
    del /q python_installer.exe 2>nul
    set "PYTHON_INSTALADO=true"
    echo ✓ Python instalado com sucesso
    
    :: Atualizar PATH 
    set "PATH=%PATH%;%ProgramFiles%\Python311;%ProgramFiles%\Python311\Scripts"
)

:: 2. Instala o Pip se necessario
if "%PYTHON_INSTALADO%"=="true" (
    if "%PIP_INSTALADO%"=="false" (
        echo [3/3] Configurando pip...
        python -m ensurepip --upgrade
        python -m pip install --upgrade pip --proxy "%PROXY_COMPLETO%" || (
            echo ERRO: Falha ao configurar pip
            goto clean_up
        )
        set "PIP_INSTALADO=true"
        echo ✓ Pip configurado com sucesso
    )
)

:: 3. Instalar Bibliotecas faltantes
if "%PIP_INSTALADO%"=="true" (
    echo.
    echo [VERIFICANDO BIBLIOTECAS NOVAMENTE...]
    for %%b in (%BIBLIOTECAS%) do (
        pip show %%b >nul 2>&1
        if !errorlevel! neq 0 (
            echo Instalando %%b...
            pip install %%b --proxy "%PROXY_COMPLETO%" --user --no-warn-script-location || (
                echo ERRO CRÍTICO: Falha ao instalar %%b
                goto clean_up
            )
            echo ✓ %%b instalado com sucesso
        )
    )
)

:: ==============================================
:: LIMPEZA FINAL
:: ==============================================
:clean_up
del /q python_installer.exe 2>nul
python -m pip config unset global.proxy 2>nul
set "PROXY_COMPLETO="
set "SENHA_TRATADA="
set "SENHA="
set "HTTP_PROXY="
set "HTTPS_PROXY="

:: ==============================================
:: RELATÓRIO FINAL
:: ==============================================
echo.
echo [RELATÓRIO FINAL]
if "%PYTHON_INSTALADO%"=="true" (
    python --version
    pip --version
    echo.
    echo Bibliotecas instaladas:
    for %%b in (%BIBLIOTECAS%) do (
        pip show %%b >nul 2>&1 && echo ✓ %%b || echo × %%b
    )
)
echo.
echo ✓ Processo concluído com segurança
pause
goto menu_principal

:: ==============================================
:: EXECUTAR PROGRAMA
:: ==============================================
:executar_programa
cls
echo.
echo  ============================================
echo    EXECUTANDO PROGRAMA PRINCIPAL
echo  ============================================
echo.

if not exist "%SCRIPT_DIR%%PYTHON_SCRIPT%" (
    echo Erro: Arquivo %PYTHON_SCRIPT% não encontrado!
    pause
    goto menu_principal
)

call :center_text "Executando %PYTHON_SCRIPT%..."
python "%SCRIPT_DIR%%PYTHON_SCRIPT%"

pause
goto menu_principal

:: ==============================================
:: FUNÇÕES AUXILIARES
:: ==============================================

:center_text
setlocal enabledelayedexpansion
set "text=%~1"

for /f "tokens=2 delims=:" %%A in ('mode con ^| findstr "Colunas"') do set /a width=%%A
set /a spaces=(%width% - 18) / 2
set "space="
for /l %%i in (1,1,%spaces%) do set "space=!space! "

echo.
echo !space!!text!
echo.
endlocal
goto :eof

:encodeURL <string> <var_result>
setlocal enabledelayedexpansion
set "str=%~1"
set "result="

:encode_loop
if defined str (
    set "char=!str:~0,1!"
    set "str=!str:~1!"
    
    if "!char!"=="@" set "char=%%40"
    if "!char!"=="!" set "char=%%21"
    if "!char!"=="#" set "char=%%23"
    if "!char!"=="$" set "char=%%24"
    if "!char!"=="^" set "char=%%5E"
    if "!char!"=="&" set "char=%%26"
    if "!char!"==";" set "char=%%3B"
    if "!char!"=="+" set "char=%%2B"
    if "!char!"=="," set "char=%%2C"
    if "!char!"=="<" set "char=%%3C"
    if "!char!"==">" set "char=%%3E"
    if "!char!"=="/" set "char=%%2F"
    if "!char!"=="?" set "char=%%3F"
    if "!char!"=="[" set "char=%%5B"
    if "!char!"=="]" set "char=%%5D"
    if "!char!"=="|" set "char=%%7C"
    if "!char!"=="\" set "char=%%5C"
    if "!char!"=="'" set "char=%%27"
    if "!char!"==" " set "char=%%20"
    
    set "result=!result!!char!"
    goto encode_loop
)
endlocal & set "%~2=%result%"
goto :eof

:getPassword <prompt> <var_name>
setlocal disabledelayedexpansion
set "psCommand=powershell -Command "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8; $p=read-host '%1' -AsSecureString; $m=[Runtime.InteropServices.Marshal]; $str=$m::PtrToStringAuto($m::SecureStringToBSTR($p)); [Console]::WriteLine($str)""
for /f "usebackq delims=" %%p in (`%psCommand%`) do set "password=%%p"
endlocal & set "%2=%password%"
goto :eof