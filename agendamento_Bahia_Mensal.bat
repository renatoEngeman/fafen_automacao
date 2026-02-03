echo on
setlocal

rem BASE = pasta onde este .bat está (sempre com \ no final)
set "BASE=%~dp0"

rem Garante que o working dir é a pasta do projeto, mesmo se a tarefa iniciar em System32
pushd "%BASE%"

rem (Opcional) Ativar venv — mas em tarefa agendada é melhor chamar o python.exe direto
rem call "%BASE%venv\Scripts\activate.bat"

rem Melhor prática: chamar o interpretador do venv diretamente (sem ativar)
"%BASE%venv\Scripts\python.exe" "%BASE%main_BA_Mensal.py"



popd
endlocal