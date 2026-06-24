@echo off
cd /d %~dp0

echo =====================================================
echo  Gestao360 Contabil - modo rede local
echo =====================================================
echo.
echo Este computador sera o servidor principal.
echo Deixe esta janela aberta enquanto outras pessoas usam o sistema.
echo.
echo Links para acesso:
echo - Neste computador: http://127.0.0.1:5000
echo - Em outros computadores: http://SEU-IP:5000
echo.
echo Enderecos IPv4 encontrados neste computador:
ipconfig | findstr /i "IPv4"
echo.
echo Use o IPv4 da sua rede, normalmente parecido com 192.168.0.x ou 192.168.1.x.
echo Exemplo para sua colega: http://192.168.0.10:5000
echo.
echo Se o Firewall do Windows pedir permissao, marque Redes privadas e clique em Permitir acesso.
echo.

python -m venv .venv
call .venv\Scripts\activate
pip install -r requirements.txt
set GESTAO360_HOST=0.0.0.0
set GESTAO360_PORT=5000
python app.py
pause
