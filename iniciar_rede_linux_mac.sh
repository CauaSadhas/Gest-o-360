#!/usr/bin/env bash
cd "$(dirname "$0")"
echo "====================================================="
echo " Gestão360 Contábil - modo rede local"
echo "====================================================="
echo "Este computador será o servidor principal."
echo "Deixe esta janela aberta enquanto outras pessoas usam o sistema."
echo ""
echo "Neste computador: http://127.0.0.1:5000"
echo "Em outros computadores, use o IP local deste computador com a porta 5000."
echo "Exemplo: http://192.168.0.10:5000"
echo ""
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
export GESTAO360_HOST=0.0.0.0
export GESTAO360_PORT=5000
python app.py
