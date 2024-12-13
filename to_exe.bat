python3 -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

pyinstaller \
    --noconfirm \
    --onefile \
    -w \
    --add-data "svod_parser_v1:svod_parser_v1/" \
    --add-data "svod_parser_v2:svod_parser_v2/" \
    -n "Svod parser" \
    "app.py"