

0

windows:

winget install -e --id Tesseract.TesseractOCR
winget install -e --id poppler.poppler

linux:

sudo apt-get install tesseract-ocr tesseract-ocr-all
sudo apt-get install poppler-utils


1

pip install -r requirements.txt




3
set user name and pw:

'{"username": "DEIN_USERNAME", "password": "DEIN_PASSWORTE", "USERNAME_popup": "p", "PASSWORD_popup": "p"}' | Set-Content -Path credentials.json -Encoding utf8

4

execute:

python main.py --config bwl_master_config