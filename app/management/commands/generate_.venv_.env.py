import os
import sys
sys.path.insert(0, '../mapsProj')
from django.core.management.utils import get_random_secret_key
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent


secret_key = get_random_secret_key()
with open(os.path.join(BASE_DIR, '.env'), 'w') as f:
    f.write('DEBUG=True\n')
    f.write('PORT=8000\n')
    f.write(f'SECRET_KEY={secret_key}\n')

os.system('python3 -m venv ' + str(BASE_DIR / '.venv'))


#! this generates the .venv environment and .env file in the root directory
#* ctrl + alt + numpad1         to run this snippet