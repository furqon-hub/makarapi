# Isi berkas WSGI di PythonAnywhere (tab Web -> klik tautan "WSGI configuration file").
# Ganti GANTI_USERNAME dengan username PythonAnywhere Anda.
import sys

sys.path.insert(0, "/home/GANTI_USERNAME/makarapi/app")

from a2wsgi import ASGIMiddleware  # noqa: E402
from api import app  # noqa: E402

application = ASGIMiddleware(app)
