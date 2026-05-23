#!/opt/alt/python311/bin/python3.11
import os
import sys
import traceback
from wsgiref.handlers import CGIHandler

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
PACKAGES_DIR = "/home/softdayt/python311_packages"
PREFIX = "/app.cgi"

os.environ.setdefault("PYTHONDONTWRITEBYTECODE", "1")
os.environ.setdefault("SOFTDAYTA_RISK_DATA_DIR", "/home/softdayt/risk_data")
os.environ.setdefault("PEP_DATABASE_URI", "sqlite:////home/softdayt/risk_data/pep_portal.db")
os.environ.setdefault("SEED_DEMO_USERS", "0")
os.environ.setdefault("SECRET_KEY", "oq1kk6A9XY5JCU8KpT9mi_DQZTsgi9VnMNfA-Mv3eDureS7pYHKMGkkt_636CqmO")

if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)
if PACKAGES_DIR not in sys.path:
    sys.path.insert(1, PACKAGES_DIR)

os.chdir(BASE_DIR)


class RootMountedApp:
    def __init__(self, app):
        self.app = app

    def __call__(self, environ, start_response):
        script_name = environ.get("SCRIPT_NAME", "")
        path_info = environ.get("PATH_INFO", "")
        request_uri = environ.get("REQUEST_URI", "")

        if script_name.startswith(PREFIX):
            environ["SCRIPT_NAME"] = ""

        request_path = (request_uri or "").split("?", 1)[0].strip() or ""
        if request_path.startswith(PREFIX):
            request_path = request_path[len(PREFIX) :] or "/"
        elif path_info.startswith(PREFIX):
            request_path = path_info[len(PREFIX) :] or "/"
        elif not request_path:
            request_path = path_info or "/"

        if not str(request_path).startswith("/"):
            request_path = f"/{request_path}"

        environ["PATH_INFO"] = request_path or "/"
        return self.app(environ, start_response)


try:
    import app as app_module

    flask_app = getattr(app_module, "app", None)
    if flask_app is None:
        flask_app = app_module.create_app()
    CGIHandler().run(RootMountedApp(flask_app))
except Exception:
    body = "\n".join(
        [
            "CGI bootstrap error",
            "",
            traceback.format_exc(),
        ]
    ).encode("utf-8", errors="replace")
    print("Status: 500 Internal Server Error")
    print("Content-Type: text/plain; charset=utf-8")
    print(f"Content-Length: {len(body)}")
    print()
    sys.stdout.buffer.write(body)
