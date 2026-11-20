"""Run from the repository root. Only a loopback server and dummy keys are used."""
import hashlib
import json
import os
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
import subprocess
import tempfile
import threading
from urllib.parse import parse_qs, urlsplit

UUID = "00000000-0000-4000-8000-000000000001"
TEXT = "Тестовое сообщение от redsms.ru"
calls = []


class Handler(BaseHTTPRequestHandler):
    def handle_api(self):
        path = urlsplit(self.path)
        data = path.query if self.command == "GET" else self.rfile.read(int(self.headers["Content-Length"])).decode()
        fields = {key: value[0] for key, value in parse_qs(data).items()}
        values = "".join(fields[key] for key in sorted(fields))
        expected = hashlib.md5((values + self.headers["ts"] + "dummy").encode()).hexdigest()
        status = 200
        body = {"success": True}
        if self.headers["login"] != "test" or expected != self.headers["sig"]:
            status = 401
        elif self.command == "POST" and path.path == "/api/message":
            if fields["from"] == "reject":
                status, body = 420, {"error_message": "Недостаточно средств", "success": False}
            elif fields != dict(route="sms", to="+70000000000", text=TEXT, **{"from": "A & B"}):
                status = 400
            else:
                body.update(items=[{"uuid": UUID}], errors=[])
        elif self.command == "GET" and path.path == "/api/message/" + UUID and fields == {"fields": "uuid,status,errCode"}:
            body.update(item={"uuid": UUID, "status": "reject", "errCode": 0})
        elif self.command == "GET" and path.path == "/api/client/info" and not fields:
            body.update(info={"login": "test"})
        else:
            status = 404
        calls.append((self.command, path.path, status))
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.end_headers()
        self.wfile.write(json.dumps(body, ensure_ascii=False).encode())

    do_POST = do_GET = handle_api

    def log_message(self, *args):
        pass


with tempfile.TemporaryDirectory() as tmp, HTTPServer(("127.0.0.1", 0), Handler) as server:
    # Run a temporary copy with a loopback address, keeping the published sample pointed at RED SMS.
    source = Path("RedSms.bas")
    code = source.read_text().replace("https://cp.redsms.ru/api/", f"http://127.0.0.1:{server.server_port}/api/")
    # Host the same procedures under Windows Script Host in CI (Excel is not installed).
    code = code.partition("\n")[2].replace("Debug.Print ", "WScript.Echo ")
    driver = """
Select Case WScript.Arguments.Count
Case 0
    ClientInfo
Case 1
    SendSms "Тестовое сообщение от redsms.ru"
Case 2
    MessageStatus WScript.Arguments(1)
End Select
"""
    Path(tmp, "test.vbs").write_text(code + driver, encoding="utf-16")
    env = dict(os.environ, REDSMS_LOGIN="test", REDSMS_API_KEY="dummy", REDSMS_FROM="A & B", REDSMS_TO="+70000000000")
    thread = threading.Thread(target=server.serve_forever)
    thread.start()
    try:
        command = ["cscript", "//nologo", "//U", "test.vbs"]
        for args, field in [([], "info"), (["send"], "items"), (["status", UUID], "item")]:
            result = subprocess.run(command + args, cwd=tmp, env=env, check=True, capture_output=True, encoding="utf-16-le")
            print(args, ascii(result.stdout), ascii(result.stderr), calls, flush=True)
            assert field in json.loads(result.stdout.lstrip('\ufeff'))
        env["REDSMS_FROM"] = "reject"
        result = subprocess.run(command + ["send"], cwd=tmp, env=env, capture_output=True, encoding="utf-16-le")
        assert result.returncode and "420" in result.stderr, result
        assert calls == [("GET", "/api/client/info", 200), ("POST", "/api/message", 200),
                         ("GET", "/api/message/" + UUID, 200), ("POST", "/api/message", 420)], calls
    finally:
        server.shutdown()
        thread.join()
