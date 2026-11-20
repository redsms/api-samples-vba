"""Run from the repository root. Only a loopback server and dummy keys are used."""
import hashlib
import json
import os
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
import subprocess
import tempfile
import threading
from xml.sax.saxutils import escape
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
    # LibreOffice Basic in VBA compatibility mode hosts the module on Windows.
    profile = Path(tmp, "profile")
    basic = profile / "user/basic"
    standard = basic / "Standard"
    office = str(Path(os.environ["ProgramFiles"], "LibreOffice/program/soffice.com"))
    subprocess.run([office, "-env:UserInstallation=" + profile.as_uri(), "--headless", "--terminate_after_init"], check=True, timeout=60)
    standard.mkdir(parents=True, exist_ok=True)
    ns = 'xmlns:library="http://openoffice.org/2000/library" xmlns:xlink="http://www.w3.org/1999/xlink"'
    (basic / "script.xlc").write_text(f'<library:libraries {ns}><library:library library:name="Standard" library:link="false"/></library:libraries>')
    (standard / "script.xlb").write_text(f'<library:library {ns} library:name="Standard" library:readonly="false" library:passwordprotected="false"><library:element library:name="RedSms"/></library:library>')
    code = "Option VBASupport 1\n" + code.partition("\n")[2].replace("Debug.Print ", "Print #1, ")
    source_lines = code.splitlines()
    # VBA line labels make automation errors diagnosable on a headless runner.
    code = "\n".join(f"{i} {line}" if line.startswith("    ") and not line.lstrip().startswith(("Dim ", "Else", "End If", "Next")) else line
                     for i, line in enumerate(source_lines, 1))
    driver = """
Sub RunTest
    On Error GoTo Failed
    Open Environ("REDSMS_OUTPUT") For Output As #1
    Select Case Environ("REDSMS_COMMAND")
        Case "info"
            ClientInfo
        Case "send"
            SendSms "Тестовое сообщение от redsms.ru"
        Case "status"
            MessageStatus "00000000-0000-4000-8000-000000000001"
    End Select
    Close #1
    Exit Sub
Failed:
    Print #1, "ERROR " & Erl & " " & Err.Description
    Close #1
End Sub
"""
    xml = '<script:module xmlns:script="http://openoffice.org/2000/script" script:name="RedSms" script:language="StarBasic">' + escape(code + driver) + '</script:module>'
    (standard / "RedSms.xba").write_text(xml, encoding="utf-8")
    env = dict(os.environ, REDSMS_LOGIN="test", REDSMS_API_KEY="dummy", REDSMS_FROM="A & B", REDSMS_TO="+70000000000")
    thread = threading.Thread(target=server.serve_forever)
    thread.start()
    try:
        command = [str(Path(os.environ["ProgramFiles"], "LibreOffice/program/soffice.com")),
                   "-env:UserInstallation=" + profile.as_uri(), "--headless", "--norestore",
                   "macro:///Standard.RedSms.RunTest()"]
        for action, field in [("info", "info"), ("send", "items"), ("status", "item"), ("send", "420")]:
            output = Path(tmp, field + ".txt")
            env.update(REDSMS_OUTPUT=str(output), REDSMS_COMMAND=action)
            if field == "420":
                env["REDSMS_FROM"] = "reject"
            subprocess.run(command, env=env, timeout=60, check=True)
            body = output.read_bytes().decode("utf-8", errors="replace")
            if field == "420":
                assert "ERROR" in body and "420" in body, body
            else:
                assert not body.startswith("ERROR"), body
                assert field in json.loads(body), body
        assert calls == [("GET", "/api/client/info", 200), ("POST", "/api/message", 200),
                         ("GET", "/api/message/" + UUID, 200), ("POST", "/api/message", 420)], calls
    finally:
        server.shutdown()
        thread.join()
