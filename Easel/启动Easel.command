#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"
export PATH="$PWD/.venv/node/bin:$PWD/.venv/bin:$HOME/.local/bin:$PATH"

.venv/bin/python - <<'PY'
import fcntl
import json
from pathlib import Path
import socket
import subprocess
import sys
import time
import urllib.request

root = Path.cwd()
url = 'http://127.0.0.1:7860'
http = urllib.request.build_opener(urllib.request.ProxyHandler({}))

def ready():
    try:
        with http.open(url + '/openapi.json', timeout=2) as response:
            return json.load(response).get('info', {}).get('title') == 'Easel'
    except Exception:
        return False

with (root / 'outputs' / '.launch.lock').open('w') as lock:
    fcntl.flock(lock, fcntl.LOCK_EX)
    if not ready():
        with socket.socket() as probe:
            if probe.connect_ex(('127.0.0.1', 7860)) == 0:
                raise SystemExit('7860 端口已被其他程序占用，请关闭该程序后重试。')

        from easel.direct_api import read_settings
        from easel import harnesses
        runtime = harnesses.settings(root / 'outputs', read_settings(root / '.env'))['runtime']
        if runtime == 'openclaw':
            from easel.gateway_endpoint import resolve_gateway_port
            with socket.socket() as probe:
                occupied = probe.connect_ex(('127.0.0.1', resolve_gateway_port())) == 0
            if not occupied:
                try:
                    with (root / 'outputs' / 'gateway-launch.log').open('a') as log:
                        subprocess.run(['bash', 'scripts/gateway.sh', 'start'],
                                       stdout=log, stderr=subprocess.STDOUT, timeout=75)
                except subprocess.TimeoutExpired:
                    print('网关启动较慢，可在工作台设置中查看状态。')

        with (root / 'outputs' / 'web-local.log').open('a') as log:
            server = subprocess.Popen(
                [sys.executable, '-m', 'uvicorn', 'web.app:app',
                 '--host', '127.0.0.1', '--port', '7860', '--log-level', 'warning'],
                cwd=root, stdin=subprocess.DEVNULL, stdout=log,
                stderr=subprocess.STDOUT, start_new_session=True)
        for _ in range(100):
            if ready():
                break
            if server.poll() is not None:
                raise SystemExit('启动失败，请查看 outputs/web-local.log。')
            time.sleep(0.3)
        else:
            raise SystemExit('启动超时，请查看 outputs/web-local.log。')

print('Easel 已启动：' + url)
subprocess.run(['open', url], check=False)
PY
