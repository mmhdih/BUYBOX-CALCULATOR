import socket

import launch


def _listener(host: str = "0.0.0.0") -> socket.socket:
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.bind((host, 0))
    s.listen()
    return s


def test_default_port_is_not_streamlit_default():
    assert launch.DEFAULT_PORT != 8501


def test_port_used_on_any_interface_is_busy_and_skipped():
    other = _listener("0.0.0.0")  # like another Streamlit app (e.g. DK ATP)
    try:
        busy = other.getsockname()[1]
        assert not launch.port_is_free(busy)
        assert launch.free_port(busy) != busy
    finally:
        other.close()


def test_port_used_on_localhost_is_busy():
    other = _listener("127.0.0.1")
    try:
        assert not launch.port_is_free(other.getsockname()[1])
    finally:
        other.close()


def _serve(body: bytes):
    import http.server
    import threading

    class H(http.server.BaseHTTPRequestHandler):
        def do_GET(self):  # noqa: N802
            self.send_response(200)
            self.end_headers()
            self.wfile.write(body)

        def log_message(self, *a):
            pass

    srv = http.server.HTTPServer(("127.0.0.1", 0), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv


def test_running_instance_detects_our_app_only(tmp_path, monkeypatch):
    monkeypatch.setattr(launch, "INSTANCE_FILE", tmp_path / ".instance.json")
    assert launch.running_instance() is None  # no file
    ours, other = _serve(b"buybox-analyzer\n"), _serve(b"<html>DK ATP</html>")
    try:
        for srv, expected in ((ours, True), (other, False)):
            port = srv.server_address[1]
            launch.remember_instance(port)
            assert (launch.running_instance() == port) is expected
    finally:
        ours.shutdown()
        other.shutdown()
    launch.remember_instance(ours.server_address[1])  # stale file, server gone
    assert launch.running_instance() is None


def test_local_check_ignores_system_proxy(monkeypatch):
    monkeypatch.setenv("HTTP_PROXY", "http://127.0.0.1:9")  # unreachable proxy
    monkeypatch.setenv("http_proxy", "http://127.0.0.1:9")
    srv = _serve(b"buybox-analyzer")
    try:
        assert launch.is_our_app(srv.server_address[1])
    finally:
        srv.shutdown()
