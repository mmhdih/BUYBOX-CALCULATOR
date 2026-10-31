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
