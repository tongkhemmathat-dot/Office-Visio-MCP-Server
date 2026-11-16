import pytest

pytest.importorskip("win32com")

from visio_mcp_server.visio_server import svg_arrow_size  # noqa: E402


@pytest.mark.parametrize("length,expected", [(0, 1), (6, 0), (10, 0), (14, 1), (20, 2), (28, 3), (60, 4)])
def test_arrow_size_buckets(length, expected):
    assert svg_arrow_size(length) == expected
