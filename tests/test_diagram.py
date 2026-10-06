import pytest

from visio_mcp_server.diagram import (
    RACK_UNIT,
    SpecError,
    check_rack_fit,
    free_ranges,
    group_bounds,
    normalize_spec,
    rack_device_bottom,
    tier_layout,
)


def spec(**overrides):
    base = {"nodes": [{"id": "a"}, {"id": "b", "tier": 1}], "links": [{"from": "a", "to": "b"}]}
    base.update(overrides)
    return base


def test_normalize_fills_defaults():
    out = normalize_spec(spec())
    assert out["direction"] == "top-down"
    assert out["icon_size"] == 1.0
    assert out["nodes"][0]["label"] == "a"
    assert out["nodes"][0]["tier"] == 0


@pytest.mark.parametrize("bad", [
    {"nodes": []},
    {"nodes": [{"id": "a"}, {"id": "a"}]},
    {"nodes": [{"label": "no id"}]},
    {"nodes": [{"id": "a", "stencil": "SERVER_U"}]},          # stencil without master
    {"nodes": [{"id": "a", "tier": -1}]},
    {"nodes": [{"id": "a"}], "links": [{"from": "a", "to": "zzz"}]},
    {"nodes": [{"id": "a"}, {"id": "b"}], "links": [{"from": "a", "to": "b", "type": "wavy"}]},
    {"nodes": [{"id": "a"}, {"id": "b"}], "links": [{"from": "a", "to": "b", "arrow": "up"}]},
    {"nodes": [{"id": "a"}], "groups": [{"nodes": ["zzz"]}]},
    {"nodes": [{"id": "a"}], "direction": "diagonal"},
    {"nodes": [{"id": "a"}], "icon_size": 0},
])
def test_normalize_rejects_invalid(bad):
    with pytest.raises(SpecError):
        normalize_spec(bad)


def test_tier_layout_top_down():
    nodes = [{"id": "fw", "tier": 0}, {"id": "s1", "tier": 1}, {"id": "s2", "tier": 1}]
    pos = tier_layout(nodes, 10, 8, margin=1)
    assert pos["fw"] == (5.0, 7.0)                 # top row, centred
    assert pos["s1"][1] == pos["s2"][1] == 1.0     # bottom row
    assert pos["s1"][0] < pos["s2"][0]
    assert pos["s1"][0] == pytest.approx(3.0)
    assert pos["s2"][0] == pytest.approx(7.0)


def test_tier_layout_single_tier_is_vertically_centred():
    pos = tier_layout([{"id": "a"}], 10, 8, margin=1)
    assert pos["a"] == (5.0, 4.0)


def test_tier_layout_left_right_and_reserved_title_space():
    nodes = [{"id": "a", "tier": 0}, {"id": "b", "tier": 1}]
    pos = tier_layout(nodes, 10, 8, direction="left-right", margin=1, top_reserved=1)
    assert pos["a"][0] < pos["b"][0]
    assert pos["a"][1] == pytest.approx((6 + 1) / 2)   # top = 8 - 1 - 1


def test_group_bounds_surround_members():
    pos = {"a": (2, 2), "b": (6, 2)}
    size = {"a": (1, 1), "b": (1, 1)}
    left, bottom, right, top = group_bounds(["a", "b"], pos, size, pad=0.5, label_space=0.3)
    assert left < 1.5 and right > 6.5
    assert bottom < 1.5 and top > 2.5


def test_rack_device_bottom():
    assert rack_device_bottom(14.0, 1) == 14.0
    assert rack_device_bottom(14.0, 42) == pytest.approx(14.0 + 41 * RACK_UNIT)


def test_check_rack_fit():
    occupied = [(30, 2, "server")]
    check_rack_fit(42, 28, 2, occupied)       # U28-29, touches but does not overlap
    check_rack_fit(42, 41, 2, occupied)       # top of rack
    for pos, size in ((42, 2), (0, 1), (29, 2), (30, 1), (31, 3)):
        with pytest.raises(SpecError):
            check_rack_fit(42, pos, size, occupied)
    with pytest.raises(SpecError):
        check_rack_fit(42, 5, 0, occupied)


def test_free_ranges():
    assert free_ranges(10, []) == [(1, 10)]
    assert free_ranges(10, [(1, 2, "a"), (5, 1, "b"), (9, 2, "c")]) == [(3, 4), (6, 8)]
    assert free_ranges(4, [(1, 4, "full")]) == []
