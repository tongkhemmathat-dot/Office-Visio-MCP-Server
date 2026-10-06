"""Pure helpers for system diagrams: spec validation, tier layout and rack arithmetic.

Nothing here touches Visio, so it can be unit tested anywhere. All coordinates are
in inches; for rack diagrams they are drawing (real-world) inches.
"""

from typing import Dict, List, Optional, Tuple

RACK_UNIT = 1.75  # inches per rack unit

DIRECTIONS = ("top-down", "left-right")
ARROWS = ("none", "end", "begin", "both")
CONNECTOR_TYPES = ("dynamic", "straight", "curved")


class SpecError(ValueError):
    """The diagram spec is not valid."""


def normalize_spec(spec: dict) -> dict:
    """Validate a diagram spec and return a copy with defaults filled in.

    Spec shape (all keys optional except nodes):
        title, subtitle, direction ("top-down" | "left-right"), icon_size,
        page {width, height},
        groups [{label, nodes: [ids], color}],
        nodes  [{id, label, tier, stencil, master, fill_color, text_color, width, height}],
        links  [{from, to, label, type, color, pattern, weight, arrow}]
    """
    if not isinstance(spec, dict):
        raise SpecError("spec must be an object")
    nodes = spec.get("nodes")
    if not nodes or not isinstance(nodes, list):
        raise SpecError("spec.nodes must be a non-empty list")

    ids = set()
    norm_nodes = []
    for i, node in enumerate(nodes):
        if not isinstance(node, dict) or not node.get("id"):
            raise SpecError(f"nodes[{i}] needs an 'id'")
        node_id = str(node["id"])
        if node_id in ids:
            raise SpecError(f"Duplicate node id '{node_id}'")
        ids.add(node_id)
        if bool(node.get("stencil")) != bool(node.get("master")):
            raise SpecError(f"Node '{node_id}': 'stencil' and 'master' must be given together")
        tier = node.get("tier", 0)
        if not isinstance(tier, int) or isinstance(tier, bool) or tier < 0:
            raise SpecError(f"Node '{node_id}': tier must be a non-negative integer")
        norm_nodes.append({**node, "id": node_id, "label": node.get("label", node_id), "tier": tier})

    norm_links = []
    for i, link in enumerate(spec.get("links", [])):
        if not isinstance(link, dict) or "from" not in link or "to" not in link:
            raise SpecError(f"links[{i}] needs 'from' and 'to'")
        for end in ("from", "to"):
            if str(link[end]) not in ids:
                raise SpecError(f"links[{i}].{end} refers to unknown node '{link[end]}'")
        if str(link.get("type", "straight")).lower() not in CONNECTOR_TYPES:
            raise SpecError(f"links[{i}].type must be one of {', '.join(CONNECTOR_TYPES)}")
        if str(link.get("arrow", "none")).lower() not in ARROWS:
            raise SpecError(f"links[{i}].arrow must be one of {', '.join(ARROWS)}")
        norm_links.append({**link, "from": str(link["from"]), "to": str(link["to"])})

    norm_groups = []
    for i, group in enumerate(spec.get("groups", [])):
        members = [str(m) for m in group.get("nodes", [])]
        if not members:
            raise SpecError(f"groups[{i}] needs a non-empty 'nodes' list")
        for m in members:
            if m not in ids:
                raise SpecError(f"groups[{i}] refers to unknown node '{m}'")
        norm_groups.append({**group, "nodes": members})

    direction = spec.get("direction", "top-down")
    if direction not in DIRECTIONS:
        raise SpecError(f"direction must be one of {', '.join(DIRECTIONS)}")
    icon_size = float(spec.get("icon_size", 1.0))
    if icon_size <= 0:
        raise SpecError("icon_size must be positive")

    return {**spec, "nodes": norm_nodes, "links": norm_links, "groups": norm_groups,
            "direction": direction, "icon_size": icon_size}


def tier_layout(nodes: List[dict], page_width: float, page_height: float,
                direction: str = "top-down", margin: float = 0.8,
                top_reserved: float = 0.0) -> Dict[str, Tuple[float, float]]:
    """Return the centre (x, y) of every node, with nodes grouped into rows (or columns) by tier.

    top-down: tier 0 is the top row. left-right: tier 0 is the left column.
    Nodes inside a tier keep their order and are spread evenly. top_reserved keeps
    room at the top of the page for a title.
    """
    if direction not in DIRECTIONS:
        raise SpecError(f"direction must be one of {', '.join(DIRECTIONS)}")
    tiers: Dict[int, List[dict]] = {}
    for node in nodes:
        tiers.setdefault(node.get("tier", 0), []).append(node)
    ordered = [tiers[t] for t in sorted(tiers)]

    left, right = margin, page_width - margin
    bottom, top = margin, page_height - margin - top_reserved
    positions: Dict[str, Tuple[float, float]] = {}

    for t_index, members in enumerate(ordered):
        count = len(ordered)
        for m_index, node in enumerate(members):
            if direction == "top-down":
                y = (top + bottom) / 2 if count == 1 else top - t_index * (top - bottom) / (count - 1)
                x = left + (m_index + 0.5) * (right - left) / len(members)
            else:
                x = (left + right) / 2 if count == 1 else left + t_index * (right - left) / (count - 1)
                y = top - (m_index + 0.5) * (top - bottom) / len(members)
            positions[node["id"]] = (x, y)
    return positions


def group_bounds(member_ids: List[str], positions: Dict[str, Tuple[float, float]],
                 sizes: Dict[str, Tuple[float, float]], pad: float = 0.4,
                 label_space: float = 0.35) -> Tuple[float, float, float, float]:
    """Return (left, bottom, right, top) of a box around the members, with room for a label on top."""
    lefts, bottoms, rights, tops = [], [], [], []
    for m in member_ids:
        x, y = positions[m]
        w, h = sizes[m]
        lefts.append(x - w / 2)
        rights.append(x + w / 2)
        bottoms.append(y - h / 2)
        tops.append(y + h / 2)
    return min(lefts) - pad, min(bottoms) - pad - 0.25, max(rights) + pad, max(tops) + pad + label_space


# --- Rack arithmetic ----------------------------------------------------------

def rack_device_bottom(rack_u1_y: float, u_position: int) -> float:
    """Drawing-inch y of the bottom edge of a device whose lowest unit is u_position."""
    return rack_u1_y + (u_position - 1) * RACK_UNIT


def check_rack_fit(u_height: int, u_position: int, u_size: int,
                   occupied: List[Tuple[int, int, str]]) -> None:
    """Raise SpecError if the device is out of range or overlaps another one.

    occupied is a list of (first_u, size, label) for devices already in the rack.
    """
    if u_size < 1:
        raise SpecError("u_size must be at least 1")
    if u_position < 1 or u_position + u_size - 1 > u_height:
        raise SpecError(
            f"U{u_position}-{u_position + u_size - 1} does not fit in a {u_height}U rack")
    last = u_position + u_size - 1
    for first, size, label in occupied:
        if u_position <= first + size - 1 and first <= last:
            raise SpecError(
                f"U{u_position}-{last} overlaps '{label}' at U{first}-{first + size - 1}")


def free_ranges(u_height: int, occupied: List[Tuple[int, int, str]]) -> List[Tuple[int, int]]:
    """Return the free (first_u, last_u) ranges of a rack, lowest first."""
    used = [False] * (u_height + 1)
    for first, size, _ in occupied:
        for u in range(first, min(first + size, u_height + 1)):
            used[u] = True
    ranges: List[Tuple[int, int]] = []
    start: Optional[int] = None
    for u in range(1, u_height + 1):
        if not used[u] and start is None:
            start = u
        if used[u] and start is not None:
            ranges.append((start, u - 1))
            start = None
    if start is not None:
        ranges.append((start, u_height))
    return ranges
