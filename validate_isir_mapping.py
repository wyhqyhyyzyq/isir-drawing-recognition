from __future__ import annotations

import argparse
import json
import math
from pathlib import Path
from typing import Any


def intersects(a: list[float], b: list[float]) -> bool:
    return not (a[2] < b[0] or a[0] > b[2] or a[3] < b[1] or a[1] > b[3])


def rectangle_gap(a: list[float], b: list[float]) -> float:
    dx = max(b[0] - a[2], a[0] - b[2], 0)
    dy = max(b[1] - a[3], a[1] - b[3], 0)
    return math.hypot(dx, dy)


def require_sequence(rows: list[dict[str, Any]], errors: list[str]) -> None:
    actual = [row.get("no") for row in rows]
    expected = list(range(1, len(rows) + 1))
    if actual != expected:
        errors.append(f"row numbers must be consecutive: expected {expected}, got {actual}")


def validate(path: Path, radius: float, max_gap: float) -> dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8"))
    view_rows = list(data.get("rows", []))
    non_balloon_rows = list(data.get("non_balloon_rows", []))
    all_rows = view_rows + non_balloon_rows
    errors: list[str] = []

    source_document = data.get("source_document")
    source_image = data.get("source_image")
    if not source_document:
        errors.append("mapping must record source_document")
    elif not Path(source_document).is_file():
        errors.append(f"source_document does not exist: {source_document}")
    if not source_image:
        errors.append("mapping must record source_image")
    elif not Path(source_image).is_file():
        errors.append(f"source_image does not exist: {source_image}")
    image_size = data.get("image_size")
    if (
        not isinstance(image_size, list)
        or len(image_size) != 2
        or not all(isinstance(value, (int, float)) and value > 0 for value in image_size)
    ):
        errors.append("mapping must record two-value image_size")
        image_width = image_height = 0
    else:
        image_width, image_height = (float(image_size[0]), float(image_size[1]))

    require_sequence(all_rows, errors)

    for row in all_rows:
        no = row.get("no")
        required_aliases = {
            "char_no": row.get("no"),
            "type": row.get("table_type"),
            "requirement_text": row.get("table_requirement"),
        }
        for field, expected in required_aliases.items():
            if row.get(field) != expected:
                errors.append(
                    f"row {no} field {field} must match normalized value {expected!r}"
                )
        if not isinstance(row.get("confidence"), (int, float)):
            errors.append(f"row {no} must record numeric confidence")
        if not isinstance(row.get("needs_review"), bool):
            errors.append(f"row {no} must record boolean needs_review")
        if row.get("visual_confirmed") is not True:
            errors.append(f"row {no} must set visual_confirmed=true")

        for field in ("nominal", "upper_tol", "lower_tol"):
            value = row.get(field)
            if field not in row or (
                value is not None and not isinstance(value, (int, float))
            ):
                errors.append(f"row {no} must record numeric or null {field}")
        if "unit" not in row or not isinstance(row.get("unit"), str):
            errors.append(f"row {no} must record string unit")

        source_text = str(row.get("source_parameter_text", ""))
        requirement = str(row.get("table_requirement", ""))
        if "（" in requirement and "（" not in source_text:
            if not row.get("position_qualifier") or not row.get("qualifier_source"):
                errors.append(
                    f"row {no} adds a parenthetical qualifier without recording its source"
                )

    source_boxes: list[tuple[int, list[float]]] = []
    for row in view_rows:
        bbox = row.get("source_bbox")
        if not isinstance(bbox, list) or len(bbox) != 4:
            errors.append(f"row {row.get('no')} has invalid source_bbox")
            continue
        if not all(isinstance(value, (int, float)) for value in bbox):
            errors.append(f"row {row.get('no')} source_bbox must be numeric")
            continue
        numeric_bbox = [float(value) for value in bbox]
        if numeric_bbox[0] >= numeric_bbox[2] or numeric_bbox[1] >= numeric_bbox[3]:
            errors.append(f"row {row.get('no')} source_bbox has invalid ordering")
        if (
            numeric_bbox[0] < 0
            or numeric_bbox[1] < 0
            or numeric_bbox[2] > image_width
            or numeric_bbox[3] > image_height
        ):
            errors.append(f"row {row.get('no')} source_bbox lies outside image_size")
        source_boxes.append((row["no"], numeric_bbox))

    balloon_boxes: list[tuple[int, list[float], list[float]]] = []
    for row in view_rows:
        no = row.get("no")
        center = row.get("balloon_center")
        bbox = row.get("source_bbox")
        if not isinstance(center, list) or len(center) != 2:
            errors.append(f"row {no} has no valid balloon_center")
            continue
        if not all(isinstance(value, (int, float)) for value in center):
            errors.append(f"row {no} balloon_center must be numeric")
            continue
        if not isinstance(bbox, list) or len(bbox) != 4:
            continue

        cx, cy = (float(center[0]), float(center[1]))
        if cx - radius < 0 or cy - radius < 0 or cx + radius > image_width or cy + radius > image_height:
            errors.append(f"balloon {no} lies outside image_size")
        circle_box = [cx - radius, cy - radius, cx + radius, cy + radius]
        own_bbox = [float(value) for value in bbox]
        gap = rectangle_gap(circle_box, own_bbox)
        if gap > max_gap:
            errors.append(
                f"balloon {no} is too far from its parameter: gap={gap:.1f}, max={max_gap:.1f}"
            )

        for source_no, source_bbox in source_boxes:
            if intersects(circle_box, source_bbox):
                errors.append(f"balloon {no} overlaps parameter text for row {source_no}")

        leader_to = row.get("leader_to")
        if leader_to is not None:
            if not isinstance(leader_to, list) or len(leader_to) != 2:
                errors.append(f"balloon {no} has invalid leader_to")
            else:
                leader_length = math.hypot(cx - leader_to[0], cy - leader_to[1])
                if leader_length > radius * 2:
                    errors.append(
                        f"balloon {no} has a long leader: length={leader_length:.1f}"
                    )

        balloon_boxes.append((no, [cx, cy], circle_box))

    for index, (no_a, center_a, _) in enumerate(balloon_boxes):
        for no_b, center_b, _ in balloon_boxes[index + 1 :]:
            distance = math.hypot(
                center_a[0] - center_b[0], center_a[1] - center_b[1]
            )
            if distance < radius * 2:
                errors.append(
                    f"balloons {no_a} and {no_b} overlap: center distance={distance:.1f}"
                )

    non_balloon_types = [row.get("table_type") for row in non_balloon_rows]
    if non_balloon_types != ["材质", "外观"]:
        errors.append(
            "non_balloon_rows must end with exactly two rows in this order: 材质, 外观"
        )

    for row in non_balloon_rows:
        no = row.get("no")
        if row.get("balloon_center") is not None or row.get("leader_to") is not None:
            errors.append(f"non-balloon row {no} must not have a balloon or leader")
        if row.get("balloon_required") is not False:
            errors.append(f"non-balloon row {no} must set balloon_required=false")
        if not row.get("source_region"):
            errors.append(f"non-balloon row {no} must record source_region")

    return {
        "status": "passed" if not errors else "failed",
        "view_rows": len(view_rows),
        "non_balloon_rows": len(non_balloon_rows),
        "errors": errors,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("mapping", type=Path)
    parser.add_argument("--radius", type=float, default=32)
    parser.add_argument("--max-gap", type=float, default=24)
    args = parser.parse_args()

    result = validate(args.mapping, args.radius, args.max_gap)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if result["errors"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
