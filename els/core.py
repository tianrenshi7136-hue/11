from __future__ import annotations

import hashlib
import json
import math
import sqlite3
import uuid
from contextlib import contextmanager, closing
from datetime import datetime, timezone
from pathlib import Path


def new_project():
    return {
        "schema_version": 1, "id": str(uuid.uuid4()), "name": "新基坑项目", "subject_name": '',
        "length_m": 0.0, "width_m": 0.0, "depth_m": 0.0,
        "embedment_m": 0.0, "analysis_reference": "", "site_notes": "",
        "external_checks": "", "members": [], "template_manifest": "",
        "sheetpile": {},
    }


def finite(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def validate(project):
    errors = []
    if not isinstance(project, dict):
        return ["项目须为JSON对象"]
    if project.get("schema_version") != 1:
        errors.append("不支持的项目数据版本")
    if not isinstance(project.get("id"), str) or not project["id"].strip():
        errors.append("项目ID不能为空")
    if not str(project.get("name", "")).strip():
        errors.append("项目名称不能为空")
    for key, label in [("length_m", "长度"), ("width_m", "宽度"), ("depth_m", "开挖深度"), ("embedment_m", "入土深度")]:
        if not finite(project.get(key)) or project[key] <= 0:
            errors.append(f"{label}必须为正数，单位m")
    if not isinstance(project.get("analysis_reference"), str) or not project["analysis_reference"].strip():
        errors.append("请填写外部分析来源及版本")
    members = project.get("members")
    sheetpile = project.get('sheetpile', {})
    if sheetpile:
        from .sheetpile import validate_request
        errors.extend(validate_request(sheetpile))
    support = project.get('support_design') or (sheetpile.get('support_package') if isinstance(sheetpile,dict) else None)
    has_support=isinstance(support,dict) and bool(support.get('members'))
    if isinstance(support,dict) and (support.get('enabled') or has_support):
        from .supports import validate_package
        errors.extend(validate_package(support))
    if isinstance(members, list) and not members and (sheetpile or has_support):
        return errors
    if not isinstance(members, list) or not members:
        return errors + ["请录入至少一项构件／工况内力"]
    seen = set()
    geometry_by_member = {}
    for i, m in enumerate(members, 1):
        tag = f"第{i}项"
        if not isinstance(m, dict):
            errors.append(f"{tag}必须为构件记录")
            continue
        key = (m.get("member_id"), m.get("case"))
        if not all(isinstance(v, str) and v.strip() for v in key):
            errors.append(f"{tag}缺少构件编号或工况")
        elif key in seen:
            errors.append(f"{tag}构件编号与工况重复")
        if all(isinstance(v, str) for v in key):
            seen.add(key)
        member_id = m.get("member_id")
        geometry = tuple(m.get(field) for field in ("kind", "x_m", "y_m", "level_m", "effective_length_m"))
        if isinstance(member_id, str):
            if member_id in geometry_by_member and geometry_by_member[member_id] != geometry:
                errors.append(f"{tag}同一构件在不同工况的几何或有效长度不一致")
            geometry_by_member[member_id] = geometry
        if m.get("kind") not in ("钢板桩", "围檩", "支撑"):
            errors.append(f"{tag}构件类型无效")
        if m.get("value_basis") not in ("特征值", "设计值"):
            errors.append(f"{tag}须明确特征值或设计值")
        if m.get("force_basis") not in ("每延米", "单根"):
            errors.append(f"{tag}须明确每延米或单根")
        if m.get("kind") == "支撑" and m.get("force_basis") != "单根":
            errors.append(f"{tag}支撑轴力须按单根输入")
        for field in ("x_m", "y_m", "level_m", "effective_length_m", "n_kn", "m_knm", "v_kn"):
            if not finite(m.get(field)):
                errors.append(f"{tag} {field}必须为有限数值")
        if finite(m.get("effective_length_m")) and m["effective_length_m"] <= 0:
            errors.append(f"{tag}有效长度须大于0")
        for field, limit in (("x_m", "length_m"), ("y_m", "width_m")):
            if finite(m.get(field)) and finite(project.get(limit)) and not 0 <= m[field] <= project[limit]:
                errors.append(f"{tag}平面位置超出基坑范围")
        if finite(m.get("level_m")) and finite(project.get("depth_m")) and not -project["depth_m"] <= m["level_m"] <= 0:
            errors.append(f"{tag}标高须位于地面0与坑底之间")
    return errors


def digest(project):
    data = json.dumps(project, ensure_ascii=False, sort_keys=True, allow_nan=False)
    return hashlib.sha256(data.encode("utf-8")).hexdigest()


class ProjectStore:
    def __init__(self, path):
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        self.path = Path(path)
        with self.connect() as conn:
            conn.execute("CREATE TABLE IF NOT EXISTS revisions (project_id TEXT NOT NULL, revision INTEGER NOT NULL, name TEXT NOT NULL, saved_at TEXT NOT NULL, fingerprint TEXT NOT NULL, payload TEXT NOT NULL, PRIMARY KEY(project_id, revision))")

    @contextmanager
    def connect(self):
        conn = sqlite3.connect(self.path)
        try:
            with conn:
                yield conn
        finally:
            conn.close()

    def save(self, project):
        fingerprint = digest(project)
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            row = conn.execute("SELECT revision, fingerprint FROM revisions WHERE project_id=? ORDER BY revision DESC LIMIT 1", (project["id"],)).fetchone()
            if row and row[1] == fingerprint:
                return row[0]
            revision = row[0] + 1 if row else 1
            conn.execute("INSERT INTO revisions VALUES (?,?,?,?,?,?)", (project["id"], revision, project["name"], datetime.now(timezone.utc).isoformat(), fingerprint, json.dumps(project, ensure_ascii=False, allow_nan=False)))
            return revision

    def list_revisions(self):
        with self.connect() as conn:
            return conn.execute("SELECT project_id, revision, name, saved_at FROM revisions ORDER BY saved_at DESC").fetchall()

    def load(self, project_id, revision):
        with self.connect() as conn:
            row = conn.execute("SELECT payload FROM revisions WHERE project_id=? AND revision=?", (project_id, revision)).fetchone()
        if row is None:
            raise ValueError("项目版本不存在")
        return json.loads(row[0])

    def backup(self, target):
        if Path(target).resolve() == self.path.resolve():
            raise ValueError("备份路径不能与数据库相同")
        with self.connect() as source, closing(sqlite3.connect(target)) as destination:
            source.backup(destination)


def export_project(project, target):
    Path(target).write_text(json.dumps(project, ensure_ascii=False, indent=2, allow_nan=False), encoding="utf-8")


def import_project(source):
    project = json.loads(Path(source).read_text(encoding="utf-8-sig"))
    if not isinstance(project, dict) or project.get("schema_version") != 1:
        raise ValueError("不是受支持的ELS项目文件")
    errors = validate(project)
    if errors:
        raise ValueError("导入数据未通过校验：\n" + "\n".join(errors))
    normalized = new_project()
    normalized.update(project)
    for key in ("site_notes", "external_checks", "template_manifest"):
        if not isinstance(normalized[key], str):
            raise ValueError(f"{key}须为文本")
    return normalized
