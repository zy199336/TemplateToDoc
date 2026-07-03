from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
from typing import Any

from .compat import ensure_legacy_imports
from .docx_replace import apply_replacements
from .docx_scan import inspect_template
from .llm import deepseek_provider, generate_local_text
from .yaml_io import dotted_get, dotted_set, load_yaml, write_yaml


GLOBAL_FIELDS = [
    {
        "id": "topic",
        "path": "global.topic",
        "label": "新文档主题 / 项目主题",
        "default": "",
        "multiline": True,
    },
    {
        "id": "background",
        "path": "global.background",
        "label": "项目背景 / 基本情况",
        "default": "",
        "multiline": True,
    },
    {
        "id": "materials",
        "path": "global.materials",
        "label": "全局原始素材",
        "default": "",
        "multiline": True,
    },
    {
        "id": "requirements",
        "path": "global.requirements",
        "label": "全局写作要求",
        "default": "",
        "multiline": True,
    },
]


@dataclass(frozen=True)
class V2Paths:
    root: Path
    profiles_dir: Path
    projects_dir: Path

    @classmethod
    def from_root(cls, root: str | Path) -> "V2Paths":
        base = Path(root).resolve()
        return cls(base, base / "profiles", base / "projects")

    def profile_dir(self, profile_id: str) -> Path:
        return self.profiles_dir / profile_id

    def project_dir(self, project_id: str) -> Path:
        return self.projects_dir / project_id


def safe_id(value: str, fallback: str = "profile") -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\r\n\t]+', "_", str(value or "")).strip(" ._")
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned[:120] or fallback


class TemplateToDocV2Pipeline:
    def __init__(self, root: str | Path) -> None:
        self.paths = V2Paths.from_root(root)
        self.paths.profiles_dir.mkdir(parents=True, exist_ok=True)
        self.paths.projects_dir.mkdir(parents=True, exist_ok=True)

    def import_template(
        self,
        template_path: str | Path,
        profile_id: str | None = None,
        project_id: str | None = None,
    ) -> dict[str, Any]:
        source = Path(template_path).resolve()
        if not source.exists():
            raise FileNotFoundError(source)
        profile = safe_id(profile_id or source.stem, "profile")
        profile_dir = self.paths.profile_dir(profile)
        profile_dir.mkdir(parents=True, exist_ok=True)
        reference = profile_dir / "reference.docx"
        prepared = self._prepare_reference_docx(source, profile_dir)
        shutil.copy2(prepared, reference)

        inspected = inspect_template(reference)
        profile_data = {
            "schema_version": 1,
            "engine": "template_to_doc_v2",
            "profile_id": profile,
            "source_template": str(source),
            "template": "reference.docx",
            "targets": {
                "table_fields": inspected["table_fields"],
                "placeholders": inspected["placeholders"],
                "paragraph_fields": inspected["paragraph_fields"],
                "sections": inspected["sections"],
                "images": inspected["images"],
            },
            "project_fields": inspected["project_fields"],
            "global_fields": GLOBAL_FIELDS,
            "stats": inspected["stats"],
            "qa": {
                "render_pdf": True,
                "compare_pages": True,
                "max_delta_ratio": 0.03,
                "pixel_delta_threshold": 0,
                "soffice": "soffice",
                "pdftoppm": "pdftoppm",
            },
        }
        write_yaml(profile_dir / "profile.yaml", profile_data)
        if project_id:
            self.init_project(profile, project_id)
        return profile_data

    def _prepare_reference_docx(self, source: Path, profile_dir: Path) -> Path:
        if source.suffix.lower() == ".docx":
            return source
        if source.suffix.lower() != ".doc":
            raise ValueError(f"Only .docx and .doc templates are supported: {source}")
        converted_dir = profile_dir / "_converted"
        converted_dir.mkdir(parents=True, exist_ok=True)
        soffice = shutil.which("soffice") or shutil.which("soffice.com")
        if not soffice:
            raise RuntimeError("LibreOffice soffice is required to import .doc templates.")
        with tempfile.TemporaryDirectory(prefix="template_to_doc_v2_lo_") as profile:
            completed = subprocess.run(
                [
                    soffice,
                    f"-env:UserInstallation={Path(profile).resolve().as_uri()}",
                    "--headless",
                    "--convert-to",
                    "docx",
                    "--outdir",
                    str(converted_dir),
                    str(source),
                ],
                text=True,
                capture_output=True,
                timeout=120,
                check=False,
            )
        converted = converted_dir / f"{source.stem}.docx"
        if completed.returncode != 0 or not converted.exists():
            detail = completed.stderr.strip() or completed.stdout.strip()
            raise RuntimeError(f"Failed to convert .doc template: {detail}")
        return converted

    def load_profile(self, profile_id: str) -> dict[str, Any]:
        profile_path = self.paths.profile_dir(profile_id) / "profile.yaml"
        profile = load_yaml(profile_path)
        changed = False
        if "global_fields" not in profile:
            profile["global_fields"] = GLOBAL_FIELDS
            changed = True
        targets = profile.setdefault("targets", {})
        paragraph_targets = targets.get("paragraph_fields")
        if "paragraph_fields" not in targets or (
            isinstance(paragraph_targets, list) and not paragraph_targets
        ):
            reference = self.paths.profile_dir(profile_id) / str(profile.get("template", "reference.docx"))
            if reference.exists():
                inspected = inspect_template(reference)
                refreshed_fields = inspected.get("paragraph_fields", [])
                if refreshed_fields or "paragraph_fields" not in targets:
                    targets["paragraph_fields"] = refreshed_fields
                    project_fields = profile.setdefault("project_fields", {})
                    for path, spec in inspected.get("project_fields", {}).items():
                        project_fields.setdefault(path, spec)
                    stats = profile.setdefault("stats", {})
                    stats["paragraph_field_count"] = len(targets["paragraph_fields"])
                    changed = True
        image_targets = targets.get("images")
        if "images" not in targets or (isinstance(image_targets, list) and not image_targets):
            reference = self.paths.profile_dir(profile_id) / str(profile.get("template", "reference.docx"))
            if reference.exists():
                inspected = inspect_template(reference)
                refreshed_images = inspected.get("images", [])
                if refreshed_images or "images" not in targets:
                    targets["images"] = refreshed_images
                    stats = profile.setdefault("stats", {})
                    stats["image_count"] = len(targets["images"])
                    changed = True
        sections = profile.get("targets", {}).get("sections", [])
        if isinstance(sections, list) and any(not item.get("template_sample") for item in sections if isinstance(item, dict)):
            reference = self.paths.profile_dir(profile_id) / str(profile.get("template", "reference.docx"))
            if reference.exists():
                inspected = inspect_template(reference)
                refreshed = {
                    str(item.get("id")): item
                    for item in inspected.get("sections", [])
                    if isinstance(item, dict)
                }
                for section in sections:
                    if not isinstance(section, dict) or section.get("template_sample"):
                        continue
                    replacement = refreshed.get(str(section.get("id")))
                    if not replacement:
                        continue
                    section["template_sample"] = replacement.get("template_sample", "")
                    section["sample_source"] = replacement.get("sample_source", "refreshed")
                    changed = True
        sections = targets.get("sections", [])
        if isinstance(sections, list):
            reference = self.paths.profile_dir(profile_id) / str(profile.get("template", "reference.docx"))
            if reference.exists():
                inspected = inspect_template(reference)
                refreshed_by_heading = {
                    int(item.get("heading_paragraph_index", -1)): item
                    for item in inspected.get("sections", [])
                    if isinstance(item, dict)
                }
                filtered_sections: list[dict[str, Any]] = []
                for section in sections:
                    if not isinstance(section, dict):
                        continue
                    try:
                        heading_index = int(section.get("heading_paragraph_index", -1))
                    except (TypeError, ValueError):
                        heading_index = -1
                    refreshed = refreshed_by_heading.get(heading_index)
                    if refreshed is None:
                        changed = True
                        continue
                    for key in (
                        "title",
                        "level",
                        "heading_paragraph_index",
                        "body_start_paragraph_index",
                        "body_end_paragraph_index",
                        "template_sample",
                        "sample_source",
                    ):
                        if section.get(key) != refreshed.get(key):
                            section[key] = refreshed.get(key)
                            changed = True
                    filtered_sections.append(section)
                if len(filtered_sections) != len(sections):
                    targets["sections"] = filtered_sections
                    stats = profile.setdefault("stats", {})
                    stats["section_target_count"] = len(filtered_sections)
                    changed = True
        if changed:
            write_yaml(profile_path, profile)
        return profile

    def profile_reference(self, profile_id: str) -> Path:
        profile = self.load_profile(profile_id)
        return self.paths.profile_dir(profile_id) / str(profile.get("template", "reference.docx"))

    def init_project(self, profile_id: str, project_id: str) -> dict[str, Any]:
        profile = self.load_profile(profile_id)
        project_dir = self.paths.project_dir(project_id)
        project_dir.mkdir(parents=True, exist_ok=True)
        project_path = project_dir / "project.yaml"
        if project_path.exists():
            return load_yaml(project_path)
        project: dict[str, Any] = {
            "project_id": project_id,
            "profile_id": profile_id,
            "global": {},
            "fields": {},
            "sections": {},
            "images": {},
        }
        for field in profile.get("global_fields") or GLOBAL_FIELDS:
            dotted_set(project, str(field.get("path")), field.get("default", ""))
        for path, spec in (profile.get("project_fields") or {}).items():
            dotted_set(project, str(path), spec.get("default", ""))
        for section in profile.get("targets", {}).get("sections", []):
            dotted_set(project, str(section.get("prompt_path")), "")
            dotted_set(project, str(section.get("path")), "")
        for image in profile.get("targets", {}).get("images", []):
            image_id = str(image.get("id") or "")
            if image_id:
                project["images"][image_id] = {"mode": "keep", "file": None}
        write_yaml(project_path, project)
        return project

    def generate_missing(
        self,
        profile_id: str,
        project: dict[str, Any],
        api_key: str,
        model: str = "deepseek-v4-flash",
    ) -> dict[str, Any]:
        profile = self.load_profile(profile_id)
        provider = deepseek_provider(api_key, model)
        context = {
            "global": project.get("global", {}),
            "fields": project.get("fields", {}),
            "project_id": project.get("project_id", ""),
            "profile_id": profile_id,
        }
        for section in profile.get("targets", {}).get("sections", []):
            path = str(section.get("path") or "")
            prompt_path = str(section.get("prompt_path") or "")
            existing = dotted_get(project, path, "")
            prompt = dotted_get(project, prompt_path, "")
            if existing:
                continue
            text = generate_local_text(
                prompt=str(prompt),
                template_sample=str(section.get("template_sample") or ""),
                target_label=str(section.get("title") or section.get("id") or ""),
                project_context=context,
                provider=provider,
            )
            dotted_set(project, path, text)
        return project

    def build(
        self,
        profile_id: str,
        project_id: str,
        project: dict[str, Any] | None = None,
        *,
        generate: bool = False,
        api_key: str = "",
        model: str = "deepseek-v4-flash",
        qa: bool = False,
    ) -> dict[str, Any]:
        profile = self.load_profile(profile_id)
        project_dir = self.paths.project_dir(project_id)
        project_dir.mkdir(parents=True, exist_ok=True)
        if project is None:
            project = self.init_project(profile_id, project_id)
        project["project_id"] = project_id
        project["profile_id"] = profile_id
        if generate:
            if not api_key:
                raise ValueError("DeepSeek API key is required when generate=True")
            project = self.generate_missing(profile_id, project, api_key, model)
        write_yaml(project_dir / "project.yaml", project)

        reference = self.profile_reference(profile_id)
        final_docx = project_dir / "final.docx"
        report = apply_replacements(reference, final_docx, profile, project)
        write_yaml(project_dir / "replacement_report.yaml", report.as_dict())
        outputs: dict[str, Any] = {
            "project_dir": str(project_dir),
            "final_docx": str(final_docx),
            "replacement_report": report.as_dict(),
        }
        if qa:
            outputs["qa"] = self.compare(profile_id, project_id)
        return outputs

    def compare(self, profile_id: str, project_id: str) -> dict[str, Any]:
        ensure_legacy_imports()
        from template_to_doc.qa import QAReport, RenderComparator

        profile = self.load_profile(profile_id)
        qa_config = profile.get("qa") or {}
        project_dir = self.paths.project_dir(project_id)
        final_docx = project_dir / "final.docx"
        reference = self.profile_reference(profile_id)
        render_dir = project_dir / "render"
        candidate_dir = render_dir / "candidate"
        reference_dir = render_dir / "reference"
        diff_dir = render_dir / "diff"
        renderer = RenderComparator(
            soffice=str(qa_config.get("soffice", "soffice")),
            pdftoppm=str(qa_config.get("pdftoppm", "pdftoppm")),
        )
        candidate = renderer.render_docx(final_docx, candidate_dir)
        reference_report = renderer.render_docx(reference, reference_dir)
        result: dict[str, Any] = {
            "candidate_render": _qa_report_dict(candidate),
            "reference_render": _qa_report_dict(reference_report),
        }
        if isinstance(candidate, QAReport) and isinstance(reference_report, QAReport):
            if candidate.ok and reference_report.ok:
                comparison = renderer.compare_rendered_pages(
                    candidate_dir,
                    reference_dir,
                    max_delta_ratio=float(qa_config.get("max_delta_ratio", 0.03)),
                    pixel_delta_threshold=int(qa_config.get("pixel_delta_threshold", 0)),
                    diff_dir=diff_dir,
                )
                result["compare"] = _qa_report_dict(comparison)
        write_yaml(project_dir / "qa_compare.yaml", result)
        return result


def _qa_report_dict(report: Any) -> dict[str, Any]:
    return {
        "ok": bool(getattr(report, "ok", False)),
        "messages": list(getattr(report, "messages", []) or []),
        "artifacts": dict(getattr(report, "artifacts", {}) or {}),
    }
