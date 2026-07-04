from __future__ import annotations

from pathlib import Path
import sys
from typing import Optional

if __package__ in {None, ""}:  # pragma: no cover - direct script use
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import typer
from rich.console import Console

from template_to_doc_v2.pipeline import TemplateToDocV2Pipeline
from template_to_doc_v2.yaml_io import load_yaml


app = typer.Typer(help="TemplateToDocV2 template-preserving agent.")
console = Console()


def _pipeline(root: Path) -> TemplateToDocV2Pipeline:
    return TemplateToDocV2Pipeline(root)


@app.command("import-template")
def import_template(
    template: Path = typer.Argument(..., exists=True, help="Word .docx/.doc template."),
    profile: Optional[str] = typer.Option(None, "--profile", "-p"),
    project_id: Optional[str] = typer.Option(None, "--project-id"),
    root: Path = typer.Option(Path("."), "--root"),
) -> None:
    result = _pipeline(root).import_template(template, profile, project_id)
    console.print(f"profile: {result['profile_id']}")
    console.print(f"targets: {result['stats']}")


@app.command("init-project")
def init_project(
    profile: str = typer.Argument(...),
    project_id: str = typer.Argument(...),
    root: Path = typer.Option(Path("."), "--root"),
) -> None:
    project = _pipeline(root).init_project(profile, project_id)
    console.print(f"project: {project_id}")
    console.print(root / "projects" / project_id / "project.yaml")
    console.print(project)


@app.command("build")
def build(
    profile: str = typer.Argument(...),
    project_id: str = typer.Argument(...),
    project_yaml: Optional[Path] = typer.Option(None, "--project-yaml", exists=True),
    root: Path = typer.Option(Path("."), "--root"),
    generate: bool = typer.Option(False, "--generate/--no-generate"),
    api_key: str = typer.Option("", "--api-key", envvar="DEEPSEEK_API_KEY"),
    model: str = typer.Option("deepseek-v4-flash", "--model"),
    qa: bool = typer.Option(False, "--qa/--no-qa"),
) -> None:
    project = load_yaml(project_yaml) if project_yaml else None
    outputs = _pipeline(root).build(
        profile,
        project_id,
        project=project,
        generate=generate,
        api_key=api_key,
        model=model,
        qa=qa,
    )
    console.print(outputs)


@app.command("compare")
def compare(
    profile: str = typer.Argument(...),
    project_id: str = typer.Argument(...),
    root: Path = typer.Option(Path("."), "--root"),
) -> None:
    console.print(_pipeline(root).compare(profile, project_id))


@app.command("serve")
def serve(
    root: Path = typer.Option(Path("."), "--root"),
    host: str = typer.Option("127.0.0.1", "--host"),
    port: int = typer.Option(8770, "--port"),
) -> None:
    from template_to_doc_v2.web import serve as serve_web

    console.print(f"TemplateToDocV2 web: http://{host}:{port}/")
    serve_web(root=root, host=host, port=port)


if __name__ == "__main__":
    app()
