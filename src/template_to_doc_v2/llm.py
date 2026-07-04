from __future__ import annotations

import json
import re
from typing import Any

from .compat import ensure_legacy_imports


DEEPSEEK_BASE_URL = "https://api.deepseek.com"
DEEPSEEK_MODELS = {"deepseek-v4-flash", "deepseek-v4-pro"}


def deepseek_provider(api_key: str, model: str = "deepseek-v4-flash", timeout: float = 90.0):
    ensure_legacy_imports()
    from template_to_doc.llm.http_chat import OpenAICompatibleProvider

    if model not in DEEPSEEK_MODELS:
        raise ValueError(f"Unsupported DeepSeek model: {model}")
    return OpenAICompatibleProvider(
        model=model,
        base_url=DEEPSEEK_BASE_URL,
        api_key=api_key,
        api_key_env="",
        max_tokens=3500,
        timeout=timeout,
    )


def _compact_generated_text(value: object) -> str:
    text = re.sub(r"\s+", "", str(value or ""))
    return re.sub(r"[，。；：、,.!！?？;:]+$", "", text)


def _looks_like_unexpanded_prompt(generated: object, prompt: object) -> bool:
    generated_text = str(generated or "").strip()
    prompt_text = str(prompt or "").strip()
    if not generated_text:
        return True
    compact_generated = _compact_generated_text(generated_text)
    compact_prompt = _compact_generated_text(prompt_text)
    if not compact_prompt:
        return False
    if compact_generated == compact_prompt:
        return True
    is_near_copy = compact_generated in compact_prompt or compact_prompt in compact_generated
    return is_near_copy and len(compact_generated) <= len(compact_prompt) + 12


def _looks_underexpanded(generated: object, template_sample: object, target_label: object) -> bool:
    compact_generated = _compact_generated_text(generated)
    compact_template = _compact_generated_text(template_sample)
    compact_label = _compact_generated_text(target_label)
    if not compact_generated:
        return True
    if compact_label and compact_generated in {compact_label, compact_label.rstrip("；;")}:
        return True
    if len(compact_template) >= 120 and len(compact_generated) < 70:
        return True
    return False


def _template_metrics(template_sample: str) -> dict[str, int]:
    text = str(template_sample or "").replace("\r\n", "\n").replace("\r", "\n").strip()
    paragraphs = [part for part in re.split(r"\n\s*\n", text) if part.strip()]
    nonempty_lines = [line for line in text.split("\n") if line.strip()]
    list_like_lines = [
        line
        for line in nonempty_lines
        if re.match(r"^\s*(?:[-*•]|[0-9]+[\.、．)]|[a-zA-Z][\.、．)])\s*", line)
    ]
    return {
        "char_count": len(text),
        "paragraph_count": len(paragraphs) or (1 if text else 0),
        "nonempty_line_count": len(nonempty_lines),
        "list_like_line_count": len(list_like_lines),
    }


def generate_local_text(
    *,
    prompt: str,
    template_sample: str,
    target_label: str,
    project_context: dict[str, Any],
    provider: Any,
    attachments: list[dict[str, Any]] | None = None,
) -> str:
    metrics = _template_metrics(template_sample)
    payload = {
        "task": (
            "仿照 template_reference 中的模板原文内容，"
            "结合 local_prompt 和 project_context.global，生成可直接写回当前 Word 局部位置的正文。"
        ),
        "target": {
            "label": target_label,
            "local_prompt": prompt,
            "template_reference": template_sample,
            "template_metrics": metrics,
        },
        "project_context": project_context,
        "attachments": attachments or [],
        "instructions": [
            "只返回可以直接写入该 Word 局部位置的正文。",
            "不要解释生成过程，不要包裹代码块，不要输出字段名。",
            "template_reference 是右侧模板内容，请把它当作写作范例。",
            "模仿 template_reference 的结构、行文风格、详略程度、术语密度和表格/列表表达。",
            "严格控制篇幅：生成内容的字符数、段落数、列表项数量应接近 target.template_metrics；除非用户明确要求，不要超过模板样例约 20%。",
            "local_prompt 是本节补充素材或写作要求，不是最终答案；不要把 local_prompt 原样保留。",
            "如果 local_prompt 为空，必须根据 project_context.global 和 template_reference 主动生成，不要追问。",
            "project_context.global 是全文共同背景、主题、原始素材和写作要求。",
            "生成结果必须围绕新项目主题展开，并替换模板中的原项目主题和对象。",
            "不要只输出 target.label、章节标题、目录项或一句话概要；应输出与 template_reference 详略程度相匹配的正文。",
            "如果 template_reference 包含标题、段落、表格或列表，请优先保持相同表达形态。",
            "保持 template_reference 的主要语言；不要主动新增模板、local_prompt、project_context 或附件中没有出现过的英文机构名、标准名、缩写或专有名词。",
            "不要输出 Markdown 标题符号、代码块、额外解释，除非 template_reference 本身就是这种表达形态。",
            "如果附件文本与提示词冲突，以用户提示词和全局信息为准，并保持与模板结构一致。",
        ],
    }
    system = (
        "你是模板化 Word 文档的局部写作助手。"
        "你的输出会直接覆盖网页输入框，并写回原 Word 模板。"
        "你要仿照模板原文的内容组织方式，以用户资料和全局信息为新主题仿写。"
        "必须只输出正文。"
    )
    prompt_json = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    generated = str(provider.generate(prompt_json, system)).strip()
    if _looks_like_unexpanded_prompt(generated, prompt) or _looks_underexpanded(generated, template_sample, target_label):
        retry_payload = dict(payload)
        retry_payload["retry_reason"] = (
            "上一次输出过短、只像标题/提示词，或没有按模板样例完成仿写扩写。"
        )
        retry_payload["previous_output"] = generated
        retry_payload["instructions"] = [
            "必须重新生成，不得只复述标题、local_prompt 或一句话概要。",
            "请仿照 template_reference 的原文内容，围绕新项目主题生成可直接写入该框的新正文。",
            *payload["instructions"],
        ]
        generated = str(
            provider.generate(json.dumps(retry_payload, ensure_ascii=False, indent=2, default=str), system)
        ).strip()
    if _looks_like_unexpanded_prompt(generated, prompt) or _looks_underexpanded(generated, template_sample, target_label):
        raise ValueError("模型返回内容仍过短或与提示词/标题过于接近，请补充全局信息或重新生成")
    return generated
