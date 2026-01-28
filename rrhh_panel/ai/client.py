from __future__ import annotations

import httpx
from zai import ZaiClient

from rrhh_panel.utils.secrets import get_secret


def get_zai_client() -> ZaiClient:
    api_key = get_secret("ZAI_API_KEY")
    base_url = get_secret("ZAI_BASE_URL", "https://api.z.ai/api/paas/v4/")

    if not api_key:
        raise RuntimeError("Falta ZAI_API_KEY. Configúrala en Streamlit Cloud → Secrets.")

    return ZaiClient(
        api_key=api_key,
        base_url=base_url,
        timeout=httpx.Timeout(timeout=120.0, connect=8.0),
        max_retries=2,
    )


def chat_stream(
    messages,
    model: str = "glm-4.7-flash",
    stream: bool = True,
    temperature: float = 0.2,
    max_tokens: int = 1200,
):
    client = get_zai_client()
    return client.chat.completions.create(
        model=model,
        messages=messages,
        stream=stream,
        temperature=temperature,
        max_tokens=max_tokens,
    )
