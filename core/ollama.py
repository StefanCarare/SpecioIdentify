from typing import Optional
import re
import requests


class OllamaError(Exception):
    """Eroare produsă în timpul comunicării cu Ollama."""
    pass


def check_ollama(ollama_url: str) -> bool:
    """
    Verifică dacă serverul Ollama este accesibil.
    """
    try:
        response = requests.get(
            f"{ollama_url}/api/tags",
            timeout=5
        )
        response.raise_for_status()
        return True

    except requests.RequestException:
        return False


def list_models(ollama_url: str) -> list[str]:
    """
    Returnează numele modelelor instalate pe serverul Ollama
    (/api/tags). Listă goală dacă serverul nu este accesibil
    sau răspunsul nu poate fi interpretat.
    """
    try:
        response = requests.get(
            f"{ollama_url}/api/tags",
            timeout=5
        )
        response.raise_for_status()
        data = response.json()

    except (requests.RequestException, ValueError):
        return []

    models = data.get("models") or []

    names = []
    for model in models:
        name = model.get("name")
        if name:
            names.append(name)

    return names


def generate(
    ollama_url: str,
    model: str,
    prompt: str,
    image_base64: Optional[str] = None,
    num_ctx: int = 8192,
    num_predict: int = 4000,
    thinking: bool = True,
) -> dict:
    """
    Trimite un request către Ollama /api/generate.

    Returnează răspunsul complet al API-ului sub forma unui dicționar.
    """
    
    payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        "think": thinking,
        "options": {
            "num_ctx": num_ctx,
            "num_predict": num_predict,
        },
    }    

    if image_base64:
        payload["images"] = [image_base64]

    try:
        response = requests.post(
            f"{ollama_url}/api/generate",
            json=payload,
            timeout=None,
        )

        response.raise_for_status()

    except requests.RequestException as exc:
        details = response.text.strip()
        if (
            response.status_code == 500
            and "CUDA" in details.upper()
            and "num_gpu" not in payload["options"]
        ):
            payload["options"]["num_gpu"] = 0
            try:
                response = requests.post(
                    f"{ollama_url}/api/generate",
                    json=payload,
                    timeout=None,
                )
                response.raise_for_status()
            except requests.RequestException as retry_exc:
                retry_details = response.text.strip()
                if retry_details:
                    retry_details = f"; detalii Ollama: {retry_details[:1000]}"
                raise OllamaError(
                    "GPU și fallback-ul CPU Ollama au eșuat: "
                    f"{retry_exc}{retry_details}"
                ) from retry_exc
        else:
            if details:
                details = f"; detalii Ollama: {details[:1000]}"
            raise OllamaError(
                f"Nu am putut comunica cu Ollama: {exc}{details}"
            ) from exc

    try:
        data = response.json()

    except ValueError as exc:
        raise OllamaError(
            "Ollama a returnat un răspuns care nu este JSON valid."
        ) from exc

    if "error" in data:
        raise OllamaError(
            f"Ollama a raportat o eroare: {data['error']}"
        )

    return data


def get_response(data: dict) -> str:
    """
    Extrage textul generat de model.
    """
    response = data.get("response")
    if isinstance(response, str) and response.strip():
        return _clean_protocol_text(response)

    message = data.get("message")
    if isinstance(message, dict):
        content = message.get("content", "")
        if isinstance(content, str):
            return _clean_protocol_text(content)

    return ""


def _clean_protocol_text(text: str) -> str:
    """Removes tool/protocol routing tokens that are not user-facing content."""
    text = re.sub(r"(?im)^\s*(?:to|from|recipient)\s*=\s*[^\n]+\s*$", "", text)
    text = re.sub(r"(?im)^\s*(?:to|from|recipient)\s*:\s*[^\n]+\s*$", "", text)
    return text.strip()


def get_thinking(data: dict) -> str:
    """
    Extrage partea de reasoning/thinking, dacă este disponibilă.
    """
    return data.get("thinking", "")


def format_duration(nanoseconds: int | None) -> str:
    """
    Transformă o durată exprimată în nanosecunde
    într-o formă ușor de citit.
    """

    if nanoseconds is None:
        return ""

    total_seconds = nanoseconds / 1_000_000_000

    hours = int(total_seconds // 3600)
    minutes = int((total_seconds % 3600) // 60)
    seconds = total_seconds % 60

    if hours > 0:
        return f"{hours}h {minutes}m {seconds:.1f}s"

    if minutes > 0:
        return f"{minutes}m {seconds:.1f}s"

    return f"{seconds:.1f}s"


def get_statistics(data: dict) -> dict:
    """
    Extrage statisticile returnate de Ollama și
    adaugă variantele ușor de citit pentru durate.
    """

    statistics = {
        "total_duration": data.get("total_duration"),
        "load_duration": data.get("load_duration"),
        "prompt_eval_count": data.get("prompt_eval_count"),
        "prompt_eval_duration": data.get("prompt_eval_duration"),
        "eval_count": data.get("eval_count"),
        "eval_duration": data.get("eval_duration"),
    }

    # Variante lizibile
    statistics["total_duration_human"] = format_duration(
        data.get("total_duration")
    )

    statistics["load_duration_human"] = format_duration(
        data.get("load_duration")
    )

    statistics["prompt_eval_duration_human"] = format_duration(
        data.get("prompt_eval_duration")
    )

    statistics["eval_duration_human"] = format_duration(
        data.get("eval_duration")
    )

    return statistics