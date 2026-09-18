from pathlib import Path


def save_result(
    output_path: Path,
    response: str,
    thinking: str = "",
    image_name: str = "",
    prompt: str = "",
    model: str = "",
    statistics: dict | None = None,
    image_metadata: dict | None = None,
    manual_location: str = "",
    preset_name: str = "",
) -> Path:
    """
    Salvează rezultatul analizei într-un fișier text.

    Păstrează:
    - imaginea analizată
    - metadata fotografiei
    - promptul utilizatorului
    - răspunsul modelului
    - thinking/reasoning, dacă există
    - statisticile Ollama
    """

    output_path = Path(output_path)
    output_path.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    image_metadata = image_metadata or {}

    photo_date = image_metadata.get(
        "photo_date"
    )

    file_created = image_metadata.get(
        "file_created"
    )

    file_modified = image_metadata.get(
        "file_modified"
    )

    camera_make = image_metadata.get(
        "camera_make"
    )

    camera_model = image_metadata.get(
        "camera_model"
    )

    latitude = image_metadata.get(
        "latitude"
    )

    longitude = image_metadata.get(
        "longitude"
    )

    lines = [
        "========================================",
        "Specio Identify - ANALYSIS RESULT",
        "========================================",
        "",
        f"Image: {image_name}",
    ]

    # --------------------------------------------------------
    # File information
    # --------------------------------------------------------

    if file_created:
        lines.append(
            f"File created: {file_created}"
        )

    if file_modified:
        lines.append(
            f"File modified: {file_modified}"
        )

    lines.extend([
        f"Model: {model}",
    ])

    if preset_name:
        lines.append(
            f"Preset: {preset_name}"
        )

    lines.extend([
        "",
        "=== PHOTO CONTEXT ===",
    ])

    # --------------------------------------------------------
    # Photo context
    # --------------------------------------------------------

    if photo_date:
        lines.append(
            f"Photo date: {photo_date}"
        )

    if camera_make:
        lines.append(
            f"Camera make: {camera_make}"
        )

    if camera_model:
        lines.append(
            f"Camera model: {camera_model}"
        )

    if latitude is not None and longitude is not None:
        lines.append(
            f"GPS: {latitude}, {longitude}"
        )

    if manual_location:
        lines.append(
            f"Location: {manual_location} (manual)"
        )

    if not any([
        photo_date,
        camera_make,
        camera_model,
        latitude is not None and longitude is not None,
    ]):
        lines.append(
            "No photo metadata available."
        )

    lines.extend([
        "",
        "=== PROMPT ===",
        prompt.strip(),
        "",
        "=== RESPONSE ===",
        response.strip(),
        "",
    ])

    # --------------------------------------------------------
    # Thinking
    # --------------------------------------------------------

    if thinking:
        lines.extend([
            "=== THINKING ===",
            thinking.strip(),
            "",
        ])

    # --------------------------------------------------------
    # Statistics
    # --------------------------------------------------------

    if statistics:
        lines.append(
            "=== STATISTICS ==="
        )

        for key, value in statistics.items():

            if value is not None:
                lines.append(
                    f"{key}: {value}"
                )

        lines.append("")

    output_path.write_text(
        "\n".join(lines),
        encoding="utf-8",
    )

    return output_path

def save_text(
    output_path: Path,
    text: str,
) -> Path:
    """
    Salvează simplu un text într-un fișier UTF-8.
    Util pentru salvarea manuală din GUI.
    """

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    output_path.write_text(
        text,
        encoding="utf-8",
    )

    return output_path


def create_output_path(
    output_dir: Path,
    image_path: Path,
    suffix: str = "_description",
) -> Path:
    """
    Creează automat numele fișierului de output pornind
    de la numele imaginii.

    Exemplu:
        floare_005.jpg
        →
        floare_005_description.txt
    """

    output_dir = Path(output_dir)
    image_path = Path(image_path)

    filename = f"{image_path.stem}{suffix}.txt"

    return output_dir / filename