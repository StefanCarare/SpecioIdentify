from pathlib import Path
import base64
from datetime import datetime
from PIL import Image, ExifTags
import xml.etree.ElementTree as ET
import piexif


def image_exists(image_path: Path) -> bool:
    """Verifică dacă fișierul imagine există."""
    return image_path.is_file()


def image_to_base64(image_path: Path) -> str:
    """
    Citește o imagine și o transformă în Base64,
    formatul folosit de Ollama pentru input vizual.
    """
    if not image_exists(image_path):
        raise FileNotFoundError(
            f"Imaginea nu există: {image_path}"
        )

    with image_path.open("rb") as f:
        image_bytes = f.read()

    return base64.b64encode(image_bytes).decode("utf-8")


def get_image_path(app_root: Path, relative_path: str) -> Path:
    """
    Construiește calea completă către o imagine
    pornind de la rădăcina aplicației.
    """
    return (app_root / relative_path).resolve()


def extract_exif_with_piexif(image_path: Path) -> dict:
    """
    Extrage EXIF metadata folosind piexif pentru suport mai bun.
    piexif poate accesa tag-uri EXIF care nu sunt accesibile de PIL.
    """
    exif_data = {}

    try:
        exif_dict = piexif.load(str(image_path))
        
        if "Exif" in exif_dict:
            exif = exif_dict["Exif"]
            
            # Look for date fields in EXIF using numeric tag IDs
            # DateTimeOriginal = 36867 (0x9003)
            # DateTimeDigitized = 36868 (0x9004)
            # DateTime = 306 (0x0132)
            date_fields = {
                36867: "DateTimeOriginal",
                36868: "DateTimeDigitized",
                306: "DateTime",
            }
            
            for tag_id, tag_name in date_fields.items():
                if tag_id in exif:
                    value = exif[tag_id]
                    exif_data[tag_name] = value.decode('utf-8') if isinstance(value, bytes) else value
                    
    except Exception:
        pass

    return exif_data


def extract_xmp_metadata(image_path: Path) -> dict:
    """
    Extrage metadata XMP din imagine.
    XMP poate conține date precum DateTaken care nu sunt în EXIF standard.
    """
    xmp_data = {}

    try:
        with Image.open(image_path) as image:
            # XMP metadata is often stored in APP1 segment
            if hasattr(image, 'info'):
                for key, value in image.info.items():
                    if key.lower() in ('xmlpacket', 'xmp', 'xmpdata'):
                        try:
                            # Parse XML
                            root = ET.fromstring(value)
                            
                            # Look for date fields in various XMP namespaces
                            namespaces = {
                                'xmp': 'http://ns.adobe.com/xap/1.0/',
                                'xmpexif': 'http://ns.adobe.com/exif/1.0/',
                                'photoshop': 'http://ns.adobe.com/photoshop/1.0/',
                                'dc': 'http://purl.org/dc/elements/1.1/',
                            }
                            
                            # Try various date field names
                            date_fields = [
                                ('xmp:CreateDate', 'xmp'),
                                ('xmp:ModifyDate', 'xmp'),
                                ('xmp:Date', 'xmp'),
                                ('xmpexif:DateTimeOriginal', 'xmpexif'),
                                ('photoshop:DateCreated', 'photoshop'),
                                ('dc:date', 'dc'),
                            ]
                            
                            for field, ns in date_fields:
                                try:
                                    elem = root.find(f'.//{{{namespaces[ns]}}}{field}')
                                    if elem is not None and elem.text:
                                        xmp_data['photo_date'] = elem.text
                                        break
                                except Exception:
                                    continue
                                    
                        except Exception:
                            pass
                            
    except Exception:
        pass

    return xmp_data
    

def get_image_metadata(image_path: Path) -> dict:
    """
    Extrage metadata relevantă a imaginii.

    Prioritate pentru data fotografierii:
    EXIF DateTimeOriginal -> EXIF DateTime -> data modificării fișierului.

    Returnează și datele filesystem-ului ca informații auxiliare.
    """

    image_path = Path(image_path)

    if not image_path.is_file():
        raise FileNotFoundError(
            f"Imaginea nu există: {image_path}"
        )

    metadata = {
        "photo_date": None,
        "file_created": None,
        "file_modified": None,
        "camera_make": None,
        "camera_model": None,
        "latitude": None,
        "longitude": None,
    }

    # Datele fișierului din Windows/filesystem
    stat = image_path.stat()

    metadata["file_created"] = datetime.fromtimestamp(
        stat.st_ctime
    ).strftime("%Y-%m-%d %H:%M:%S")

    metadata["file_modified"] = datetime.fromtimestamp(
        stat.st_mtime
    ).strftime("%Y-%m-%d %H:%M:%S")

    # EXIF
    try:
        with Image.open(image_path) as image:
            exif = image.getexif()

            if not exif:
                # Try XMP if no EXIF
                xmp_data = extract_xmp_metadata(image_path)
                if xmp_data.get('photo_date'):
                    metadata["photo_date"] = xmp_data['photo_date']
                return metadata

            exif_data = {}

            for tag_id, value in exif.items():
                tag_name = ExifTags.TAGS.get(tag_id, tag_id)
                exif_data[tag_name] = value

            # GPS
            gps_data = exif_data.get("GPSInfo")

            if gps_data:
                try:
                    gps_latitude = gps_data.get(
                        ExifTags.GPSIFD.GPSLatitude
                    )

                    gps_latitude_ref = gps_data.get(
                        ExifTags.GPSIFD.GPSLatitudeRef
                    )

                    gps_longitude = gps_data.get(
                        ExifTags.GPSIFD.GPSLongitude
                    )

                    gps_longitude_ref = gps_data.get(
                        ExifTags.GPSIFD.GPSLongitudeRef
                    )

                    if (
                        gps_latitude
                        and gps_latitude_ref
                        and gps_longitude
                        and gps_longitude_ref
                    ):
                        metadata["latitude"] = gps_to_decimal(
                            gps_latitude,
                            gps_latitude_ref
                        )

                        metadata["longitude"] = gps_to_decimal(
                            gps_longitude,
                            gps_longitude_ref
                        )

                except Exception:
                    pass

            # Data reală a fotografierii - prioritar
            photo_date = (
                exif_data.get("DateTimeOriginal")
                or exif_data.get("DateTimeDigitized")
                or exif_data.get("DateTime")
            )

            # If EXIF doesn't have DateTimeOriginal or DateTimeDigitized, try piexif
            if not exif_data.get("DateTimeOriginal") and not exif_data.get("DateTimeDigitized"):
                piexif_data = extract_exif_with_piexif(image_path)
                if piexif_data.get("DateTimeOriginal"):
                    photo_date = piexif_data["DateTimeOriginal"]
                elif piexif_data.get("DateTimeDigitized"):
                    photo_date = piexif_data["DateTimeDigitized"]
                else:
                    # Try XMP as fallback
                    xmp_data = extract_xmp_metadata(image_path)
                    if xmp_data.get('photo_date'):
                        photo_date = xmp_data['photo_date']

            if photo_date:
                try:
                    # Try EXIF format first
                    dt = datetime.strptime(
                        photo_date,
                        "%Y:%m:%d %H:%M:%S"
                    )
                    metadata["photo_date"] = dt.strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                except ValueError:
                    try:
                        # Try ISO 8601 format (XMP)
                        dt = datetime.fromisoformat(
                            photo_date.replace('Z', '+00:00')
                        )
                        metadata["photo_date"] = dt.strftime(
                            "%Y-%m-%d %H:%M:%S"
                        )
                    except (ValueError, TypeError):
                        metadata["photo_date"] = str(photo_date)

            metadata["camera_make"] = exif_data.get("Make")
            metadata["camera_model"] = exif_data.get("Model")

    except Exception:
        # Dacă imaginea nu are EXIF sau metadata nu poate fi citită,
        # păstrăm informațiile filesystem disponibile.
        pass

    # Dacă nu există EXIF, folosim data modificării ca fallback.
    if metadata["photo_date"] is None:
        metadata["photo_date"] = metadata["file_modified"]

    return metadata    
    
def gps_to_decimal(value, ref: str) -> float | None:
    """
    Transformă o coordonată GPS EXIF din grade/minute/secunde
    în grade zecimale.
    """

    try:
        degrees = float(value[0])
        minutes = float(value[1])
        seconds = float(value[2])

        decimal = degrees + (
            minutes / 60
        ) + (
            seconds / 3600
        )

        if ref in ("S", "W"):
            decimal = -decimal

        return decimal

    except (TypeError, ValueError, IndexError, ZeroDivisionError):
        return None
        
        
def gps_to_decimal(value, ref: str) -> float | None:
    """
    Transformă o coordonată GPS EXIF din
    grade/minute/secunde în grade zecimale.
    """

    try:
        degrees = float(value[0])
        minutes = float(value[1])
        seconds = float(value[2])

        decimal = (
            degrees
            + minutes / 60
            + seconds / 3600
        )

        if ref in ("S", "W"):
            decimal = -decimal

        return decimal

    except (
        TypeError,
        ValueError,
        IndexError,
        ZeroDivisionError,
    ):
        return None        
        