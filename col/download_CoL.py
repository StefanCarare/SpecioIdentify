import requests

url_base = "https://download.checklistbank.org/col/latest_coldp.zip"
url_extended = "https://download.checklistbank.org/col/xr_latest_coldp.zip"
url = url_extended
output_file = "catalogue_of_life.zip"

response = requests.get(url, stream=True, timeout=60)
response.raise_for_status()

with open(output_file, "wb") as f:
    for chunk in response.iter_content(chunk_size=1024 * 1024):
        if chunk:
            f.write(chunk)

print(f"Arhiva a fost salvată în: {output_file}")