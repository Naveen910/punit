"""
Refactored helpers for MAIN_LIMS_CSV.py
- Centralises LIMS CSV export logic into `export_to_lims_csv` and `export_all_presets`.
- Provides the `Client` upload class.
- Includes an optional integration helper `replace_export_blocks_in_main()` which can
  programmatically replace repeated LIMS export blocks in the original file with a
  single call to `export_all_presets()`.

Usage:
    # Inspect, then run the integrator (it will rewrite MAIN_LIMS_CSV.py if you run it)
    from MAIN_LIMS_CSV_refactored import replace_export_blocks_in_main
    replace_export_blocks_in_main("/path/to/MAIN_LIMS_CSV.py")

Note: The integration helper performs textual replacement—please review and commit
or back up the original file before running it.
"""

import re
import os
import csv
import json
import mimetypes
import uuid
import shutil
import requests
from datetime import datetime


def export_to_lims_csv(record, lims_folder=r"C:\LIMS_UPLOAD"):
    """Export a single test record as a CSV for LIMS pickup.

    This mirrors the behaviour added to the original file but centralises it here.
    """
    try:
        if not os.path.exists(lims_folder):
            os.makedirs(lims_folder)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"LIMS_{record.get('Test_ID', 'NA')}_{timestamp}.csv"
        dest_path = os.path.join(lims_folder, file_name)

        with open(dest_path, "w", newline="") as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=list(record.keys()))
            writer.writeheader()
            writer.writerow(record)

        print(f"LIMS CSV exported: {dest_path}")
        return True
    except Exception as e:
        print(f"LIMS CSV export failed: {e}")
        return False


def export_all_presets(globals_map):
    """Export P1..P6 preset records to LIMS using values from the provided globals map.

    The original code uses many global variables; this helper accepts a mapping
    (for example, the original module's globals()) so it doesn't depend on importing
    those names directly.
    """
    # Expected keys used by the export logic
    keys = [
        'test_id', 'date', 'time1', 'P1_FileName', 'P1_Drop_Temp', 'P1_Drop_Point',
        'P2_FileName', 'P2_Drop_Temp', 'P2_Drop_Point',
        'P3_FileName', 'P3_Drop_Temp', 'P3_Drop_Point',
        'P4_FileName', 'P4_Drop_Temp', 'P4_Drop_Point',
        'P5_FileName', 'P5_Drop_Temp', 'P5_Drop_Point',
        'P6_FileName', 'P6_Drop_Temp', 'P6_Drop_Point',
        'Block_float'
    ]

    missing = [k for k in keys if k not in globals_map]
    if missing:
        raise KeyError(f"Missing globals required for export_all_presets: {missing}")

    Block_float = globals_map.get('Block_float')

    presets = [
        ('P1_FileName', 'P1_Drop_Temp', 'P1_Drop_Point'),
        ('P2_FileName', 'P2_Drop_Temp', 'P2_Drop_Point'),
        ('P3_FileName', 'P3_Drop_Temp', 'P3_Drop_Point'),
        ('P4_FileName', 'P4_Drop_Temp', 'P4_Drop_Point'),
        ('P5_FileName', 'P5_Drop_Temp', 'P5_Drop_Point'),
        ('P6_FileName', 'P6_Drop_Temp', 'P6_Drop_Point'),
    ]

    for fname_k, dtemp_k, dpoint_k in presets:
        record = {
            'Test_ID': globals_map.get('test_id'),
            'Date': globals_map.get('date'),
            'Time': globals_map.get('time1'),
            'File_Name': globals_map.get(fname_k),
            'Drop_Temp': globals_map.get(dtemp_k),
            'Block_Temp': round(Block_float) if isinstance(Block_float, (int, float)) else 0,
            'Drop_Point': round(globals_map.get(dpoint_k)) if isinstance(globals_map.get(dpoint_k), (int, float)) else 0,
            'Instrument_ID': 'DP-01',
            'Operator': os.getenv('USERNAME', 'Unknown'),
            'Upload_Timestamp': datetime.now().isoformat()
        }
        export_to_lims_csv(record)


class Client:
    """HTTP client for uploading files to the LIMS-like endpoint.

    This is the same logic ported from the original file, placed here for reuse.
    """
    def __init__(self, server_url):
        self.is_connected = False
        self.server_url = server_url
        self.token = None

    def login(self, username, password):
        response = requests.post(f"{self.server_url}/login", json={'username': username, 'password': password})
        if response.status_code == 200:
            self.token = response.json().get('token')
            return True
        return False

    def is_alive(self):
        try:
            response = requests.get(f"{self.server_url}/is_alive")
            return response.status_code == 200
        except Exception:
            return False

    def upload_file(self, file_name, data):
        # simplified behaviour: keep the same return codes as original
        try:
            content_type, _ = mimetypes.guess_type(file_name)
            data['Content-Type'] = content_type
            data = json.loads(json.dumps(data))

            with open(file_name, 'rb') as fh:
                file_content = fh.read()

            if content_type is None:
                content_type = 'application/octet-stream'

            boundary = str(uuid.uuid4())
            headers = {
                'Authorization': f'Bearer {self.token}',
                'Content-Type': f'multipart/form-data; boundary={boundary}'
            }

            only_file_name = os.path.basename(file_name)
            str_payload = (
                f"--{boundary}\r\n"
                f"Content-Disposition: form-data; name=\"data\"\r\n\r\n"
                f"{data}\r\n"
                f"--{boundary}\r\n"
                f"Content-Disposition: form-data; name=\"file\"; filename=\"{only_file_name}\"\r\n"
                f"Content-Type: {content_type}\r\n\r\n"
            )
            payload = str_payload.encode('utf-8') + file_content + f"\r\n--{boundary}--\r\n".encode('utf-8')

            response = requests.post(f"{self.server_url}/attachments", data=payload, headers=headers)
            if response.status_code == 200:
                return 1
            return -1
        except Exception as e:
            print(f"upload_file error: {e}")
            return -1


# Integration helper: textual replacement of repeated export blocks in the original file.
EXPORT_BLOCK_RE = re.compile(r"\n\s*# --- LIMS CSV Export ---(?:.|\n)*?# ------------------------\n", re.MULTILINE)


def replace_export_blocks_in_main(main_path):
    """Replace repeated LIMS export blocks in `main_path` with a single call to
    `export_all_presets(globals())`.

    This function edits the file in-place. It makes a backup next to the original
    as `<filename>.bak` before writing.
    """
    if not os.path.isfile(main_path):
        raise FileNotFoundError(main_path)

    with open(main_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Count original occurrences
    matches = list(EXPORT_BLOCK_RE.finditer(content))
    if not matches:
        print('No export blocks found to replace.')
        return 0

    # Replace each contiguous group of repeated export blocks with a single call.
    new_content = EXPORT_BLOCK_RE.sub('\n    try:\n        export_all_presets(globals())\n    except Exception:\n        pass\n', content)

    backup_path = main_path + '.bak'
    with open(backup_path, 'w', encoding='utf-8') as bf:
        bf.write(content)

    with open(main_path, 'w', encoding='utf-8') as f:
        f.write(new_content)

    print(f'Replaced {len(matches)} export block(s). Backup written to {backup_path}')
    return len(matches)


if __name__ == '__main__':
    print('This module provides helpers for MAIN_LIMS_CSV.py. Import and use the functions.')
