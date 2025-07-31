import os
import sys
import json
import shutil
import datetime
from PIL import Image
import piexif
import time
import re
import platform
import glob
import zipfile

LOG_LINES = []

def log(msg):
    print(msg)
    LOG_LINES.append(msg)

def print_bar(i: int, l: int, n_bars=50) -> None:
    n_pipes = int((i / l) * n_bars)
    bar = '|' * n_pipes + '-' * (n_bars - n_pipes)
    print(f'\r\t{bar} ({i}/{l})', end='', flush=True)

def deg_to_dms_rational(deg_float):
    deg_abs = abs(deg_float)
    minutes, seconds = divmod(deg_abs * 3600, 60)
    degrees, minutes = divmod(minutes, 60)
    return [
        (int(degrees), 1),
        (int(minutes), 1),
        (int(seconds * 100), 100)
    ]

WINDOWS = platform.system().lower().startswith('win')
if WINDOWS:
    try:
        import pywintypes
        import win32file
        import win32con
    except ImportError:
        print('\n[ERROR] pywin32 is required for setting creation time on Windows.\nRun: pip install pywin32\n')
        sys.exit(1)

def set_file_creation_time(path, timestamp):
    if not WINDOWS:
        return
    try:
        local_time = time.localtime(timestamp)
        wintime = pywintypes.Time(time.mktime(local_time))
        filehandle = win32file.CreateFile(
            path, win32con.GENERIC_WRITE,
            win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
            None, win32con.OPEN_EXISTING, win32con.FILE_ATTRIBUTE_NORMAL, None)
        win32file.SetFileTime(filehandle, wintime, wintime, wintime)
        filehandle.close()
    except Exception:
        pass

def set_exif_jpeg(file_path, json_data, people):
    try:
        img = Image.open(file_path)
        exif_dict = {"0th":{}, "Exif":{}, "GPS":{}, "1st":{}, "thumbnail":None}
        try:
            dt = datetime.datetime.utcfromtimestamp(int(json_data['photoTakenTime']['timestamp']))
            exif_dict['Exif'][piexif.ExifIFD.DateTimeOriginal] = dt.strftime("%Y:%m:%d %H:%M:%S")
        except Exception:
            pass
        gps = json_data.get('geoData') or {}
        if gps.get('latitude') and gps.get('longitude'):
            try:
                exif_dict['GPS'][piexif.GPSIFD.GPSLatitudeRef] = b'N' if gps['latitude'] >= 0 else b'S'
                exif_dict['GPS'][piexif.GPSIFD.GPSLatitude] = deg_to_dms_rational(gps['latitude'])
                exif_dict['GPS'][piexif.GPSIFD.GPSLongitudeRef] = b'E' if gps['longitude'] >= 0 else b'W'
                exif_dict['GPS'][piexif.GPSIFD.GPSLongitude] = deg_to_dms_rational(gps['longitude'])
                if gps.get('altitude') is not None:
                    exif_dict['GPS'][piexif.GPSIFD.GPSAltitudeRef] = 0 if gps['altitude'] >= 0 else 1
                    exif_dict['GPS'][piexif.GPSIFD.GPSAltitude] = (int(abs(gps['altitude'] * 100)), 100)
            except Exception:
                pass
        if people:
            try:
                keywords = ";".join(people).encode('utf-16le')
                exif_dict['0th'][piexif.ImageIFD.XPKeywords] = keywords
            except Exception:
                pass
        try:
            exif_bytes = piexif.dump(exif_dict)
            img.save(file_path, "jpeg", exif=exif_bytes)
            return True
        except Exception:
            return False
    except Exception:
        return False

def extract_main_media_name_and_index(filename):
    base = os.path.basename(filename)
    name, _ = os.path.splitext(base)
    m = re.search(r'\((\d+)\)$', name)
    index = m.group(1) if m else ''
    name = re.sub(r'\(\d+\)$', '', name)
    name = re.sub(r'([_-]edited.*)$', '', name, flags=re.I)
    name = name.rstrip(' _-')
    return name.lower(), index

def build_json_lookup(all_json_files):
    lookup = {}
    for json_path in all_json_files:
        json_base = os.path.basename(json_path).lower()
        json_base_noext = json_base[:-5] if json_base.endswith('.json') else json_base
        m = re.search(r'\((\d+)\)$', json_base_noext)
        index = m.group(1) if m else ''
        base = re.sub(r'\(\d+\)$', '', json_base_noext)
        base = re.sub(r'([_-]edited.*)$', '', base, flags=re.I)
        base = base.rstrip(' _-')
        fullkey = base
        indexkey = f"{base}({index})" if index else None
        for key in filter(None, [fullkey, indexkey]):
            lookup.setdefault(key, []).append((json_base_noext, json_path))
    return lookup

def find_json_for_media(media_filename, json_lookup):
    base_name = os.path.basename(media_filename).lower()
    main_part, index = extract_main_media_name_and_index(media_filename)
    key_with_index = f"{main_part}({index})" if index else None
    key_plain = main_part

    # 1. Exact match: JSON whose base starts with the full media filename (including extension)
    candidates = []
    for key, entries in json_lookup.items():
        for jb, jp in entries:
            if jb.startswith(base_name):
                candidates.append((jb, jp))
    if candidates:
        unindexed = []
        indexed = []
        for jb, jp in candidates:
            if re.search(r'\(\d+\)$', jb):
                indexed.append((len(jb[len(base_name):]), jb[len(base_name):], jp))
            else:
                unindexed.append((len(jb[len(base_name):]), jb[len(base_name):], jp))
        if unindexed:
            unindexed.sort()
            return unindexed[0][2]
        elif indexed:
            indexed.sort()
            return indexed[0][2]

    # 2. If indexed, only allow indexed matches (as before)
    if index and key_with_index in json_lookup:
        candidates = json_lookup[key_with_index]
        suffix_candidates = [(len(jb[len(main_part):]), jb[len(main_part):], jp)
                             for jb, jp in candidates if jb.startswith(main_part)]
        if suffix_candidates:
            suffix_candidates.sort()
            return suffix_candidates[0][2]

    # 3. If not indexed, prefer non-indexed JSONs over indexed (as before)
    if not index and key_plain in json_lookup:
        candidates = json_lookup[key_plain]
        unindexed = []
        indexed = []
        for jb, jp in candidates:
            if re.search(r'\(\d+\)$', jb):
                indexed.append((len(jb[len(main_part):]), jb[len(main_part):], jp))
            else:
                unindexed.append((len(jb[len(main_part):]), jb[len(main_part):], jp))
        if unindexed:
            unindexed.sort()
            return unindexed[0][2]
        elif indexed:
            indexed.sort()
            return indexed[0][2]

    # 4. Fallback: Try matching by best prefix (hash case)
    media_name_noext, _ = os.path.splitext(base_name)
    best_prefix = None
    best_json = None
    best_len = 0
    for key, entries in json_lookup.items():
        for jb, jp in entries:
            jb_noext = jb
            if jb_noext.endswith('.supplemental-metadata'):
                jb_noext = jb_noext[:-len('.supplemental-metadata')]
            if jb_noext.endswith('.supplemental-metadata(1)'):
                jb_noext = jb_noext[:-len('.supplemental-metadata(1)')]
            jb_noext, _ = os.path.splitext(jb_noext)
            if media_name_noext.startswith(jb_noext) and len(jb_noext) >= 12:
                if len(jb_noext) > best_len:
                    best_prefix = jb_noext
                    best_json = jp
                    best_len = len(jb_noext)
    if best_json:
        return best_json

    # 5. fallback: substring match anywhere
    for key, entries in json_lookup.items():
        for jb, jp in entries:
            if main_part and main_part in jb:
                return jp
    for key, entries in json_lookup.items():
        for jb, jp in entries:
            if main_part and re.sub(r'[^a-z0-9]', '', main_part) in re.sub(r'[^a-z0-9]', '', jb):
                return jp
    return None

def unzip_takeout_zips(source_folder, temp_folder):
    zip_files = glob.glob(os.path.join(source_folder, "takeout-*.zip"))
    if not zip_files:
        return False
    print(f"Found {len(zip_files)} Takeout zip file(s). Unzipping them to: {temp_folder}")
    if not os.path.exists(temp_folder):
        os.makedirs(temp_folder)
    for i, zip_file in enumerate(zip_files):
        print_bar(i+1, len(zip_files))
        with zipfile.ZipFile(zip_file, "r") as z:
            # Extract all files flat into temp_folder
            for member in z.infolist():
                if member.is_dir():
                    continue
                # Remove any leading directories
                out_name = os.path.basename(member.filename)
                # If a file with the same name exists, add a unique suffix
                out_path = os.path.join(temp_folder, out_name)
                suffix = 1
                while os.path.exists(out_path):
                    parts = os.path.splitext(out_name)
                    out_path = os.path.join(temp_folder, f"{parts[0]}_{suffix}{parts[1]}")
                    suffix += 1
                with z.open(member) as fsrc, open(out_path, "wb") as fdst:
                    shutil.copyfileobj(fsrc, fdst)
    print()
    return True

def gather_files_flat(folder, ext_list):
    files = []
    for root, dirs, fls in os.walk(folder):
        for f in fls:
            if ext_list is None or f.lower().endswith(ext_list):
                files.append(os.path.join(root, f))
    return files

def main():
    # Set up folders and output
    source_folder = sys.argv[1] if len(sys.argv) > 1 else os.path.expanduser('~/Downloads')
    output_folder = os.path.join(source_folder, f'Output-{datetime.datetime.now().strftime("%Y%m%dT%H%M%S")}')
    temp_flat_folder = os.path.join(output_folder, "TEMP_FLAT")
    os.makedirs(output_folder, exist_ok=True)

    # Step 1: Unzip Takeout if necessary
    unzipped = unzip_takeout_zips(source_folder, temp_flat_folder)
    if unzipped:
        scan_folder = temp_flat_folder
    else:
        scan_folder = source_folder

    # Step 2: Gather files
    media_exts = ('.jpg', '.jpeg', '.png', '.gif', '.mp4', '.mov', '.avi', '.mkv', '.heic', '.webp', '.mpg')
    all_files = gather_files_flat(scan_folder, None)
    all_json_files = [f for f in all_files if f.lower().endswith('.json')]
    all_media_files = [f for f in all_files if f.lower().endswith(media_exts)]

    if not all_media_files:
        print("[ERROR] No media files found to process.")
        return

    log(f"[INFO] Found {len(all_media_files)} media files and {len(all_json_files)} JSONs.")

    already_processed = set(os.listdir(output_folder))
    valid_pairs = []
    unmatched_media = []
    matched_jsons = set()

    log("[INFO] Building JSON lookup for fast matching...")
    json_lookup = build_json_lookup(all_json_files)

    log("[INFO] Matching media to JSONs (with extension and hash-aware matching)...")
    for i, fl in enumerate(all_media_files):
        print_bar(i + 1, len(all_media_files))
        base_name = os.path.basename(fl)
        if base_name in already_processed:
            continue
        match = find_json_for_media(fl, json_lookup)
        if match:
            valid_pairs.append((fl, match))
            matched_jsons.add(match)
        else:
            unmatched_media.append(fl)
    print()

    unmatched_jsons = [jsonf for jsonf in all_json_files if jsonf not in matched_jsons]

    unmatched_media_path = os.path.join(output_folder, 'unmatched_media.txt')
    with open(unmatched_media_path, 'w', encoding='utf-8') as f:
        for item in unmatched_media:
            f.write(item + '\n')
    log(f"[INFO] Wrote unmatched media list: {unmatched_media_path}")

    unmatched_jsons_path = os.path.join(output_folder, 'unmatched_jsons.txt')
    with open(unmatched_jsons_path, 'w', encoding='utf-8') as f:
        for item in unmatched_jsons:
            f.write(item + '\n')
    log(f"[INFO] Wrote unmatched JSON list: {unmatched_jsons_path}")

    log(f"[INFO] Matched {len(valid_pairs)} media files with JSON. {len(unmatched_media)} unmatched media files. {len(unmatched_jsons)} unmatched JSONs.")

    log("[INFO] Copying files and applying metadata...")
    for i, (media, jsonf) in enumerate(valid_pairs):
        print_bar(i + 1, len(valid_pairs))
        dest_name = os.path.basename(media)
        dest_path = os.path.join(output_folder, dest_name)
        shutil.copy2(media, dest_path)
        json_dest_name = os.path.basename(jsonf)
        json_dest_path = os.path.join(output_folder, json_dest_name)
        shutil.copy2(jsonf, json_dest_path)
        meta = None
        try:
            with open(jsonf, encoding='utf-8') as jf:
                meta = json.load(jf)
        except Exception as e:
            log(f"[ERROR] Could not parse JSON: {jsonf} (skipping file: {media}) - {e}")
            continue

        people = []
        if 'people' in meta and isinstance(meta['people'], list):
            people = [p['name'] for p in meta['people'] if 'name' in p]

        exif_result = None
        if dest_name.lower().endswith(('.jpg', '.jpeg')):
            exif_result = set_exif_jpeg(dest_path, meta, people)
        elif dest_name.lower().endswith(('.mp4', '.mov', '.avi', '.mkv', '.mpg')):
            exif_result = f"People: {people}" if people else "No EXIF, but timestamp set"
        else:
            exif_result = "No EXIF for this type"

        ts = int(meta.get('photoTakenTime', {}).get('timestamp', '0'))
        if not ts:
            ts = int(meta.get('creationTime', {}).get('timestamp', '0'))
        if ts:
            os.utime(dest_path, (ts, ts))
            set_file_creation_time(dest_path, ts)
        else:
            log(f"    [WARN] No usable timestamp in JSON for: {dest_name}")

        log(f"[MATCH] MEDIA: {media} | JSON: {jsonf} | OUT: {dest_path} | People: {people if people else 'n/a'} | EXIF: {exif_result}")

    print()

    if unmatched_media:
        failed_dir = os.path.join(output_folder, 'FAILED')
        os.makedirs(failed_dir, exist_ok=True)
        log(f"[INFO] Copying {len(unmatched_media)} unmatched files to FAILED...")
        for i, media in enumerate(unmatched_media):
            print_bar(i + 1, len(unmatched_media))
            dest_failed = os.path.join(failed_dir, os.path.basename(media))
            shutil.copy2(media, dest_failed)
            log(f"[FAILED] MEDIA: {media} | Reason: No JSON match")
        print()

    log_path = os.path.join(output_folder, 'process_log.txt')
    with open(log_path, 'w', encoding='utf-8') as f:
        for line in LOG_LINES:
            f.write(line + '\n')

    if WINDOWS:
        log("[INFO] Windows detected: File creation, modification, and access times are set for enriched files.")
    else:
        log("[INFO] Non-Windows OS detected: Only file modification and access times are set (creation time is not supported by Python on this OS).")

    log(f"\n[INFO] Processing complete.")
    log(f"    Total media files: {len(all_media_files)}")
    log(f"    Matched with JSON: {len(valid_pairs)}")
    log(f"    Unmatched media:  {len(unmatched_media)}")
    log(f"    Unmatched JSON:   {len(unmatched_jsons)}")
    log(f"    Output in: {output_folder}")
    log(f"    Log written to: {log_path}")

    # Cleanup temp folder
    if unzipped:
        shutil.rmtree(temp_flat_folder, ignore_errors=True)

if __name__ == '__main__':
    main()
