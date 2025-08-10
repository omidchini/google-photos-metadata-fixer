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
        print('\n[ERROR] pywint32 is required for setting creation time on Windows.\nRun: pip install pywin32\n')
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
        exif_bytes = img.info.get('exif')
        if exif_bytes:
            exif_dict = piexif.load(exif_bytes)
        else:
            exif_dict = {"0th": {}, "Exif": {}, "GPS": {}, "1st": {}, "thumbnail": None}

        # Update DateTimeOriginal if present in JSON
        try:
            dt = datetime.datetime.utcfromtimestamp(int(json_data['photoTakenTime']['timestamp']))
            exif_dict['Exif'][piexif.ExifIFD.DateTimeOriginal] = dt.strftime("%Y:%m:%d %H:%M:%S")
        except Exception:
            pass

        # Update GPS if present in JSON
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

        # Add XPKeywords if any people
        if people:
            try:
                keywords = ";".join(people).encode('utf-16le')
                exif_dict['0th'][piexif.ImageIFD.XPKeywords] = keywords
            except Exception:
                pass

        exif_bytes = piexif.dump(exif_dict)
        img.save(file_path, "jpeg", exif=exif_bytes)
        return True
    except Exception as e:
        print(f"[EXIF ERROR] {file_path}: {e}")
        return False

def gather_files_flat(folder, ext_list):
    files = []
    for root, dirs, fls in os.walk(folder):
        for f in fls:
            if ext_list is None or f.lower().endswith(ext_list):
                files.append(os.path.join(root, f))
    return files

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
            for member in z.infolist():
                if member.is_dir():
                    continue
                out_name = os.path.basename(member.filename)
                out_path = os.path.join(temp_flat_folder, out_name)
                suffix = 1
                while os.path.exists(out_path):
                    parts = os.path.splitext(out_name)
                    out_path = os.path.join(temp_flat_folder, f"{parts[0]}_{suffix}{parts[1]}")
                    suffix += 1
                with z.open(member) as fsrc, open(out_path, "wb") as fdst:
                    shutil.copyfileobj(fsrc, fdst)
    print()
    return True

def parse_indices(s):
    """Remove trailing (N) indices from a string and return (base, [indices])."""
    indices = []
    while True:
        m = re.search(r'\((\d+)\)$', s)
        if not m:
            break
        indices.insert(0, int(m.group(1)))
        s = s[:m.start()]
    return s, indices

def parse_json_core_indices(json_filename):
    """
    Extract the base, indices, and ext from a Takeout JSON.
    Handles any number of dots in the base.
    """
    SUPPORTED_EXTS = ('.jpg', '.jpeg', '.png', '.gif', '.mp4', '.mov', '.avi', '.mkv', '.heic', '.webp', '.mpg')
    base = os.path.basename(json_filename)
    if not base.lower().endswith('.json'):
        return None, None, None
    s = base[:-5]  # remove .json

    # Find the *last* supported extension in the string
    last_ext_pos = -1
    ext_found = None
    for ext in SUPPORTED_EXTS:
        pos = s.lower().rfind(ext)
        if pos > last_ext_pos:
            last_ext_pos = pos
            ext_found = ext
    if ext_found is not None and last_ext_pos != -1:
        core = s[:last_ext_pos]
        after_ext = s[last_ext_pos + len(ext_found):]
        core_base, indices1 = parse_indices(core)
        indices2 = [int(i) for i in re.findall(r'\((\d+)\)', after_ext)]
        indices = tuple(indices1 + indices2)
        return core_base, indices, ext_found
    else:
        # fallback: no recognized extension found
        core_base, indices = parse_indices(s)
        return core_base, tuple(indices), None

def build_media_lookup(all_media_files):
    # {(core, indices, ext): path, ...}
    lookup = {}
    # Also: {(core, ext): [(indices, path), ...]}
    loose_lookup = {}
    for f in all_media_files:
        base = os.path.basename(f)
        root, ext = os.path.splitext(base)
        core, indices = parse_indices(root)
        k = (core, tuple(indices), ext.lower())
        lookup[k] = f
        # For loose match
        k_loose = (core, ext.lower())
        loose_lookup.setdefault(k_loose, []).append((tuple(indices), f))
        # 1-char truncation for Takeout bug
        if len(core) > 1:
            k2 = (core[:-1], tuple(indices), ext.lower())
            lookup[k2] = f
            k2_loose = (core[:-1], ext.lower())
            loose_lookup.setdefault(k2_loose, []).append((tuple(indices), f))
    return lookup, loose_lookup

def indices_set(indices):
    return set(indices)

def match_json_to_media(jsonf, strict_lookup, loose_lookup):
    jb, jindices, jext = parse_json_core_indices(jsonf)
    if jb is None:
        return None

    exts_to_try = [jext] if jext else ['.jpg', '.jpeg', '.png', '.gif', '.mp4', '.mov', '.avi', '.mkv', '.heic', '.webp', '.mpg']

    # Pass 1: strict (core, indices, ext)
    for ext in exts_to_try:
        k = (jb, jindices, ext)
        if k in strict_lookup:
            return strict_lookup[k]
    # Pass 1b: strict, 1-char truncation
    for ext in exts_to_try:
        if len(jb) > 1:
            k = (jb[:-1], jindices, ext)
            if k in strict_lookup:
                return strict_lookup[k]
    # Pass 2: loose, match on core/ext, and indices as sets (allow superset/subset)
    for ext in exts_to_try:
        k_loose = (jb, ext)
        if k_loose in loose_lookup:
            for mindices, f in loose_lookup[k_loose]:
                if not jindices or not mindices:
                    return f
                if indices_set(jindices) == indices_set(mindices):
                    return f
                if indices_set(jindices).issubset(indices_set(mindices)) or indices_set(mindices).issubset(indices_set(jindices)):
                    return f
    # Pass 2b: loose, 1-char truncation
    for ext in exts_to_try:
        if len(jb) > 1:
            k_loose = (jb[:-1], ext)
            if k_loose in loose_lookup:
                for mindices, f in loose_lookup[k_loose]:
                    if not jindices or not mindices:
                        return f
                    if indices_set(jindices) == indices_set(mindices):
                        return f
                    if indices_set(jindices).issubset(indices_set(mindices)) or indices_set(mindices).issubset(indices_set(jindices)):
                        return f
    return None

def main():
    source_folder = sys.argv[1] if len(sys.argv) > 1 else os.path.expanduser('~/Downloads')
    output_folder = os.path.join(os.getcwd(), f'Output-{datetime.datetime.now().strftime("%Y%m%dT%H%M%S")}')
    temp_flat_folder = os.path.join(output_folder, "TEMP_FLAT")
    os.makedirs(output_folder, exist_ok=True)

    unzipped = unzip_takeout_zips(source_folder, temp_flat_folder)
    scan_folder = temp_flat_folder if unzipped else source_folder

    media_exts = ('.jpg', '.jpeg', '.png', '.gif', '.mp4', '.mov', '.avi', '.mkv', '.heic', '.webp', '.mpg')
    all_files = gather_files_flat(scan_folder, None)
    all_json_files = [f for f in all_files if f.lower().endswith('.json')]
    all_media_files = [
        f for f in all_files
        if f.lower().endswith(media_exts) and not re.search(r'(-edited|-edit|-crop)', os.path.basename(f), re.IGNORECASE)
    ]

    if not all_media_files:
        print("[ERROR] No media files found to process.")
        return

    log(f"[INFO] Found {len(all_media_files)} media files and {len(all_json_files)} JSONs.")

    already_processed = set(os.listdir(output_folder))

    log("[INFO] Building media lookup for robust JSON-to-media matching...")
    strict_lookup, loose_lookup = build_media_lookup(all_media_files)

    valid_pairs = []
    unmatched_jsons = []
    matched_media = set()

    log("[INFO] Matching JSONs to media files (robust Takeout normalization)...")
    for i, jsonf in enumerate(all_json_files):
        print_bar(i + 1, len(all_json_files))
        base_name = os.path.basename(jsonf)
        if base_name in already_processed:
            continue
        found = match_json_to_media(jsonf, strict_lookup, loose_lookup)
        if found and found not in matched_media:
            valid_pairs.append((found, jsonf))
            matched_media.add(found)
        else:
            unmatched_jsons.append(jsonf)
    print()
    log(f"[INFO] Starting move and EXIF loop for {len(valid_pairs)} pairs")

    unmatched_media = [mediaf for mediaf in all_media_files if mediaf not in matched_media]

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

    matched_pairs_path = os.path.join(output_folder, 'matched_pairs.txt')
    with open(matched_pairs_path, 'w', encoding='utf-8') as f:
        for media, jsonf in valid_pairs:
            f.write(f"{media}\t{jsonf}\n")
    log(f"[INFO] Wrote matched pairs list: {matched_pairs_path}")

    log(f"[INFO] Matched {len(valid_pairs)} JSON files with media. {len(unmatched_media)} unmatched media files. {len(unmatched_jsons)} unmatched JSONs.")

    # ==== Uncomment the following block to perform move/EXIF operations ====
    
    log("[INFO] Moving files and applying metadata...")
    for i, (media, jsonf) in enumerate(valid_pairs):
        print_bar(i + 1, len(valid_pairs))
        dest_name = os.path.basename(media)
        dest_path = os.path.join(output_folder, dest_name)
        json_dest_name = os.path.basename(jsonf)
        json_dest_path = os.path.join(output_folder, json_dest_name)

        # Move files (not copy)
        shutil.move(media, dest_path)
        shutil.move(jsonf, json_dest_path)
        meta = None
        try:
            with open(json_dest_path, encoding='utf-8') as jf:
                meta = json.load(jf)
        except Exception as e:
            log(f"[ERROR] Could not parse JSON: {json_dest_path} (skipping file: {dest_path}) - {e}")
            continue

        people = []
        if 'people' in meta and isinstance(meta['people'], list):
            people = [p['name'] for p in meta['people'] if 'name' in p]

        exif_result = None
        if dest_name.lower().endswith(('.jpg', '.jpeg')):
            exif_result = "Skipped EXIF for JPEG"  # No EXIF for JPEGs in this script
            #exif_result = set_exif_jpeg(dest_path, meta, people)
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
            os.utime(json_dest_path, (ts, ts))
            set_file_creation_time(json_dest_path, ts)
        else:
            log(f"    [WARN] No usable timestamp in JSON for: {dest_name}")

        log(f"[MATCH] MEDIA: {media} | JSON: {jsonf} | OUT: {dest_path} | People: {people if people else 'n/a'} | EXIF: {exif_result}")

    log("[INFO] Finished move and EXIF loop.")
    print()
    if unmatched_media:
        failed_dir = os.path.join(output_folder, 'FAILED')
        os.makedirs(failed_dir, exist_ok=True)
        log(f"[INFO] Moving {len(unmatched_media)} unmatched files to FAILED...")
        for i, media in enumerate(unmatched_media):
            print_bar(i + 1, len(unmatched_media))
            dest_failed = os.path.join(failed_dir, os.path.basename(media))
            shutil.move(media, dest_failed)
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

    if unzipped:
        shutil.rmtree(temp_flat_folder, ignore_errors=True)

if __name__ == '__main__':
    main()
