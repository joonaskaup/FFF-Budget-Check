"""Sanitise the Mac Forms library reference in a closed standalone XLS copy.

Preserve worksheet bytes and VBA source; replace only the local username with
an equally long generic placeholder. Reopen and test the output in native Excel.
Adapted from the Actors Budget Assistant release utility (GPL-3.0-only).
"""
import argparse
from pathlib import Path
import re
import struct
import olefile
from oletools.olevba import decompress_stream

def compress(data):
    result = bytearray(b"\x01")
    for start in range(0, len(data), 4096):
        chunk = data[start:start + 4096]
        output = bytearray()
        pos = 0
        while pos < len(chunk):
            flag_at = len(output)
            output.append(0)
            for bit in range(8):
                if pos >= len(chunk):
                    break
                offset_bits = max(4, (pos - 1).bit_length())
                length_bits = 16 - offset_bits
                max_length = min(len(chunk) - pos, (1 << length_bits) + 2)
                best_length = 0
                best_distance = 0
                for distance in range(1, min(pos, 1 << offset_bits) + 1):
                    length = 0
                    while length < max_length and chunk[pos + length] == chunk[pos + length - distance]:
                        length += 1
                    if length > best_length:
                        best_length, best_distance = length, distance
                if best_length >= 3:
                    output[flag_at] |= 1 << bit
                    token = ((best_distance - 1) << length_bits) | (best_length - 3)
                    output.extend(struct.pack("<H", token))
                    pos += best_length
                else:
                    output.append(chunk[pos])
                    pos += 1
        if len(output) >= 4096:
            if len(chunk) != 4096:
                raise ValueError("This sanitizer expects a small compressible dir stream")
            result.extend(struct.pack("<H", 0x3FFF) + chunk)
        else:
            result.extend(struct.pack("<H", 0xB000 | (len(output) - 1)) + output)
    assert decompress_stream(bytes(result)) == data
    return bytes(result)


def scrub(path):
    dir_path = "_VBA_PROJECT_CUR/VBA/dir"
    cache_path = "_VBA_PROJECT_CUR/VBA/_VBA_PROJECT"
    with olefile.OleFileIO(path, write_mode=True) as ole:
        original_workbook = ole.openstream("Workbook").read()
        old_dir = ole.openstream(dir_path).read()
        plain = decompress_stream(old_dir)
        match = re.search(rb"/Users/([^/]+)/Library/Containers/com\.microsoft\.Excel/", plain)
        if not match:
            raise ValueError("Expected the known Mac Forms library reference")
        username = match.group(1)
        if len(username) < 9:
            raise ValueError("Username too short for fixed-width placeholder")
        placeholder = b"ExcelUser" + b"_" * (len(username) - 9)
        old_path = b"/Users/" + username + b"/"
        new_path = b"/Users/" + placeholder + b"/"
        plain = plain.replace(old_path, new_path)
        new_dir = compress(plain)
        if len(new_dir) > len(old_dir):
            raise ValueError("Replacement will not fit; original file is unchanged")
        ole.write_stream(dir_path, new_dir + bytes(len(old_dir) - len(new_dir)))
        cache = ole.openstream(cache_path).read()
        cache = cache.replace(old_path.decode().encode("utf-16-le"), new_path.decode().encode("utf-16-le"))
        cache = cache.replace(old_path, new_path)
        ole.write_stream(cache_path, cache)
        sid = ole._find(dir_path)
        directory_sector = ole.first_dir_sector
        for _ in range(sid // (ole.sectorsize // 128)):
            directory_sector = ole.fat[directory_sector]
        size_position = (directory_sector + 1) * ole.sectorsize + (sid % (ole.sectorsize // 128)) * 128 + 120
    data = bytearray(path.read_bytes())
    data[size_position:size_position + 8] = struct.pack("<Q", len(new_dir))
    path.write_bytes(data)
    with olefile.OleFileIO(path) as ole:
        assert decompress_stream(ole.openstream(dir_path).read()) == plain
        assert ole.openstream("Workbook").read() == original_workbook
    print("Mac library username sanitised; worksheet bytes unchanged.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path)
    scrub(parser.parse_args().workbook)
