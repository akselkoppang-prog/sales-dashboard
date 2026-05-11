"""
build_vba_bin.py — Patch the original xlsxwriter example vbaProject.bin,
replacing Module1's source with our full VBA macro suite.

Strategy: binary-patch approach
  - Keep all original CFB structures intact (mini-stream, FAT, dir tree, CLSIDs)
  - Move Module1 from mini-stream to regular sectors (it grows to ~10 KB)
  - Free old Module1 mini-sectors in the mini-FAT
  - Add new regular sectors at end of file
  - Update FAT chain and Module1 directory entry

Run once:  python3 build_vba_bin.py
Output:    vba_project.bin  (used by excel_generator.py)
"""
import struct, math, os, urllib.request, tarfile, io

FREESECT   = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT    = 0xFFFFFFFD

# ── MS-OVBA compression (uncompressed chunks) ─────────────────────────────────

def ovba_compress(data: bytes) -> bytes:
    """MS-OVBA compress: uncompressed 4096-byte chunks + literal final chunk."""
    result = bytearray(b'\x01')
    i = 0; n = len(data)
    while i < n:
        chunk = data[i : i + 4096]
        csz = len(chunk)
        if csz == 4096:
            result += struct.pack('<H', 0x3FFF)   # uncompressed chunk header
            result += chunk
        else:
            buf = bytearray()
            j = 0
            while j < csz:
                buf.append(0x00)
                buf += chunk[j : j + 8]
                j += 8
            result += struct.pack('<H', 0xB000 | (2 + len(buf) - 3))
            result += buf
        i += 4096
    return bytes(result)


# ── Get original template from PyPI ──────────────────────────────────────────

def _fetch_original_bin() -> bytes:
    url = 'https://files.pythonhosted.org/packages/source/x/xlsxwriter/xlsxwriter-3.2.9.tar.gz'
    with urllib.request.urlopen(url, timeout=30) as r:
        tardata = r.read()
    with tarfile.open(fileobj=io.BytesIO(tardata)) as t:
        member = t.getmember('xlsxwriter-3.2.9/examples/vbaProject.bin')
        return t.extractfile(member).read()


# ── CFB helpers ───────────────────────────────────────────────────────────────

def _read_fat(raw: bytes, difat_list: list) -> list:
    """Read all FAT entries (covers sectors 0..127 per FAT sector)."""
    fat = [FREESECT] * 256
    for idx, fat_sec in enumerate(difat_list):
        offset = (fat_sec + 1) * 512
        entries = struct.unpack_from('<128I', raw, offset)
        for i, v in enumerate(entries):
            fat[idx * 128 + i] = v
    return fat


def _follow_chain(fat: list, start: int) -> list:
    chain = []
    cur = start
    while cur not in (ENDOFCHAIN, FREESECT, FATSECT, 0xFFFFFFFC):
        chain.append(cur)
        if cur >= len(fat):
            break
        cur = fat[cur]
    return chain


# ── Main patcher ──────────────────────────────────────────────────────────────

def build_vba_project_bin(vba_source_code: str) -> bytes:
    """
    Patch the xlsxwriter example vbaProject.bin to embed our VBA macros.
    Returns the patched binary ready to be saved as vba_project.bin.
    """
    raw = bytearray(_fetch_original_bin())

    # ── Parse header ──────────────────────────────────────────────────────────
    sector_sz    = 1 << struct.unpack_from('<H', raw, 30)[0]  # 512
    mini_sec_sz  = 1 << struct.unpack_from('<H', raw, 32)[0]  # 64
    fat_count    = struct.unpack_from('<I', raw, 44)[0]
    first_dir    = struct.unpack_from('<I', raw, 48)[0]
    mini_cutoff  = struct.unpack_from('<I', raw, 56)[0]       # 4096
    first_minifat= struct.unpack_from('<I', raw, 60)[0]
    minifat_count= struct.unpack_from('<I', raw, 64)[0]

    difat = []
    for k in range(109):
        v = struct.unpack_from('<I', raw, 76 + k * 4)[0]
        if v != FREESECT:
            difat.append(v)

    fat = _read_fat(raw, difat)
    dir_chain   = _follow_chain(fat, first_dir)
    minifat_chain = _follow_chain(fat, first_minifat)

    # ── Read mini-FAT ─────────────────────────────────────────────────────────
    minifat = []
    for ms in minifat_chain:
        offset = (ms + 1) * sector_sz
        minifat.extend(struct.unpack_from('<128I', raw, offset))

    # ── Find Module1 directory entry (entry index 4) ─────────────────────────
    # Dir chain [1, 7, 13, 27] → entry 4 is at chain[1]=sector7, first entry
    def dir_entry_offset(entry_idx: int) -> int:
        chain_idx = entry_idx // 4
        intra_idx = entry_idx % 4
        sector = dir_chain[chain_idx]
        return (sector + 1) * sector_sz + intra_idx * 128

    MODULE1_IDX = 4
    m1_off = dir_entry_offset(MODULE1_IDX)

    m1_mini_start = struct.unpack_from('<I', raw, m1_off + 116)[0]  # = 32
    m1_orig_size  = struct.unpack_from('<I', raw, m1_off + 120)[0]  # = 1056
    m1_mini_count = math.ceil(m1_orig_size / mini_sec_sz)           # = 17

    # ── Build new Module1 stream data ─────────────────────────────────────────
    # Keep original p-code header (bytes 0..954); replace compressed source
    # Read old Module1 data from mini-stream
    root_start  = struct.unpack_from('<I', raw, dir_entry_offset(0) + 116)[0]  # = 3
    root_size   = struct.unpack_from('<I', raw, dir_entry_offset(0) + 120)[0]  # = 11456
    root_chain  = _follow_chain(fat, root_start)
    mini_container = b''.join(
        bytes(raw[(s + 1) * sector_sz : (s + 2) * sector_sz]) for s in root_chain
    )

    # Extract original Module1 header from mini-stream
    m1_mini_chain = _follow_chain(minifat, m1_mini_start)
    m1_orig_bytes = b''.join(
        mini_container[ms * mini_sec_sz : (ms + 1) * mini_sec_sz]
        for ms in m1_mini_chain
    )[:m1_orig_size]
    m1_header = m1_orig_bytes[:955]

    # Compress new VBA source
    src_line = 'Attribute VB_Name = "Module1"\r\n' + vba_source_code
    compressed_src = ovba_compress(src_line.encode('latin1'))
    new_m1_data = m1_header + compressed_src
    new_m1_size = len(new_m1_data)
    new_m1_sectors_needed = math.ceil(new_m1_size / sector_sz)  # regular sectors

    # ── Free old Module1 mini-sectors in mini-FAT ────────────────────────────
    # Mini-FAT sector 0 is at minifat_chain[0] = sector 2 in the file
    # Mini-sectors 32..48 (m1_mini_chain) → mark as FREESECT
    for ms in m1_mini_chain:
        minifat[ms] = FREESECT

    # Write modified mini-FAT back
    for idx, ms_sector in enumerate(minifat_chain):
        chunk = minifat[idx * 128 : (idx + 1) * 128]
        packed = struct.pack('<128I', *chunk)
        start  = (ms_sector + 1) * sector_sz
        raw[start : start + sector_sz] = packed

    # ── Append new regular sectors for Module1 ───────────────────────────────
    # Find current last sector number
    file_sectors = (len(raw) - sector_sz) // sector_sz  # = 30 (sectors 0..29)
    new_start_sector = file_sectors  # first new sector = 30

    # Pad new Module1 data to sector boundary and append
    padded = new_m1_data + b'\x00' * (new_m1_sectors_needed * sector_sz - new_m1_size)
    raw.extend(padded)

    # ── Update FAT for new sectors ────────────────────────────────────────────
    # Extend fat list if needed
    while len(fat) < new_start_sector + new_m1_sectors_needed + 2:
        fat.append(FREESECT)

    for k in range(new_m1_sectors_needed - 1):
        fat[new_start_sector + k] = new_start_sector + k + 1
    fat[new_start_sector + new_m1_sectors_needed - 1] = ENDOFCHAIN

    # Write modified FAT back (single FAT sector at sector difat[0]=0)
    # Total sectors now = 30 + new_m1_sectors_needed ≤ 128 → fits in 1 FAT sector
    assert new_start_sector + new_m1_sectors_needed <= 128, "Need second FAT sector!"
    fat_sector_offset = (difat[0] + 1) * sector_sz
    packed_fat = struct.pack('<128I', *fat[:128])
    raw[fat_sector_offset : fat_sector_offset + sector_sz] = packed_fat

    # ── Update Module1 directory entry ───────────────────────────────────────
    # start_sec (bytes 116-119) = new_start_sector
    # size       (bytes 120-123) = new_m1_size  (low DWORD; high DWORD stays 0)
    struct.pack_into('<I', raw, m1_off + 116, new_start_sector)
    struct.pack_into('<I', raw, m1_off + 120, new_m1_size)

    return bytes(raw)


# ── CLI entry point ───────────────────────────────────────────────────────────

if __name__ == '__main__':
    HERE    = os.path.dirname(os.path.abspath(__file__))
    vba_txt = os.path.join(HERE, 'VBA_MODULES.txt')
    out     = os.path.join(HERE, 'vba_project.bin')

    with open(vba_txt, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    start = next(i for i, l in enumerate(lines) if l.strip() == 'Option Explicit')
    code  = ''.join(lines[start:])
    code  = code.encode('ascii', errors='replace').decode('ascii')
    code  = code.replace('\r\n', '\n').replace('\n', '\r\n')

    print('Fetching original vbaProject.bin from PyPI...')
    data = build_vba_project_bin(code)

    with open(out, 'wb') as f:
        f.write(data)
    print(f'Written: {out}  ({len(data):,} bytes)')

    # Verify VBA is readable
    from oletools.olevba import VBA_Parser, decompress_stream
    import olefile
    with olefile.OleFileIO(out) as ole:
        m1 = ole.openstream('VBA/Module1').read()
        src = decompress_stream(m1[955:]).decode('latin1')
        for sub in ['FilterYear2023', 'FilterAllYears', 'GoToDashboard', 'AddYearButtons', 'ExportActiveToPDF']:
            status = 'OK' if sub in src else 'MISSING'
            print(f'  {status}: {sub}')
    print('Done.')
