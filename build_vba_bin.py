"""
build_vba_bin.py — Build a custom vbaProject.bin embedding our VBA macros.

Run once:  python3 build_vba_bin.py
Output:    vba_project.bin  (used by excel_generator.py)
"""
import struct, os
import olefile

FREESECT   = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT    = 0xFFFFFFFD
NOSTREAM   = 0xFFFFFFFF
SECTOR_SZ  = 512

# ── MS-OVBA compression ───────────────────────────────────────────────────────

def ovba_compress(data: bytes) -> bytes:
    """Compress using uncompressed chunks for full 4096-byte blocks,
    literal-token compressed chunks for the final partial block."""
    result = bytearray(b'\x01')   # signature byte
    i = 0; n = len(data)
    while i < n:
        chunk = data[i : i + 4096]
        csz = len(chunk)
        if csz == 4096:
            # Uncompressed chunk: CompressedChunkFlag=0, chunk_size must be 4098
            result += struct.pack('<H', 0x3FFF)
            result += chunk
        else:
            # Last partial chunk: literal tokens (flag byte = 0 = all literals)
            buf = bytearray()
            j = 0
            while j < csz:
                grp = chunk[j : j + 8]
                buf.append(0x00)     # flag byte: all literals
                buf += grp
                j += 8
            # chunk_total = 2 (header) + len(buf); header bits[11:0] = chunk_total - 3
            chunk_total = 2 + len(buf)
            result += struct.pack('<H', 0xB000 | (chunk_total - 3))
            result += buf
        i += 4096
    return bytes(result)


# ── CFB helpers ───────────────────────────────────────────────────────────────

def _pad(data: bytes) -> bytes:
    """Pad bytes to next sector boundary."""
    r = len(data) % SECTOR_SZ
    if r:
        data += b'\x00' * (SECTOR_SZ - r)
    return data


def _dir_entry(name: str, obj_type: int, color: int,
               left: int, right: int, child: int,
               clsid: bytes, start_sec: int, size: int) -> bytes:
    """Build a 128-byte CFB directory entry."""
    name_enc = (name + '\x00').encode('utf-16-le')
    name_bytes = name_enc + b'\x00' * (64 - len(name_enc))
    entry = (
        name_bytes                    +   # 64 bytes
        struct.pack('<H', len(name_enc)) +  # 2: name length incl. null
        struct.pack('<B', obj_type)    +   # 1: 0=empty, 1=storage, 2=stream, 5=root
        struct.pack('<B', color)       +   # 1: 0=red, 1=black
        struct.pack('<I', left)        +   # 4: left sibling
        struct.pack('<I', right)       +   # 4: right sibling
        struct.pack('<I', child)       +   # 4: child (first child for storage/root)
        clsid                          +   # 16: CLSID
        struct.pack('<I', 0)           +   # 4: state bits
        struct.pack('<Q', 0)           +   # 8: created time
        struct.pack('<Q', 0)           +   # 8: modified time
        struct.pack('<I', start_sec)   +   # 4: starting sector
        struct.pack('<Q', size)            # 8: size (QWORD, high DWORD = 0 for v3)
    )
    assert len(entry) == 128, f"dir entry size {len(entry)} != 128"
    return entry


# ── Build the binary ──────────────────────────────────────────────────────────

def build_vba_project_bin(source_vba: str, template_path: str) -> bytes:
    """
    source_vba: the full VBA source code to embed (ASCII-clean string)
    template_path: original vbaProject.bin to copy non-Module1 streams from
    """
    # Load all original streams
    streams: dict[str, bytes] = {}
    with olefile.OleFileIO(template_path) as ole:
        for parts in ole.listdir():
            path = '/'.join(parts)
            streams[path] = ole.openstream(path).read()

    # Build new Module1: keep original p-code header (bytes 0..954), replace source
    m1_header = streams['VBA/Module1'][:955]
    compressed_src = ovba_compress(
        ('Attribute VB_Name = "Module1"\r\n' + source_vba).encode('latin1')
    )
    streams['VBA/Module1'] = m1_header + compressed_src

    # Ordered list of streams to lay out in the file
    # (VBA storage children first, then root-level streams)
    stream_order = [
        'VBA/Module1',
        'VBA/Sheet1',
        'VBA/Sheet2',
        'VBA/ThisWorkbook',
        'VBA/ThisWorkbook1',
        'VBA/_VBA_PROJECT',
        'VBA/__SRP_0',
        'VBA/__SRP_1',
        'VBA/__SRP_2',
        'VBA/__SRP_3',
        'VBA/dir',
        'PROJECTwm',
        'PROJECT',
    ]

    # ── Assign sectors ────────────────────────────────────────────────────────
    # Layout: sector 0 = FAT, sectors 1..4 = directory (16 entries × 128 B = 2048 B = 4 sectors)
    # sectors 5+ = data
    DIR_START   = 1
    DIR_SECTORS = 4
    DATA_START  = DIR_START + DIR_SECTORS   # = 5

    sector_of: dict[str, int] = {}
    cur = DATA_START
    for key in stream_order:
        data = streams.get(key, b'')
        sector_of[key] = cur
        n_sec = max(1, (len(data) + SECTOR_SZ - 1) // SECTOR_SZ)
        cur += n_sec
    total_sectors = cur

    # ── FAT ──────────────────────────────────────────────────────────────────
    fat = [FREESECT] * max(128, total_sectors + 2)
    fat[0] = FATSECT    # sector 0 is the FAT itself

    # Directory chain: sectors 1..4
    for s in range(DIR_START, DIR_START + DIR_SECTORS - 1):
        fat[s] = s + 1
    fat[DIR_START + DIR_SECTORS - 1] = ENDOFCHAIN

    # Stream chains
    for key in stream_order:
        data = streams.get(key, b'')
        start = sector_of[key]
        n_sec = max(1, (len(data) + SECTOR_SZ - 1) // SECTOR_SZ)
        for k in range(n_sec - 1):
            fat[start + k] = start + k + 1
        fat[start + n_sec - 1] = ENDOFCHAIN

    fat_bytes = struct.pack('<' + 'I' * len(fat), *fat)[:SECTOR_SZ]

    # ── Directory entries ─────────────────────────────────────────────────────
    # Entry indices:
    # [0]  Root Entry (child → [1] VBA)
    # [1]  VBA storage  (child → [2] Module1, right → [13] PROJECT)
    # [2]  Module1    right→[3]
    # [3]  Sheet1     right→[4]
    # [4]  Sheet2     right→[5]
    # [5]  ThisWorkbook  right→[6]
    # [6]  ThisWorkbook1 right→[7]
    # [7]  _VBA_PROJECT  right→[8]
    # [8]  __SRP_0    right→[9]
    # [9]  __SRP_1    right→[10]
    # [10] __SRP_2    right→[11]
    # [11] __SRP_3    right→[12]
    # [12] dir        (last VBA child)
    # [13] PROJECT    right→[14]
    # [14] PROJECTwm
    # [15] (empty)

    ROOT_CLSID = bytes.fromhex('06090200000000000000000000000046')
    NO_CLSID   = b'\x00' * 16

    def _s(key):  # start sector for a stream
        return sector_of.get(key, FREESECT)

    def _sz(key): # size for a stream
        return len(streams.get(key, b''))

    vba_children = [
        ('Module1',       'VBA/Module1'),
        ('Sheet1',        'VBA/Sheet1'),
        ('Sheet2',        'VBA/Sheet2'),
        ('ThisWorkbook',  'VBA/ThisWorkbook'),
        ('ThisWorkbook1', 'VBA/ThisWorkbook1'),
        ('_VBA_PROJECT',  'VBA/_VBA_PROJECT'),
        ('__SRP_0',       'VBA/__SRP_0'),
        ('__SRP_1',       'VBA/__SRP_1'),
        ('__SRP_2',       'VBA/__SRP_2'),
        ('__SRP_3',       'VBA/__SRP_3'),
        ('dir',           'VBA/dir'),
    ]

    dir_entries = []

    # [0] Root Entry — child = 1 (VBA)
    dir_entries.append(_dir_entry(
        'Root Entry', 5, 1, NOSTREAM, NOSTREAM, 1,
        ROOT_CLSID, FREESECT, 0))

    # [1] VBA storage — child = 2 (Module1, first VBA child), right = 13 (PROJECT)
    dir_entries.append(_dir_entry(
        'VBA', 1, 1, NOSTREAM, 13, 2, NO_CLSID, FREESECT, 0))

    # [2..12] VBA stream children (linear right-sibling chain)
    for i, (name, key) in enumerate(vba_children):
        right = 3 + i if i < len(vba_children) - 1 else NOSTREAM
        dir_entries.append(_dir_entry(
            name, 2, 1, NOSTREAM, right, NOSTREAM,
            NO_CLSID, _s(key), _sz(key)))

    # [13] PROJECT — right = 14 (PROJECTwm)
    dir_entries.append(_dir_entry(
        'PROJECT', 2, 1, NOSTREAM, 14, NOSTREAM,
        NO_CLSID, _s('PROJECT'), _sz('PROJECT')))

    # [14] PROJECTwm
    dir_entries.append(_dir_entry(
        'PROJECTwm', 2, 1, NOSTREAM, NOSTREAM, NOSTREAM,
        NO_CLSID, _s('PROJECTwm'), _sz('PROJECTwm')))

    # [15] empty
    dir_entries.append(_dir_entry(
        '', 0, 1, NOSTREAM, NOSTREAM, NOSTREAM, NO_CLSID, FREESECT, 0))

    assert len(dir_entries) == 16
    dir_bytes = b''.join(dir_entries)   # 2048 bytes = 4 sectors

    # ── CFB Header ───────────────────────────────────────────────────────────
    hdr = bytearray(512)
    hdr[0:8]  = bytes.fromhex('D0CF11E0A1B11AE1')   # signature
    struct.pack_into('<H', hdr, 24, 0x003E)           # minor version
    struct.pack_into('<H', hdr, 26, 0x0003)           # major version 3
    struct.pack_into('<H', hdr, 28, 0xFFFE)           # little-endian
    struct.pack_into('<H', hdr, 30, 0x0009)           # sector shift (2^9=512)
    struct.pack_into('<H', hdr, 32, 0x0006)           # mini sector shift (2^6=64)
    struct.pack_into('<I', hdr, 44, 1)                # FAT sector count
    struct.pack_into('<I', hdr, 48, DIR_START)        # first dir sector
    struct.pack_into('<I', hdr, 56, 0x1000)           # mini stream cutoff = 4096
    struct.pack_into('<I', hdr, 60, FREESECT)         # first mini FAT sector (none)
    struct.pack_into('<I', hdr, 64, 0)                # mini FAT sector count
    struct.pack_into('<I', hdr, 68, FREESECT)         # first DIFAT sector (none)
    struct.pack_into('<I', hdr, 72, 0)                # DIFAT sector count
    struct.pack_into('<I', hdr, 76, 0)                # DIFAT[0] = sector 0 (our FAT)
    for k in range(1, 109):
        struct.pack_into('<I', hdr, 76 + k * 4, FREESECT)

    # ── Assemble file ─────────────────────────────────────────────────────────
    out = bytes(hdr) + fat_bytes + dir_bytes
    for key in stream_order:
        out += _pad(streams.get(key, b''))

    return out


# ── Main ──────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    HERE = os.path.dirname(os.path.abspath(__file__))
    template = os.path.join(HERE, 'vba_project.bin')
    vba_txt  = os.path.join(HERE, 'VBA_MODULES.txt')
    out_path = os.path.join(HERE, 'vba_project.bin')

    with open(vba_txt, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # Use only from 'Option Explicit' onwards (strips Unicode-heavy header comments)
    start = next(i for i, l in enumerate(lines) if l.strip() == 'Option Explicit')
    code = ''.join(lines[start:])
    code = code.encode('ascii', errors='replace').decode('ascii')
    code = code.replace('\r\n', '\n').replace('\n', '\r\n')   # normalise CRLF

    data = build_vba_project_bin(code, template)
    with open(out_path, 'wb') as f:
        f.write(data)
    print(f'Written: {out_path}  ({len(data):,} bytes)')

    # Quick verification
    print('Verifying streams:')
    with olefile.OleFileIO(out_path) as ole:
        for parts in ole.listdir():
            path = '/'.join(parts)
            sz = ole.get_size(path)
            tag = ' *** NEW ***' if path == 'VBA/Module1' else ''
            print(f'  {path}: {sz:,} bytes{tag}')
