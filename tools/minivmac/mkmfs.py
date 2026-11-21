"""Build an MFS (Macintosh File System) floppy image.

usage: mkmfs.py OUT.dsk SIZE_K VOLNAME  name=data_path[,rsrc_path]:TYPE:CRTR ...
"""
import struct, sys, time

def main():
    out, size_k, vname, specs = sys.argv[1], int(sys.argv[2]), sys.argv[3], sys.argv[4:]
    sectors = size_k * 2
    blk = 1024 if size_k <= 400 else 2048          # keep the block map inside sectors 2-3
    dir_start, dir_len = 4, 12
    first = dir_start + dir_len                    # first allocation-block sector
    nblocks = (sectors - first) * 512 // blk
    assert 64 + (nblocks * 3 + 1) // 2 <= 1024, 'block map does not fit'
    img = bytearray(sectors * 512)
    bmap = [0] * nblocks                           # index i is allocation block i + 2
    nextfree = 0
    now = int(time.time()) + 2082844800            # seconds since 1904

    def alloc(data):
        nonlocal nextfree
        if not data:
            return 0, 0
        count = (len(data) + blk - 1) // blk
        start = nextfree
        assert start + count <= nblocks, 'disk full'
        for i in range(count):
            bmap[start + i] = (start + i + 3) if i < count - 1 else 1
            off = (first * 512) + (start + i) * blk
            chunk = data[i * blk:(i + 1) * blk]
            img[off:off + len(chunk)] = chunk
        nextfree += count
        return start + 2, count * blk

    entries = []
    for num, spec in enumerate(specs, start=1):
        name, rest = spec.split('=', 1)
        paths, ftype, creator = rest.rsplit(':', 2)
        dpath, _, rpath = paths.partition(',')
        data = open(dpath, 'rb').read() if dpath else b''
        rsrc = open(rpath, 'rb').read() if rpath else b''
        dstart, dphys = alloc(data)
        rstart, rphys = alloc(rsrc)
        nm = name.encode('mac_roman')
        location = (40 << 16) | (30 + 90 * ((num - 1) % 4))   # icon position (v, h) in the disk window
        e = struct.pack('>BB4s4sHIHIHIIHIIII', 0x80, 0, ftype.encode('mac_roman'), creator.encode('mac_roman'),
                        0, location, 0, num, dstart, len(data), dphys, rstart, len(rsrc), rphys, now, now)
        e += bytes([len(nm)]) + nm
        if len(e) % 2:
            e += b'\0'
        entries.append(e)

    # directory: entries never cross a sector boundary
    d = bytearray(dir_len * 512); p = 0
    for e in entries:
        if p // 512 != (p + len(e) - 1) // 512:
            p = (p // 512 + 1) * 512
        d[p:p + len(e)] = e; p += len(e)
    img[dir_start * 512:(dir_start + dir_len) * 512] = d

    vn = vname.encode('mac_roman')
    mdb = struct.pack('>HIIHHHHHIIHIH', 0xD2D7, now, 0, 0, len(entries), dir_start, dir_len,
                      nblocks, blk, 8 * blk, first, len(entries) + 1, nblocks - nextfree)
    mdb += bytes([len(vn)]) + vn
    img[1024:1024 + len(mdb)] = mdb
    m = bytearray()
    for i in range(0, nblocks, 2):
        a = bmap[i]; b = bmap[i + 1] if i + 1 < nblocks else 0
        m += bytes([a >> 4, ((a & 15) << 4) | (b >> 8), b & 255])
    img[1024 + 64:1024 + 64 + len(m)] = m[:(nblocks * 3 + 1) // 2]
    open(out, 'wb').write(img)
    print(f'{out}: {len(entries)} files, {nextfree}/{nblocks} blocks used')

main()
