"""Minimal MFS (400K Macintosh File System) extractor: python3 mfs.py image outdir"""
import struct, sys, os
img = open(sys.argv[1], 'rb').read(); out = sys.argv[2]; os.makedirs(out, exist_ok=True)
m = img[1024:]
sig, cr, bk, atrb, nfiles, dirst, dirlen, nblocks, blksz, clp, alst, nxt, free = struct.unpack('>HIIHHHHHIIHIH', m[:36])
assert sig == 0xD2D7
vname = m[37:37 + m[36]].decode('mac_roman')
# 12-bit block map after the 64-byte MDB
mp = m[64:]
def entry(n):  # n >= 2
    i = n - 2; o = i * 3 // 2
    v = (mp[o] << 8) | mp[o + 1]
    return (v >> 4) if i % 2 == 0 else (v & 0xFFF)
def fork(start, length):
    data = bytearray(); b = start; seen = 0
    while b not in (0, 1) and len(data) < length and seen < nblocks + 2:
        off = alst * 512 + (b - 2) * blksz
        data += img[off:off + blksz]; b = entry(b); seen += 1
    return bytes(data[:length])
print(f'volume {vname!r}: {nfiles} files, alloc block {blksz}')
d = img[dirst * 512:(dirst + dirlen) * 512]
for sec in range(dirlen):
    p = sec * 512
    while p < (sec + 1) * 512 - 51:
        flags = d[p]
        if not flags & 0x80: break
        typ, cre = d[p + 2:p + 6].decode('mac_roman'), d[p + 6:p + 10].decode('mac_roman')
        num, dst, dlen, dpy, rst, rlen, rpy, crd, mdd = struct.unpack('>IHIIHIIII', d[p + 18:p + 50])
        nl = d[p + 50]; name = d[p + 51:p + 51 + nl].decode('mac_roman')
        safe = name.replace('/', ':')
        open(os.path.join(out, safe), 'wb').write(fork(dst, dlen))
        if rlen: open(os.path.join(out, safe + '.rsrc'), 'wb').write(fork(rst, rlen))
        print(f'  {typ}/{cre} {dlen:7d} {rlen:7d}  {name}')
        p += 51 + nl; p += p & 1
