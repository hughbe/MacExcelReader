#!/usr/bin/env python3
"""Disassembles the p-code in Excel 1.x's PCOD resources (spec, "Background: how Excel 1.x runs").

    python3 -I tools/pcode.py EXCEL.rsrc              # every segment
    python3 -I tools/pcode.py EXCEL.rsrc 18           # one segment
    python3 -I tools/pcode.py EXCEL.rsrc 18.12 18.13  # procedures

EXCEL.rsrc is the application's resource fork (the `PCOD` 0 and 2 resources are read from it), or a directory holding
PCOD_00.bin and PCOD_02.bin. Native 68000 code inside procedures is disassembled with Capstone when it is installed
(`pip install capstone`); otherwise it is shown as hex.

The opcodes come from the interpreter in Excel 1.00's `CODE` 1 resource: its main loop reads an opcode byte and jumps
through a 256-entry table of handlers at `CODE` 1 offset 0x2F7C, each relative to offset 0x2F6A. Notes on the encoding:

* A segment's image starts with a word giving the end of its import table (A3), a word the interpreter uses as a
  least-recently-used stamp, the number of procedures and their offsets. The imports run down from A3: call operands
  index them (`seg.proc` words). Above A3, up to the first procedure, is a table of global offsets that global operands
  index. One-byte forms of the commonest calls and globals index the first entries of those tables.
* A procedure starts with a word: its frame size, or with bit 15 set, a procedure written in 68000 code.
* `native` (0xE7) runs the 68000 code that follows. It calls Toolbox traps with `jsr -$3a(a4)` and the trap word, and
  returns to p-code with `jsr -$36(a4)` plus two bytes, or `jsr -$16(a4)`.
* `gosub`/`endsub` are the compiler's shared sequences: `gosub` runs the instructions at its target up to the next
  `endsub` and comes back; elsewhere `endsub` does nothing.
* Locals are numbered from the frame pointer down (L0 is the word at the frame pointer), globals are offsets into
  Excel's global data (`PCOD` 0).
"""
import os
import struct
import sys

# --- resources -------------------------------------------------------------------------------------------------------


def read_resources(path):
    if os.path.isdir(path):
        return {n: open(os.path.join(path, 'PCOD_%02d.bin' % n), 'rb').read() for n in (0, 2)}
    data = open(path, 'rb').read()
    data_offset, map_offset = struct.unpack('>II', data[:8])
    rmap = data[map_offset:]
    type_list = struct.unpack('>H', rmap[24:26])[0]
    types = struct.unpack('>H', rmap[type_list:type_list + 2])[0] + 1
    found = {}
    for i in range(types):
        kind, count, refs = struct.unpack('>4sHH', rmap[type_list + 2 + 8 * i:type_list + 10 + 8 * i])
        if kind != b'PCOD':
            continue
        for j in range(count + 1):
            rid, _, attrs = struct.unpack('>hhI', rmap[type_list + refs + 12 * j:type_list + refs + 12 * j + 8])
            offset = data_offset + (attrs & 0xFFFFFF)
            length = struct.unpack('>I', data[offset:offset + 4])[0]
            found[rid] = data[offset + 4:offset + 4 + length]
    return found


class Program:
    def __init__(self, pcod0, pcod2):
        self.code = pcod2
        starts = []
        offset = 0x38
        while offset + 4 <= len(pcod0):
            start = struct.unpack('>I', pcod0[offset:offset + 4])[0]
            if start == 0 or start + 6 >= len(pcod2) or (starts and start <= starts[-1]):
                break     # past the last segment: the table ends with the end of the code
            starts.append(start)
            offset += 4
        self.segments = [Segment(self, n, s) for n, s in enumerate(starts)]


class Segment:
    def __init__(self, program, number, start):
        self.number = number
        self.base = start + 6            # addresses in the segment count from here
        self.size = struct.unpack('>H', program.code[start + 4:start + 6])[0]
        self.image = program.code[self.base:self.base + self.size]
        self.a3 = self.word(0)
        count = self.word(4)
        self.procs = [self.word(6 + 2 * i) for i in range(count)]

    def word(self, address):
        return struct.unpack('>H', self.image[address:address + 2])[0] if address + 2 <= len(self.image) else 0

    def byte(self, address):
        return self.image[address] if address < len(self.image) else 0

    def global_(self, index):
        return 'g%04x' % self.word(self.a3 + 2 * index)

    def import_(self, index):
        w = self.word(self.a3 - 2 - 2 * index)
        return '%02x.%02x' % (w >> 8, w & 0xFF)


# --- opcodes ---------------------------------------------------------------------------------------------------------

SYS = {0x00: 'callind', 0x05: 'bmovefarL', 0x06: 'longjmp', 0x07: 'fld', 0x08: 'fst', 0x09: 'str2num', 0x0A: 'num2str',
       0x0B: 'fadd', 0x0C: 'fsub', 0x0D: 'fmul', 0x0E: 'fdiv', 0x0F: 'fsign', 0x10: 'fneg', 0x11: 'fzero', 0x12: 'fld_int',
       0x13: 'f2int', 0x14: 'f2long', 0x15: 'fabs', 0x16: 'fint', 0x17: 'muldiv', 0x18: 'exit', 0x19: 'memA', 0x1A: 'memB',
       0x1B: 'segload', 0x1D: 'seglock', 0x1F: 'mkthunk', 0x20: 'cbret', 0x21: 'ostrap', 0x22: 'calljt', 0x23: 'calljtW',
       0x24: 'calljtL', 0x25: 'callptr', 0x26: 'callptrW', 0x27: 'callptrL', 0x28: 'intl0', 0x29: 'intl1', 0x2A: 'memC',
       0x2B: 'memD', 0x2C: 'xchg24c', 0x2D: 'segunlock', 0x2E: 'setjmp', 0x2F: 'segalloc', 0x30: 'segfree'}

# 0xE6 prefix: name and the number of operand bytes.
EXT = {0x01: ('xor', 0), 0x02: ('ldmulti', 1), 0x03: ('stmulti', 1), 0x04: ('lmul', 0), 0x05: ('ldiv', 0), 0x06: ('sar', 0),
       0x07: ('ldwT', 0), 0x08: ('dupl', 0), 0x0A: ('mul_wl', 0), 0x0C: ('fillw', 0), 0x0D: ('fillb', 0),
       0x0E: ('bmove_d', 0), 0x0F: ('bmove_far', 0), 0x10: ('stb_far+', 1), 0x11: ('stw_far+', 1), 0x12: ('stl_far+', 1),
       0x13: ('ldl_so+', 1), 0x14: ('ldb_so_idx', 0), 0x15: ('stb_so_idx', 0), 0x16: ('ldw_so_idx', 0),
       0x17: ('stw_so_idx', 0), 0x18: ('so2far', 0), 0x19: ('far2so', 0), 0x1A: ('st3_so+', 1), 0x1B: ('ulcmp', 0),
       0x1D: ('bfins_far', 1), 0x1E: ('bfins_so', 2), 0x20: ('bmove_1f', 0), 0x21: ('bmove_so', 0), 0x25: ('pushDSseg', 0),
       0x26: ('stl_so+', 1), 0x27: ('trapL', 1)}

SIMPLE = {0x00: 'nop', 0x56: 'ldw', 0x78: 'dup', 0x7A: 'add', 0x7B: 'sub', 0x7C: 'ldl', 0x7D: 'stl', 0x7F: 'endsub',
          0x85: 'stw', 0x86: 'stw+2', 0x87: 'stw+4', 0x88: 'stb', 0x89: 'stw_keep', 0xA7: 'ldb', 0xAA: 'pushT',
          0xAB: 'popT', 0xAC: 'xload', 0xAD: 'stwT', 0xAE: 'ldbT', 0xAF: 'stbT', 0xD2: 'dec', 0xD3: 'and', 0xD4: 'or',
          0xD5: 'shl', 0xD6: 'shl1', 0xD7: 'shr', 0xD8: 'mul', 0xD9: 'div', 0xDA: 'drop', 0xDB: 'neg', 0xDD: 'ret',
          0xDF: 'cmp3', 0xE0: 'swapl', 0xE1: 'addidx2', 0xE2: 'lcmp', 0xE3: 'ladd', 0xE4: 'lsub', 0xE5: 'ext',
          0xE8: 'setDS', 0xE9: 'resetDS', 0xEA: 'tofar', 0xEB: 'tosegoff', 0xEF: 'ladd_w', 0xF0: 'segoff_add',
          0xFF: 'native!'}

PUSH = {0x60: 0, 0x61: 2, 0x62: 4, 0x63: 6, 0x64: 8, 0x65: 10, 0x66: -1, 0x67: 1, 0x68: 3, 0x69: 5, 0x6A: 7, 0x6B: 9}
ADDC = {0x6D: 2, 0x6E: 4, 0x6F: 6, 0x70: 8, 0x71: 10, 0x72: 12, 0x73: 14, 0x74: 16, 0x75: 1, 0x76: 3}
LDWO = {0x57: 2, 0x58: 4, 0x59: 6, 0x5A: 8, 0x5B: 10, 0x5C: 12, 0x5D: 14}
LDBO = {0xA5: 1, 0xA6: 3, 0xA8: 2, 0xA9: 4}
CALLF1 = {0x8F: 0, 0x8E: 1, 0x8D: 2, 0x8C: 3, 0x8B: 4, 0x84: 11, 0x83: 12, 0x82: 13, 0x81: 14}
CALL1 = {0x9A: 5, 0x99: 6, 0x98: 7, 0x97: 8, 0x96: 9, 0x95: 10, 0x94: 11, 0x93: 12, 0x92: 13, 0x91: 14}
BRANCH = {0xB8: 'jeq', 0xB9: 'jz', 0xC2: 'jne', 0xC8: 'jnz', 0xCE: 'jle', 0xCF: 'jlt', 0xD0: 'jgt', 0xD1: 'jge'}
# One-byte jumps: how far ahead of the opcode each goes.
SHORT = {op: (name, d) for name, first, distances in (('jmp', 0xB2, (3, 5, 7, 2, 4, 6)), ('jz', 0xBA, (3, 5, 7, 9, 2, 4, 6, 8)),
                                                      ('jne', 0xC3, (3, 5, 2, 4, 6)), ('jnz', 0xC9, (3, 5, 2, 4, 6)))
         for op, d in zip(range(first, first + len(distances)), distances)}
FAR = {0xEC: 'ldb_far+', 0xED: 'ldw_far+', 0xEE: 'ldl_far+', 0xF1: 'stb_so+', 0xF2: 'stw_so+', 0xF4: 'ldb_so+',
       0xF5: 'ldw_so+', 0xF6: 'ld3_so+'}
TRAP = {0xF9: ('trap', 0xA800), 0xFA: ('trapW', 0xA800), 0xFC: ('trap', 0xA900), 0xFD: ('trapW', 0xA900),
        0xFE: ('trapL', 0xA900)}


def s16(v):
    return v - 0x10000 if v & 0x8000 else v


def decode(seg, a):
    """Decodes the p-code instruction at a: (length, text, switches to 68000 code)."""
    b = seg.byte
    op = b(a)
    if op in SIMPLE:
        return 1, SIMPLE[op], False
    if op == 0xE7:
        return 1, 'native', True
    if op == 0x01:
        return 2, 'pushL L%d' % b(a + 1), False
    if op == 0x02:
        return 3, 'pushL L%d' % seg.word(a + 1), False
    if 0x03 <= op <= 0x16:
        return 1, 'pushL L%d' % (0x16 - op), False
    if op == 0x17:
        return 2, 'lea L%d' % b(a + 1), False
    if op == 0x18:
        return 3, 'lea L%d' % seg.word(a + 1), False
    if 0x19 <= op <= 0x22:
        return 1, 'lea L%d' % (0x22 - op), False
    if op == 0x23:
        return 2, 'popL L%d' % b(a + 1), False
    if op == 0x24:
        return 3, 'popL L%d' % seg.word(a + 1), False
    if 0x25 <= op <= 0x33:
        return 1, 'popL L%d' % (0x33 - op), False
    if op == 0x34:
        return 2, 'postinc L%d' % b(a + 1), False
    if 0x35 <= op <= 0x3D:
        return 1, 'postinc L%d' % (0x3D - op), False
    if op == 0x3E:
        return 2, 'pushG ' + seg.global_(b(a + 1)), False
    if 0x3F <= op <= 0x4E:
        return 1, 'pushG ' + seg.global_(op - 0x3F), False
    if op == 0x4F:
        return 2, 'leaG ' + seg.global_(b(a + 1)), False
    if 0x50 <= op <= 0x54:
        return 1, 'leaG ' + seg.global_(op - 0x50), False
    if op == 0x55:
        return 2, 'ldw+ %d' % (2 * b(a + 1)), False
    if op in LDWO:
        return 1, 'ldw+ %d' % LDWO[op], False
    if op == 0x5E:
        return 2, 'push %d' % b(a + 1), False
    if op == 0x5F:
        return 3, 'push 0x%04x' % seg.word(a + 1), False
    if op in PUSH:
        return 1, 'push %d' % PUSH[op], False
    if op == 0x6C:
        return 2, 'addc %d' % b(a + 1), False
    if op in ADDC:
        return 1, 'addc %d' % ADDC[op], False
    if op in (0x77, 0x8A):
        v = b(a + 1)
        return 2, '%s shift=%d bits=%d' % ('bfext' if op == 0x77 else 'bfins', v & 0xF, (v >> 4) + 1), False
    if op == 0x79:
        return 2, 'jmp %04x' % (a + 2 + b(a + 1) - 0x100), False     # always back
    if op == 0x7E:
        return 2, 'gosub %04x' % (a + 2 - b(a + 1)), False
    if op == 0x80:
        return 2, 'callF ' + seg.import_(b(a + 1)), False
    if op in CALLF1:
        return 1, 'callF ' + seg.import_(CALLF1[op]), False
    if op == 0x90:
        return 2, 'call ' + seg.import_(b(a + 1)), False
    if op in CALL1:
        return 1, 'call ' + seg.import_(CALL1[op]), False
    if op == 0x9B:
        return 2, 'popG ' + seg.global_(b(a + 1)), False
    if 0x9C <= op <= 0xA3:
        return 1, 'popG ' + seg.global_(op - 0x9C), False
    if op == 0xA4:
        return 2, 'ldb+ %d' % b(a + 1), False
    if op in LDBO:
        return 1, 'ldb+ %d' % LDBO[op], False
    if op == 0xB0:
        return 2, 'jmp %04x' % (a + 2 + b(a + 1)), False
    if op == 0xB1:
        return 3, 'jmp %04x' % ((a + 3 + s16(seg.word(a + 1))) & 0xFFFF), False
    if op in SHORT:
        name, distance = SHORT[op]
        return 1, '%s %04x' % (name, a + distance), False
    if op in BRANCH:
        return 2, '%s %04x' % (BRANCH[op], a + 2 + b(a + 1)), False
    if op == 0xDC:
        v = b(a + 1)
        name = SYS.get(v & 0x7F, '0x%02x' % (v & 0x7F))
        return 2, 'sys %s%s' % (name, ' (keep)' if v & 0x80 else ''), False
    if op == 0xDE:
        return switch(seg, a)
    if op == 0xE6:
        sub = b(a + 1)
        name, n = EXT.get(sub, ('ext 0x%02x' % sub, 0))
        if sub == 0x1E:
            return 4, '%s +%d %02x' % (name, b(a + 2), b(a + 3)), False
        return 2 + n, name + (' %d' % b(a + 2) if n else ''), False
    if op in FAR:
        return 2, '%s %d' % (FAR[op], b(a + 1)), False
    if op == 0xF3:
        return 2, 'pushGL ' + seg.global_(b(a + 1)), False
    if op == 0xF7:
        return 2, 'pushLL L%d' % b(a + 1), False
    if op == 0xF8:
        return 2, 'popLL L%d' % b(a + 1), False
    if op == 0xFB:
        return 2, 'popGL ' + seg.global_(b(a + 1)), False
    if op in TRAP:
        name, base = TRAP[op]
        return 2, '%s %04X' % (name, base | b(a + 1)), False
    return 1, 'db 0x%02x' % op, False


def switch(seg, a):
    """The interpreter's switch (0xDE): clauses, each a header byte and its table, until one with bit 6 set.

    Bit 7 set: values lo to lo + (header & 0x3F) with a byte offset each. Bit 5 set: values (header & 0x1F) to that
    plus the next byte, with a word offset each. Otherwise one value (a word) with a 14-bit offset from the header and
    the next byte. Offsets count from the byte after the one holding them; an unmatched value falls through.
    """
    p = a + 1
    cases = []
    while p < len(seg.image):    # bytes that are not code may never end the switch
        h = seg.byte(p)
        if h & 0x80:
            count, lo = h & 0x3F, seg.byte(p + 1)
            for i in range(count + 1):
                q = p + 2 + i
                cases.append((lo + i, q + 1 + seg.byte(q)))
            p += 3 + count
        elif h & 0x20:
            lo, count = h & 0x1F, seg.byte(p + 1)
            for i in range(count + 1):
                q = p + 2 + 2 * i
                cases.append((lo + i, (q + 2 + seg.word(q)) & 0xFFFF))
            p += 4 + 2 * count
        else:
            offset = ((h << 8) | seg.byte(p + 1)) & 0x3FFF
            cases.append((seg.word(p + 2), (p + 4 + offset) & 0xFFFF))
            p += 4
        if h & 0x40:
            break
    return p - a, 'switch ' + ' '.join('%d->%04x' % c for c in cases), False


# --- 68000 code ------------------------------------------------------------------------------------------------------

try:
    import capstone
    _md = capstone.Cs(capstone.CS_ARCH_M68K, capstone.CS_MODE_M68K_000)
except ImportError:
    _md = None


def native(seg, a, end):
    """Yields (address, text) for 68000 code from a until it returns to p-code or reaches end; returns the resume
    address, or None at end."""
    while a < end:
        w = seg.word(a)
        if w == 0x4EAC:
            target = seg.word(a + 2)
            if target == 0xFFC6:
                yield a, 'trap %04X' % seg.word(a + 4)
                a += 6
                continue
            if target == 0xFFCA:
                yield a, '-> pcode (adj=%d, seg=%02x)' % (seg.byte(a + 4), seg.byte(a + 5))
                return a + 6
            if target == 0xFFEA:
                yield a, '-> pcode resume'
                return a + 4
        insn = None
        if _md is not None:
            for i in _md.disasm(bytes(seg.image[a:min(a + 10, end)]), a, 1):
                insn = i
        if insn is None or insn.size == 0:
            yield a, 'dc.w $%04x' % w
            a += 2
        else:
            yield a, ('%s %s' % (insn.mnemonic, insn.op_str)).strip()
            a += insn.size
    return None


# --- listing ---------------------------------------------------------------------------------------------------------

def proc_ranges(seg):
    starts = sorted(set(p for p in seg.procs if p < len(seg.image)))
    for i, p in enumerate(seg.procs):
        later = [s for s in starts if s > p]
        yield i, p, (later[0] if later else len(seg.image))


def list_proc(seg, number, start, end, out):
    frame = seg.word(start)
    out.write('--- proc %02x.%02x @%04x frame=%04x\n' % (seg.number, number, start, frame))
    a = start + 2
    in_native = bool(frame & 0x8000)
    while a < end:
        if in_native:
            gen = native(seg, a, end)
            try:
                while True:
                    addr, text = next(gen)
                    out.write('  %04x: [68k] %s\n' % (addr, text))
            except StopIteration as stop:
                if stop.value is None:
                    return
                a = stop.value
                in_native = False
            continue
        length, text, to_native = decode(seg, a)
        out.write('  %04x: %s\n' % (a, text))
        a += length
        in_native = to_native


def main(argv):
    if len(argv) < 2:
        print(__doc__.strip())
        return 2
    res = read_resources(argv[1])
    program = Program(res[0], res[2])
    wanted = argv[2:]
    out = sys.stdout
    for seg in program.segments:
        selected = [w for w in wanted if int(w.split('.')[0], 16) == seg.number]
        if wanted and not selected:
            continue
        procs = None if not wanted or any('.' not in w for w in selected) else {int(w.split('.')[1], 16) for w in selected}
        out.write('=== SEG %02x size=%x A3=%x nprocs=%d\n' % (seg.number, seg.size, seg.a3, len(seg.procs)))
        for number, start, end in proc_ranges(seg):
            if procs is None or number in procs:
                list_proc(seg, number, start, end, out)
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
