# ISA v3 — 404 Angels CPU

Source of truth for **both** the CPU decoder (`modCore`, Alicia) and the
assembler (`modAssembler`, Paola). Any change here: PR approved by both.

> Status: **AGREED** (reviewed by Alicia and Paola, PR #4).

## Carried over from Parcial 1 (ISA v2)

Register codes: `AX=00`, `BX=01`, `CX=10`, `DX=11`.

| Opcode | Instruction | Bytes |
|---|---|---|
| `00` | NOP | 1 |
| `1X–7X` | MOV/ADD/SUB/CMP/AND/OR/XOR reg,reg (high nibble = op, low nibble = `dd ss`) | 1 |
| `80–9B` | MOV…XOR reg,imm8 | 2 |
| `A8–AB` | LOAD reg,[BX] (effective address `100h + BX`) | 1 |
| `AC–AF` | STORE [BX],reg (effective address `100h + BX`) | 1 |
| `B0–BB` | INC / DEC / NOT reg | 1 |
| `FF` | HLT | 1 |

## Changed in v3: 9-bit addresses

Virtual memory is 512 bytes, so an address needs **9 bits** (`000h–1FFh`).
`PC` and `MAR` become 9-bit registers. Data registers stay 8-bit.

Instructions with an absolute address now carry **2 address bytes**
(little-endian: low byte, then high byte `0` or `1`):

| Opcode | Instruction | Bytes |
|---|---|---|
| `A0–A3` | LOAD reg,[addr9] | 3 |
| `A4–A7` | STORE [addr9],reg | 3 |
| `C0–C6` | JMP, JZ, JNZ, JC, JNC, JS, JNS addr9 | 3 |

`[BX]` addressing points to the **data area**: effective address = `100h + BX`
(`BX = 00h` → `100h`, `BX = FFh` → `1FFh`). This way BX can walk through the
data pages 8–15 and never touches the code area.

## New in v3: I/O

| Opcode | Instruction | Bytes | Bus signal |
|---|---|---|---|
| `D0–D3` | IN reg,[KBD] | 1 | `En_KBD` |
| `D4–D7` | OUT [LCD],reg | 1 | `Ld_LCD` |

Bits 1–0 = register code. **IN and OUT do not modify the flags** (same as x86).

## Memory map (virtual)

| Range | Use |
|---|---|
| `000h–0FFh` | Code (pages 0–7) |
| `100h–1DFh` | Data (pages 8–14) |
| `1E0h–1E7h` | Keyboard queue (8 bytes) — page 15 |
| `1E8h–1EFh` | LCD video buffer (8 bytes) — page 15 |

Page 15 is **pinned** to physical frame 7 (`E0h–FFh`): it is always present
and is **never chosen by FIFO**, so I/O never page-faults. FIFO replacement
only uses frames 0–6.

The keyboard queue and the LCD buffer are **stored in RAM** (frame 7):

| Virtual | Physical | Content |
|---|---|---|
| `1E0h–1E7h` | `E0h–E7h` | Keyboard queue, 8 bytes, circular FIFO (see below) |
| `1E8h–1EFh` | `E8h–EFh` | LCD video buffer, 8 bytes (one ASCII char per cell) |

The I/O range `1E0h–1EFh` is **non-cacheable**: the keyboard can change it at
any time, so the cache always reads it directly from RAM.

### Keyboard queue (circular FIFO, as in the class lab)

The 8 chars live in RAM `E0h–E7h`. Three pointers in `modIO` control the queue:
`HEAD` (next char to read), `TAIL` (next free position) and `COUNT` (chars waiting).

| Operation | Condition | Action |
|---|---|---|
| Keyboard types a char | `COUNT < 8` | `RAM[E0h + TAIL] = char`, `TAIL = (TAIL + 1) mod 8`, `COUNT + 1` |
| `IN reg, [KBD]` | `COUNT > 0` | `reg = RAM[E0h + HEAD]`, `HEAD = (HEAD + 1) mod 8`, `COUNT - 1` |
| `IN reg, [KBD]` | `COUNT = 0` | `reg = 0` (queue empty) |
| Keyboard types a char | `COUNT = 8` | The char is discarded (overflow) |

### LCD 8x1

`OUT [LCD], reg` writes the char at the cursor (`E8h + cursor`) and moves the cursor.
When the line is full the text scrolls one position to the left. Char `08h`
(backspace) moves the cursor back and writes a space (`20h`).

