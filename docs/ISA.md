# ISA v3 — 404 Angels CPU

Source of truth for **both** the CPU decoder (`modCore`, Alicia) and the
assembler (`modAssembler`, Paola). Any change here: PR approved by both.

> Status: **DRAFT — agree on it in Sprint 0 before starting the assembler.**

## Carried over from Parcial 1 (ISA v2)

Register codes: `AX=00`, `BX=01`, `CX=10`, `DX=11`.

| Opcode | Instruction | Bytes |
|---|---|---|
| `00` | NOP | 1 |
| `1X–7X` | MOV/ADD/SUB/CMP/AND/OR/XOR reg,reg (high nibble = op, low nibble = `dd ss`) | 1 |
| `80–9B` | MOV…XOR reg,imm8 | 2 |
| `A8–AB` | LOAD reg,[BX] | 1 |
| `AC–AF` | STORE [BX],reg | 1 |
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

`[BX]` addressing stays 8-bit (pointer inside page 0–7 of the virtual space).

## New in v3: I/O

| Opcode | Instruction | Bytes | Bus signal |
|---|---|---|---|
| `D0–D3` | IN reg,[KBD] | 1 | `En_KBD` |
| `D4–D7` | OUT [LCD],reg | 1 | `Ld_LCD` |

Bits 1–0 = register code.

## Memory map (virtual, proposal)

| Range | Use |
|---|---|
| `000h–0FFh` | Code (pages 0–7) |
| `100h–1DFh` | Data (pages 8–14) |
| `1E0h–1E7h` | Keyboard queue (8 bytes) — page 15 |
| `1E8h–1EFh` | LCD video buffer (8 bytes) — page 15 |

Page 15 is **pinned** to frame 7 (never evicted), so I/O never page-faults.
