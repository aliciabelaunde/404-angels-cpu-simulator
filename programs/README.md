# Program catalog

These are the 3 programs of the catalog (sheet **Programs**, columns K, L and M).
The programs are not inside the VBA code: they live in the cells of the sheet,
and these files keep a copy in the repository.

| # | File | What it shows |
|---|---|---|
| 1 | `01_loops_factorial.asm` | Nested loops with AX, BX, CX and DX: `5! = 120 = 78h` stored at `101h` |
| 2 | `02_keyboard_echo.asm` | `IN` / `OUT`: every char typed in sheet IO is shown on the LCD 8x1 |
| 3 | `03_page_faults.asm` | Scattered accesses to pages 8-14: page faults, FIFO, write back and cache misses. `SUM = 23h` at `1DFh` |

## Put a program in the catalog

1. Open the `.asm` file in VS Code, select everything (Ctrl+A) and copy it (Ctrl+C).
2. In Excel, sheet **Programs**, click cell **K6** (program 1), **L6** (program 2) or **M6** (program 3) and paste (Ctrl+V). Each line goes to its own cell.

## Run a program

1. Sheet **Programs**: press **OPEN 1**, **OPEN 2** or **OPEN 3**. The program is copied to the editor (column B), where it can be changed.
2. Press **ASSEMBLE**. Each line shows its address and bytes; lines with errors turn red.
3. Press **LOAD PROGRAM**. The program goes to virtual memory (swap) and the whole computer is reset.
4. Sheet **CPU**: press **STEP** or **RUN**.

## Assembler syntax

```
LABEL:  MNEMONIC operand1, operand2   ; comment
```

- Labels end with `:`. Comments start with `;`. Upper or lower case.
- Numbers: decimal (`12`) or hexadecimal ending in `h` (`0Ch`, `1E0h`).
- `ORG address`: the next byte goes to that address (`ORG 100h` for the data area).
- `DB value, value, ...`: data bytes (0-255).
- Instructions: see `docs/ISA.md`. Memory operands use brackets: `LOAD AX, [N]`, `STORE [1DFh], DX`, `LOAD AX, [BX]`.
