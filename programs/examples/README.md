# Example programs

Small programs, one idea each, so every instruction type can be tested on its own.
Each file starts with a comment that says the input, the expected output and which sheet to watch.

The 4 programs marked with ★ are the ones for the presentation. They show their result on the LCD or on the sheets, so the result can be seen without reading the registers.

## How to run an example

1. Open the `.asm` file in Notepad or VS Code and copy everything.
2. In sheet **Programs**, click cell **B11** (first line of the editor) and paste.
3. Press **ASSEMBLE** and then **LOAD PROGRAM**.
4. For E12 and E13: go to sheet **IO**, type the text in **B5** and press **SEND** before RUN.
5. In sheet **CPU**, use **STEP INSTR**, or **RUN** with a *Run delay* of 100 to 300 ms so every pulse is drawn. With *Run delay* = 0 RUN is fast but only repaints every 25 pulses.

The 3 catalog programs (`programs/01`, `02`, `03`) are still loaded with OPEN 1, 2 and 3.

## List of examples

All results below were checked by running each program in the simulator and comparing it with a reference emulator written from `docs/ISA.md` (registers, flags, memory, LCD and the one-driver-per-bus rule after every instruction).

### Registers and ALU

| # | File | Shows | Input | Output |
|---|---|---|---|---|
| E01 | `e01_mov.asm` | `MOV` reg, imm and reg, reg | 12h, 34h | AX = CX = 12h, BX = DX = 34h |
| ★ E02 | `e02_add_sub_flags.asm` | `ADD` / `SUB` and the flags | AX = 0F0h, BX = 20h | carry (CF = 1, AX = 10h), borrow (CF = 1, SF = 1, AX = 0F0h), zero (ZF = 1, AX = 0) |
| E03 | `e03_cmp.asm` | `CMP` changes only the flags | AX = 7 compared with 7, 9, 3 | ZF = 1, then CF = 1, then both 0. AX stays 7 |
| E04 | `e04_logic.asm` | `AND`, `OR`, `XOR`, `NOT` | AX = 0B4h | BX = 04h, CX = 0F4h, DX = 4Bh, AX = 4Bh |
| E05 | `e05_inc_dec.asm` | `INC` / `DEC` keep CF | AX = 0FFh with CF = 1 | AX = 0 (ZF = 1), then 0FFh (SF = 1). CF stays 1 |

### Memory and jumps

| # | File | Shows | Input | Output |
|---|---|---|---|---|
| E06 | `e06_load_store.asm` | `LOAD` / `STORE [addr9]` | 25 at 100h, 17 at 101h | 42 = 2Ah at 102h |
| ★ E07 | `e07_bx_array.asm` | `[BX]` addressing and a loop | array 3, 1, 4, 1, 5 at 100h | 0Eh at 105h, LCD shows `SUM=14` |
| E08 | `e08_loop_jnz.asm` | loop with `DEC` + `JNZ` | N = 10 | AX = 55 = 37h |
| E09 | `e09_max_jc.asm` | `JC` after `CMP` | X = 47, Y = 83 | MAX = 83 = 53h at 102h |
| E10 | `e10_sign_js.asm` | `JS` / `JNS` and two's complement | 5 − 9 | AX = 0FCh (−4), SF = 1, DX = 1 |

### Keyboard and LCD

| # | File | Shows | Input | Output |
|---|---|---|---|---|
| E11 | `e11_out_lcd.asm` | `OUT [LCD]` | ASCII codes in the program | LCD shows `HI 404` |
| E12 | `e12_in_one_char.asm` | `IN [KBD]` with polling | type `Q` + SEND | AX = 51h, LCD shows `Q` |
| ★ E13 | `e13_uppercase.asm` | keyboard → ALU → LCD | type `hola` + SEND | LCD shows `HOLA` |
| E14 | `e14_lcd_scroll.asm` | LCD scroll and backspace | digits 0 to 9, then 08h | LCD shows `2345678 ` |

### Virtual memory, swap and cache

| # | File | Shows | Input | Output |
|---|---|---|---|---|
| E15 | `e15_page_faults.asm` | demand paging, no replacement | reads from pages 8, 9, 10 | 4 page faults; the second read of page 8 has none |
| E16 | `e16_fifo_swap.asm` | FIFO when RAM is full | reads from pages 8 to 14 | 9 page faults; page 0 (code) and then page 8 leave; no write back |
| ★ E17 | `e17_swap_write_back.asm` | swap with write back | 77h written in page 8, then 6 other pages read | page 8 leaves with write back and comes back from swap: BX = 77h, LCD shows `OK`. 11 page faults |
| E18 | `e18_cache_hits.asm` | cache hits vs misses | 8 reads in one line, 4 reads 8 bytes apart | L1D: 7 hits and 5 misses |

## Why these 4 for the presentation

- **E02** shows the ALU and the three flags changing in only 6 instructions.
- **E07** uses `[BX]`, a loop, memory and the LCD in one short program.
- **E13** is the clearest input/output example: lower case comes in from the keyboard and upper case goes out on the LCD.
- **E17** shows the full swap cycle that was asked for: page fault, FIFO victim, write back to swap, and the value coming back. The `OK` on the LCD proves the data survived.

## What to say for E17 (swap)

1. `STORE [100h]` writes 77h in page 8, so its Dirty bit becomes 1.
2. Pages 9 to 13 fill the RAM. Page 14 needs a frame, so FIFO takes out page 0 (the code). It was not modified, so nothing is written back.
3. The next instruction is in page 0, so it comes back and FIFO now takes out page 8. Page 8 has Dirty = 1, so it is copied to swap first (the WRITE BACK arrow turns orange).
4. `LOAD BX, [100h]` causes a page fault and page 8 comes back from swap with 77h. The program checks it and writes `OK` on the LCD.
