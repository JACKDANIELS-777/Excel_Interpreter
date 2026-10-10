# Excel Interpreter

A small programming language that runs **inside Excel**, written from scratch in about **1,800 lines of VBA**:
a hand-written lexer, a recursive interpreter, and a bridge out to Python for things VBA can't do on its own.

Programs are plain text, typed into a text box on the first sheet and run with a button. They can read and write cells, branch, loop, define
variables and functions, format the sheet, and even generate new VBA functions while running.

## Example

```
if_statement_cell_val;
a2;=;10;[
    print;b1;10;
    if_statement_cell_val;
    b1;=;10;[print;c1;10;];
];
```

If `A2` is 10, write 10 to `B1`. Then, inside the same block, if `B1` is now 10, write 10 to `C1`.
Blocks in `[ ]` nest to any depth.

Functions with parameters:

```
function;show;(a,b);[print_var;D1;a;print_var;D2;b;];
call_function;show;(10,20);
```

`call_function` binds `a = 10` and `b = 20` as variables, then runs the body, so `D1` gets 10 and `D2` gets 20.

## How it works

1. **Lexer** (`lex`). Splits the source on `;` into tokens. A `[ ... ]` block or a `'...'` string becomes a
   single `statement` token holding its raw text. A bracket stack makes sure nested blocks stay whole.
2. **Interpreter** (`parser`). Walks the tokens. Each keyword reads its arguments from the next few tokens. When
   it reaches a block it calls `parser(lex(block))` again, so nesting is plain recursion.
3. **Expressions** are handed to Excel's own `Evaluate`, so anything Excel can calculate (`=A1*2`, `=SUM(A1:A5)`)
   works as a value. Language variables aren't visible to `Evaluate`, so expressions work on cells and numbers, and
   variables are read with `print_var`.

## What the language can do

| Area | Keywords |
|---|---|
| Output | `print`, `msg`, `msg_cell`, `check_cell_value` |
| Conditions | `if_statement_cell_val` (runs a block), `if_print_cell_val`, `if_print_val_cell`, `if_else_print_cell_val`, `if_else_print_cells`, `if_else_print_vals`. Operators `=`, `>`, `<`, `<>` |
| Loops | `for_each;cell;in;A1:A10;print;cell;+;5;`, `for_inc_cell_by_var` |
| Variables | `define_var`, `print_var`, `get_var` |
| Control flow | `switch` ... `end switch`, `function` / `call_function` with parameters |
| Strings | `print_concat`, `print_substr`, `print_len_cell`, `append_val`, `replace_val` |
| Ranges | `clear_cell`, `clear_range`, `add_range_dup`, `add_range_inc` |
| Formatting | `format;bold`, `format;number`, `align`, `highlight_bg_cell_color`, `highlight_fg_cell_color`, `merge`, `unmerge`, `add_hyperlink` |
| Sheets | `add_sheet`, `rename_sheet`, `del_sheet` |
| Input | `input` (prompt the user, store the answer in a cell) |
| Files | `load_run` (load a program from a text file and run it) |
| Meta | `lambda` and `mask_expr` write **new VBA functions into the workbook at runtime** through the VBA project object model, which can then be used from cells like built-in formulas |
| Python bridge | `random`, `py_input`, `py_input_eval`, `python` (a Python REPL), `create_tk` (a Tkinter window), `tk` (an editor that colours keywords as you type), `show_custom_msgbox` |

## Running it

1. Open `dostuffxlsx.xlsm` in Excel on Windows and enable macros.
2. For `lambda` and `mask_expr`: File → Options → Trust Center → Macro Settings →
   tick **Trust access to the VBA project object model**.
3. Type a program into the text box on the first sheet and click the button. Results appear on that sheet.

The Python features call a local Python install, and the paths to it are currently hard-coded in the VBA
(`RunPythonScript`, `CreateCustomTkForm`, `show_custom_msgbox`). Change those paths to match your machine before
using them.

## Code

| File | What's in it |
|---|---|
| `Code/a.vba` | `Module1`: the lexer, the interpreter and the Python bridge helpers |
| `Code/b.vba` | `Module2`: the custom Tkinter message box |
| `Code/c.vba` | Sheet code: the first version of the interpreter, kept for reference |

The workbook's `Sheet1` code (not exported) is the entry point: its button runs `parser lex(TextBox1.Text)`.

## Status

A personal experiment, not production code. It works for the examples above. Known rough edges: hard-coded
paths, and a few debug message boxes still left in.
