# Automaton code generation tool
This is a copycat of [Automaton Machine Editor](https://github.com/tinkerspy/Automaton-Machine-Editor).
It produces a single header file with inline codes. To generate code, refer to `template.xlsx` and keep the
consistent sheet names. Run `python generate.py path_to_xlsx.xlsx` will write the code to 1) `output path` 
defined in the spreadsheet; 2) `output` argument; or 3) print to `stdout`.

The generate machine should be subclassed and implement the pure virtual functions to use.
This prevents user code being overwritten after update the machine.

