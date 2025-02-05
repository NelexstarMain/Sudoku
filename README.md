# Sudoku Generator

A Python tool that generates Sudoku puzzles of varying difficulty. It creates a complete Sudoku solution, removes cells based on the selected difficulty, and exports both the puzzle and the solution into an Excel file with nicely formatted grids. The puzzles and solutions are also printed to the console.

---

## Features

- **Sudoku Generation:** Uses a backtracking algorithm to generate a complete, valid Sudoku grid.
- **Puzzle Creation:** Removes cells according to difficulty levels (1 to 5) to create the puzzle.
- **Excel Export:** Saves two worksheets in an Excel file:
  - **Puzzle:** Sudoku with removed cells.
  - **Solution:** The complete solution.
- **Formatted Output:** Excel sheets are formatted to display classic Sudoku grids with bold borders separating 3x3 blocks.

---

## Requirements

- Python 3.7+
- [openpyxl](https://openpyxl.readthedocs.io/)

---

## Installation

Install the required dependency with:
```bash
pip install openpyxl
```

---

## Usage

Run the script:
```bash
python sudoku.py
```
You can set the difficulty level (1–5) in the `Sudoku` class constructor:
```python
sudoku = Sudoku(lvl=3)  
```

---

## File Structure

```
sudoku-generator/
├── sudoku.py     # Main script with the Sudoku generator code
└── README.md     # This file
```

---

## License

This project is licensed under the MIT License.

---

Enjoy generating Sudoku puzzles!