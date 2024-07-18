# Excel Sudoku Solver

An advanced Sudoku puzzle solver implemented as a Microsoft Excel spreadsheet, leveraging the power of linear programming techniques to efficiently solve puzzles of varying difficulty. This tool combines Excel's native functionalities with custom VBA macros and scripts to provide a user-friendly interface for inputting unsolved Sudoku grids and obtaining solutions through mathematical optimization methods.

## Features

- User-friendly interface for inputting unsolved Sudoku puzzles
- Powerful solver using linear programming algorithms
- VBA Macro and script integration for enhanced functionality
- Visual representation of the solving process

## How to Use

1. Open the Excel file
2. Enter the unsolved Sudoku puzzle in the designated area
3. Click the "Populate Solver" button to initialize the solving process
4. Press "Solve Sudoku" to find the solution
5. Use "Clear Solver" to reset for a new puzzle

## Technical Details

The solver employs linear programming techniques to efficiently solve Sudoku puzzles. It utilizes Excel's built-in features along with custom macros and scripts for optimal performance. For more technical explanation: visit [py_sudoku_solver](https://github.com/choppystick/py-sudoku-solver)

### Solver Logic
![image](https://github.com/user-attachments/assets/02c03c8f-9e24-48f2-9675-8366dfa43533)![image](https://github.com/user-attachments/assets/a05f0f77-74e8-476d-8fe5-6561c252d0c9)

- The solver uses a series of 9x9 grids to represent the Sudoku puzzle
- Each cell is represented by binary variables (0 or 1) for each possible digit (1-9)
- Constraints are applied to ensure Sudoku rules are followed (row, column, and 3x3 box uniqueness)
- The linear programming model optimizes these constraints to find a valid solution

## Requirements

- Microsoft Excel (version X or higher)
- OpenSolver (An open-source Excel LP solver)
- Enabled macros for full functionality

## Installation

1. Download the Excel file
2. Enable macros when prompted
3. Install OpenSolver 2.93x and enable it in addon
4. You're ready to start solving Sudoku puzzles!

## Contributing

Contributions to improve the solver are welcome. Please feel free to submit pull requests or open issues for any bugs or enhancements.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE.txt) file for details.
