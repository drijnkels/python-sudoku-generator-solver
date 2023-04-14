# Python Sudoku Generator and Sudoku Solver

## about

Python based Sudoku generator that can create unique Sudoku board based on 4 difficulty levels. This code also includes a brute force Sudoku solver that is capable of solving even the most difficult Sudoku puzzles!

## Sudoku Generator Usage

### Install environment

For a Windows system with git bash use the following commands to set up an virtual python environment.

``` sh
python -m venv venv
```

### Activate environment

``` sh
source venv/Scripts/activate
```

### Install requirements

``` sh
pip install -r requirements.txt
```

### Adjust settings

#### Difficulty

To adjust the difficulty level of the generated Sudoku puzzle, browse to sudoku.py and change the variable 'level' to the desired difficulty levels, which include `Easy`, `Medium`, `Hard` and `Insane`.

#### Size

To adjust the size of the generated Sudoku set the variable 'size' to an value from 2 or higher.
The standard Sudoku size value is `3`.
This means a $3^2 = 9$ is the highest possible number.

| size | number range |
| :--: | :----------: |
|  2   |   1 ... 4    |
|  3   |   1 ... 9    |
|  4   |   1 ... 16   |
|  5   |   1 ... 25   |
| ...  |     ...      |

#### Amount

To adjust the amount of created Sudokus set the variable 'amount' to desired quantity.

#### Output

To adjust the way of output, set the variables 'print_console' and 'export_excel' to `True` or `False`.

### Run script

``` sh
python sudoku.py
```

**B.Y.O.T (bring your own tests)**

See more at www.callmejoe.net
Made by Joe Carlson 2015
