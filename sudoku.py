##Working code to Generate a puzzle
# -*- coding: utf-8 -*-
import copy
import math
import os
import random
import time
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter

level = "Medium"
size = 3
amount = 2
print_console = False
export_excel = True

# [Level of Difficulty] = Input the level of difficulty of the Sudoku puzzle. Difficulty levels
#        include ‘Easy’ ‘Medium’ ‘Hard’ and ‘Insane’. Outputs a Sudoku of desired
#        difficulty.

# 2 -> 2 * 2 = 4 -> numbers 1 ... 4
# 3 -> 3 * 3 = 9 -> numbers 1 ... 9
# 4 -> 4 * 4 = 16 -> numbers 1 ... 16
# 5 -> 5 * 5 = 25 -> numbers 1 ... 25
# ...

# 2 -> 18 per page
# 3 -> 2 per page
# 4 -> 1 per page
# ...

class cell():
    #   Initializes cell object. A cell is a single box of a Sudoku puzzle. 81 cells make up the body of a
    #   3*3 Sudoku puzzle. Initializes puzzle with all possible answers available, solved to false, and position of cell within the
    #   Sudoku puzzle
    def __init__(self, position, size):
        self.size = size
        self.possibleAnswers = list(range(1,size**2+1))
        self.answer = None
        self.position = position
        self.solved = False
        
    def remove(self, num):
        """Removes num from list of possible answers in cell object."""
        if num in self.possibleAnswers and self.solved == False:
            self.possibleAnswers.remove(num)
            if len(self.possibleAnswers) == 1:
                self.answer = self.possibleAnswers[0]
                self.solved = True
        if num in self.possibleAnswers and self.solved == True:
            self.answer = 0

    def solvedMethod(self):
        """ Returns whether or not a cell has been solved"""
        return self.solved

    def checkPosition(self):
        """ Returns the position of a cell within a Sudoku puzzle. x = row; y = col; z = box number"""
        return self.position

    def returnPossible(self):
        """ Returns a list of possible answers that a cell can still use"""
        return self.possibleAnswers

    def lenOfPossible(self):
        """ Returns an integer of the length of the possible answers list"""
        return len(self.possibleAnswers)

    def returnSolved(self):
        """ Returns whether or not a cell has been solved"""
        if self.solved == True:
            return self.possibleAnswers[0]
        else:
            return 0
        
    def returnSize(self):
        # returns the size of the Sudoku
        return self.size
        
    def setAnswer(self, num):
        """ Sets an answer of a puzzle and sets a cell's solved method to true. This
            method also eliminates all other possible numbers"""
        if num in list(range(1,self.size**2+1)):
            self.solved = True
            self.answer = num
            self.possibleAnswers = [num]
        else:
            raise(ValueError)
       
    def reset(self):
        """ Resets all attributes of a cell to the original conditions""" 
        self.possibleAnswers = list(range(1,self.size**2+1))
        self.answer = None
        self.solved = False

def emptySudoku(size):
    # Creates an empty Sudoku in row major form. Sets up all of the x, y, and z coordinates for the Sudoku cells
    ans = []
    for x in list(range(1,size**2+1)):
        intz = math.floor((x-1)/size)*size+1
        for y in list(range(1,size**2+1)):
            z = intz
            z += math.floor((y-1)/size)
            c = cell((x,y,z),size)
            ans.append(c)
    return ans

def printSudoku(sudoku):
    size = sudoku[0].returnSize()
    width = len(str(size**2))
    columns = size**2

    for row in range(columns*2+1):
        print("")
        for column in range(columns*4+1):
            if row == 0:
                if column == 0:
                    print("╔", end = '')
                elif column % 4 == 0:
                    if column == columns*4:
                        print("╗", end = '')
                    elif column % (4*size) == 0:
                        print("╦", end = '')
                    else:
                        print("╤", end = '')
                elif (column+2) % 4 == 0:
                    print("═"*width, end = '')
                elif column % 4 > 0:
                    print("═", end = '')
            elif row % 2 > 0:
                if column % (4*size) == 0:
                    print("║", end = '')
                elif column % 4 == 0:
                    print("│", end = '')
                elif (column+2) % 4 == 0:
                    i = int((row-1)/2*columns + (column-1)/4)
                    value = sudoku[i].returnSolved()
                    if value == 0:
                        value = ' '
                    print(f"{str(value).rjust(width)}", end = '')
                else:
                    print(' ', end = '')
            elif row % (columns*2) == 0:
                if column == 0:
                    print("╚", end = '')
                elif column % 4 == 0:
                    if column == columns*4:
                        print("╝", end = '')
                    elif column % (4*size) == 0:
                        print("╩", end = '')
                    else:
                        print("╧", end = '')
                elif (column+2) % 4 == 0:
                    print("═"*width, end = '')
                elif column % 4 > 0:
                    print("═", end = '')
            elif row % (size*2) == 0:
                if column == 0:
                    print("╠", end = '')
                elif column % 4 == 0:
                    if column == columns*4:
                        print("╣", end = '')
                    elif column % (4*size) == 0:
                        print("╬", end = '')
                    else:
                        print("╪", end = '')
                elif (column+2) % 4 == 0:
                    print("═"*width, end = '')
                elif column % 4 > 0:
                    print("═", end = '')
            else:
                if column == 0:
                    print("╟", end = '')
                elif column % 4 == 0:
                    if column == columns*4:
                        print("╢", end = '')
                    elif column % (4*size) == 0:
                        print("╫", end = '')
                    else:
                        print("┼", end = '')
                elif (column+2) % 4 == 0:
                    print("─"*width, end = '')
                elif column % 4 > 0:
                    print("─", end = '')
    print("")

def exportSudoku(sudoku, count = 0, solution=False):
    size = sudoku[0].returnSize()
    columns = size**2

    if solution:
        filename = f'sudoku_{columns}x{columns}_{level}_Solution.xlsx'
    else:
        filename = f'sudoku_{columns}x{columns}_{level}.xlsx'

    MAX_COLUMN_PER_PAGE = 16
    MAX_ROW_PER_PAGE = 28

    max_sudoku_per_row = max(1,math.floor(MAX_COLUMN_PER_PAGE/(columns+1)))

    position_column = count%max_sudoku_per_row
    position_row = math.floor(count/max_sudoku_per_row)
    
    offset_column = position_column * (columns + 1)
    offset_row = position_row * (columns + 1)

    if os.path.isfile(filename):
        wb = load_workbook(filename)
    else:
        wb = Workbook()
    ws = wb.active

    thin = Side(border_style="thin", color="00808080")
    thick = Side(border_style="thick", color="00808080")

    border = Border(left=thin,
                    right=thin,
                    top=thin,
                    bottom=thin,
                )
    borderlefttop = Border(left=thick,
                    right=thin,
                    top=thick,
                    bottom=thin,
                )
    borderleft = Border(left=thick,
                    right=thin,
                    top=thin,
                    bottom=thin,
                )
    borderleftbottom = Border(left=thick,
                    right=thin,
                    top=thin,
                    bottom=thick,
                )
    borderbottom = Border(left=thin,
                    right=thin,
                    top=thin,
                    bottom=thick,
                )
    borderrightbottom = Border(left=thin,
                    right=thick,
                    top=thin,
                    bottom=thick,
                )
    borderright = Border(left=thin,
                    right=thick,
                    top=thin,
                    bottom=thin,
                )
    borderrighttop = Border(left=thin,
                    right=thick,
                    top=thick,
                    bottom=thin,
                )
    bordertop = Border(left=thin,
                    right=thin,
                    top=thick,
                    bottom=thin,
                )
    
    alignment=Alignment(horizontal='center',
                     vertical='center',
                     text_rotation=0,
                     wrap_text=False,
                     shrink_to_fit=False,
                     indent=0)

    for row in range(columns):
        ws.row_dimensions[row+1+offset_row].height = 25
        for column in range(columns):
            ws.column_dimensions[get_column_letter(column+1+offset_column)].width = 5
            value = sudoku[row*columns+column].returnSolved()
            if value == 0:
                value = ''
            cell = ws.cell(row=row+1+offset_row, column=column+1+offset_column, value=value)
            if row%size == 0 and column%size == 0:
                cell.border = borderlefttop
            elif row%size == 0 and (column+1)%size == 0:
                cell.border = borderrighttop
            elif row%size == 0:
                cell.border = bordertop
            elif (row+1)%size == 0 and column%size == 0:
                cell.border = borderleftbottom
            elif (row+1)%size == 0 and (column+1)%size == 0:
                cell.border = borderrightbottom
            elif (row+1)%size == 0:
                cell.border = borderbottom
            elif column%size == 0:
                cell.border = borderleft
            elif (column+1)%size == 0:
                cell.border = borderright
            else:
                cell.border = border
            cell.alignment = alignment

    wb.save(filename)

def sudokuGen(size):
    # Generates a completed Sudoku. Sudoku is completely random
    cells = [i for i in range(size**4)] ## our cells is the positions of cells not currently set
    sudoku = emptySudoku(size)
    while len(cells) != 0:
        lowestNum = []
        Lowest = []
        for i in cells:
            lowestNum.append(sudoku[i].lenOfPossible())  ## finds all the lengths of of possible answers for each remaining cell
        m = min(lowestNum)  ## finds the minimum of those
        '''Puts all of the cells with the lowest number of possible answers in a list titled Lowest'''
        for i in cells:
            if sudoku[i].lenOfPossible() == m:
                Lowest.append(sudoku[i])
        '''Now we randomly choose a possible answer and set it to the cell'''
        choiceElement = random.choice(Lowest)
        choiceIndex = sudoku.index(choiceElement) 
        cells.remove(choiceIndex)                 
        position1 = choiceElement.checkPosition()
        if choiceElement.solvedMethod() == False:  ##the actual setting of the cell
            possibleValues = choiceElement.returnPossible()
            finalValue = random.choice(possibleValues)
            choiceElement.setAnswer(finalValue)
            for i in cells:  ##now we iterate through the remaining unset cells and remove the input if it's in the same row, col, or box
                position2 = sudoku[i].checkPosition()
                if position1[0] == position2[0]:
                    sudoku[i].remove(finalValue)
                if position1[1] == position2[1]:
                    sudoku[i].remove(finalValue)
                if position1[2] == position2[2]:
                    sudoku[i].remove(finalValue)

        else:
            finalValue = choiceElement.returnSolved()
            for i in cells:  ##now we iterate through the remaining unset cells and remove the input if it's in the same row, col, or box
                position2 = sudoku[i].checkPosition()
                if position1[0] == position2[0]:
                    sudoku[i].remove(finalValue)
                if position1[1] == position2[1]:
                    sudoku[i].remove(finalValue)
                if position1[2] == position2[2]:
                    sudoku[i].remove(finalValue)
    return sudoku

def sudokuChecker(sudoku):
    """ Checks to see if an input a completed sudoku puzzle is of the correct format and abides by all
        of the rules of a sudoku puzzle. Returns True if the puzzle is correct. False if otherwise"""
    for i in range(len(sudoku)):
        for n in range(len(sudoku)):
            if i != n:
                position1 = sudoku[i].checkPosition()
                position2 = sudoku[n].checkPosition()
                if position1[0] == position2[0] or position1[1] == position2[1] or position1[2] == position2[2]:
                    num1 = sudoku[i].returnSolved()
                    num2 = sudoku[n].returnSolved()
                    if num1 == num2:
                        return False
    return True

def perfectSudoku(size):
    '''Generates a completed sudoku. Sudoku is in the correct format and is completely random'''
    result = False
    while result == False:
        s = sudokuGen(size)
        result = sudokuChecker(s)
    return s

def solver(sudoku, size, f = 0):
    """ Input an incomplete Sudoku puzzle and solver method will return the solution to the puzzle. First checks to see if any obvious answers can be set
        then checks the rows columns and boxes for obvious solutions. Lastly the solver 'guesses' a random possible answer from a random cell and checks to see if that is a
        possible answer. If the 'guessed' answer is incorrect, then it removes the guess and tries a different answer in a different cell and checks for a solution. It does this until
        all of the cells have been solved. Returns a printed solution to the puzzle and the number of guesses that it took to complete the puzzle. The number of guesses is
        a measure of the difficulty of the puzzle. The more guesses that it takes to solve a given puzzle the more challenging it is to solve the puzzle"""
    if f > 900:
        return False
    guesses = 0
    copy_s = copy.deepcopy(sudoku)
    cells = [i for i in range(size**4)] ## our cells is the positions of cells not currently set
    solvedCells = []
    for i in cells:
        if copy_s[i].lenOfPossible() == 1:
            solvedCells.append(i)
    while solvedCells != []:
        for n in solvedCells:
            cell = copy_s[n]
            position1 = cell.checkPosition()
            finalValue = copy_s[n].returnSolved()
            for i in cells:  ##now we iterate through the remaining unset cells and remove the input if it's in the same row, col, or box
                position2 = copy_s[i].checkPosition()
                if position1[0] == position2[0]:
                    copy_s[i].remove(finalValue)
                if position1[1] == position2[1]:
                    copy_s[i].remove(finalValue)
                if position1[2] == position2[2]:
                    copy_s[i].remove(finalValue)
                if copy_s[i].lenOfPossible() == 1 and i not in solvedCells and i in cells:
                    solvedCells.append(i)
                ##print(n)
            solvedCells.remove(n)
            cells.remove(n)
        if cells != [] and solvedCells == []:
            lowestNum=[]
            lowest = []
            for i in cells:
                lowestNum.append(copy_s[i].lenOfPossible())
            m = min(lowestNum)
            for i in cells:
                if copy_s[i].lenOfPossible() == m:
                    lowest.append(copy_s[i])
            randomChoice = random.choice(lowest)
            randCell = copy_s.index(randomChoice)
            randGuess = random.choice(copy_s[randCell].returnPossible())
            copy_s[randCell].setAnswer(randGuess)
            solvedCells.append(randCell)
            guesses += 1
    if sudokuChecker(copy_s):
        if guesses == 0:
            level = 'Easy'
        elif guesses <= 2:
            level = 'Medium'
        elif guesses <= 7:
            level = 'Hard'
        else:
            level = 'Insane'
        return copy_s, guesses, level
    else:
        return solver(sudoku, size, f+1)
    
def solve(sudoku, size, n = 0):
    """ Uses the solver method to solve a puzzle. This method was built in order to avoid recursion depth errors. Returns True if the puzzle is solvable and
        false if otherwise"""
    if n < 30:
        s = solver(sudoku, size)
        if s != False:
            return s
        else:
            solve(sudoku, size, n+1)
    else:
        return False
    
def puzzleGen(sudoku, size):
    """ Generates a puzzle with a unique solution. """
    cells = [i for i in range(size**4)]
    while cells != []:
        copy_s = copy.deepcopy(sudoku)
        randIndex = random.choice(cells)
        cells.remove(randIndex)
        copy_s[randIndex].reset()
        s = solve(copy_s, size)
        if s[0] == False:
            f = solve(sudoku, size)
            print("Guesses: " + str(f[1]))
            print("Level: " + str(f[2]))
            return printSudoku(sudoku)
        elif equalChecker(s[0],solve(copy_s, size)[0]):
            if equalChecker(s[0],solve(copy_s, size)[0]):
                sudoku[randIndex].reset()
        else:
            f = solve(sudoku, size)
##            print("Guesses: " + str(f[1]))
##            print("Level: " + str(f[2]))
            return sudoku, f[1], f[2]

def equalChecker(s1,s2):
    """ Checks to see if two puzzles are the same"""
    for i in range(len(s1)):
        if s1[i].returnSolved() != s2[i].returnSolved():
            return False
    return True

def main(level, size, count):
    # Input the level of difficulty of the sudoku puzzle. Difficulty levels
    #    include ‘Easy’ ‘Medium’ ‘Hard’ and ‘Insane’. Outputs a sudoku of desired
    #    difficulty.

    t1 = time.time()
    n = 0

    print("------------------------------")
    print(f"Sudoku number: {count}")

    p = perfectSudoku(size)
    if export_excel:
        exportSudoku(p, count, solution=True)
    s = puzzleGen(p, size)
    
    if level == 'Easy':
        if s[2] != 'Easy':
            return main(level, size, count)
    if level == 'Medium':
        while s[2] == 'Easy':
            n += 1
            s = puzzleGen(p, size)
            if n > 50:
                return main(level, size, count)
        if s[2] != 'Medium':
            return main(level, size, count)
    if level == 'Hard':
        while s[2] == 'Easy':
            n += 1
            s = puzzleGen(p, size)
            if n > 50:
                return main(level, size, count)
        while s[2] == 'Medium':
            n += 1
            s = puzzleGen(p, size)
            if n > 50:
                return main(level, size, count)
        if s[2] != 'Hard':
            return main(level, size, count)
    if level == 'Insane':
        while s[2] != 'Insane':
            n += 1
            s = puzzleGen(p, size)
            if n > 50:
                return main(level, size, count)
    
    t2 = time.time()
    t3 = t2 - t1
    print("Runtime is " + str(t3) + " seconds")
    print("Guesses: " + str(s[1]))
    print("Level: " + str(s[2]))
    if export_excel:
        exportSudoku(s[0], count)
    if print_console:
        printSudoku(s[0])
    return

for i in range(amount):
    main(level, size, i)