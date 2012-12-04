#! /usr/bin/python

import argparse
import os
import csv
import sys

##### UTiLS #####

class KTSpreadsheet:
    
    spreadsheet = None
    
    def __init__(self, fullpath):
        if os.path.isdir(fullpath):
            print "Please pass in the path to a CSV file and not a directory."
        elif os.stat(fullpath).st_size == 0:
            print os.stat(fullpath)
            print "The file either does not exist or is empty."
        else:
            # then file exists/isn't empty and we can continue with the parsing
            f = open(fullpath, 'r')
            reader = csv.reader(f)
    
            self.spreadsheet = []
            for i, row in enumerate(reader):
                spreadsheet_row = []
                for cell in row:
                    spreadsheet_row.append(cell.strip())
                self.spreadsheet.append(spreadsheet_row)
            f.close()

    # small utility to convert alphabetic letters to numbers 1-26 (for the column)
    def letter_to_num(self, letter):
        return ord(str.lower(letter)) - ord('a')
    
    # function to read a certain row number from csv file
    def read_row(self, num, file_path):
        f = open(file_path, 'r')
        reader = csv.reader(f)
    
        # enumerate(reader) utilizes .next() to avoid loading ALL the file into memory
        for i, line in enumerate(reader):
            if  reader.line_num == num:
                f.close()
                return line
            if reader.line_num > num:
                break
        f.close()
        return nil
    
    
    ##### LOGiC #####
    
    ## Handle number operations
    # assuming only one operation
    def handle_num(self, cell_args):
        eval_str = ""
        for num_str in cell_args[0:-1]:
            # must convert integer num_str's to floats
            num_str = str(float(num_str))
    
            if eval_str == "":
                # first element
                eval_str += num_str
            else:
                eval_str += cell_args[-1] + num_str
        ret_str = eval(eval_str)
    
        # specific formatting to return integer whenever possible
        # extremely hacky solution to check if float ends in ___.0
        # if so, then convert to int
        if str(ret_str)[-2:] =='.0':
            ret_str = int(ret_str)
        return str(ret_str)
    
    ## Handle cell reference
    # assuming there is only ONE cell reference
    # ** In order to handle circular referencing, breaks a loop and returns -1 after seeing
    #      the initial reference again.
    # ^ Did not need to implement... Just thought about it
    def handle_ref(self, cell_args, fullpath):
        ref = cell_args[0]
        col_num = self.letter_to_num(ref[0])
        row_num = int(ref[1:])
    
        ref_row = self.read_row(row_num, fullpath)
        return str.strip(ref_row[col_num])
    
    
    ## Function that handles evaluation of cells
    # 
    # Input: (Note that all input is originally a string)
    #   int -> strips whitespace and returns integer
    #   eval -> begins with "=". Evaluates value and returns integer
    #   default -> Any other input should not be handled
    #
    # ** needs fullpath because we call read_row in handle_ref
    def eval_cell(self, cell, fullpath):
        stripped_cell = str.strip(cell)
        valid_ops = ['+', '-', '*', '/']
        op = None
    
        if stripped_cell == "":
            return ""
    
        try:
            if float(stripped_cell):
                # this cell is a number
                return stripped_cell
        except ValueError:
            # Not a float/integer
            if stripped_cell[0] == "=":
                # need to evaluate cell
                # remove '=' from cell
                cell_args = stripped_cell[1:].split()
                if stripped_cell[-1] in valid_ops:
                    # there is an operator and will assume that this is of the form =## ... ## op
                    return self.handle_num(cell_args)
                return self.handle_ref(cell_args, fullpath)
            else:
                print "Improper input. Only handles the floats/ints >= 0 and '=_____' notation"
    
    # Main logic function for spreadsheet evaluator
    def eval_spreadsheet(self):
        for row_index, row in enumerate(self.spreadsheet):
            for cell_index, cell in enumerate(row):
                self.spreadsheet[row_index][cell_index] = self.eval_cell(cell, fullpath)
            
    def print_spreadsheet(self):
        for row in self.spreadsheet:
            sys.stdout.write("\t".join(row) + '\n')


##### MAiN LOOP #####

# code that provides "pretty" command line help and argument checking
parser = argparse.ArgumentParser(description='Simple evaluator for a CSV spreadsheet')
parser.add_argument('csv_file', metavar='CSV_FILE', type=str,
                                help='Input CSV spreadsheet filepath')
args = parser.parse_args()


# find full path rather than relative path for os module's use
fullpath =  os.path.realpath(args.csv_file)
kt_spreadsheet = KTSpreadsheet(fullpath)
kt_spreadsheet.eval_spreadsheet()
kt_spreadsheet.print_spreadsheet()
