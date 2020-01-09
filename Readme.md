# excel-formula-generator

This is my submission for the take-home assessment of generating excel formula in textual format.

## Logic

As this problem requires manipulation of string and pattern matching, I went with the 'Regular Expression' approach. First, I gazed through the given Excel and found the different formulas used to generate the cells and started finding the patterns. I came up with two Regular Expression that solved the problem for me like a charm. The first one is to split the formula into individual parts and the second one to parse each part and generate the corresponding textual representation of it. The algorithm is described in [main.py](https://github.com/Girish21/excel-formula-generator/blob/master/main.py#L101) 'build_excel' function.

## Scripts

1. [main.py](https://github.com/Girish21/excel-formula-generator/blob/master/main.py)
    - Driver script which parses the given excel and generates the target excel with formulas in textual format.
2. [test_base.py](https://github.com/Girish21/excel-formula-generator/blob/master/test_base.py)
    - Test script which tests the helper functions
3. [test_main.py](https://github.com/Girish21/excel-formula-generator/blob/master/test_main.py)
    - Test script to test the main logic of the solution

## Technology Used

> 1. Python (3.7)
> 2. openpyxl (3)
