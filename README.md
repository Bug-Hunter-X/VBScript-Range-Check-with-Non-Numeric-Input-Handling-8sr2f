# VBScript Range Check with Non-Numeric Input Handling

This repository contains a VBScript function that checks if a number is within a specified range.  The original function lacked robust error handling for non-numeric inputs.  This improved version addresses this issue.

## Bug

The `IsWithinRange` function in `bug.vbs` does not correctly handle non-numeric inputs. Passing a string or other non-numeric value causes a type mismatch error.

## Solution

The solution provided in `bugSolution.vbs` adds error handling using `IsNumeric` to check if the input is a number before performing the range check.  If the input is not numeric, it returns `False` without causing an error.

## Usage

1.  Save both `bug.vbs` and `bugSolution.vbs`.
2.  Run either file using a VBScript interpreter (e.g., by double-clicking the file in Windows).

The improved function provides more reliable range checking for VBScript applications.