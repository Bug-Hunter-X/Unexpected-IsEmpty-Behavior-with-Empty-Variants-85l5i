# Unexpected IsEmpty Behavior in VBScript

This repository demonstrates a subtle bug in VBScript's `IsEmpty` function when dealing with empty variants.  The function doesn't always reliably detect truly empty variants, potentially causing issues in conditional logic.

## Bug Description
The `IsEmpty` function in VBScript is designed to check if a variant is empty. However, there are cases, particularly with variants passed as arguments that might seem empty, where `IsEmpty` returns `False` even when the value should be considered empty. This can lead to unexpected code execution paths.

## Reproduction
The `bug.vbs` file contains a VBScript function that highlights this issue.  Run the script to observe the behavior.

## Solution
The `bugSolution.vbs` file provides a corrected version of the function, demonstrating a robust way to handle potentially empty variants. This involves explicitly checking for both `IsEmpty` and an empty string value, ensuring a more reliable check.

## How to use
1. Save both `bug.vbs` and `bugSolution.vbs` files.
2. Run each file from a VBScript environment (e.g. double-click on the file). 
3. Compare the outputs to see the difference in handling empty variants.