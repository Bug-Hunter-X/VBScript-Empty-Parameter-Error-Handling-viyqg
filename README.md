# VBScript Empty Parameter Error Handling
This repository demonstrates a common error in VBScript related to handling empty parameters.  The `bug.vbs` file shows the flawed error handling.  The `bugSolution.vbs` demonstrates a more robust approach.

**The Problem:**
In VBScript, passing an empty parameter to a function can easily lead to unexpected runtime errors. The original code uses `Err.Raise`, which is good for raising errors, but lacks context or proper error trapping.

**The Solution:**
The solution introduces more sophisticated error handling using `On Error Resume Next` and `Err.Number` to properly handle and report the error, improving the robustness and reliability of the code.