# VBScript ByRef Parameter Passing Bug

This repository demonstrates a subtle but important behavior of `ByRef` parameter passing in VBScript when dealing with immutable data types.  The following code snippet illustrates that while the function appears to modify the original variable, the modification doesn't persist outside the function's scope.