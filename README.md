# VBA-utilities
VBA-utilities is a collection of helpful modules for programming in VBA. It is important to note, these modules are all programed using the `Option Base 0` and `Option Compare Text` settings.

## Modules
This section lists each of the modules along with a brief description of their purpose.

### Arrays
The `Arrays` module contains various methods for manipulating arrays in VBA, e.g. copying and sorting. For example, an array `A()` can be sorted using the following method call

    Arrays.MergeSort A()


### Collections
The `Collections` module contains various methods for manipulating collections in VBA. Specifically, the module provides methods to check  if a collection contains a specific element and also to retrieve its index, and to sort a collection (requires the `Arrays` module be loaded). For a collection `col`, the following is a list of example method calls:

    Dim bVal1 As Boolean
    bVal1 = Collections.contains("hello world", col)
    
    Dim bVal2 As Boolean
    bVal = Collections.hasKey(5, col)
    
    Dim iVal As Long
    iVal = Collections.indexOf("hello world", col)
    
    Collections.mergeSort col
