Attribute VB_Name = "Main"
Option Explicit
Option Base 0



Sub Main()

    Dim Arr As ArrayListCol
    Set Arr = New ArrayListCol



    Dim col As iCollection
    Set col = New ArrayListCol

    col.add "blah"
    col.add "blah2"


    Arr.addAll col
    Debug.Print (Arr.size = 2) & " ## addAll "

    col.remove 1
    col.remove 1
    col.add "a"
    col.add "b"

    Arr.addAllAtIndex 0, col
    Arr.addAllAtIndex 1, col
    Arr.addAllAtIndex 2, col
    Debug.Print ("[ blah, blah, blah, blah2, a, b, blah2, a, b, blah2, a, b, blah, blah2]" = Arrays.toString(Arr.ToArray)) & " ## addAllAtIndex "

    Arr.clear
    Arr.add "something"
    Debug.Print (Arr.size = 1) & " ## add"
    Debug.Print (Arr.size = 1) & " ## size"

    
    Arr.addAtIndex 0, "fun thing"
    Arr.addAtIndex 2, "cool thing"
    Debug.Print (Arr.getIndex(2) = "cool thing" And Arr.getIndex(0) = "fun thing") & " ## addAtIndex "

    Arr.clear
    Debug.Print (Arr.size = 0) & " ## clear "

    Arr.add "foo"
    Debug.Print (Arr.contains("foo") And Not Arr.contains("bar")) & " ## contains"

    Arr.add "bar"
    Debug.Print (Arr.getIndex(0) = "foo" And Arr.getIndex(1) = "bar") & " ## getIndex"

    Arr.add "boze"
    Arr.add "boze"
    Debug.Print (Arr.indexOf("foo") = 0 And Arr.indexOf("bar") = 1 And Arr.indexOf("baz") = -1 And Arr.indexOf("boze") = 2) & " ## indexOf "
    Debug.Print (Arr.lastIndexOf("foo") = 0 And Arr.lastIndexOf("bar") = 1 And Arr.lastIndexOf("baz") = -1 And Arr.lastIndexOf("boze") = 3) & " ## lastIndexOf "

    Debug.Print (Arr.remove("foo") = True And Arr.remove("foo") = False) & " ## Remove"

    Dim isEmpty As Boolean
    isEmpty = Arr.isEmpty
    Arr.clear
    Debug.Print (isEmpty = False And Arr.isEmpty = True) & " ## isEmpty "

    Arr.add "foo"
    Arr.addAll col
    Arr.add "bar"
    Arr.addAll col
    Arr.add "baz"


    Dim removeAllContains1 As Boolean
    Dim removeAllContains2 As Boolean
'    Arr.removeAll col
'    removeAllContains1 = (Arr.contains("a") And Arr.contains("b"))
'    Arr.removeAll col
'    removeAllContains2 = (Arr.contains("a") And Arr.contains("b"))
'
'    Debug.Print (removeAllContains1 And Not removeAllContains2 And Not Arr.isEmpty) & " ## removeAll "


End Sub



