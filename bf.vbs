arg = WScript.Arguments.Item(0)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(arg, 1)

' Basic variables
Dim BUF(29999)          ' Main memory
Dim BUF_SZ, PTR, DEBUG
BUF_SZ = 30000          ' Size of the memory
PTR = 0                 ' Current cell of the pointer
DEBUG = False           ' If debug logging is enabled

' AST accumulator
ReDim statements(-1)

function appendToStatements (stm, isObjectP)
    ReDim preserve statements(UBound(statements) + 1)

    if isObjectP then
        set statements(UBound(statements)) = stm
    else
        statements(UBound(statements)) = stm
    end if
end function


' Standard methods for BF
function validatePointer()
    if BUF(PTR) = Empty then
        BUF(PTR) = 0
    end if
end function

function changePtr (diff)
    validatePointer

    if PTR + diff = BUF_SZ or PTR + diff < 0 then
        exit function
    end if

    if DEBUG then WScript.echo "Px" & ptr + diff & " " & diff end if
    ptr = ptr + diff
end function

function increment ()
    validatePointer

    if BUF(PTR) = 255 then
        BUF(PTR) = 0
    else
        BUF(PTR) = BUF(PTR) + 1
    end if

    if DEBUG then WScript.echo "Ix" & PTR & " " & BUF(PTR) end if
end function

function decrement ()
    validatePointer

    if BUF(PTR) = 0 then
        BUF(PTR) = 255
    else
        BUF(PTR) = BUF(PTR) - 1
    end if

    if DEBUG then WScript.echo "Dx" & PTR & " " & BUF(PTR) end if
end function

function log ()
    validatePointer

    if DEBUG then
        WScript.echo ">>O " & BUF(PTR)
        exit function
    end if

    WScript.StdOut.Write(CHR(BUF(PTR)))
end function

' Run a statement
function exec (statement)
    if TypeName(statement) = "Collection" then
        if DEBUG then WScript.echo ">>L" end if
        statement.run
        if DEBUG then WScript.echo "<<L" end if

        exit function
    end if

    select case statement
        case ">" changePtr 1
        case "<" changePtr -1
        case "+" increment
        case "-" decrement
        case "." log
        case "," BUF(PTR) = asc(InputBox("Input request"))
    end select
end function

' Pre-process - parse loops
class Collection
    private collected

    private sub Class_Initialize
        reDim coll_array(-1)
        collected = coll_array
    end sub

    public sub append (statement)
        reDim preserve collected(UBound(collected) + 1)
        if TypeName(statement) = "Collection" then
            set collected(UBound(collected)) = statement
        else
            collected(UBound(collected)) = statement
        end if
    end sub

    public sub run ()
        if BUF(PTR) = 0 then
            exit sub
        end if

        while not BUF(PTR) = 0
            for k=0 to UBound(collected)
                exec(collected(k))
            next
        wend
    end sub
end class

dim depth
reDim accHelper(-1)
depth = -1

function preProcess (char)
    select case char
        case "["
            depth = depth + 1

            if UBound(accHelper) < depth then
                reDim preserve accHelper(UBound(accHelper) + 1)
            end if

            set accHelper(depth) = new Collection
        case "]"
            depth = depth - 1

            if depth > -1 then
                accHelper(depth).append(accHelper(depth + 1))
            else
                appendToStatements accHelper(depth + 1), True
            end if
        case else
            ' WScript.echo UBound(accHelper)

            if depth > -1 then
                accHelper(depth).append(char)
            else
                appendToStatements char, False
            end if
    end select
end function


' Exec the program
do while objFile.AtEndOfStream = False
	strLine = objFile.ReadLine

	for i=1 To Len(strLine)
	    preProcess Mid(strLine, i, 1)
	next

    for j=0 to UBound(statements)
        exec(statements(j))
    next
Loop