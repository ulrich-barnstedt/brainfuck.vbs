# Brainfuck.vbs

A brainfuck interpreter, in 175 lines of Visual Basic Script.  

### Why?

I've decided to implement brainfuck interpreters in diverse languages, such as JS and [LOLCODE](https://github.com/ulrich-barnstedt/brainfuck.lol)
, and this is the VBS version.

### Usage

```shell
cscript bf.vbs <source file>

# For example
cscript bf.vbs test.bf # Run the included "Hello World!"
```

### Modifying

See Lines 6 - 10 for config variables such as buffer size and debug logging.