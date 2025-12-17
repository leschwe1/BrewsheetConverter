## Chapter 2: Objects and Classes 
> → all
>- Objects are useful without classes, but classes make them easier to understand.
> - A well-designed class defines a contract that code using its instances can rely on.
> - Objects that respect the same contract are polymorphic, i.e., they can be used interchangeably even if they do different specific things.
>- Objects and classes can be thought of as dictionaries with stereotyped behavior.
>- Most languages allow functions and methods to take a variable number of arguments.
>- Inheritance can be implemented in several ways that differ in the order in which objects and classes are searched for methods.


*Objects are useful without classes, but classes make them easier to understand.*


Reasons for OOP:
- natural way to represent real-world “things” in code
- organize code to make it easier to understand, test, and extend


### Contract: 
An object must satisfy it in order to be considered an instance of a subclass, i.e., must provide methods with these names


```py3
class Shape:
    def perimeter(self):
        raise NotImplementedError("perimeter")

class Square(Shape):
    def perimeter(self):
        return 4 * self.side
```

### Polymorphism:
**Example:** Superclass has the method perimeter not implemented. The two subclasses Square and shape implement it with their own functions. This means two different functions under the same name 

- Objects that respect the same contract are polymorphic
- reduces cognitive load 
- allows people using related things to ignore their differences
- can be used interchangeably even if they do different specific things

### Classes without classes
- store function objects in lists and dicts 
- tricky: all squares should have different values but use the same object --> methods stored in dict that corresponds to class `_class`, each square contains that higher level dict
```py3
######CLASS Methods (for class Square and superclass Shape)
def square_perimeter(thing):
    return 4 * thing["side"]

def square_area(thing):
    return thing["side"] ** 2

def square_larger(thing, size): #uses extra arg!
    return call(thing, "area") > size

def shape_density(thing, weight):
    return weight / call(thing, "area")


######CLASS Dicts (Square inherits from Shape)
Shape = {
    "density": shape_density,
    "_classname": "Shape",
    "_parent": None
    "_new": shape_new
}

Square = {
    "perimeter": square_perimeter,
    "area": square_area,
    "larger": square_larger,
    "_classname": "Square"
    "_parent": Shape
    "_new": square_new
}

######INIT/MAKE Method
def make(cls, *args):
    return cls["_new"](*args)

#-------------
def shape_new(name):
    return {
        "name": name,
        "_class": Shape
    }

def square_new(name, side):
    return make(Shape, name) | { #using | to combine two dictionaries
        "side": side,
        "_class": Square
    }


######CALL Method
def call(thing, method_name, *args):
    method = find(thing["_class"], method_name)
    return method(thing, *args)

    #helper to search in parent classes
def find(cls, method_name):
    while cls is not None:
        if method_name in cls:
            return cls[method_name]
        cls = cls["_parent"]
    raise NotImplementedError("method_name")

######SAMPLE CALL
examples = [make(Square, "sq", 3), make(Circle, "ci", 2)]
for ex in examples:
    n = ex["name"]
    d = call(ex, "density", 5)
    print(f"{n}: {d:.2f}")
```

### *args and **kwargs (= 'varargs')
> spreading: To automatically match the values from a list or dictionary supplied by the caller to the parameters of a function.
- with a leading *, it captures any “extra” values passed to the function that don’t line up with named parameters
- Similarly, if we define a parameter with two leading stars **, it captures any extra named parameters


### Further notes
- Each function is an object in memory as well!
- An alias is a second or subsequent reference to the same object. The same object (eg a function) can then be called under both reference names
- argument versus parameter: 
    - arguments: values passed into a function
    - parameters: names the function uses to refer to them. 


## Chapter 6: Running Tests 
>→ all
>- Functions are objects you can save in data structures or pass to other functions.
>- Python stores local and global variables in dictionary-like structures.
>- A unit test performs an operation on a fixture and `passes`, `fails`, or produces an `error`.
>- A program can use introspection to find functions and other objects at runtime.

### How tests work, unit tests
**Unit Test definition**: A test that exercises **one function or feature** of a piece of software and produces pass, fail, or error.

- **Fixture:** The thing on which a test is run, such as the parameters to the function being tested or the file being processed

- **Assertion:** A Boolean expression that must be true at a certain point in a program. Assertions may be built into the language or provided as functions.

**Test functionality:** Each test does something to a fixture (such as the number 19) and uses assertions to compare the actual result against the expected result. The outcome of each test can be:
- **Pass**: the test subject works as expected.
- **Fail**: something is wrong with the test subject.
- **Error**: something is wrong in the test itself, which means we don’t know if the thing we’re testing is working properly or not.


Evaluated through the following logic:

1. If a test function completes without raising any kind of exception, it passes. (We don’t care if it returns something, but by convention tests don’t return a value.)
2. If the function raises an AssertionError exception, then the test has failed. Python’s assert statement does this automatically when the condition it is checking is false, so almost all tests use assert for checks.
3. If the function raises any other kind of exception, then we assume the test itself is broken and count it as an error.

translated to cvode:
```py3
def run_tests(all_tests):
    results = {"pass": 0, "fail": 0, "error": 0}
    for test in all_tests:
        try: # = 1. 
            test()
            results["pass"] += 1
        except AssertionError: #2. 
            results["fail"] += 1
        except Exception: #3.
            results["error"] += 1
    print(f"pass {results['pass']}")
    print(f"fail {results['fail']}")
    print(f"error {results['error']}")
```

### Introspection
To avoid passing functions as lists, we want the test runner to find the tests by itself. For this, we use introspection and name each test `test_`

globals() is a dictionary and shows all variables in the gobal scope (see below). As this is the case, we can just iterate thorugh this globals() dictionary to find the testing fucntions

```py3
def find_tests(prefix):
    for (name, func) in globals().items():
        if name.startswith(prefix):
            print(name, func)

find_tests("test_")
```
output:
>test_sign_negative <function test_sign_negative at 0x105bcd440> \
>test_sign_positive <function test_sign_positive at 0x105bcd4e0> \
>test_sign_zero <function test_sign_zero at 0x105bcd580> \
>test_sign_error <function test_sign_error at 0x105bcd620> 

The two code snippets above combined create a test runner using introspection:

```py3
def run_tests():
    results = {"pass": 0, "fail": 0, "error": 0}
    for (name, test) in globals().items():
        if not name.startswith("test_"):
            continue
        try:
            test()
            results["pass"] += 1
        except AssertionError:
            results["fail"] += 1
        except Exception:
            results["error"] += 1
    print(f"pass {results['pass']}")
    print(f"fail {results['fail']}")
    print(f"error {results['error']}")

```



### Further notes
**Pretty Print (pprint)** prints nicely formatted output. For example we can inspect globals()
```py3
pprint.pprint(globals())
{'__annotations__': {},
 '__builtins__': <module 'builtins' (built-in)>,
 '__cached__': None,
 '__doc__': None,
 '__file__': '/sdx/test/globals.py',
 '__loader__': <_frozen_importlib_external.SourceFileLoader object \
at 0x109d65290>,
 '__name__': '__main__',
 '__package__': None,
 '__spec__': None,
'my_variable': 123,
 'pprint': <module 'pprint' from \
'/sdx/conda/envs/sdxpy/lib/python3.11/pprint.py'>}
```

## Chapter 7: An Interpreter 
> → all
> - Compilers and interpreters are just programs.
> - Basic arithmetic operations are just functions that have special notation
> - Programs can be represented as trees, which can be stored as nested lists.
> - Interpreters recursively dispatch operations to functions that implement low-level steps.
> - Programs store variables in stacked dictionaries called environments.
> - One way to evaluate a program's design is to ask how extensible it is.

### Terms

**compiler:** An application that translates programs written in some languages into machine instructions or bytecode.

**bytecode** A set of instructions designed to be executed efficiently by an interpreter.

**interpreter** A program that runs programs written in a high-level interpreted language. Interpreters can run interactively but may also execute commands saved in a file.

**interpreted language**  A high-level language that is not executed directly by the computer, but instead is run by an interpreter that translates program instructions into machine commands on the fly.

**runtime**    A program that implements the basic operations used in a programming language.

**notation types**   
- infix notation `1 + 2` to add 1 and 2.
- prefix notation `+ 3 4` to add 3 and 4.
- postfix notation `2 3 +` to add 2 and 3.


### Building expressions
- each expressen is a list with [op_name, values]
- nested lists for multiple operations
- interpreter uses prefix notation: easier to find operations
- not using symbols like (+, - *) but names, it is the compiler's job to convert this
```
["add", 1, 2]            # 1 + 2
["abs", -3.5]            # abs(-3.5)
["add", ["abs", -5], 9]  # abs(-5) + 9
```

Code Idea: we have all functions defined with the prefix `do_`. This prefix prevents colisions with built-in PYthon funcs as well. These functions are called with `args`, so unnamed arguments. First they assert the correct number of arguments, then they perform the operation and then return the result.

Dynamic dispatch (see below) is used as the program decides who to give work to on the fly. 

### Variables
To work with variables, we need to pass the environment with the function. In addition to that, we need to have setter and getter methods tomake new or retreive existing variables 

### Making sequences
To have several expressions that set variables and wotk on them, we need a function to run a sequence of expressions one by one. This is a control flow that controls when and how other expressions are evaluated.

With the sequence runner qe can now evaluate somethingh like:

```
[
    "seq", 
    ["set", "alpha", 1],
    ["set", "beta", 2],
    ["add", ["get", "alpha"], ["get", "beta"]]
]
```
> Note: Python distinguishes expressions that produce values from statements that don’t. But it doesn’t have to, and many languages don’t. This is a design choice

### Introspection
We can improve the do function by using introspection. This may be don as all our functions start with `do_`. This removes the sequence of ifs, It's exactly the same as in the testing part. 

Here we use introspection to create a lookup dict that stores every function that starts with do_.

The lookup dict contains `key: value` pairs with `name: function`. Similar to this
```
OPS = {
    "abs": do_abs,
    "add": do_add,
    "get": do_get,
    "seq": do_seq,
    "set": do_set,
}
```

### Final code
```py3
def do_add(args):
    assert len(args) == 2
    left = do(args[0])
    right = do(args[1])
    return left + right


def do_abs(env, args):
    assert len(args) == 1
    val = do(env, args[0])
    return abs(val)

# VARIABLE GETTER
def do_get(env, args):
    assert len(args) == 1
    assert isinstance(args[0], str)
    assert args[0] in env, f"Unknown variable {args[0]}"
    return env[args[0]]

# VARIABLE SETTER
def do_set(env, args):
    assert len(args) == 2
    assert isinstance(args[0], str)
    value = do(env, args[1])
    env[args[0]] = value
    return value

# SEQUENCE RUNNER
def do_seq(env, args):
    assert len(args) > 0
    for item in args:
        result = do(env, item)
    return result


###### Lookup Table generation using dictionary comprehensions
OPS = {
    name.replace("do_", ""): func 
    for (name, func) in globals().items()
    if name.startswith("do_")
}



###### SIMPLE evaluation function using dynamic dispacth and recursion
def do(expr):
    # Integers evaluate to themselves.
    if isinstance(expr, int):
        return expr

    # Lists trigger function calls.
    assert isinstance(expr, list)
    assert expr[0] in OPS, f"Unknown operation {expr[0]}"
    func = OPS[expr[0]]
    return func(env, expr[1:]) #recursive call

##### MAIN function to call interpreter with a file (json) containing operations and calling do and print
def main():
    assert len(sys.argv) == 2, "Usage: expr.py filename"
    with open(sys.argv[1], "r") as reader:
        program = json.load(reader)
    result = do(program)
    print(f"=> {result}")

if __name__ == "__main__":
    main()

```



### Further notes

**Dynamic dispatch:** to find a function or a property of an object by name while a program is running. For example, instead of getting a specific property of an object using obj.name, a program might use obj[someVariable], where someVariable could hold "name" or some other property name.

**environment**  The set of variables currently defined in a program.

## Chapter 8: Functions (and Closures) 
>→ all except section '3. Closures'
> -    When we define a function, our programming system saves instructions for later use.
> -    Since functions are just data, we can separate creation from naming.
> -    Most programming languages use eager evaluation, in which arguments are evaluated before a function is called.
> -    Programming languages can also use lazy evaluation, in which expressions are passed to functions for just-in-time evaluation.
> -    Every call to a function creates a new stack frame on the call stack.
> -    When a function looks up variables it checks its own stack frame and the global frame.
> -    A closure stores the variables referenced in a particular scope.


We want to extend the interpreter from the previous chapter to be able to define and call our own functions as well, not just using predefined functions. 

### Definition and Storage

We need to replace the definition structure of python and do so with the term "set" and "func". the first option is an unnamed function (easy as it is an object too), in the second one, we name it for later use.

### Calling Functions
For the named case, we also need a method to later call it. We name this call 

The system of a call is
1. Evaluate all of these expressions.
2.    Look up the function.
3.    Create a new environment from the function’s parameter names and the expressions’ values.
4.    Call do to run the function’s action and capture the result.
5.    Discard the environment created in Step 3.
6.    Return the function’s result.

Now, to make this work, we need to implement **scoping**. This is needed to prevent name collisions. Here it means, that the environment, which was a simple dict, must now be a list of dictionaries. This list is the **call stack** with each dict being a **stack frame**. When searching for a value, we iterate through the frames from most recent to oldest

>do_call contains the line:
>
>`env.append(dict(zip(params, values)))`
>
>Working from the inside out, it uses the **built-in function zip to create a list of pairs of corresponding items from params and values,** then passes that list of pairs to dict to create a dictionary, which it then appends to the list env. The exercises will explore whether rewriting this would make it easier to read.

### Final Code
```py3
def same(num):
    return num

#define a func  unnamed
["func", ["num"], ["get", "num"]]

#define a named function for later use
["set", "same", ["func", ["num"], ["get", "num"]]]

#call the named function with a value
["call", "same", 3]

#-------------------------------- new implementation

def do_func(env, args):
    assert len(args) == 2
    params = args[0]
    body = args[1]
    return ["func", params, body]

def do_call(env, args):
    # Set up the call.
    assert len(args) >= 1
    name = args[0]
    values = [do(env, a) for a in args[1:]]

    # Find the function.
    func = env_get(env, name)
    assert isinstance(func, list) and (func[0] == "func")
    params, body = func[1], func[2]
    assert len(values) == len(params)

    # Run in new environment.
    env.append(dict(zip(params, values)))
    result = do(env, body)
    env.pop()

    # Report.
    return result
```
now we can run stuff like
```
["seq",
  ["set", "double",
    ["func", ["num"],
      ["add", ["get", "num"], ["get", "num"]]
    ]
  ],
  ["set", "a", 1],
  ["repeat", 4, ["seq",
    ["set", "a", ["call", "double", ["get", "a"]]],
    ["print", ["get", "a"]]
  ]]
]
```
### Further notes
**anonymous function:**    A function without a name. Languages like JavaScript make frequent use of anonymous functions; Python provides a limited form called a lambda expression.

**lambda expression**    An expression that takes zero or more parameters and produces a result. A lambda expression is sometimes called an anonymous function; the name comes from the mathematical symbol λ used to represent such expressions.

**eager evaluation**: evaluating a function's arguments before we run

**lazy evaluation**: Evaluating expressions only when absolutely necessary, during the run. We would pass the argument sub-lists into the function and let it evaluate them when it needed their values.

>*Python and most other languages are eager, but a handful of languages, such as R, are lazy.*

Example:
```py3
double = lambda x: 2 * x
double(3)
```

**dynamic scoping** To find the value of a variable by looking at what is on the call stack at the moment the lookup is done. Almost all programming languages use lexical scoping instead, since it is more predictable.

**lexical scoping**     To look up the value associated with a name according to the textual structure of a program.

> *Python etc use lexical scoping but our interpreter here uses dynamic scoping (easier implementation)*

## Chapter 9: Protocols
>→ all except section '3. Decorators'
>-    Temporarily replacing functions with mock objects can simplify testing.
>-    Mock objects can record their calls and/or return variable results.
>-    Python defines protocols so that code can be triggered by keywords in the language.
>-    Use the context manager protocol to ensure cleanup operations always execute.
>-    Use decorators to wrap functions after defining them.
>-    Use closures to create decorators that take extra parameters.
>-    Use the iterator protocol to make objects work with for loops.

### Mock Objects
Functions are objects referred to by variable names, this can be used to change functions at runtime to **make testing easier**. 

For example time: here we can replace the real timt.time function with a specific value that can be tested 

**Mock Objects:** A simplified replacement for part of a program whose behavior is easy to control and predict. Mock objects are used in unit tests to simulate databases, web services, and other complex systems.

```py3
#simple mock object for time
import time

def elapsed(since):
    return time.time() - since

def mock_time():
    return 200

def test_elapsed():
    time.time = mock_time #mocking
    assert elapsed(50) == 150
```

If an object obj has a `__call__` method, then obj(…) is automatically turned into `obj.__call__(…)` just as a == b is automatically turned into `a.__eq__(b)` 

So we can create a mock object class of the following system
1. defines a `__call__` method so that instances can be called like functions;

1. declares the parameters of that method to be *args and **kwargs so that it can be called with any number of regular or keyword arguments;

3. stores those arguments so we can see how the replaced function was called; and

4. returns either a fixed value or a value produced by a user-defined function

We use this together wite a function to make it more easy to use:

```py3
class Fake:
    def __init__(self, func=None, value=None):
        self.calls = [] #list tracking the calls and 'surviving'
        self.func = func #None
        self.value = value #99

    def __call__(self, *args, **kwargs):
        self.calls.append([args, kwargs]) #call added to list
        if self.func is not None:
            return self.func(*args, **kwargs)
        return self.value

function to use the fake class, works with a second function or fixed value
def fakeit(name, func=None, value=None):
    assert name in globals()
    fake = Fake(func, value) #make instance of fake class
    globals()[name] = fake #replace in globals -> globals[adder] = fake
    return fake


#-------------Example use
#func to test the mock class
def adder(a, b):
    return a + b

#normal use 2 + 3 = 5
def test_with_real_function():
    assert adder(2, 3) == 5

#use fakeit, fixed value 99
def test_with_fixed_return_value():
    fakeit("adder", value=99)
    assert adder(2, 3) == 99 #= assert fake(2,3) == 99

```
![mock operation system with Fake class](image.png)

### Protocols
Mock objects have one big issue that will lead to errors: each test replaces the function with the mock object (here adder).
Thus, any test using the original function will fail (as fake is called). We would have to set back and revert to the original function. Instead of by hand // user input, we can use a protocol provided by Python.

**protocol:**    A rule that specifies how programs can tell Python to do specific things at specific moments.

Examples are
- when there is a `__call__` method and thing() is used, Python checks if that method exists
- if `__init__` is defined, it is automatically called when a new instance is created

**Format:**
```
with C(…args…) as name:
    …do things…
```
This 
1. calls constructor of C and creates object
2. call the objects `__enter__`, the res of enter is assigned to the var name
3. code in `with` is run
4. `__exit__` is called after that block

So we use this to make a **context manager** (object that automatically executes some operations at the start of a code block and some other operations at the end of the block.). for our Fake class

```py3
#new fake class with protocols, still bases on fake for init and call
class ContextFake(Fake):
    def __init__(self, name, func=None, value=None):
        super().__init__(func, value)
        self.name = name
        self.original = None

    #no extra params, all provided via constructior
    def __enter__(self): 
        assert self.name in globals()
        self.original = globals()[self.name]
        globals()[self.name] = self
        return self

    #three params to handle exeptions
    def __exit__(self, exc_type, exc_value, exc_traceback): 
        globals()[self.name] = self.original


# CALL on protocol version
def subber(a, b):
    return a - b

def check_no_lasting_effects():
    assert subber(2, 3) == -1 #normal func
    with ContextFake("subber", value=1234) as fake:
        assert subber(2, 3) == 1234 #faked 
        assert len(fake.calls) == 1 #call counter
    assert subber(2, 3) == -1 #normal func again

```
### Iterators
Iterators are also examples of protocols. The iterator protocol ahs two parts, built as follows:

1. If an object has an `__iter__` method, that method is called to create an iterator object. It must always return self

2. That iterator object must have a `__next__` method, which must return a value each time it is called. When there are no more values to return, it must raise a StopIteration exception.

```py3
class NaiveIterator:
    def __init__(self, text):
        self._text = text[:]

    def __iter__(self):
        self._row, self._col = 0, -1
        return self

    def __next__(self):
        self._advance() #forward within row 
        if self._row == len(self._text): #if end of row -> next row
            raise StopIteration
        return self._text[self._row][self._col]

 #---------- simple example for a class with iterator:
 #prints numbers 1-20
 class MyNumbers:
    def __iter__(self):
        self.a = 1
        return self

    def __next__(self):
        if self.a <= 20:
            x = self.a
            self.a += 1
            return x
        else:
            raise StopIteration

myiter = iter(myclass)
for x in myiter:
    print(x)
```

But above doesn't work for rows and colums.



## Chapter 10: A File Archiver 
>→ all
>-   Version control tools use hashing to uniquely identify each saved file.
>-    Each snapshot of a set of files is recorded in a manifest.
>-    Using a mock filesystem for testing is safer and faster than using the real thing.
>-    Operations involving multiple files may suffer from race conditions.
>-    Use a base class to specify what a component must be able to do and derive child classes to implement those operations.

We only want to (re)archive a file if it changed. To prevent that, we hash the file. This means we produce a short identifier (hash) to check if the content of two files is equal.

For each file we generate a ["Hash"].bck file. Then, the original filenames and hash keys are saved in each snapshot. So we now the filename plus the filecontent of a file in a snapshot.
To restore, the .bck is copied back to its original location

### Hash all files 
With the following function, all files below a root are hashed, using the python glob module . The hash is cut to the first 16 digits (just here for simplicity)
```py3
HASH_LEN = 16
def hash_all(root):
    result = []
    for name in glob("**/*.*", root_dir=root, recursive=True):
        full_name = Path(root, name)
        with open(full_name, "rb") as reader:
            data = reader.read()
            hash_code = sha256(data).hexdigest()[:HASH_LEN]
            result.append((name, hash_code))
    return result
```
Given folder 
```
sample_dir
|-- a.txt
|-- b.txt
`-- sub_dir
    `-- c.txt
```
get get the output of `python hash_all.py sample_dir`
```
filename,hash
b.txt,3cf9a1a81f6bdeaf
a.txt,17e682f060b5f8e4
sub_dir/c.txt,5695d82a086b6779
```
### Testing
Testing can be done using the module  `pyfakefs`, this is a mock object for file system (similar to mock objects in previous chapter)

This is needed to make sure early tests don’t contaminate later ones, otherwise we would have to recreate those files and directories after each test.

`pyfakefs` replaces key functions like open with functions that behave the same way but act on “files” stored in memory. And it is much faster. 

By importing the lib, we get `fs` that we can use to create files. we pass it to pytest as well to write tests
```py3
from pathlib import Path
import pyfakefs 
import pytest 

def test_simple_example(fs):
    sentence = "This file contains one sentence."
    with open("alpha.txt", "w") as writer:
        writer.write(sentence)
    assert Path("alpha.txt").exists()
    with open("alpha.txt", "r") as reader:
        assert reader.read() == sentence

```

### Tracking Backups
To track which files have and haven’t been backed up, we make a file that contains the .bck files and create a manifest

The manifest describes the content of each snapshot. The manifest is named `ssssssss.csv` with ssssss being the timestamp (UTC) of the backup (fails if two backups are created in the same second!). This manifest uses the hash_all function to generate its content

### Backup system
```py3
def write_manifest(backup_dir, timestamp, manifest):
    backup_dir = Path(backup_dir) #make Path item
    if not backup_dir.exists():
        backup_dir.mkdir() #create dir if not there
    manifest_file = Path(backup_dir, f"{timestamp}.csv") #filename = timestamp
    with open(manifest_file, "w") as raw:
        writer = csv.writer(raw)
        writer.writerow(["filename", "hash"])
        writer.writerows(manifest)

def copy_files(source_dir, backup_dir, manifest):
    for (filename, hash_code) in manifest:
        source_path = Path(source_dir, filename)
        backup_path = Path(backup_dir, f"{hash_code}.bck")
        if not backup_path.exists():
            shutil.copy(source_path, backup_path)


def backup(source_dir, backup_dir):
    manifest = hash_all(source_dir) #as seen above, csv, list of filename,hash
    timestamp = current_time() #time now in utc
    write_manifest(backup_dir, timestamp, manifest) #write list to file
    copy_files(source_dir, backup_dir, manifest) #copy changed files
    return manifest
```

This can also be put into a base class now:
```py3
class Archive:
    def __init__(self, source_dir):
        self._source_dir = source_dir

    def backup(self):
        manifest = hash_all(self._source_dir)
        self._write_manifest(manifest)
        self._copy_files(manifest)
        return manifest
```
with this as base class, we can make child classes for different archiving types like local vs remote: `archiver = ArchiveLocal(source_dir, backup_dir)`

![Backup System Chart](image-1.png)

### Further notes

**successive refinement:** Writing a high-level function first and then filling in the things it needs, synonym for top-down design


## Chapter 16: Object Persistence 
>→ all (including '3. Aliasing', which was assigned for self-study)
>-    A persistence framework saves and restores objects.
>-    Persistence must handle aliasing and circularity.
>-    Users should be able to extend persistence to handle objects of their own types.
>-    Software designs should be open for extension but closed for modification.

Two options to sotre data persistent:
- pickle: very Python-like
- json: widespred, close to JavasScript where it originated from

A persistence framework needs to choose one of the following options: 
1. Only handle built-in types, or even only handle types that are common across many languages, so that data saved by Python can be read by JavaScript and vice versa.
2. Provide a way for programs to convert from user-defined types to built-in types and then save those. This is less restrictive than (1) but information might be lost
3. Save class definitions as well as objects’ values so that when a program reads saved data it can reconstruct the classes and then create fully functional instances of them

(3) is most powerfulbut hardest to implement especially for multi-language support

### Built-in types (option 1)
```py3
def save(writer, thing):
    if isinstance(thing, bool): #easy
        print(f"bool:{thing}", file=writer)

    elif isinstance(thing, float): #easy
        print(f"float:{thing}", file=writer)

    elif isinstance(thing, int): #easy
        print(f"int:{thing}", file=writer)

    elif isinstance(thing, list):
        print(f"list:{len(thing)}", file=writer) #print "list: length"
        for item in thing: #add list items
            save(writer, item) #recursive call 

    elif isinstance(thing, dict):
        print(f"dict:{len(thing)}", file=writer) #print "dict length"
        for (key, value) in thing.items():
            save(writer, key) #recursive call 
            save(writer, value) #recursive call 

    else:
        raise ValueError(f"unknown type of thing {type(thing)}")

def load(reader):
    line = reader.readline()[:-1]
    assert line, "Nothing to read"
    fields = line.split(":", maxsplit=1)
    assert len(fields) == 2, f"Badly-formed line {line}"
    key, value = fields

    if key == "bool":
        names = {"True": True, "False": False}
        assert value in names, f"Unknown Boolean {value}"
        return names[value]

    elif key == "float":
        return float(value)

    elif key == "int":
        return int(value)
    
    elif key == "list":
    return [load(reader) for _ in range(int(value))]

    else:
        raise ValueError(f"unknown type of thing {line}")
```
output:

```
save(sys.stdout, [False, 3.14, "hello", {"left": 1, "right": [2, 3]}])
list:4
bool:False
float:3.14
str:1
hello
dict:2 #dict starts here and contains the next 2*2 items
str:1 
left 
int:1 
str:1 
right
list:2 #list starts here and includes the next two items
int:2
int:3

```

### Converting to classes
We can rewrite above functions as classes. And by using dynamic dispatch to handle an item without a separate if statement for each type we can further improve it 
```py3


class SaveObjects:
    def __init__(self, writer):
        self.writer = writer

    #we define the methods as befor but it must be *save_*
    def save_int(self, thing):
        self._write("int", thing)

    def save_str(self, thing):
        lines = thing.split("\n")
        self._write("str", len(lines))
        for line in lines:
            print(line, file=self.writer)

    #here we define which method we choose to save
    def save(self, thing):
        typename = type(thing).__name__ #get type of element
        method = f"save_{typename}"
        #check if we have an function called (save_typename)
        assert hasattr(self, method), \ 
            f"Unknown object type {typename}"
        getattr(self, method)(thing)

class LoadObjects:
    def __init__(self, reader):
        self.reader = reader

    def load_float(self, value):
    return float(value)

    def load(self):
        line = self.reader.readline()[:-1]
        assert line, "Nothing to read"
        fields = line.split(":", maxsplit=1)
        assert len(fields) == 2, f"Badly-formed line {line}"
        key, value = fields
        method = f"load_{key}"
        assert hasattr(self, method), f"Unknown object type {key}"
        return getattr(self, method)(value)
```

### Aliasing
**Issue:**
a list shared = ["content"] and a list fixture = [shared, shared]

The list fixed points to the same object twice
But if we save this and reload it, the two lists are built separately, whcih resluts in two different copies of ["content"]. This causes issue as chages are not made on both.

To solve we need to:
1. remember during saving which objects are already saved 
2. if an object should be save the second time, we do not save another object but an alias (Verweis) to the first
3. When re-loading the aliases go back to the correct single object (reverse)

for this we use the id function (built-in)

So we store the IDs of all objects we already saved and write an entry with alias and the id if we meet an object the second time.

To make this work, we need to save each objects id as well.
```py3
class SaveAlias(SaveObjects):
    def __init__(self, writer):
        super().__init__(writer)
        self.seen = set() #track already saved objects

    #the saving methods now save type, id, len/value
    def save_list(self, thing):
    self._write("list", id(thing), len(thing))
    for item in thing:
        self.save(item)


    def save(self, thing):
        thing_id = id(thing) #get id
        if thing_id in self.seen: #if already seen: "alias id", stop
            self._write("alias", thing_id, "")
            return

        self.seen.add(id(thing)) #else add id to seen
        typename = type(thing).__name__ #get type
        method = f"save_{typename}" #save with correct method
        assert hasattr(self, method), f"Unknown object type {typename}"
        getattr(self, method)(thing) 

#saving format:     type:id:len/value
#                   list:4484025600:1
#                   alias:4484025600:
#                   str:4539552048:1
#                   word

class LoadAlias(LoadObjects):
    def __init__(self, reader):
        super().__init__(reader)
        self.seen = {}

    def load_list(self, ident, length): #ensuring recursion works
        result = []
        self.seen[ident] = result
        for _ in range(int(length)):
            result.append(self.load())
        return result


    def load(self):
        line = self.reader.readline()[:-1]
        assert line, "Nothing to read"
        fields = line.split(":", maxsplit=2)
        assert len(fields) == 3, f"Badly-formed line {line}"
        key, ident, value = fields

        # the lines below contain a bug
        if key == "alias":
            assert ident in self.seen
            return self.seen[ident]

        method = f"load_{key}"
        assert hasattr(self, method), f"Unknown object type {key}"
        result = getattr(self, method)(value)
        self.seen[ident] = result
        return result

```


### Further notes
**persistence:**     The act of saving and restoring data, particularly heterogeneous data with irregular structure.

**Open-Closed Principle:**    A design rule stating that software should be open for extension but closed for modification, i.e., it should be possible to extend functionality without having to rewrite existing code.



## Chapter 17: Binary Data 
>→ all
>- Programs usually store integers using two's complement rather than sign and magnitude.
>- Characters are usually encoded as bytes using either ASCII, UTF-8, or UTF-32.
>- Programs can use bitwise operators to manipulate the bits representing data directly.
>- Low-level compiled languages usually store raw values, while high-level interpreted languages use boxed values.
>- Sets of values can be packed into contiguous byte arrays for efficient transmission and storage.


### Different Data Types
#### Integers
- usually saved as base 2 in binary --> 9 (in base 10) = 1001
- sign and magnitude: for negative numbers, we use the top bit for the sign --> 01001 = +9, 11001 = -9
    - two zeros are possible (+0/-0)
    - more complicated hardware needed
- two's complement: rolls over when going below zero like an odometer
    - 2: 010, 1: 001, 0: 000, -1: 111, -2: 110
    - looking at the first bit still tells us the sign (1 = neg)
    - as it is asymetric (0 is pos), we range from -4 to 3 here (or similar)

**We can write binary directly in python using 0b prefix**
`print(0b101101) = 45`

- hexadecimal:  we sum up four bits (=range 0-15) and convert to 0-f
    - to write out we again go via binary: F7 = 1111 0001 

**We can write hexadecimal directly in python using 0x prefix**
`print(0xf7) = 247,  print(0b11110111) = 247`

#### Bitwise Operations
> An operation that manipulates individual bits in memory. 
- & (and):   yields a 1 only if both its inputs are 1’s
- | (or):   yields 1 if either or both are 1
- ^ (xor):  called exclusive or or “xor” (pronounced “ex-or”), produces 1 if the bits are different 
- ~ (not): flips its argument: 1 becomes 0, and 0 becomes 1
- bit shifting operators: move bits left or right
    - 0110 << 1  = 1100 (equal to *2)
    - 0110 >> 1 = 0011 (equal to /2, throwing away the remainder)

> Python always fills with zeroes, while Java provides two versions of right shift: >> fills in the high end with zeroes, while >>> copies in the topmost (sign) bit of the original value. 

Examples:
- 12 & 6 	-> 1100 & 0110 -> 0100 -> 	4
- ~ 6 	-> ~ 0110	-> 1001	-> 9
- 12 << 2	-> 1100 << 2 -> 110000 	48

#### Text
We find different character encodings:

ASCII
- unaccented Latin chars with numbers 32-127 (7bits)
- 128-255 (8th bit) for other special chars, used differently

ANSI
- Standard for chars 128-255 
- solved a small part of a large problem, many chars still nbot supported
- we could go up to 16/32 bits per char buit all ANSI would be invalid, all docs 4x larger

Unicode
- code point for every character like U+0065 or  U+2605
- each of these code points was defined how to store 
    - UTF-32: each char is a 32bit number - wastes memory if only ASCI chars used
    - UTF-8: variable length, most popular  all points 0-127 in single byte / 8 bit like in ASCII 
        - if top bit is 1, the bits after the top bit before the first 0 tells how many more bytes are used for the next char
        - 11101101 = multi-byte: 1, 2 next bytes as well: 11, separation: 0, actual value: 1101

### Persistance
Why Binary / machine code is better than human-readable text:
- binary is way more memory efficient than saving human-readable text
- integer/float operations are much faster than operations on strings 
- no better solution

To save abd load binary, we need to read it correctly. 
If we use `open("filename", "r")`, Python assumes we want to read character strings from the file. Thus it asks OS for the default (UTF-8) char encoding, uses this to convert and converts end of line markers if necessary (\r -> \n).

This works for text but not with binary (for example a png read). we can use `open(filename, "rb")`, then Python reads proper binary as bytes objects instead of strings. But here we wnat to use `reader.read(N)` to read N bytes at a time rather than for line in reader because there aren’t actually lines of text in the file.

Python and other dynamic languages, on the other hand, put each value in a data structure that keeps track of its type along with a bit of extra administrative information. Something stored this way is called a **boxed value**, and this extra information is what **allows the interpreter to do introspection at runtime**.

The **format string** specifies what types of data are being packed, how big they are (e.g., is this a 32-bit or 64-bit floating point number?), and how many values there are, which in turn exactly determines how much memory is required by the packed representation.

#### The struct module

This Python module packs and unpacks data for us to and from binary. It takes a format string and a bunch of values as arguments and packs them into a bytes object.
```py3
import struct

fmt = "ii"  # two 32-bit integers
x = 31
y = 65

binary = struct.pack(fmt, x, y)  # pack(format, val_1, val_2, …)
print("binary representation:", repr(binary))
#binary representation: b'\x1f\x00\x00\x00A\x00\x00\x00'

normal = struct.unpack(fmt, binary)
print("back to normal:", normal)
#back to normal: (31, 65)


print(pack("3i", 1, 2, 3))
# b'\x01\x00\x00\x00\x02\x00\x00\x00\x03\x00\x00\x00'
print(pack("5s", bytes("hello", "utf-8")))
# b'hello'
print(pack("5s", bytes("a longer string", "utf-8")))
# b'a lon' -- worng length, must get correct one
```


The `x1f`: Python finds a byte in a string that doesn’t correspond to a printable character, it prints a 2-digit escape sequence in hexadecimal. Python is therefore telling us that our string contains the eight bytes

Struct types:
- "c" 	Single character (i.e., string of length 1)
- "B" 	Unsigned 8-bit integer with all 8 bits used for value
- "h" 	16-bit integer
- "i" 	32-bit integer
- "d" 	64-bit float

Any format can be preceded by a count, so the format "3i" means “three integers”

To get rid of wrong size issues during packing and unpacking, we save the size along with the data. If we always use exactly the same number of bytes to store the size, we can read it back safely

 ```py3
 def pack_string(as_string):
    as_bytes = bytes(as_string, "utf-8")
    header = pack("i", len(as_bytes))
    format = f"{len(as_bytes)}s"
    body = pack(format, as_bytes)
    return header + body

def unpack_string(buffer):
    header, body = buffer[:4], buffer[4:]
    length = unpack("i", header)[0]
    format = f"{length}s"
    result = unpack(format, body)[0]
    return str(result, "utf-8")

#In practice, programmers use the struct module’s calcsize function to figure out how large (in bytes) the data represented by a format is:
from struct import calcsize

for format in ["4s", "3i4s5d"]:
    print(f"format '{format}' needs {calcsize(format)} bytes")

 ```

## Chapter 25: A Virtual Machine 
>→ all
>-   Every computer has a processor with a particular instruction set, some registers, and memory.
>-    Instructions are just numbers but may be represented as assembly code.
>-   Instructions may refer to registers, memory, both, or neither.
>-    A processor usually executes instructions in order but may jump to another location based on whether a conditional is true or false.

### parts of a VM
- **Instruction Pointer:** holds the memory address of the next instruction to execute. It is automatically initialized to point at address 0, so that is where every program must start. This requirement is part of our VM’s Application Binary Interface (ABI).
- **Register R0-R3**  instructions can access directly. There are no memory-to-memory operations in our VM: everything happens in or through registers.
- **Words (256)** memory, each of which can store a single value. Both the program and its data live in this single block of memory; we chose the size 256 so that the address of each word will fit in a single byte.
- **instruction set**, three bytes each, defines what it can do. Instructions are just numbers, but we will write them in a simple text format called assembly code that gives those number human-readable names, r = register identifier | v = constant value
```py3
NUM_REG = 4  # number of registers
RAM_LEN = 256  # number of words in RAM

OPS = {
    "hlt": {"code": 0x1, "fmt": "--"},  # Halt program
    "ldc": {"code": 0x2, "fmt": "rv"},  # Load value
    "ldr": {"code": 0x3, "fmt": "rr"},  # Load register
    "cpy": {"code": 0x4, "fmt": "rr"},  # Copy register
    "str": {"code": 0x5, "fmt": "rr"},  # Store register
    "add": {"code": 0x6, "fmt": "rr"},  # Add
    "sub": {"code": 0x7, "fmt": "rr"},  # Subtract
    "beq": {"code": 0x8, "fmt": "rv"},  # Branch if equal
    "bne": {"code": 0x9, "fmt": "rv"},  # Branch if not equal
    "prr": {"code": 0xA, "fmt": "r-"},  # Print register
    "prm": {"code": 0xB, "fmt": "r-"},  # Print memory
}

OP_MASK = 0xFF  # select a single byte
OP_SHIFT = 8  # shift up by one byte
OP_WIDTH = 6  # op width in characters when printing

class VirtualMachine:
    def __init__(self):
        self.initialize([]) #calls initialize with an empty array
        self.prompt = ">>"

    def initialize(self, program):
        assert len(program) <= RAM_LEN, "Program too long"
        self.ram = [
            program[i] if (i < len(program)) else 0
            for i in range(RAM_LEN)
        ]
        self.ip = 0 #instruction pointer
        self.reg = [0] * NUM_REG

    #to execute the next instruction, getting current pointer, moves it by 1
    #bitwise operations to extract the op code and operands from the instruction
    def fetch(self):
        instruction = self.ram[self.ip]
        self.ip += 1
        op = instruction & OP_MASK
        instruction >>= OP_SHIFT
        arg0 = instruction & OP_MASK
        instruction >>= OP_SHIFT
        arg1 = instruction & OP_MASK
        return [op, arg0, arg1]

    def run(self):
    running = True
    while running:
        op, arg0, arg1 = self.fetch()
        if op == OPS["hlt"]["code"]:
            running = False
        elif op == OPS["ldc"]["code"]:
            self.reg[arg0] = arg1
        elif op == OPS["ldr"]["code"]:
            self.reg[arg0] = self.ram[self.reg[arg1]]
        elif op == OPS["cpy"]["code"]:
            self.reg[arg0] = self.reg[arg1]
        elif op == OPS["str"]["code"]:
            self.ram[self.reg[arg1]] = self.reg[arg0]
        elif op == OPS["add"]["code"]:
            self.reg[arg0] += self.reg[arg1]
        elif op == OPS["beq"]["code"]: #conditional jump
            if self.reg[arg0] == 0:
                self.ip = arg1
        else:
            assert False, f"Unknown op {op:06x}"


```

### Assembly code 
Much easier to use an assembler, which is just a small compiler for a language that very closely represents actual machine instructions.

Each command in our assembly languages matches an instruction in the VM.

```
prr R1
hlt

00010a
000001
```

One thing the assembly language has that the instruction set doesn’t is **labels on addresses** in memory. The label loop doesn’t take up any space; instead, it tells the assembler to give the address of the next instruction a name so that we can **refer to @loop** in jump instructions. For example, this program prints the numbers from 0 to 2

### Arrays
We can do a lot more once we have arrays, so let’s add those to our assembler. We don’t have to make any changes to the virtual machine, which doesn’t care if we think of a bunch of numbers as individuals or elements of an array, but we do need a way to create arrays and refer to them

```py3
class Assembler:
    def assemble(self, lines):
        lines = self._get_lines(lines)
        to_compile, to_allocate = self._split(lines) #for arrays

        labels = self._find_labels(lines)
        instructions = [
            ln for ln in lines if not self._is_label(ln)
        ]

        base_of_data = len(instructions) #for arrays
        self._add_allocations(base_of_data, labels, to_allocate) #for arrays

        compiled = [
            self._compile(instr, labels) for instr in instructions
        ]
        program = self._to_text(compiled)
        return program

    def _find_labels(self, lines):
    result = {}
    loc = 0
    for ln in lines:
        if self._is_label(ln):
            label = ln[:-1].strip()
            assert label not in result, f"Duplicated {label}"
            result[label] = loc
        else:
            loc += 1
    return result

    def _is_label(self, line):
        return line.endswith(":")

    def _compile(self, instruction, labels):
    tokens = instruction.split()
    op, args = tokens[0], tokens[1:]
    fmt, code = OPS[op]["fmt"], OPS[op]["code"]

    if fmt == "--":
        return self._combine(code)

    elif fmt == "r-":
        return self._combine(self._reg(args[0]), code)

    elif fmt == "rr":
        return self._combine(
            self._reg(args[1]), self._reg(args[0]), code
        )

    elif fmt == "rv":
        return self._combine(
            self._val(args[1], labels),
            self._reg(args[0]), code
        )

    def _combine(self, *args):
        assert len(args) > 0, "Cannot combine no arguments"
        result = 0
        for a in args:
            result <<= OP_SHIFT
            result |= a
        return result

    def _split(self, lines): #for arrays
    try:
        split = lines.index(self.DIVIDER)
        return lines[0:split], lines[split + 1:]
    except ValueError:
        return lines, []

    def _add_allocations(self, base_of_data, labels, to_allocate): #for arrays
    for alloc in to_allocate:
        fields = [a.strip() for a in alloc.split(":")]
        assert len(fields) == 2, f"Invalid allocation directive '{alloc}'"
        lbl, num_words_text = fields
        assert lbl not in labels, f"Duplicate label '{lbl}' in allocation"
        num_words = int(num_words_text)
        assert (base_of_data + num_words) < RAM_LEN, \
            f"Allocation '{lbl}' requires too much memory"
        labels[lbl] = base_of_data
        base_of_data += num_words

```
```
# Count up to 3.
# - R0: loop index.
# - R1: loop limit.
# - R2: array index.
# - R3: temporary.
ldc R0 0
ldc R1 3
ldc R2 @array
loop:
str R0 R2
ldc R3 1
add R0 R3
add R2 R3
cpy R3 R1
sub R3 R0
bne R3 @loop
hlt
.data
array: 10

----------

R000000 = 000003
R000001 = 000003
R000002 = 00000e
R000003 = 000000
000000:   000002  030102  0b0202  020005
000004:   010302  030006  030206  010304
000008:   000307  030309  000001  000000
00000c:   000001  000002  000000  000000
```
.


### Further notes

## FUrther Topics

### PY Addons

#### Comprehensions
Comprehensions provide a concise way to create lists and dictionaries. They are generally faster and more readable than traditional for loops.

**List Comprehensions**

Structure: `[expression for item in iterable if condition]`
- Simple: `[x**2 for x in range(5)] → [0, 1, 4, 9, 16]`
- With Condition: `[x for x in range(10) if x % 2 == 0]`
- Nested Logic (Ternary): `[x if x > 0 else 0 for x in data]`

**Dictionary Comprehensions**

Structure: `{key_expression: value_expression for item in iterable if condition}`

- Simple: `{x: x**2 for x in range(3)} → {0: 0, 1: 1, 2: 4}`
- From Lists: `{k: v for k, v in zip(keys, values)}`
- Filtering: `{k: v for k, v in my_dict.items() if v > 10}`


#### File I/O (Read and Write)
Python uses file objects to interact with files on the disk.
The with open(...) as ... Pattern

The Context Manager (with statement) is the best practice because it automatically closes the file, even if an exception occurs.
```py3
# Writing
with open("test.txt", "w", encoding="utf-8") as f:
    f.write("Hello World\n")

# Reading
with open("test.txt", "r") as f:
    content = f.read()  # Reads whole file
    # OR: lines = f.readlines() (returns a list)
```

**Modes:**
- 'r': Read (default).
- 'w': Write (overwrites existing).
- 'a': Append (adds to the end).
- 'b': Binary mode (e.g., 'rb' for images).

#### Standard Streams: stdin / stdout

These are handled by the sys module and behave like file objects.
-    sys.stdin: Used for reading input (similar to input(), but more flexible for large data/pipes).
-    sys.stdout: Used for output (what print() uses internally).
-    sys.stderr: Used for error messages.

```py3
import sys

# Reading from pipe/stdin
data = sys.stdin.read()

# Writing to stdout
sys.stdout.write("Direct output\n")
```

#### IO Class
The io module provides the main facilities for dealing with various types of I/O. The most common use case in exams is io.StringIO, which allows you to treat a string as a file.

- io.StringIO: In-memory text stream. Useful for testing functions that expect a file object.
- io.BytesIO: In-memory binary stream.

```py3
import io

# Treat a string like a file
fake_file = io.StringIO("First line\nSecond line")
print(fake_file.read())
```


### Debugging

#### A Debugger

>**A debugger is a separate software process that runs alongside the program being debugged, acting as a controller that can start, stop, and inspect that program.**
>**It often opens a local communication channel (such as a socket or port) so an editor can talk to it.**
>
> It needs:
> - A Python VE (to work on project's dependencies, not system-wide)
> - The Python extension for VS Code for language support
> - debugpy debugger: component that actually enables debugging features (breakpoints, stepping, inspection)

Most modern debuggers (like the one in VS Code) act as a "client" that communicates with a Debug Engine or Interpreter (the "server") using a protocol like the Debug Adapter Protocol (DAP).

-    **Instruction Control**: The debugger hooks into the Python interpreter to pause execution at specific lines.

-    **Breakpoints**: You mark a line where execution should stop. The debugger monitors the "Program Counter" and halts the CPU/Interpreter when it hits that address.

-    **Introspection**: While paused, the debugger reads the current Stack Frame. This allows it to show you the current value of variables in memory.

- **Profiler**: is a tool that analyzes a running program to measure where time and resources are spent. It observes the execution of the code and reports which functions run, how long they take, and whether the program is waiting on things like I/O or CPU work. Profilers help identify performance bottlenecks, so developers can understand why a program is slow and where to optimize.

- **distributed tracing system:** tracks a request as it flows through multiple microservices, assigning it a unique trace ID and collecting timing information from each service it touches. This makes it possible to see the complete execution path end-to-end, including which services were called, how long each step took, and where bottlenecks or failures occurred. --> Open Telemetry e.g.

- Stepping:

    - **Step Over:** Execute the next line without entering functions.

    - **Step Into:** Enter a function to see what happens inside.

    - **Step Out:** Finish the current function and return to the caller.


#### Bug types 

>**A bug is an error, flaw, or unintended behavior in software that causes it to produce an incorrect or unexpected result, or to behave in unintended ways**

**Syntax Bug**	
- Code that breaks the language rules (e.g., missing : or )).	
- Found at: Compile-time (before the code even runs) by the compiler or interpreter or IDE. The interpreter will refuse to start.
**Runtime Bug** (Exceptions)	
- Valid syntax that crashes during execution (e.g., ZeroDivisionError, IndexError).	
- Found at: Execution time with a **Debugger**. The program stops and provides a Traceback (error log).
**Logic Bug**	
- The code runs without crashing, but the output is wrong (e.g., using + instead of *).	
- Found at: Testing or Production. These are the hardest to find and require Unit Tests or a **Debugger**.
**Performance Bug**
- Code is correct but inefficient or too slow.
**Integration Bug**
- Components or systems fail to work together properly
- Found with a debugger



### Call Stack
The Call Stack is a specialized data structure that follows the Last-In, First-Out (LIFO) principle to track active function calls in a program. In Python, every time a function is invoked, the interpreter creates a new Frame Object containing the function’s local variables, arguments, and the return address. These frames are "pushed" onto the stack as functions call each other and "popped" off once the function execution completes, returning control to the caller. When an error occurs, Python prints a Traceback, which is essentially a visual representation of the current state of the call stack at the moment of the crash.

- Stack Overflow: Occurs when too many frames are pushed onto the stack (commonly via infinite recursion), exceeding the memory limit.

- Frame Objects: Unlike some lower-level languages, Python frames are stored on the heap but managed logically as a stack, allowing them to persist in some cases (like in generators or tracebacks).

### Java

#### Differences
- Private: default 
- Protected: For Subclass Access
- Public: Very Rare

1. Static vs. Dynamic Typing
- Python: Dynamically typed – types are determined at runtime.
- Java: Statically typed – types are set at compile-time and cannot change.
2. Variable Declaration and Type Safety
- Python: Variable types are inferred; no explicit type declaration.
- Java: All variables require explicit type declaration, ensuring type safety.
3. Syntax
- Python: Relies on indentation to define code blocks.
- Java: Uses braces {} for code blocks, thus making indentation optional but recommended.
4. Object-Oriented Programming (OOP)
- Python: Fully object-oriented but more flexible – no need for strict OOP adherence.
- Java: Strongly enforces OOP principles; (almost) everything is an object and OOP concepts like inheritance and polymorphism areheavily used.
5. Compilation vs. Interpretation
- Python: Interpreted (kind of) – runs line-by-line, making it flexible but generally slower.
- Java: Compiled to bytecode and run (possibly via interpretation) on the Java Virtual Machine (JVM), making it more performant and platform-independent.
6. Error Handling
- Python: All errors and exceptions are caught at runtime, giving flexibility but potentially hiding issues until they occur.
- Java: Has compile-time checking as well as both checked exceptions (must be handled or declared) and unchecked exceptions, allowing more control over error management at compile-time.

#### Compilation
**Python**
1. Source Code to Bytecode
- When we run a Python script, the Python interpreter compiles the source code (.py file) to bytecode. This bytecode is a low-level, platform-independent representation, which is not directly machine code but an intermediate form that the Python Virtual Machine (PVM) can execute.
- This Python bytecode is saved as .pyc files in the __pycache__ directory for future runs.
2. Interpreting Bytecode
-  Python’s bytecode is interpreted by the Python Virtual Machine (PVM) line-by-line at runtime. This is why we refer to Python as
“interpreted” in practice, even though a compilation step occurs


**Java**

1. Source Code to Bytecode
- When we compile Java (javac), the compiler generates an intermediate binary form of your code that can be interpreted and run by
the Java Virtual Machine (JVM) of your choice.
- This Java bytecode is saved as .class files.
- NOTE: It is technically possible to compile Java to native code ahead-of-time and run the resulting binary.
2. Initial Interpretation
- When you first run Java code, the JVM interprets the bytecode. This initial interpretation enables quick startup.
3. Just-In-Time (JIT) Compilation
- As the program runs, the JVM monitors the bytecode and identifies frequently used code paths, called “hot spots.”
- For these hot spots, the JVM dynamically compiles bytecode into machine-specific code (native machine code) and caches it.
- This JIT-compiled native code is then executed directly by the CPU, bypassing the need for further interpretation for those.

#### Code for File Archiver
```java
// --------------------- Main
public class Main {

    public static void main(String[] args) {
        if (args.length != 2) {
            System.out.println("Usage: java Main source_dir backup_dir");
            System.exit(1);
        }
        String source_dir = args[0];
        String backup_dir = args[1];
        Archiver a = new LocalArchiver(source_dir,backup_dir);
        a.backup();

    }

}
// --------------------- Archiver
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

public abstract class Archiver {
    protected Path sourcePath;

    public Archiver(String sourceDir) {
        this.sourcePath = Paths.get(sourceDir);
    }

    public List<FileEntry> backup() {
        //FileHasher fh = new FileHasher();
        //List<FileEntry> manifest = fh.hashAll(this.sourcePath);
        List<FileEntry> manifest = FileHasher.hashAll(this.sourcePath);
        writeManifest(manifest);
        copyFiles(manifest);
        return manifest;
    }

    abstract protected void writeManifest(List<FileEntry> manifest);

    abstract protected void copyFiles(List<FileEntry> manifest);

}

// --------------------- Local Archiver

import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

public class LocalArchiver extends Archiver {

    private Path backupDir;

    public LocalArchiver(String sourceDir, String backupDir) {
        super(sourceDir);
        this.backupDir = Paths.get(backupDir);
    }

    // f(x) -> x^2
    // f(2) -> 4

    private String getTimestamp() {
        return "test";
    }


    @Override
    protected void writeManifest(List<FileEntry> manifest) {
        if (!Files.exists(this.backupDir)) {
            try {
                Files.createDirectories(this.backupDir);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        Path manifestFile = this.backupDir.resolve(this.getTimestamp() + ".csv");
        try (FileWriter writer = new FileWriter(manifestFile.toFile())) {
            writer.write("filename,hash\n");
            for (FileEntry entry : manifest) {
                writer.write(entry.getFilename() + "," + entry.getHash() + "\n");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    protected void copyFiles(List<FileEntry> manifest) {
        for (FileEntry entry : manifest) {
            Path sourcePath = this.sourcePath.resolve(entry.getFilename());
            Path backupPath = this.backupDir.resolve(entry.getHash() + ".bck");
            if (!Files.exists(backupPath)) {
                try {
                    Files.copy(sourcePath, backupPath);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    

    public void mySpecialLocalMethod(int number) {
        // does very very fancy stuff
    }
    
}


// --------------------- File Hasher

import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class FileHasher {

    private static final int HASH_LEN = 16;

    public static List<FileEntry> hashAll(Path root) {
        List<FileEntry> result = new ArrayList<FileEntry>();

        try (Stream<Path> paths = Files.walk(root)) {
            List<Path> files = paths.filter(Files::isRegularFile)
                                     .collect(Collectors.toList());

            // for file in files:
            for (Path file : files) {
                String relativePath = root.relativize(file).toString();
                String hashCode = FileHasher.calculateHash(file);
                result.add(new FileEntry(relativePath, hashCode));
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        return result;
    }

    private static String calculateHash(Path file) {
        try {
            byte[] data = Files.readAllBytes(file);  // Read entire file at once
            MessageDigest digest = MessageDigest.getInstance("SHA-256");
            byte[] hashBytes = digest.digest(data);

            // Convert hashBytes to hex with desired truncation length
            return FileHasher.bytesToHex(hashBytes, HASH_LEN);

        } catch (IOException | NoSuchAlgorithmException e) {
            e.printStackTrace();
            return null;
        }
    }

    private static String bytesToHex(byte[] bytes, int length) {
        StringBuilder hexString = new StringBuilder();
        for (int i = 0; i < bytes.length && hexString.length() < length; i++) {
            String hex = Integer.toHexString(0xff & bytes[i]);
            if (hex.length() == 1) hexString.append('0');  // Pad with leading zero if needed
            hexString.append(hex);
        }
        return hexString.substring(0, Math.min(length, hexString.length()));
    }
}

// --------------------- File Entry

public class FileEntry {
    private String filename;
    private String hash;

    public FileEntry(String filename, String hash) {
        this.filename = filename;
        this.hash = hash;
    }

    public String getFilename() {
        return this.filename;
    }

    public String getHash() {
        return this.hash;
    }
    
}

```
#### Code for VM
```java

// --------------------- Virtual Machine
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.Reader;
import java.io.Writer;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Scanner;

public class VirtualMachine {
    private String prompt = ">>";
    private int instructionPointer = 0;
    private int[] registers = new int[Architecture.NUM_REG];
    private int[] ram = new int[Architecture.RAM_LEN];
    private final int COLUMNS = 4;

    public VirtualMachine() {
        this.initialize(new int[0]);
    }

    public void initialize(int[] program) {
        if (program.length > ram.length) {
            throw new IllegalArgumentException("Program too long");
        }
        System.arraycopy(program, 0, this.ram, 0, program.length);
        if (program.length < this.ram.length) {
            Arrays.fill(this.ram, program.length, this.ram.length, 0);
        }
        this.instructionPointer = 0;
        Arrays.fill(this.registers, 0);
    }

    public void show(Writer writer) throws IOException {
        // Show registers
        for (int i = 0; i < this.registers.length; i++) {
          writer.write(String.format(
              "R%d: %06X%n", i, this.registers[i]));
        }
        writer.write("\n");
    
        // How much memory to show
        int top = 0;
        for (int i = 0; i < ram.length; i++) {
          if (ram[i] != 0) {
            top = i;
          }
        }
    
        // Show memory
        int base = 0;
        while (base <= top) {
          StringBuilder output = new StringBuilder(String.format("%06X: ", base));
          for (int i = 0; i < COLUMNS; i++) {
            output.append(String.format("  %06X", this.ram[base + i]));
          }
          writer.write(output.toString() + "\n");
          base += COLUMNS;
        }
    
        writer.flush();
      }


    private Instruction fetch() {
        int instruction = this.ram[this.instructionPointer];
        this.instructionPointer++;

        int op = instruction & Architecture.OP_MASK;
        instruction >>= Architecture.OP_SHIFT;
        int arg0 = instruction & Architecture.OP_MASK;
        instruction >>= Architecture.OP_SHIFT;
        int arg1 = instruction & Architecture.OP_MASK;
        Instruction inst = new Instruction(op, arg0, arg1);
        return inst;
        // return new int[] {op, arg0, arg1};
    }
    
    public void run() {
        boolean running = true;
        while (running) {
            Instruction inst = this.fetch();

            if (inst.opcode == Architecture.OPS.get("hlt").code) {
                running = false; 
            } else if (inst.opcode == Architecture.OPS.get("prr").code) {
                System.out.println(this.prompt + " " + this.registers[inst.arg0]);
            } else {
                throw new IllegalStateException("Unknown opcode: " + inst.opcode);
            }

        }
    }

    public static void main(String[] args) throws IOException {
        if (args.length != 2) {
            System.err.println("Usage: java VirtualMachine input|- output|-");
            System.exit(1);
        }

        Reader reader;
        if (args[0].equals("-")) {
            reader = new InputStreamReader(System.in);
        } else {
            reader = new FileReader(args[0]);
        }

        Writer writer;
        if (args[1].equals("-")) {
            writer = new OutputStreamWriter(System.out);
        } else {
            writer = new FileWriter(args[1]);
        }

        int[] program;
        try ( Scanner scanner = new Scanner(reader)) {
            List<Integer> programList = new ArrayList<>();
            while (scanner.hasNext()) {
                String line = scanner.nextLine().trim();
                if (!line.isEmpty()) {
                    programList.add(Integer.parseInt(line, 16));
                }
            }
            program = new int[programList.size()];
            for (int i = 0; i < programList.size(); i++) {
                program[i] = programList.get(i);
            }
        }

        VirtualMachine vm = new VirtualMachine();
        vm.initialize(program);
        vm.show(writer);
        vm.run();
        vm.show(writer);

    }

}

// --------------------- Architecture
import java.util.HashMap;
import java.util.Map;

public class Architecture {
    public static final int NUM_REG = 4;
    public static final int RAM_LEN = 256;

    public static final int OP_MASK = 0xFF;   // Select a single byte
    public static final int OP_SHIFT = 8;    // Shift up by one byte
    public static final int OP_WIDTH = 6;    // Op width in characters when printing

    public static final Map<String,Operation> OPS = new HashMap<>();

    static {
        OPS.put("hlt", new Operation(0x1, "--"));
        OPS.put("ldc", new Operation(0x2, "rv"));
        OPS.put("ldr", new Operation(0x3, "rr"));
        OPS.put("cpy", new Operation(0x4, "rr"));
        OPS.put("str", new Operation(0x5, "rr"));
        OPS.put("add", new Operation(0x6, "rr"));
        OPS.put("sub", new Operation(0x7, "rr"));
        OPS.put("beq", new Operation(0x8, "rv"));
        OPS.put("bne", new Operation(0x9, "rv"));
        OPS.put("prr", new Operation(0xA, "r-"));
        OPS.put("prm", new Operation(0xB, "r-"));
    }
}
// --------------------- Instruction
public class Instruction {
    final int opcode;
    final int arg0;
    final int arg1;

    public Instruction(int opcode, int arg0, int arg1) {
        this.opcode = opcode;
        this.arg0 = arg0;
        this.arg1 = arg1;
    }
}

// --------------------- Operation
public class Operation {
    int code;
    String fmt;

    public Operation(int code, String fmt){
        this.code = code;
        this.fmt = fmt;
    }
}

```

>Please note that this list does not include the following topics, which we covered through slides, in-class explanations and coding, and exercises:
>    - Call Stack
>    - Java
>    - Debugging
>
> These topics can also be part of the exam.

