# Mathcad
## Application Overview
### Introduction
Mathcad is a computational notebook software available on Windows, known for its graphical what-you-see-is-what-you-get (WYSIWYG) editor and integrated engineering units system.

See the [Official Documentation for Mathcad Prime 8](https://support.ptc.com/help/mathcad/r8.0/en/index.html).

Mathcad has been around since 1986, and in 2011 the code base saw a major rewrite, switching from “Mathcad” to “PTC Mathcad Prime.” See [Differences Between PTC Mathcad 15 and PTC Mathcad Prime](https://support.ptc.com/help/wnc/r12.0.0.0/en/index.html#page/Windchill_Help_Center/WWGMMathcadUseMathcadPrime.html). The conversion is still hotly contested from longtime Mathcad users due to many features still missing in Prime.

### Features
Nonetheless, Mathcad is an awesome tool I have come to enjoy. It produces printable, human-readable documents that are easy to distribute for review, all formatted in typical mathematical notation. Units are handled brilliantly without any special syntax, and Mathcad can even be embedded inside Creo documents for manipulating dimensions and properties. 

### Limitations
Mathcad lacks in cross-platform functionality, and locks you into proprietary Mathcad binary files. Mathcad Express, a free but limited version of Mathcad, still allows you to open Prime files without calculating blocks with paid functions. I created an Excel document that allows you to iterate a table of values with units through a single Mathcad worksheet. See [[#Excel Iterator]].

## API
### Introduction
Mathcad is lackluster at iterating and abstracting worksheets. Attempting to turn every expression into a function misses out on the benefits of Mathcad’s graphical representation. I recommend creating a worksheet that does a single (and possibly complex) calculation well for a single case, and then using another scripting language to change variables and run iterative studies.

The Mathcad API is a [Component Object Model](https://support.ptc.com/help/mathcad/r7.0/en/index.html#page/PTC_Mathcad_Help%2Fcomponent_object_model.html%23) (COM) interface for interacting with Mathcad using a library of different commands. The API works with PTC Mathcad Prime 3.1 and later. COM is a standard interface for automating Windows applications. See the [Official Documentation](https://support.ptc.com/help/mathcad/r8.0/en/index.html#page/PTC_Mathcad_Help/API/mathcad_and_automation_api.html#).

The gist of using the API includes instantiating instances of the Mathcad application or worksheet classes as objects, then calling methods on that object to manipulate data. 

The API supports JavaScript, C++, C#, VB, VB Script, VBA, and any other language with a COM library. The Python wrapper [MathcadPy](https://pypi.org/project/MathcadPy/) exists but is unmaintained. VBA is especially useful as it is pre-installed on any Windows machine that has MS Office, and works directly inside Excel for easily passing tabulated data through a Mathcad worksheet. 

### Library
The type library file `Ptc.MathcadPrime.Automation.tlb` lives in the root directory of your Mathcad installation. It can be opened and explored inside [Visual Studio]([https://visualstudio.microsoft.com](https://visualstudio.microsoft.com/)) using the **Object Browser** to see which methods and classes exist and how to use them. I do not believe that VS Code has the ability to open this library out of the box, but please reach out if I am wrong. 

#### Visual Studio setup
The *.NET desktop development* workload must be installed with Visual Studio. Open Visual Studio Installer, and modify your installation of Visual Studio. Under *workloads*, enable **.NET desktop development**.

### Designate inputs and outputs
Variables must be flagged as either an input or an output if you want it to be exposed to the API. Part of this process is also specifying an alias for that variable that is machine readable (Greek characters and underscores are supported). See [To Designate Input and Output Regions (ptc.com)](https://support.ptc.com/help/mathcad/r7.0/en/index.html#page/PTC_Mathcad_Help%2Fto_designate_input_and_output_regions.html%23) for help on managing I/O in a Mathcad worksheet.

#### Unique variable names are not enforced by Mathcad
Mathcad does not check for unique values across inputs and outputs. This could be problematic depending on how you interact with the worksheet. I highly recommend using unique names for all I/O aliases.

### Units
Units shall be in string format. Use asterisks `*` and forward-slashes `/` for chaining units. The units are not well-documented, but are believed to follow the same input notation at Mathcad e.g. pound-force as `lbf`.

### VBA Reference
Manipulating a [[mathcad]] worksheet with Excel is easier than it sounds. VBA natively supports communication with the Mathcad API with any additional installations or libraries. You only need to have Mathcad installed on the machine. 

#### Initialize Mathcad and open worksheet
```vb
Dim Mathcad As Object
Dim MathcadWorksheet As Object
Set Mathcad = CreateObject("MathcadPrime.Application")
Mathcad.Visible = False ' Hide Mathcad GUI
Set MathcadWorksheet = Mathcad.Open(filePath)
```

#### Get inputs from worksheet
```vb
Dim MathcadInputs: Set MathcadInputs = MathcadWorksheet.Inputs
Dim countMathcadInputs As Integer
countMathcadInputs = MathcadInputs.Count
If countMathcadInputs = 0 Then
	MsgBox ("Error: This worksheet has no inputs.")
	Exit Sub
End If
```

#### Get outputs from worksheet
```vb
Dim MathcadOutputs: Set MathcadOutputs = MathcadWorksheet.Outputs
Dim countMathcadOutputs As Integer
countMathcadOutputs = MathcadOutputs.Count
If countMathcadOutputs = 0 Then
	MsgBox ("Error: This worksheet has no outputs.")
	Exit Sub
End If
```