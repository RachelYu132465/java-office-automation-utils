# EEQ — Excel Equation / Goal Seek Automation

EEQ is a Java/Maven project for inspecting Excel workbooks, identifying formula cells and their input cells, applying specification limits, running Excel Goal Seek, and generating Word reports.

The project is organized so that you can either run the complete workflow or reuse individual Java utility classes in another project.

> **Project status:** This repository is an older, work-in-progress implementation. The workflow depends on workbook formatting and structure, Microsoft Excel, and Windows Script Host. Review the configuration and test with copies of your files before using it on important workbooks.

## What it does

For a supported Excel workbook, the intended end-to-end process is:

1. Open the workbook and inspect each worksheet.
2. Find formula cells marked with blue font.
3. Identify the formula's referenced input cells.
4. Find the specification associated with each formula, or allow the user to enter a specification manually.
5. Store the formula, input cells, output cell, and specification in a ValidGoal data object.
6. Generate VBScript files that use Microsoft Excel's Goal Seek feature to calculate input values that meet the target.
7. Run the scripts and save result workbooks.
8. Read the result workbooks and create Word reports for the original analysis and generated test cases.

Apache POI can read and write Excel files, but it does not provide Excel's Goal Seek feature. This project therefore delegates that calculation to desktop Microsoft Excel through VBScript.

## Quick start

### Requirements

- Java 11 (the Maven compiler plugin is configured for source/target 11).
- Apache Maven.
- Windows.
- Desktop Microsoft Excel installed and available to the Windows Script Host (cscript).
- Excel workbooks that follow the conventions described below.

The project includes Apache POI and other dependencies in pom.xml. Maven downloads these dependencies during the build.

### 1. Clone the repository

~~~bash
git clone https://github.com/RachelYu132465/EEQ.git
cd EEQ
~~~

### 2. Build the project

~~~bash
mvn clean package
~~~

### 3. Configure the input workbook

The current runnable example is:

src/main/java/test/ExcelForRu.java

Before running it:

1. Put your input workbook in the current user's home directory (the code uses FileSystemView.getFileSystemView().getHomeDirectory()).
2. Open ExcelForRu.java and change the output filename to the workbook you want to process.
3. Check the output folders and worksheet names expected by your workbook.
4. Make sure the workbook's formulas, font colors, and specification labels match the conventions below.

The current example uses a hard-coded sample filename. It is not yet a general command-line application with configurable input and output arguments.

### 4. Run the workflow

Run test.ExcelForRu from your IDE, or configure your IDE to use the Maven project classpath and run this class:

~~~text
test.ExcelForRu
~~~

The program writes Word documents and intermediate Excel/VBS files under the user's home directory. It may modify and save the input workbook. **Use a backup copy.**

## Workbook conventions and assumptions

The extraction logic is not fully generic. The current implementation assumes:

- Relevant output cells contain Excel formulas and use a blue font. The code scans for blue formula cells.
- A specification can be found by searching nearby cells for the word specification, or can be entered manually when prompted.
- Formula references can be parsed by the project's formula-processing logic.
- The workbook has worksheet names that can be used to locate corresponding sheets in generated result files.
- The worksheet named History of versions is skipped by the main workflow.
- File names and paths used by the generated VBScript must be valid Windows paths. The original source comments also warn against spaces in input file names; test your own paths before batch processing.

These are implementation assumptions, not a formal workbook specification. If your workbook uses different colors, labels, formula patterns, or sheet layouts, the extraction logic may need adjustment.

## Project structure

Main source files are under src/main/java.

| Package / file | Responsibility |
| --- | --- |
| test/ExcelForRu.java | Main workflow runner / orchestration example |
| mainFlow/extract.java | Finds formula cells and extracts input/output/specification information |
| mainFlow/VBS.java | Builds and executes VBScript files for Excel Goal Seek |
| mainFlow/ProduceWordFile.java | Coordinates Word report generation |
| dataStructure/ValidGoal.java | Data model for one validation goal |
| msexcel/Excel.java, ExcelCell.java, Range.java | Excel workbook, cell, and range helpers |
| validate/ | Formula parsing, input-cell discovery, specification/range parsing, and validation helpers |
| msword/ | Word document text, table, and style helpers |
| common/ | General file, string, and number utilities |
| copyExcel/ | Excel row/sheet copying and conversion helpers |
| customExcel/ | Custom Excel-related helpers |

## Main workflow: how the classes fit together

The main entry point is **test.ExcelForRu**. It coordinates the other classes; it is not the place where all the Excel, formula, VBScript, and Word logic is implemented.

### 1. test.ExcelForRu — workflow orchestration

This class:

- Opens the configured workbook.
- Loops through worksheets.
- Calls extract.extractData(...) to identify goals.
- Applies conditional formatting and prepares Word report content.
- Calls VBS.storeTargetInFile(...) and VBS.produceVBSFiles(...).
- Executes generated scripts with VBS.execAllVBSFiles(...).
- Reads the generated workbooks using getValidatedValues(...).
- Calls ProduceWordFile to create general and test-case reports.

Use this class when you want to run the complete process. To adapt the project to a different workbook, start by changing its input filename and then check the workbook-specific assumptions in the extraction and reporting classes.

### 2. mainFlow.extract — identify validation goals

The main method is:

~~~java
HashMap<String, ValidGoal> goals = extract.extractData(excel);
~~~

It scans the selected worksheet for blue formula cells, searches for the associated specification, optionally asks for a manually entered specification, and collects the formula's input cells.

The returned map is keyed by the output cell address. Each ValidGoal stores the output cell, an input cell, all discovered input cells, worksheet title, and parsed range/specification.

You can reuse this extraction step without running VBScript or generating Word files, provided the workbook follows the expected conventions.

### 3. mainFlow.VBS — generate and run Goal Seek scripts

This class separates the Goal Seek work from formula extraction:

- storeTargetInFile(...) stores target values in the workbook so that Excel's Goal Seek can reference them.
- produceVBSFiles(...) creates VBScript files for the goals.
- execVBSFile(...) and execAllVBSFiles(...) run scripts through cscript.

This part requires Windows and desktop Microsoft Excel. It is not a cross-platform replacement for Goal Seek. The current source also contains a hard-coded example path in VBS.main(); configure paths before using that method.

### 4. mainFlow.ProduceWordFile — report generation

This class assembles Word reports from the extracted goal data. It delegates the actual table layouts and styles to the msword package.

You can use the report-generation methods separately if you already have the required ValidGoal data and a Word document object.

## Reusing the Java utility classes independently

You do not have to run test.ExcelForRu to use every class in this repository. Several classes provide reusable operations. They can be called from another Java class, as long as you include the project dependencies and any required supporting classes.

### A. msexcel.Excel — reusable Excel helper

This is the central Excel wrapper around Apache POI. It provides methods for loading or creating workbooks, selecting sheets/rows/cells, reading and writing values, and saving workbooks.

Example:

~~~java
import msexcel.Excel;
import java.io.File;

public class ReadWorkbook {
    public static void main(String[] args) {
        Excel workbook = Excel.loadExcel(new File("input.xlsx"));

        if (workbook == null) {
            throw new IllegalArgumentException("Could not open input.xlsx");
        }

        workbook.assignSheet(0);

        System.out.println("Sheet: " + workbook.getSheet().getSheetName());
        System.out.println("Sheets: " + workbook.getNumberOfSheets());
    }
}
~~~

This example uses only the Excel helper; it does not run the EEQ Goal Seek or Word-report workflow.

### B. common.StringProcessor — reusable string operations

Provides independent static methods for tasks such as:

- Extracting strings with regular expressions.
- Sorting strings by embedded numeric values.
- Removing special characters.
- Checking whether text contains specified keywords.
- Matching text against regular expressions.
- Removing an outer function wrapper from a formula string.

Example:

~~~java
import common.StringProcessor;

public class StringExample {
    public static void main(String[] args) {
        System.out.println(
            StringProcessor.ifContain("specification", "Output specification: 95%")
        );
    }
}
~~~

### C. common.NumberProcessor — reusable number operations

Contains static numeric helpers, including finding the index of the largest number and working with decimal precision / powers of ten. These methods can be called directly without running the main workflow.

### D. common.FileHandler — reusable file operations

Provides file-related helpers such as recursive file discovery by keyword, moving files into subfolders, reading text files, writing text files, and saving Word documents.

Review the method before using it on important files: some operations move or overwrite files.

### E. validate — formula and specification utilities

The validate package contains the logic used by the main extraction process, including:

- excelFormulaProcessor: finds specification labels and values, inspects cell font colors, and extracts formula-related text.
- excelFormulaValidator: discovers referenced input cells, scans for manual specification input, and reads generated validation workbooks.
- MyRange, OperatorConvertor, customStringFormatter: represent and process specification ranges and comparison rules.

These classes can be reused independently when their input objects and supporting classes are available. Some methods prompt for console input, and some expect an Excel workbook or Apache POI cell rather than a plain string.

### F. msword — Word document helpers

The msword package contains methods for constructing and styling Word tables and text using Apache POI's XWPF APIs. You can reuse these helpers in another Java project if you provide the required XWPFDocument, goal data, and project dependencies.

## Using these classes in another Maven project

The simplest way to reuse a helper is to copy the required source files and their dependent classes into your project, then include the relevant dependencies from this repository's pom.xml.

For a quick experiment, you can also open this repository as a Maven project in IntelliJ IDEA and create a separate Java class under src/main/java that imports the helper you want to test.

Not every class is standalone: for example, extract depends on the Excel wrapper and validation classes, while Excel itself depends on the Maven libraries declared in this project. Check imports and method dependencies before copying a single file.

## Important limitations and safety notes

- **Back up input workbooks.** The main workflow saves workbook changes and creates intermediate files.
- **Windows-only Goal Seek execution.** The VBScript execution path uses cscript and desktop Microsoft Excel.
- **Paths are currently hard-coded in places.** In particular, the sample runner selects a fixed filename, and the VBScript helper includes a hard-coded sample path in its main method. Update these before running in a different environment.
- **No automated test suite is configured.** The repository contains example/test classes, but they are not a complete regression test suite.
- **Workbook compatibility is limited by assumptions.** Font color, specification labels, formula syntax, sheet names, and file naming all affect results.
- **Review generated output.** Confirm Goal Seek results and Word reports against the source workbook before relying on them.

## Dependencies

The Maven project currently declares:

- Apache POI (poi, poi-ooxml) for Excel and Word document handling.
- Apache XMLBeans.
- GemBox Spreadsheet.
- Apache Commons Lang / Commons Lang 3 and Commons IO.
- JetBrains annotations.

See pom.xml for the exact versions.

## License

No license file is currently provided. Unless a license is added, reuse and redistribution are not automatically granted.
