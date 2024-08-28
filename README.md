# dsv2xlsx
This script provides a quick way to turn a set of delimiter-separated value (DSV) files into .xslx files (the originals are preserved).

## Documentation

### Introduction
This script will convert every DSV file in a given input directory to a xlsx file in a given output directory, and it works on both Mac and Windows. If there are folders within the input directory, those folders will be made in the output directory and filled with the corresponding xlsx files. This script (as-is) will not read or copy non-DSV files. Although there are comments in the code, I figured a document with a tutorial and warnings would be better. In this document, "terminal window" (for Mac) will mean "command prompt" for Windows.

### Libraries
The libraries this script uses are listed below, as well as the download instructions. **My assumption is that you already have Python 3 installed on your computer.** To check, open a terminal window and type `python3 -V`. If the output does not display a version number, try `python -V`. If the latter command works, use `python` and `pip` instead of `python3` and `pip3`, respectively. If neither command shows a version number, install Python 3 first, and then return here. To see which non-built-in libraries are already installed, open a terminal window and type `pip3 list`.
1. os
    * This library is built-in, so you should not need to install anything.
2. xlsxwriter
    * This library is not built-in, so you need to open a terminal window and enter `pip3 install xlsxwriter` if you do not have the library.

### Warning
In Spring 2023, my semester of AME-341b, the xls files from the wind tunnel lab were opening in TextEdit as **tab-delimited** values. My assumption when maintaining this script is that the files will continue to be delimiter-separated. If the values are not displaying properly when the first line is displayed, maybe the format changed. In that case, this file may no longer be of use for the wind tunnel lab.

Also, I am using a Mac, so the path slashes in the tutorial are different from those for Windows: "/" versus "\\" (the script accounts for this difference in input, but the tutorial does not).
* Note that Python's `os.path.normpath()` will correct "/" to "\\" on Windows, but will not correct "\\" to "/" on Mac.

### Example Folder Organization
Note that the ellipses in the folder paths below are meant to indicate that there could be a longer path. Suppose you have your wind tunnel xls files in folders that are all in one folder called XLS_Data:

.../XLS_Data
* /WT_Monday
* /WT_Tuesday
* /WT_Wednesday
* /WT_Thursday
* (etc.)

If you enter the XLS_Data folder path as the input and enter a new path for the output (for example, .../XLSX_Data), the following directory will be created, with every xls file in each folder (including the main one) converted to an xlsx file:

.../XLSX_Data
* /WT_Monday
* /WT_Tuesday
* /WT_Wednesday
* /WT_Thursday
* (etc.)

Note: there can be other non-DSV files in the folders, but they will not be read or copied.

### Tutorial
If all of the libraries are installed, and the DSV files are all in one folder (no matter how many subfolders there are), you are ready for the tutorial.
1. Run the script. **You can do this either from your IDE or a terminal window.** For this tutorial, I will use the terminal window (I used the `cd` command to get to the directory with the code). Beyond this step, the process is the same whether you use an IDE or terminal window.
```
brad@Brads-MBP ~ % cd "/Users/brad/Desktop/Courses/AME 341b/HW/Assignment Submissions/SE2/Python_XLS_Conv"
brad@Brads-MBP Python_XLS_Conv % python3 dsv2xlsx.py
Enter the delimiter-separated value (DSV) file extension (without the '.')
 >>
```

2. Enter the file extension ("xls" in this case). Afterwards, it will ask for the path to the existing DSV folder, then for the path to the xlsx folder. Note the following:
    * The xlsx folder does not have to be currently existing; the code can make new folders. If the xlsx folder does exist, it will give you a warning and ask you to confirm.
    * The DSV and xlsx folders **cannot** be the same.
    * The xlsx folder **cannot** be inside the DSV folder.
    * **Do not use quotation marks.** The script evaluates these inputs as strings, so it does not matter if there are spaces in the file paths.
```
Enter the delimiter-separated value (DSV) file extension (without the '.')
 >> xls
Enter file path to the folder containing the xls files:
 >> /Users/brad/Desktop/Courses/AME 341b/HW/Assignment Submissions/SE2/Python_XLS_Conv/XLS_Folder
Enter file path to the folder that will contain the converted xlsx files:
 >> /Users/brad/Desktop/Courses/AME 341b/HW/Assignment Submissions/SE2/Python_XLS_Conv/XLSX_Folder

Input (xls) root directory:
    /Users/brad/Desktop/Courses/AME 341b/HW/Assignment Submissions/SE2/Python_XLS_Conv/XLS_Folder
Output (xlsx) root directory:
    /Users/brad/Desktop/Courses/AME 341b/HW/Assignment Submissions/SE2/Python_XLS_Conv/XLSX_Folder
Continue? ('y'/'n'):
```

3. If the input and output directories are correct, continue. It will then ask you to choose the appropriate delimiter (tab in this case, assuming the xls file format has not been changed). Note that the single quotes are generated by the code.
```
First line:
  'U0[m/s]	U0_Sd	U[m/s ]	U_Sd	L[N ]	D[N]	Horiz.	Verti.	3/20/2023 5:05 PM'
Choose the delimiter of the above line, or -1 to exit:
0: '	'
1: ','
2: Other
Choice:
 >> 0
```

4. Finally, the code will begin going through the input directory to find files that have the specified extension. Note that the ellipsis in the output below is meant to show that there could be more files processed, and is **not** generated by the code.
```
Current file: /Users/brad/Desktop/Courses/AME 341b/HW/Assignment Submissions/SE2/Python_XLS_Conv/XLS_Folder/WT_Monday/NACA0010U10Ap05_4.xls
Current file: /Users/brad/Desktop/Courses/AME 341b/HW/Assignment Submissions/SE2/Python_XLS_Conv/XLS_Folder/WT_Monday/NACA0010U10Ap05_3.xls
...
Done.
```

5. Your xlsx folder should now be created (or updated if it already existed).
    * Note that if the output xlsx folder already existed, any file in it without a corresponding DSV file in the input folder will **not** be overwritten or deleted.
