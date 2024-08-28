'''
Brad Barakat
Made for AME-341b

This script will convert every delimiter-separated value (DSV) file in a given input directory to a corresponding
xlsx file in a given output directory.
If there are folders within the input directory, those folders will be made in the output directory and filled
with the corresponding xlsx files.
This script (as-is) will not read files without the user-specified extension in the input directory.
'''

# Python has a built-in os library
import os
# If xlsxwriter is not installed, type "pip3 install xlsxwriter" into a Terminal window
import xlsxwriter


# This function makes all "Continue?" requests consistent
def confirmContinue():
    return input("Continue? ('y'/'n'): ").upper() == "Y";


# This function gets a valid integer input from the user
def getValidIntInput(prompt, lBnd=None, hBnd=None):
    valid = False;
    while (not valid):
        x = input(prompt).strip();
        try:
            x = int(x);
            if (lBnd == None): lBnd = x;
            if (hBnd == None): hBnd = x;
            valid = (lBnd <= x) and (x <= hBnd);
            if (not valid): print(f"Error: Integer out of range [{lBnd},{hBnd}]");
        except:
            print("Error: Numeric input not an integer");
    return x;


# This function checks to make sure the inputted directories are valid
def confirmAndCheckRootDirs(inRootDir, outRootDir, dsvExt):
    # Confirm directory choices with user
    print("");
    print(f"Input ({dsvExt}) root directory:\n    " + inRootDir);
    print("Output (xlsx) root directory:\n    " + outRootDir);
    confirm = confirmContinue();
    print("");
    if (confirm) and (not os.path.exists(inRootDir)):
        print("The input directory does not exist. Terminating...");
        confirm = False;
    else:
        # Make sure the paths are not the same, otherwise there may be an infinite loop
        # Also make sure that the xlsx folder is not in the DSV folder
        if (confirm) and (os.path.commonpath([inRootDir, outRootDir]) == inRootDir):
            print("The output directory cannot be the same as (or inside) the input directory.");
            print("Exiting...");
            confirm = False;
        if (confirm):
            # If the the output root directory exists, give a warning
            if (not os.path.exists(outRootDir)):
                os.makedirs(outRootDir);
            else:
                print("Warning: Path '" + outRootDir + "' already exists.");
                print(f"Any xlsx file in it with a corresponding {dsvExt} file will be overwritten.");
                confirm = confirmContinue();
                print("");
    return confirm;


def getDelim(inRootDir, dsvExt, delimList):
    foundFile = False;
    for (root, dirs, files) in os.walk(inRootDir):
        for filename in files:
            # Make sure the file is not hidden AND is a wanted file
            if (not filename.startswith(".")) and (filename.endswith(dsvExt)):
                # Determine the current full input and output file names
                full_name = os.path.join(root, filename);
                # Read the first line
                with open(full_name, "rt") as f: line1 = f.readline().strip();
                foundFile = True;
                break;
        if (foundFile): break;
    # Print the first line and have the user choose the delimiter
    print(f"First line:\n  '{line1}'");
    print("Choose the delimiter of the above line, or -1 to exit:");
    d_ind = 0;
    for d in delimList:
        print(f"{d_ind}: '{d}'");
        d_ind += 1;
    print(f"{d_ind}: Other"); # At this point, d_ind = len(delimList)
    ind = getValidIntInput("Choice:\n >> ", -1, d_ind);
    # Act according to user input
    if (ind == -1): delim = None;
    elif (ind == d_ind): delim = input("Enter delimiter:\n >> ");
    else: delim = delimList[ind]
    print("");
    return delim;


# This function does the conversions from DSV to xlsx
def convertFiles(inRootDir, outRootDir, dsvExt, delim):
    # Get length of input root directory string
    len_ird = len(inRootDir);
    # Set xlsx extension
    xlsxExt = ".xlsx";
    # Begin iterating through the directory tree
    for (root, dirs, files) in os.walk(inRootDir):
        outputDir = outRootDir + root[len_ird:];
        for filename in files:
            # Make sure the file is not hidden AND is a DSV file
            if (not filename.startswith(".")) and (filename.endswith(dsvExt)):
                # Apply xlsx extension
                filename_base = os.path.splitext(filename)[0];
                filename_x = filename_base + xlsxExt;
                # Determine the current full input and output file names
                full_name = os.path.join(root, filename);
                print("Current file: " + full_name);
                outputFile = os.path.join(outputDir, filename_x);
                # Make sure the folder exists
                if (not os.path.exists(outputDir)):
                    os.makedirs(outputDir);
                # Turn the DSV into an xlsx
                with open(full_name) as inFile:
                    workbook = xlsxwriter.Workbook(outputFile, {'constant_memory': True});
                    sheetName = filename_base;
                    sheet = workbook.add_worksheet(sheetName);
                    rowNum = 0;
                    for line in inFile:
                        row = line.split(delim);
                        for col in range(len(row)):
                            cellData = row[col];
                            try: cellData = float(cellData);
                            except: pass;
                            sheet.write(rowNum, col, cellData);
                        rowNum += 1;
                    workbook.close();

# main()
def main():
    # Get file type
    dsvExt = input("Enter the delimiter-separated value (DSV) file extension (without the '.'):\n >> ").strip().lower().replace(".", "");
    # Get directories from the user
    inRootDir = os.path.normpath(input(f"Enter the path to the folder containing the {dsvExt} files:\n >> "));
    # Example inRootDir: "/Users/brad/Desktop/Courses/AME 341b/HW/Assignment Submissions/SE2/Data/"
    outRootDir = os.path.normpath(input("Enter the path to the folder that will contain the converted xlsx files:\n >> "));
    # Example outRootDir: "/Users/brad/Desktop/Courses/AME 341b/HW/Assignment Submissions/SE2/Data_XLSX/"
    
    # Check user-inputted directories
    confirm = confirmAndCheckRootDirs(inRootDir, outRootDir, dsvExt);
    # Begin conversions
    if (confirm):
        # Get delimiter
        dsvExt = "." + dsvExt;
        delimList = ["\t", ","];
        delim = getDelim(inRootDir, dsvExt, delimList);
        if (isinstance(delim, str)): convertFiles(inRootDir, outRootDir, dsvExt, delim);
    # Print a confirmation
    print("\nDone.");


# Run main()
main();
