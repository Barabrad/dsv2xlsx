'''
Brad Barakat
Made for AME-341b

This script will convert every xls file in a given input directory to a xlsx file in a given output directory.
If there are folders within the input directory, those folders will be made in the output directory and filled
with the corresponding xlsx files.
This script (as-is) will not read non-xls files.
'''

# Python has a built-in os library
import os
# If xlsxwriter is not installed, type "pip3 install xlsxwriter" into a Terminal window
import xlsxwriter


# This main() function could be broken up more, but given that this is not an assignment, I can relax the standards a bit.
def main():
    # Get directories from the user
    inRootDir = input("Enter the path to the folder containing the xls files:\n >> ");
    # Example inRootDir: "/Users/brad/Desktop/Courses/AME 341b/HW/Assignment Submissions/SE2/Data/"
    outRootDir = input("Enter the path to the folder that will contain the converted xlsx files:\n >> ");
    # Example outRootDir: "/Users/brad/Desktop/Courses/AME 341b/HW/Assignment Submissions/SE2/Data_XLSX/"
    xlsDelim = "\t"; # Specify delimeter (found empirically)

    # Confirm directory choices with user
    print("");
    print("Input (xls) root directory:\n    " + inRootDir);
    print("Output (xlsx) root directory:\n    " + outRootDir);
    confirm = input("Continue? ('y'/'n'): ").upper() == "Y";
    print("");
    if confirm and not os.path.exists(inRootDir):
        print("The input directory does not exist. Terminating...");
        confirm = False;
    else:
        # If the the output root directory exists, give a warning
        if not os.path.exists(outRootDir):
            os.mkdir(outRootDir);
        else:
            print("Warning: Path '" + outRootDir + "' already exists.");
            print("Any xlsx file in it with a corresponding xls file will be overwritten.");
            confirm = input("Continue? ('y'/'n'): ").upper() == "Y";
            # Make sure the paths are not the same, otherwise there may be an infinite loop
            if confirm and os.path.samefile(inRootDir, outRootDir):
                print("The output directory cannot be the same as the input directory.");
                print("Exiting...");
                confirm = False;
            print("");

    # Begin conversions
    if confirm:
        # Get length of input root directory string
        len_ird = len(inRootDir);
        # Begin iterating through the directory tree
        for (root, dirs, files) in os.walk(inRootDir):
            outputDir = outRootDir + root[len_ird:];
            for filename in files:
                # Make sure the file is not hidden AND is a xls file
                if filename[0] != "." and filename[-4:] == ".xls":
                    # Put the second "x" in "xlsx"
                    filename_x = filename + "x";
                    # Determine the current full input and output file names
                    full_name = os.path.normpath(root + "/" + filename);
                    print("Current file: " + full_name);
                    outputFile = os.path.normpath(outputDir + "/" + filename_x);
                    # Make sure the folder exists
                    if not os.path.exists(outputDir):
                        os.mkdir(outputDir);
                    # Turn the tab-delimited xls into an xlsx
                    with open(full_name) as inFile:
                        workbook = xlsxwriter.Workbook(outputFile);
                        sheetName = filename_x[:-5];
                        sheet = workbook.add_worksheet(sheetName);
                        rowNum = 0;
                        for line in inFile:
                            row = line.split(xlsDelim);
                            for col in range(len(row)):
                                cellData = row[col];
                                try:
                                    cellData = float(cellData);
                                except:
                                    pass;
                                sheet.write(rowNum, col, cellData);
                            rowNum += 1;
                        workbook.close();
    # Print a confirmation
    print("\nDone.");


# Run main()
main();
