import Text_Manipulation
import Excel_File_Manipulation

# The batch script runner for the text file and Excel file manipulations
#################### My Testing ############################################
#folder_path = input('\033[1m' + "\nPlease enter the path to the folder where the desired files are stored.\n")  ##
path = r"C:\Users\mnarcisi\Documents\Mike\Scientific Affairs\Data"        ##
                                                                          ##
############################################################################


print("\n\033[1m" + "Current Directory: " + path)
file_decision = 0
while file_decision != 1 and file_decision != 2:
    try:
        file_decision = int(input("\nManipulate a text file or an Excel sheet?\n(1) Text File\n(2) Excel Sheet\n"))
        if file_decision != 1 and file_decision != 2:
            print("\nPlease enter a valid number.")
    except ValueError:
        print("\nPlease enter the number of the desired option.")
        file_decision = 0
if file_decision == 1:
    Text_Manipulation.defib_shock_text_manipulation(path)
else:
    Excel_File_Manipulation.add_title_column(path)
