import Text_Manipulation
import Excel_File_Manipulation

# The batch script runner for the text file and Excel file manipulations
#################### My Testing ############################################
#folder_path = input('\033[1m' + "\nPlease enter the path to the folder where the desired files are stored.\n")  ##
path = r"C:\Users\mnarcisi\Documents\Mike\Scientific Affairs\Data_Group_2"        ##
                                                                          ##
############################################################################


print("\n\033[1m" + "Current Directory: " + path)
file_decision = 0
while file_decision != 1 and file_decision != 2:
    try:
        file_decision = int(input("\nManipulate text files or Excel sheets?\n(1) Text File\n(2) Excel Sheet\n"))
        if file_decision != 1 and file_decision != 2:
            print("\nPlease enter a valid number.")
    except ValueError:
        print("\nPlease enter the number of the desired option.")
        file_decision = 0
if file_decision == 1:
    data_extraction_file, case_file = Text_Manipulation.defib_shock_data_consolidation(path)
    if data_extraction_file != "N/A":
        data_csv_file = Text_Manipulation.create_csv(data_extraction_file, "Element Name", "Time (sec)", "File Name")
        case_csv_file = Text_Manipulation.create_csv(case_file, "File Name", "Shock (Y/N)", "Number of Shocks")
        excel_file = Text_Manipulation.write_excel_remove_csv(data_csv_file, case_csv_file)
        Text_Manipulation.add_stats(excel_file)
else:
    Excel_File_Manipulation.add_title_column(path)
