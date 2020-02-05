import os
import Text_Manipulation
import Excel_File_Manipulation

# The batch script runner for the text file and Excel file manipulations
path_finder = r"C:\Users\mnarcisi\Documents\Mike\Scientific Affairs"
first_folder_list = os.listdir(path_finder)
folder_list = []
for i in range(len(first_folder_list)):
    if first_folder_list[i].find(".") == -1:
        folder_list.append((first_folder_list[i]))

# Show user the relevant files
first_user_choice = 0
while first_user_choice < 1 or first_user_choice > len(folder_list):
    print('\033[1m' + "\nPlease choose which folder to open:\n")
    for j in range(len(folder_list)):
        print("(" + str(j + 1) + ") " + folder_list[j])
    first_user_choice = input()
    try:
        first_user_choice = int(first_user_choice)
        if first_user_choice < 1 or first_user_choice > len(folder_list):
            print("\nPlease select one of the given selections")
    except ValueError:
        print("\nPlease enter a numerical value")
        user_choice = 0

folder_path = path_finder + "\\" + folder_list[first_user_choice - 1]

#################### My Testing ############################################
#folder_path = input('\033[1m' + "\nPlease enter the path to the folder where the desired files are stored.\n")  ##
#folder_path = r"C:\Users\mnarcisi\Documents\Mike\Scientific Affairs\Data" ##
                                                                          ##
############################################################################

choice_to_continue = 1
while choice_to_continue == 1:
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
        Text_Manipulation.defib_shock_text_manipulation(folder_path)
    else:
        Excel_File_Manipulation.add_title_column(folder_path)
    choice_to_continue = int(input("\nManipulate more files?\n(1) Yes\n"))
print("\nDone")
