import os
import shutil
import pandas as pd
import csv

#  This function is used to filter all text files in a folder of sample defib data and pick out all "defib shock"
#  rows and their times
def defib_shock_data_consolidation(path):
    file_list = os.listdir(path)

    # Ensure only .txt files are included
    user_list = []
    for i in range(len(file_list)):
        list_element = file_list[i]
        if (list_element[len(list_element) - 4: len(list_element)] == ".txt" or list_element[len(list_element) - 4:
        len(list_element)] == ".log") and list_element != "Defib_Shock_Master_Data_File.txt":
            user_list.append(list_element)

    # Exit function if there are no .txt files in the folder
    if len(user_list) == 0:
        print("\nNo .txt files to manipulate.  Aborting Script.")
        return "N/A"

    # Create a directory to place all processed .txt files (if one hasn't already been created)
    directory_name = path + "\\" + "Processed_.txt_Files"
    try:
        os.mkdir(directory_name)
        print("\nSuccessfully created the directory %s to store processed .txt files." % directory_name)
    except OSError:
        print("\nProcessed .txt files are stored in the directory %s." % directory_name)

    # Create a directory to place all .log files (if one hasn't already been created)
    log_directory_name = path + "\\" + ".log Files"
    try:
        os.mkdir(log_directory_name)
        print("\nSuccessfully created the directory %s to store .log files." % log_directory_name)
    except OSError:
        pass

    # Iterate over every relevant .txt file in the folder and manipulate them
    print("\nOne moment please...")
    master_file_name = path + "\\" + "Defib_Shock_Master_Data_File.txt"
    for j in range(0, len(user_list)):
        file_name = user_list[j]
        file_path = path + "\\" + user_list[j]
        # Move any .log files to specified directory
        if file_name[len(file_name) - 4: len(file_name)] == ".log":
            shutil.move(file_path, log_directory_name + "\\" + user_list[j])
            continue
        # Remove '.txt' and trailing digits from the end of the file name (this is what will be pasted in the file)
        updated_name_format = ""
        reversed_file_name = ""
        is_extra_digit = False
        # Reverse string to find last instance of '_'
        for k in range(len(file_name), 0, -1):
            reversed_file_name = reversed_file_name + file_name[k - 1]
        underscore_marker = reversed_file_name.find("_")
        # If '_' has a character (letter) to its left, it doesn't have extra digits
        if underscore_marker != -1:
            try:
                character_or_integer = int(file_name[len(file_name) - 1 - underscore_marker - 1])
                is_extra_digit = True
            except ValueError:
                is_extra_digit = False
            # If there are extra digits, remove them along with the file extension
            if is_extra_digit:
                text_to_remove = file_name[
                                 len(file_name) - 1 - underscore_marker: len(file_name)]
                updated_name_format = file_name.replace(text_to_remove, "")
            else:
                updated_name_format = file_name.replace(".txt", "")
        else:
            updated_name_format = file_name.replace(".txt", "")

        # Read the selected .txt file line by line and copy the Defib Shock times to a new file
        file = open(file_path, "r")
        totals_file = open(master_file_name, "a+")
        for line in file:
            line = line.lstrip()
            try:
                int_check = int(line[0])  # Use this statement to check if the leading character is an integer
                time = 0.0
                defib_shock_position = line.find("DEFIB SHOCK")
                if defib_shock_position != -1:
                    # Save name of line item
                    defib_shock = line[defib_shock_position: defib_shock_position + 11]
                    # Save time of element
                    front_bracket = line.find("[")
                    rear_bracket = line.find("]")
                    line = line[front_bracket + 1: rear_bracket]
                    line = line.lstrip()
                    try:
                        time = float(line)
                        totals_file.write(defib_shock + "  |  " + str(time) + "  |  " + updated_name_format + "\n")
                    except ValueError:
                        pass
            except IndexError:
                continue
            except ValueError:
                continue
        file.close()
        totals_file.close()

        # Move current .txt to the directory for processed .txt files
        shutil.move(file_path, directory_name + "\\" + user_list[j])

    print("\n.txt files manipulated successfully.")
    return master_file_name


# Create a .csv file containing all of the data from the text file
def createCSV(path):
    text_file = path
    csv_file = path.replace(".txt", ".csv")
    text_input = csv.reader(open(text_file), delimiter = "|")
    csv_output = csv.writer(open(csv_file, "w", newline = "\n"))
    csv_output.writerow(["Element Name", "Time (sec)", "File Name"])
    csv_output.writerows(text_input)
    return csv_file


# Create an Excel file based on the .csv file
def createExcel(path):
    df = pd.read_csv(path)
    print(df)
    #df.to_excel(path + "\\" + "Defib_Shock_Master_Data_File.xlsx", "Sheet1")
