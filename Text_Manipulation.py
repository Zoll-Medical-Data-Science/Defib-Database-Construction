import os

#  This function is used to filter a text file of sample defib data and pick out all "defib shock" rows,
#  their times, and the file name
def defib_shock_text_manipulation(path):
    file_list = os.listdir(path)

    # Ensure only .txt files are shown to user
    user_list = []
    for i in range(len(file_list)):
        list_element = file_list[i]
        if list_element[len(list_element) - 4: len(list_element)] == ".txt" and list_element.find("(Filtered)") == -1:
            user_list.append(list_element)

    # Exit function if there are no .txt files in the folder
    if len(user_list) == 0:
        print("\nNo .txt files to manipulate.")
        return

    # Show user the relevant files
    user_choice = 0
    while user_choice < 1 or user_choice > len(user_list):
        print('\033[1m' + "\nPlease choose which .txt file to manipulate (Enter number only):\n")
        for j in range(len(user_list)):
            print("(" + str(j + 1) + ") " + user_list[j])
        user_choice = input()
        try:
            user_choice = int(user_choice)
            if user_choice < 1 or user_choice > len(user_list):
                print("\nPlease select one of the given selections")
        except ValueError:
            print("\nPlease enter a numerical value")
            user_choice = 0
    selected_file_name = path + "\\" + user_list[user_choice - 1]

    # Remove '.txt' and trailing digits from the end of the file name
    formatted_file_name = user_list[user_choice - 1]
    updated_name_format = ""
    reversed_file_name = ""
    underscore_marker = -1
    is_extra_digit = False
    # Reverse string to find last instance of '_'
    for k in range(len(formatted_file_name), 0, -1):
        reversed_file_name = reversed_file_name + formatted_file_name[k - 1]
    underscore_marker = reversed_file_name.find("_")
    # If '_' has a character to its left, it doesn't have extra digits
    if underscore_marker != -1:
        try:
            character_or_integer = int(formatted_file_name[len(formatted_file_name) - 1 - underscore_marker - 1])
            is_extra_digit = True
        except ValueError:
            is_extra_digit = False
        # If there are extra digits, remove them along with the file extension
        if is_extra_digit:
            text_to_remove = formatted_file_name[
                             len(formatted_file_name) - 1 - underscore_marker: len(formatted_file_name)]
            updated_name_format = formatted_file_name.replace(text_to_remove, "")
        else:
            updated_name_format = formatted_file_name.replace(".txt", "")
    else:
        updated_name_format = formatted_file_name.replace(".txt", "")

    # Read the selected .txt file line by line and copy the Defib Shock times to a new file
    file = open(selected_file_name, "r")
    new_file = open(selected_file_name[0 : len(selected_file_name) - 4] + "_(Filtered)" + ".txt", "w")
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
                    new_file.write(defib_shock + "  |  " + str(time) + "  |  " + updated_name_format + "\n")
                except ValueError:
                    pass
        except IndexError:
            continue
        except ValueError:
            continue

    print("\n.txt files manipulated successfully.")
