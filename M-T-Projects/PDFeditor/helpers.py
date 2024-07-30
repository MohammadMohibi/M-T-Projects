def checkIfBlank(inputStr, typeStr):
    while inputStr.strip() == "":
        print(f"Please enter non blank {typeStr}")
        inputStr = input("Provide correct input: ")
        
    return inputStr