import os

def checkIfBlank(inputStr, typeStr):
    while inputStr.strip() == "":
        print(f"Please enter non blank {typeStr}.")
        inputStr = input("Provide correct input: ")
        
    return inputStr

def checkPathCreate(pathName, site):
    newSite = f"{pathName}\\{site}"
    if not os.path.exists(newSite):
        os.makedirs(newSite)

    return newSite