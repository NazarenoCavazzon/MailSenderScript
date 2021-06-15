import os
from os import walk
import smtplib
import pandas as pd
import shutil
import docx
from getpass import getpass

# ------------------------------------------VARIABLES------------------------------------------

lines = '-'*60
optionLine = "Select an Option: "
option = 0
config = False
empty = False
abc = ["A", "B", "C", "D", "E", "F", "G", "H",
       "I", "J", "K", "L", "M", "N", "O", "P",
       "Q", "R", "S", "T", "U", "V", "X", "Y", "Z"]

# ------------------------------------------FUNCTIONS------------------------------------------


def clear(): return os.system("cls")


def getFilesInFolder(path):
    f = []
    for (dirpath, dirnames, filenames) in walk(path):
        f.extend(filenames)
        names = filenames
        break
    return names


def getFoldersInFolder(path):
    f = []
    for (dirpath, dirnames, filenames) in walk(path):
        f.extend(dirnames)
        break
    return dirnames


def getDocsInFolder(path):
    count = []
    files = getFilesInFolder(path)
    for file in files:
        file = str(file)
        name, extension = os.path.splitext(path+file)
        if extension == ".docx":
            count.append(file)
    return count


def checkExcelInFolder(path):
    count = []
    files = getFilesInFolder(path)
    for file in files:
        file = str(file)
        name, extension = os.path.splitext(path+file)
        if extension == ".xlsx":
            count.append(file)
    return count


def createTextFile(name):
    txt = open(name+'.txt', 'a')
    txt.close()


def createFolderIN(path, name):
    try:
        os.chdir(path)
        os.makedirs(name)
    except:
        pass


def createFolderON(name):
    try:
        path = os.getcwd()
        os.chdir(path)
        os.makedirs(name)
    except:
        pass


def removeTextFiles(*args):
    for i in args:
        os.remove(i+'.txt')


def getDataFrom():

    excelsFolder = os.path.abspath(os.path.join(os.path.dirname(
        __file__), '..', 'MailScript', 'Excel Repository'))

    # Transform the excelFolder string into a raw string to not have problems with the document reader
    excelsFolder = r"{}".format(excelsFolder)

    excelsInFolder = checkExcelInFolder(excelsFolder)

    excelSelected = []

    clear()
    print(lines)
    print("We found", len(excelsInFolder),
          "excels in the folder, select the ones you want to get data from")
    print(lines)
    input("Press enter to continue... ")
    clear()

    for excel in excelsInFolder:
        error = False
        clear()
        print(lines)
        print("Get data from", excel+'?')
        print(lines)
        print("1_ Yes")
        print("2_ No")
        print(lines)
        option = int(input(optionLine))

        if option > 2 or option < 1:
            error = True

        while error == True:
            clear()
            print(lines)
            print("Invalid Option, please select a number between 1 and 2")
            print(lines)
            print("1_ Yes")
            print("2_ No")
            print(lines)
            option = int(input(optionLine))
            if numerodesc >= 0:
                error = False

        if option == 1:
            excelSelected.append(excel)

        elif option == 2:
            pass

    data_dir = os.path.abspath(os.path.join(os.path.dirname(
        __file__), '..', 'MailScript', 'Data'))

    for ex in excelSelected:
        excelFile = os.path.abspath(os.path.join(os.path.dirname(
            __file__), '..', 'MailScript', 'Excel Repository', ex))

        excelFile = r"{}".format(excelFile)

        data_dir = os.path.abspath(os.path.join(os.path.dirname(
            __file__), '..', 'MailScript', 'Data'))

        ex = ex[:len(ex)-5]

        createFolderIN(data_dir, ex)

        excelFolder = os.path.abspath(os.path.join(os.path.dirname(
            __file__), '..', 'MailScript', 'Data', ex))

        cont = 0
        text_file = ""
        txt = 'New.txt'
        file_list = []

        for car in abc:
            cont = 0
            df = pd.read_excel(excelFile, sheet_name=0, usecols=car)
            df.to_csv(txt, header=True, index=False)
            fo = open(txt, 'r')
            content = fo.readlines()

            if (os.stat(txt).st_size > 2) == False:
                fo.close()
                empty == True
                os.remove(txt)

            for i in content:
                var = str(i)
                if cont == 0:
                    l = len(var)
                    rem_last = var[:l-1]
                    createTextFile(rem_last)
                    text_file = rem_last+".txt"
                    file_list.append(str(text_file))
                    cont = 1
                else:
                    add = open(text_file, 'a')
                    add.write(var)
                    add.close()
            fo.close()

        file_list = set(file_list)
        file_list.remove(".txt")
        for file in file_list:
            try:
                shutil.move(file, excelFolder)
            except shutil.Error:
                length = len(file)
                file_woTXT = file[:length-4]
                print(file_woTXT, "--- Already Uploaded")
                os.remove(file)
        os.remove(".txt")


def enter(mail, password):
    try:
        # Inicio de sesion
        conexion.login(user=mail, password=password)
        print("Valid Credentials")
        return True

    except:
        input("Wrong Credentials... ")
        return False


def removeExtensions(array):
    array_list = []
    for i in array:
        length = len(i)
        i = i[:length-4]
        array_list.append(i)
    return array_list


def printArray(array):
    for i, j in enumerate(array):
        print(str(i+1)+'_ '+str(j))


def dataBase():
    mainPath = os.path.abspath(os.path.join(os.path.dirname(
        __file__), '..', 'MailScript', 'Data'))
    return getFoldersInFolder(mainPath)


def sendMail(sender, reciever):
    subject = 'Test'
    body = 'I wanna go to europe'
    msg = f'Subject: {subject}\n\n{body}'
    conexion.sendmail(sender, reciever, msg)


def sumTextToList(dataSelectedFolder, *args):
    dataList = []
    array = []
    for i in args:  # ["Region.txt", "First Name.txt"], ["Region.txt", "First Name.txt"]
        # array [["Region.txt", "First Name.txt"], ["Region.txt", "First Name.txt"]]
        array.append(i)
    # ["Region.txt", "First Name.txt"], ["Region.txt", "First Name.txt"]
    for textFiles in array:
        path = os.path.abspath(os.path.join(os.path.dirname(
            __file__), '..', 'MailScript', 'Data', dataSelectedFolder, textFiles))
        text = open(path, "r").readlines()
        for lines in textFiles:
            dataList.append(lines)
    return print(dataList)


def sendMassiveMails(selectedDataBase, textFiles, wordFile, mails, mail):
    textFilesWE = removeExtensions(textFiles)
    allFiles = []
    mailWT = []
    number = 0
    mails = path_data = os.path.abspath(os.path.join(os.path.dirname(
        __file__), '..', 'MailScript', 'Data', selectedDataBase, mails))
    for file in textFiles:
        file = os.path.abspath(os.path.join(os.path.dirname(
            __file__), '..', 'MailScript', 'Data', selectedDataBase, file))
        semiList = []
        file = open(file, "r").readlines()
        for word in file:
            semiList.append(word.strip())
        allFiles.append(semiList)
    mails = open(mails, "r").readlines()
    for i in mails:
        mailWT.append(i.strip())
    for mn, ma in enumerate(mailWT):
        first, second = readDocs(wordFile)
        body = subject = msg = ' '
        for i, j in enumerate(textFilesWE):
            for m, p in enumerate(first):
                if p == '['+j+']':
                    print("Primer parrafo", mn, ma)
                    first[m] = allFiles[i][mn]
            for f, g in enumerate(second):
                if g == '['+j+']':
                    print("Segundo parrafo", mn, ma)
                    second[f] = allFiles[i][mn]
            print("For textFilesWE", mn, ma)
        subject = u' '.join(first).encode('utf-8')
        body = u' '.join(second).encode('utf-8')
        msg = f'Subject: {subject}\n\n{body}'
        conexion.sendmail(mail, ma, msg)
        print("For mails", mn, ma)


def readDocs(docFile):
    path = mainPath = os.path.abspath(os.path.join(os.path.dirname(
        __file__), '..', 'MailScript', 'Mails', docFile))
    doc = docx.Document(path)
    firstParagraph = doc.paragraphs[0].text.split()
    secondParagraph = doc.paragraphs[1].text.split()
    return firstParagraph, secondParagraph


# ------------------------------------------START/SETUP------------------------------------------


# Establecer conexion con el servidor de SMTP Gmail
conexion = smtplib.SMTP(host='smtp.gmail.com', port=587)
conexion.ehlo()

# Encriptacion TLS
conexion.starttls()

createTextFile("Messssssssssssssi")

createFolderON("Excel Repository")
createFolderON("Data")
createFolderON("Mails")


# ------------------------------------------FUNCTIONS------------------------------------------
while option != 5:
    print(lines)
    print("-----------------Welcome to the Mail Script-----------------")
    print(lines)
    print("Select any option")
    print("1_ Upload an Excel file with Data")
    print("2_ Send Mails")
    print(lines)
    option = int(input(optionLine))

    if option == 1:
        print(lines)
        print("Put your Excels in the Excel directory")
        # print("")
        print(lines)
        input("Press enter to Confirm... ")
        getDataFrom()
        input("New Data Uploaded... ")
        clear()

    if option == 2:
        data_Files = []
        error = False
        clear()
        print(lines)
        print("Login with your gmail")
        print(lines)
        mail = input("Mail: ")
        password = getpass()
        print(lines)
        succesConection = enter(mail, password)

        while succesConection == False:
            clear()
            print(lines)
            print("Please Retry")
            print(lines)
            mail = input("Mail: ")
            password = getpass()
            print(lines)
            succesConection = enter(mail, password)

        print(lines)
        input("Press enter to continue... ")
        clear()
        print(lines)
        print("From what Excel Data base you want to import data")
        print(lines)

        printArray(dataBase())

        print(lines)
        dataBaseSelection = int(input(optionLine))

        if dataBaseSelection > len(dataBase()) or dataBaseSelection < 1:
            error = True

        while error == True:
            clear()
            print(lines)
            print("Invalid Option, please select a valid number")
            print(lines)

            printArray(dataBase())

            print(lines)
            dataBaseSelection = int(input(optionLine))

            if dataBaseSelection > 0 and dataBaseSelection < int(len(dataBase())+1):
                error = False

        selectedDataBase = dataBase()[dataBaseSelection-1]

        dataSelectedFolder = os.path.abspath(os.path.join(os.path.dirname(
            __file__), '..', 'MailScript', 'Data', selectedDataBase))

        mailFolder = os.path.abspath(os.path.join(os.path.dirname(
            __file__), '..', 'MailScript', 'Mails'))

        data_Files = getFilesInFolder(dataSelectedFolder)

        clear()
        print(lines)
        print("Select what data file have the Mails")
        print(lines)
        printArray(data_Files)
        print(lines)
        email_textfile = int(input(optionLine))
        email_database = data_Files[email_textfile-1]
        data_Files.remove(email_database)
        clear()
        data_FilesWE = []
        data_FilesWE = removeExtensions(data_Files)
        docsInMails = []
        print(lines)
        print("Data loaded succesfully, now you can write a Email in word or in a text file with the next variables")
        print(*data_FilesWE, sep=", ")
        print(lines)
        input("Press enter when you got your mail ready in the folder... ")
        docsInMails = getDocsInFolder(mailFolder)
        for i in docsInMails:
            if i[0] == "~" and i[1] == "$":
                docsInMails.remove(i)
        clear()
        print(lines)
        print("Select the word document with the mail draft")
        print(lines)
        printArray(docsInMails)
        print(lines)
        wordSelection = int(input(optionLine))
        if wordSelection > len(dataBase()) or wordSelection < 1:
            error = True

        while error == True:
            clear()
            print(lines)
            print("Invalid Option, please select a valid number")
            print(lines)
            printArray(docsInMails)
            print(lines)
            wordSelection = int(input(optionLine))

            if wordSelection > 0 and wordSelection < int(len(docsInMails)+1):
                error = False

        wordSelected = docsInMails[wordSelection-1]
        sendMassiveMails(dataSelectedFolder, data_Files,
                         wordSelected, email_database, mail)
        print(wordSelected)
        print(data_Files)

    if option == 2:
        pass
