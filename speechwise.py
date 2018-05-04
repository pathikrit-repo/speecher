import speech_recognition as sr
import openpyxl


def getAudioInputFromUser(speechRecognizer, soundSource, row, column, excelWorkBook):

    with soundSource as source:
        audio = speechRecognizer.listen(source)
    print ("input recieved")
    text = speechRecognizer.recognize_google(audio)
    excelSheet = excelWorkBook['Sheet1']
    excelSheet.cell(row=row, column=column, value=text)
    print(text)
    excelWorkBook.save("Book1.xlsx")
    return text



r = sr.Recognizer()

# with sr.Microphone() as source:

try:
    varIdentifier = "start"
    varExcelrow = 1
    varExcelcolumn = 1
    varIteration = 1
    varMainChoices = ["given", "and", "when", "then", "wrap it up"]

    wb = openpyxl.load_workbook(filename='C:\\Users\\pathikrit\\Desktop\\Book1.xlsx')
    print("Welcome to the automated test creation application.")

    while varIdentifier != "wrap it up":
        print("select any one of the command by saying the command avilable commands are.\n given, and, when, then or just say wrap it up to finish")
        print("Listening...\n")
        resultCommand = getAudioInputFromUser(r, sr.Microphone(), varExcelrow, varExcelcolumn, wb)
        if resultCommand.lower() in varMainChoices:
            if resultCommand.lower() != "wrap it up":
                varExcelcolumn = varExcelcolumn+1
                # command is now in the excel file now we need a sentence for the command
                print("say the sentence for the "+resultCommand+" Command")
                print("Listening...\n")
                resultDetail = getAudioInputFromUser(r, sr.Microphone(), varExcelrow, varExcelcolumn, wb)
                print("you said "+resultDetail)
                varExcelrow = varExcelrow+1
                varExcelcolumn=1
            else:
                print("finish because user said wrap it up")
                # take everything to .txt file except the last row where wrap it up is written.
        else:
            print("error because what user said is not among the commands")
           # last input command was something wrong handle that atleast delete it from the workbook.


except Exception as e:
    print (e)

