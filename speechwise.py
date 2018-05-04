import speech_recognition as sr
import openpyxl

r = sr.Recognizer()

with sr.Microphone() as source:
    print('Say Something!')
    audio = r.listen(source)
    print('Done!')
try:
    text = r.recognize_google(audio)
    print('Speech Recognizer thinks you said:\n' + text)
    wb = openpyxl.load_workbook('C:\\Users\\pathikrit\\Desktop\\Book1.xlsx')
    print (text)


except Exception as e:
    print(e)

