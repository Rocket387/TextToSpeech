# This program reads the docx file from the user input of the file path
# and then convert it to a speech file (mp3)

from gtts import gTTS
import docx
from playsound import playsound

#Depending on the size of docx file there can be a delay between the program reading the doc
# and the program reading the doc, The program will start on its own when it is ready

#
file_path = input("Please enter the file path you would like to convert to speech: ")


def getText():
    try:
        doc = docx.Document(file_path)  # Creating word reader object.
        data = ""
        fullText = []
        for para in doc.paragraphs:
            fullText.append(para.text)
            data = '\n'.join(fullText)

        print("Reading file, will commence speech shortly...")
        tts = gTTS(data)
        tts.save('output.mp3')
        print("Conversion complete, your text to speech file is ready")
        file = "output.mp3"
        playsound(file, True)

    except IOError:
        print('There was an error opening the file!')
        return


if __name__ == "__main__":
    getText()
