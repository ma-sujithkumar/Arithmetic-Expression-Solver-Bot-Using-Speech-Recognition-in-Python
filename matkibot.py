
# coding: utf-8

# In[1]:

import speech_recognition as spec
import win32com.client as wincl
import wikipedia
speak = wincl.Dispatch("SAPI.SpVoice")
def bot():
    x=spec.Recognizer() 
    with spec.Microphone() as source:
        audio=x.listen(source)
    y=x.recognize_google(audio)
    print(y)
    y=y.encode("ascii")
    process(y)


# In[2]:

def calculate(y):
    try:
        sol=eval(y)
    except:
        speak.Speak("I can't understand, Sorry, Tell it again")
        bot()
    else:
        print(sol)
        speak.Speak("the answer is")
        speak.Speak(sol)
        speak.Speak("Anything else! say Yes or No")
        terminate()


# In[3]:

def count(y,defined):
    y=y.split()
    c=0
    for i in y:
        for j in defined:
            if i==j:
                c=c+1
    return c



# In[4]:

def terminate():
    x=spec.Recognizer() 
    with spec.Microphone() as source:
        audio=x.listen(source)
    y=x.recognize_google(audio)
    y=y.encode("ascii")
    print(y)
    if "no" in y.split():
        speak.Speak("Thank you Dood! Happy to help you always. For queries, please contact Sujithkumar")
        exit()
    if "yes" in y.split():
        speak.Speak("How may I help you?")
        bot()
    else:
        speak.Speak("Please Repeat again")
        terminate()


# In[5]:

def search():
    x=spec.Recognizer() 
    with spec.Microphone() as source:
        audio=x.listen(source)
    y=x.recognize_google(audio)
    y=y.encode("ascii")
    print(y)
    z=wikipedia.summary((y),sentences=5)
    print(z)
    speak.Speak(z)
    speak.Speak("Anything else! say Yes or No")
    terminate()


# In[6]:

def process(y):
    greet=["how","are","you"]
    c=count(y,greet)
    if c>=2:
        speak.Speak("I am fine! How are you doing?")
        bot()
    mat=["calculate","expression","mathematical"]
    d=count(y,mat)
    wiki=["wikipedia","search"]
    e=count(y,wiki)
    if e>=1:
        speak.Speak("Tell me something which you want to search! Note:The thing alone!")
        search()
    if d>=2:
        speak.Speak("Tell me the expression. Please use numbers and symbols only. to add, use plus. to subtract, use minus. to multiply use astereik. to divide, use slash")
        bot()
    else:
         calculate(y)
    terminate()
    
elements=["Getting Started","Hey Dood! I am your matkibot","I can calculate mathematical expressions and I can also search something in wikipedia","That's y mask named me Matkibot","now,ask me"]
for i in elements:
    speak.Speak(i)
bot()
