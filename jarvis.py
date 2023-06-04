import datetime
import speech_recognition as sr #pip installation
import wikipedia #pip installation
import webbrowser
import smtplib

def speak(str):
    from win32com.client import Dispatch

    speak=Dispatch("SAPI.spVoice")
    speak.Speak(str)

def wish():
    hour=int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")

    else:
        speak("Good evening!")
    speak("I am jarvis. please tell me how may i help you sir ")

def takeCommad():
    # it take microphone input from the user and return string output
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("listing....")
        r.pause_threshold=1
        audio=r.listen(source)

    try:
        print("Recognizing....")
        query=r.recognize_google(audio,language="en-in")
        print(f"User said:  {query}\n")

    except Exception as e:
        # print(e)
        print("say that again please...")
        return "None"
    return query       

def sendEmail(to, content):
    server=smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('yourgmail.com','your-password')
    server.sendmail('youremail@gmail.com', to, content)
    server.close()

if __name__ == "__main__":
    # speak("Hello i am jarvis, your personal digital assistant.")
    wish()
    while True:
      query=takeCommad().lower()

      if 'wilipedia' in query:
          speak('Searching wikipedia....')
          query=query.replace("wikipedia", "")
          results=wikipedia.summary(query, sentences=2)
          speak("According to wikipedia")
          print(results)
          speak(results)

      elif'open youtube' in query:
          webbrowser.open("youtube.com")

      elif'open google' in query:
          webbrowser.open("google.com")

      elif'open stackoverflow' in query:
          webbrowser.open("stackoverflow.com")

      elif'open Linkedin' in query:
          webbrowser.open("linkedin.com")

      elif'open twitter' in query:
          webbrowser.open("twitter.com")

      elif'open gfg' in query:
          webbrowser.open("geeksforgeeks.com")

      elif'open javatpoint' in query:
          webbrowser.open("javatpoint.com")

      elif 'the time' in query:
          strTime=datetime.datetime.now().strftime("%H:%M:%S")
          print(strTime )
          speak(f"Sir, the time is {strTime}")

      elif 'email to rishi' in query:
          try:
              speak("what should I say?")
              content=takeCommad()
              to="rishiyouremail@gmail.com"
              sendEmail(to, content)
              speak("Email has been sent!")
          except Exception as e:
              print(e)
              speak("Sorry my friend rishi.I am not able to send this email")

    # logic for executing tasks based on query