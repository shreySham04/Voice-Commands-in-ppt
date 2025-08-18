#Importing Libraries 
import speech_recognition as sr
import pyautogui
import pygetwindow as gw
import time
import re


def focus_powerpoint():
    windows = [w for w in gw.getWindowsWithTitle('PowerPoint') if w.visible] #Command helps making the script able to only react with ppt and not other applications, this stops any distraction for the script
    if windows:
        ppt_window = windows[0]
        ppt_window.activate()
        time.sleep(0.3) #Note- open only one ppt at a time and un script in bg
        return True
    return False

def voice_control():
    recognizer = sr.Recognizer()
    mic = sr.Microphone()

    print("🎤 Voice control started. Available commands: next, previous, start slideshow, resume slideshow, end slideshow, go to slide X, blank screen, show slide, close PowerPoint.")
    
    while True:
        with mic as source:
            recognizer.adjust_for_ambient_noise(source)
            print("Listening...")
            audio = recognizer.listen(source)

        try:
            command = recognizer.recognize_google(audio).lower()
            print(f"You said: {command}")

            if not focus_powerpoint():
                print("❌ PowerPoint not found or not open!")
                continue

            if "next" in command:
                pyautogui.press("right")
            elif "previous" in command or "back" in command:
                pyautogui.press("left")
            elif "start slideshow" in command:
                pyautogui.press("f5")
            elif "resume slideshow" in command:
                pyautogui.hotkey("shift", "f5")
            elif "end slideshow" in command or "stop slideshow" in command:  #In many case it will take "and" as input instead of "end" so use and instead.
                pyautogui.press("esc")
            elif "go to slide" in command:
                match = re.search(r"go to slide (\d+)", command)
                if match:
                    slide_number = match.group(1)
                    pyautogui.typewrite(slide_number)
                    pyautogui.press("enter")
            elif "blank screen" in command:
                pyautogui.press("b")
            elif "show slide" in command:
                pyautogui.press("b")
            elif "close powerpoint" in command:
                pyautogui.hotkey("alt", "f4")
            elif "exit" in command or "quit" in command:
                print("🛑 Stopping voice control.")
                break

        except sr.UnknownValueError:
            print("Didn't catch that. Try again.")
        except sr.RequestError:
            print("⚠️ Could not request results from Google Speech Recognition service.")

if __name__ == "__main__":
    voice_control()

