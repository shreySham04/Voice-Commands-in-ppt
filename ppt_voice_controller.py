import speech_recognition as sr
import pyautogui

recognizer = sr.Recognizer()
mic = sr.Microphone()

print("🎤 Voice control started. Say 'next' or 'previous'...")

while True:
    with mic as source:
        recognizer.adjust_for_ambient_noise(source)
        print("Listening...")
        audio = recognizer.listen(source)

    try:
        command = recognizer.recognize_google(audio).lower()
        print(f"You said: {command}")

        if "next" in command:
            pyautogui.press("right")  # Move to next slide
        elif "previous" in command:
            pyautogui.press("left")   # Move to previous slide
        elif "stop" in command:
            print("Stopping voice control.")
            break

    except sr.UnknownValueError:
        print("Could not understand audio.")
    except sr.RequestError:
        print("Speech service is down.")
