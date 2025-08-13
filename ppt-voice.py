import speech_recognition as sr
import pyautogui
import pygetwindow as gw
import time

def focus_powerpoint():
    # Find PowerPoint window
    windows = [w for w in gw.getWindowsWithTitle('PowerPoint') if w.visible]

    if windows:
        ppt_window = windows[0]
        ppt_window.activate()
        time.sleep(0.3)  # Small delay to ensure it gets focus
        return True
    return False

def voice_control():
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

            if not focus_powerpoint():
                print("❌ PowerPoint not found or not open!")
                continue

            if "next" in command:
                pyautogui.press("right")
            elif "previous" in command or "back" in command:
                pyautogui.press("left")
            elif "exit" in command or "stop" in command:
                print("🛑 Stopping voice control.")
                break
        except sr.UnknownValueError:
            print("Didn't catch that. Try again.")
        except sr.RequestError:
            print("⚠️ Could not request results from Google Speech Recognition service.")

if __name__ == "__main__":
    voice_control()
