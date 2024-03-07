import win32com.client
import os
import speech_recognition as sr
import pyautogui
import time

def record_text():
    recognizer = sr.Recognizer()
    microphone = sr.Microphone()
    with microphone as source:
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)
    try:
        return recognizer.recognize_google(audio).lower()
    except sr.UnknownValueError:
        return "unknown"
    except sr.RequestError:
        return "error"

def open_powerpoint_presentation(file_path):
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True

    presentation = powerpoint.Presentations.Open(file_path)
    slides_count = presentation.Slides.Count

    # Set slideshow settings
    slideshow_settings = presentation.SlideShowSettings
    slideshow_settings.ShowWithNarration = False
    slideshow_settings.ShowWithAnimation = False
    slideshow_settings.RangeType = 2  # ppShowAll
    slideshow_settings.StartingSlide = 1
    slideshow_settings.EndingSlide = slides_count

    # Start slideshow
    presentation.SlideShowSettings.Run()

    while True:
        command = record_text()
        print(f"Command recognized: {command}")

        if command == "previous":
            pyautogui.hotkey('left')
        elif command == "next":
            pyautogui.hotkey('right')
        elif command.startswith("go to"):
            try:
                slide_number = int(command.split()[2])
                if 1 <= slide_number <= slides_count:
                    presentation.SlideShowWindow.View.GotoSlide(slide_number)
                    print(f"Going to slide {slide_number}")
                else:
                    print("Invalid slide number")
            except IndexError:
                print("No slide number provided")
            except ValueError:
                print("Invalid slide number")
        elif command == "exit":
            powerpoint.Quit()
            break
        else:
            print("Invalid command")

        # Give some time for PowerPoint to process the commands
        time.sleep(0.5)

if __name__ == "__main__":
    file_path = input("Enter the path to the PowerPoint presentation: ")
    if os.path.exists(file_path):
        open_powerpoint_presentation(file_path)
    else:
        print("File not found.")
