import cv2
import pytesseract
import serial
import time
import speech_recognition as sr
import win32com.client
import requests
import json
import datetime

# ================= TESSERACT PATH =================
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# ================= SPEAKER =================
speaker = win32com.client.Dispatch("SAPI.SpVoice")

def speak(text):
    print("\nAssistant:", text)
    speaker.Speak(text)

# ================= ESP32 =================
try:
    esp = serial.Serial("COM9", 115200, timeout=1)  # 🔴 CHANGE COM PORT
    time.sleep(2)
    speak("ESP32 connected")
except:
    esp = None
    speak("ESP32 not connected")

# ================= MICROPHONE =================
recognizer = sr.Recognizer()
recognizer.energy_threshold = 300
recognizer.dynamic_energy_threshold = True

def listen():
    try:
        with sr.Microphone() as source:
            print("\n🎤 Listening...")
            recognizer.adjust_for_ambient_noise(source, duration=0.8)
            audio = recognizer.listen(source, timeout=6, phrase_time_limit=6)

        text = recognizer.recognize_google(audio)
        print("You:", text)
        return text.lower()

    except sr.WaitTimeoutError:
        speak("Please say again")
        return ""

    except sr.UnknownValueError:
        speak("I did not understand")
        return ""

    except sr.RequestError:
        speak("Speech service error")
        return ""

# ================= PHI AI =================
OLLAMA_URL = "http://localhost:11434/api/generate"
MODEL = "phi"

def ask_phi(question):
    payload = {
        "model": MODEL,
        "prompt": question,
        "stream": False
    }

    try:
        r = requests.post(OLLAMA_URL, json=payload)
        return r.json()["response"]
    except:
        return "AI is not available"

# ================= DATE & TIME =================
def handle_date_time(cmd):
    now = datetime.datetime.now()

    if "time" in cmd:
        return now.strftime("The time is %I %M %p")

    if "date" in cmd or "day" in cmd:
        return now.strftime("Today is %A %d %B %Y")

    return None

# ================= OCR =================
def read_book():
    speak("Opening camera. Show the book and press S to read. Press Q to exit.")

    cap = cv2.VideoCapture(0, cv2.CAP_ANY)

    if not cap.isOpened():
        speak("Camera not accessible")
        return

    while True:
        ret, frame = cap.read()
        if not ret:
            speak("Camera error")
            break

        cv2.imshow("Book Reader | S = Scan | Q = Exit", frame)
        key = cv2.waitKey(1) & 0xFF

        if key == ord('s'):
            speak("Reading the book")

            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            gray = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)[1]

            text = pytesseract.image_to_string(gray)

            if text.strip() == "":
                speak("No readable text found")
            else:
                print("\n--- OCR TEXT ---\n")
                print(text)
                speak(text)

        elif key == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()

# ================= OBJECT PICK =================
def pick_ultrasonic():
    if esp is None:
        speak("ESP32 not connected")
        return

    speak("Show the ultrasonic sensor to the camera")

    cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
    if not cap.isOpened():
        speak("Camera error")
        return

    stable = 0

    while True:
        ret, frame = cap.read()
        if not ret:
            break

        frame = cv2.resize(frame, (640, 480))
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        blur = cv2.GaussianBlur(gray, (7, 7), 0)

        thresh = cv2.adaptiveThreshold(
            blur, 255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY_INV,
            11, 2
        )

        contours, _ = cv2.findContours(
            thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
        )

        found = False
        for cnt in contours:
            area = cv2.contourArea(cnt)
            if 2500 < area < 12000:
                x,y,w,h = cv2.boundingRect(cnt)
                if 0.6 < w/h < 2.0:
                    found = True
                    cv2.rectangle(frame,(x,y),(x+w,y+h),(0,255,0),2)

        if found:
            stable += 1
        else:
            stable = 0

        if stable > 10:
            speak("Ultrasonic sensor detected. Picking now.")
            esp.write(b"pick\n")
            time.sleep(2)
            break

        cv2.imshow("Ultrasonic Detection", frame)
        if cv2.waitKey(1) & 0xFF == 27:
            break

    cap.release()
    cv2.destroyAllWindows()

# ================= MAIN =================
speak("Smart Student Assistant started. You can speak now.")

while True:
    command = listen()
    if command == "":
        continue

    if any(word in command for word in ["exit", "stop", "bye"]):
        speak("Goodbye. Have a nice day.")
        break

    if "read" in command:
        read_book()
        continue

    if "pick" in command and "ultrasonic" in command:
        pick_ultrasonic()
        continue

    dt = handle_date_time(command)
    if dt:
        speak(dt)
        continue

    answer = ask_phi(command)
    speak(answer)