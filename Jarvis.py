#JARVIS
from pprint import pprint
import methods
import time
import apiai
import json
import speech_recognition as sr

CLIENT_ACCESS_TOKEN = '8c564ca4f1ae4c45acea584bc16116b5'

ai = apiai.ApiAI(CLIENT_ACCESS_TOKEN)
request = ai.text_request()
request.lang = 'en'
request.session_id = 'user_id_lelo'

# handle = methods.getWindowWithTitle("untitled - notepad")
# monHandle = win32api.MonitorFromWindow(handle)
# print(win32api.GetMonitorInfo(monHandle))


# this is called from the background thread
def callback(recognizer, audio):
    print("-------------")
    try:
        print("DEBUG: recognizing..")
        recognizedText = r.recognize_google(audio)
        print(">> " + recognizedText)
    except:
        print('DEBUG: Error audio')
        return
    try:
        print("DEBUG: requesting... for: " + recognizedText)
        request = ai.text_request()
        request.query = recognizedText
        response = request.getresponse()
        responseStr = response.read().decode('utf-8')
        response_obj = json.loads(responseStr)
    except:
        print("Error getting response")
    evaluateIntent(response_obj)


def evaluateIntent(response_obj):
    intent = response_obj['result']['action']
    print('DEBUG: intent: ' + intent)

    if intent == "minimize":
        target_window = response_obj['result']['parameters']['target_window']
        print('DEBUG: target_window: ', target_window)
        target_monitor = response_obj['result']['parameters']['target_monitor']
        print('DEBUG: target_monitor: ', target_monitor)
        methods.minimize(target_window, target_monitor)
    elif intent == "create_shortcut":
        methods.createShortcut("")


# obtain audio from the microphone
r = sr.Recognizer()
m = sr.Microphone()
with m as source:
    print("Say something: ")
    r.adjust_for_ambient_noise(source, duration=1)  # listen for 1 second to calibrate the energy threshold for ambient noise levels

# start listening in the background (note that we don't have to do this inside a `with` statement)
stop_listening = r.listen_in_background(m, callback)
# `stop_listening` is now a function that, when called, stops background listening

# main loop
while True:
    time.sleep(0.1)
