import google.generativeai as genai
from dotenv import load_dotenv
import logging, verboselogs
import win32com.client as wincl
from time import sleep
import win32api
import win32gui

from deepgram import (
    DeepgramClient,
    DeepgramClientOptions,
    LiveTranscriptionEvents,
    LiveOptions,
    Microphone,
)

load_dotenv()

is_finals = []

if True:
    WM_APPCOMMAND = 0x319
    APPCOMMAND_MICROPHONE_VOLUME_MUTE = 0x180000
    hwnd_active = win32gui.GetForegroundWindow()
    win32api.SendMessage(hwnd_active, WM_APPCOMMAND, None, APPCOMMAND_MICROPHONE_VOLUME_MUTE)

def main():
    try:

        genai.configure(api_key="ENTER_API_KEY")
        model = genai.GenerativeModel('gemini-pro')
        deepgram: DeepgramClient = DeepgramClient("ENTER_API_KEY")

        dg_connection = deepgram.listen.live.v("1")

        def on_message(self, result, **kwargs):
            global is_finals
            sentence = result.channel.alternatives[0].transcript
            if len(sentence) == 0:
                return
            if result.is_final:
               
                is_finals.append(sentence)

                if result.speech_final:
                    utterance = ' '.join(is_finals)
                    print(f"Speech Final: {utterance}")
                    is_finals = []
                    response = model.generate_content(utterance)
                    print(response.text)

                    WM_APPCOMMAND = 0x319
                    APPCOMMAND_MICROPHONE_VOLUME_MUTE = 0x180000
                    hwnd_active = win32gui.GetForegroundWindow()
                    win32api.SendMessage(hwnd_active, WM_APPCOMMAND, None, APPCOMMAND_MICROPHONE_VOLUME_MUTE)

                    speak = wincl.Dispatch("SAPI.SpVoice")
                    speak.Speak(response.text)

                    WM_APPCOMMAND = 0x319
                    APPCOMMAND_MICROPHONE_VOLUME_MUTE = 0x180000
                    hwnd_active = win32gui.GetForegroundWindow()
                    win32api.SendMessage(hwnd_active, WM_APPCOMMAND, None, APPCOMMAND_MICROPHONE_VOLUME_MUTE)
                    
        def on_utterance_end(self, utterance_end, **kwargs):
            global is_finals
            if len(is_finals) > 0:
                utterance = ' '.join(is_finals)
                is_finals = []



        dg_connection.on(LiveTranscriptionEvents.Transcript, on_message)
        dg_connection.on(LiveTranscriptionEvents.UtteranceEnd, on_utterance_end)


        options: LiveOptions = LiveOptions(
            model="nova-2",
            language="en-US",
            smart_format=True,
            encoding="linear16",
            channels=1,
            sample_rate=16000,
            interim_results=True,
            utterance_end_ms="1000",
            vad_events=True,
            endpointing=500
        )

        addons = {
            "no_delay": "true"
        }

        print("\n\nPress Enter to stop recording...\n\n")
        if dg_connection.start(options, addons=addons) is False:
            print("Failed to connect to Deepgram")
            return
        microphone = Microphone(dg_connection.send)
        microphone.start()
        input("")
        microphone.finish()
        dg_connection.finish()

        print("Finished")

    except Exception as e:
        print(f"Could not open socket: {e}")
        return


if __name__ == "__main__":
    main()
