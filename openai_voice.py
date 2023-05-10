from tendo import singleton
import sys

try:
    me = singleton.SingleInstance()
except:
    from tkinter import messagebox

    messagebox.showwarning("Warning", "OpenAI Virtual Assistant is already running")
    sys.exit()
import pvporcupine
import pvcobra
import pyaudio
import textwrap
import time
import openai
import struct
import pystray
import wave
import os
from tkinter import simpledialog
from tkinter import *
from pynput.keyboard import Key, Controller
import pyttsx3
from winreg import *
from PIL import Image
import string

# from gtts import gTTS
# import pygame
import winsound
from ctypes import wintypes
import ctypes
import comtypes
import re
from pystray import MenuItem as item
from threading import Thread
from XInput import *
import wmi


def get_resource_path(relative_path):
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def is_voicemeeter():
    default_output_device_index = pa.get_default_output_device_info()["index"]
    name = pa.get_device_info_by_index(default_output_device_index)["name"].lower()
    return ("virtual" in name) or ("vb-audio" in name)


MMDeviceApiLib = comtypes.GUID("{2FDAAFA3-7523-4F66-9957-9D5E7FE698F6}")
IID_IMMDevice = comtypes.GUID("{D666063F-1587-4E43-81F1-B948E807363F}")
IID_IMMDeviceCollection = comtypes.GUID("{0BD7A1BE-7A1A-44DB-8397-CC5392387B5E}")
IID_IMMDeviceEnumerator = comtypes.GUID("{A95664D2-9614-4F35-A746-DE8DB63617E6}")
IID_IAudioEndpointVolume = comtypes.GUID("{5CDF2C82-841E-4546-9722-0CF74078229A}")
CLSID_MMDeviceEnumerator = comtypes.GUID("{BCDE0395-E52F-467C-8E3D-C4579291692E}")
eRender = 0
keyboard = Controller()
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume

# ERole
c = wmi.WMI()
eConsole = 0  # games, system sounds, and voice commands
eMultimedia = 1  # music, movies, narration
eCommunications = 2  # voice communications
LPCGUID = REFIID = ctypes.POINTER(comtypes.GUID)
LPFLOAT = ctypes.POINTER(ctypes.c_float)
LPDWORD = ctypes.POINTER(wintypes.DWORD)
LPUINT = ctypes.POINTER(wintypes.UINT)
LPBOOL = ctypes.POINTER(wintypes.BOOL)
PIUnknown = ctypes.POINTER(comtypes.IUnknown)


class IMMDevice(comtypes.IUnknown):
    _iid_ = IID_IMMDevice
    _methods_ = (
        comtypes.COMMETHOD(
            [],
            ctypes.HRESULT,
            "Activate",
            (["in"], REFIID, "iid"),
            (["in"], wintypes.DWORD, "dwClsCtx"),
            (["in"], LPDWORD, "pActivationParams", None),
            (["out", "retval"], ctypes.POINTER(PIUnknown), "ppInterface"),
        ),
        comtypes.STDMETHOD(ctypes.HRESULT, "OpenPropertyStore", []),
        comtypes.STDMETHOD(ctypes.HRESULT, "GetId", []),
        comtypes.STDMETHOD(ctypes.HRESULT, "GetState", []),
    )


PIMMDevice = ctypes.POINTER(IMMDevice)


class IMMDeviceCollection(comtypes.IUnknown):
    _iid_ = IID_IMMDeviceCollection


PIMMDeviceCollection = ctypes.POINTER(IMMDeviceCollection)


class IMMDeviceEnumerator(comtypes.IUnknown):
    _iid_ = IID_IMMDeviceEnumerator
    _methods_ = (
        comtypes.COMMETHOD(
            [],
            ctypes.HRESULT,
            "EnumAudioEndpoints",
            (["in"], wintypes.DWORD, "dataFlow"),
            (["in"], wintypes.DWORD, "dwStateMask"),
            (["out", "retval"], ctypes.POINTER(PIMMDeviceCollection), "ppDevices"),
        ),
        comtypes.COMMETHOD(
            [],
            ctypes.HRESULT,
            "GetDefaultAudioEndpoint",
            (["in"], wintypes.DWORD, "dataFlow"),
            (["in"], wintypes.DWORD, "role"),
            (["out", "retval"], ctypes.POINTER(PIMMDevice), "ppDevices"),
        ),
    )

    @classmethod
    def get_default(cls, dataFlow, role):
        enumerator = comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator, cls, comtypes.CLSCTX_INPROC_SERVER
        )
        return enumerator.GetDefaultAudioEndpoint(dataFlow, role)


class IAudioEndpointVolume(comtypes.IUnknown):
    _iid_ = IID_IAudioEndpointVolume
    _methods_ = (
        comtypes.STDMETHOD(ctypes.HRESULT, "RegisterControlChangeNotify", []),
        comtypes.STDMETHOD(ctypes.HRESULT, "UnregisterControlChangeNotify", []),
        comtypes.COMMETHOD(
            [],
            ctypes.HRESULT,
            "GetChannelCount",
            (["out", "retval"], LPUINT, "pnChannelCount"),
        ),
        comtypes.COMMETHOD(
            [],
            ctypes.HRESULT,
            "SetMasterVolumeLevel",
            (["in"], ctypes.c_float, "fLevelDB"),
            (["in"], LPCGUID, "pguidEventContext", None),
        ),
        comtypes.COMMETHOD(
            [],
            ctypes.HRESULT,
            "SetMasterVolumeLevelScalar",
            (["in"], ctypes.c_float, "fLevel"),
            (["in"], LPCGUID, "pguidEventContext", None),
        ),
        comtypes.COMMETHOD(
            [],
            ctypes.HRESULT,
            "GetMasterVolumeLevel",
            (["out", "retval"], LPFLOAT, "pfLevelDB"),
        ),
        comtypes.COMMETHOD(
            [],
            ctypes.HRESULT,
            "GetMasterVolumeLevelScalar",
            (["out", "retval"], LPFLOAT, "pfLevel"),
        ),
        comtypes.COMMETHOD(
            [],
            ctypes.HRESULT,
            "SetChannelVolumeLevel",
            (["in"], wintypes.UINT, "nChannel"),
            (["in"], ctypes.c_float, "fLevelDB"),
            (["in"], LPCGUID, "pguidEventContext", None),
        ),
        comtypes.COMMETHOD(
            [],
            ctypes.HRESULT,
            "SetChannelVolumeLevelScalar",
            (["in"], wintypes.UINT, "nChannel"),
            (["in"], ctypes.c_float, "fLevel"),
            (["in"], LPCGUID, "pguidEventContext", None),
        ),
        comtypes.COMMETHOD(
            [],
            ctypes.HRESULT,
            "GetChannelVolumeLevel",
            (["in"], wintypes.UINT, "nChannel"),
            (["out", "retval"], LPFLOAT, "pfLevelDB"),
        ),
        comtypes.COMMETHOD(
            [],
            ctypes.HRESULT,
            "GetChannelVolumeLevelScalar",
            (["in"], wintypes.UINT, "nChannel"),
            (["out", "retval"], LPFLOAT, "pfLevel"),
        ),
        comtypes.COMMETHOD(
            [],
            ctypes.HRESULT,
            "SetMute",
            (["in"], wintypes.BOOL, "bMute"),
            (["in"], LPCGUID, "pguidEventContext", None),
        ),
        comtypes.COMMETHOD(
            [], ctypes.HRESULT, "GetMute", (["out", "retval"], LPBOOL, "pbMute")
        ),
        comtypes.COMMETHOD(
            [],
            ctypes.HRESULT,
            "GetVolumeStepInfo",
            (["out", "retval"], LPUINT, "pnStep"),
            (["out", "retval"], LPUINT, "pnStepCount"),
        ),
        comtypes.COMMETHOD(
            [],
            ctypes.HRESULT,
            "VolumeStepUp",
            (["in"], LPCGUID, "pguidEventContext", None),
        ),
        comtypes.COMMETHOD(
            [],
            ctypes.HRESULT,
            "VolumeStepDown",
            (["in"], LPCGUID, "pguidEventContext", None),
        ),
        comtypes.COMMETHOD(
            [],
            ctypes.HRESULT,
            "QueryHardwareSupport",
            (["out", "retval"], LPDWORD, "pdwHardwareSupportMask"),
        ),
        comtypes.COMMETHOD(
            [],
            ctypes.HRESULT,
            "GetVolumeRange",
            (["out", "retval"], LPFLOAT, "pfLevelMinDB"),
            (["out", "retval"], LPFLOAT, "pfLevelMaxDB"),
            (["out", "retval"], LPFLOAT, "pfVolumeIncrementDB"),
        ),
    )

    @classmethod
    def get_default(cls):
        endpoint = IMMDeviceEnumerator.get_default(eRender, eMultimedia)
        interface = endpoint.Activate(cls._iid_, comtypes.CLSCTX_INPROC_SERVER)
        return ctypes.cast(interface, ctypes.POINTER(cls))


engine = pyttsx3.init()
voices = engine.getProperty("voices")
engine.setProperty("voice", voices[1].id)
keyboard = Controller()
WAKE_WORD = "computer"
WAVE_OUTPUT_FILENAME = os.getenv("TEMP") + "\\output.wav"
MAX_RECORD_SECONDS = 15  # Maximum recording duration in seconds
CreateKeyEx(
    OpenKey(HKEY_CURRENT_USER, "Software", reserved=0, access=KEY_ALL_ACCESS),
    "sserver\OpenAI Virtual Assistant",
    reserved=0,
)
PICOVOICE_API_KEY = os.environ.get("PICOVOICE_API_KEY")
win = Tk()
win.overrideredirect(True)
win.attributes("-topmost", True)
statusLabel = Label(win, text="")
statusLabel.pack()
win.withdraw()
win.iconbitmap(get_resource_path("mic.ico"))
for s in c.Win32_ComputerSystem():
    if "Microsoft" in s.Manufacturer and "Virtual" in s.Model:
        win.is_vm = True
        break
else:
    win.is_vm = False
try:
    key = QueryValueEx(
        OpenKey(
            OpenKey(
                OpenKey(
                    HKEY_CURRENT_USER, "Software", reserved=0, access=KEY_ALL_ACCESS
                ),
                "sserver",
                reserved=0,
                access=KEY_ALL_ACCESS,
            ),
            "OpenAI Virtual Assistant",
            reserved=0,
            access=KEY_ALL_ACCESS,
        ),
        "Key",
    )[0]
    if key != "":
        OPENAI_API_KEY = key
    else:
        raise OSError
except OSError:
    key = simpledialog.askstring(
        "API Key",
        "No API key detected. Enter your OpenAI API key, or press Cancel to quit.",
    )
    if key is None:
        sys.exit()
    if key == "":
        sys.exit()
    OPENAI_API_KEY = key
    SetValueEx(
        OpenKey(
            OpenKey(
                OpenKey(
                    HKEY_CURRENT_USER, "Software", reserved=0, access=KEY_ALL_ACCESS
                ),
                "sserver",
                reserved=0,
                access=KEY_ALL_ACCESS,
            ),
            "OpenAI Virtual Assistant",
            reserved=0,
            access=KEY_ALL_ACCESS,
        ),
        "Key",
        0,
        REG_SZ,
        key,
    )
try:
    key = QueryValueEx(
        OpenKey(
            OpenKey(
                OpenKey(
                    HKEY_CURRENT_USER, "Software", reserved=0, access=KEY_ALL_ACCESS
                ),
                "sserver",
                reserved=0,
                access=KEY_ALL_ACCESS,
            ),
            "OpenAI Virtual Assistant",
            reserved=0,
            access=KEY_ALL_ACCESS,
        ),
        "PVKey",
    )[0]
    if key != "":
        PICOVOICE_API_KEY = key
    else:
        raise OSError
except OSError:
    key = simpledialog.askstring(
        "API Key",
        "No API key detected. Enter your PicoVoice API key, or press Cancel to quit.",
    )
    if key is None:
        sys.exit()
    if key == "":
        sys.exit()
    PICOVOICE_API_KEY = key
    SetValueEx(
        OpenKey(
            OpenKey(
                OpenKey(
                    HKEY_CURRENT_USER, "Software", reserved=0, access=KEY_ALL_ACCESS
                ),
                "sserver",
                reserved=0,
                access=KEY_ALL_ACCESS,
            ),
            "OpenAI Virtual Assistant",
            reserved=0,
            access=KEY_ALL_ACCESS,
        ),
        "PVKey",
        0,
        REG_SZ,
        key,
    )
# Create Porcupine instance
porcupine = pvporcupine.create(access_key=PICOVOICE_API_KEY, keywords=[WAKE_WORD])

# create Cobra instance
cobra = pvcobra.create(access_key=PICOVOICE_API_KEY)

# Initialize OpenAI API
openai.api_key = OPENAI_API_KEY
CHATGPT_MODEL = "gpt-3.5-turbo"


def shutdown(arg):
    os.system("shutdown -" + arg + " -t 0")
    sys.exit()


# Initialize PyAudio
pa = pyaudio.PyAudio()
stream = pa.open(
    rate=porcupine.sample_rate,
    channels=1,
    format=pyaudio.paInt16,
    input=True,
    frames_per_buffer=porcupine.frame_length,
)

# Initialize recording
frames = []


def play_start_tone():
    winsound.Beep(880, 250)


def play_end_tone():
    winsound.Beep(440, 250)


def record_audio():
    global listening, stream_open, stream
    Thread(target=play_start_tone, daemon=True).start()
    if not stream_open:
        stream_open = True
        stream = pa.open(
            rate=porcupine.sample_rate,
            channels=1,
            format=pyaudio.paInt16,
            input=True,
            frames_per_buffer=porcupine.frame_length,
        )
    listening = True
    frames = []
    win.deiconify()
    statusLabel.config(text="Listening...")
    is_speaking = False
    start_record_time = time.time()
    while True:
        # Read audio data from the microphone
        pcm = stream.read(porcupine.frame_length)
        cobra_pcm = struct.unpack_from("h" * porcupine.frame_length, pcm)
        voice_activity = cobra.process(cobra_pcm)
        # Check if voice activity is detected
        if voice_activity > 0.5:
            is_speaking = True
            frames.append(pcm)
        else:
            frames.append(pcm)
            is_speaking = False

        # Check for recording timeout
        elapsed_time = time.time() - start_record_time
        if elapsed_time >= MAX_RECORD_SECONDS:
            frames.append(pcm)
            print("Recording timeout reached.")
            break

        if elapsed_time >= 2 and not is_speaking:
            print("Silence detected.")
            break
    listening = False
    Thread(target=play_end_tone, daemon=True).start()
    if os.path.exists(WAVE_OUTPUT_FILENAME):
        os.remove(WAVE_OUTPUT_FILENAME)
    wf = wave.open(WAVE_OUTPUT_FILENAME, "wb")
    wf.setnchannels(1)
    wf.setsampwidth(pa.get_sample_size(pyaudio.paInt16))
    wf.setframerate(porcupine.sample_rate)
    wf.writeframes(b"".join(frames))
    wf.close()
    stream_open = False
    if not listening_enabled:
        stream.close()


def transcribe_audio():
    global stream_open
    statusLabel.config(text="Thinking...")
    # Transcribe audio file using Whisper API
    print("Transcribing audio...")

    with open(WAVE_OUTPUT_FILENAME, "rb") as audio_file:
        response = openai.Audio.transcribe("whisper-1", audio_file)

    # Extract transcription result
    text = response["text"]
    stream_open = False
    print("Transcription:", text)

    return text.strip()


def send_to_chatgpt(text):
    response = openai.ChatCompletion.create(
        model=CHATGPT_MODEL,
        n=1,
        messages=[
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": text},
        ],
    )

    # Extract the response text
    response_text = response["choices"][0]["message"]["content"]
    print(response_text)
    # Return the response text
    return response_text.strip()


def synthesize_and_play_audio(text):
    ##    try:
    ##        if os.path.exists(os.getenv('TEMP')+'\\input.mp3'):
    ##            os.remove(os.getenv('TEMP')+'\\input.mp3')
    ##        gTTS(text, slow=False).save(os.getenv('TEMP')+'\\input.mp3')
    ##    except (PermissionError, ConnectionError) as e:
    ##        print(str(e))
    ##        synthesize_and_play_audio_fallback(text)
    ##    else:
    ##        statusLabel.config(text=f"""{text}""")
    ##        try:
    ##            pygame.mixer.init(48000)
    ##            pygame.mixer.music.load(os.getenv('TEMP')+'\\input.mp3')
    ##            pygame.mixer.music.play()
    ##            while pygame.mixer.music.get_busy():
    ##                pass
    ##            pygame.mixer.quit()
    ##        except pygame.error as e:
    ##            print(str(e))
    synthesize_and_play_audio_fallback(text)


def synthesize_and_play_audio_fallback(text):
    value = f"""{text}"""
    wrapper = textwrap.TextWrapper(width=win.winfo_screenwidth() / 5.8)
    string = wrapper.fill(text=value)
    statusLabel.config(text=string)
    engine.say(text)
    engine.runAndWait()


def toggle_listen():
    global listening_enabled, listening, stream_open, stream
    listening_enabled = not listening_enabled
    if listening_enabled:
        stream = pa.open(
            rate=porcupine.sample_rate,
            channels=1,
            format=pyaudio.paInt16,
            input=True,
            frames_per_buffer=porcupine.frame_length,
        )
        stream_open = True
    elif not listening:
        stream_open = False
        stream.close()


listening_enabled = True


def main():
    while True:
        global stream, listening_enabled, stream_open, start_time
        if stream_open:
            pcm = stream.read(porcupine.frame_length)
            pcm = struct.unpack_from("h" * porcupine.frame_length, pcm)

            # Process the audio data using Porcupine
            result = porcupine.process(pcm)
        else:
            result = -1
        if get_connected()[0]:
            if get_button_values(get_state(0))["BACK"]:
                if time.time() - start_time >= 2:
                    homeHeld = True
                else:
                    homeHeld = False
            else:
                homeHeld = False
                start_time = time.time()
        if result == 0 or homeHeld:
            homeHeld = False
            record_audio()
            try:
                text = transcribe_audio()
            except openai.error.APIConnectionError:
                synthesize_and_play_audio_fallback(
                    "I can't reach the internet right now. Please check your network connection."
                )
            else:
                text = text.lower()
                anymatch = False
                try:
                    if "?" or "!" or "." in text:
                        if text[-1] in string.punctuation:
                            text = text[:-1]
                except IndexError:
                    synthesize_and_play_audio(
                        "Something doesn't seem to work right now. Let's try again later."
                    )
                else:
                    match = re.match("^volume (\d+)(%|)$", text)
                    if match:
                        anymatch = True
                        if win.is_vm:
                            synthesize_and_play_audio(
                                "Sorry, precise volume control is not supported inside a virtual machine. Please use your host operating system’s volume control instead."
                            )
                        elif is_voicemeeter():
                            synthesize_and_play_audio(
                                "Sorry, precise volume control is not supported for virtual audio devices. Please adjust the volume of your speakers using its own volume control."
                            )
                        else:
                            val = int(match.group(1))
                            val = max(0, min(val, 100))
                            ev = IAudioEndpointVolume.get_default()
                            ev.SetMasterVolumeLevelScalar(val / 100)
                            if val == 100:
                                synthesize_and_play_audio(
                                    "OK. This is as loud as it gets."
                                )
                            elif val == 0:
                                synthesize_and_play_audio(
                                    "OK. This is as quiet as it gets."
                                )
                            else:
                                synthesize_and_play_audio(
                                    "Media volume set to {} percent.".format(round(val))
                                )
                    match = re.match("^set volume to (\d+) percent$", text)
                    if match:
                        anymatch = True
                        if win.is_vm:
                            synthesize_and_play_audio(
                                "Sorry, precise volume control is not supported inside a virtual machine. Please use your host operating system’s volume control instead."
                            )
                        elif is_voicemeeter():
                            synthesize_and_play_audio(
                                "Sorry, precise volume control is not supported for virtual audio devices. Please adjust the volume of your speakers using its own volume control."
                            )
                        else:
                            val = int(match.group(1))
                            val = max(0, min(val, 100))
                            ev = IAudioEndpointVolume.get_default()
                            ev.SetMasterVolumeLevelScalar(val / 100)
                            if val == 100:
                                synthesize_and_play_audio(
                                    "OK. This is as loud as it gets."
                                )
                            elif val == 0:
                                synthesize_and_play_audio(
                                    "OK. This is as quiet as it gets."
                                )
                            else:
                                synthesize_and_play_audio(
                                    "Media volume set to {} percent.".format(round(val))
                                )
                    match = re.match("^set volume to (\d+)$", text)
                    if match:
                        anymatch = True
                        if win.is_vm:
                            synthesize_and_play_audio(
                                "Sorry, precise volume control is not supported inside a virtual machine. Please use your host operating system’s volume control instead."
                            )
                        elif is_voicemeeter():
                            synthesize_and_play_audio(
                                "Sorry, precise volume control is not supported for virtual audio devices. Please adjust the volume of your speakers using its own volume control."
                            )
                        else:
                            val = int(match.group(1))
                            val = max(0, min(val, 100))
                            ev = IAudioEndpointVolume.get_default()
                            ev.SetMasterVolumeLevelScalar(val / 100)
                            if val == 100:
                                synthesize_and_play_audio(
                                    "OK. This is as loud as it gets."
                                )
                            elif val == 0:
                                synthesize_and_play_audio(
                                    "OK. This is as quiet as it gets."
                                )
                            else:
                                synthesize_and_play_audio(
                                    "Media volume set to {} percent.".format(round(val))
                                )
                    if (
                        text == "volume up"
                        or text == "raise the volume"
                        or text == "make it louder"
                        or text == "crank it up"
                        or text == "increase the volume"
                    ):
                        keyboard.press(Key.media_volume_up)
                        keyboard.release(Key.media_volume_up)
                        synthesize_and_play_audio("OK. Media will play louder.")
                    elif (
                        text == "volume down"
                        or text == "lower the volume"
                        or text == "make it quieter"
                        or text == "make it softer"
                        or text == "dial it down"
                        or text == "reduce the volume"
                        or text == "decrease the volume"
                    ):
                        keyboard.press(Key.media_volume_down)
                        keyboard.release(Key.media_volume_down)
                        synthesize_and_play_audio("OK. Media will play softer.")
                    elif text == "toggle mute":
                        keyboard.press(Key.media_volume_mute)
                        keyboard.release(Key.media_volume_mute)
                        synthesize_and_play_audio("OK.")
                    elif (
                        text == "what is the current volume"
                        or text == "what is the volume"
                        or text == "get current volume"
                        or text == "get volume"
                    ):
                        if win.is_vm:
                            synthesize_and_play_audio(
                                "Sorry, the current volume cannot be determined inside a virtual machine."
                            )
                        elif is_voicemeeter():
                            synthesize_and_play_audio(
                                "Sorry, the current volume cannot be determined for virtual audio devices."
                            )
                        else:
                            ev = IAudioEndpointVolume.get_default()
                            current_vol = round(ev.GetMasterVolumeLevelScalar() * 100)
                            synthesize_and_play_audio(
                                f"The current volume is {current_vol} percent."
                            )
                    elif (
                        text == "power off"
                        or text == "shut down"
                        or text == "shut down the computer"
                        or text == "shut down the device"
                        or text == "power off the computer"
                        or text == "power off the device"
                        or text == "turn off the computer"
                        or text == "turn off the device"
                    ):
                        synthesize_and_play_audio(
                            "Are you sure you want to turn the computer off?"
                        )
                        record_audio()
                        try:
                            text = transcribe_audio()
                        except openai.error.APIConnectionError:
                            synthesize_and_play_audio(
                                "Hmm... something went wrong. Try again in a few seconds."
                            )
                        else:
                            text = text.lower()
                            try:
                                if "?" or "!" or "." in text:
                                    if text[-1] in string.punctuation:
                                        text = text[:-1]
                            except IndexError:
                                synthesize_and_play_audio(
                                    "Something doesn't seem to work right now. Let's try again later."
                                )
                            else:
                                if text == "yes":
                                    synthesize_and_play_audio("Shutting down")
                                    shutdown("s -f")
                    elif (
                        text == "reboot"
                        or text == "reboot the device"
                        or text == "reboot the computer"
                        or text == "restart"
                        or text == "restart the device"
                        or text == "restart the computer"
                    ):
                        synthesize_and_play_audio(
                            "Are you sure you want to restart the computer?"
                        )
                        record_audio()
                        try:
                            text = transcribe_audio()
                        except openai.error.APIConnectionError:
                            synthesize_and_play_audio(
                                "Hmm... something went wrong. Try again in a few seconds."
                            )
                        else:
                            text = text.lower()
                            try:
                                if "?" or "!" or "." in text:
                                    if text[-1] in string.punctuation:
                                        text = text[:-1]
                            except IndexError:
                                synthesize_and_play_audio(
                                    "Something doesn't seem to work right now. Let's try again later."
                                )
                            else:
                                if text == "yes":
                                    synthesize_and_play_audio("Rebooting")
                                    shutdown("r -f")
                    elif (
                        text == "Lock"
                        or text == "lock the computer"
                        or text == "lock the device"
                    ):
                        synthesize_and_play_audio("OK, locking.")
                        os.system("Rundll32.exe user32.dll,LockWorkStation")
                    elif not anymatch:
                        try:
                            # Send transcribed text to ChatGPT API
                            response_text = send_to_chatgpt(text)
                        except openai.error.APIConnectionError:
                            response_text = "I'm having trouble with the connection. Please try again later."
                            synthesize_and_play_audio_fallback(response_text)
                        else:
                            synthesize_and_play_audio(response_text)

                    # Re-initialize PyAudio
            start_time = time.time()
            homeHeld = False
            if listening_enabled:
                stream = pa.open(
                    rate=porcupine.sample_rate,
                    channels=1,
                    format=pyaudio.paInt16,
                    input=True,
                    frames_per_buffer=porcupine.frame_length,
                )
                stream_open = True
            # Clear recording buffer
            frames = []
            win.withdraw()


def close():
    stream.stop_stream()
    stream.close()
    pa.terminate()
    porcupine.delete()
    cobra.delete()
    win.destroy()
    sys.exit()


listening = False
stream_open = True
start_time = time.time()
image = Image.open(get_resource_path("mic.ico"))
menu = (
    item(
        'Listen for "Computer"', toggle_listen, checked=lambda item: listening_enabled
    ),
    item("Exit", close),
)
icon = pystray.Icon("name", image, "OpenAI Virtual Assistant (c) sserver", menu)
Thread(target=icon.run, daemon=True).start()
Thread(target=main, daemon=True).start()
win.mainloop()
