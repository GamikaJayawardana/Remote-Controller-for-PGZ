from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
import serial.tools.list_ports
import pyautogui
import serial
import time

# Get default audio device (speakers/headphones)
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

# --- Set volume to 50% (0.5) ---
volume.SetMasterVolumeLevelScalar(0.5, None)

time.sleep(2)

mode = 0
sePort = None
deviceId = "#ha6kl"
devPort = None
portList = []

buttonPreviousState = None
buttonPreviousValue = None

def action(buttonAction):

    match buttonAction:
        case 'M':
            pyautogui.press('space')

        case 'L':
            pyautogui.press('left')

        case 'R':
            pyautogui.press('right')

        case 'U':
            muteVolume()

        case 'D':
            unMuteVolume()

    setVolume()

def getButtonAction(receive):

    if len(receive) == 5:
        global buttonPreviousState, buttonPreviousValue

        if receive[0] == 'B' and receive[2] == 'S':
            if buttonPreviousValue != receive[1]:
                buttonPreviousValue = receive[1]
                return receive[1]
            elif buttonPreviousState != receive[3]:
                buttonPreviousState = receive[3]
                return receive[1]
        else:
            return
        
    else:
        return

def setVolume():
    target_volume = (ord(receive[4]) - 33) / 20.0
    target_volume = max(0.0, min(1.0, target_volume))  # Clamp
    current_volume = volume.GetMasterVolumeLevelScalar()
    if abs(current_volume - target_volume) > 0.01:
        volume.SetMasterVolumeLevelScalar(target_volume, None)

def muteVolume():
    volume.SetMute(1, None)

def unMuteVolume():
    volume.SetMute(0, None)

while True:
    if mode == 0:
        volume.SetMasterVolumeLevelScalar(0.5, None)
        print(">> Initiating device search...")
        sePort = None
        devPort = None
        mode = 1
        receive = ""

    elif mode == 1:  # Get Port List
        portList = serial.tools.list_ports.comports()
        if portList:
            mode = 2

    elif mode == 2:  # Find The Device
        try:
            for devPort in portList:
                ser = serial.Serial(port=devPort.device, baudrate=9600,
                                    bytesize=8, timeout=1, stopbits=serial.STOPBITS_ONE)
                counter = 0
                while counter < 100:
                    rece = ser.readline().decode('ascii', errors='ignore').strip()
                    if rece == deviceId:
                        mode = 3
                        print(">> Status: Device ID received successfully.")
                        ser.close()
                        break
                    else:
                        ser.write("idRequest\r\n".encode('ascii'))
                    time.sleep(0.01)
                    counter += 1
        except:
            mode = 1

    elif mode == 3:
        print(">> Status: Device connection successful.")
        try:
            sePort = serial.Serial(port=devPort.device, baudrate=9600,
                                   bytesize=8, timeout=1, stopbits=serial.STOPBITS_ONE)
            mode = 4
        except serial.SerialException as e:
            print(f">> Serial Connection Error: {e}")
            mode = 0

    elif mode == 4:
        print(">> Entered live communication mode.")
        last_data_time = time.time()
        try:
            while True:
                if sePort.in_waiting > 0:
                    receive = sePort.readline().decode('ascii', errors='ignore').strip()
                    if receive:
                        last_data_time = time.time()
                        buttonAction = getButtonAction(receive)

                        if buttonAction:
                            action(buttonAction)

                else:
                    if time.time() - last_data_time >= 2:
                        print(">> No data received in the last 2 seconds.")
                        sePort.close()
                        mode = 0
                        last_data_time = time.time()
                        break

        except KeyboardInterrupt:
            print("\nExiting live communication.")
            break
        finally:
            if sePort and sePort.is_open:
                sePort.close()
                print("Serial port closed.")
