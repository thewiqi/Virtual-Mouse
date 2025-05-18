import cv2
import numpy as np
import pyautogui
import keyboard
from cvzone.HandTrackingModule import HandDetector
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
import screen_brightness_control as sbc
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# Webcam size
wCam, hCam = 640, 480

cap = cv2.VideoCapture(0)
cap.set(3, wCam)
cap.set(4, hCam)

# Hand detector
detector = HandDetector(detectionCon=1, maxHands=1)

# Screen size
wScr, hScr = pyautogui.size()

# Frame Reduction
frameR = 100

# Smoothening
smoothening = 7
plocX, plocY = 0, 0
clocX, clocY = 0, 0

# Toggle control
mouseControl = False

# Setup for Volume Control (pycaw)
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
volRange = volume.GetVolumeRange()
minVol = volRange[0]
maxVol = volRange[1]

# Brightness control range
minBright = 0
maxBright = 100

# Shift toggle management
shiftPressed = False

while True:
    # 1. Get image frame
    success, img = cap.read()
    img = cv2.flip(img, 1)

    # 2. Find hand landmarks
    img = detector.findHands(img)
    lmList, bbox = detector.findPosition(img)

    # 3. Toggle mouse control with Shift key
    if keyboard.is_pressed('shift') and not shiftPressed:
        mouseControl = not mouseControl
        shiftPressed = True
    if not keyboard.is_pressed('shift'):
        shiftPressed = False

    # 4. If hand landmarks detected
    if lmList:
        x1, y1 = lmList[8]   # Index finger tip
        x2, y2 = lmList[12]  # Middle finger tip
        x_thumb, y_thumb = lmList[4]  # Thumb tip

        # Fingers state
        fingers = detector.fingersUp()

        # 5. Volume Control: Thumb + Index
        if fingers[0] == 1 and fingers[1] == 1 and fingers[2] == 0:
            length, img, lineInfo = detector.findDistance(4, 8, img)
            vol = np.interp(length, [30, 200], [minVol, maxVol])
            volume.SetMasterVolumeLevel(vol, None)
            cv2.putText(img, f'Volume', (50, 100),
                        cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 0, 255), 3)
            cv2.circle(img, (lineInfo[4], lineInfo[5]), 15, (0, 255, 0), cv2.FILLED)

        # 6. Brightness Control: Thumb + Middle Finger
        if fingers[0] == 1 and fingers[1] == 0 and fingers[2] == 1:
            length, img, lineInfo = detector.findDistance(4, 12, img)
            bright = np.interp(length, [30, 200], [minBright, maxBright])
            sbc.set_brightness(int(bright))
            cv2.putText(img, f'Brightness', (50, 100),
                        cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 255), 3)
            cv2.circle(img, (lineInfo[4], lineInfo[5]), 15, (0, 255, 255), cv2.FILLED)

        # 7. Mouse control
        if mouseControl:
            # Only Index Finger: Moving Mode
            if fingers[1] == 1 and fingers[2] == 0:
                # Convert Coordinates
                x3 = np.interp(x1, (frameR, wCam-frameR), (0, wScr))
                y3 = np.interp(y1, (frameR, hCam-frameR), (0, hScr))

                # Smoothen Values
                clocX = plocX + (x3 - plocX) / smoothening
                clocY = plocY + (y3 - plocY) / smoothening

                # Move Mouse
                pyautogui.moveTo(clocX, clocY)
                plocX, plocY = clocX, clocY

                cv2.circle(img, (x1, y1), 15, (255, 0, 255), cv2.FILLED)

            # Both Index and Middle fingers are up: Clicking Mode
            if fingers[1] == 1 and fingers[2] == 1:
                length, img, _ = detector.findDistance(8, 12, img)
                if length < 40:
                    pyautogui.click()
                    cv2.circle(img, ((x1+x2)//2, (y1+y2)//2), 15, (0, 255, 0), cv2.FILLED)

    # 8. Display status
    if mouseControl:
        cv2.putText(img, "Mouse Mode: ON (Press Shift to Toggle)", (10, 50),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)
    else:
        cv2.putText(img, "Mouse Mode: OFF (Press Shift to Toggle)", (10, 50),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)

    # 9. Display
    cv2.imshow("Virtual Mouse + Volume + Brightness", img)
    if cv2.waitKey(1) == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()