import cv2
import cvzone
import numpy as np
from cvzone.FaceMeshModule import FaceMeshDetector
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

cap = cv2.VideoCapture(0)
detector = FaceMeshDetector(maxFaces=1)


devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
volume.GetMute()
volume.GetMasterVolumeLevel()
volRange = volume.GetVolumeRange()
minVol = volRange[0]
maxVol = volRange[1]

while True:
    Success, img = cap.read()
    img, faces = detector.findFaceMesh(img, draw=False)
    if faces:
        face = faces[0]
        pointLeft = face[145]
        pointRight = face[374]

        w, _ = detector.findDistance(pointLeft, pointRight)
        W = 6.3

        # Finding Distance
        f = 840
        d = (W*f)/w
       # print(d)

        # Face Range 50 - 150
        # Volume Range -65 - 0

        vol = np.interp(d, [50, 150], [minVol, maxVol])
        print(int(d), vol)
        volume.SetMasterVolumeLevel(vol, None)

        cvzone.putTextRect(img, f'Depth: {int(d)}cm', (face[10][0]-100, face[10][1]-50), scale=2)

    cv2.imshow("Images", img)
    cv2.waitKey(1)
