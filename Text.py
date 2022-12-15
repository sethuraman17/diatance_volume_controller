import cv2
import cvzone
from cvzone.FaceMeshModule import FaceMeshDetector
import numpy as np

cap = cv2.VideoCapture(0)
detector = FaceMeshDetector(maxFaces=1)

sen = 10

textList = ["Please Check", "Your Distance", "From TV.", "Enjoy Your Games"]

while True:
    Success, img = cap.read()
    imgText = np.zeros_like(img)
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
        print(d)

        cvzone.putTextRect(img, f'Depth: {int(d)}cm', (face[10][0]-100, face[10][1]-50), scale=2)
        for i, text in enumerate(textList):
            singleHeight = 20 + int((int(d/sen)*sen)/4)
            scale = 0.4 + (int(d/sen)*sen)/75
            cv2.putText(imgText, text, (50, 50+(i*singleHeight)),
                        cv2.FONT_ITALIC, scale, (255, 255, 255), 2)

    imgStacked = cvzone.stackImages([img, imgText], 2, 1)
    cv2.imshow("Images", imgStacked)
    cv2.waitKey(1)
