import cv2
import numpy as np
import re
import gspread
import time

sa = gspread.service_account("nice.json")
sh = sa.open_by_key("1C0qA8xQ9ExL5EeL-7lQkxt4hUv1TEeVDT0RKnWPx3ao")
wks = sh.worksheet("Form Responses 1")

detector = cv2.QRCodeDetector()
cap = cv2.VideoCapture(0)

while True:
    ret, frame = cap.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    ret, binary = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY)
    binary = cv2.bitwise_not(binary)
    
    data, points, ok = detector.detectAndDecode(frame)
    
    if re.search(r"'key': '12137",data):        
        data = data.split("\\\\")
        for all in range(len(data)-1):
            data[all] = data[all].split(": ")[1][1:]
        data = data[:-2]
        # cell = wks.find(data[0])
        # print(data)
        cell_list = wks.findall(data[0])
        for cell in cell_list:
            if cell:
                row = cell.row
                col = chr(cell.col+64)
                if str(col) != 'B':
                    continue
                col2 = chr(cell.col+74)
                Color_cell = str(col) + str(row)
                Color_cell2 = str(col2) + str(row)
                c = "L" + str(cell.row)
                val = wks.acell(c).value
                if val is not None:
                    val = int(val)
                    if val <= len(data)-6:
                        continue
                adder = 70
                if val is None:
                    wks.update(c, '1')
                    c = "N" + str(cell.row)
                    tis = int(time.time())
                    wks.update(c, tis)
                    wks.format(Color_cell + ":" + Color_cell2, {'backgroundColor': {'red': 0.0 , 'green' : 1.0 , 'blue' : 0.0}})
                else:
                    if data[5] == "Team":
                        val = int(val)
                        if val <= len(data)-6:
                            if len(data) > 6 and len(data) <= 10:
                                tis = int(time.time())
                                if tis - int(wks.acell("N"+str(cell.row)).value) > 20:
                                    val+=1
                                    wks.update(c, str(val))
                                    c = "N" + str(cell.row)
                                    tis = int(time.time())
                                    wks.update(c, tis)
                        else:
                            wks.format(Color_cell + ":" + Color_cell2, {'backgroundColor': {'red': 1.0 , 'green' : 1.0 , 'blue' : 0.0}})
                    else:
                        tis = int(time.time())
                        if tis - int(wks.acell("N"+str(cell.row)).value) > 20:
                            wks.format(Color_cell + ":" + Color_cell2, {'backgroundColor': {'red': 1.0 , 'green' : 1.0 , 'blue' : 0.0}})                
                
    cv2.imshow("Webcam Stream", frame)
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
