import platform
import cv2
import time

c = cv2.VideoCapture(0, cv2.CAP_DSHOW if platform.system() == 'Windows' else 0)
print('opened', c.isOpened())
if c.isOpened():
    for i in range(3):
        start = time.time()
        ret, frame = c.read()
        print('read', i, 'ret=', ret, 'elapsed=', round(time.time() - start, 3), 'shape=', None if frame is None else frame.shape)
        if ret:
            cv2.imwrite('debug_frame.jpg', frame)
            break
    c.release()
