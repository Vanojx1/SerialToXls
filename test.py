from serial import Serial
import time

com = Serial('COM3')
com.write(b'@00EX001')
time.sleep(1)
com.write(b'72000250')
time.sleep(1)
com.write(b'06000120')
time.sleep(1)
com.write(b'05B*')
com.close()