import serial

COM_PORT = 'COM5'

def main():
    serial.Serial(COM_PORT, 19200, timeout=0.1).write(5)

if __name__ == '__main__':
    main()