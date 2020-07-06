import sys
import time
import serial

COM_PORT = 'COM6'
LIN_NAD = 0xa0
TESTER_PRESENT_PERIODE = 0.5  # time in seconds


class T_heikostick:

    def __init__(self, COM_port_name):
        self._COM = serial.Serial(COM_port_name, 19200, timeout=0.1)
        print(self._COM)
        return

    def _add_parity(self,PID):
        p0 = ((PID>>0) ^ (PID>>1) ^ (PID>>2) ^ (PID>>4)) & 1
        p1 =  (~((PID>>1) ^ (PID>>3) ^ (PID>>4) ^ (PID>>5))) & 1
        return (PID & 0x3f) | (p0<<6) | (p1<<7)

    def _t_and_r(self, txb):
        xorsum = 0
        for i in range(1, txb[1]+1):
            xorsum ^= txb[i]
        txb += [xorsum, 0x03]
        print('sending  (',len(txb),'):',[hex(x) for x in txb])
        self._COM.write(txb)
        rxb = list(self._COM.read(2))
        if len(rxb) == 0: raise Exception('Timeout, no answer!')
        if len(rxb) != 2: raise Exception('Not enough bytes received!')
        if rxb[0] != 0x02: raise Exception('Invalid start byte!')
        rxb += list(self._COM.read(rxb[1] + 1))
        #print(' received(', len(rxb), '):', [hex(x) for x in rxb])
        if len(rxb) != (rxb[1] + 3): raise Exception('Not enough bytes received!')
        xorsum = rxb[1]
        for i in range(2, rxb[1]+1):
            xorsum ^= rxb[i]
        if xorsum != rxb[rxb[1]+1]: raise Exception('Received invalid xorsum!')
        if rxb[rxb[1]+2] != 0x03: raise Exception('Received end byte!')
        return rxb[5:len(rxb)-3]

    def send_request(self, frame, PID=0x3c):
        dlc = len(frame)
        txb = [0x02, 4+dlc+1, ord('M'), self._add_parity(PID), dlc]
        txb += frame
        chksum = frame[0]
        for i in range(1, dlc):
            chksum += frame[i]
            if chksum > 255: chksum -= 255
        chksum ^= 255
        #print('frame:', [hex(x) for x in frame], ' chksum:', hex(chksum))
        txb += [chksum]
        return self._t_and_r(txb)

    def get_answer(self, expected=8, PID=0x3d):
        return self._t_and_r([0x02, 4, ord('S'), self._add_parity(PID), expected])

    def close(self):
        self._COM.close()
        return

    pass  # end of T_heikostick
################################################################################################


def crc_CCIT16(datablock, crc=0xffff):
    for b in datablock:
        for _ in range(8):
            if (crc^b)&1:crc=(crc>>1)^0x8408
            else:crc>>=1
            b>>=1
    return crc


def send_Tester_present():
    #global LIN
    LIN.send_request([ LIN_NAD, 2, 0x3e, 0x80, 0xff, 0xff, 0xff, 0xff ])  # tester present, no answer expected
    return

################################################################################################

def requestSolar1():
    print("Request Solar channel 1:")
    ans = LIN.send_request([LIN_NAD, 0x3, 0x22, 0xf1, 0xf5, 0xff, 0xff, 0xff])
    print()

    print("Response:")
    ans1 = LIN.get_answer(8, 0x3D)
    # check for error (if all bytes are 0xFF)
    if all(b == 0xFF for b in ans1):
        raise Exception("Unable to read current from solar channel 1.")

    ans2 = LIN.get_answer(8, 0x3D)

    data = ans1[6:8] + ans2[2:4]

    value = 0
    for byte in data:
        value <<= 8
        value += byte

    #print([hex(i) for i in ans1])
    #print([hex(i) for i in ans2])
    #print(data)
    return value


def requestSolar2():
    print("Request Solar channel 2:")
    ans = LIN.send_request([LIN_NAD, 0x3, 0x22, 0xf1, 0xf6, 0xff, 0xff, 0xff])
    print()

    print("Response:")
    ans1 = LIN.get_answer(8, 0x3D)
    # check for error (if all bytes are 0xFF)
    if all(b == 0xFF for b in ans1):
        raise Exception("Unable to read current from solar channel 2.")
    ans2 = LIN.get_answer(8, 0x3D)
    data = ans1[6:8] + ans2[2:4]

    value = 0
    for byte in data:
        value <<= 8
        value += byte

    #print([hex(i) for i in ans1])
    #print([hex(i) for i in ans2])
    #print(data)
    return value


def writeSolarCoeff2NVM(solar_L, solar_R):

    # build the UDS frame
    payload = [0x2e, 0xf1, 0xf8]
    payload += [solar_L & 0xFF, (solar_L >> 8) & 0xFF, (solar_L >> 16) & 0xFF, (solar_L >> 24) & 0xFF, (solar_L >> 32) & 0xFF, (solar_L >> 40) & 0xFF, (solar_L >> 48) & 0xFF, (solar_L >> 56) & 0xFF]		# solar channel L
    payload += [solar_R & 0xFF, (solar_R >> 8) & 0xFF, (solar_R >> 16) & 0xFF, (solar_R >> 24) & 0xFF, (solar_R >> 32) & 0xFF, (solar_R >> 40) & 0xFF, (solar_R >> 48) & 0xFF, (solar_R >> 56) & 0xFF]		# solar channel R
    crc = crc_CCIT16(payload[3:])
    payload += [crc & 255, crc >> 8]
    print('Total', len(payload), 'bytes:', [hex(x) for x in payload])
    #print('CRC:', hex(crc))

    # build the first frame_8B
    print(' ')
    print("Writing coefficients to NVM - 1st frame")
    print(' ')
    frame_8B = [LIN_NAD, 0x10 | (len(payload) >> 8), len(payload) & 0xff]
    cf_offs = 5 	#consecutive frame offset
    frame_8B += payload[0:cf_offs]
    cf_seq_num = 0  #consecutive frame sequence number
    LIN.send_request(frame_8B)
    while cf_offs < len(payload):
        print(' ')
        print("Writing coefficients to NVM - n frame")
        print(' ')
        cf_seq_num = (cf_seq_num + 1) & 0xFF # masking with 1 byte
        frame_8B = [LIN_NAD, 0x20 | cf_seq_num] # sending A0 2x, x = consecutive frame number
        frame_8B += payload[cf_offs:cf_offs + 6] # sending the payload of 6 bytes
        cf_offs += 6
        while len(frame_8B) < 8: frame_8B.append(0xff)
        LIN.send_request(frame_8B)
    print(' ')
    print("Writing coefficients to NVM - response")
    frame_8B = LIN.get_answer()
    if len(frame_8B) < 3: raise Exception('LIN:no answer!')

    # sent it through LIN
    return


def main():
    
    if((str(sys.argv[1])) == 'help'):
        print(" ------ HELP ------")
        print("excepted arguments are:")
        print("1: Calibration value in the Right Channel")
        print("2: Calibration value in the Left Channel")
        sys.exit(1)
	
    if len(sys.argv) != (2 + 1): # maximum number of arguments is 3
        print("arguments are not correct")
        print("excepted arguments are:")
        print("1: Calibration value in the Right Channel")
        print("2: Calibration value in the Left Channel")
        sys.exit(1)

    # wake up the sensor
    global LIN
    LIN = T_heikostick(COM_PORT)

    send_Tester_present()
    print()
    time.sleep(1)		# wait until it boots up

    writeSolarCoeff2NVM(int(sys.argv[2]), int(sys.argv[1]))

    # we are done - close the connection
    LIN.close()
    print()
    print('Calibration successful.')


if __name__ == '__main__':
    main()
