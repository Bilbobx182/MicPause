import pyaudio
import struct
import math
import win32com.client

INITIAL_TAP_THRESHOLD = 0.070
FORMAT = pyaudio.paInt16
SHORT_NORMALIZE = (1.0/32768.0)
CHANNELS = 2
RATE = 44100
INPUT_BLOCK_TIME = 0.05
INPUT_FRAMES_PER_BLOCK = int(RATE*INPUT_BLOCK_TIME)
errorcount=0

MAX_TAP_BLOCKS = 0.15/INPUT_BLOCK_TIME


def get_rms(block):

    count = len(block)/2
    format = "%dh"%(count)
    shorts = struct.unpack(format, block)

    sum_squares = 0.0
    for sample in shorts:
        n = sample * SHORT_NORMALIZE
        sum_squares += n*n

    return math.sqrt(sum_squares / count)

pa = pyaudio.PyAudio()
stream = pa.open(format=FORMAT, channels=CHANNELS, rate=RATE, input=True, frames_per_buffer=INPUT_FRAMES_PER_BLOCK)

tap_threshold = INITIAL_TAP_THRESHOLD
while True:
    try:
        block = stream.read(INPUT_FRAMES_PER_BLOCK)
    except IOError as e:
        errorcount += 1
        print("(%d) ERROR IN INPUT: %s" % (errorcount, e))

    amplitude = get_rms(block)
    if amplitude > tap_threshold:
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys("^a")
        shell.SendKeys(" ")
