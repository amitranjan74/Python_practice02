import sounddevice

from scipy.io.wavfile import write

## sample recording
fs = 44100

#ask to enter recording time
second = int(input("Enter the recording time in seconds: "))

print("Recording....\n")

record_voice = sounddevice.rec(int(second * fs), samplerate = fs, channels = 2)
sounddevice.wait()

write("myrecording.wav", fs, record_voice)
print("Recording is done, please check your folder to listen")