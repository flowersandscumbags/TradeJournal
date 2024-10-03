import winsound
try:
    winsound.PlaySound("C:\\Windows\\Media\\chimes.wav", winsound.SND_FILENAME)

except RuntimeError as e:
    print(f"Error playing sound: {e}")

