# time module
import win32com.client
import random
print("-"*40)
print("     🔊 SHOUT-OUT VOICE ASSISTANT")
print("-"*40)
prompts = [
     "I'm ready! Please provide the names separated by commas.",
    "Let me know who you'd like me to announce.",
    "Who deserves a shout-out today?",
    "Enter the names, and I’ll take care of the rest.",
    "Provide the list of names, separated by commas.",
]


speaker = win32com.client.Dispatch("SAPI.SpVoice")
voices = speaker.GetVoices()
 # Show available voices
for i in range(voices.Count):
    print(i, voices.Item(i).GetDescription())
try:
    choice = int(input("Select voice index: "))

    if 0 <= choice < voices.Count:
        speaker.Voice = voices.Item(choice)
    else:
        print(f"Invalid index! Choose between 0 and {voices.Count-1}")
        print("Using default voice")

except:
    print("Invalid input, using default voice")

speaker.Rate = 0
speaker.Volume = 100

while True:
    msg = random.choice(prompts)
    print(f"\n👉 {msg} (type 'exit' to quit): ", end="")
    speaker.Speak(msg)

    i = input()

    if i.lower().strip() == "exit":
        speaker.Speak("Have a nice day")
        break

    if not i.strip():
        print("Please enter valid names!!")
        speaker.Speak("Please enter valid names!!")
        continue

    names = i.split(",")

    for name in names:
        name = name.strip()
        print(f"Shout-out to {name}")
        speaker.Speak(f"Shout out to... {name}")