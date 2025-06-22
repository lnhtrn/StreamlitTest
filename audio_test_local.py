import whisper

def load_whisper_model():
    return whisper.load_model("base")

whisper_model = load_whisper_model()

result = whisper_model.transcribe("audio_test/testfile1.m4a")
print(result)