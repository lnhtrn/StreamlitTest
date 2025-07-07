import whisper

def load_whisper_model():
    return whisper.load_model("base")

whisper_model = load_whisper_model()

result_text = ''
result = whisper_model.transcribe("audio_test\Chase Propper - behavioral observation.m4a")
result_text += result['text']

result = whisper_model.transcribe("audio_test\Chase Propper - developmental history.m4a")
result_text += '\n\n ' + result['text']

print(result_text)