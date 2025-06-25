import whisper

def load_whisper_model():
    return whisper.load_model("base")

whisper_model = load_whisper_model()

result_text = ''
result = whisper_model.transcribe("audio_test/rozanna_behave_obs.m4a")
result_text += result['text']

result = whisper_model.transcribe("audio_test/rozanna_behave_obs_2.m4a")
result_text += ' ' + result['text']

print(result_text)