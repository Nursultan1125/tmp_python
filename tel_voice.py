import telebot
import speech_recognition as speech_recog


from pydub import AudioSegment

API_KEY = "1090642940:AAGVbbjPUJV5HpebsGthIbDlM1oVfQ8wY7U"

tb = telebot.TeleBot(API_KEY)


@tb.message_handler(content_types=['voice'])
def send_echo(message):
    voice = tb.get_file(message.voice.file_id)
    downloaded_file = tb.download_file(voice.file_path)
    with open('new_file.ogg', 'wb') as new_file:
        new_file.write(downloaded_file)
        new_file.close()

    AudioSegment.from_file("new_file.ogg").export("new_file.wav", format="wav")
    sample_audio = speech_recog.AudioFile('new_file.wav')
    with sample_audio as audio_file:
        recog = speech_recog.Recognizer()
        audio_content = recog.record(audio_file)
        a = recog.recognize_google(audio_content, language='ru-RU')
        tb.reply_to(message, str(a))


tb.polling(none_stop=True)
