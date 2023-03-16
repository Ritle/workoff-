import telebot
import create_new_list

token='6282860384:AAHlneTWcQHjv4ZAA6UIH9JnsFX3sousr_g'

bot = telebot.TeleBot(token)
sid = ""

class englishBot:
    def __init__(self, bot):
        self.bot = bot
        

    def send_msg(self, chat_id, msg_text):
        self.bot.send_message(chat_id, msg_text)

    def start(self):
        print("Запускаем бота")
        self.bot.polling(none_stop=True)
      
my_bot = englishBot(bot) 


@bot.message_handler(commands=['start'])
def get_sid(message):
    
    my_bot.send_msg(message.chat.id,"Введите sid")
    

@bot.message_handler(content_types=["text"])
def check_sid(message):
    if len(message.text) == 39:
        my_bot.send_msg(message.chat.id,"Начинаю работу")
        create_new_list.start(message.text)
        my_bot.send_msg(message.chat.id,"Все готово")

    else:
        my_bot.send_msg(message.chat.id,"Некорректный sid. Введите sid заново")


my_bot.start()