import os

import telebot
from telebot import types

from data_management import find_client, save_ratings, load_data

token = '5395555491:AAEOZ4b7htMaX2keSFbsuQvSd7y2Se9LH6I'
bot = telebot.TeleBot(token, parse_mode=None)


@bot.message_handler(commands=['start'])
def start_message(message):
    bot.send_message(message.chat.id, "Привет ✌️ "
                                      "Чтобы узнать список всех команды пишите /help")


@bot.message_handler(commands=['help'])
def help_message(message):
    bot.send_message(message.chat.id, "/number 88005553535 - Найти клиента по номеру телефона\n"
                                      "/card 001 - Найти клиента по номеру карты\n"
                                      "/download - Сгрузить рейтинг вин")


def extract_arg(arg, message):
    try:
        return str(arg.split()[1:][0])
    except IndexError:
        bot.send_message(message.chat.id, 'Неверное значение',
                         parse_mode='Markdown')


def send_vines(message, all_data):
    bot.send_message(message.chat.id,
                     f"Имя: *{all_data[0]['name']}*\nНомер карты: *{all_data[0]['loyal_card_number']}*\nНомер телефона: *{all_data[0]['phone_number']}*",
                     parse_mode='Markdown')

    for element in all_data[:15]:
        if not element['articulate']:
            element['articulate'] = 'Отсутствует'
        markup_inline = types.InlineKeyboardMarkup()
        item_1 = types.InlineKeyboardButton(text='1', callback_data=f"{element['articulate']} 1")
        item_2 = types.InlineKeyboardButton(text='2', callback_data=f"{element['articulate']} 2")
        item_3 = types.InlineKeyboardButton(text='3', callback_data=f"{element['articulate']} 3")
        item_4 = types.InlineKeyboardButton(text='4', callback_data=f"{element['articulate']} 4")
        item_5 = types.InlineKeyboardButton(text='5', callback_data=f"{element['articulate']} 5")
        markup_inline.add(item_1, item_2, item_3, item_4, item_5)
        bot.send_message(message.chat.id,
                         f"Дата покупки: *{element['date']}*\nАртикул: *{element['articulate']}*\nНазвание вина: *{element['nomn']}*\nЦена: *{element['price']}*\nЦена со скидкий: *{element['price_with_discount']}*\n\nОцените вино:",
                         parse_mode='Markdown', reply_markup=markup_inline)


@bot.message_handler(commands=['number'])
def findbynumber(message):
    number = extract_arg(message.text, message)
    all_data = find_client(phone_number=number)
    if len(all_data) == 0:
        bot.send_message(message.chat.id, 'Ничего не найдено',
                         parse_mode='Markdown')
        return
    send_vines(message, all_data)


@bot.message_handler(commands=['card'])
def findbycard(message):
    card = extract_arg(message.text, message)
    all_data = find_client(card=card)
    if len(all_data) == 0:
        bot.send_message(message.chat.id, 'Ничего не найдено',
                         parse_mode='Markdown')
        return
    send_vines(message, all_data)


@bot.callback_query_handler(func=lambda call: True)
def callback_inline(call):
    art, rate = call.data.split()
    data = load_data()
    title = ''
    for e in data:
        if e['articulate'] == art:
            title = e['nomn']
    if art == 'Отсутствует':
        bot.answer_callback_query(callback_query_id=call.id, text=f'Вы не можете оставить рейтинг к этому товару!')
        return
    save_ratings(articulate=art, title=title, rate=rate)
    bot.answer_callback_query(callback_query_id=call.id, text=f'Вы оставили рейтинг {rate} к товару {art}!')


@bot.message_handler(content_types=['document'])
def upload_base(message):
    if not (str(message.document.file_name).endswith('.xls') or str(message.document.file_name).endswith('.xlsx')):
        bot.reply_to(message, "Неверный формат файла")
        return
    os.remove('all_data/data.xls')
    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)
    src_old = 'all_data/' + message.document.file_name
    src_new = 'all_data/' + 'data.xls'
    with open(src_old, 'wb') as new_file:
        new_file.write(downloaded_file)
    os.rename(src_old, src_new)
    bot.reply_to(message, "Новая база успешно установлена")


@bot.message_handler(commands=['download'])
def start_message(message):
    rating_xlsx = open('all_data/ratings.xlsx', 'rb')
    bot.send_document(message.chat.id, rating_xlsx)


def main():
    bot.infinity_polling()


if __name__ == '__main__':
    main()
