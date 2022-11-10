import datetime

import telebot, xlsxwriter, sqlalchemy
from sqlalchemy.orm import mapper
from sqlalchemy.orm import sessionmaker

engine = sqlalchemy.create_engine('sqlite:///users.db', echo=False, pool_recycle=7200)

metadata = sqlalchemy.MetaData()
metadata.create_all(engine)
users_table = sqlalchemy.Table('users', metadata,
				sqlalchemy.Column('name', sqlalchemy.String, primary_key=True),
				sqlalchemy.Column('birthday', sqlalchemy.String),
				sqlalchemy.Column('role', sqlalchemy.String))


class User(object):
	def __init__(self, name, birthday, role):
		self.name = name
		self.birthday = birthday
		self.role = role


mapper(User, users_table)
metadata.create_all(engine)

bot = telebot.TeleBot("5755560169:AAGKcRL8Qc7QlVzFRIBrww-YNk46DjGXYvc")
fio = False


@bot.message_handler(commands=['start'])
def send_welcome(message):
	user_id = message.from_user.id
	markup = telebot.types.InlineKeyboardMarkup(row_width=1)
	item1 = telebot.types.InlineKeyboardButton("Записать работника", callback_data="button_1")
	markup.add(item1)
	bot.send_message(user_id, "Вы в главном меню!", reply_markup=markup)


@bot.callback_query_handler(func=lambda call:True)
def callback(call):
	if call.message:
		user_id = call.message.chat.id
		if call.data == "button_1":
			bot.send_message(user_id, "Введите ФИО работника")
			global fio
			fio = True


@bot.message_handler(content_types=['text'])
def msg(message):
	global fio
	if fio:
		fio = False
		user = User(message.text, datetime.datetime.today().strftime("%Y.%m.%d"), "543543gr")

		Session = sessionmaker(bind=engine)
		session = Session()
		session.add(user)
		session.commit()

		query = session.query(User)
		query_list = list(query)

		workbook = xlsxwriter.Workbook('users.xlsx')
		worksheet = workbook.add_worksheet()
		worksheet.write('A1', 'ФИО')
		worksheet.write('B1', 'Дата рождения')
		worksheet.write('C1', 'Наименование роли')

		expenses = []
		for item in query_list[::-1]:
			expenses.append([item.name, item.birthday, item.role])
			if len(expenses) == 5:
				break

		row = 1
		col = 0

		for a, b, c in (expenses):
			worksheet.write(row, col, a)
			worksheet.write(row, col + 1, b)
			worksheet.write(row, col + 2, c)
			row += 1

		workbook.close()
		session.close()

# Запуск бота
bot.infinity_polling()