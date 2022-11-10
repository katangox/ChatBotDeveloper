import telebot, xlsxwriter, sqlalchemy, datetime, random
from sqlalchemy.orm import mapper
from sqlalchemy.orm import sessionmaker

engine_users = sqlalchemy.create_engine('sqlite:///users.db', echo=False, pool_recycle=7200)
engine_roles = sqlalchemy.create_engine('sqlite:///roles.db', echo=False, pool_recycle=7200)

metadata = sqlalchemy.MetaData()
users_table = sqlalchemy.Table('users', metadata,
				sqlalchemy.Column('name', sqlalchemy.String, primary_key=True),
				sqlalchemy.Column('birthday', sqlalchemy.String),
				sqlalchemy.Column('role', sqlalchemy.String))

roles_table = sqlalchemy.Table('roles', metadata,
				sqlalchemy.Column('id', sqlalchemy.Integer, primary_key=True),
				sqlalchemy.Column('name', sqlalchemy.String))


class User(object):
	def __init__(self, name, birthday, role):
		self.name = name
		self.birthday = birthday
		self.role = role


class Role(object):
	def __init__(self, id, name):
		self.id = id
		self.name = name


mapper(User, users_table)
mapper(Role, roles_table)
metadata.create_all(engine_users)
metadata.create_all(engine_roles)

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
		Session_roles = sessionmaker(bind=engine_roles)
		session_roles = Session_roles()
		query = session_roles.query(Role)
		query_list = list(query)

		user = User(message.text, datetime.datetime(random.randint(1960, 1998), random.randint(1, 12),
													random.randint(1, 30)).strftime("%Y.%m.%d"),
					query_list[random.randint(0, 1)].name)

		session_roles.close()
		query_list = None

		Session = sessionmaker(bind=engine_users)
		session = Session()
		session.add(user)
		session.commit()

		query = session.query(User)
		query_list = list(query)

		workbook = xlsxwriter.Workbook('users.xlsx')
		worksheet = workbook.add_worksheet()
		worksheet.set_column(0, 2, 30)
		worksheet.write('A1', 'ФИО')
		worksheet.write('B1', 'Дата рождения')
		worksheet.write('C1', 'Наименование роли')

		expenses = []
		for item in query_list[:-6:-1]:
			expenses.append([item.name, item.birthday, item.role])

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