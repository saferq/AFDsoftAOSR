from docxtpl import DocxTemplate, RichText
from res import afd_sqlite
import datetime
import re


class DocxT():
	""" Заполнение шаблонов """
	def __init__(self):
		self.db = afd_sqlite.Work_with_datebase()
		self.dtime = datetime.datetime.now()
	
	def get_value_object(self, column):
		query = f"SELECT {column} FROM Объект WHERE rowid==3;"
		val = self.db.db_query(query)
		return val[0][0]

	def get_value_vk(self, column, number):
		query = f"SELECT {column} FROM ВК WHERE number=='{number}';"
		val = self.db.db_query(query)
		return val[0][0]

	def get_value_fio(self, id_two, column, number):
		""" Добавить данные ответственных лиц \n
		id_two - выбрать ZK GP SKK AN PD SK \n
		column - выбрать fio fio_vsn fio_full"""
		# получение id1
		query = f"SELECT fiо FROM ВК WHERE number=='{number}';"
		id_one = self.db.db_query(query)[0][0]
		if column == 'fio_full':
			# получить фамилию
			query = f"SELECT {column} FROM ФИО WHERE id1=='{id_one}' AND id2=='{id_two}'"
			text = ''
			count = 1
			fio = self.db.db_query(query)
			for i in fio:
				if count == len(fio):
					text += f"{i[0]} \t"
				else:
					text += f"{i[0]} \t \n"
					count += 1
			return text
		else:
			# получить фамилию
			query = f"SELECT {column} FROM ФИО WHERE id1=='{id_one}' AND id2=='{id_two}'"
			query_date = f"SELECT date FROM ВК WHERE number=='{number}';"
			fio_date = self.db.db_query(query_date)[0][0]
			fio = self.db.db_query(query)
			text = ''
			count = 1
			fio = self.db.db_query(query)
			for i in fio:
				if count == len(fio):
					text += f"{i[0]} \t {fio_date}"
				else:
					text += f"{i[0]} \t {fio_date} \n"
					count += 1
			return text

	def vk_08_rd(self, path, number):
		""" заполнение """
		doc = DocxTemplate("tempel\\AVK_88_RD2.docx")
		context = {}
		doc.render({
			# Шапка
			"object": self.get_value_object('object'),
			"bulder": self.get_value_object('build_name'),
			"vsn_min": self.get_value_object('vsn_min'),
			"vsn_union": self.get_value_object('vsn_union'),
			"vsn_build": self.get_value_object('vsn_build'),
			"vsn_type": self.get_value_object('vsn_type'),
			# Заполнение
			"number": self.get_value_vk('number', number),
			"date": self.get_value_vk('date', number),
			"material": self.get_value_vk('material', number),
			"material_full": self.get_value_vk('material_full', number),
			"type_control": self.get_value_vk('type_control', number),
			"project": self.get_value_vk('project', number),
			"site": self.get_value_vk('site', number),
			"app": self.get_value_vk('app', number),
			"gost_tu": self.get_value_vk('gost_tu', number),
			"parameters": self.get_value_vk('parameters', number),
			"ovp1": self.get_value_vk('ovp1', number),
			"ovp2": self.get_value_vk('ovp2', number),
			# ФИО
			"fio_full_ZK": RichText(self.get_value_fio('ZK', 'fio_full', number)),
			"fio_full_SKK": RichText(self.get_value_fio('SKK', 'fio_full', number)),
			"fio_full_PD": RichText(self.get_value_fio('PD', 'fio_full', number)),
			"fio_full_SK": RichText(self.get_value_fio('SK', 'fio_full', number)),
			# 
			"fio_ZK": RichText(self.get_value_fio('ZK', 'fio', number)),
			"fio_SKK": RichText(self.get_value_fio('SKK', 'fio', number)),
			"fio_PD": RichText(self.get_value_fio('PD', 'fio', number)),
			"fio_SK": RichText(self.get_value_fio('SK', 'fio', number)),
			# 
			"fio_vsn_ZK": RichText(self.get_value_fio('ZK', 'fio_vsn', number)),
			"fio_vsn_SKK": RichText(self.get_value_fio('SKK', 'fio_vsn', number)),
			"fio_vsn_PD": RichText(self.get_value_fio('PD', 'fio_vsn', number)),
			"fio_vsn_SK": RichText(self.get_value_fio('SK', 'fio_vsn', number))
		})
		# now = self.dtime.strftime('%m%d')
		name = re.compile('[^a-zA-Zа-яА-Я0-9 ]').sub("", self.get_value_vk('material', number))
		print(f"{path}VK_{number}_{name}.docx")
		doc.save(f"{path}VK_{number}_{name}.docx")	
	
	def aosr_v2019(self, path, number, context, name_act):
		""" заполнение АОСР """
		doc = DocxTemplate("tempel\\AOSR_v2019.docx")
		doc.render(
			context
		)
		# now = self.dtime.strftime('%m%d-%H%M%S')
		name_act = re.compile('[^a-zA-Zа-яА-Я0-9, ]').sub("", name_act)
		name_act = name_act.replace(' ', '_').replace('__', '_')
		if len(name_act) > 50:
			print(f"{path}АОСР-{number}_{name_act[:30]}..{name_act[-17:-1]}.docx")
			doc.save(f"{path}АОСР-{number}_{name_act[:30]}..{name_act[-17:-1]}.docx")
		else:	
			print(f"{path}АОСР-{number}_{name_act[:50]}.docx")
			doc.save(f"{path}АОСР-{number}_{name_act[:50]}.docx")


if __name__ == "__main__":
	print('DocxT')
	x = DocxT()
	x.get_value_fio('SK', 'fio', '01.266-АР-12')
	x.get_value_fio('SK', 'fio_vsn', '01.266-АР-12')
	x.get_value_fio('SK', 'fio_full', '01.266-АР-12')
	path = 'files\\'
	# acts = ['01.266-АР-02.1', '01.266-АР-03.1', '01.266-АР-12']
	acts = ['АР-12']
	for act in acts:
		x.vk_08_rd(path, act)
		print(f"{path}VK_{act}.docx")

	