from docxtpl import DocxTemplate, RichText
import datetime
import re


class DocxT():
	""" Заполнение шаблонов """
	
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
			print(f"АОСР-{number}_{name_act[:30]}..{name_act[-17:-1]}.docx")
			doc.save(f"{path}АОСР-{number}_{name_act[:30]}..{name_act[-17:-1]}.docx")
		else:	
			print(f"АОСР-{number}_{name_act[:50]}.docx")
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

	