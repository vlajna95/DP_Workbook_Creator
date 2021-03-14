import openpyexcel as ope

def create_workbook(textfile_path, workbook_path, author="", title="", subject=""):
	"""Creates and saves a new workbook."""
	with open(textfile_path, "r") as f:
		source = f.read()
		sheets = source.split("-----\n")
		sheet_number = 0
		new_workbook = ope.Workbook()
		for sheet in sheets:
			if sheet_number == 0:
				s = new_workbook[new_workbook.sheetnames[0]]
			else:
				s = new_workbook.create_sheet()
			rows = sheet.split("\n")
			s.title = rows[0]
			for r in range(1, len(rows)):
				columns = rows[r].split("\t")
				for c in range(len(columns)):
					cell = s.cell(r, c+1)
					cell.value = columns[c]
			sheet_number += 1
		new_workbook.properties.creator = author
		new_workbook.properties.title = title
		new_workbook.properties.subject = subject
		new_workbook.save(workbook_path)
		return True
