import os
import wx
from wx.adv import Sound
import workbook_sdk

def _(text):
	return text


"""
Create Excel workbooks with multiple sheets 
using just formatted text as source.
"""


class Window(wx.Frame):
	"""Creates a new Excel workbook from textual source."""
	def __init__(self, title):
		super(Window, self).__init__(None, -1, title)
		self._textfile_path = ""
		self._workbook_path = ""
		self._init_window()

	@property
	def textfile_path(self):
		"""Returns the current textfile path."""
		return self._textfile_path

	@textfile_path.setter
	def textfile_path(self, path):
		"""Sets a new path of the textfile."""
		self._textfile_path = path

	@property
	def workbook_path(self):
		"""Returns the current path for the new workbook."""
		return self._workbook_path

	@workbook_path.setter
	def workbook_path(self, path):
		"""Sets a new path for the new workbook."""
		self._workbook_path = path

	@property
	def author(self):
		"""Returns the current name of the author of the workbook."""
		return self._author

	@author.setter
	def author(self, author):
		"""Sets a new name of the author of the workbook."""
		self._author = author

	@property
	def title(self):
		"""Returns the current title of the workbook."""
		return self._title

	@title.setter
	def title(self, title):
		"""Sets a new title of the workbook."""
		self._title = title

	@property
	def subject(self):
		"""Returns the current subject of the workbook."""
		return self._subject

	@subject.setter
	def subject(self, subject):
		"""Sets a new subject of the workbook."""
		self._subject = subject

	def _init_window(self):
		"""Creates the main window and binds the needed event handlers."""
		panel = wx.Panel(self)
		panel.SetLabel(self.GetTitle())
		textfile_path_label = wx.StaticText(panel, label=_("&Tekstualni fajl *"))
		self._textfile_path_field = wx.TextCtrl(panel)
		self._textfile_path_field.Bind(wx.EVT_TEXT, self._on_textfile_path_field_text_change)
		self._textfile_path_button = wx.Button(panel, label=_("Izaberi &izvor tabele..."))
		self._textfile_path_button.Bind(wx.EVT_BUTTON, self._on_textfile_path_button)
		workbook_path_label = wx.StaticText(panel, label=_("&Excel fajl *"))
		self._workbook_path_field = wx.TextCtrl(panel)
		self._workbook_path_field.Bind(wx.EVT_TEXT, self._on_workbook_path_field_text_change)
		self._workbook_path_button = wx.Button(panel, label=_("Izaberi &lokaciju za čuvanje..."))
		self._workbook_path_button.Bind(wx.EVT_BUTTON, self._on_workbook_path_button)
		author_label = wx.StaticText(panel, label=_("&Autor"))
		self._author_field = wx.TextCtrl(panel)
		self._author_field.Bind(wx.EVT_TEXT, self._on_author_field_text_change)
		title_label = wx.StaticText(panel, label=_("&Naslov"))
		self._title_field = wx.TextCtrl(panel)
		self._title_field.Bind(wx.EVT_TEXT, self._on_title_field_text_change)
		subject_label = wx.StaticText(panel, label=_("&Kategorija"))
		self._subject_field = wx.TextCtrl(panel)
		self._subject_field.Bind(wx.EVT_TEXT, self._on_subject_field_text_change)
		notes_label = wx.StaticText(panel, label=_("Polja označena zvezdicom (*) su obavezna."))
		self._save_button = wx.Button(panel, label=_("&Sačuvaj"))
		self._save_button.Bind(wx.EVT_BUTTON, self._on_save_button)
		self._save_button.SetDefault()
		sizer = wx.GridBagSizer(5, 5)
		sizer.Add(textfile_path_label, (0, 0), (1, 1))
		sizer.Add(self._textfile_path_field, (0, 1), (1, 1), flag=wx.EXPAND|wx.ALL)
		sizer.Add(self._textfile_path_button, (0, 2), (1, 1), flag=wx.ALL)
		sizer.Add(workbook_path_label, (1, 0), (1, 1))
		sizer.Add(self._workbook_path_field, (1, 1), (1, 1), flag=wx.EXPAND|wx.ALL)
		sizer.Add(self._workbook_path_button, (1, 2), (1, 1), flag=wx.ALL)
		sizer.Add(author_label, (2, 0), (1, 1))
		sizer.Add(self._author_field, (2, 1), (1, 2), flag=wx.ALL)
		sizer.Add(title_label, (3, 0), (1, 1))
		sizer.Add(self._title_field, (3, 1), (1, 2), flag=wx.ALL)
		sizer.Add(subject_label, (4, 0), (1, 1))
		sizer.Add(self._subject_field, (4, 1), (1, 2), flag=wx.ALL)
		sizer.Add(notes_label, (5, 0), (1, 3), flag=wx.EXPAND)
		sizer.Add(self._save_button, (6, 0), (1, 3), flag=wx.EXPAND|wx.ALIGN_CENTER_HORIZONTAL|wx.ALL)
		panel.SetSizerAndFit(sizer)
		self.Fit()
		self.Center()
		self.Show(True)

	def _on_textfile_path_field_text_change(self, event):
		"""Fires when the text in the textfile path field is changed."""
		self.textfile_path = self._textfile_path_field.GetValue()

	def _on_workbook_path_field_text_change(self, event):
		"""Fires when the text in the workbook path field is changed."""
		self.workbook_path = self._workbook_path_field.GetValue()

	def _on_author_field_text_change(self, event):
		"""Fires when the text in the author field is changed."""
		self.author = self._author_field.GetValue()

	def _on_title_field_text_change(self, event):
		"""Fires when the text in the title field is changed."""
		self.title = self._title_field.GetValue()

	def _on_subject_field_text_change(self, event):
		"""Fires when the text in the subject field is changed."""
		self.subject = self._subject_field.GetValue()

	def _on_textfile_path_button(self, event):
		"""Fires when the textfile path button is activated."""
		with wx.FileDialog(self, _("Izaberi fajl koji sadrži izvor tabele"), wildcard=f"{_('Tekstualni fajlovi')} (*.txt)|*.txt|{_('Sve vrste fajlova')} (*.*)|*.*", style=wx.FD_OPEN|wx.FD_FILE_MUST_EXIST) as textfile_dialog:
			if textfile_dialog.ShowModal() == wx.ID_OK:
				self.textfile_path = textfile_dialog.GetPath()
				self._textfile_path_field.SetValue(self.textfile_path)
			else:
				pass

	def _on_workbook_path_button(self, event):
		"""Fires when the workbook path button is activated."""
		with wx.FileDialog(self, _("Izaberi gde će Excel dokument biti sačuvan"), wildcard=f"{_('Excel dokumenti')} (*.xlsx;*.xls)|*.xlsx;*.xls", style=wx.FD_SAVE|wx.FD_OVERWRITE_PROMPT) as workbook_dialog:
			if workbook_dialog.ShowModal() == wx.ID_OK:
				self.workbook_path = workbook_dialog.GetPath()
				self._workbook_path_field.SetValue(self.workbook_path)
			else:
				pass

	def _on_save_button(self, event):
		"""Fires when the Save button is activated."""
		if self.textfile_path == "":
			wx.MessageDialog(self, _("Mora se odabrati ili upisati putanja do tekstualnog izvora tabele."), _("Greška!"), style=wx.ICON_ERROR|wx.OK).ShowModal()
			self._textfile_path_field.SetFocus()
		elif self.workbook_path == "":
			wx.MessageDialog(self, _("Mora se odabrati ili upisati putanja na kojoj će Excel dokument biti kreiran."), _("Greška!"), style=wx.ICON_ERROR|wx.OK).ShowModal()
			self._workbook_path_field.SetFocus()
		elif os.path.splitext(self.textfile_path)[1] in [".docx", ".doc"]:
			wx.MessageDialog(self, _("Ne može se koristiti Word dokument kao izvor. Odaberite tekstualni dokument u drugom formatu."), _("Greška!"), style=wx.ICON_ERROR|wx.OK).ShowModal()
			self._textfile_path_field.SetFocus()
		elif os.path.splitext(self.workbook_path)[1] not in [".xlsx", ".xls"]:
			wx.MessageDialog(self, _("Excel dokument mora imati ekstenziju .xlsx ili .xls i potrebno je da ponovo odaberete putanju ili ispravite upisano."), _("Greška!"), style=wx.ICON_ERROR|wx.OK).ShowModal()
			self._workbook_path_field.SetFocus()
		else:
			if os.path.exists(self.workbook_path):
				with wx.MessageDialog(self, _("Excel dokument sa odabranim nazivom već postoji. \n{} \nDa li sigurno želiš da ga ukloniš i sačuvaš novi pod tim nazivom?").format(self.workbook_path), _("Neophodna je potvrda"), style=wx.ICON_QUESTION|wx.ICON_WARNING|wx.YES_NO|wx.NO_DEFAULT) as dlg:
					if dlg.ShowModal() == wx.ID_YES:
						self.save_workbook(self.textfile_path, self.workbook_path)
			else:
				self.save_workbook(self.textfile_path, self.workbook_path, self.author, self.title, self.subject)

	def save_workbook(self, t_path, w_path, author="", title="", subject=""):
		"""Creates and saves the new workbook.
		
		:param textfile: The full path to a text document containing the source for the workbook. (required)
		:param workbook_path: Full path to a location for the new workbook file. (required)
		:param author: The name of the author of the new workbook. (optional)
		:param title: The title of the new workbook. (optional)
		:param subject: The subject (category) of the new workbook. (optional)
		"""
		if workbook_sdk.create_workbook(self.textfile_path, self.workbook_path, self.author, self.title, self.subject):
			Sound("c:/windows/media/tada.wav").Play(wx.adv.SOUND_ASYNC)
			wx.MessageDialog(self, _("Excel dokument je sačuvan na izabranoj putanji: \n{}".format(self.workbook_path)), _("Uspeh!"), style=wx.ICON_INFORMATION|wx.OK).ShowModal()


if __name__ == "__main__":
	app = wx.App()
	w = Window(_("DP Workbook kreator"))
	app.MainLoop()
