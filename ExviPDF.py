import tkinter as tk
from tkinter import ttk
import tkinter.filedialog as fd
from tkinter.messagebox import showinfo
from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
from PIL import ImageTk, Image 
import os
import sys


class ExviPDF:

	files_len = 0
	BUTTON_PADY = 1
	FRAME_PADX = 2
	FRAME_PAD_IN = 5
	BUTTON1_X = 20
	BUTTON2_X = 12
	ENTRY_X = 50

	def merge_addf(self):
		fip = fd.askopenfilenames(title = 'Choose files to merge', filetypes =[('PDF Files', '*.pdf')])
		if(len(fip) == 0):
			return
		self.files_len += len(fip)
		for file in fip:
			self.listbox_files.insert(self.files_len + 1, file)


	def merge_removef(self):
		selection = self.listbox_files.curselection()
		try:
			self.listbox_files.delete(selection[0])
			self.files_len -= 1
		except:
			pass


	def merge_moveup(self):
		if(self.files_len < 1):
			return
		selection = self.listbox_files.curselection()
		if(selection[0] == 0):
			return
		try:
			file = self.listbox_files.get(selection[0])
			self.listbox_files.delete(selection[0])
			self.listbox_files.insert(selection[0] - 1, file)
			self.listbox_files.select_set(selection[0] - 1)
		except:
			pass
		

	def merge_movedn(self):
		if(self.files_len < 1):
			return
		selection = self.listbox_files.curselection()
		if(selection[0] == self.files_len - 1):
			return
		try:
			file = self.listbox_files.get(selection[0])
			self.listbox_files.delete(selection[0])
			self.listbox_files.insert(selection[0] + 1, file)
			self.listbox_files.select_set(selection[0] + 1)
		except:
			pass


	def merge_savef(self):
		if(self.files_len < 2):
			showinfo('', 'Not Enough Files Chosen!')
			return
		fop = fd.asksaveasfilename(defaultextension = '.pdf', filetypes = [('PDF Files', '*.pdf')])
		if(fop == ''):
			return
		mergedObject = PdfFileMerger()
		for file in self.listbox_files.get(0, tk.END):
			mergedObject.append(PdfFileReader(file, 'rb'))
		save = open(fop, 'w')
		mergedObject.write(save.name)
		showinfo('', 'File Saved!')


	def merge_clear(self):
		self.listbox_files.delete(0, tk.END)
		self.files_len = 0


	def split_selectf(self):
		fip = fd.askopenfile(mode = 'r', title = 'Choose file to split', filetypes =[('PDF Files', '*.pdf')])
		if(fip == None):
			return
		self.entry_file.config(state = tk.NORMAL)
		self.entry_file.delete(0, tk.END)
		self.entry_file.insert(0, fip.name)
		self.entry_file.config(state = tk.DISABLED)

	
	def split_savef(self):
		total_pages = 0
		fin_pages = []
		try:
			pdfreader = PdfFileReader(open(self.entry_file.get(), 'rb'))
			total_pages = pdfreader.numPages
		except:
			showinfo('', 'Error Reading File!')
			return
		try:
			page_ranges = [r for r in self.entry_split.get().replace(' ', '').split(',')]
			for r in page_ranges:
				nums = [int(x) for x in r.split('-')]
				if(nums[0] <= nums[-1]):
					fin_pages = fin_pages + list(range(nums[0], nums[-1] + 1))
				else:
					C = nums[0]
					while(C >= nums[-1]):
						fin_pages.append(C)
						C -= 1
		except:
			showinfo('', 'Recheck Numbering!')
			return
		pdfwriter = PdfFileWriter()
		for p in fin_pages:
			if(p > total_pages):
				showinfo('', 'Invalid Page Number(s)!')
				return
			pdfwriter.addPage(pdfreader.getPage(p - 1))
		try:
			fop = fd.asksaveasfilename(defaultextension = '.pdf', filetypes = [('PDF Files', '*.pdf')])
			with open(fop, 'wb') as output_file_open:
				pdfwriter.write(output_file_open)
			showinfo('', 'File Saved!')
		except:
			return


	def merge_tab_frame(self):
		frame_root = tk.LabelFrame(self.tab_merge, padx = self.FRAME_PAD_IN, pady = self.FRAME_PAD_IN)
		frame_files = tk.LabelFrame(frame_root, text = 'Files:', padx = self.FRAME_PAD_IN, pady = self.FRAME_PAD_IN)
		frame_buttons = tk.LabelFrame(frame_root, padx = self.FRAME_PAD_IN, pady = self.FRAME_PAD_IN)

		frame_root.pack()
		frame_files.grid(row = 0, column = 0, padx = self.FRAME_PADX)
		frame_buttons.grid(row = 0, column = 1, padx = self.FRAME_PADX)
		
		scrollx_files = tk.Scrollbar(frame_files, orient = 'horizontal')
		scrolly_files = tk.Scrollbar(frame_files, orient = 'vertical')
		self.listbox_files = tk.Listbox(frame_files, width = 60, height = 15, xscrollcommand = scrollx_files.set, yscrollcommand = scrolly_files.set)
		scrollx_files.config(command = self.listbox_files.xview)
		scrolly_files.config(command = self.listbox_files.yview)
		
		button_add = tk.Button(frame_buttons, text = 'Add Files..', command = self.merge_addf, width = self.BUTTON1_X)
		button_remove = tk.Button(frame_buttons, text = 'Remove File', command = self.merge_removef, width = self.BUTTON1_X)
		button_clear = tk.Button(frame_buttons, text = 'Remove All Files', command = self.merge_clear, width = self.BUTTON1_X)
		button_moveup = tk.Button(frame_buttons, text = 'Move Up', command = self.merge_moveup, width = self.BUTTON1_X)
		button_movedn = tk.Button(frame_buttons, text = 'Move Down', command = self.merge_movedn, width = self.BUTTON1_X)
		button_save = tk.Button(frame_buttons, text = 'Save As..', command = self.merge_savef, width = self.BUTTON1_X)
		
		scrollx_files.pack(side = tk.BOTTOM, fill = tk.X)
		scrolly_files.pack(side = tk.RIGHT, fill = tk.Y)
		self.listbox_files.pack()
		button_add.grid(row = 0, column = 0, pady = self.BUTTON_PADY)
		button_remove.grid(row = 1, column = 0, pady = self.BUTTON_PADY)
		button_clear.grid(row = 2, column = 0, pady = self.BUTTON_PADY)
		button_moveup.grid(row = 3, column = 0, pady = self.BUTTON_PADY)
		button_movedn.grid(row = 4, column = 0, pady = self.BUTTON_PADY)
		button_save.grid(row = 5, column = 0, pady = self.BUTTON_PADY)


	def split_tab_frame(self):
		frame_root = tk.LabelFrame(self.tab_split, padx = self.FRAME_PAD_IN, pady = self.FRAME_PAD_IN, width = 1000, height = 1000)
		frame_file = tk.LabelFrame(frame_root, text = 'File:', padx = self.FRAME_PAD_IN, pady = self.FRAME_PAD_IN)
		frame_pages = tk.LabelFrame(frame_root, text = 'Pages:', padx = self.FRAME_PAD_IN, pady = self.FRAME_PAD_IN)

		frame_root.pack()
		frame_file.grid(row = 0, column = 0, padx = self.FRAME_PADX)
		frame_pages.grid(row = 1, column = 0, padx = self.FRAME_PADX)

		self.entry_file = tk.Entry(frame_file, state = tk.DISABLED, width = self.ENTRY_X)
		button_select = tk.Button(frame_file, text = 'Select File..', command = self.split_selectf, width = self.BUTTON2_X)
		self.entry_split = tk.Entry(frame_pages, width = self.ENTRY_X)
		button_split = tk.Button(frame_pages, text = 'Split..', command = self.split_savef, width = self.BUTTON2_X)

		self.entry_file.grid(row = 0, column = 1)
		button_select.grid(row = 0, column = 2, pady = self.BUTTON_PADY)		
		self.entry_split.grid(row = 0, column = 1)
		button_split.grid(row = 0, column = 2, pady = self.BUTTON_PADY)




	def frames_in_root(self):
		self.tabControl = ttk.Notebook(self.root)
		self.tab_merge = ttk.Frame(self.tabControl)
		self.tab_split = ttk.Frame(self.tabControl)
		self.tabControl.add(self.tab_merge, text = f'{"Merge": ^30s}')
		self.tabControl.add(self.tab_split, text = f'{"Split": ^30s}')
		self.tabControl.pack()

		txt_infodev = 'Kaushik Choudhury. March 21, 2021. github.com/kaushik-cy'
		self.frame_infodev = tk.LabelFrame(self.root)
		self.label_infodev = tk.Label(self.frame_infodev, text = txt_infodev, font = ('Helvetica', 8))
		self.frame_infodev.pack(fill = tk.X)
		self.label_infodev.pack()

		self.merge_tab_frame()
		self.split_tab_frame()


	def application_gui(self):
		self.root = tk.Tk()
		self.root.resizable(0, 0)
		self.frames_in_root()
		
		self.root.mainloop()


	def __init__(self):
		self.application_gui()


ExviPDF()
