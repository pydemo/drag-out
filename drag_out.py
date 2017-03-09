from __future__ import print_function
import wx
import os
import time
import win32gui
import urllib
from win32com.client import gencache
 
########################################################################
class MyListCtrl(wx.ListCtrl):
 
	#----------------------------------------------------------------------
	def __init__(self, parent, id):
		wx.ListCtrl.__init__(self, parent, id, style=wx.LC_REPORT)
 
		files = os.listdir('.')
 
		self.InsertColumn(0, 'Name')
		self.InsertColumn(1, 'Ext')
		self.InsertColumn(2, 'Size', wx.LIST_FORMAT_RIGHT)
		self.InsertColumn(3, 'Modified')
 
		self.SetColumnWidth(0, 220)
		self.SetColumnWidth(1, 70)
		self.SetColumnWidth(2, 100)
		self.SetColumnWidth(3, 420)
 
		j = 0
		for i in files:
			(name, ext) = os.path.splitext(i)
			ex = ext[1:]
			size = os.path.getsize(i)
			sec = os.path.getmtime(i)
			self.InsertStringItem(j, "%s%s" % (name, ext))
			self.SetStringItem(j, 1, ex)
			self.SetStringItem(j, 2, str(size) + ' B')
			self.SetStringItem(j, 3, time.strftime('%Y-%m-%d %H:%M', 
												   time.localtime(sec)))
 
			if os.path.isdir(i):
				self.SetItemImage(j, 1)
			elif ex == 'py':
				self.SetItemImage(j, 2)
			elif ex == 'jpg':
				self.SetItemImage(j, 3)
			elif ex == 'pdf':
				self.SetItemImage(j, 4)
			else:
				self.SetItemImage(j, 0)
 
			if (j % 2) == 0:
				self.SetItemBackgroundColour(j, '#e6f1f5')
			j = j + 1
 
########################################################################
class FileHunter(wx.Frame):
	#----------------------------------------------------------------------
	def __init__(self, parent, id, title):
		wx.Frame.__init__(self, parent, -1, title)
		panel = wx.Panel(self)
 
		p1 = MyListCtrl(panel, -1)
		p1.Bind(wx.EVT_LIST_BEGIN_DRAG, self.onDrag)
		sizer = wx.BoxSizer()
		sizer.Add(p1, 1, wx.EXPAND)
		panel.SetSizer(sizer)
 
		self.Center()
		self.Show(True)
 
	#----------------------------------------------------------------------
	def onDrag(self, event):
		data = wx.FileDataObject()
		obj = event.GetEventObject()
		dropSource = wx.DropSource(obj)

		dropSource.SetData(data)

		#next line will make the drop target window come to top, allowing us
		#to get the info we need to do the work, if it's Explorer
		result = dropSource.DoDragDrop(0)

		#get foreground window hwnd
		h = win32gui.GetForegroundWindow()

		#get explorer location
		
		s = gencache.EnsureDispatch('Shell.Application')
		#s = win32com.client.Dispatch("Shell.Application")
		loc, outdir = None, None
		for w in s.Windows():
			if int(w.Hwnd) == h:
				loc = w.LocationURL
		if loc:
			outdir = loc.split('///')[1]
			outdir = urllib.unquote(outdir)
		print (outdir)
		#got what we need, now download to outfol
		#if outdir and os.path.isdir(outdir):
		#	self.dloadItems(event, outdir)


		return
	
	def onDrag0(self, event):
		""""""
		data = wx.FileDataObject()
		obj = event.GetEventObject()
		id = event.GetIndex()
		filename = obj.GetItem(id).GetText()
		dirname = os.path.dirname(os.path.abspath(os.listdir(".")[0]))
		fullpath = str(os.path.join(dirname, filename))

		data.AddFile(fullpath)
 
		dropSource = wx.DropSource(obj)
		dropSource.SetData(data)
		#result = dropSource.DoDragDrop()
		result =  dropSource.DoDragDrop(wx.Drag_AllowMove) 
		print (fullpath)
		print(result)
 
#----------------------------------------------------------------------
app = wx.App(False)
FileHunter(None, -1, 'File Hunter')
app.MainLoop()
