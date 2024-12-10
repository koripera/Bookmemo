# -*- coding: utf-8 -*-
import os
import shutil

import pickle
import glob

import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog
from tkinter import scrolledtext
import keyboard

import re
import pyautogui
import pyperclip
import sys

import unicodedata

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A5, portrait
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
from decimal import Decimal, ROUND_HALF_UP
import subprocess
import win32api
			
def main():
#ﾒｲﾝになるｳｲﾝﾄﾞｳの作成
	
	root = tk.Tk()#-----------------------------------------------------------------------
	#ﾀｲﾄﾙﾊﾞｰを非表示にする
	root.wm_overrideredirect(True)


	root.geometry("1126x1130")	

	tab_L = tk.Frame(root,height=1038, width=100)
	tab_L.place(x=0 , y=0)

	tree = ttk.Treeview(root,show=["tree"])
	tree.place(x=100  , y =0  , width=150 ,relheight=1)

	#fontsize8,height94,width64のとき、ypx=1038,xpx=388
	tbox_L = tk.Text(height=94 ,width=64 ,font=('ＭＳ ゴシック', 8))
	tbox_L.place(x=100+150 , y=0)

	tbox_R = tk.Text(height=94 ,width=64 ,font=('ＭＳ ゴシック', 8))	
	tbox_R.place(x=488+150 , y=0)

	tab_R = tk.Frame(root,height=1038, width=100)
	tab_R.place(x=876+150 , y=0)

	tbox_L.configure(bg="#eee",state='disabled')
	tbox_R.configure(bg="#eee",state='disabled')

	com = CommonData(root,tree,tbox_L,tbox_R)
	com.tree_reset()

	tree.bind("<<TreeviewSelect>>",lambda e:com.tree_choice())
	tree.bind("<ButtonRelease-3>",lambda e:com.tree_menu(e))

	#window移動操作を可能にする
	window = Windowmove(root)
	window.set(tab_L)
	window.set(tab_R)

	root.mainloop()#-----------------------------------------------------------------------




class CommonData:
	#treeのIDは自動で割り振り
	#edit_bookは treeID : (絶対path、OBJ) の辞書
	#edit_pageは treeID : OBJ           の辞書


	def __init__(self,root,tree,tbox_L,tbox_R):
		self.root = root
		self.tree = tree
		self.tbox_L = tbox_L
		self.tbox_R = tbox_R

		self.edit_book = {}
		self.edit_page = {}
		self.opened_page = None

		self.sub_win = None

		os.chdir('data')
		self.path = os.getcwd().replace("\\", "/")

	def tree_menu(self,event):
	#ﾂﾘｰ右ｸﾘｯｸ時に出現させるﾒﾆｭｰ

		#選択物の状態確認
		ID        = self.tree.focus()
		parent_ID = self.tree.parent(ID)


		menu = tk.Menu(self.root,tearoff=0)		
		menu.add_command(label="新しいブック"                  ,command=lambda: self.newbook())

	
		#ｱｲﾃﾑ選択時のみ追加する操作	
		if ID != "":
			menu.add_separator()
			menu.add_command(label="子ページの追加"            ,command=lambda: self.newpage())
			if parent_ID != "":
				
				menu.add_command(label="次ページの追加"        ,command=lambda: self.newpage_insert())

			menu.add_separator() 
			menu.add_command(label="上に"                      ,command=lambda: self.swap_up())
			menu.add_command(label="下に"                      ,command=lambda: self.swap_down())
			menu.add_separator()
			menu.add_command(label="選択中のものをリネーム"    ,command=lambda: self.rename(event))

			menu.add_separator()
			menu.add_separator()	
			menu.add_command(label="選択中のものを削除"        ,command=lambda: self.delete())
			menu.add_separator()
			menu.add_separator()

			menu.add_command(label="選択中ﾍﾟｰｼﾞの印刷(単体)"        ,command=lambda:self.pdf_print("one"))
			menu.add_command(label="選択中ﾍﾟｰｼﾞの印刷(子孫含)"      ,command=lambda:self.pdf_print())

		#いつでも可能な操作
		menu.add_separator()
		menu.add_command(label="保存"                ,command=lambda:self.save())
		menu.add_command(label="保存して終了"        ,command=lambda:self.close())
		
		menu.post(event.x_root, event.y_root)
			
	def newbook(self):
	#新しいﾌｧｲﾙの作成

		#ﾌｧｲﾙ名決定をGUIでする
		path = tk.filedialog.asksaveasfilename(
			title            = "新規作成",
			filetypes        = [("memﾌｧｲﾙ", ".mem")],
			initialdir       = self.path,
			defaultextension = "mem"
			)

		#入力があるなら実行
		if path != "":

			#表示する名前をﾌｧｲﾙ名から作成
			_ , name = os.path.split(path)
			name = name.replace(".mem","")

			#名前だけの空ﾍﾟｰｼﾞを作って保存
			txtdata=("","")
			newbook = Page(txtdata , name = name)
			with open(path, 'wb') as f : pickle.dump( newbook , f )

			#ﾂﾘｰに表示・編集中ﾃﾞｰﾀに登録
			ID = self.tree.insert("", "end", text = newbook.name)
			self.edit_book[ID] = (path,newbook)
			self.edit_page[ID] = newbook

	def newpage(self):
	#下の階層にﾍﾟｰｼﾞを増やす		

		#空ﾍﾟｰｼﾞの作成
		txtdata=("","")
		newpage = Page(txtdata)

		#選択中のものを親にﾍﾟｰｼﾞ追加
		ID = self.tree.focus()
		self.edit_page[ID].sub.append(newpage)

		#ﾂﾘｰに表示・編集中ﾃﾞｰﾀに登録
		newID = self.tree.insert(ID, "end", text = newpage.name)
		self.edit_page[newID] = newpage

	def newpage_insert(self):
	#現ﾍﾟｰｼﾞの次にﾍﾟｰｼﾞを増やす(最上位以外が対象)

		#空ﾍﾟｰｼﾞの作成
		txtdata=("","")
		newpage = Page(txtdata)

		#tree表示上のｲﾝﾃﾞｯｸｽを取る
		ID = self.tree.focus()		
		tree_index = self.tree.index(ID)

		#ﾃﾞｰﾀ内部上のｲﾝﾃﾞｯｸｽの次に挿入
		parent_ID = self.tree.parent(ID)
		list_index = self.edit_page[parent_ID].sub.index(self.edit_page[ID])
		self.edit_page[parent_ID].sub.insert(list_index+1,newpage)

		#ﾂﾘｰに表示・編集中ﾃﾞｰﾀに登録
		newID = self.tree.insert(parent_ID, tree_index+1, text = newpage.name)
		self.edit_page[newID] = newpage

	def rename(self,event):
	#ﾍﾟｰｼﾞの名前の変更

		def set_name(self,ID,name):
			self.tree.item(ID ,text = name)#tree上の表示変更
			self.edit_page[ID].name = name #ﾃﾞｰﾀの中身変更
			self.sub_win.destroy()         #ｳｲﾝﾄﾞｳの消去

		#選択を取る
		ID = self.tree.focus()

		#ｻﾌﾞｳｲﾝﾄﾞｳで新しい名前を入力、ｴﾝﾀｰで確定
		if self.sub_win == None or not self.sub_win.winfo_exists():
			self.sub_win = tk.Toplevel()
			self.sub_win.attributes("-topmost", True)

			entry = tk.Entry(self.sub_win,width=20)
			entry.insert(0, self.edit_page[ID].name )
			entry.focus_set()
			entry.pack(padx=50,pady=20)

			self.sub_win.geometry("+" + str(event.x_root) + "+" + str(event.y_root))

			entry.bind("<Return>",lambda e:set_name(self,ID,entry.get()))

	def close(self):
	#終了時の処理

		if self.opened_page != None:
			self.opened_page.L = self.tbox_L.get("1.0", 'end-1c') 
			self.opened_page.R = self.tbox_R.get("1.0", 'end-1c')

		self.file_save()
		sys.exit()

	def delete(self):
	#開いているﾍﾟｰｼﾞを消す処理

		#再保存を防ぐため、ﾍﾟｰｼﾞを無しにしておく
		self.opened_page = None

		#選択を取る
		ID = self.tree.focus()
			
		#親のID取る
		#最上位ならﾌｧｲﾙを消す
		#下位なら親のsubから消す
		parent_ID = self.tree.parent(ID)

		if parent_ID == "":
			path , _ = self.edit_book[ID]
			os.remove(path)
			#book辞書からの削除
			del self.edit_book[ID]

		else:
			self.edit_page[parent_ID].sub.remove(self.edit_page[ID])
		
		#treeからの削除
		self.tree.delete(ID)

		#page辞書からの削除
		del self.edit_page[ID]

	def tree_reset(self):
	#ﾂﾘｰの読み込みと表示

		#treeのIDとpageOBJを紐づけるのに、辞書を使って登録しておく	

		def pageload(self,page,parent=""):
			ID = self.tree.insert(parent, "end", text = page.name , open=True)

			#IDとﾍﾟｰｼﾞの紐づけ
			self.edit_page[ID] = page

			#IDとﾌﾞｯｸの紐づけ
			if parent == "":
				self.edit_book[ID] = (path,book)
			
			#子ﾍﾟｰｼﾞの読み込み
			for subpage in page.sub:
				pageload(self,subpage,parent=ID)


		#作業ﾃﾞｨﾚｸﾄﾘ内の全memﾌｧｲﾙを取得する
		file_pathlist = [path.replace("\\", "/") for path in glob.glob(self.path+"/**/*", recursive=True) if os.path.splitext(path)[1]==".mem"]

		#ﾌｧｲﾙをﾂﾘｰに追加
		for path in file_pathlist:
			with open(path, 'rb') as f : book = pickle.load(f)

			#ｻﾌﾞﾍﾟｰｼﾞまで回帰的に処理
			pageload(self,book)

	def swap_up(self):
	#ﾍﾟｰｼﾞの順序を前に

		ID = self.tree.focus()

		if ID != "":
			#ﾂﾘｰの表示
			index = self.tree.index(ID)
			self.tree.move(ID, self.tree.parent(ID) , index-1)

			#ﾃﾞｰﾀ内
			#親のID取る
			#親のsub
			parent_ID = self.tree.parent(ID)

			if parent_ID == "":
				pass
			else:
				index = self.edit_page[parent_ID].sub.index(self.edit_page[ID])
				if index != 0: 
					self.edit_page[parent_ID].sub[index-1] , self.edit_page[parent_ID].sub[index]\
					= self.edit_page[parent_ID].sub[index] , self.edit_page[parent_ID].sub[index-1]

	def swap_down(self):
	#ﾍﾟｰｼﾞの順序を後に

		ID = self.tree.focus()

		if ID != "":
			#ﾂﾘｰの表示
			index = self.tree.index(ID)
			self.tree.move(ID, self.tree.parent(ID) , index+1)

			#ﾃﾞｰﾀ内
			#親のID取る
			#親のsub
			parent_ID = self.tree.parent(ID)

			if parent_ID == "":
				pass
			else:
				index = self.edit_page[parent_ID].sub.index(self.edit_page[ID])
				max_index = len(self.edit_page[parent_ID].sub)-1
				if index != max_index: 
					self.edit_page[parent_ID].sub[index] , self.edit_page[parent_ID].sub[index+1]\
					= self.edit_page[parent_ID].sub[index+1] , self.edit_page[parent_ID].sub[index]

	def tree_choice(self):
		self.tbox_L.configure(bg="#fff",state='normal')
		self.tbox_R.configure(bg="#fff",state='normal')

		#選択前の変更を自動で更新する
		if self.opened_page != None:
			self.opened_page.L = self.tbox_L.get("1.0", 'end-1c') 
			self.opened_page.R = self.tbox_R.get("1.0", 'end-1c')

		#表示のﾘｾｯﾄ
		self.tbox_L.delete('1.0', 'end')
		self.tbox_R.delete('1.0', 'end')

		ID = self.tree.focus()
		if ID == "":
			self.tbox_L.configure(bg="#eee",state='disabled')
			self.tbox_R.configure(bg="#eee",state='disabled')
			self.opened_page = None

		else:
			nextpage = self.edit_page[ID]

			self.tbox_L.insert("1.0",nextpage.L)
			self.tbox_R.insert("1.0",nextpage.R)

			self.opened_page = self.edit_page[ID]
		
	def save(self):
		if self.opened_page != None:
			self.opened_page.L = self.tbox_L.get("1.0", 'end-1c') 
			self.opened_page.R = self.tbox_R.get("1.0", 'end-1c')

		self.file_save()

	def file_save(self):
		for path , OBJ in list(self.edit_book.values()):
			with open(path, 'wb') as f : pickle.dump(OBJ , f)

	def pdf_print(self,mode="normal"):
		#空白のﾍﾟｰｼﾞ作成
		sheet = pdfdata("book.pdf",148,210)

		#ﾃﾞｰﾀを取得
		ID = self.tree.focus()
		page = self.edit_page[ID]

		#ﾃﾞｰﾀの書き込み
		sheet.notebook(page,mode)

		#ﾌｧｲﾙとして保存
		sheet.save()

		#規定のﾌﾟﾘﾝﾀで印刷
		#win32api.ShellExecute(0,"print","book.pdf",None,".",0);

		#規定のﾌﾟﾛｸﾞﾗﾑで開く
		#subprocess.Popen(["start", output_pdf], shell=True)

def get_east_asian_width_count(text):
	count = 0
	for c in text:
		if unicodedata.east_asian_width(c) in 'FWA':
			count += 2
		else:
			count += 1
	return count

def strcut(txt,wantlong):#文字列先頭から欲しい長さを求める
	num = 0
	maxnum = len(txt)
	localtxt = txt
	txtlist = []

	while True:
		count = get_east_asian_width_count(localtxt[0:num])#文字数のカウント
		if count > wantlong:#欲しい長さを超えた時
			txtlist.append( localtxt[0:num-1] )#頭の文字を抜き取る
			localtxt = localtxt[num-1:]
			num = 0
			continue

		num = num+1

		if num > maxnum:
			txtlist.append( localtxt[:num-1] )
			break
	
	return txtlist

class pdfdata:
#pdfdata()  (str/名前 int/横ｻｲｽﾞ int/縦ｻｲｽﾞ int/ﾌｫﾝﾄｻｲｽﾞ str/ﾌｫﾝﾄ名 dic/余白)
	setup = False

	def setting(self):
		pdfmetrics.registerFont(TTFont("MSゴシック", "C:/Windows/Fonts/msgothic.ttc"))
		pdfdata.setup = True

	def __init__( self , name , width , height , fontsize=6 , font="MSゴシック",margin = {"top":5, "bottom":5, "right":5, "left":5}):

		if pdfdata.setup == False : self.setting()
			
		self.x,self.y = 0,0  #原点
		self.name = name
		self.width = width
		self.height = height
		self.fontsize = fontsize
		self.font = font
		self.margin = margin
 
		self.file_path = self.name
		self.page = canvas.Canvas(self.file_path)

		#フォントサイズ(mm)を計算
		self.fontsize_mm = float(Decimal(str( fontsize * 0.35278 )).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
		self.page.setPageSize((self.width*mm, self.height*mm))

		# フォントの設定(第1引数：フォント、第2引数：サイズ)
		self.page.setFont(font, fontsize)

		#線を細くして見やすく
		self.page.setStrokeColorRGB(0.1, 0.1, 0.1)  #どっちかあることで空白のページになる、他に記述ないとPDF開けない
		self.page.setLineWidth(0.05)                #→↑ 

	def newpage(self):
		self.page.showPage()
		self.page.setFont(self.font, self.fontsize)

	def notebook(self,page,mode):

		#レイアウト
		x_center = self.width/2
		chara_half = self.fontsize_mm/2
		x_center_hidari = x_center - chara_half
		x_center_migi = x_center + chara_half
		writearea_yoko = self.width - self.margin["left"] - self.margin["right"]
		lines = int((self.height - self.margin["top"] - self.margin["bottom"])/self.fontsize_mm)

		def getpage(page):
			data=[page,]

			#親(i)の後ろに、subのﾘｽﾄを個別挿入
			for i , page in enumerate(data):				
				data[i+1:i+1] = page.sub
			
			return data

		if mode == "normal":
			data = getpage(page)
		elif mode == "one":
			data = [page] 

		for page in data:

			#改行の調整
			dataL=[]
			dataR=[]

			for i in page.L.split("\n"):
				count = get_east_asian_width_count(i)

				if count <= 64:
					dataL.append(i)
				
				if count > 64:
					for n in strcut(i,64):
						dataL.append(n)

			for i in page.R.split("\n"):
				count = get_east_asian_width_count(i)

				if count <= 64:
					dataR.append(i)
				
				if count > 64:
					for n in strcut(i,64):
						dataR.append(n)

			#左のコラム分
			rows = 1
			for i in dataL:
				#x,yは一行目左下の座標
				x = self.margin["left"]
				y = self.height - self.margin["top"] - rows * self.fontsize_mm
				self.page.drawString(x*mm, y*mm, i)
				rows = rows + 1

			#右のコラム分
			rows = 1
			for i in dataR:
				#x,yは一行目左下の座標
				x = x_center_migi
				y = self.height - self.margin["top"] - rows * self.fontsize_mm
				self.page.drawString(x*mm, y*mm, i)
				rows = rows + 1
			
			self.newpage()

	def save(self):
		self.page.save()





class Page:
	def __init__(self,txtdata,name="notitle"):
		self.name   = name
		self.L      = txtdata[0]
		self.R      = txtdata[1]
		self.sub    = []
		
class Windowmove:#ｳｲﾝﾄﾞｳ移動用
	def __init__(self,root):
		self.root = root
		self.click = False
		self.pos       = (0,0)
		self.start_pos = (0,0)

	def set(self,item):
		item.bind("<ButtonPress-1>",  lambda event : self.on(event))
		item.bind("<Motion>",         lambda event : self.move(event))
		item.bind("<ButtonRelease-1>",lambda event : self.off(event))

	def on(self,event):
		self.click = True
		self.start_pos   = pyautogui.position()#クリック座標を覚えておく
		
	def off(self,event):
		self.click=False
		tmp = self.root.geometry().split("+")
		self.pos = (int(tmp[1]),int(tmp[2]))
		
	def move(self,event):
		if self.click==True:
			move_pos = pyautogui.position()
			x_add = move_pos[0]-self.start_pos[0]
			y_add = move_pos[1]-self.start_pos[1]

			x = x_add + self.pos[0]
			y = y_add + self.pos[1]

			self.root.geometry("1126x1038+"+str(x)+"+"+str(y))


if __name__ == "__main__":
	main() 


