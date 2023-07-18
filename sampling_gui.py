#!/usr/bin/python
# -*- encoding: utf-8 -*-
# File    :   sampling_gui.py
# Time    :   2023/07/05 00:30
# Author  :   Chang, Chun-Wei
# Email:   jimmychwchang@gmail.com
# Description : NTUWS抽樣GUI操作介面

import os
from ntuws_sampling.ws_sampling import ws_sample
os.chdir(os.path.dirname(__file__))
import tkinter as tk
from tkinter import filedialog as fdl

'''
### default files' path
## Optional
member_path = "member-29723.csv"  # 會員名單路徑
population_path = "population.xlsx"  # 母體底冊路徑
blacklist_path = "blacklist0613.csv"  # 黑名單路徑
sampled_path = "sampled.csv"  # 已抽出名單路徑 (每次抽樣完會自動更新)
output_path = "第三套樣本.xlsx"  # 樣本輸出存放路徑，需含副檔名.xlsx
# 是否使用活躍會員來抽樣，是的話請在Judgement填 Yes，否的話請在Judgement填 No
Judgement = 'Yes'
actived_path = 'member29723_active.csv'
'''

class sampling_gui():
    def __init__(self):
        self.win = tk.Tk()
        self.win.geometry("550x500")
        self.win.title("Select the Files")
        
        self.member_lb = tk.Label(text = "member file")
        self.population_lb = tk.Label(text ="population file")
        self.blacklist_lb = tk.Label(text ="blacklist file")
        self.sampled_lb = tk.Label(text ="sampled file")
        self.number_lb = tk.Label(text = "樣本套數")
        self.sample_nos_lb = tk.Label(text="抽樣數量")   

        self.set_num_en = tk.Entry(width=5)
        self.member_en = tk.Entry(width=40)
        self.population_en = tk.Entry(width=40)
        self.blacklist_en = tk.Entry(width=40)
        self.sampled_en = tk.Entry(width=40)
        self.sample_nos_en = tk.Entry(width=5)
        self.act_en = tk.Entry(width=40)
        self.mem_bt = tk.Button(
            self.win, 
            height = 1,
            text = "選擇檔案",
            command = lambda: self.select_file(filetype = "member_path")
        )
        self.popu_bt = tk.Button(
            self.win, 
            height = 1,
            text = "選擇檔案",
            command = lambda: self.select_file(filetype = "population_path")
        )
        self.bckl_bt = tk.Button(
            self.win, 
            height = 1,
            text = "選擇檔案",
            command = lambda: self.select_file(filetype = "blacklist_path")
        )
        self.sp_bt = tk.Button(
            self.win, 
            height = 1,
            text = "選擇檔案",
            command = lambda: self.select_file(filetype = "sampled_path")
        )
        self.Judgement_var = tk.IntVar(self.win)
        self.Judgement_bt = tk.Checkbutton(
            self.win,
            height = 1,
            text = "是否以活躍會員名單抽樣",
            variable = self.Judgement_var,
            onvalue = 1, offvalue = 0,
            command = lambda: self.active_member_judge()
        )
        self.act_bt = tk.Button(
            self.win,
            height = 1,
            text = "選擇檔案",
            command = lambda: self.active_file()
        )

        self.exe_btn = tk.Button(
            self.win, 
            height = 1,
            text = "執行",
            command = self.execute_sampling
        )
        
        self.member_lb.grid(row = 0, column = 0, columnspan = 5,padx=10, ipadx=5, sticky = "w")
        self.member_en.grid(row = 1, column = 0, columnspan = 4, padx=10, ipadx=5)
        self.mem_bt.grid(row = 1, column = 4, padx=10, ipadx=5)
        self.population_lb.grid(row = 2, column = 0, columnspan = 5,padx=10, ipadx=5, sticky = "w")
        self.population_en.grid(row = 3, column = 0, columnspan = 4, padx=10, ipadx=5)
        self.popu_bt.grid(row = 3, column = 4, padx=10, ipadx=5)
        self.blacklist_lb.grid(row = 4, column = 0, columnspan = 5,padx=10, ipadx=5, sticky = "w")
        self.blacklist_en.grid(row = 5, column = 0, columnspan = 4, padx=10, ipadx=5)
        self.bckl_bt.grid(row = 5, column = 4, padx=10, ipadx=5)
        self.sampled_lb.grid(row = 6, column = 0, columnspan = 5,padx=10, ipadx=5, sticky = "w")
        self.sampled_en.grid(row = 7, column = 0, columnspan = 4, padx=10, ipadx=5)
        self.sp_bt.grid(row = 7, column = 4, padx=10, ipadx=5)
        self.number_lb.grid(row = 8, column = 0,padx=10, ipadx=5, sticky = "E")
        self.set_num_en.grid(row = 8, column = 1, padx=10, ipadx=5, sticky = "W")
        self.sample_nos_lb.grid(row = 8, column = 2, padx=10, ipadx=5, sticky = "E")
        self.sample_nos_en.grid(row = 8, column = 3, columnspan = 2, padx=10, ipadx=5, sticky = "W")
        self.Judgement_bt.grid(row = 9, column = 0, padx=10, ipadx=5)
        self.act_en.grid(row = 10, column = 0, columnspan = 4,padx=10, ipadx=5)
        self.act_bt.grid(row = 10, column = 4, padx=10, ipadx=5)
        self.exe_btn.grid(row= 11, column = 2, padx=10, ipadx=5)
        self.win.mainloop()
    
    def select_file(self, filetype: str):
        if filetype == "member_path":
            self.member_path = fdl.askopenfilename()
            self.member_en.delete(0, tk.END)
            self.member_en.insert(0, self.member_path)
        elif filetype == "population_path":
            self.population_path = fdl.askopenfilename()
            self.population_en.delete(0, tk.END)
            self.population_en.insert(0, self.population_path)
        elif filetype == "blacklist_path":
            self.blacklist_path = fdl.askopenfilename()
            self.blacklist_en.delete(0, tk.END)
            self.blacklist_en.insert(0, self.blacklist_path)
        elif filetype == "sampled_path":
            self.sampled_path = fdl.askopenfilename()
            self.sampled_en.delete(0, tk.END)
            self.sampled_en.insert(0, self.sampled_path)
        elif filetype == "output_path":
            self.output_path = f"第{self.sample_setno()}套樣本.xlsx"
        else:
            print("file does not exsit")
            
    def sample_setno(self):
        smaple_set_number = self.set_num_en.get()
        return smaple_set_number
          
    def get_output_num(self):
        self.select_file("output_path")
        return self.output_path
        
    def active_file(self):        
        self.active_path = fdl.askopenfilename()
        self.act_en.delete(0, tk.END)
        self.act_en.insert(0, self.active_path)

    def active_member_judge(self):
        if self.Judgement_var.get() == 1:
            self.act_en.config(state = "normal")
            self.act_bt.config(state = "normal")
            self.Judgement = "Yes"
        else:
            self.act_en.config(state = "disabled")
            self.act_bt.config(state = "disabled")
            self.Judgement = "No"

    def sample_numbers(self):
        self.nos_samples = self.sample_nos_en.get()
        return self.nos_samples

    def execute_sampling(self):
        self.get_output_num()
        n = self.sample_numbers()
        n = int(n)
        sample_ex = ws_sample(self.member_path)
        sample_ex = sample_ex.popu_sample(
            n, self.Judgement, self.active_path, self.population_path, self.blacklist_path, self.sampled_path
        )
        sample_ex.to_excel(self.output_path, index=False, encoding="utf_8_sig")
        self.win.destroy()
    
sampling_gui()