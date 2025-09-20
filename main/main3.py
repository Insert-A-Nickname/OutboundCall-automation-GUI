import time
import os
import tkinter as tk
from tkinter import ttk, filedialog
import threading
import pyautogui  # For controlling mouse and keyboard
from pywinauto.application import Application  # For finding and focusing windows
from pywinauto.findwindows import ElementNotFoundError
from pywinauto import Desktop
import math, sys
import pandas as pd
import pyexcel as pe
import sounddevice as sd
import numpy as np
import wave
import phonenumbers
from phonenumbers import carrier
from phonenumbers.phonenumberutil import number_type, NumberParseException
import pyttsx3
from pypdf import PdfReader
# --- Version 1.5 ---

# --- Configuration ---
LOOP_DELAY_SECONDS = 20  # Time to wait at the end of the loop
POST_CALL_INITIATION_DELAY = 10  # Time to wait after pressing Enter
# --- IMPORTANT: USER ACTION REQUIRED ---
# The script will look for a window title that STARTS WITH this text.
# The audio file to play after the call connects
global file_path, config_path, AppName
# Coordinates for the number input field.
config_path='configuration\config.json'
data=pd.read_json(config_path,orient='index')
global X0,Y0, X1,Y1,X2,Y2,X3,Y3
X0,Y0, X1,Y1=data.loc[0,'click1'],data.loc[1,'click1'],data.loc[0,'click2'],data.loc[1,'click2']
X2,Y2,X3,Y3=data.loc[0,'click3'],data.loc[1,'click3'],data.loc[0,'click4'],data.loc[1,'click4']
# --- IMPORTANT: Please update these coordinates ---
# Coordinates for the RED HANG UP button.
# The values below are set to the same as the input field as requested,
# but you should update them to the actual coordinates of the hang-up button.
global c1_on, c2_on, c3_on,c4_on, conframe_status, cc, n
cc=0
c1_on, c2_on,c3_on,c4_on,conframe_status=False,False,False, False,False
AppName =data["appname"][0]
file_path=data["filepath"][0]
folder_path=os.getcwd()
#audiorecording
global sample_rate,channels,record_name,dtype, recorded_frames, audio_data,read_txt,txt_path, currentrecord, reading
reading=False
read_txt=True
sample_rate = 48000  # 48 kHz my computer uses this setting for input, check if yours uses a different frequency
channels = 2  # 1 = Mono, 2 = Stereo
record_name= "0.wav"
dtype = 'int16'
audio_data=[]
crecordlist=[]
currentrecord=False
# Buffer to hold recorded data
recorded_frames = []
recording = False
class AutomationApp(tk.Tk):
    def __init__(self):
        super().__init__()
        # --- Window Setup ---
        self.title("Desktop Automation UI v1.5")
        self.geometry("500x600")
        self.configure(bg="#2E2E2E")
        self.attributes('-topmost', True)
        # --- Style Configuration ---
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#2E2E2E")
        style.configure("TLabel", background="#2E2E2E", foreground="#FFFFFF", font=("Segoe UI", 10))
        #style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Data.TLabel", font=("Segoe UI", 12, "italic"), foreground="#A0E9FF")
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=5)
        style.map("TButton", background=[("active", "#4A4A4A")], foreground=[("active", "white")])
        # --- UI Variables ---
        self.timer_text = tk.StringVar(value=f"{LOOP_DELAY_SECONDS}")
        self.typing_text = tk.StringVar(value="--")
        self.cursor_pos_text = tk.StringVar(value="X: ---, Y: ---")
        self.is_running = False
        self.automation_thread = None
        self.countdown_end_time = None
        # --- UI Layout ---
        main_frame = ttk.Frame(self, padding="15 15 15 15")
        main_frame.pack(expand=True, fill=tk.BOTH)
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(20, 0))
        self.config_frame = ttk.Frame(main_frame)
        con_frame_l= ttk.Frame(self.config_frame)
        con_frame_l.pack(side=tk.LEFT,fill=tk.X)
        con_frame_r= ttk.Frame(self.config_frame)
        con_frame_r.pack(side=tk.LEFT,fill=tk.X)
        # UI Elements
        ttk.Label(main_frame, text="Time Until Next Loop:", style="Header.TLabel").pack(pady=(0, 5))
        ttk.Label(main_frame, textvariable=self.timer_text, style="Data.TLabel").pack()
        ttk.Label(main_frame, text="Current Action:", style="Header.TLabel").pack(pady=(20, 5))
        ttk.Label(main_frame, textvariable=self.typing_text, style="Data.TLabel").pack()
        ttk.Label(main_frame, text="Live Cursor Position:", style="Header.TLabel").pack(pady=(20, 5))
        ttk.Label(main_frame, textvariable=self.cursor_pos_text, style="Data.TLabel").pack()
        # Control Buttons
        self.start_button = ttk.Button(control_frame, text="Start", command=self.start_automation)
        self.start_button.pack(side=tk.LEFT, expand=True, padx=2)
        self.stop_button = ttk.Button(control_frame, text="Stop", command=self.stop_automation, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, expand=True, padx=2)
        self.reset_button = ttk.Button(control_frame, text="Reset", command=self.reset_automation, state=tk.DISABLED)
        self.reset_button.pack(side=tk.LEFT, expand=True, padx=2)
        #config frame
        self.select_app=ttk.Combobox(con_frame_l,height=1)
        self.select_app.pack(side=tk.BOTTOM,expand=True,pady=(2,0))
        self.browse_sheet=tk.Button(con_frame_r, text="Browse Sheet",command=self.selectsheet, state=tk.NORMAL)
        windows = Desktop(backend="uia").windows()
        list=[w.window_text() for w in windows]#get open tabs to configure
        self.select_app['values']=list
        self.select_app.bind('<ButtonRelease-1>', self.selectapp)
        self.browse_sheet.pack(side=tk.TOP,expand=False,padx=(4,0))
        self.select_click1=tk.Text(con_frame_l,height=1,width=40)
        self.select_click1.pack(side=tk.TOP,expand=True,pady=(2,0))
        self.select_click2=tk.Text(con_frame_l,height=1,width=40)
        self.select_click2.pack(side=tk.TOP,expand=False,pady=(2,0))
        self.select_click3=tk.Text(con_frame_l,height=1,width=40)
        self.select_click3.pack(side=tk.TOP,expand=False,pady=(2,0))
        self.select_click4=tk.Text(con_frame_l,height=1,width=40)
        self.select_click4.pack(side=tk.TOP,expand=False,pady=(2,0))
        self.click1=tk.Button(con_frame_r, text="click1",command=self.setc1, state=tk.NORMAL)
        self.click1.pack(side=tk.TOP,expand=False,padx=(4,0),pady=(2,0))
        self.click2=tk.Button(con_frame_r, text="click2",command=self.setc2, state=tk.NORMAL)
        self.click2.pack(side=tk.TOP,expand=False,padx=(4,0),pady=(2,0))
        self.click3=tk.Button(con_frame_r, text="click3",command=self.setc3, state=tk.NORMAL)
        self.click3.pack(side=tk.TOP,expand=False,padx=(4,0),pady=(2,0))
        self.click4=tk.Button(con_frame_r, text="click3",command=self.setc4, state=tk.NORMAL)
        self.click4.pack(side=tk.TOP,expand=False,padx=(4,0),pady=(2,0))
        self.confbutton=ttk.Button(control_frame,text='Configuration',command=self.configstate)
        self.confbutton.pack(side=tk.TOP)
        self.startingclicks()
        #read
        self.read=tk.Button(con_frame_l, text="pdf_status",command=self.readstatus, state=tk.NORMAL)
        self.read.pack(side=tk.LEFT,expand=False,padx=(4,0),pady=(2,0))
        self.testreader=tk.Text(con_frame_l,height=1,width=40)
        self.testreader.pack(side=tk.TOP,expand=False,padx=(4,0),pady=(2,0))
        self.startreading=tk.Button(control_frame, text="Read",command=self.readtext, state=tk.NORMAL)
        self.startreading.pack(side=tk.LEFT,expand=False,padx=(4,0),pady=(2,0))
        #table frame
        table_frame=ttk.Frame(main_frame)
        self.table=ttk.Treeview(table_frame)
        self.table['columns'] = ('idx', 'Acc-name', 'Phone-Number','Status','Call-File')
        self.table.column('#0', width=0,anchor=tk.W)
        self.table.column('idx', anchor=tk.W, width=20)
        self.table.column('Acc-name', anchor=tk.W, width=100)
        self.table.column('Phone-Number', anchor=tk.W, width=100)
        self.table.column('Status', anchor=tk.W, width=100)
        self.table.column('Call-File', anchor=tk.W, width=100)
        table_frame.pack()
        # Create the headings
        self.table.heading('#0', text='', anchor=tk.W)
        self.table.heading('idx', text='idx', anchor=tk.W)
        self.table.heading('Acc-name', text='Acc-name', anchor=tk.W)
        self.table.heading('Phone-Number', text='Phone-Number', anchor=tk.W)
        self.table.heading('Status', text='Status', anchor=tk.W)
        self.table.heading('Call-File', text='Call-File', anchor=tk.W)
        self.table.bind('<ButtonRelease-1>', self.selectItem)
        self.table.pack(expand=True, fill=tk.BOTH)
        try:
            self.updatetable()
        except FileNotFoundError:
            global file_path
            file_path=''
        self.periodic_ui_update()
    #Record audio
        self.stream=sd.InputStream(samplerate=sample_rate, 
                                channels=channels, 
                                dtype=dtype, 
                                callback=self.callback)
    def callback(self,indata, frames, time, status):
        global recorded_frames,crecordlist,currentrecord
        if status:
            print(status, file=sys.stderr)
        recorded_frames.append(indata.copy())
    def doaudio(self,currentcall):
        # Buffer to hold recorded data
        global recording, recorded_frames,audio_data,currentrecord
        recording = True
        self.stream.start()
        sd.sleep(100)
        audio_data = np.concatenate(recorded_frames)
        recordlist=audio_data[-1000000:]
        df=pd.DataFrame(recordlist)
        currentrecord=True
        for i in range(len(df)):
            if (int(df[0][i])>200):
                currentrecord=False
        print(currentrecord)
        df.to_csv('audio\startingsequence.csv')
        print(f'{len(audio_data)}, {len(recordlist)}')
    def stopaudio(self,currentcall):
        global sample_rate,channels,record_name,dtype,audio_data, recording,recorded_frames
        record_name=currentcall
        recording=False
        with wave.open(record_name, 'wb') as wf:
            wf.setnchannels(channels)
            wf.setsampwidth(np.dtype(dtype).itemsize)
            wf.setframerate(sample_rate)
            wf.writeframes(audio_data.tobytes())
        print(f"Saved recording as {record_name}")
        df=pd.DataFrame(audio_data)
        a=[]
        audio_data=np.array(a)
        recorded_frames=[]
        self.stream.stop()
    def selectapp(self,event):
        global AppName
        b=self.select_app.get()
        df=pd.read_json(config_path,orient="index")
        df.loc[0,"appname"]=b
        AppName=df["appname"]
        df.to_json(config_path,orient="index")
        print(b)
    def configstate(self):
        global conframe_status
        if conframe_status:
            conframe_status=False
            self.config_frame.pack_forget()
        else:
            conframe_status=True
            self.config_frame.pack(side=tk.TOP,expand=True, pady=(2, 0))
    def startingclicks(self):
        global config_path
        data=pd.read_json(config_path,orient='index')
        x1,y1=data['click1'][0],data['click1'][1]
        x2,y2=data['click2'][0],data['click2'][1]
        x3,y3=data['click3'][0],data['click3'][1]
        x4,y4=data['click4'][0],data['click4'][1]
        self.select_click1.insert(index=1.0,chars=f'{x1},{y1}')
        self.select_click2.insert(index=1.0,chars=f'{x2},{y2}')
        self.select_click3.insert(index=1.0,chars=f'{x3},{y3}')
        self.select_click4.insert(index=1.0,chars=f'{x4},{y4}')
        self.select_click1['state'],self.select_click2['state']='disabled','disabled'
        self.select_click3['state'],self.select_click4['state']='disabled','disabled'
    def clickhandler(self,currentclick,clicked,display,cc,x,y):
        data=pd.read_json(config_path,orient='index')
        if currentclick:
            x=data[f'click{cc}'][0]
            y=data[f'click{cc}'][1]
            clicked.configure(background="green")
            display.delete(index1=1.0,index2=3.0)
            display.insert(index=1.0,chars=f'{x},{y}')
            display.delete(index1=2.0,index2=3.0)
            display['state']='normal'
        else:
            clicked.configure(background="SystemButtonFace")
            txt=display.get(index1=1.0,index2=2.0).split('\n')[0].split(',')
            x, y=int(txt[0]),int(txt[1])
            data.loc[0,f'click{cc}']=x
            data.loc[1,f'click{cc}']=y
            data.to_json(config_path,orient='index')
            display['state']='disabled'
            print(x,y)
    def setc1(self):
        global c1_on, cc, X0, Y0
        if c1_on:
            c1_on=False
        else:
            c1_on=True
            cc=1
        self.clickhandler(c1_on,self.click1,self.select_click1,cc,X0,Y0)
    def setc2(self):
        global c2_on, cc, X1, Y1
        if c2_on:
            c2_on=False
        else:
            c2_on=True
            cc=2
        self.clickhandler(c2_on,self.click2,self.select_click2,cc,X1,Y1)
    def setc3(self):
        global c3_on, cc, X2, Y2
        if c3_on:
            c3_on=False
        else:
            c3_on=True
            cc=3
        self.clickhandler(c3_on,self.click3,self.select_click3,cc,X2,Y2)
    def setc4(self):
        global c4_on, cc, X3, Y3
        if c4_on:
            c4_on=False
        else:
            c4_on=True
            cc=4
        self.clickhandler(c4_on,self.click4,self.select_click4,cc,X3,Y3)
    def selectsheet(self):
        global file_path, config_path
        root = tk.Tk()
        root.withdraw()
        config=pd.read_json(config_path,orient='index')
        print(config["filepath"][0])
        file_path = filedialog.askopenfilename()
        #file_path = filedialog.askdirectory()
        config.loc[0,"filepath"]=file_path
        config.to_json(config_path, orient="index")
        self.select_app.set(file_path)
        print(config["filepath"][0])
        self.updatetable()
        child_id=self.table.get_children()[0]
        self.table.focus(child_id)
        self.table.selection_set(child_id)
    def updatetable(self):
        global file_path
        df=pd.read_excel(file_path)
        a=df[df.columns[2]]
        self.table.delete(*self.table.get_children())
        for i in range(len(df)):
            data=str(df[df.columns[0]][i]),str(df[df.columns[1]][i]),str(a[i]),str(df[df.columns[3]][i]),str(df[df.columns[4]][i])
            self.table.insert(parent='',index=i,values=data)
        self.table.pack(expand=False, fill=tk.BOTH,)
    def selectItem(self,a):
        curItem = self.table.focus()
        print (self.table.item(curItem)["values"])
    ######################
    #function to do reading#
    ######################
    def readstatus(self):
        global file_path,read_txt,txt_path
        txt_path=filedialog.askopenfilename()
        config=pd.read_json(config_path,orient='index')
        config.loc[0,"pdftoread"]=txt_path
    def doread(self):
        global config_path,reading
        df=pd.read_json(config_path,orient="index")
        bookname=df["pdftoread"][0]
        self.startreading.configure(background='green')
        reader = PdfReader(bookname)
        acc=pd.read_csv('configuration\\acc-log.csv',index_col=0)
        print(acc["Reading"])
        acc["Account-Name"]
        print(acc["Current-Page"])
        page = reader.pages[10]
        engine = pyttsx3.init()
        text = page.extract_text()
        engine.setProperty("rate", 200)
        engine.say(text)
        engine.runAndWait()
    def readtext(self):
        global config_path,reading
        if reading:
            df=pd.read_json(config_path,orient="index")
            bookname=df["pdftoread"][0]
            self.startreading.configure(background='green')
            reader = PdfReader(bookname)
            acc=pd.read_csv('configuration\\acc-log.csv',index_col=0)
            print(acc["Reading"])
            acc["Account-Name"]
            print(acc["Current-Page"])
            page = reader.pages[10]
            engine = pyttsx3.init()
            text = page.extract_text()
            engine.setProperty("rate", 200)
            engine.say(text)
            engine.runAndWait()
            reading=False
            try:
                engine.runAndWait()
            except KeyboardInterrupt:
                engine.stop
            reading=False
        else:
            self.startreading.configure(background="red")
            reading =True
    ######################
    # end of function to do reading#
    ######################
    def updatespreadsheet(self):
        global file_path
        records = pe.get_records(file_name=f'{file_path}.csv')
        # Save records to an ODS file
        pe.save_as(records=records, dest_file_name=file_path)
        os.remove(f'{file_path}.csv')
        self.updatetable()
    def start_automation(self):
        if self.is_running: return
        self.is_running = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.reset_button.config(state=tk.NORMAL)
        self.typing_text.set("Starting...")
        self.automation_thread = threading.Thread(target=self.run_automation_loop, daemon=True)
        self.automation_thread.start()
    def stop_automation(self):
        if not self.is_running: return
        self.is_running = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.reset_button.config(state=tk.DISABLED)
        self.typing_text.set("Stopped by user.")
        self.countdown_end_time = None
        self.timer_text.set(f"{LOOP_DELAY_SECONDS}")
    def reset_automation(self):
        if self.is_running:
            self.stop_automation()
            self.after(500, self.start_automation)
    def periodic_ui_update(self):
        try:
            x, y = pyautogui.position()
            self.cursor_pos_text.set(f"X: {x}, Y: {y}")
        except Exception:
            self.cursor_pos_text.set("Could not get position")
        if self.countdown_end_time and self.is_running:
            remaining_time = self.countdown_end_time - time.time()
            self.timer_text.set(f"{math.ceil(remaining_time) if remaining_time > 0 else 0}")
        self.after(100, self.periodic_ui_update)
    def run_automation_loop(self):
        """Main automation logic using desktop control, not Selenium."""
        while self.is_running:
            global AppName, X0, X1,X2,Y0,Y1,Y2,X3,Y3,config_path,currentrecord, n, txt_path
            data=pd.read_json(config_path,orient='index')
            X0,Y0=data['click1'][0],data['click1'][1]
            X1,Y1=data['click2'][0],data['click2'][1]
            X2,Y2=data['click3'][0],data['click3'][1]
            X3,Y3=data['click4'][0],data['click4'][1]
            try:
                # 1. Find and focus the browser window
                currentrecord=False
                self.typing_text.set(f"Finding window: '{AppName}'")
                #get item on table
                textreader=pd.read_excel
                curItem = self.table.focus()
                comp=self.table.item(curItem)["values"]
                df=pd.read_excel(file_path,index_col=0)
                lastindex=df.last_valid_index()
                data=df[df.index==comp[0]]
                phonenumber=f'+{data["Phone-Number"][comp[0]]}'
                name=f'call-{comp[1]}-{comp[0]}.wav'
                try:
                    d=carrier._is_mobile(number_type(phonenumbers.parse(phonenumber)))
                    try:
                        app = Application(backend="uia").connect(title_re=f".*{AppName}.*", timeout=15,found_index=0)
                        target_window = app.top_window()
                    except ElementNotFoundError:
                        self.typing_text.set(f"ERROR: Window with title '{AppName}' NOT FOUND.")
                        print(f"Could not find a window with title containing '{AppName}'. Please check the title and ensure the window is open.")
                        time.sleep(5)
                        continue # Skip to the next loop iteration to try again
                    self.typing_text.set("Bringing window to front...")
                    target_window.set_focus()
                    time.sleep(0.5)
                    # 2. Move mouse to target coordinates
                    self.typing_text.set(f"Moving mouse to ({X0}, {Y0})")
                    pyautogui.moveTo(X0, Y0, duration=0.25)
                    # 3. Click to focus the text box
                    self.typing_text.set("Clicking to focus...")
                    pyautogui.click()
                    time.sleep(0.5)
                    self.typing_text.set(f"Moving mouse to ({X1}, {Y1})")
                    pyautogui.moveTo(X1, Y1, duration=0.25)
                    pyautogui.click()
                    time.sleep(0.5)
                    # 4. Clear the text field
                    self.typing_text.set("Clearing text field (Ctrl+A, Del)...")
                    pyautogui.hotkey('ctrl', 'a')
                    pyautogui.press('delete')
                    time.sleep(0.5)
                    # 5. Type the number directly
                    self.typing_text.set(f"Typing: {phonenumber}")
                    pyautogui.write(phonenumber, interval=0.05)
                    # 6. Pause before pressing Enter
                    self.typing_text.set("Pausing for 1 second...")
                    time.sleep(1)
                    # 7. Press Call
                    self.typing_text.set("Pressing call button")
                    pyautogui.moveTo(X2, Y2, duration=0.25)
                    pyautogui.click()
                    time.sleep(0.5)
                    # 8. Wait for the next loop
                    self.typing_text.set("Waiting...")
                    n=0
                    while currentrecord==False:
                        self.doaudio(name)
                        self.countdown_end_time = time.time() + LOOP_DELAY_SECONDS
                        time.sleep(LOOP_DELAY_SECONDS)
                        self.countdown_end_time = None
                        self.doaudio(name)
                        n=n+1
                        if n<3:
                            df.loc[comp[0],"Status"]='Unanswered'
                    if (df.loc[comp[0],"Status"]!='Unanswered'):
                        df.loc[comp[0],"Status"]='Done'
                    self.typing_text.set("Ending Call")
                    pyautogui.moveTo(X3, Y3, duration=0.25)
                    pyautogui.click()
                    time.sleep(0.5)
                    self.stopaudio(f'audio\{name}')
                    df.loc[comp[0],"Call-File"]=name
                    df.to_csv(f'{file_path}.csv')
                    self.updatespreadsheet()
                    if (comp[0]<lastindex):
                        nextitem=self.table.get_children()[int(data.index[0]+1)]
                        self.table.focus(nextitem)
                        self.table.selection_set(nextitem)
                    else:
                        self.stop_automation()
                        self.typing_text.set("spreadsheetover")
                except NumberParseException:
                    print(data.loc[comp[0],"Status"])
                    df.loc[comp[0],"Status"]='is not number'
                    df.to_csv(f'{file_path}.csv')
                    self.updatespreadsheet()
                if (comp[0]<lastindex):
                    nextitem=self.table.get_children()[int(data.index[0]+1)]
                    self.table.focus(nextitem)
                    self.table.selection_set(nextitem)
                else:
                    self.stop_automation()
                    self.typing_text.set("spreadsheetover") 
            except Exception as e:
                if not self.is_running: break
                self.typing_text.set("ERROR: An unexpected action failed.")
                print(f"An error occurred: {e}")
                time.sleep(3)
        print("Automation loop has ended.")
        if self.is_running:
            self.stop_automation()
if __name__ == "__main__":
    app = AutomationApp()
    app.mainloop()
