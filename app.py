from serial.tools.list_ports import comports
from serial import Serial
from threading import Thread, currentThread
import Tkinter as tk
import xlsxwriter
from datetime import datetime
import re


class MainApplication(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent.title('Serial to XLS')
        self.parent.resizable(False, False)

        self.pack(pady=20, padx=40)
        
        self.tkvar = tk.StringVar(self.parent)
        self.ports = {p.device: p for p in comports()}
        self.tkvar.set(self.ports.keys()[0]) # set the default option
        self.popupMenu = tk.OptionMenu(self, self.tkvar, *self.ports.keys())
        
        tk.Label(self, text="Seleziona la porta").grid(row=1, column=1)
        
        self.popupMenu.grid(row=2, column=1)

        self.button = tk.Button(self, text="Start", fg="red", height=5, width=10, command=self.toggle)
       	self.button.grid(row=1, column=3, rowspan=2)
       	
       	self.running = False
       	self.current_port = None
       	self.thread = None

    def handle_data(self, data):
        if data != '':
            A = ord('A')
            for i, match in enumerate(re.findall(r'[0-9]{4}', data)):
                self.worksheet.write('A%s' % self.row_count, datetime.now().strftime("%m/%d/%Y %H:%M:%S"))
                self.worksheet.write_number('%s%s' % (chr(A+i+1), self.row_count), int(match))
            self.row_count += 1
            # print('RECEIVED', data)

    def read_from_port(self, port):
        s = Serial(port.device, timeout=0)
        t = currentThread()
        while getattr(t, "do_run", True):
            reading = s.readline().decode()
            self.handle_data(reading)
        s.close()

    def toggle(self):
    	self.running = not self.running
    	self.button.configure(text='Stop' if self.running else 'Start')

    	if self.running:
            xls_title = 'Letture_%s.xlsx' % datetime.now().strftime("%m%d%Y%H%M%S")
            self.workbook = xlsxwriter.Workbook(xls_title)
            self.worksheet = self.workbook.add_worksheet()
            self.row_count = 1
            self.current_port = self.ports[self.tkvar.get()]
            self.thread = Thread(target=self.read_from_port, args=(self.current_port,))
            self.thread.start()
        else:
            self.thread.do_run = False
            self.workbook.close()

if __name__ == "__main__":
    root = tk.Tk()
    MainApplication(root).pack(side="top", fill="both", expand=True)
    root.mainloop()