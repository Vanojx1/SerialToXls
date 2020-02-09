from serial.tools.list_ports import comports
from serial import (
    Serial,
    SerialException,
    FIVEBITS, SIXBITS, SEVENBITS, EIGHTBITS,
    PARITY_NONE, PARITY_EVEN, PARITY_ODD, PARITY_MARK, PARITY_SPACE,
    STOPBITS_ONE, STOPBITS_ONE_POINT_FIVE, STOPBITS_TWO
)
from threading import Thread, currentThread
import tkinter as tk
import xlsxwriter
from datetime import datetime
import re
from tkinter.scrolledtext import ScrolledText
from tkinter import N, S, W, E


class MainApplication(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent.title('Serial to XLS')
        self.parent.resizable(False, False)

        self.pack(pady=20, padx=40)

        # port
        self.port_var = tk.StringVar(self.parent)
        self.ports = {p.device: p for p in comports()}
        self.port_var.set(list(self.ports.keys()).pop(0))
        port_label = tk.Label(self, text="Port")
        port_menu = tk.OptionMenu(self, self.port_var, *self.ports.keys())

        # baudrate
        self.baud_var = tk.IntVar(self.parent)
        self.baud_var.set(9600)
        baud_label = tk.Label(self, text='Baudrate')
        baud_field = tk.Entry(self)
        baud_field.insert(0, 9600)

        # bytesize
        self.byte_var = tk.IntVar(self.parent)
        self.byte_var.set(SEVENBITS)
        byte_label = tk.Label(self, text='Bytetype')
        byte_menu = tk.OptionMenu(
            self, self.byte_var, *[FIVEBITS, SIXBITS, SEVENBITS, EIGHTBITS])

        # parity
        self.parity_var = tk.StringVar(self.parent)
        self.parity_var.set(PARITY_EVEN)
        parity_label = tk.Label(self, text='Parity')
        parity_menu = tk.OptionMenu(
            self, self.parity_var, *[PARITY_NONE, PARITY_EVEN, PARITY_ODD, PARITY_MARK, PARITY_SPACE])

        # stopbit
        self.stopbit_var = tk.DoubleVar(self.parent)
        self.stopbit_var.set(STOPBITS_ONE)
        stopbit_label = tk.Label(self, text='Stopbit')
        stopbit_menu = tk.OptionMenu(
            self, self.stopbit_var, *[STOPBITS_ONE, STOPBITS_ONE_POINT_FIVE, STOPBITS_TWO])

        # logarea
        self.log_area = ScrolledText(self.parent, width=40, height=10)

        self.button = tk.Button(
            self, text="Start", fg="red", height=5, width=10, command=self.toggle)
        
        
        port_label.grid(row=0, column=0)
        port_menu.grid(row=0, column=1, sticky=E+W)
        baud_label.grid(row=1, column=0)
        baud_field.grid(row=1, column=1, sticky=E+W)
        byte_label.grid(row=2, column=0)
        byte_menu.grid(row=2, column=1, sticky=E+W)
        parity_label.grid(row=3, column=0)
        parity_menu.grid(row=3, column=1, sticky=E+W)
        stopbit_label.grid(row=4, column=0)
        stopbit_menu.grid(row=4, column=1, sticky=E+W)
        self.button.grid(row=0, column=2, rowspan=5, sticky=N+S+E)
        self.log_area.pack()

        self.running = False
        self.current_port = None
        self.thread = None

    def log(self, text):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.INSERT, text + '\n')
        self.log_area.config(state='disabled')

    def handle_data(self, data):
        if data != '':
            self.log('Ricevuto: %s' % data)
            A = ord('A')
            for i, match in enumerate(re.findall(r'[0-9]{4}', data)):
                self.worksheet.write(
                    'A%s' % self.row_count, datetime.now().strftime("%m/%d/%Y %H:%M:%S"))
                self.worksheet.write_number(
                    '%s%s' % (chr(A+i+1), self.row_count), int(match))
            self.row_count += 1
            # print('RECEIVED', data)

    def read_from_port(self, port, baudrate, byte, parity, stopbit):
        # print(port, baudrate, byte, parity, stopbit)
        s = Serial(
            port,
            baudrate=baudrate,
            bytesize=byte,
            parity=parity,
            stopbits=stopbit,
            timeout=0
        )
        t = currentThread()
        self.log('Porta: %s|%s|%s|%s|%s connessa' % (port, baudrate, byte, parity, stopbit))
        while getattr(t, "do_run", True):
            try:
                reading = s.readline().decode()
                self.handle_data(reading)
            except SerialException as e:
                self.toggle()
                self.log('Errore di comunicazione')
        s.close()
        self.log('Porta disconnessa')

    def toggle(self):
        self.running = not self.running
        self.button.configure(text='Stop' if self.running else 'Start')

        if self.running:
            xls_title = 'Letture_%s.xlsx' % datetime.now().strftime("%m%d%Y%H%M%S")
            self.workbook = xlsxwriter.Workbook(xls_title)
            self.worksheet = self.workbook.add_worksheet()
            self.log('Xls: %s' % xls_title)
            self.row_count = 1
            self.current_port = self.ports[self.port_var.get()]
            self.thread = Thread(target=self.read_from_port,
                                 args=(self.current_port.device, self.baud_var.get(), self.byte_var.get(), self.parity_var.get(), self.stopbit_var.get()))
            self.thread.start()
        else:
            self.thread.do_run = False
            self.workbook.close()
            self.log('Xls salvato')


if __name__ == "__main__":
    root = tk.Tk()
    MainApplication(root).pack(side="top", fill="both", expand=True)
    root.mainloop()
