import tkinter as tk 
from tkinter import messagebox

import sys
import os

import pyodbc

import serial
import threading
import queue

from datetime import datetime
from decimal import *

import socket
import time

#CONFIGURACION 
BAUDRATE = 9600
TIMEOUT = 1

SERVER = 'localhost\\SQLEXPRESS'  # Puede ser una dirección IP o nombre del servidor
DATABASE = 'db_ensacadora'

#COLA PARA ALMACENAR DATOS SIN PERDERLOS
data_queue = queue.Queue()

def get_port():
    port = 'COM1'
    text = 'DEMO DEMO'
    try:
        # Conectar a SQL Server local
        conn = pyodbc.connect(f'DRIVER={{ODBC Driver 18 for SQL Server}};SERVER={SERVER};DATABASE={DATABASE};Trusted_Connection=yes;Encrypt=no')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM puertos")
        result = cursor.fetchone()
        cursor.close()
        conn.close()

        if result:
            port = result.num_puerto
            text = result.ct

    except:
        pass

    data = {
        'port': port,
        'text': text
    }

    return data
  
class indexGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("EnsacTrack")

        # IMG
        if getattr(sys, 'frozen', False):  # Si está ejecutándose como .exe
            base_path = sys._MEIPASS
        else:
            base_path = os.path.abspath(".")

        image_path = os.path.join(base_path, "logo.png")

        self.icon = tk.PhotoImage(file=image_path)
        self.iconphoto(False, self.icon)

        #PANTALLA COMPLETA
        self.attributes('-fullscreen', True)
        self.bind('<Escape>', self.no_close)

        #BLOQUEO DE CIERRE DE PANTALLA
        self.protocol("WM_DELETE_WINDOW", self.no_close)

        #PUERTO 
        self.serial_port = ''
        self.stop_thread = threading.Event()
        
        #MENU
        self.create_menu()

        #CREAR INTERFAZ
        self.show_init()

        self.start_reading()

    def create_menu(self):
        menu_bar = tk.Menu(self)
        
        edit_menu = tk.Menu(menu_bar, tearoff=0)
        edit_menu.add_command(label="PUERTO", command=self.edit_port)
        
        menu_bar.add_cascade(label='Configuración', menu=edit_menu)

        self.config(menu=menu_bar)

    def edit_port(self):
        # Crear una ventana para editar el puerto (puedes personalizarlo)
        self.edit_window = tk.Toplevel(self)
        self.edit_window.title("Editar PUERTO")

        self.edit_window.iconphoto(False, self.icon)

        label = tk.Label(self.edit_window, text="Introduce el puerto:")
        label.pack(pady=10)

        entry_port = tk.Entry(self.edit_window, width=30)
        entry_port.pack(pady=5)

        label_password = tk.Label(self.edit_window, text="Introduce el password:")
        label_password.pack(pady=10)

        entry_password = tk.Entry(self.edit_window, width=30, show="*")
        entry_password.pack(pady=5)

        button_save = tk.Button(self.edit_window, text="Guardar", command=lambda: self.save_port(entry_port.get(), entry_password.get()))
        button_save.pack(pady=10)

    def save_port(self, port, password):
        if password == 'GC25ENSAC':
            self.stop_thread.set()
            self.thread.join(timeout=3)

            try:

                conn = pyodbc.connect(f'DRIVER={{ODBC Driver 18 for SQL Server}};SERVER={SERVER};DATABASE={DATABASE};Trusted_Connection=yes;Encrypt=no')
                cursor = conn.cursor()
                query = """
                        UPDATE puertos
                        SET num_puerto = ?
                        WHERE id = 1
                    """

                # Ejecutar la consulta con los valores
                cursor.execute(query, port)

                # Confirmar los cambios
                conn.commit()

                # Cerrar la conexión
                cursor.close()
                conn.close()

            except:
                messagebox.showerror("Error", "Ha ocurrido un error en el sistema al actualizar el PUERTO")

            #PORT
            init_data = get_port()
            self.serial_port = init_data['port']
            #HILO PARA LECTURA EN SEGUNDO PLANO
            self.stop_thread.clear()
            self.thread = threading.Thread(target=self.read_serial, daemon=True)
            self.thread.start()

            self.edit_window.destroy()

        else: 
            messagebox.showerror("Error", "PASSWORD INCORRECTO")

    #GUI
    def show_init(self):
        self.main_frame = tk.Frame(self, bg="#2c3e50")
        self.main_frame.pack(side="right", expand=True, fill="both")

        self.title_lable = tk.Label(self.main_frame, text="EnsacTrack", font=("Arial", 30, "bold"), fg="white", bg="#2c3e50")
        self.title_lable.pack(pady=20)

        self.display_data = tk.Label(self.main_frame, text='------', font=("Arial", 40, "bold"), fg="#000000", bg="white", height=4)
        self.display_data.pack(fill="x", pady=20, padx=20)

        self.display_history_title = tk.Label(self.main_frame, text="Último Registro", font=("Arial", 12), fg="white", bg="#2c3e50")
        self.display_history_title.pack(pady=5)
        
        self.display_history_data = tk.Label(self.main_frame, text="------", font=("Arial", 16), fg="#000000", bg="#d3d3d3", height=2)
        self.display_history_data.pack(fill="x", pady=5, padx=20)
        
        #ENSACADORA
        self.frame_ensacadora = tk.Frame(self.main_frame, bg="#2c3e50")
        self.frame_ensacadora.pack(fill="x", pady=5, padx=20)

        self.status_ensacadora = tk.Label(self.frame_ensacadora, text="Estatus Conexión Ensacadora: ", font=("Arial", 12), fg="white", bg="#2c3e50")
        self.status_ensacadora.pack(side="left", pady=5)

        self.status_ensacadora_data = tk.Label(self.frame_ensacadora, text="------", font=("Arial", 16, "bold"), fg="white", bg="#2c3e50")
        self.status_ensacadora_data.pack(side="left", pady=5)

        #DB
        self.frame_db = tk.Frame(self.main_frame, bg="#2c3e50")
        self.frame_db.pack(fill="x", pady=5, padx=20)

        self.status_db = tk.Label(self.frame_db, text="Estatus Conexión DB: ", font=("Arial", 12), fg="white", bg="#2c3e50")
        self.status_db.pack(side="left", pady=5)

        self.status_db_data = tk.Label(self.frame_db, text="------", font=("Arial", 16, "bold"), fg="white", bg="#2c3e50")
        self.status_db_data.pack(side="left", pady=5)
        
        #INTERNET
        self.frame_online = tk.Frame(self.main_frame, bg="#2c3e50")
        self.frame_online.pack(fill="x", pady=5, padx=20)

        self.status_online = tk.Label(self.frame_online, text="Estatus Conexión Internet: ", font=("Arial", 12), fg="white", bg="#2c3e50")
        self.status_online.pack(side="left", pady=5)

        self.status_online_data = tk.Label(self.frame_online, text="------", font=("Arial", 16, "bold"), fg="white", bg="#2c3e50")
        self.status_online_data.pack(side="left", pady=5)

        #PUERTO
        self.frame_puerto= tk.Frame(self.main_frame, bg="#2c3e50")
        self.frame_puerto.pack(fill="x", pady=5, padx=20)

        self.status_puerto = tk.Label(self.frame_puerto, text="Puerto: ", font=("Arial", 12), fg="white", bg="#2c3e50")
        self.status_puerto.pack(side="left", pady=5)

        self.status_puerto_data = tk.Label(self.frame_puerto, text="------", font=("Arial", 16, "bold"), fg="white", bg="#2c3e50")
        self.status_puerto_data.pack(side="left", pady=5)

        self.display_sms = tk.Text(self.main_frame, font=("Arial", 12), fg="#000000", bg="#d3d3d3", height=2, wrap="word")
        self.display_sms.pack(fill="x", pady=5, padx=20)

    def start_reading(self):
        init_data = get_port() 
        #TITLE
        self.title_lable.config(text=f"EnsacTrack - {init_data['text']}")

        #PUERTO
        self.serial_port = init_data['port']
        
        #HILO PARA LECTURA EN SEGUNDO PLANO
        self.thread = threading.Thread(target=self.read_serial, daemon=True)
        self.thread.start()

        #SEGUNDO HILO PARA PROCESAR DATOS
        process_thread = threading.Thread(target=self.process_data, daemon=True)
        process_thread.start()

        #TERCER HILO PARA VALIDACION ONLINE
        verification_online_thread = threading.Thread(target=self.verification_online, daemon=True)
        verification_online_thread.start()
    
    #LECTURA DEL PUERTO
    def read_serial(self):
        self.status_puerto_data.config(text=self.serial_port)
        #FUNCION QUE LEE CONTINUAMENTE LOS DATOS DEL PUERTO
        try:
            with serial.Serial(self.serial_port, BAUDRATE, timeout=TIMEOUT) as ser:
                text_status_ensacadora = 'CONECTADO'
                self.update_status_ensacadora(text_status_ensacadora, 1)
                while not self.stop_thread.is_set(): 
                    try:
                        data = ser.readline().decode('utf-8').strip()
                        if data:
                            data_queue.put(data)  # Guardar en la cola
                            text_status_ensacadora = 'CONECTADO'
                            self.update_status_ensacadora(text_status_ensacadora, 1)

                    except Exception as e:
                        text_sms = f"Error al leer el puerto {self.serial_port}: {e}\n"
                        self.update_sms(text_sms)
                        text_status_ensacadora = 'DESCONECTADO'
                        self.update_status_ensacadora(text_status_ensacadora, 2)

                ser.close()
        
        except serial.SerialException as e:
            text_sms = f"No se pudo abrir el puerto {self.serial_port}: {e}\n"
            self.update_sms(text_sms)
            text_status_ensacadora = 'DESCONECTADO'
            self.update_status_ensacadora(text_status_ensacadora, 2)

    #PROCESAMIENTO DE DATOS
    def process_data(self):
        #FUNCION PARA PROCESAR DATOS SIN BLOQUEAR LA LECTURA.
        try:
            conn = pyodbc.connect(f'DRIVER={{ODBC Driver 18 for SQL Server}};SERVER={SERVER};DATABASE={DATABASE};Trusted_Connection=yes;Encrypt=no')
            cursor = conn.cursor()
                
            #INIT
            #DB
            text_status_db_data = 'CONECTADO'
            self.update_status_db_data(text_status_db_data, 1)
            
            #HISTORY
            query = 'SELECT TOP 1 * FROM register_data ORDER BY id DESC;'
            cursor.execute(query)
            result = cursor.fetchall()
            if not result: 
                result = '------'

            else: 
                result = f'FECHA / HORA: {result[0].date_time_register.strftime("%Y-%m-%d / %H:%M:%S")}  -  PESO: {result[0].kilograms} KG  -  NUMERO: {result[0].number_register}'

            self.display_history_data.config(text=result)
            

            while True:
                try:
                    data = data_queue.get(timeout=5) # Espera hasta que haya datos
                    valores = str(data).split(',')

                    fecha = valores[0] # FECHA
                    hora = valores[1] # HORA
                    cantidad = valores[2] # CANTIDAD
                    peso = valores[3] # PESO
                    unidad = valores[4] # kg

                    self.display_data.config(text=f'No. {cantidad} / Peso: {peso} {unidad}')

                    try:
                        insert_data = 'INSERT INTO register_data (date_time_register, date_time_ensac, kilograms, number_register) VALUES (?, ?, ?, ?)'
                        
                        hoy = datetime.now()
                        date_ensac = datetime.strptime(f'{fecha} {hora}', "%d/%m/%Y %H:%M")
                        cursor.execute(insert_data, (hoy, date_ensac, Decimal(peso), cantidad))
                        # Confirmar la transacción
                        conn.commit()

                        #HISTORY
                        query = 'SELECT TOP 1 * FROM register_data ORDER BY id DESC;'
                        cursor.execute(query)
                        result = cursor.fetchall()
                        if not result: 
                            result = '------'

                        else: 
                            result = f'FECHA / HORA: {result[0].date_time_register.strftime("%Y-%m-%d / %H:%M:%S")}  -  PESO: {result[0].kilograms} KG  -  NUMERO: {result[0].number_register}'

                        self.display_history_data.config(text=result)
                        
                        #DB
                        text_status_db_data = 'CONECTADO'
                        self.update_status_db_data(text_status_db_data, 1)

                    except Exception as e:
                        text_sms = f"Sin conexión a la base de datos: {e}\n"
                        self.update_sms(text_sms)
                        text_status_db_data = 'DESCONECTADO'
                        self.update_status_db_data(text_status_db_data, 2)

                except queue.Empty:
                    self.display_data.config(text='------')

        except Exception as e: 
            text_sms = f"Sin conexión a la base de datos: {e}\n"
            self.update_sms(text_sms)
            text_status_db_data = 'DESCONECTADO'
            self.update_status_db_data(text_status_db_data, 2)
            
    #COMPROBACION CONEXION A INTERNET
    def verification_online(self):
        while True:
            text = ''
            id_color = 2
            try:
                socket.create_connection(("8.8.8.8", 53), timeout=3)
                text = 'CONECTADO'
                id_color = 1

            except:
                text = 'DESCONECTADO'
                
            self.update_status_online(text, id_color)
            time.sleep(5)

    #ESTATDOS
    def update_sms(self, text):
        self.display_sms.insert(tk.END, text)

    def update_status_ensacadora(self, text, id_color):
        if id_color == 1:
            self.status_ensacadora_data.config(text=text, fg='#008f39')

        elif id_color == 2:
            self.status_ensacadora_data.config(text=text, fg='#ff0000')
            
    def update_status_db_data(self, text, id_color):
        if id_color == 1:
            self.status_db_data.config(text=text, fg='#008f39')

        elif id_color == 2:
            self.status_db_data.config(text=text, fg='#ff0000')

    def update_status_online(self, text, id_color):
        if id_color == 1:
            self.status_online_data.config(text=text, fg='#008f39')

        elif id_color == 2:
            self.status_online_data.config(text=text, fg='#ff0000')
    
    #CLOSE
    def no_close(self, event=None):
        messagebox.showwarning("Advertencia", "No puedes cerrar la aplicación.")

if __name__ == "__main__":
    login = indexGUI()
    login.mainloop()