#-*-coding:UTF-8 -*
import tkinter as tk
import tkinter.font
import tkinter.messagebox
from tkinter.ttk import *
import pyodbc
import os, wmi, subprocess, psutil, socket, platform
from time import sleep
from PIL import Image, ImageTk
import webbrowser
from threading import Thread
import getpass
import pickle
from ldap3 import Server, Connection, ALL, NTLM
from cryptography.fernet import Fernet
from datetime import date, datetime, timedelta


class Application(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.create_widgets()
        self.config(bg="black")
        
    
    def create_widgets(self):
        self.option_add("*Frame*background", "black")
        self.option_add("*Button*background", "grey")
        self.option_add("*Label*foreground", "white")
        self.option_add("*Label*background", "black")
        self.option_add("*Button*width", 15)
        self.option_add("*Label*width", 20)
        self.option_add("*Label*anchor", 'w')
        self.option_add("*Text*width", 30)
        self.option_add("*Text*background", 'black')
        self.option_add("*Text*foreground", 'white')
        self.option_add("*Text*font", ('Calibri', '9'))
        self.option_add("*Text*height", 1)
        self.option_add("*Text*selectbackground", 'grey')

        #MENU
        self.menu = tk.Menu(self)
        self['menu'] = self.menu
        self.MenuFichier = tk.Menu(self.menu, tearoff=0)
        self.MenuApplications = tk.Menu(self.menu, tearoff=0)
        self.MenuAdmin = tk.Menu(self.menu, tearoff=0)
        self.MenuAide = tk.Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label='Fichier', menu=self.MenuFichier)
        self.menu.add_cascade(label='Admin', menu=self.MenuAdmin)
        self.menu.add_cascade(label='Applications', menu=self.MenuApplications)
        self.menu.add_cascade(label='Aide', menu=self.MenuAide)
        self.MenuAide.add_command(label='Historique')
        self.MenuAide.add_command(label='A propos')
        self.MenuAide.add_command(label='Aide')
        self.MenuApplications.add_command(label='Smart', command= self.Open_Smart)
        self.MenuApplications.add_command(label='GRC', command= self.Open_GRC_Acc)
        self.MenuAdmin.add_command(label='Iguazu', command= self.Open_Iguazu)
        self.MenuAdmin.add_command(label='Wifi Guest', command= self.Open_WifiGuest)
        self.MenuAdmin.add_command(label='ServicePilot', command= self.Open_SPilot)
        self.MenuAdmin.add_command(label='Wiki TA', command= self.Open_Wiki)
        self.MenuAdmin.add_command(label='Console AD', command= self.ADConsol)

        #FRAME PRINCIPALE
        self.frame_1 = tk.Frame(self, width=500, height=230)
        self.label_1 = tk.Label(self.frame_1, text="Poste/Trigramme :")
        self.entry_rech = tk.Entry(self.frame_1, bg="white")
        self.btn_ok = tk.Button(self.frame_1, text="OK", command=self.WriteInfo)

        # FRAME INFO 
        self.frame_info = tk.LabelFrame(self, text="Informations", labelanchor='nw', width=500, height=230,  bg="black", fg="white")
        self.label_name = tk.Label(self.frame_info, text="Nom :")
        self.txt_name1 = tk.Text(self.frame_info, bd=0)
        self.label_idwin = tk.Label(self.frame_info, text="Identifiant Windows :")
        self.txt_idwin1 = tk.Text(self.frame_info, bd=0)
        self.label_idsig = tk.Label(self.frame_info, text="Identifiant Sigma :")
        self.txt_idsig1 = tk.Text(self.frame_info, bd=0)
        self.label_idgrc = tk.Label(self.frame_info, text="Identifiant GRC :")
        self.txt_idgrc1 = tk.Text(self.frame_info, bd=0)
        self.label_computer = tk.Label(self.frame_info, text="Poste :")
        self.combo = tk.ttk.Combobox(self.frame_info)
        self.label_serv = tk.Label(self.frame_info, text="Service :")
        self.txt_serv1 = tk.Text(self.frame_info, bd=0)
        self.label_address = tk.Label(self.frame_info, text="Ville :")
        self.txt_address1 = tk.Text(self.frame_info, width=45, bd=0)
        self.label_phone = tk.Label(self.frame_info, text="Téléphone :")
        self.txt_phone1 = tk.Text(self.frame_info, bd=0)
        self.label_email = tk.Label(self.frame_info, text="Email :")
        self.txt_email1 = tk.Text(self.frame_info, bd=0, width=45)

        self.btn_grc = tk.Button(self.frame_info, text="y", font="Wingdings", width=3, command=self.Open_GrC)
        self.btn_remote = tk.Button(self.frame_info, text=":", font="Wingdings", width=3, command=self.Remote)
        
        #FRAME WINDOWS
        self.frame_compte = tk.LabelFrame(self, text="Compte Windows", labelanchor='nw', width=500, height=100, bg='black', fg='white')
        self.label_wind = tk.Label(self.frame_compte, text="Trigramme :", width=10)
        self.entry_win = tk.Entry(self.frame_compte, bg='white')
        self.btn_win_rech = tk.Button(self.frame_compte, text='Rechercher', command=self.UseCheckAD)
        self.lbl_pwd = tk.Label(self.frame_compte, text='Expiration mot de passe :')
        self.lbl_state = tk.Label(self.frame_compte, text= 'Etat du compte :')
        self.txt_pwd = tk.Text(self.frame_compte, bd=0, width=20)
        self.txt_state = tk.Text(self.frame_compte, bd=0, width=20)

        #FRAME POSTE
        self.frame_poste = tk.LabelFrame(self, text="Poste", labelanchor='nw', width=500, height=150, bg='black', fg='white')
        self.label_poste = tk.Label(self.frame_poste, text="Poste :")
        self.combo_poste = tk.ttk.Combobox(self.frame_poste)
        self.btn_poste_rech = tk.Button(self.frame_poste, text='Rechercher', command=self.WriteInfoPoste)
        self.lbl_poste_ip = tk.Label(self.frame_poste, text='Adresse IP :')
        self.lbl_poste_mod = tk.Label(self.frame_poste, text='Modèle :')
        self.lbl_poste_os = tk.Label(self.frame_poste, text='Version OS :')
        self.lbl_poste_boot = tk.Label(self.frame_poste, text='Dernier Démarrage :')
        self.lbl_poste_dd = tk.Label(self.frame_poste, text='Disque dur :')
        self.txt_poste_ip = tk.Text(self.frame_poste, bd=0)
        self.txt_poste_mod = tk.Text(self.frame_poste, bd=0)
        self.txt_poste_os = tk.Text(self.frame_poste, bd=0)
        self.txt_poste_boot = tk.Text(self.frame_poste, bd=0)
        self.txt_poste_dd = tk.Text(self.frame_poste, bd=0)


        #FRAME GENESYS
        self.frame_genesys = tk.LabelFrame(self, text='Genesys', labelanchor='nw', width=500, height=230, bg='black', fg='white')
        self.label_genid = tk.Label(self.frame_genesys, text='Trigramme :', width=10)
        self.entry_gen_id = tk.Entry(self.frame_genesys, bg='white', width=15)
        self.label_gennum = tk.Label(self.frame_genesys, text='Numéro :', width=10)
        self.entry_gen_num = tk.Entry(self.frame_genesys, bg='white', width=15)
        self.onglet = tk.ttk.Notebook(self.frame_genesys, width=470, height=100)
        self.onglet_1 = tk.ttk.Frame(self.onglet)
        self.onglet_2 = tk.ttk.Frame(self.onglet)
        self.btn_gen_id = tk.Button(self.frame_genesys, text='Rechercher', width=12, command=self.AlimTabGenID)
        self.btn_gen_num = tk.Button(self.frame_genesys, text='Rechercher', width=12, command= self.AlimTabGenNum)
        
        self.tableau_id = tk.ttk.Treeview(self.onglet_1, columns=('num', 'trigramme','nom','prenom','login','logout'), height=3, selectmode='browse')
        self.scrollbar_id_x = tk.ttk.Scrollbar(self.onglet_1, orient='horizontal', command=self.tableau_id.xview)
        self.scrollbar_id_y = tk.ttk.Scrollbar(self.onglet_1, orient='vertical', command=self.tableau_id.yview)
        #self.scrollbar_id.grid(row=1, column=0, sticky='we')
        self.tableau_id.configure(xscrollcommand=self.scrollbar_id_x.set)
        self.tableau_id.configure(xscrollcommand=self.scrollbar_id_y.set)

        self.tableau_id.heading('num', text='Num')
        self.tableau_id.column('num', minwidth=0, width=45, stretch='NO')
        self.tableau_id.heading('trigramme', text='Trigramme')
        self.tableau_id.column('trigramme', minwidth=10, width=75, stretch='NO')
        self.tableau_id.heading('nom', text='Nom')
        self.tableau_id.column('nom', minwidth=10, width=70, stretch='NO')
        self.tableau_id.heading('prenom', text='Prénom')
        self.tableau_id.column('prenom', minwidth=10, width=60, stretch='NO')
        self.tableau_id.heading('login', text='Login')
        self.tableau_id.column('login', minwidth=10, width=100, stretch='NO')
        self.tableau_id.heading('logout', text='Logout')
        self.tableau_id.column('logout', minwidth=10, width=100, stretch='NO')
        self.tableau_id['show'] = 'headings'
        
        self.tableau_num = tk.ttk.Treeview(self.onglet_2, columns=('num', 'trigramme','nom','prenom','login','logout'), height=3, selectmode='browse')
        self.scrollbar_num_x = tk.ttk.Scrollbar(self.onglet_2, orient='horizontal', command=self.tableau_num.xview)
        self.scrollbar_num_y = tk.ttk.Scrollbar(self.onglet_2, orient='vertical', command=self.tableau_num.yview)
        #self.scrollbar_id.grid(row=1, column=0, sticky='we')
        self.tableau_num.configure(xscrollcommand=self.scrollbar_num_x.set)
        self.tableau_num.configure(xscrollcommand=self.scrollbar_num_y.set)

        self.tableau_num.heading('num', text='Num')
        self.tableau_num.column('num', minwidth=0, width=45, stretch='NO')
        self.tableau_num.heading('trigramme', text='Trigramme')
        self.tableau_num.column('trigramme', minwidth=10, width=75, stretch='NO')
        self.tableau_num.heading('nom', text='Nom')
        self.tableau_num.column('nom', minwidth=10, width=70, stretch='NO')
        self.tableau_num.heading('prenom', text='Prénom')
        self.tableau_num.column('prenom', minwidth=10, width=60, stretch='NO')
        self.tableau_num.heading('login', text='Login')
        self.tableau_num.column('login', minwidth=10, width=100, stretch='NO')
        self.tableau_num.heading('logout', text='Logout')
        self.tableau_num.column('logout', minwidth=10, width=100, stretch='NO')
        self.tableau_num['show'] = 'headings'
        
        self.btn_quit = tk.Button(self, text="Quitter", command = self.quit)
        
        # PLACEMENT GRID
        self.frame_1.grid()
        self.label_1.grid(row=0, column=0)
        self.entry_rech.grid(row=0, column=1)
        self.btn_ok.grid(row=0, column=3, padx=20)
        self.frame_info.grid(row=1 ,column=0, padx=10)
        self.frame_info.grid_propagate(0)

        self.label_name.grid(row=0, column=0, sticky='w')
        self.txt_name1.grid(row=0, column=1, sticky='w')
        self.btn_remote.grid(row=4, column=2)
        self.label_idwin.grid(row=1, column=0, sticky='w')
        self.txt_idwin1.grid(row=1, column=1, sticky='w')
        self.label_idsig.grid(row=2, column=0, sticky='w')
        self.txt_idsig1.grid(row=2, column=1, sticky='w')
        self.label_idgrc.grid(row=3, column=0, sticky='w')
        self.txt_idgrc1.grid(row=3, column=1, sticky='w')
        self.btn_grc.grid(row=3,column=2)
        self.label_computer.grid(row=4, column=0, sticky='w')
        self.combo.grid(row=4, column=1, sticky='w')
        self.label_serv.grid(row=5, column=0, sticky='w')
        self.txt_serv1.grid(row=5, column=1, sticky='w')
        self.label_address.grid(row=6, column=0, sticky='w')
        self.txt_address1.grid(row=6, column=1, sticky='w')
        self.label_phone.grid(row=7, column=0, sticky='w')
        self.txt_phone1.grid(row=7, column=1, sticky='w')
        self.label_email.grid(row=8, column=0, sticky='w')
        self.txt_email1.grid(row=8, column=1, sticky='w')

        self.frame_compte.grid(row=2,column=0, padx=10)
        self.frame_compte.grid_propagate(0)
        self.label_wind.grid(row=0, column=0, sticky='w')
        self.entry_win.grid(row=0, column=1, sticky='w')
        self.btn_win_rech.grid(row=0, column=2, padx=45)
        self.lbl_pwd.grid(row=1, column=0)
        self.lbl_state.grid(row=2, column=0)
        self.txt_pwd.grid(row=1, column=1)
        self.txt_state.grid(row=2, column=1)

        self.frame_poste.grid(row=3, column=0, padx=10)
        self.frame_poste.grid_propagate(0)
        self.label_poste.grid(row=0, column=0, sticky='w')
        self.combo_poste.grid(row=0, column=1, sticky='w')
        self.btn_poste_rech.grid(row=0, column=2, padx=40)
        self.lbl_poste_ip.grid(row=2, column=0)
        self.lbl_poste_mod.grid(row=3, column=0)
        self.lbl_poste_os.grid(row=4, column=0)
        self.lbl_poste_boot.grid(row=5, column=0)
        #self.lbl_poste_dd.grid(row=6, column=0)
        self.txt_poste_ip.grid(row=2, column=1)
        self.txt_poste_mod.grid(row=3, column=1)
        self.txt_poste_os.grid(row=4, column=1)
        self.txt_poste_boot.grid(row=5, column=1)
        #self.txt_poste_dd.grid(row=6, column=1)

        self.frame_genesys.grid(row=4, column=0, padx=10)
        self.frame_genesys.grid_propagate(0)
        self.label_genid.grid(row=0, column=0,sticky='w')
        self.label_gennum.grid(row=0, column=2, padx=20)
        self.entry_gen_id.grid(row=0, column=1,sticky='w')
        self.entry_gen_num.grid(row=0, column=3,sticky='w')
        self.onglet.add(self.onglet_1, text='Utilisateur')
        self.onglet.add(self.onglet_2, text='Numéro')
        self.onglet.grid(row=2, columnspan=4, sticky='w', pady=10, padx=10)
        self.btn_gen_id.grid(row=1, column=0, columnspan=2, pady=5)
        self.btn_gen_num.grid(row=1, column=2, columnspan=2, pady=5)
       
        self.tableau_id.grid(row=0, column=0)
        self.scrollbar_id_x.grid(row=1, column=0, sticky='we')
        self.scrollbar_id_y.grid(row=0, column=1, sticky='ns')
        self.tableau_num.grid(row=0, column=0)
        self.scrollbar_num_x.grid(row=1, column=0, sticky='we')
        self.scrollbar_num_y.grid(row=0, column=1, sticky='ns')

        self.btn_quit.grid(sticky='e',padx=10, pady=5)

    def recup_text(self):
        recup = self.entry_rech.get()
        if recup:
            return recup
        else:
            return False

    def recup_text_poste(self):
        poste = self.combo_poste.get()  
        if poste:
            return poste
        else:
            return False          

    def CreateCred(self):
        if not os.access("D:/ut/user/Scripts/Python_2/tkinter/credential",1) :
            key = Fernet.generate_key()
            admin = 'Domain\\' + getpass.getuser()[0:3]
            password= Fernet(key).encrypt(getpass('Mot de passe :').encode())
            cred = {
                "user": admin,
                "key": key,
                "pwd": password , }
            with open("D:/ut/user/Scripts/Python_2/tkinter/credential", "wb") as credential:
                picki = pickle.Pickler(credential)
                picki.dump(cred)

    def GetCred(self):
        self.CreateCred()
        with open("D:/ut/user/Scripts/Python_2/tkinter/credential", "rb") as cred:
            recup = pickle.Unpickler(cred)
            mdp_recup = recup.load()

        return Fernet(mdp_recup['key']).decrypt(mdp_recup['pwd']).decode()

    def Get_Admin(self):
        with open("D:/ut/user/Scripts/Python_2/tkinter/credential", "rb") as cred:
            recup = pickle.Unpickler(cred)
            mdp_recup = recup.load()
        return mdp_recup['user']

    def CheckAD(self, user):
        authent = self.GetCred()
        admin = self.Get_Admin()
        server = Server('Server', get_info=ALL)
        conn = Connection(server, user=admin, password=authent, authentication=NTLM)
        conn.bind()
        conn.search('OU=XX,OU=XX,OU=XX,OU=GLBR,DC=xx,DC=xx,DC=xx,DC=fr','(&(objectclass=person)(sAMAccountName=%s))' % user , attributes=['sAMAccountName','pwdLastSet', 'name','lastLogon','lockoutTime'])

        result = conn.entries
        return result

    def UseCheckAD(self):
        self.txt_state.delete(1.0, tk.END)
        self.txt_pwd.delete(1.0, tk.END)

        user = self.entry_win.get()
        data_user = self.CheckAD(user)
        verr = datetime.fromisoformat(str(data_user[0].lockoutTime)).timestamp()
        if verr < 0:
            self.txt_state.insert(1.0, "Non verrouillé")
        else:
            self.txt_state.insert(1.0, "Verrouillé")

        expire_date = datetime.fromisoformat(str(data_user[0].pwdLastSet)) + timedelta(days=90)
        self.txt_pwd.insert(1.0, expire_date.strftime('%d/%m/%Y %H:%M'))

    def ConnectSQL(self, id):
        #recup = 'user'
        #id = 'N44N100'
        try:
            Server = "Server"
            Database = 'Database'
            conn = pyodbc.connect('DRIVER={SQL Server}; SERVER='+Server+'; DATABASE='+Database+';TRUSTED_CONNECTION=True;')
            cursor = conn.cursor()
            cursor.execute("""SELECT TRIGRAMME,FIRSTNAME,LASTNAME,ID_RESSOURCE,SERVICE,CP,VILLE,TELEPHONE,EMAIL,LOGIN_SIGMA,LOGIN_GRC FROM PERSONNES
            INNER JOIN RESSOURCE ON RESSOURCE.SID_PERS = PERSONNES.SID_PERS
            INNER JOIN SITEENTREPRISE on SITEENTREPRISE.ID_SITEENTREPRISE = PERSONNES.ID_SITEENTREPRISE
            WHERE TRIGRAMME = ? OR ID_RESSOURCE =? """,id, id)

            info = cursor.fetchall()
            if not info[0]:
                tk.messagebox.showwarning(title="Information", message="Aucun résultat pour cette recherche", icon='warning')
            else:
                return info
        except AttributeError:
            tk.messagebox.showwarning(title="Information", message="Saisie incorrecte !", icon='warning')
        except IndexError:
            tk.messagebox.showwarning(title="Information", message="Saisie inconnue au bataillon !", icon='warning')

    def GenesysSQLID(self, id):
        try:
            Server_Gen = "Server"
            Database_Gen = "DATABASE" 
            conn = pyodbc.connect('DRIVER={SQL Server}; SERVER='+Server_Gen+'; DATABASE='+Database_Gen+';TRUSTED_CONNECTION=True;')
            cursor = conn.cursor()
            cursor.execute("""SELECT TOP 10 GIDB_G_LOGIN_SESSION_V.ID, GIDB_GC_AGENT.EMPLOYEEID Trigramme,
		                      GIDB_GC_AGENT.FIRSTNAME Prenom,
		                      GIDB_GC_AGENT.LASTNAME Nom,
		                      GIDB_GC_PLACE.NAME Place,
		                      GIDB_GC_ENDPOINT.DN Telephone_Associe,
		                      GIDB_GC_FOLDER.NAME Localisation,
		                      GIDB_G_LOGIN_SESSION_V.CREATED DebutDeSession,
		                      GIDB_G_LOGIN_SESSION_V.TERMINATED FinDeSession
	                    FROM GEN_GIM.dbo.GIDB_G_LOGIN_SESSION_V GIDB_G_LOGIN_SESSION_V, GEN_GIM.dbo.GIDB_GC_AGENT GIDB_GC_AGENT, GEN_GIM.dbo.GIDB_GC_ENDPOINT GIDB_GC_ENDPOINT, GEN_GIM.dbo.GIDB_GC_FOLDER GIDB_GC_FOLDER, GEN_GIM.dbo.GIDB_GC_PLACE GIDB_GC_PLACE
	                    WHERE GIDB_G_LOGIN_SESSION_V.AGENTID = GIDB_GC_AGENT.ID AND GIDB_G_LOGIN_SESSION_V.PLACEID = GIDB_GC_PLACE.ID AND GIDB_G_LOGIN_SESSION_V.PRIMARYDEVICEID = GIDB_GC_ENDPOINT.ID AND GIDB_GC_AGENT.FOLDERID = GIDB_GC_FOLDER.ID AND ((GIDB_GC_AGENT.EMPLOYEEID= ?)) ORDER BY GIDB_G_LOGIN_SESSION_V.ID DESC
            """, id)
            info_gen_id = cursor.fetchall()
            return info_gen_id
        except AttributeError:
            tk.messagebox.showwarning(title="Information", message="Saisie incorrecte !", icon='warning')

    def AlimTabGenID(self):
        self.tableau_id.delete(*self.tableau_id.get_children())
        id = self.entry_gen_id.get()
        info_gen = self.GenesysSQLID(id)
        for line in info_gen:
            self.tableau_id.insert("", 'end', text="L", 
            values= (line.Place, line.Trigramme, line.Nom, line.Prenom, line.DebutDeSession, line.FinDeSession))

    def GenesysSQLNum(self, num):
        try:
            Server_Gen = "Server"
            Database_Gen = "DATABASE" 
            conn = pyodbc.connect('DRIVER={SQL Server}; SERVER='+Server_Gen+'; DATABASE='+Database_Gen+';TRUSTED_CONNECTION=True;')
            cursor = conn.cursor()
            cursor.execute("""SELECT TOP 10 GIDB_G_LOGIN_SESSION_V.ID, GIDB_GC_AGENT.EMPLOYEEID Trigramme,
                            GIDB_GC_AGENT.FIRSTNAME Prenom,
                            GIDB_GC_AGENT.LASTNAME Nom,
                            GIDB_GC_PLACE.NAME Place,
                            GIDB_GC_ENDPOINT.DN Telephone_Associe,
                            GIDB_GC_FOLDER.NAME Localisation,
                            GIDB_G_LOGIN_SESSION_V.CREATED DebutDeSession,
                            GIDB_G_LOGIN_SESSION_V.TERMINATED FinDeSession
                    FROM GEN_GIM.dbo.GIDB_G_LOGIN_SESSION_V GIDB_G_LOGIN_SESSION_V, GEN_GIM.dbo.GIDB_GC_AGENT GIDB_GC_AGENT, GEN_GIM.dbo.GIDB_GC_ENDPOINT GIDB_GC_ENDPOINT, GEN_GIM.dbo.GIDB_GC_FOLDER GIDB_GC_FOLDER, GEN_GIM.dbo.GIDB_GC_PLACE GIDB_GC_PLACE
                    WHERE GIDB_G_LOGIN_SESSION_V.AGENTID = GIDB_GC_AGENT.ID AND GIDB_G_LOGIN_SESSION_V.PLACEID = GIDB_GC_PLACE.ID AND GIDB_G_LOGIN_SESSION_V.PRIMARYDEVICEID = GIDB_GC_ENDPOINT.ID AND GIDB_GC_AGENT.FOLDERID = GIDB_GC_FOLDER.ID AND ((GIDB_GC_ENDPOINT.DN= ?)) ORDER BY GIDB_G_LOGIN_SESSION_V.ID DESC 
            """, num)
            info_gen_num = cursor.fetchall()
            return info_gen_num
        except AttributeError:
            tk.messagebox.showwarning(title="Information", message="Saisie incorrecte !", icon='warning')

    def AlimTabGenNum(self):
        self.tableau_num.delete(*self.tableau_num.get_children())
        num = self.entry_gen_num.get()
        info_gen = self.GenesysSQLNum(num)
        for line in info_gen:
            self.tableau_num.insert("", 'end', text="L", 
            values= (line.Place, line.Trigramme, line.Nom, line.Prenom, line.DebutDeSession, line.FinDeSession))
    
    def get_size(self, bytes, suffix="B"):
        factor = 1024
        for unit in ["", "K", "M", "G", "T", "P"]:
            if bytes < factor:
                return f"{bytes:.2f}{unit}{suffix}"
            bytes /= factor

    def InfoPoste(self, poste):
        try:
            ipaddress = socket.gethostbyname(poste)
            info_computer = wmi.WMI(poste)
            operatingsystem = {}
            for os in info_computer.Win32_OperatingSystem():
                operatingsystem['Caption'] = os.Caption
                operatingsystem['CSName'] = os.CSName
                operatingsystem['LastBootUpTime'] = os.LastBootUpTime
                operatingsystem['InstallDate'] = os.InstallDate
                operatingsystem['OSArchitecture'] = os.OSArchitecture
                operatingsystem['Version'] = os.Version
            computersystem = {}
            for os in info_computer.Win32_ComputerSystem():
                computersystem['BootStatus'] = os.BootStatus
                computersystem['DNSHostName'] = os.DNSHostName
                computersystem['Domain'] = os.Domain
                computersystem['Model'] = os.Model
                computersystem['Manufacturer'] = os.Manufacturer
                computersystem['SystemFamily'] = os.SystemFamily
                computersystem['UserName'] = os.UserName
            logicaldisk_list = {}
            i = 0
            for os in info_computer.Win32_LogicalDisk():
                logicaldisk = {}
                logicaldisk['Caption'] = os.Caption
                logicaldisk['DeviceID'] = os.DeviceID
                logicaldisk['Description'] = os.Description
                logicaldisk['DriveType'] = os.DriveType
                logicaldisk['FileSystem'] = os.FileSystem
                logicaldisk['FreeSpace'] = os.FreeSpace
                logicaldisk['Size'] = os.Size
                logicaldisk['VolumeName'] = os.VolumeName
                logicaldisk['ProviderName'] = os.ProviderName
                logicaldisk_list[i] = logicaldisk
                i += 1
            networkadapter_list = {}
            j = 0
            for os in info_computer.Win32_NetworkAdapter(NetConnectionStatus = 2):
                networkadapter = {}
                networkadapter['InterfaceIndex'] = os.InterfaceIndex
                networkadapter['Description'] = os.Description
                networkadapter['NetConnectionID'] = os.NetConnectionID
                networkadapter['NetConnectionStatus'] = os.NetConnectionStatus
                networkadapter['PhysicalAdapter'] = os.PhysicalAdapter
                networkadapter['MACAddress'] = os.MACAddress
                networkadapter_list[j] = networkadapter
                j += 1

            return ipaddress, operatingsystem, computersystem, logicaldisk_list
        except socket.gaierror as error:
            tk.messagebox.showwarning(title="Information", message="Erreur rencontrée : {}".format(error), icon='warning')

    def Delete_data(self):
        self.txt_name1.delete(1.0, tk.END)
        self.txt_idwin1.delete(1.0, tk.END)
        self.txt_idsig1.delete(1.0, tk.END)
        self.txt_idgrc1.delete(1.0, tk.END)
        self.txt_serv1.delete(1.0, tk.END)
        self.txt_address1.delete(1.0, tk.END)
        self.txt_phone1.delete(1.0, tk.END)
        self.txt_email1.delete(1.0, tk.END)

        self.txt_state.delete(1.0, tk.END)
        self.txt_pwd.delete(1.0, tk.END)

        self.entry_gen_id.delete(0, tk.END)
        self.entry_gen_num.delete(0, tk.END)
        self.entry_win.delete(0, tk.END)

        self.combo.delete(0, tk.END)
        self.combo_poste.delete(0, tk.END)
        self.Delete_data_poste()

        self.tableau_id.delete(*self.tableau_id.get_children())
        self.tableau_num.delete(*self.tableau_num.get_children())

    def Delete_data_poste(self):
        self.txt_poste_ip.delete(1.0, tk.END)
        self.txt_poste_mod.delete(1.0, tk.END)
        self.txt_poste_os.delete(1.0, tk.END)
        self.txt_poste_boot.delete(1.0, tk.END)

    def WriteInfo(self):
        self.Delete_data()
        recup = self.recup_text()
        if recup != False:
            get_info = self.ConnectSQL(recup)
            get_num = self.GenesysSQLID(recup)

            self.txt_name1.insert(1.0, get_info[0].FIRSTNAME.capitalize()+" "+get_info[0].LASTNAME)
            self.txt_idwin1.insert(1.0,get_info[0].TRIGRAMME)
            
            self.txt_idsig1.insert(1.0,get_info[0].LOGIN_SIGMA)
            self.txt_idgrc1.insert(1.0,get_info[0].LOGIN_GRC)
            if len(get_info) > 1:
                computer_tuple = ()
                for comp in get_info:
                    computer_tuple += (comp.ID_RESSOURCE,)
                self.combo['values'] = computer_tuple
                self.combo.current(0)
                self.combo_poste['values'] = computer_tuple
                self.combo_poste.current(0)
            else : 
                self.combo['values'] = get_info[0].ID_RESSOURCE
                self.combo.current(0)
                self.combo_poste['values'] = get_info[0].ID_RESSOURCE
                self.combo_poste.current(0)
            self.txt_serv1.insert(1.0,get_info[0].SERVICE)
            self.txt_address1.insert(1.0,get_info[0].CP+ " "+get_info[0].VILLE)
            if get_num:
                self.txt_phone1.insert(1.0,get_num[0].Place)
                self.entry_gen_id.insert(1,get_info[0].TRIGRAMME)
                self.entry_gen_num.insert(1,get_num[0].Place)
            else:
                self.txt_phone1.insert(1.0,get_info[0].TELEPHONE)
            self.txt_email1.insert(1.0,get_info[0].EMAIL )

            self.entry_win.insert(1, get_info[0].TRIGRAMME)

        else:
            tk.messagebox.showwarning(title="Information", message="Le champs est vide !", icon='warning')

    def WriteInfoPoste(self):
        self.Delete_data_poste()
        get_info_poste = self.InfoPoste(self.recup_text_poste())

        lastboot = get_info_poste[1]['LastBootUpTime'][:4]+"/"+get_info_poste[1]['LastBootUpTime'][4:6]+"/"+get_info_poste[1]['LastBootUpTime'][6:8]+" "+get_info_poste[1]['LastBootUpTime'][8:10]+":"+get_info_poste[1]['LastBootUpTime'][10:12]

        self.txt_poste_ip.insert(1.0,get_info_poste[0])
        self.txt_poste_mod.insert(1.0,get_info_poste[2]['Model'])
        self.txt_poste_os.insert(1.0,get_info_poste[1]['Version'])
        self.txt_poste_boot.insert(1.0,lastboot ) 

    def Open_GrC(self):
        webbrowser.open("URL")

    def Open_Smart(self):
        webbrowser.open("URL")

    def Open_GRC_Acc(self):
        webbrowser.open("URL")

    def Open_Iguazu(self):
        webbrowser.open("URL")
    
    def Open_WifiGuest(self):
        webbrowser.open("URL")

    def Open_Wiki(self):
        webbrowser.open("URL")

    def Open_SPilot(self):
         webbrowser.open("URL")

    def Remote(self):
        args = r"C://Program Files (x86)//Microsoft Configuration Manager//AdminConsole//bin//i386//CmRcViewer.exe"+ " " + self.combo.get()
        subprocess.Popen(args)

    def ADConsol(self):
        os.system('runas /user:"Domain\%username:~0,3%" "cmd /c start mmc C:\AppDsi\EXPGLB\RSAT\AD_Domain.msc /domain=Domain"')

    def Historique(self):
        pass


if __name__ == "__main__":
    toolbox  = Application()
    toolbox.title("ToolBox")
    toolbox.geometry("+0+0")
    toolbox.mainloop()  
