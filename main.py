from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename
import tkinter as tk
from tkinter import messagebox
from tkinter.filedialog import askdirectory
from tkinter.filedialog import asksaveasfile
from datetime import datetime
import configparser #Configparser is used for reading .ini file
# =============================================================================
# Bulutcontrol and controlPortal2 files are used to compare functions of Parameter File and Energy Label
# =============================================================================
from Bulutcontrol import controlBulut 
from controlPortal2 import controlPortal2

# =============================================================================
# .ini file is used for replaceable data such as Sheet_name or Revision list.
# =============================================================================
# Data file name
Config_filename = "assets/config.ini"
# Resolving Data
config = configparser.ConfigParser()
# Reading Data
config.read([Config_filename],encoding='utf-8')
keys = [
    "SHEET_NAME",
    "port",
    "REV_LIST",
    "HELP_TEXT"
]
for key in keys:
    try:
        print("key: ", key)
        value = config.get("SETTINGS", key)
        #print(f"{key}:", value)
        if key == "REV_LIST":
            Revision_list = value.split(",")
        if key == "SHEET_NAME":
            #print(value)
            SHEET_NAME = value
        if key == "HELP_TEXT":
            HelpText = value
            HelpText = HelpText.replace("*","\n")
    except configparser.NoOptionError:    
        print(f"No option '{key}' in section 'SETTINGS'")



# datetime object containing current date and time
now = datetime.now()
#today = date.today()
current_date = now.strftime("Date: %d/%m/%y\nTime: %H:%M")
print("Current Date =",current_date)

print("Hoşgeldiniz \nlütfen referans dosyasını ve karşılaştırmak üzere .h dosyasını seçiniz.")

#main properties of GUI
root = tk.Tk()
root.geometry("700x400")
root.title('Login')
root.resizable(0, 0)
root.config(bg='#D3D3D3')
root.title('Lotus')
root.iconbitmap('assets/lotus.ico')
#logo (bottom)
logo = PhotoImage(file=r"assets/arcelik_logo_kirmizi.png").subsample(16)
logowrap = Label(root, width=500, height=40, image=logo, background='#D3D3D3').pack(side=BOTTOM, padx=270)

# configure the grid
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=3)

filepath1 = "first global variable"
filepath2 = "second global variable"
message = ["",""]
#Text box to show file paths
path1 = Text(root, height=1, width=65, state="disabled")
path1.place(x=155, y=11)

path2 = Text(root, height=1, width=65, state="disabled")
path2.place(x=155, y=41)

#Buttons (select file and start)
btnFind = ttk.Button(root, width=18, text=" Select .h File ", command=lambda: [clearpath1(), getFilePath_h()])
btnFind.place(x=10, y=10)

btnFind2 = ttk.Button(root, width=18, text="Energy Labelling", command=lambda: [clearpath2(), getFilePath_xl()])
btnFind2.place(x=10, y=40)

btnStart = ttk.Button(root, width=18, text="Start", command=lambda: [startbutton()])
btnStart.place(x=10, y=70)

#btnSave = ttk.Button(root, text = 'Save as .txt', command = lambda : save())
#btnSave.place(x=30, y=350)

button = Button(root, text = 'Save', command = lambda : save2())
button.place(x=30,y=350)

btnExit = ttk.Button(root, text = 'Exit', command=root.destroy)
btnExit.place(x=600, y=350)

v=Scrollbar(root, orient='vertical')
v.pack(side=RIGHT, fill='y', pady=(100,13),padx=8)
#Textbox to print returned string
T = Text(root, height=15, width=80, wrap=WORD, yscrollcommand=v.set)
T.insert(tk.END,"Lotus Beta 0.0.1\nLütfen Energy Label (.xlsx) dosyasını ve karşılaştırmak üzere Parametre (.h) dosyasını seçiniz\nStart'a bastığınızda karşılaştırma başlayacaktır.")
T.config(state= "disabled")
T.place(x=30,y=100)




v.config(command=T.yview)

btnHelp = ttk.Button(root, text = 'Help', command=lambda : createPopup())
btnHelp.place(x=515, y=350)


# =============================================================================
# Yardım popup'ının görüntülenmesi için yazılmış fonksiyon. Dökümantasyon texti
# 
# =============================================================================
def createPopup():
    win = Toplevel(root)
    btnHelp.config(state="disabled")
    win.geometry("700x400")
    win.resizable(0, 0)
    
    v2=Scrollbar(win, orient='vertical')
    v2.pack(side=RIGHT, fill='y', pady=(40,40),padx=(0,10))
    
    HelpBox = Text(win, width=650,wrap=WORD, yscrollcommand=v2.set)
    HelpBox.insert(tk.END,HelpText)
    HelpBox.config(state="disabled")
    HelpBox.pack(pady=40,padx=(15,0))
    v2.config(command=HelpBox.yview)
    
    Label(win, text="Lotus Documentation",wraplength=350).place(x=2, y=3)
    Label(win, text="Utku Akyüz").place(x=2, y=370)
    btnExit2 = ttk.Button(win, text="Exit", command=win.destroy)
    btnExit2.place(x=600,y=370)
    root.wait_window(win)
    btnHelp.config(state="normal")


def getFilePath_xl():
   path2.configure(state="normal")
   filepath = askopenfilename(filetypes=[("Excel files", "*.xlsx")])
   #filepath = filepath.replace("/", "\\\\")
   global filepath1
   filepath1 = filepath
   path2.insert(tk.END, filepath1)
   path2.configure(state="disabled")


def getFilePath_h():
   path1.configure(state="normal")
   filepath = askopenfilename(filetypes=[("Header files", "*.h")])
   #filepath = filepath.replace("/", "\\\\")
   global filepath2
   filepath2 = filepath
   path1.insert(tk.END, filepath2)
   path1.configure(state="disabled")

def get_start(x):
   y = 0
   return '%d.%d' % (x, y)

def get_end(x):
   z = x+1
   y = 0
   return '%d.%d' % (z, y)

def createLabel():
   T.config(state= "normal")
   
   global message
   message = compareFunc(filepath2, filepath1)
   #print("Mesaj listesi : ", message, "Mesaj Listesi Sonu")
   i = 1
   #Bu try exceptten önceki hali (3satır)
   count = 0
   for line in message[0].split('\n'):
      count += 1

   if count == 1:
      T.insert(tk.END, message[0])
      T.config(state="disabled")
      return message

   else:
      for line in message[0].split('\n'):
         if "bulunamadı" in line.split() or "eşleşmedi" in line.split() or "hatalı" in line.split() or "değil" in line.split():
            T.insert(tk.END, line + '\n')
            T.tag_config("start", foreground="red")
            T.tag_add("start", get_start(i), get_end(i))

         elif "uyarı!" in line.split() or "'BLEACH'" in line.split():
             T.insert(tk.END, line + '\n')
             T.tag_config("start3", foreground="orange")
             T.tag_add("start3", get_start(i), get_end(i))   

         elif "Fonksiyonlar" in line.split() or "tespit" in line.split():
             T.insert(tk.END, line + '\n')
             T.tag_config("start4", foreground="black")
             T.tag_add("start4", get_start(i), get_end(i))             

         else:
            T.insert(tk.END, line + '\n')
            T.tag_config("start2", foreground="green")
            T.tag_add("start2", get_start(i), get_end(i))

         i = i+1
      T.config(state= "disabled")
      return message



def cleartext():
   T.config(state= "normal")
   T.delete('1.0', END)
   T.config(state= "disabled")

def clearpath1():
   path1.config(state= "normal")
   path1.delete('1.0', END)
   path1.config(state= "disabled")

def clearpath2():
   path2.config(state= "normal")
   path2.delete('1.0', END)
   path2.config(state= "disabled")

def save():
    print("message in save : ", message , "Enddddddddddddd")
    saveFile = askdirectory(initialdir=filepath1)
    #value=T.get("1.0","end-1c")
    with open(saveFile+'/readme.txt', 'w') as f:
        #f.write(value)
        f.write(current_date+"\n\n"+"Karşılaştırılan Dosyalar :\n"+filepath1+"\n"+filepath2+"\n"+message[1]+message[0])

def save2():
    hfilaname = filepath2.strip().split('/')[-1].split('.')[0]
    Files = [('Text Document', '*.txt')]
    file = asksaveasfile( filetypes = Files, defaultextension = Files, initialfile = "Lotus-"+hfilaname+"-"+ now.strftime("%d%m%y")+now.strftime("-%H%M"))
    a = (str(file).split("'"))[1]
    print("\n\na\n\n", a)
    with open(a, "w") as data:
        data.write(current_date+"\n\n"+"Karşılaştırılan Dosyalar :\n"+filepath1+"\n"+filepath2+"\n"+message[1]+message[0])
    
def startbutton():
   if filepath1=='' or filepath2=='' or filepath2 == 'second global variable' or filepath1 == 'first global variable':
      messagebox.showerror('error', 'Dosya seçiniz')
   else:
      cleartext();
      createLabel();
    
def rotate(l, n):
    return l[n:] + l[:n]


def compareFunc(filepath_h, filepath_xl):
    import pandas as pd  # import pandas
    import time
    
    start = time.time()  # Time starts with the program
    # Read Excel with just only one sheet
    
    try:
        temp = pd.read_excel(filepath_xl, SHEET_NAME)
    except Exception:
        #return ['+//*************************//\n  Excel Dosyası Okunamadı\n'
        #        '  Lütfen Dosyanın kapalı olduğundan olduğundan emin olunuz\n//*************************//']
        return ["# ============================================================================="
        '# Excel Dosyası Okunamadı\n''  Lütfen Dosyanın kapalı olduğundan olduğundan emin olunuz\n'
        '# =============================================================================#']

    #print("Excel dosyasındaki satır sayısı : ", len(temp))
    #print("Excel dosyasındaki sütun sayısı : ", len(temp.columns))
    end = time.time()  # Time ends after excel read
    print("\nExcel dosyasının açılması için geçen süre:", end - start)

    # .h dosyası açılır ve satırları tek tek okunur, 3. satır kart isminin olduğu satırdır.
    file = open(filepath_h)
    content = file.readlines()
    
    #Eğer .h dosyası boş ise, return eden kısım;
    if len(content) == 0:
        return ["Parametre dosyası (.h) Boş Gözükmektedir, lütfen doğru dosyayı seçtiğinizden emin olun."]
    
    string = str(content[2]).rstrip('\n')  # kart ismi elde edilir.

    # extract final word
    res = string.split(' ')
    fin = res[len(res) - 1]

    if fin[len(fin)-1] == 'B' or fin[len(fin)-1] == 'M':
        name = fin[1:len(fin)-1]
    else:
        name = fin[1:len(fin)]

    #EXCEL DOSYASI COLUMN LISTESI, column_lists.index() is used to find a column's index
    column_lists = []
    for i in temp.iloc[:1]:
        column_lists.append(i)
    #print("Col list : ", column_lists.index("Range"))
    
    
    # Range'in sütunu, Range-Türev kartlar cell_obj1 olarak assign edilir.
    rangeCol_index = 0
    for i in temp.iloc[:1]:
        if str(i) == 'Range - Türev Kartlar (Yeni Etiket Geçişi)':
            break
        rangeCol_index += 1

    # Reference(YeniEtiketGeçişi)'nin sütunu, Range-Türev kartlar cell_obj1 olarak assign edilir.
    refCol_index = 0
    for colname in temp.iloc[:1]:
        if str(colname) == 'Ref(YeniEtiketGeçişi)':
            break
        refCol_index += 1
        
    
# =============================================================================
#   Program names are right after "SoftFunc7F" string, so iterate through h file to find the line.
#   strIntro is the introduction part of parameter file. Containing card's general information, it is returned at the end.
# =============================================================================
    strIntro = ""
    word = "SoftFunc7F"
    checkStar = 0
    with open(filepath_h, "r") as file:
        for line_number, line in enumerate(file, start=1):
            if checkStar != 2:
                if "//***" in line:
                    checkStar +=1
                strIntro = strIntro + line #YILDIZLI SATIRDAN İTİBAREN GÖSTERİLMESİ İÇİN
            if word in line:
                # linenmbr is assigned to index of programs' starting line -2.
                linenmbr = line_number
                print("\nHeader dosyası programların listelendiği satır: ", linenmbr+2, "\n")
                break
    
    
    # =============================================================================
    #   FindStr function is used for finding the index of a line, containing the area parameter as string.
    # =============================================================================
    def find_str(area):
        with open(filepath2, "r") as file:
            for stringlinenumber, line in enumerate(file, start=1):
                if area in line:
                    break
            return stringlinenumber
    
    
    
# =============================================================================
# Check linenumber exist in case cart is not found.    
# SON REVİZE, revize kodu ile birlikte kontrol sağlanır.
# =============================================================================
    row_card = 0
    try:
        N = linenmbr + 1
    except:
        print("Kart Ataması Yapılamadı")
        return ["Kart Ataması Yapılamadı"]
    
    
# =============================================================================
#     
# =============================================================================
    notfoundrefprog = 0
    for i in range(0, len(temp)):  # iterate through excel and display data
        cell_obj2 = str(temp.iloc[i, rangeCol_index])
        if name in str(cell_obj2):
            cell_ref = str(temp.iloc[i, refCol_index])
            ref_code = cell_ref.rstrip('\n').split('_')
            print("ref_code : ", ref_code)
            if len(ref_code) != 1:
                print("ref_code[1] :", ref_code[1])
                if ref_code[1] in Revision_list:
                    updt_ref_code = ref_code[0]+'_'+ref_code[1]
                    while str(content[N]) != "\n":
                        if updt_ref_code in str(content[N]):
                            row_card = i
                            notfoundrefprog = 0
                            break
                        else:
                            notfoundrefprog = 1  #referans programı bulunamazsa 1 olacak
                        N = N + 1
                else:
                    N = linenmbr+1
                    updt_ref_code = ref_code[0]
                    while str(content[N]) != "\n":
                        if updt_ref_code in str(content[N]):
                            notfoundrefprog = 0
                            row_card = i
                            break
                        else:
                            notfoundrefprog = 1
                        N = N+1
            else:
                N = linenmbr+1
                updt_ref_code = ref_code[0]
                while str(content[N]) != "\n":
                    if updt_ref_code in str(content[N]):
                        row_card = i
                        notfoundrefprog = 0
                        break
                    else:
                       notfoundrefprog = 1
                    N = N+1
        else:
            continue
        break



# =============================================================================
# 
# =============================================================================
    warning_notfoundrefprog = ''
    if notfoundrefprog == 1:
        warning_notfoundrefprog = 'Referans programı eşleşmedi\n'
        for i in range(0, len(temp)):  # iterate through excel and display data
            # cell_obj2 = sh.cell(row=i, column=j)
            cell_obj2 = str(temp.iloc[i, rangeCol_index])
            # print("cell_obj2 : ", cell_obj2)
            if name in str(cell_obj2):
                row_card = i




    print("Excel dosyasında kart'ın bulunduğu satır numarası: ", row_card)
    if row_card == 0 and notfoundrefprog == 1:
        return ['Kart bulunamadı']
    # SON REVİZE

    # exceptions
    col_index = -1
    for i in temp.iloc[:1]:
        col_index += 1
        if str(i) == 'P0':
            p0 = col_index
        if str(i) == 'P0 - Matbuat':
            p10 = col_index
        if str(i) == 'Download Cycle - 1':
            dc1 = col_index
        if str(i) == 'RL2_10A':
            dc7 = col_index
            break

# =============================================================================
#     # count the number of programs, if it includes Download Cycle, checks for dc1 to dc7 in excel file
#     # to determine download programs, else, it checks other programs and increase programcounter
# =============================================================================
    programCounter2 = 0
    codeof_programsExcel = []  # Excel program codes will be appended to
    for m in range(p0, p10):
        #print(temp.iloc[row_card,m])
        
        if str(temp.iloc[row_card, m]) != '-' and str(temp.iloc[row_card, m]) != 'Call Service' and str(temp.iloc[row_card, m]) != 'Settings':
            if str(temp.iloc[row_card, m]) == 'Download Cycle':
                #print("\nDownload Cycle Programları aşağıda listelenir \n")
                for a in range(dc1, dc7):
                    # check if download program is not empty
                    if str(temp.iloc[row_card, a]) != '-':
                        #print(temp.iloc[row_card, a])
                        programCounter2 = programCounter2 + 1  # if it is not empty, increase counter
            else:
                programCounter2 = programCounter2 + 1

# =============================================================================
#      Determine the programs and program count in .h file, append program codes to H file
#      Counts until a space character
# =============================================================================
    N = linenmbr + 1
    programsHfile = []
    programCounter = 0  # counter of program number for .h file.
    while str(content[N]) != "\n": 
        programCounter = programCounter + 1
        programsHfile.append(content[N])
        #print("Content N : ", content[N])
        N = N + 1

    codeof_programsHfile = []
    for i in programsHfile:
        a = i.split(" ")[1]
        #print("i :", a)
        a = a.split("_")
        updt_a = a[0]+"_"+a[1]+"_"+a[2]
        codeof_programsHfile.append(updt_a)

    print("\nExcel dosyasında bulunan program sayısı : ", programCounter2)
    print("Header dosyasında bulunan program sayısı : ", programCounter)
    
    
# =============================================================================
#   # Parameters created for final iteration
#   # Control array shows the unmatched programs with 0
#   # ProgramControl array holds column index of programs exist in excel
#   # STRfinal is the final string to be modified
# =============================================================================
    control = [0] * programCounter
    programControl = [0] * programCounter 
    strtextbox = [{}] * programCounter
    strfinal = "Karşılaştırma Tamamlandı\n\n" 
    p = 0
    programCol_index = 0


# =============================================================================
#     # /****COMPARE PROGRAMS OF BOTH FILES AND DETECT MISSING PROGRAMS/TYPOS IN PARAMETER FILE ******/#
#   Reset N to index of programs' starting line + 1
#   Create final string to be attached to text widget in GUI
#   Iterating through p0 to p10 if program counters of .h and excel files are equal
#   Iterate by programcodes through P column
# =============================================================================
    N = linenmbr + 1 
    if programCounter == programCounter2:
        for programCode in range(p0, p10):
            if str(temp.iloc[row_card, programCode]) != '-' and str(
                    temp.iloc[row_card, programCode]) != 'Settings' and str(
                    temp.iloc[row_card, programCode]) != 'Call Service':
                if str(temp.iloc[row_card, programCode]) == 'Download Cycle':
                    for programCode2 in range(dc1, dc7):
                        if str(temp.iloc[row_card, programCode2]) != '-' and str(
                                temp.iloc[row_card, programCode2]) != 'Settings' and str(
                                temp.iloc[row_card, programCode2]) != 'Call Service':
                            codecell = str(temp.iloc[row_card, programCode2]).rstrip('\n')
                            #print(" CODEcell : ", codecell)
                            codeof_programsExcel.append(codecell)
                            code = codecell.split('-')
                            codenew = code[1].split('_')
                            programControl[programCol_index] = programCode2
                            programCol_index += 1
                            #print("CODENEW : ", codenew)
                            if len(codenew) != 1:
                                if codenew[1] in Revision_list:
                                    updt_code = codenew[0] + '_' + codenew[1]
                                else:
                                    updt_code = codenew[0]    
                                #codeof_programsExcel.append(updt_code)
                            else:
                                updt_code = codenew[0]
                                #codeof_programsExcel.append(updt_code)
                            while str(content[N]) != "\n":
                                if updt_code in str(content[N]):
                                    res2 = str(content[N]).rstrip('\n').split(' ')
                                    word2 = res2[1]
                                    word2 = word2.replace('_Size', "")
                                    control[p] = 1
                                    strtextbox[p] = updt_code + \
                                                    '  -  ' + word2 + '\n'
                                N = N + 1
                            p += 1
                            N = linenmbr + 1
                else:
                    codecell = str(temp.iloc[row_card, programCode]).rstrip("\n")
                    #print(" CODEcell : ", codecell)
                    codeof_programsExcel.append(codecell)
                    code = codecell.split('-')
                    codenew = code[1].split('_')
                    programControl[programCol_index] = programCode
                    programCol_index += 1
                    if len(codenew) != 1:
                        if codenew[1] in Revision_list:
                            updt_code = codenew[0] + '_' + codenew[1]
                        else:
                            updt_code = codenew[0]
                        #codeof_programsExcel.append(updt_code)
                    else:
                        updt_code = codenew[0]
                        #codeof_programsExcel.append(updt_code)

                    while str(content[N]) != "\n":
                        if updt_code in str(content[N]):
                            res2 = str(content[N]).rstrip('\n').split(' ')
                            word2 = res2[1]
                            word2 = word2.replace('_Size', "")
                            control[p] = 1
                            strtextbox[p] = updt_code + '  -  ' + word2 + '\n'

                        N = N + 1
                    p += 1
                N = linenmbr + 1

        for s in range(0, programCounter):
            if control[s] == 0:
                notfoundcodecell = str(temp.iloc[row_card, programControl[s]]).rstrip('\n')
                notfoundcode = str(notfoundcodecell).split('-')[1].split('_')
                if len(notfoundcode) != 1:
                    if notfoundcode[1] in Revision_list:
                        updt_nf_code = notfoundcode[0] + '_' + notfoundcode[1]
                    else:
                        updt_nf_code = notfoundcode[0]
                else:
                    updt_nf_code = notfoundcode[0]
                strtextbox[s] = str(updt_nf_code) + \
                                ' parametre dosyasında bulunamadı' + '\n'

        if str(temp.iloc[row_card, p0]).rstrip("\n") != '-':
            strtextbox = rotate(strtextbox, 1)

        for s in range(0, programCounter):
            strfinal = strfinal + str(strtextbox[s])
    else:
        return ['Eksik/Fazla Program Seçimi']
    
    
    
    print("codeof_programsExcel :" , codeof_programsExcel)

    # =============================================================================
    # Programexcel ve Hfile dosyalarının sıralama açısından karşılaştırlıması, programpronounlarının tutulduğu listeden
    # Auto,Eco,MT programlarının indexinin bulunması ve bit no için kullanılması.
    # Exceldeki programlar ters çevrilir, sıralama hatası olduğu taktirde uyarı verilir.
    # =============================================================================
    if str(temp.iloc[row_card, p0]).rstrip("\n") != '-':
        codeof_programsExcel = rotate(codeof_programsExcel,1)

    ProgramPronouns = []
    for i in range(len(codeof_programsExcel)):
        ProgramPronouns.append(codeof_programsExcel[i].split()[0].split("-")[0]) 
        try:
            if codeof_programsExcel[i].split()[2] != codeof_programsHfile[i].split("_")[0] and codeof_programsHfile[i].split("_")[0] not in codeof_programsExcel[i].split()[2] != codeof_programsHfile[i].split("_")[0]  :
                print("Programlar için iki dosyada farklı Sıralama mevcut")
        except Exception:
            pass

    # =============================================================================
    #     # Karşılaştırma sırasında Bit'no kısmında kullanılmak üzere Auto,Eco,MT programlarının indexinin bulunması.
    # =============================================================================
    if "Auto" in ProgramPronouns:
        AutoProgIndex = str(ProgramPronouns.index("Auto"))
    else:
        AutoProgIndex = "255"
    if "Eco50" in ProgramPronouns :
        EcoProgIndex = str(ProgramPronouns.index("Eco50"))
    else:
        EcoProgIndex = "255"
    if "MT" in ProgramPronouns:
        MtProgIndex = str(ProgramPronouns.index("MT"))
    else:
        MtProgIndex = "255"
    if "Prewash" in ProgramPronouns:
        PWProgIndex = str(ProgramPronouns.index("Prewash"))
    else:
        PWProgIndex = "255"
    if "Hygiene" in ProgramPronouns:
        HygProgIndex = str(ProgramPronouns.index("Hygiene"))
    else:
        HygProgIndex = "255"
        
    # =============================================================================
    # Güç kartına göre anakart tipinin tespit edilmesi ve uygun fonksiyon karşılaştırma işleminin gerçekleştirilmesi
    # pd.isna pandas kütüphanesinin "nan" değerlerini tespit eden fonksiyonudur. Güç kartı sütununda nan değerler mevcut olduğu için kontrolü sağlanmıştır
    # strfinal program karşılaştırmasının dönütünü içerir. 
    # strIntro parametre dosyasının giriş kısmını içerir, return edilen objenin 2. elemanıdır. [1. eleman, 2.eleman]
    # warning_notfoundrefprog eğer referans kodu bulunamazsa ortaya çıkar, nadir bir istisnadır, program karşılaştırılması yine de yapılır
    # Programın çalışma süresi time2 - start şeklinde bulunur. time2 comparefunc'ın en son returnünden hemen önce çağırılır.
    # =============================================================================
    if pd.isna(temp.iloc[row_card,column_lists.index('Güç Kartı')]) == False:
        if "PORTAL-2" in temp.iloc[row_card,column_lists.index('Güç Kartı')]:
            FunctionString = controlPortal2(filepath2,content,AutoProgIndex,EcoProgIndex,MtProgIndex,PWProgIndex,HygProgIndex,temp,row_card,column_lists,name)
            pass   
        else:
            FunctionString = controlBulut(filepath2,content,AutoProgIndex,EcoProgIndex,MtProgIndex,PWProgIndex,HygProgIndex,temp,row_card,column_lists,name)
    else:
        return ["Seçilen kart Portal-2 veya YeniMeşe/Bulut kartı değil."]
    
    time2 = time.time()
    return [warning_notfoundrefprog + strfinal + "\n" + FunctionString + "\nProgramın çalışması: " + str(round(time2 - start,2)) + " Saniye sürmüştür", strIntro]
root.mainloop()