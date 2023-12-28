from tkinter import *
from tkinter import ttk
import random
#from random import randint
from tkinter import filedialog
import sys
#import docx
#import docx2txt
#import rsa
#from tkinter import PhotoImage
#from tkinter import messagebox
#import docx
#from docx import Document
#Cài thư viện để sử dụng
    #pip install tkinter
    #pip install python-docx
    #pip install docx2txt
    #
    #
    #
window = Tk()
window.title('RSA của nhóm 8 ')
window.geometry('850x580')
#window.iconbitmap("C:\Py\L.ico")
window.resizable(False, False)

label_in = Label(window, text="Input", fg='blue', font='arial 20')
label_in.place(x=20, y=10)
label_out = Label(window, text="Output", fg='blue', font='arial 20')
label_out.place(x=400, y=10)
text = Text(window, width=15, height=3)
text.grid(row=3, column=0, columnspan=2)
text.place(x=190, y = 75)

text = Text(window, width=15, height=3)
text.grid(row=3, column=0, columnspan=2)
text.place(x=190, y = 175)


def gcd(a, b):
    while b != 0:
        a, b = b, a % b
    return a


def multiplicative_inverse(e, phi):
    d = 0
    x1 = 0
    x2 = 1
    y1 = 1
    temp_phi = phi
    while e > 0:
        temp1 = temp_phi // e
        temp2 = temp_phi - temp1 * e
        temp_phi = e
        e = temp2

        x = x2 - temp1 * x1
        y = d - temp1 * y1

        x2 = x1
        x1 = x
        d = y1
        y1 = y

    if temp_phi == 1:
        return d + phi


def generate_keypair():
    p = int(entry_input.get())
    q = int(entry_input1.get())
    n = p * q
    phi = (p - 1) * (q - 1)

    e = random.randrange(1, phi)

    g = gcd(e, phi)
    while g != 1:
        e = random.randrange(1, phi)
        g = gcd(e, phi)

    d = multiplicative_inverse(e, phi)

    public = (e, n)
    private = (d, n)
    return public, private


def encrypt():
    #p = int(entry_input1.get())
    #q = int(entry_input2.get())

    public, _ = generate_keypair()
    plaintext = input_ent.get()
    cipher = [(ord(char) ** public[0]) % public[1] for char in plaintext]
    #messagebox.showinfo('Mã hóa', ''.join(map(str, cipher)))
    encoded_text = ''.join(map(str, cipher))
    output_text.delete('1.0', END)
    output_text.insert(END, encoded_text)




def decrypt():
   # p = int(entry_input1.get())
   # q = int(entry_input2.get())
    public, private = generate_keypair()
    plaintext = input_ent.get()
    #cipher = [int(text[i:i+3]) for i in range(0, len(text), 3)]
    cipher = [(ord(char) ** public[0]) % public[1] for char in plaintext]
    plain = [chr((char ** private[0]) % private[1]) for char in cipher]
    #output_label.config(text= ''.join(plain))
    decoded_text = ''.join(plain)
    output_text1.delete('1.0', END)
    output_text1.insert(END, decoded_text)
def documen():
    # Mở file word cần mã hóa
    filepath = filedialog.askopenfilename(title="Select a Word file", filetypes=[("Word files", "*.docx")])
    # Đọc nội dung file word
    text = docx2txt.process(filepath)
    #with open(filepath, 'rb') as f:
    #    data = f.read()
    text.delete('0,1', END)
    text.insert(END, text)



# Hàm hiển thị nội dung file Word
def display_text(filename):
    document = Document(filename)
    text = "\n".join([para.text for para in document.paragraphs])
    text_box.delete("1.0", END)
    text_box.insert(END, text)

global text_box
text_box = Text(window,width= 20, height= 2, bg= 'white')
text_box.place(x= 20, y = 390) 


def encrypt_word():
    p = int(entry_input1.get())
    q = int(entry_input2.get())
    if text_word.get('1.0', END) != "":
        public, _ = generate_keypair(p,q)
        text = Document()
        for line in text_word.get('1.0', END).split('\n'):
            text.add_paragraph(''.join(map(lambda x: chr((ord(x) ** public[0]) % public[1]), line)))
            text.save('encrypted.docx')
            text_word.delete('1.0', END)
            text_word.insert(END, "Đã mã hóa và lưu thành công vào encrypted.docx!")
    else:
        text_word.delete('1.0', END)
        text_word.insert(END, "Bạn chưa chọn file Word để mã hóa!")

def decrypt_word():
    p = int(entry_input1.get())
    q = int(entry_input2.get())
    if text_word.get('1.0', END) != "":
        public, private = generate_keypair(p,q)
        text = Document()
        for line in text_word.get('1.0', END).split('\n'):
            text.add_paragraph(''.join(map(lambda x: chr((ord(x) ** private[0]) % private[1]), line)))
            text.save('decrypted.docx')
            text_word.delete('1.0', END)
            text_word.insert(END, "Đã giải mã và lưu thành công vào decrypted.docx!")
    else:
        text_word.delete('1.0', END)
        text_word.insert(END, "Bạn chưa chọn file Word để giải mã!")

def save_file():
    if text_word.get('1.0', END) != "":
        filename = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
        if filename:
            if not filename.endswith('.docx'):
                filename += '.docx'
            with open(filename, 'wb') as f:
                f.write(text_word.get('1.0', END).encode())
            text_word.delete('1.0', END)
            text_word.insert(END, "Đã lưu file Word đã mã hóa vào " + filename + " !")
    else:
        text_word.delete('1.0', END)
        text_word.insert(END, "Bạn chưa có bất cứ nội dung gì để lưu!")

output_text = Text(window, bg='white', width=29, height=15, fg='black', font='arial 10')
output_text.place(x=400, y=300)

text_word = Text(window, bg='white', width=29, height=15, fg='black', font='arial 10')
text_word.place(x=620, y=300)


btn_encrypt_word = Button(window, text='Mã hóa file Word', command=encrypt_word)
btn_encrypt_word.place(x=20, y=430)

btn_decrypt_word = Button(window, text='Giải mã file Word', command=decrypt_word)
btn_decrypt_word.place(x=20, y=460)

btn_save = Button(window, text='Lưu file Word đã mã hóa', command=save_file)
btn_save.place(x=200, y=430)


entry_input = Entry(window, bg='white', width=15, fg='black', font='arial 10')
entry_input.place(x=20, y=80)
entry_input1 = Entry(window, bg='white', width=15, fg='black', font='arial 10')
entry_input1.place(x=20, y=150)

input_ent = StringVar()

entry_input2 = Entry(window, bg='white', width=40, fg='black', font='arial 10', textvariable=input_ent)
entry_input2.place(x=20, y=290)

btn_en = Button(window, text='Mã hóa', command=encrypt)
btn_en.place(x=20, y=325)
btn_de = Button(window, text='Giải mã', command=decrypt)
btn_de.place(x=80, y=325)

label_nt1 = Label(window, text="Số nguyên tố 1", background= 'yellow')
label_nt1.place(x=20, y=55)
label_nt2 = Label(window, text="Số nguyên tố 2", background= 'yellow')
label_nt2.place(x=20, y=125)
label_n3 = Label(window, text="Nhập thông điệp", background= 'yellow')
label_n3.place(x=20, y=260)
label_n4 = Label(window, text= "Kết quả mã hóa", bg= 'yellow')
label_n4.place(x= 410, y = 50)

public_label = Label(window, text="", bg= 'white')
public_label.place(x=200, y=80)
private_label = Label(window, text="", bg= 'white')
private_label.place(x=200, y=180)


#text cua ban Text nhap tu ban phim
output_text = Text(window, width=25, height=10)
output_text.place(x=400, y=75)
output_text1 = Text(window, width=25, height=10)
#output_text1.grid(row=3, column=0, columnspan=2)
output_text1.place(x=620, y = 75)

# Tạo nút nhập file word
browse_button = Button(window, text="Browse", command= documen)
browse_button.place(x= 200, y = 398)

# Tạo ô để hiển thị dữ liệu đã được mã hóa
text = Text(window,width= 20, height= 2, bg= 'white')
text.place(x= 20, y = 390)
text = Label(window, text= '', bg='white')
text.place(x= 23 , y=395)

def update_public_private_labels():
    public, private = generate_keypair()
    public_label.config(text=str(public))
    private_label.config(text=str(private))

btn_keypair = Button(window, text='Tạo khóa', command=update_public_private_labels)
btn_keypair.place(x=20, y=200)

lbl_w_in = Label(window, text= "Word đã được mã hóa", bg = 'yellow')
lbl_w_in.place(x = 410, y = 275)
lbl_w_ou = Label(window, text= "Word đã được lưu tại", bg = 'yellow')
lbl_w_ou.place(x = 630, y = 275)

label_pub = Label(window, text="Khóa công khai", background= 'yellow')
label_pub.place(x=200, y=50)
label_pri = Label(window, text="Khóa bí mật", bg= 'yellow')
label_pri.place(x=200, y=150)
label_out = Label(window, text="Kết quả giải mã", bg= 'yellow')
label_out.place(x=630, y=50)


window.mainloop()