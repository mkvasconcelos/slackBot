import sys
import pandas as pd
import slack
# BIBLIOTECAS INTERFACE GRÁFICA
from tkinter import*
from tkinter import filedialog
################################

# Token SLACK API
client = slack.WebClient(token='TOKEN SLACK')


def Excel():
    global excelfile
    excelfile = filedialog.askopenfilename()
    texto_file["text"] = excelfile


def Code():
    # Leitura de dados
    data = pd.read_excel(excelfile)
    i, j, k = 0, 0, 0
    UserFind = data['User Find']
    email = data['User']
    message1 = data['Message 1']
    message2 = data['Message 2']
    message3 = data['Message 3']
    message, user, emailsNotFound = [], [], []

    for i in range(len(UserFind)):
        if i == 0:
            i += 1
            continue
        elif UserFind[i] == UserFind[i-1]:
            message2[i] = message2[i-1] + message2[i]
            i += 1
        else:
            i += 1
            continue

    for j in range(len(email)):
        try:
            if type(email[j]) == float:
                user.append("SKIP")
            else:
                userEmail = "@" + \
                    str(client.users_lookupByEmail(
                        email=email[j])['user']['id'])
                user.append(userEmail)
        except:
            emailsNotFound.append(email[j])
            user.append("SKIP")
            pass
        j += 1

    for k in range(len(user)):
        if user[k] == "SKIP":
            message.append("SKIP")
            k += 1
            continue
        else:
            message.append(message1[k]+message2[k]+message3[k])
            client.chat_postMessage(
                channel=user[k], text=message[k], username=botname.get(), icon_emoji=boticon.get())
            if emailsNotFound != []:
                var_user.set("Last Message send to: " +
                             email[k] + "\n Problem with some emails, users not found. Look at the file 'usersNotFound.xlsx'!!!")
                pd.DataFrame(data=emailsNotFound).to_excel(
                    'usersNotFound.xlsx')
            else:
                var_user.set("Last Message send to: " + email[k])
            k += 1

    texto_final["text"] = "FINISHED!!!"


def Exit():
    sys.exit()


# INTERFACE GRÁFICA
win = Tk()
win.title("Slack Bot - Sending Messages")
win.configure(background='#f71962')
global var_user
var_user = StringVar()
var_user.set("")

Label(win, bg='#f71962').grid(column=0, row=0, columnspan=2)

texto_orientacao = Label(win, text="Click on the button to send messages in Slack!",
                         bg='#f71962', foreground="white", font=('Helvetica'))
texto_orientacao.grid(column=0, row=1, padx=10, pady=10, columnspan=2)

botao = Button(win, text="CHOOSE FILE", command=Excel,
               font=('Helvetica', 12, 'bold'))
botao.grid(column=0, row=2, columnspan=2)

texto_file = Label(win, text="", bg='#f71962',
                   foreground="white", font=('Helvetica', 10))
texto_file.grid(column=0, row=3, padx=10, pady=10, columnspan=2)

Label(win, text='BOT NAME:', bg='#f71962', foreground="white",
      font=('Helvetica', 12, 'bold')).grid(row=4)
Label(win, text='BOT ICON:', bg='#f71962', foreground="white",
      font=('Helvetica', 12, 'bold')).grid(row=5)

botname = Entry(win)
boticon = Entry(win)

botname.grid(row=4, column=1)
boticon.grid(row=5, column=1)

botao = Button(win, text="SEND MESSAGES", command=Code,
               font=('Helvetica', 12, 'bold'))
botao.grid(column=0, row=6)

botao2 = Button(win, text="EXIT CODE", command=Exit,
                font=('Helvetica', 12, 'bold'))
botao2.grid(column=1, row=6)

texto_user = Label(win, textvariable=var_user, bg='#f71962',
                   foreground="white", font=('Helvetica', 10))
texto_user.grid(column=0, row=7, padx=10, pady=10, columnspan=2)

texto_final = Label(win, text="", bg='#f71962',
                    foreground="white", font=('Helvetica', 18, 'bold'))
texto_final.grid(column=0, row=8, padx=10, pady=10, columnspan=2)

win.mainloop()
