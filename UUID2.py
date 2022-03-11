# A python 3 script to generate an alpha numeric "UUID"

import random
import string
import tkinter

from tkinter import *

uuid2 = Tk()
uuid2.title("UUID2")
uuid2.resizable(height=False, width=False)

def genUUID2():
  str = ""
  stdout = ""
  str = ''.join(random.SystemRandom().choice(string.ascii_uppercase + string.digits) for _ in range(33))
  stdout = "{}-{}-{}-{}-{}".format(str[0:8], str[8:12], str[12:16], str[16:20], str[21:])
  if braces.get() == 1:
    stdout = "{" + stdout + "}"
  print(stdout)
  uuid2.clipboard_clear()
  uuid2.clipboard_append(stdout)
  uuid2.update()

braces = IntVar(value=1)

braceCheck = Checkbutton(uuid2, text="{Use Braces}", variable=braces, onvalue=1, offvalue=0)
genButton = Button(uuid2, text="Generate", command=genUUID2)

braceCheck.pack(side = RIGHT)
genButton.pack(side = LEFT)

uuid2.mainloop()

# EOF