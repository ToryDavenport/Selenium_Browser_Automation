import tkinter as tk

def run_script():
      arg1 = e1.get()
      print(arg1)
master = tk.Tk()
tk.Label(master, text="First Starting Row").grid(row=0)


e1 = tk.Entry(master)

e1.grid(row=0, column=1)


tk.Button(master, 
          text='Quit', 
          command=master.quit).grid(row=3, 
                                    column=0, 
                                    sticky=tk.W, 
                                    pady=4)
tk.Button(master, text='Submit', command=run_script).grid(row=3, 
                                                               column=1, 
                                                               sticky=tk.W, 
                                                               pady=4)

master.mainloop()

tk.mainloop()
