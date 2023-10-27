import tkinter as tk
from tkinter import ttk
import webbrowser
from compare_neotech import compare_neotech
from perform_vlookup import perform_vlookup


# Create main window
window = tk.Tk()
window.title("Comparing NeoTech Contract Files")
window.configure(bg="white")
window.geometry("900x500")

# Create a canvas and a vertical scrollbar
canvas = tk.Canvas(window)
scrollbar = ttk.Scrollbar(window, orient="vertical", command=canvas.yview)
canvas.configure(bg="white", yscrollcommand=scrollbar.set)

canvas.configure(yscrollcommand=scrollbar.set)
inner_frame = tk.Frame(canvas, bg="white")
canvas.create_window((window.winfo_width() / 2, 0), window=inner_frame, anchor="n")

inner_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
scrollbar.pack(side="right", fill="y")
canvas.pack(side="left", fill="both", expand=True, padx=20, pady=20)

# Canvas - Scrollbar
canvas.configure(yscrollcommand=scrollbar.set)
scrollbar.configure(command=canvas.yview)


def _on_mousewheel(event):
    canvas.yview_scroll(-1 * (event.delta // 120), "units")


# Bind the function to the MouseWheel event, to make our scrolling function more applicable
canvas.bind_all("<MouseWheel>", _on_mousewheel)

# Widgets
style = ttk.Style()
style.configure("TButton", font=("Rupee", 16, "bold"), width=30, height=2, background="white")
style.map("TButton", foreground=[('active', 'red')], background=[('active', 'blue')])

title_label = ttk.Label(inner_frame, text="Welcome Partnership Member!",
                        font=("Rupee", 32, "underline"), background="white", foreground="#103d81")
title_label.pack(pady=(20, 10), padx=20)  # Adjust the padding values as necessary

description_label = ttk.Label(inner_frame, text="Instructions for NeoTech Contract Files",
                              font=("Rupee", 20, "underline"), background="white")
description_label.pack(pady=(10, 20), padx=20)  # Adjust the padding values as necessary

description_label1 = ttk.Label(inner_frame, text="For the Comparing Neotech Files button select the files in this "
                                                 "order:\n"
                                                 "• NEW NeoTech Contract file for the week\n"
                                                 "• Then Select the Previous NeoTech Contract File",
                               font=("Rupee", 18), background="white", justify='center')
description_label1.pack(pady=(10, 20), padx=20)  # Adjust the padding values as necessary

compare_button = ttk.Button(inner_frame, text='Compare NeoTech Files', command=compare_neotech)
compare_button.pack(pady=10)

description_label2 = ttk.Label(inner_frame, text="For the Perfrom Vlook-Up button select the files in this order:\n"
                                                 "• Contract file with the data we need (could be a previous week "
                                                 "file)\n"
                                                 "• Then Select the Contract File where we have our new 'Lost Items' "
                                                 "sheet!",
                               font=("Rupee", 18), background="white", justify='center')
description_label2.pack(pady=(10, 20), padx=20)  # Adjust the padding values as necessary

vlookup_button = ttk.Button(inner_frame, text='Perform V-lookup', command=perform_vlookup)
vlookup_button.pack(pady=10)


def open_readme_link():
    webbrowser.open('https://github.com/nabilcanan/Partnership_Sort_NeoTech/blob/main/README.md', new=2)


readme_button = ttk.Button(inner_frame, text='Open README', command=open_readme_link)
readme_button.pack(pady=10)

window.mainloop()
