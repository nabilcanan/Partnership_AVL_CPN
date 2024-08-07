import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
import os
from concat_and_formulas import process_workbook


def gui_create(main_window):
    style = ttk.Style()
    main_window.configure(bg="white")
    # Adjusting button size and font size
    style.configure("TButton", font=("Roboto", 12, "bold"), width=30, height=2)
    style.map("TButton", foreground=[('active', 'white')], background=[('active', '#007BFF')])

    title_label = ttk.Label(main_window, text="Welcome Partnership Team!",
                            font=("Segoe UI", 24, "underline"), background="white", foreground="#103d81")
    title_label.pack(pady=10)  # Reduced padding

    description_label = ttk.Label(main_window,
                                  text="This tool allows you to take the desired information from the BOM\n"
                                       "and concatenate desired column as well as perform other logic",
                                  font=("Roboto", 14), background="white", anchor="center",
                                  justify="center")
    description_label.pack(pady=10)

    instructions_button = ttk.Button(main_window, text="Open Instructions", command=open_instructions)
    instructions_button.pack(pady=10)  # Added padding for the button

    process_button = ttk.Button(main_window, text="Process Workbook", command=show_warning_and_process)
    process_button.pack(pady=20)  # Added padding for the button


def open_instructions():
    # Path to the instructions PowerPoint file
    instructions_path = r"P:\Partnership_Python_Projects\ADV CPN Program\Partnership_ALV_CPN.pptx"
    # Open the file with the default application
    os.startfile(instructions_path)


def show_warning_and_process():
    # Show warning message
    messagebox.showwarning("Attention",
                           "Did you rename the MFG Name column in your workbook to 'CLEAN MANUFACTURER NAME'? \n"
                           "Otherwise, the program won't work as expected. \n"
                           "Please make sure you change the name of this column and ensure you have the CPN column as well.\n")

    # Proceed to file selection and processing
    select_and_process_workbook()


def select_and_process_workbook():
    # Open file dialog to select the Excel file to load
    file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])

    # Open file dialog to specify where to save the modified Excel file
    save_path = filedialog.asksaveasfilename(title="Save modified Excel file", defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")])

    # Process the workbook
    if file_path and save_path:
        process_workbook(file_path, save_path)
        messagebox.showinfo("Process Complete", f"The process is complete. The file is saved at:\n{save_path}")


if __name__ == "__main__":
    root = tk.Tk()
    root.title('AVL CPN Program')
    root.geometry('600x300')
    gui_create(root)
    root.mainloop()
