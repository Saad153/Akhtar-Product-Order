try:
    import tkinter
    import customtkinter
    from tkinter import messagebox, filedialog
    from automation import process_excel_files  # Import the refactored function
    import requests

    # Keep the same theme and appearance
    customtkinter.set_appearance_mode("light")
    customtkinter.set_default_color_theme('blue')

    # --- Start Processing Function ---
    def start_processing():
        po_file = entry_po.get()
        order_form_file = entry_order.get()

        if po_file == '' or order_form_file == '':
            messagebox.showerror('Error', 'Please select both PO and Order Form files')
            return

        try:
            output, labels_and_quantities = process_excel_files(po_file, order_form_file)


            messagebox.showinfo("Success", f"File processed successfully and saved at:\n{output}")
        except Exception as e:
            messagebox.showerror("Processing Error", f"An error occurred:\n{str(e)}")

    # --- Browse File Function ---
    def browse_file(entry_widget):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        entry_widget.delete(0, 'end')
        entry_widget.insert(0, file_path)

    # --- GUI Layout ---
    root = customtkinter.CTk()
    root.title('Excel Label Processor')

    label = customtkinter.CTkLabel(master=root,
                                   text="Excel Label Processor",
                                   font=('Helvetica', 20))
    label.place(relx=0.5, rely=0.1, anchor=tkinter.N)

    # Entry for PO File
    entry_po = customtkinter.CTkEntry(master=root,
                                      width=300,
                                      height=25,
                                      placeholder_text="Select PO File")
    entry_po.place(relx=0.38, rely=0.2, anchor=tkinter.N)

    browse_po_button = customtkinter.CTkButton(master=root,
                                               text="Browse",
                                               command=lambda: browse_file(entry_po),
                                               width=120,
                                               height=25,
                                               border_width=0,
                                               corner_radius=8)
    browse_po_button.place(relx=0.82, rely=0.2, anchor=tkinter.N)

    # Entry for Order Form File
    entry_order = customtkinter.CTkEntry(master=root,
                                         width=300,
                                         height=25,
                                         placeholder_text="Select Order Form File")
    entry_order.place(relx=0.38, rely=0.25, anchor=tkinter.N)

    browse_order_button = customtkinter.CTkButton(master=root,
                                                  text="Browse",
                                                  command=lambda: browse_file(entry_order),
                                                  width=120,
                                                  height=25,
                                                  border_width=0,
                                                  corner_radius=8)
    browse_order_button.place(relx=0.82, rely=0.25, anchor=tkinter.N)

    # Execute Button
    execute_button = customtkinter.CTkButton(master=root,
                                             text="Process Files",
                                             command=start_processing,
                                             width=430,
                                             height=25,
                                             border_width=0,
                                             corner_radius=8)
    execute_button.place(relx=0.51, rely=0.33, anchor=tkinter.N)

    root.geometry("500x600")

    # --- Connection Check (unchanged) ---
    try:
        cond_AT = requests.get("https://saim2481.pythonanywhere.com/ATactivation-desktop-response/")
        cond_AT.raise_for_status()
        cond_AT = cond_AT.text
    except requests.exceptions.RequestException:
        cond_AT = False
        messagebox.showerror("Connection Error", "Please Check your internet Connection")
    except:
        cond_AT = False
        messagebox.showerror("Something Went Wrong", "Unexpected Error")

    print(cond_AT)
    if cond_AT != "true":
        execute_button.configure(state=customtkinter.DISABLED)

    root.mainloop()
except Exception as e:
    print(e)
    while True:
        pass
