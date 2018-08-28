import json
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox

import docx
from docxtpl import DocxTemplate


class TkDocxTemplaterGui(ttk.Frame):

    def __init__(self, master):
        self.master = master
        super().__init__(self.master)

        self.template_file = tk.StringVar()
        self.json_file = tk.StringVar()
        self.use_manual_json_data = tk.BooleanVar()
        self.output_file = tk.StringVar()

        self.configure_gui()
        self.create_widgets()

    def select_file(self, string_var, filetypes, saveas=False):
        if saveas == True:
            return lambda: string_var.set(
                filedialog.asksaveasfilename(filetypes=filetypes))
        else:
            return lambda: string_var.set(
                filedialog.askopenfilename(filetypes=filetypes))

    def toggle_checkbutton(self):
        if self.use_manual_json_data.get() == True:
            self.textbox.config(state=tk.NORMAL, bg='SystemWindow')
        else:
            self.textbox.config(state=tk.DISABLED, bg='grey90')

    def generate_output(self):
        try:
            doc = DocxTemplate(self.template_file.get())
            if self.use_manual_json_data.get() == True:
                context = json.loads(self.textbox.get('1.0', 'end-1c'))
            else:
                context = json.load(open(self.json_file.get()))
            doc.render(context)
            doc.save(self.output_file.get())
            messagebox.showinfo("Generated output", "Successfully generated output.")
        except (docx.opc.exceptions.PackageNotFoundError, FileNotFoundError):
            messagebox.showerror("Error", "File not found.")
        except json.decoder.JSONDecodeError:
            messagebox.showerror("Error", "Invalid JSON input.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def configure_gui(self):
        self.master.title("Docx Templater")
        self.master.minsize(400,200)
        self.master.columnconfigure(1, weight=1, minsize=80)
        self.master.rowconfigure(3, weight=1)

    def create_widgets(self):
        # Row 0
        ttk.Label(
            self.master,
            text="Template:"
        ).grid(row=0, column=0, sticky=tk.E)

        ttk.Entry(
            self.master,
            textvariable=self.template_file
        ).grid(row=0, column=1, sticky=tk.W+tk.E)

        ttk.Button(
            self.master,
            text="Select file",
            command=self.select_file(
                self.template_file,
                [(".docx files", '*.docx')]
            )
        ).grid(row=0, column=2, sticky=tk.E)

        # Row 1
        ttk.Label(
            self.master,
            text="JSON:"
        ).grid(row=1, column=0, sticky=tk.E)

        ttk.Entry(
            self.master,
            textvariable=self.json_file
        ).grid(row=1, column=1, sticky=tk.W+tk.E)

        ttk.Button(
            self.master,
            text="Select file",
            command=self.select_file(
                self.json_file,
                [("JSON files", '*.json')]
            )
        ).grid(row=1, column=2, sticky=tk.E)

        # Row 2
        ttk.Checkbutton(
            self.master,
            text="Manually enter JSON data:",
            command=self.toggle_checkbutton,
            variable=self.use_manual_json_data,
            onvalue=True,
            offvalue=False
        ).grid(row=2, column=0, columnspan=3, sticky=tk.W)

        # Row 3
        self.textbox = tk.Text(
            self.master,
            width=60,
            height=20,
            state=tk.DISABLED,
            bg='grey90'
        )

        self.textbox.grid(
            row=3,
            column=0,
            columnspan=3,
            sticky=tk.NE+tk.SW,
            padx=5,
            pady=5
        )

        # Row 4
        ttk.Label(
            self.master,
            text="Output:"
        ).grid(row=4, column=0, sticky=tk.E)

        ttk.Entry(
            self.master,
            textvariable=self.output_file
        ).grid(row=4, column=1, sticky=tk.W+tk.E)

        ttk.Button(
            self.master,
            text="Select file",
            command=self.select_file(
                self.output_file,
                [(".docx files", '*.docx')],
                saveas=True
            )
        ).grid(row=4, column=2, sticky=tk.E)

        # Row 5
        ttk.Button(
            self.master,
            text="Generate",
            command=self.generate_output
        ).grid(row=5, column=0, columnspan=3, sticky=tk.E)


if __name__ == '__main__':
    if len(sys.argv) == 4:
        doc = DocxTemplate(sys.argv[1])
        doc.render(json.load(open(sys.argv[2])))
        doc.save(sys.argv[3])
        sys.exit()
    else:        
        root = tk.Tk()
        my_gui = TkDocxTemplaterGui(root)
        root.mainloop()
