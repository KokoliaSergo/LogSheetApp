import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
import os

# PDF Generation imports
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

# Global lists to store file paths and cached results
file_paths_cache = []
last_analysis_results = []

def select_files():
    global file_paths_cache
    filepaths = filedialog.askopenfilenames(
        title="Select Log Sheets",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if filepaths:
        for f in filepaths:
            if f not in file_paths_cache:
                file_paths_cache.append(f)
                file_listbox.insert(tk.END, os.path.basename(f))

def remove_selected_file():
    global file_paths_cache
    selected_indices = file_listbox.curselection()
    for index in reversed(selected_indices):
        file_listbox.delete(index)
        file_paths_cache.pop(index)

def run_analysis():
    global file_paths_cache, last_analysis_results
    if not file_paths_cache:
        messagebox.showwarning("Selection Error", "Please select at least one log sheet using 'Select Log Sheets...' first.")
        return
        
    target_name = name_entry.get().strip().lower()
    if not target_name:
        messagebox.showwarning("Input Error", "Please enter a name and surname to search.")
        return
    
    # Clear previous search results from the frame (skip row 0 headers)
    for widget in table_inner_frame.winfo_children():
        info = widget.grid_info()
        if info.get('row', 0) != 0:
            widget.destroy()
            
    all_results = []
        
    for file_path in file_paths_cache:
        filename = os.path.basename(file_path)
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            
            log_sheet_num = "N/A"
            subject = "N/A"
            
            # Updated mapping for the new log sheet template structure:
            # Log Sheet Number is at C7, Subject is at C12 on the 'Verification' sheet
            if "Verification" in wb.sheetnames:
                ver_sheet = wb["Verification"]
                val_c7 = ver_sheet["C7"].value
                val_c12 = ver_sheet["C12"].value
                
                if val_c7 is not None: 
                    log_sheet_num = str(val_c7).strip()
                if val_c12 is not None: 
                    subject = str(val_c12).strip()
            
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(max_col=13):
                    for cell in row:
                        if cell.value and isinstance(cell.value, str):
                            if target_name in cell.value.lower():
                                
                                role = "Unknown Role"
                                for r_idx in range(cell.row, max(0, cell.row - 4), -1):
                                    role_val = sheet.cell(row=r_idx, column=2).value
                                    if role_val:
                                        role = str(role_val).strip()
                                        break
                                
                                all_results.append((log_sheet_num, subject, sheet.title, role))
                                
        except Exception as e:
            all_results.append(("Error", f"Could not read {filename}", "N/A", "N/A"))
            
    last_analysis_results = all_results
    render_table(all_results)

def export_to_pdf():
    global last_analysis_results
    if not last_analysis_results:
        messagebox.showwarning("Export Warning", "No analysis results available to export. Please run an analysis first.")
        return

    target_name = name_entry.get().strip()
    
    pdf_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        title="Export Analysis Report as PDF",
        initialfile="LogSheet_Analysis_Report.pdf"
    )
    
    if not pdf_path:
        return

    try:
        doc = SimpleDocTemplate(pdf_path, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        elements = []
        
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'ReportTitle',
            parent=styles['Heading1'],
            fontName='Helvetica-Bold',
            fontSize=16,
            textColor=colors.HexColor('#0277bd'),
            spaceAfter=6
        )
        subtitle_style = ParagraphStyle(
            'ReportSubtitle',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=10,
            textColor=colors.HexColor('#546e7a'),
            spaceAfter=15
        )
        
        elements.append(Paragraph("Aeronautical Information Publication - Role Analysis Report", title_style))
        elements.append(Paragraph(f"<b>Searched Personnel:</b> {target_name.title()}", subtitle_style))
        elements.append(Spacer(1, 5))
        
        sorted_results = sorted(last_analysis_results, key=lambda x: (str(x[0]), str(x[1])))
        
        log_groups = []
        for row in sorted_results:
            if not log_groups:
                log_groups.append([row])
            elif log_groups[-1][0][0] == row[0]:
                log_groups[-1].append(row)
            else:
                log_groups.append([row])
                
        table_data = [["Log Sheet #", "Data Change Request Subject", "Tab", "Role"]]
        pdf_table_style = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#eceff1')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#37474f')),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cfd8dc'))
        ]
        
        current_row_idx = 1
        
        for log_group in log_groups:
            log_group_size = len(log_group)
            log_start_row = current_row_idx
            
            sub_groups = []
            for row in log_group:
                if not sub_groups:
                    sub_groups.append([row])
                elif sub_groups[-1][0][1] == row[1]:
                    sub_groups[-1].append(row)
                else:
                    sub_groups.append([row])
                    
            sub_row_tracker = current_row_idx
            for sub_group in sub_groups:
                sub_group_size = len(sub_group)
                sub_start_row = sub_row_tracker
                
                for inner_i, (_, _, tab, role) in enumerate(sub_group):
                    actual_row = sub_row_tracker + inner_i
                    row_bg = colors.white if (actual_row % 2 != 0) else colors.HexColor('#f4f9fd')
                    pdf_table_style.append(('BACKGROUND', (2, actual_row), (3, actual_row), row_bg))
                    
                    table_data.append([
                        "", 
                        "", 
                        Paragraph(str(tab), styles['Normal']),
                        Paragraph(str(role), ParagraphStyle('RoleStyle', parent=styles['Normal'], fontName='Helvetica-Bold', textColor=colors.HexColor('#0288d1')))
                    ])
                    
                if sub_group_size > 1:
                    pdf_table_style.append(('SPAN', (1, sub_start_row), (1, sub_start_row + sub_group_size - 1)))
                
                sub_text = sub_group[0][1]
                table_data[sub_start_row][1] = Paragraph(str(sub_text), styles['Normal'])
                
                sub_row_tracker += sub_group_size
                
            if log_group_size > 1:
                pdf_table_style.append(('SPAN', (0, log_start_row), (0, log_start_row + log_group_size - 1)))
                
            log_num_text = log_group[0][0]
            table_data[log_start_row][0] = Paragraph(str(log_num_text), ParagraphStyle('LogNumStyle', parent=styles['Normal'], fontName='Helvetica-Bold', textColor=colors.HexColor('#0277bd')))
            
            current_row_idx += log_group_size
            
        col_widths = [80, 245, 80, 130]
        pdf_table = Table(table_data, colWidths=col_widths, repeatRows=1)
        pdf_table.setStyle(TableStyle(pdf_table_style))
        
        elements.append(pdf_table)
        doc.build(elements)
        
        messagebox.showinfo("Success", f"Report successfully exported to:\n{pdf_path}")
        
    except Exception as e:
        messagebox.showerror("Export Error", f"An error occurred while generating the PDF:\n{str(e)}")

def render_table(all_results):
    if not all_results:
        table_inner_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        return

    all_results.sort(key=lambda x: (str(x[0]), str(x[1])))
    
    log_groups = []
    for row in all_results:
        if not log_groups:
            log_groups.append([row])
        elif log_groups[-1][0][0] == row[0]:
            log_groups[-1].append(row)
        else:
            log_groups.append([row])
            
    row_idx = 1 
    
    for log_group in log_groups:
        log_num = log_group[0][0]
        log_group_size = len(log_group)
        
        lbl_num = tk.Label(table_inner_frame, text=log_num, bg="#ffffff", fg="#0277bd", font=("Arial", 9, "bold"), relief="flat", padx=5, pady=5)
        lbl_num.grid(row=row_idx, column=0, rowspan=log_group_size, sticky="nsew", padx=1, pady=1)
        
        sub_groups = []
        for row in log_group:
            if not sub_groups:
                sub_groups.append([row])
            elif sub_groups[-1][0][1] == row[1]:
                sub_groups[-1].append(row)
            else:
                sub_groups.append([row])
                
        sub_row_tracker = row_idx
        for sub_group in sub_groups:
            subject = sub_group[0][1]
            sub_group_size = len(sub_group)
            
            bg_color = "#f4f9fd" if (sub_row_tracker // 2) % 2 == 0 else "#ffffff"
            lbl_sub = tk.Label(table_inner_frame, text=subject, bg=bg_color, fg="#263238", relief="flat", padx=5, pady=5, justify="left", anchor="w")
            lbl_sub.grid(row=sub_row_tracker, column=1, rowspan=sub_group_size, sticky="nsew", padx=1, pady=1)
            lbl_sub.bind("<Configure>", make_label_wrap_dynamically)
            
            for inner_i, (_, _, tab, role) in enumerate(sub_group):
                curr_row = sub_row_tracker + inner_i
                row_bg = "#f4f9fd" if curr_row % 2 == 0 else "#ffffff"
                
                lbl_tab = tk.Label(table_inner_frame, text=tab, bg=row_bg, fg="#546e7a", relief="flat", padx=5, pady=5)
                lbl_tab.grid(row=curr_row, column=2, sticky="nsew", padx=1, pady=1)
                
                lbl_role = tk.Label(table_inner_frame, text=role, bg=row_bg, fg="#0288d1", font=("Arial", 9, "bold"), relief="flat", padx=5, pady=5, justify="left", anchor="w")
                lbl_role.grid(row=curr_row, column=3, sticky="nsew", padx=1, pady=1)
                lbl_role.bind("<Configure>", make_label_wrap_dynamically)
                
            sub_row_tracker += sub_group_size
            
        row_idx += log_group_size
            
    table_inner_frame.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))

def make_label_wrap_dynamically(event):
    label = event.widget
    new_wraplength = event.width - 12
    if new_wraplength > 0 and label.cget("wraplength") != new_wraplength:
    
        label.configure(wraplength=new_wraplength)

def create_rounded_button(parent, text, bg_color, hover_color, fg_color, command, width=145):
    canvas = tk.Canvas(parent, bg="#f5f7fa", highlightthickness=0, width=width, height=36, cursor="hand2")
    
    def draw_button(color):
        canvas.delete("all")
        canvas.create_arc(0, 0, 16, 16, start=90, extent=90, fill=color, outline=color)
        canvas.create_arc(width-16, 0, width, 16, start=0, extent=90, fill=color, outline=color)
        canvas.create_arc(0, 20, 16, 36, start=180, extent=90, fill=color, outline=color)
        canvas.create_arc(width-16, 20, width, 36, start=270, extent=90, fill=color, outline=color)
        canvas.create_rectangle(8, 0, width-8, 36, fill=color, outline=color)
        canvas.create_rectangle(0, 8, width, 28, fill=color, outline=color)
        canvas.create_text(width // 2, 18, text=text, fill=fg_color, font=("Arial", 9, "bold"))

    draw_button(bg_color)
    canvas.bind("<Enter>", lambda e: draw_button(hover_color))
    canvas.bind("<Leave>", lambda e: draw_button(bg_color))
    canvas.bind("<Button-1>", lambda e: command())
    
    return canvas

# --- Graphical User Interface Setup ---
app = tk.Tk()
app.title("LogSheetApp - Role Analyzer")
app.geometry("1000x800")
app.configure(bg="#f5f7fa")

top_container = tk.Frame(app, bg="#f5f7fa")
top_container.pack(fill="x", padx=10, pady=(15, 10))

controls_frame = tk.Frame(top_container, bg="#f5f7fa")
controls_frame.pack(side="left", fill="x")

name_label = tk.Label(controls_frame, text="Insert Name and Surname:", font=("Arial", 10, "bold"), bg="#f5f7fa", fg="#37474f")
name_label.pack(side="left", padx=(0, 8))

name_entry = tk.Entry(controls_frame, font=("Arial", 10), width=25, bg="white", fg="#263238", insertbackground="#263238", relief="solid", bd=1)
name_entry.pack(side="left", padx=(0, 15))

select_btn_canvas = create_rounded_button(controls_frame, "Select Log Sheets...", "#0288d1", "#0277bd", "white", select_files, width=150)
select_btn_canvas.pack(side="left", padx=(0, 10))

analyse_btn_canvas = create_rounded_button(controls_frame, "Analyse", "#2e7d32", "#1b5e20", "white", run_analysis, width=110)
analyse_btn_canvas.pack(side="left")

credit_label = tk.Label(top_container, text="Designed and Developed by Sergo Kokolia", font=("Arial", 9, "italic"), bg="#f5f7fa", fg="#78909c")
credit_label.pack(side="right", padx=5)

files_frame = tk.LabelFrame(app, text=" Selected Log Sheets ", font=("Arial", 9, "bold"), bg="#f5f7fa", fg="#37474f", padx=10, pady=5)
files_frame.pack(fill="x", padx=10, pady=(0, 15))

list_container = tk.Frame(files_frame, bg="#f5f7fa")
list_container.pack(fill="x", expand=True)

file_listbox = tk.Listbox(list_container, height=8, font=("Arial", 9), selectmode=tk.EXTENDED, bg="white", fg="#263238", selectbackground="#b3e5fc", selectforeground="#01579b", relief="solid", bd=1)
file_scrollbar = tk.Scrollbar(list_container, orient="vertical", command=file_listbox.yview)
file_listbox.configure(yscrollcommand=file_scrollbar.set)

file_listbox.pack(side="left", fill="x", expand=True)
file_scrollbar.pack(side="right", fill="y")

remove_btn_canvas = create_rounded_button(files_frame, "Remove Selected", "#e53935", "#ef5350", "white", remove_selected_file, width=130)
remove_btn_canvas.pack(anchor="e", pady=(5, 0))

# Frozen Header Results Table Setup
table_outer_container = tk.Frame(app, bg="#f5f7fa")
table_outer_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))

header_frame = tk.Frame(table_outer_container, bg="#cfd8dc")
header_frame.pack(fill="x", padx=1, pady=(1, 0))

headers = ["Log sheet number", "Data change request subject", "Tab", "Role"]
col_widths = [16, 45, 16, 26]

for col, (text, w) in enumerate(zip(headers, col_widths)):
    lbl = tk.Label(header_frame, text=text, font=("Arial", 10, "bold"), bg="#eceff1", fg="#37474f", relief="flat", padx=5, pady=8, width=w, anchor="w")
    lbl.grid(row=0, column=col, sticky="nsew", padx=1, pady=1)

header_frame.grid_columnconfigure(0, weight=0)
header_frame.grid_columnconfigure(1, weight=1)
header_frame.grid_columnconfigure(2, weight=0)
header_frame.grid_columnconfigure(3, weight=0)

table_outer_frame = tk.Frame(table_outer_container, bg="#f5f7fa")
table_outer_frame.pack(fill="both", expand=True)

canvas = tk.Canvas(table_outer_frame, bg="white", highlightthickness=1, highlightbackground="#cfd8dc")
scrollbar = ttk.Scrollbar(table_outer_frame, orient="vertical", command=canvas.yview)
table_inner_frame = tk.Frame(canvas, bg="#cfd8dc") 

table_inner_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

def _on_table_mousewheel(event):
    canvas.yview_scroll(int(-1*(event.delta/120)), "units")

canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_table_mousewheel))
canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

canvas.create_window((0, 0), window=table_inner_frame, anchor="nw", width=canvas.winfo_width())

def _on_canvas_configure(event):
    canvas.itemconfig(canvas.find_withtag("all")[0], width=event.width)
canvas.bind('<Configure>', _on_canvas_configure)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

table_inner_frame.grid_columnconfigure(0, minsize=140, weight=0)
table_inner_frame.grid_columnconfigure(1, minsize=350, weight=1) 
table_inner_frame.grid_columnconfigure(2, minsize=130, weight=0)
table_inner_frame.grid_columnconfigure(3, minsize=220, weight=0)

bottom_frame = tk.Frame(app, bg="#f5f7fa")
bottom_frame.pack(fill="x", padx=10, pady=(0, 15))

export_btn_canvas = create_rounded_button(bottom_frame, "Export", "#0277bd", "#01579b", "white", export_to_pdf, width=130)
export_btn_canvas.pack(anchor="e")

app.mainloop()