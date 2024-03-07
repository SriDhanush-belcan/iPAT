import tkinter as tk
from tkinter import filedialog
from pptx import Presentation
from pptx.util import Inches

def insert_data():
    eng_number = eng_number_entry.get()
    engine_module = engine_module_entry.get()
    detail_pn = detail_pn_entry.get()
    assy_pn = assy_pn_entry.get()
    vendor = vendor_entry.get()
    qn = qn_entry.get()
    sn = sn_entry.get()
    
    data = {'ENG Number': eng_number,
            'Engine Module': engine_module,
            'Detail P/N': detail_pn,
            'Assy P/N': assy_pn,
            'Vendor': vendor,
            'QN': qn,
            'S/N': sn}
    
    selected_file = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
    
    if selected_file:
        prs = Presentation(selected_file)
        slide = prs.slides[0]  # Assuming the data should be inserted into the first slide
        
        bullet_points = []
        for key, value in data.items():
            bullet_points.append(f"{key}: {value}")
        
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(4))
        text_frame = textbox.text_frame
        
        for point in bullet_points:
            p = text_frame.add_paragraph()
            p.text = point
                
        prs.save(selected_file)
        tk.messagebox.showinfo("Success", "Data inserted into PowerPoint successfully!")

# GUI setup
root = tk.Tk()
root.title("Insert Data into PowerPoint")

tk.Label(root, text="ENG Number:").grid(row=0, column=0)
eng_number_entry = tk.Entry(root)
eng_number_entry.grid(row=0, column=1)

tk.Label(root, text="Engine Module:").grid(row=1, column=0)
engine_module_entry = tk.Entry(root)
engine_module_entry.grid(row=1, column=1)

tk.Label(root, text="Detail P/N:").grid(row=2, column=0)
detail_pn_entry = tk.Entry(root)
detail_pn_entry.grid(row=2, column=1)

tk.Label(root, text="Assy P/N:").grid(row=3, column=0)
assy_pn_entry = tk.Entry(root)
assy_pn_entry.grid(row=3, column=1)

tk.Label(root, text="Vendor:").grid(row=4, column=0)
vendor_entry = tk.Entry(root)
vendor_entry.grid(row=4, column=1)

tk.Label(root, text="QN:").grid(row=5, column=0)
qn_entry = tk.Entry(root)
qn_entry.grid(row=5, column=1)

tk.Label(root, text="S/N:").grid(row=6, column=0)
sn_entry = tk.Entry(root)
sn_entry.grid(row=6, column=1)

insert_button = tk.Button(root, text="Insert Data into PowerPoint", command=insert_data)
insert_button.grid(row=7, columnspan=2)

root.mainloop()
