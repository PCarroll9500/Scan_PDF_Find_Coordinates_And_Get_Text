import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from PIL import Image, ImageTk
import json
import matplotlib.pyplot as plt

class PDFViewer(tk.Tk):
    def __init__(self, pdf_path):
        super().__init__()
        self.title("PDF Viewer")
        self.geometry("800x600")

        self.pdf_path = pdf_path
        self.current_page_number = 1
        self.rectangles = {}  # Dictionary to store rectangles with names and page numbers
        self.current_rect = None
        self.start_x = self.start_y = 0

        self.doc = fitz.open(pdf_path)
        self.canvas = tk.Canvas(self, bg='white')
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.create_navigation_buttons()

        self.load_page()

        self.canvas.bind("<Button-1>", self.on_click)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)

    def create_navigation_buttons(self):
        button_frame = tk.Frame(self)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X)

        prev_button = tk.Button(button_frame, text="Previous Page", command=self.prev_page)
        prev_button.pack(side=tk.LEFT)

        next_button = tk.Button(button_frame, text="Next Page", command=self.next_page)
        next_button.pack(side=tk.LEFT)

        extract_button = tk.Button(button_frame, text="Extract Text", command=self.extract_text_from_boxes)
        extract_button.pack(side=tk.LEFT)

        save_button = tk.Button(button_frame, text="Save Boxes", command=self.save_boxes)
        save_button.pack(side=tk.LEFT)

        load_button = tk.Button(button_frame, text="Load Boxes", command=self.load_boxes)
        load_button.pack(side=tk.LEFT)

        clear_button = tk.Button(button_frame, text="Clear Boxes", command=self.clear_boxes)
        clear_button.pack(side=tk.LEFT)

    def on_click(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)

        if self.current_rect:
            self.canvas.delete(self.current_rect)

        self.current_rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="red", width=2)

    def on_drag(self, event):
        self.canvas.coords(self.current_rect, self.start_x, self.start_y, self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))

    def on_release(self, event):
        if self.current_rect:
            x1, y1, x2, y2 = self.canvas.coords(self.current_rect)
            if x1 != x2 and y1 != y2:  # Ensure the rectangle is not degenerate
                box_name = simpledialog.askstring("Box Name", "Enter a name for this box:")
                if box_name:
                    self.rectangles.setdefault(self.current_page_number, []).append({
                        'name': box_name,
                        'coords': (x1, y1, x2, y2)
                    })
            self.current_rect = None

    def prev_page(self):
        if self.current_page_number > 0:
            self.save_current_page_boxes()
            self.current_page_number -= 1
            self.load_page()

    def next_page(self):
        if self.current_page_number < len(self.doc) - 1:
            self.save_current_page_boxes()
            self.current_page_number += 1
            self.load_page()
        else:
            messagebox.showinfo("Info", "No more pages to go forward.")

    def load_page(self):
        self.page = self.doc.load_page(self.current_page_number)
        self.pagemap = self.page.get_pixmap()
        self.image = Image.frombytes("RGB", [self.pagemap.width, self.pagemap.height], self.pagemap.samples)
        self.img_tk = ImageTk.PhotoImage(image=self.image)

        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.img_tk)

        # Draw previously saved rectangles for the current page
        for box in self.rectangles.get(self.current_page_number, []):
            x1, y1, x2, y2 = box['coords']
            self.canvas.create_rectangle(x1, y1, x2, y2, outline="red", width=2, tags=box['name'])

    def save_current_page_boxes(self):
        # Save rectangles of the current page before moving to another page
        if self.current_page_number in self.rectangles:
            self.rectangles[self.current_page_number] = [
                {'name': box['name'],
                 'coords': (self.canvas.coords(self.canvas.find_withtag(box['name'])[0])[0],
                            self.canvas.coords(self.canvas.find_withtag(box['name'])[0])[1],
                            self.canvas.coords(self.canvas.find_withtag(box['name'])[0])[2],
                            self.canvas.coords(self.canvas.find_withtag(box['name'])[0])[3])}
                for box in self.rectangles[self.current_page_number]
            ]

    def extract_text_from_boxes(self):
        extracted_text = []
        for box in self.rectangles.get(self.current_page_number, []):
            x1, y1, x2, y2 = box['coords']
            pdf_rect = fitz.Rect(x1, y1, x2, y2)
            text = self.page.get_text("text", clip=pdf_rect)
            extracted_text.append(f"{box['name']}: {text}")

        if extracted_text:
            plt.figure(figsize=(10, 7))
            plt.text(0.5, 0.5, '\n'.join(extracted_text), fontsize=12, ha='center')
            plt.axis('off')
            plt.show()

    def save_boxes(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")])
        if save_path:
            with open(save_path, 'w') as file:
                json.dump(self.rectangles, file)
            messagebox.showinfo("Info", "Boxes saved successfully!")

    def load_boxes(self):
        load_path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if load_path:
            with open(load_path, 'r') as file:
                self.rectangles = json.load(file)
            self.load_page()
            messagebox.showinfo("Info", "Boxes loaded successfully!")

    def clear_boxes(self):
        if self.current_page_number in self.rectangles:
            self.rectangles[self.current_page_number] = []
        self.canvas.delete("all")
        self.load_page()

def main():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if pdf_path:
        app = PDFViewer(pdf_path)
        app.mainloop()

if __name__ == "__main__":
    main()
