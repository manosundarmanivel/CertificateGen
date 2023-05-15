
from PIL import Image, ImageTk, ImageFont, ImageDraw
import tkinter as tk
import openpyxl


image = Image.open('TECH 3RD.jpg')
window = tk.Tk()
tk_image = ImageTk.PhotoImage(image)
canvas = tk.Canvas(window, width=image.width, height=image.height)
canvas.create_image(0, 0, anchor='nw', image=tk_image)
canvas.pack()
a = []
b = []

def get_pixel_position(event):
    x, y = event.x, event.y
    pixels = image.load()
    print(f"Pixel at ({x}, {y}): {pixels[x, y]}")
    a.append(x)
    b.append(y)
    canvas.unbind('<Button-1>')
    window.quit()

canvas.bind('<Button-1>', get_pixel_position)

window.mainloop()

workbook = openpyxl.load_workbook('Technical Quiz names(1).xlsx')
worksheet = workbook['Sheet1']
column_number = 2
column_values = []
for row in worksheet.iter_rows(min_row=2, min_col=column_number, values_only=True):

    column_values.append(row[0])


font = ImageFont.truetype('Poppins.ttf', 50)

text_list = column_values
for i, text in enumerate(text_list):
    image = Image.open('./TECH 3RD.jpg')
    draw = ImageDraw.Draw(image)
    position = (a[0], b[0])
    color = (0, 0, 0)
    draw.text(position, text.upper(), font=font, fill=color)
    filename = f"{text}_{i}.jpg"
    image.save('./Certificates/'+filename)


