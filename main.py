import tkinter as tk
from tkinter import filedialog
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common import exceptions
from tkinter import messagebox
import os
from PIL import Image
import shutil


def log(out: tk.Text, message: str, timestamp: bool = True, color: str = "black"):
    time = datetime.now().strftime("[%H:%M:%S]")
    out.configure(state="normal")
    if timestamp:
        out.insert(tk.INSERT, f'{time} {message}\n')
    else:
        out.insert(tk.INSERT, f'{message}\n')
    out.configure(state="disabled")
    index = out.index(tk.INSERT)
    current_row = int(index.split('.')[0])
    if color != "black":
        out.tag_add("error", f"{current_row-1}.0", f"{current_row-1}.45")
        out.tag_config("error", foreground=color)
    out.see("end")
    out.update_idletasks()


options = Options()
options.add_argument('start-maximized')
driver = webdriver.Chrome(options=options)


def download(url: str):
    log(text1, f"Getting webpage: {url}")
    driver.get(url)
    WebDriverWait(driver, 10000).until(EC.visibility_of_element_located((By.CLASS_NAME, 'ndfHFb-c4YZDc-cYSp0e-DARUcf')))
    title = "Unnamed.pdf"
    try:
        title_element = driver.find_element_by_xpath('/html/body/div[3]/div[3]/div/div[1]/div[2]/div[1]')
        title = title_element.text
        log(text1, "Loading title")
    except exceptions.NoSuchElementException:
        log(text1, "Error occurred while loading title", color="red")
    slides = driver.find_elements_by_class_name('ndfHFb-c4YZDc-cYSp0e-DARUcf')
    js = "document.getElementsByClassName('ndfHFb-c4YZDc-q77wGc')[0].remove();"
    driver.execute_script(js)
    os.mkdir(f'output/')
    for ind, slide in enumerate(slides):
        with open(f'output/slide{ind + 1}.png', 'wb') as outfile:
            WebDriverWait(slide, 10).until(
                EC.visibility_of_element_located((By.CLASS_NAME, 'ndfHFb-c4YZDc-cYSp0e-DARUcf-PLDbbf')))
            outfile.write(slide.screenshot_as_png)
            log(text1, f"Screenshotting slide# {ind+1}")
    slides[1].screenshot(f'output/slide2.png')
    log(text1, f"Quitting driver")
    driver.quit()
    os.chdir('output')
    slides = os.listdir()
    log(text1, f"Constructing PDF")
    slides_r = []
    slide_1 = None
    for ind, slide in enumerate(slides):
        sl = Image.open(slide)
        sl_r = sl.convert('RGB')
        if ind == 0:
            slide_1 = sl_r
        else:
            slides_r.append(sl_r)
    os.chdir('..')
    log(text1, f"Saving PDF")
    filename = filedialog.asksaveasfilename(initialdir=r"%USERPROFILE%", initialfile=title,  title="Save PDF file as",
                                            filetypes=(("PDF files", "*.pdf"), ("All files", "*.*")))
    if filename.endswith('.pdf'):
        slide_1.save(f'{filename}', save_all=True, append_images=slides_r)
        log(text1, f"PDF saved at {filename}")
    else:
        slide_1.save(f'{filename}.pdf', save_all=True, append_images=slides_r)
        log(text1, f"PDF saved at {filename}.pdf")
    log(text1, f"Clearing image cache")
    try:
        shutil.rmtree('output')
    except OSError as e:
        log(text1, f"Error occurred: {e}", color="red")
    entry1.delete(0, tk.END)
    log(text1, f"---------------------END---------------------", False)


def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        root.destroy()
        driver.quit()


def main():
    global text1, entry1, root
    root = tk.Tk()
    root.resizable(False, False)
    photo = tk.PhotoImage(file="iconphoto.png")
    root.iconphoto(False, photo)
    root.title("Protected download")
    root.geometry('400x300')
    root.protocol("WM_DELETE_WINDOW", on_closing)
    frame1 = tk.Frame()
    frame1.pack(pady=10)
    label1 = tk.Label(frame1, text="URL:")
    label1.grid(column=1, row=1, ipadx=5)
    entry1 = tk.Entry(frame1, width=54)
    entry1.grid(column=2, row=1)
    frame2 = tk.Frame()
    frame2.pack(pady=20)
    text1 = tk.Text(frame2, state="disabled", width=45, height=11, bg="gainsboro")
    text1.grid(row=1)
    button1 = tk.Button(frame2, text="Download", command=lambda: download(entry1.get()))
    button1.grid(row=2, pady=10)

    root.mainloop()


if __name__ == "__main__":
    main()
