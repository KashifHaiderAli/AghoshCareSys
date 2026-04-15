import tkinter as tk
from PIL import Image, ImageTk
import os, sys

def resource_path(relative):
    try:
        base = sys._MEIPASS
    except Exception:
        base = os.path.abspath(".")
    return os.path.join(base, relative)


class SplashScreen(tk.Toplevel):
    def __init__(self, master, on_finish):
        super().__init__(master)

        self.on_finish = on_finish
        self.overrideredirect(True)
        self.attributes("-alpha", 0)
        self.configure(bg="#005BA1")

        w, h = 500, 350
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

        self.logo_label = tk.Label(self, bg="#005BA1")
        self.logo_label.pack(pady=40)

        self.title_label = tk.Label(self, font=("Helvetica", 14, "bold"), bg="#005BA1", fg="#FFFFFF")
        self.title_label.pack()

        self.tagline_label = tk.Label(self, font=("Helvetica", 12, "italic"), bg="#005BA1", fg="#FFFFFF")
        self.tagline_label.pack(pady=5)

        self.steps = [
            ("logo/akh_logo.png", "Al Khidmat Khawateen Trust Pakistan",
             "Service to Humanity With Integrity"),
            ("logo/orphan_logo.png", "Orphan Care Program",
             "Protecting, Educating & Empowering Orphans"),
            ("logo/aghosh_logo.png", "Aghosh Homes",
             "A Home, Not Just a Shelter")
        ]

        self.step = 0
        self.after(200, self.next_step)

    def next_step(self):
        if self.step >= len(self.steps):
            self.destroy()
            self.on_finish()
            return

        logo, title, tagline = self.steps[self.step]
        img = Image.open(resource_path(logo)).resize((160, 160))
        self.photo = ImageTk.PhotoImage(img)

        self.logo_label.config(image=self.photo)
        self.title_label.config(text=title)
        self.tagline_label.config(text=tagline)

        self.attributes("-alpha", 0)
        self.fade_in(0)

    def fade_in(self, a):
        if a >= 1:
            self.after(900, self.fade_out, 1)
            return
        self.attributes("-alpha", a)
        self.after(35, self.fade_in, a + 0.05)

    def fade_out(self, a):
        if a <= 0:
            self.step += 1
            self.next_step()
            return
        self.attributes("-alpha", a)
        self.after(35, self.fade_out, a - 0.05)
