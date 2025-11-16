import customtkinter
from PIL import Image

from ngs import *
from sanger import *

my_image = customtkinter.CTkImage(light_image=Image.open("Trim2Sort_icon.png"), size=(128, 72))


# 主畫面設定
class App(customtkinter.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("280x170")
        self.resizable(False, False)
        self.title("Trim2Sort")
        self.configure(fg_color="#091235")

        # 設定Grid System(2*2)
        self.grid_rowconfigure(0, weight=0)
        self.grid_columnconfigure((0, 1), weight=1)

        # Tirm2Sort Logo的參數
        self.logo = customtkinter.CTkLabel(self, image=my_image, text="")
        self.logo.grid(row=0, column=0, columnspan=2, padx=20)

        self.label = customtkinter.CTkLabel(
            self,
            text="Step0: Choose the sequence format",
            fg_color="transparent",
            text_color="#DCF3F0",
            font=("Consolas", 12, "bold"),
        )
        self.label.grid(row=1, column=0, columnspan=2, padx=20, pady=(10, 5))

        self.button_sanger = customtkinter.CTkButton(
            self, text="Sanger", command=self.OpenSangerAnalysis
        )
        self.button_sanger.configure(
            fg_color="#2B4257",
            hover_color="#4C6A78",
            text_color="#DCF3F0",
            font=("Consolas", 12, "bold"),
            corner_radius=8,
        )
        self.button_sanger.grid(row=2, column=0, padx=(10, 10), pady=(10, 10), sticky="ew")

        self.button_NGS = customtkinter.CTkButton(self, text="NGS", command=self.OpenNGSAnalysis)
        self.button_NGS.configure(
            fg_color="#2B4257",
            hover_color="#4C6A78",
            text_color="#DCF3F0",
            font=("Consolas", 12, "bold"),
            corner_radius=8,
        )
        self.button_NGS.grid(row=2, column=1, padx=(10, 10), pady=(10, 10), sticky="ew")

        self.toplevel_window = None

    def OpenSangerAnalysis(self):
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = Sanger(self)  # create window if its None or destroyed
        else:
            self.toplevel_window.focus()  # if window exists focus it

    def OpenNGSAnalysis(self):
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = NGS(self)  # create window if its None or destroyed
        else:
            self.toplevel_window.focus()  # if window exists focus it


if __name__ == "__main__":
    app = App()
    app.mainloop()
