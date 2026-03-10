import tkinter as tk
from app.controllers.pistoleo_controller import PistoleoController
from app.views.main_view import MainView

def main():
    controller = PistoleoController()
    app = MainView(controller)
    app.mainloop()

if __name__ == "__main__":
    main()