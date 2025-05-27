from gui.main_window import ExpenseApp
import tkinter as tk

def main():
    """프로그램 시작점"""
    root = tk.Tk()
    app = ExpenseApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()