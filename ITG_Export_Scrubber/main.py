import itg_scrub_pkg
from itg_scrub_pkg.Application import Application

if __name__ == "__main__":
    app = Application()
    app.grid_columnconfigure(0, weight=1)
    app.mainloop()