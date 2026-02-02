import os
import sys

def main():
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("PyInstaller ვერ მოიძებნა, მიმდინარეობს დაყენება...")
        os.system(f'"{sys.executable}" -m pip install pyinstaller')

    from PyInstaller.__main__ import run as pyinstaller_run

    base_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(base_dir, "icon.ico")
    has_icon = os.path.exists(icon_path)

    params = [
        "app.py",                # მთავარი სკრიპტი
        "--name=PremiumMed",     # exe ფაილის სახელი
        "--onedir",              # dist/MedicalApp საქაღალდე
        "--noconfirm",
        "--clean",

        # Console რეჟიმი - ეს ვერსია მუშაობს კარგად ბეჭდვაზე
        "--console",

        # რესურსები
        "--add-data=templates;templates",
        "--add-data=static;static",

        # აუცილებელი hidden-import-ები
        "--hidden-import=flask",
        "--collect-all=flask",

        "--hidden-import=docx",
        "--collect-all=docx",

        "--hidden-import=docx2pdf",
        "--collect-all=docx2pdf",

        "--hidden-import=win32com",
        "--hidden-import=win32com.client",
        "--hidden-import=pythoncom",

        "--hidden-import=comtypes",
    ]

    if has_icon:
        params.append(f"--icon={icon_path}")
    else:
        print("გაფრთხილება: icon.ico ვერ მოიძებნა, EXE იქნება ლოგოს გარეშე.")

    print("==============================================")
    print("   სამედიცინო დოკუმენტაცია - EXE აგება")
    print("==============================================")
    print("პროექტის საქაღალდე:", base_dir)
    print("PyInstaller პარამეტრები:")
    for p in params:
        print(" ", p)
    print("==============================================")

    pyinstaller_run(params)

    print("\n✅ აგება დასრულდა!")
    print("გადადით საქაღალდეში: dist\\MedicalApp")
    print("გაუშვით MedicalApp.exe (ან run_hidden.vbs, იხილეთ ქვემოთ)\n")


if __name__ == "__main__":
    main()