print("Testing imports...")

try:
    import flask
    print("✅ Flask OK:", flask.__version__)
except ImportError as e:
    print("❌ Flask FAILED:", e)

try:
    import docx
    print("✅ python-docx OK")
except ImportError as e:
    print("❌ python-docx FAILED:", e)

try:
    import werkzeug
    print("✅ Werkzeug OK:", werkzeug.__version__)
except ImportError as e:
    print("❌ Werkzeug FAILED:", e)

try:
    import lxml
    print("✅ lxml OK")
except ImportError as e:
    print("❌ lxml FAILED:", e)

print("\nAll imports tested!")
input("Press Enter to exit...")