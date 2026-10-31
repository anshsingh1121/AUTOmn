def load_template(template_path):
    if not os.path.exists(template_path):
        # Support PyInstaller --onefile bundled template fallback
        if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
            bundled = os.path.join(sys._MEIPASS, os.path.basename(template_path))
            if os.path.exists(bundled):
                template_path = bundled
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template file not found at: {template_path}")
    print(f"Loading the template from {template_path}")
    return openpyxl.load_workbook(template_path)
