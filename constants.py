def resolve_template_path(template_name: str) -> str:
    """Resolve template path across direct filesystem and PyInstaller --onefile."""
    if not template_name:
        template_name = "QAR Report.xlsx"
    if os.path.exists(template_name):
        return template_name
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        bundled = os.path.join(sys._MEIPASS, os.path.basename(template_name))
        if os.path.exists(bundled):
            return bundled
    return template_name
