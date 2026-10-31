def load_workbook(template_path):
    resolved_path = resolve_template_path(template_path)
    try:
        return excel_handler.load_template(resolved_path)
    except FileNotFoundError:
        print(f"Error: Template file '{template_path}' not found (resolved: '{resolved_path}').")
        sys.exit(1)
