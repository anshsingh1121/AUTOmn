def _load_runtime_overrides():
    """In frozen/packaged mode, check for external constants.py.
    Allows runtime modification of credentials without recompilation."""
    if not getattr(sys, "frozen", False):
        return

    if os.environ.get("_CMAA_LOADING_CONSTANTS") == "1":
        return

    os.environ["_CMAA_LOADING_CONSTANTS"] = "1"
    try:
        candidate_dirs = [os.getcwd(), os.path.dirname(sys.executable)]
        meipass = getattr(sys, "_MEIPASS", "")

        target_file = None
        for cdir in candidate_dirs:
            cand = os.path.join(cdir, "constants.py")
            if os.path.isfile(cand):
                if meipass and os.path.abspath(cand).startswith(os.path.abspath(meipass)):
                    continue
                target_file = cand
                break

        if target_file:
            try:
                import runpy
                overrides = runpy.run_path(target_file)
                this_mod = sys.modules.get("constants") or sys.modules.get(__name__)
                if this_mod:
                    for k, v in overrides.items():
                        if not k.startswith("__") and not callable(v):
                            setattr(this_mod, k, v)
                    if "INSTANCE_URL" in overrides and "BASE_URL" not in overrides:
                        this_mod.BASE_URL = f"{str(this_mod.INSTANCE_URL).rstrip('/')}/api/now/table/"
                    if "INSTANCE_URL" in overrides and "ATT_FILE_URL" not in overrides:
                        this_mod.ATT_FILE_URL = f"{str(this_mod.INSTANCE_URL).rstrip('/')}/api/now/attachment/"
                    print(f"[CONFIG] Loaded external configuration override from: {target_file}")
            except Exception as exc:
                print(f"[CONFIG WARNING] Could not load external constants from {target_file}: {exc}")
    finally:
        os.environ.pop("_CMAA_LOADING_CONSTANTS", None)


_load_runtime_overrides()
