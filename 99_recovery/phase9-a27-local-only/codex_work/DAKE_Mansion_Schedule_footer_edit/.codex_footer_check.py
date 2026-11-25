import main


class FakeWidget:
    def __init__(self, parent=None, **kwargs):
        self.parent = parent
        self.kwargs = kwargs
        self.children = []
        self._pack = {}
        self._bindings = {}
        if parent is not None and hasattr(parent, "children"):
            parent.children.append(self)

    def pack(self, **kwargs):
        self._pack = kwargs

    def pack_forget(self):
        self._pack = {}

    def pack_info(self):
        return self._pack

    def bind(self, event, callback, add=None):
        self._bindings[event] = callback

    def cget(self, key):
        return self.kwargs.get(key, "")

    def configure(self, **kwargs):
        self.kwargs.update(kwargs)

    def winfo_children(self):
        return list(self.children)


class FakeRoot(FakeWidget):
    def __init__(self):
        super().__init__(None)
        self.width = 1040

    def winfo_width(self):
        return self.width


class FakeLabel(FakeWidget):
    pass


original_frame = main.tk.Frame
original_label = main.tk.Label
try:
    main.tk.Frame = FakeWidget
    main.tk.Label = FakeLabel

    app = main.MansionScheduleApp.__new__(main.MansionScheduleApp)
    app.root = FakeRoot()
    app.font_family = "TestFont"
    app.footer_compact = None

    parent = FakeWidget()
    app._build_footer(parent)
    wide = (app.footer_left.pack_info().get("side"), app.footer_right.pack_info().get("side"))

    app.root.width = 860
    app._update_footer_layout()
    compact = (app.footer_left.pack_info().get("anchor"), app.footer_right.pack_info().get("anchor"))

    right_widgets = app.footer_right.winfo_children()
    links = [widget.cget("text") for widget in right_widgets if widget.cget("cursor") == "hand2"]
    static = [widget.cget("text") for widget in right_widgets if widget.cget("cursor") != "hand2"]
    left_text = [widget.cget("text") for widget in app.footer_left.winfo_children()]

    assert wide == ("left", "right"), wide
    assert compact == ("center", "center"), compact
    assert links == [main.UI_TEXT["footer_link_1"], main.UI_TEXT["footer_link_2"]], links
    assert main.UI_TEXT["footer_copyright"] in static, static
    assert main.UI_TEXT["footer_separator"] in static, static
    assert main.UI_TEXT["footer_left"] in left_text, left_text
    assert main.UI_TEXT["footer_tagline"] in left_text, left_text
    assert "footer_note" not in main.UI_TEXT
finally:
    main.tk.Frame = original_frame
    main.tk.Label = original_label

print("footer_layout=OK")
print("footer_links=OK")
print("footer_static_text=OK")
