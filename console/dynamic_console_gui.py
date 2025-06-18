import tkinter as tk
from tkinter import scrolledtext
from datetime import datetime # Import datetime here as well for the timestamp

class DynamicConsoleGUI:
    _instance = None # Singleton instance

    def __new__(cls, master=None):
        if cls._instance is None:
            cls._instance = super(DynamicConsoleGUI, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self, master=None):
        if self._initialized:
            return
        self.text_widget = None
        # No need to define tags on 'master' here.
        # Tags will be defined on the 'text_widget' once it's set.
        self._initialized = True

    def set_text_widget(self, text_widget: scrolledtext.ScrolledText):
        """Sets the ScrolledText widget where messages will be displayed."""
        self.text_widget = text_widget
        # Ensure tags are defined on the text_widget if it's the first time
        if not hasattr(self.text_widget, '_console_tags_defined_on_widget'):
            self.text_widget.tag_config("info", foreground="black")
            self.text_widget.tag_config("warning", foreground="orange")
            self.text_widget.tag_config("error", foreground="red")
            self.text_widget.tag_config("success", foreground="green")
            self.text_widget._console_tags_defined_on_widget = True # type: ignore


    def print_message(self, message: str, message_type: str = "info"):
        """Prints a message to the console widget with a specified type (color)."""
        if self.text_widget:
            self.text_widget.config(state='normal') # Enable editing
            self.text_widget.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n", message_type)
            self.text_widget.see(tk.END) # Scroll to the end
            self.text_widget.config(state='disabled') # Disable editing
        else:
            # Fallback to print if no text widget is set (e.g., during testing or early startup)
            print(f"[{message_type.upper()}] {message}")

    def clear_log(self):
        """Clears all messages from the console log."""
        if self.text_widget:
            self.text_widget.config(state='normal')
            self.text_widget.delete(1.0, tk.END)
            self.text_widget.config(state='disabled')