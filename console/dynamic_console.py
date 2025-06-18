import os

# ANSI escape codes for color console output
# This will work in most modern terminals.
class DynamicConsole:
    INFO = '\033[0m'
    WARNING = '\033[93m'
    ERROR = '\033[91m'
    RESET = '\033[0m'

    @staticmethod
    def print_message(message, type="info"):
        """Prints the message to the console with the corresponding color."""
        if type == "error":
            print(f"{DynamicConsole.ERROR}{message}{DynamicConsole.RESET}")
        elif type == "warning":
            print(f"{DynamicConsole.WARNING}{message}{DynamicConsole.RESET}")
        else:
            print(f"{DynamicConsole.INFO}{message}{DynamicConsole.RESET}")

# To ensure DynamicConsole can also be used in other modules without circular imports
# from os.path.basename, we move the os import here.
