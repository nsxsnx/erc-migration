"""Error/warning message handlers"""

import logging
from abc import ABC


class ErrorMessageHandler(ABC):
    "Shows error/warning messages only once"

    def __init__(self) -> None:
        self.error_messages: dict[str, list] = {}

    def _save_message(self, error_type: str, error_hash_key: str) -> bool:
        "Remembers new error message of a particular type"
        error_list: list = self.error_messages.get(error_type, [])
        error_list.append(error_hash_key)
        self.error_messages.update({error_type: error_list})

    def _is_new(self, error_type: str, error_hash_key: str) -> bool:
        "Checks if hash_key has been seen before. If not, returns True; False otherwise."
        error_list: list = self.error_messages.get(error_type, [])
        if error_hash_key in error_list:
            return False
        return True

    def show(self, *args: object) -> None:
        """
        Shows error message via 'self._show()' only once

        ARGS:
            error_type     : str - indicates type of error message
            error_hash_key : str - unique hash key to determine a message
            log_message    : str - lazy-formatted error message
            log_args       ...   - additional arguments for error message
        """
        if len(args) < 3:
            raise ValueError("Too few arguments given")
        error_type: str = args[0]
        error_hash_key: str = args[1]
        if self._is_new(error_type, error_hash_key):
            self._save_message(error_type, error_hash_key)
            self._show(args[2], *args[3:])

    def _show(self, *args):
        "Actually shows message"
        raise NotImplementedError


class ErrorMessageConsoleHandler(ErrorMessageHandler):
    "Shows messages via logging module"

    def __init__(self) -> None:
        super().__init__()
        self._show = logging.warning
