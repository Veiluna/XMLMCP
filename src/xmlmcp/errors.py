from __future__ import annotations


class XMLMCPError(Exception):
    def __init__(self, code: str, message: str, **details: object):
        super().__init__(message)
        self.code = code
        self.message = message
        self.details = details

    def as_dict(self) -> dict:
        return {"code": self.code, "message": self.message, "details": self.details}


def require(condition: object, code: str, message: str, **details: object) -> None:
    if not condition:
        raise XMLMCPError(code, message, **details)

