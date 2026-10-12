"""SpreadsheetML ST_Xstring conversion, kept separate from XML escaping."""
import re

TOKEN = re.compile(r"_x([0-9a-fA-F]{4})_")


def encode_xstring(value: str) -> str:
    # Protect literal escape tokens before introducing the CR escape itself.
    return TOKEN.sub(lambda m: "_x005F_" + m.group(0)[1:], value).replace("\r", "_x000D_")


def decode_xstring(value: str) -> str:
    decoded = TOKEN.sub(lambda m: chr(int(m[1], 16)), value)
    # Supplementary characters can be stored as two escaped UTF-16 code units.
    return decoded.encode("utf-16-le", "surrogatepass").decode("utf-16-le", "surrogatepass")
