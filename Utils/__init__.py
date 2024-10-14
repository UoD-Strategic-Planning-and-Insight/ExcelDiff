import math


def convert_alphabetic_number_to_int(alphabetic_number) -> int:
    """
    Converts a number in an alphabetic format (i.e. the format that Excel uses for columns) (e.g. Z = 26, AA = 27) to an
    integer.
    :param alphabetic_number: A string containing an alphabetic representation of a number.
    :return: The integer representation of the passed alphabetic number.
    """

    i: int = 0
    l: int = len(alphabetic_number)

    result: int = 0

    while(i < l):
        c: str = alphabetic_number[l - i - 1]
        result += (ord(c) - 64) * (26 ** i)

        i += 1

    return result


def convert_int_to_alphabetic_number(to_convert: int) -> str:

    result_chars: list[str] = []

    while(to_convert > 0):
        to_convert -= 1
        result_chars.append(chr(to_convert % 26 + ord("A")))
        to_convert //= 26

    result_chars.reverse()
    return "".join(result_chars)


def replace_quote_in_str(source: str) -> str:
    quote_replacement_base: str = "quote$"
    quote_replacement: str = "[quote$1]"
    quote_replacement_int: int = 1

    while(quote_replacement in source):
        quote_replacement_int += 1
        quote_replacement = f"[{quote_replacement_base}{quote_replacement_int}]"

    return source.replace("\"", quote_replacement)


def dict_to_str(dictionary: dict[str, any]) -> str:
    # TODO: Note in documentation that this replaces quotes in the key and value to guarantee uniqueness is retained.

    sorted_keys: list[str] = sorted(dictionary.keys())
    keys_and_vals_as_strs: list[str] = []

    for key in sorted_keys:
        adjusted_key   = replace_quote_in_str(key)
        adjusted_value = replace_quote_in_str(str(dictionary[key]))
        keys_and_vals_as_strs.append(f"\"{adjusted_key}\": \"{adjusted_value}\"")

    return "{" + (", ".join(keys_and_vals_as_strs)) + "}"

