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



