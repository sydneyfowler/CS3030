'''
custom_regular_expressions.py (Excel Command Line Tool)
Sydney Fowler and Matthew Hileman
15 December 2019
Description: Contains all regular expressions needed for some tools.
'''

# ================ REFERENCES ================
# Email rules derived from: https://help.returnpath.com/hc/en-us/articles/220560587-What-are-the-rules-for-email-address-syntax-
# Web adress derived from: https://www.regextester.com/93652

# ================ IMPORTS ================
# System
import re

# ================== REGULAR EXPRESSION ===================
strip_none_digits = re.compile(r'(\d+)')

remove_special_characters = re.compile(r'''[,.:;=?!"*%<>\-_(){}[\]\\]''')

phone_regex = re.compile(r'''(
                (?P<area_code>\d{3}|\((\s+)?\d{3}(\s+)?\)|\[(\s+)?\d{3}(\s+)?\])?       # Area code
                ((\s+)?(\s|-|\.)?(\s+)?)?                                               # Separator
                (?P<three_digits>\d{3})                                                 # First 3 digits
                ((\s+)?(\s|-|\.)?(\s+)?)?                                               # Separator
                (?P<four_digits>\d{4})                                                  # Last 4 digits
                (\s*(ext|x|ext.)\s*(?P<ext>\d{2,5}))?                                   # Extension
                )''', re.VERBOSE)

# Derived email rules from site in REFERENCES.
email_regex = re.compile(r'''(
                ([a-zA-Z0-9](([a-zA-Z0-9!#$%&'*+/=?^_`{|.-]){,62}[a-zA-Z0-9])?)     # Recipient name
                (@)                                                                 # @ symbol
                ([a-zA-Z0-9](([a-zA-Z0-9.-]){,251}[a-zA-Z0-9])?)                    # Domain name
                (\.)                                                                # . symbol
                (com|org|net|edu|co)                                                # Top-level domain
                )''', re.VERBOSE)

zip_regex = re.compile(r'''(
                (?P<five_digits>(\d)?(\d{4}))       # 5 digits
                ((\s+)?(-)(\s+)?)?                  # -
                (?P<four_digits>\d{4})?             # 4 digits
                )''', re.VERBOSE)

yyyy_mm_dd = re.compile(r'''(
                (?P<year>\d{4})                         # Year
                (-|/|\s)                                # Separator (- or / or ' ')
                (?P<month>(1[0-2])|[0]?[1-9])           # Month
                (-|/|\s)                                # Separator (- or / or ' ')
                (?P<day>(3[0-1])|[0]?[1-9]|[1-2][0-9])  # Day
                )''', re.VERBOSE)

mm_dd_yyyy = re.compile(r'''(
                (?P<month>([0]?[1-9])|([1][0-2]))               # Month
                (-|/|\s)                                        # Separator (- or / or ' ')
                (?P<day>([0]?[1-9])|([1-2][0-9])|([3][0-1]))    # Day
                (-|/|\s)                                        # Separator (- or / or ' ')
                (?P<year>(\d{2})?(\d{2}))                       # Year
                )''', re.VERBOSE)

month_word = re.compile(r'''(
                (?P<month>JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC|JANUARY|FEBRUARY|MARCH|APRIL|JUNE|JULY|AUGUST
                          |SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)                 # Month
                (\s|\.)+                                                        # Separator
                (?P<day>([0]?[1-9])|([1-2][0-9])|([3][0-1]))                    # Day
                (\s|\.|,|')+                                                    # Separator
                (?P<year>(\d{2})?(\d{2}))                                       # Year
                )''', re.VERBOSE)

# Derived web address from site in REFERENCES.
web_address_regex = re.compile(r'''(
                (http:\/\/www\.|https:\/\/www\.|http:\/\/|https:\/\/)?
                [a-z0-9]+
                ([\-\.]{1}[a-z0-9]+)*
                \.
                [a-z]{2,5}
                (:[0-9]{1,5})?
                (\/.*)?
                )''', re.VERBOSE)
