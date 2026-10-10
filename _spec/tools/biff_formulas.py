#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""biff_formulas.py v2: trích công thức trực tiếp từ BIFF8 (.xls/.xls1) và OOXML (.xlsx).

Chạy:  python3 -I biff_formulas.py --src <thư mục nguồn> --menu <menu.xml> --out <_spec/formulas>
Cần olefile (có trong /tmp/audit_venv). Không chạy mã trong tệp nguồn; chỉ đọc byte.

Quy ước đầu ra:
  - formula: chuỗi A1 không có dấu "=" đầu; CR/LF được JSON escape.
  - tham chiếu ngoài: [ext:n]Sheet!A1, n = chỉ số bản ghi SUPBOOK (đếm từ 0) trong tệp.
  - shared formula: mở rộng cho từng ô (kind=shared, shared_ref = vùng SHRFMLA).
  - array: kind=array, array_ref = vùng ARRAY; mọi ô trong vùng mang cùng công thức.
"""
import sys, os, re, struct, json, hashlib, zipfile, argparse
import xml.etree.ElementTree as ET
try:
    import olefile
except ImportError:
    sys.exit('Thiếu olefile: chạy bằng /tmp/audit_venv/bin/python -I')

GENERATOR = 'biff_formulas.py'
VERSION = '2.0.0'

FTAB = {
    0: "COUNT",
    1: "IF",
    2: "ISNA",
    3: "ISERROR",
    4: "SUM",
    5: "AVERAGE",
    6: "MIN",
    7: "MAX",
    8: "ROW",
    9: "COLUMN",
    10: "NA",
    11: "NPV",
    12: "STDEV",
    13: "DOLLAR",
    14: "FIXED",
    15: "SIN",
    16: "COS",
    17: "TAN",
    18: "ATAN",
    19: "PI",
    20: "SQRT",
    21: "EXP",
    22: "LN",
    23: "LOG10",
    24: "ABS",
    25: "INT",
    26: "SIGN",
    27: "ROUND",
    28: "LOOKUP",
    29: "INDEX",
    30: "REPT",
    31: "MID",
    32: "LEN",
    33: "VALUE",
    34: "TRUE",
    35: "FALSE",
    36: "AND",
    37: "OR",
    38: "NOT",
    39: "MOD",
    40: "DCOUNT",
    41: "DSUM",
    42: "DAVERAGE",
    43: "DMIN",
    44: "DMAX",
    45: "DSTDEV",
    46: "VAR",
    47: "DVAR",
    48: "TEXT",
    49: "LINEST",
    50: "TREND",
    51: "LOGEST",
    52: "GROWTH",
    53: "GOTO",
    54: "HALT",
    55: "RETURN",
    56: "PV",
    57: "FV",
    58: "NPER",
    59: "PMT",
    60: "RATE",
    61: "MIRR",
    62: "IRR",
    63: "RAND",
    64: "MATCH",
    65: "DATE",
    66: "TIME",
    67: "DAY",
    68: "MONTH",
    69: "YEAR",
    70: "WEEKDAY",
    71: "HOUR",
    72: "MINUTE",
    73: "SECOND",
    74: "NOW",
    75: "AREAS",
    76: "ROWS",
    77: "COLUMNS",
    78: "OFFSET",
    79: "ABSREF",
    80: "RELREF",
    81: "ARGUMENT",
    82: "SEARCH",
    83: "TRANSPOSE",
    84: "ERROR",
    85: "STEP",
    86: "TYPE",
    87: "ECHO",
    88: "SET.NAME",
    89: "CALLER",
    90: "DEREF",
    91: "WINDOWS",
    92: "SERIES",
    93: "DOCUMENTS",
    94: "ACTIVE.CELL",
    95: "SELECTION",
    96: "RESULT",
    97: "ATAN2",
    98: "ASIN",
    99: "ACOS",
    100: "CHOOSE",
    101: "HLOOKUP",
    102: "VLOOKUP",
    103: "LINKS",
    104: "INPUT",
    105: "ISREF",
    106: "GET.FORMULA",
    107: "GET.NAME",
    108: "SET.VALUE",
    109: "LOG",
    110: "EXEC",
    111: "CHAR",
    112: "LOWER",
    113: "UPPER",
    114: "PROPER",
    115: "LEFT",
    116: "RIGHT",
    117: "EXACT",
    118: "TRIM",
    119: "REPLACE",
    120: "SUBSTITUTE",
    121: "CODE",
    122: "NAMES",
    123: "DIRECTORY",
    124: "FIND",
    125: "CELL",
    126: "ISERR",
    127: "ISTEXT",
    128: "ISNUMBER",
    129: "ISBLANK",
    130: "T",
    131: "N",
    132: "FOPEN",
    133: "FCLOSE",
    134: "FSIZE",
    135: "FREADLN",
    136: "FREAD",
    137: "FWRITELN",
    138: "FWRITE",
    139: "FPOS",
    140: "DATEVALUE",
    141: "TIMEVALUE",
    142: "SLN",
    143: "SYD",
    144: "DDB",
    145: "GET.DEF",
    146: "REFTEXT",
    147: "TEXTREF",
    148: "INDIRECT",
    149: "REGISTER",
    150: "CALL",
    151: "ADD.BAR",
    152: "ADD.MENU",
    153: "ADD.COMMAND",
    154: "ENABLE.COMMAND",
    155: "CHECK.COMMAND",
    156: "RENAME.COMMAND",
    157: "SHOW.BAR",
    158: "DELETE.MENU",
    159: "DELETE.COMMAND",
    160: "GET.CHART.ITEM",
    161: "DIALOG.BOX",
    162: "CLEAN",
    163: "MDETERM",
    164: "MINVERSE",
    165: "MMULT",
    166: "FILES",
    167: "IPMT",
    168: "PPMT",
    169: "COUNTA",
    170: "CANCEL.KEY",
    171: "FOR",
    172: "WHILE",
    173: "BREAK",
    174: "NEXT",
    175: "INITIATE",
    176: "REQUEST",
    177: "POKE",
    178: "EXECUTE",
    179: "TERMINATE",
    180: "RESTART",
    181: "HELP",
    182: "GET.BAR",
    183: "PRODUCT",
    184: "FACT",
    185: "GET.CELL",
    186: "GET.WORKSPACE",
    187: "GET.WINDOW",
    188: "GET.DOCUMENT",
    189: "DPRODUCT",
    190: "ISNONTEXT",
    191: "GET.NOTE",
    192: "NOTE",
    193: "STDEVP",
    194: "VARP",
    195: "DSTDEVP",
    196: "DVARP",
    197: "TRUNC",
    198: "ISLOGICAL",
    199: "DCOUNTA",
    200: "DELETE.BAR",
    201: "UNREGISTER",
    204: "USDOLLAR",
    205: "FINDB",
    206: "SEARCHB",
    207: "REPLACEB",
    208: "LEFTB",
    209: "RIGHTB",
    210: "MIDB",
    211: "LENB",
    212: "ROUNDUP",
    213: "ROUNDDOWN",
    214: "ASC",
    215: "DBCS",
    216: "RANK",
    219: "ADDRESS",
    220: "DAYS360",
    221: "TODAY",
    222: "VDB",
    223: "ELSE",
    224: "ELSE.IF",
    225: "END.IF",
    226: "FOR.CELL",
    227: "MEDIAN",
    228: "SUMPRODUCT",
    229: "SINH",
    230: "COSH",
    231: "TANH",
    232: "ASINH",
    233: "ACOSH",
    234: "ATANH",
    235: "DGET",
    236: "CREATE.OBJECT",
    237: "VOLATILE",
    238: "LAST.ERROR",
    239: "CUSTOM.UNDO",
    240: "CUSTOM.REPEAT",
    241: "FORMULA.CONVERT",
    242: "GET.LINK.INFO",
    243: "TEXT.BOX",
    244: "INFO",
    245: "GROUP",
    246: "GET.OBJECT",
    247: "DB",
    248: "PAUSE",
    251: "RESUME",
    252: "FREQUENCY",
    253: "ADD.TOOLBAR",
    254: "DELETE.TOOLBAR",
    255: "User",
    256: "RESET.TOOLBAR",
    257: "EVALUATE",
    258: "GET.TOOLBAR",
    259: "GET.TOOL",
    260: "SPELLING.CHECK",
    261: "ERROR.TYPE",
    262: "APP.TITLE",
    263: "WINDOW.TITLE",
    264: "SAVE.TOOLBAR",
    265: "ENABLE.TOOL",
    266: "PRESS.TOOL",
    267: "REGISTER.ID",
    268: "GET.WORKBOOK",
    269: "AVEDEV",
    270: "BETADIST",
    271: "GAMMALN",
    272: "BETAINV",
    273: "BINOMDIST",
    274: "CHIDIST",
    275: "CHIINV",
    276: "COMBIN",
    277: "CONFIDENCE",
    278: "CRITBINOM",
    279: "EVEN",
    280: "EXPONDIST",
    281: "FDIST",
    282: "FINV",
    283: "FISHER",
    284: "FISHERINV",
    285: "FLOOR",
    286: "GAMMADIST",
    287: "GAMMAINV",
    288: "CEILING",
    289: "HYPGEOMDIST",
    290: "LOGNORMDIST",
    291: "LOGINV",
    292: "NEGBINOMDIST",
    293: "NORMDIST",
    294: "NORMSDIST",
    295: "NORMINV",
    296: "NORMSINV",
    297: "STANDARDIZE",
    298: "ODD",
    299: "PERMUT",
    300: "POISSON",
    301: "TDIST",
    302: "WEIBULL",
    303: "SUMXMY2",
    304: "SUMX2MY2",
    305: "SUMX2PY2",
    306: "CHITEST",
    307: "CORREL",
    308: "COVAR",
    309: "FORECAST",
    310: "FTEST",
    311: "INTERCEPT",
    312: "PEARSON",
    313: "RSQ",
    314: "STEYX",
    315: "SLOPE",
    316: "TTEST",
    317: "PROB",
    318: "DEVSQ",
    319: "GEOMEAN",
    320: "HARMEAN",
    321: "SUMSQ",
    322: "KURT",
    323: "SKEW",
    324: "ZTEST",
    325: "LARGE",
    326: "SMALL",
    327: "QUARTILE",
    328: "PERCENTILE",
    329: "PERCENTRANK",
    330: "MODE",
    331: "TRIMMEAN",
    332: "TINV",
    334: "MOVIE.COMMAND",
    335: "GET.MOVIE",
    336: "CONCATENATE",
    337: "POWER",
    338: "PIVOT.ADD.DATA",
    339: "GET.PIVOT.TABLE",
    340: "GET.PIVOT.FIELD",
    341: "GET.PIVOT.ITEM",
    342: "RADIANS",
    343: "DEGREES",
    344: "SUBTOTAL",
    345: "SUMIF",
    346: "COUNTIF",
    347: "COUNTBLANK",
    348: "SCENARIO.GET",
    349: "OPTIONS.LISTS.GET",
    350: "ISPMT",
    351: "DATEDIF",
    352: "DATESTRING",
    353: "NUMBERSTRING",
    354: "ROMAN",
    355: "OPEN.DIALOG",
    356: "SAVE.DIALOG",
    357: "VIEW.GET",
    358: "GETPIVOTDATA",
    359: "HYPERLINK",
    360: "PHONETIC",
    361: "AVERAGEA",
    362: "MAXA",
    363: "MINA",
    364: "STDEVPA",
    365: "VARPA",
    366: "STDEVA",
    367: "VARA",
    368: "BAHTTEXT",
    369: "THAIDAYOFWEEK",
    370: "THAIDIGIT",
    371: "THAIMONTHOFYEAR",
    372: "THAINUMSOUND",
    373: "THAINUMSTRING",
    374: "THAISTRINGLENGTH",
    375: "ISTHAIDIGIT",
    376: "ROUNDBAHTDOWN",
    377: "ROUNDBAHTUP",
    378: "THAIYEAR",
    379: "RTD",
    380: "CUBEVALUE",
    381: "CUBEMEMBER",
    382: "CUBEMEMBERPROPERTY",
    383: "CUBERANKEDMEMBER",
    384: "HEX2BIN",
    385: "HEX2DEC",
    386: "HEX2OCT",
    387: "DEC2BIN",
    388: "DEC2HEX",
    389: "DEC2OCT",
    390: "OCT2BIN",
    391: "OCT2HEX",
    392: "OCT2DEC",
    393: "BIN2DEC",
    394: "BIN2OCT",
    395: "BIN2HEX",
    396: "IMSUB",
    397: "IMDIV",
    398: "IMPOWER",
    399: "IMABS",
    400: "IMSQRT",
    401: "IMLN",
    402: "IMLOG2",
    403: "IMLOG10",
    404: "IMSIN",
    405: "IMCOS",
    406: "IMEXP",
    407: "IMARGUMENT",
    408: "IMCONJUGATE",
    409: "IMAGINARY",
    410: "IMREAL",
    411: "COMPLEX",
    412: "IMSUM",
    413: "IMPRODUCT",
    414: "SERIESSUM",
    415: "FACTDOUBLE",
    416: "SQRTPI",
    417: "QUOTIENT",
    418: "DELTA",
    419: "GESTEP",
    420: "ISEVEN",
    421: "ISODD",
    422: "MROUND",
    423: "ERF",
    424: "ERFC",
    425: "BESSELJ",
    426: "BESSELK",
    427: "BESSELY",
    428: "BESSELI",
    429: "XIRR",
    430: "XNPV",
    431: "PRICEMAT",
    432: "YIELDMAT",
    433: "INTRATE",
    434: "RECEIVED",
    435: "DISC",
    436: "PRICEDISC",
    437: "YIELDDISC",
    438: "TBILLEQ",
    439: "TBILLPRICE",
    440: "TBILLYIELD",
    441: "PRICE",
    442: "YIELD",
    443: "DOLLARDE",
    444: "DOLLARFR",
    445: "NOMINAL",
    446: "EFFECT",
    447: "CUMPRINC",
    448: "CUMIPMT",
    449: "EDATE",
    450: "EOMONTH",
    451: "YEARFRAC",
    452: "COUPDAYBS",
    453: "COUPDAYS",
    454: "COUPDAYSNC",
    455: "COUPNCD",
    456: "COUPNUM",
    457: "COUPPCD",
    458: "DURATION",
    459: "MDURATION",
    460: "ODDLPRICE",
    461: "ODDLYIELD",
    462: "ODDFPRICE",
    463: "ODDFYIELD",
    464: "RANDBETWEEN",
    465: "WEEKNUM",
    466: "AMORDEGRC",
    467: "AMORLINC",
    468: "CONVERT",
    469: "ACCRINT",
    470: "ACCRINTM",
    471: "WORKDAY",
    472: "NETWORKDAYS",
    473: "GCD",
    474: "MULTINOMIAL",
    475: "LCM",
    476: "FVSCHEDULE",
    477: "CUBEKPIMEMBER",
    478: "CUBESET",
    479: "CUBESETCOUNT",
    480: "IFERROR",
    481: "COUNTIFS",
    482: "SUMIFS",
    483: "AVERAGEIF",
    484: "AVERAGEIFS",
    724: "SHEETJS",
}
FTAB_ARGC = {
    2: 1,
    3: 1,
    10: 0,
    15: 1,
    16: 1,
    17: 1,
    18: 1,
    19: 0,
    20: 1,
    21: 1,
    22: 1,
    23: 1,
    24: 1,
    25: 1,
    26: 1,
    27: 2,
    30: 2,
    31: 3,
    32: 1,
    33: 1,
    34: 0,
    35: 0,
    38: 1,
    39: 2,
    40: 3,
    41: 3,
    42: 3,
    43: 3,
    44: 3,
    45: 3,
    47: 3,
    48: 2,
    53: 1,
    61: 3,
    63: 0,
    65: 3,
    66: 3,
    67: 1,
    68: 1,
    69: 1,
    70: 1,
    71: 1,
    72: 1,
    73: 1,
    74: 0,
    75: 1,
    76: 1,
    77: 1,
    79: 2,
    80: 2,
    83: 1,
    85: 0,
    86: 1,
    89: 0,
    90: 1,
    94: 0,
    95: 0,
    97: 2,
    98: 1,
    99: 1,
    101: 3,
    102: 3,
    105: 1,
    106: 1,
    108: 2,
    111: 1,
    112: 1,
    113: 1,
    114: 1,
    117: 2,
    118: 1,
    119: 4,
    121: 1,
    126: 1,
    127: 1,
    128: 1,
    129: 1,
    130: 1,
    131: 1,
    133: 1,
    134: 1,
    135: 1,
    136: 2,
    137: 2,
    138: 2,
    140: 1,
    141: 1,
    142: 3,
    143: 4,
    144: 4,
    161: 1,
    162: 1,
    163: 1,
    164: 1,
    165: 2,
    172: 1,
    175: 2,
    176: 2,
    177: 3,
    178: 2,
    179: 1,
    184: 1,
    186: 1,
    189: 3,
    190: 1,
    195: 3,
    196: 3,
    197: 1,
    198: 1,
    199: 3,
    201: 1,
    207: 4,
    210: 3,
    211: 1,
    212: 2,
    213: 2,
    214: 1,
    215: 1,
    225: 0,
    229: 1,
    230: 1,
    231: 1,
    232: 1,
    233: 1,
    234: 1,
    235: 3,
    244: 1,
    247: 4,
    252: 2,
    257: 1,
    261: 1,
    271: 1,
    273: 4,
    274: 2,
    275: 2,
    276: 2,
    277: 3,
    278: 3,
    279: 1,
    280: 3,
    281: 3,
    282: 3,
    283: 1,
    284: 1,
    285: 2,
    286: 4,
    287: 3,
    288: 2,
    289: 4,
    290: 3,
    291: 3,
    292: 3,
    293: 4,
    294: 1,
    295: 3,
    296: 1,
    297: 3,
    298: 1,
    299: 2,
    300: 3,
    301: 3,
    302: 4,
    303: 2,
    304: 2,
    305: 2,
    306: 2,
    307: 2,
    308: 2,
    309: 3,
    310: 2,
    311: 2,
    312: 2,
    313: 2,
    314: 2,
    315: 2,
    316: 4,
    325: 2,
    326: 2,
    327: 2,
    328: 2,
    331: 2,
    332: 2,
    337: 2,
    342: 1,
    343: 1,
    346: 2,
    347: 1,
    350: 4,
    351: 3,
    352: 1,
    353: 2,
    360: 1,
    368: 1,
    369: 1,
    370: 1,
    371: 1,
    372: 1,
    373: 1,
    374: 1,
    375: 1,
    376: 1,
    377: 1,
    378: 1,
    382: 3,
    385: 1,
    392: 1,
    393: 1,
    396: 2,
    397: 2,
    398: 2,
    399: 1,
    400: 1,
    401: 1,
    402: 1,
    403: 1,
    404: 1,
    405: 1,
    406: 1,
    407: 1,
    408: 1,
    409: 1,
    410: 1,
    414: 4,
    415: 1,
    416: 1,
    417: 2,
    420: 1,
    421: 1,
    422: 2,
    424: 1,
    425: 2,
    426: 2,
    427: 2,
    428: 2,
    430: 3,
    438: 3,
    439: 3,
    440: 3,
    443: 2,
    444: 2,
    445: 2,
    446: 2,
    447: 6,
    448: 6,
    449: 2,
    450: 2,
    464: 2,
    468: 3,
    476: 2,
    479: 1,
    480: 2,
    65535: 0,
}

FTAB_ARGC.setdefault(221, 0)  # TODAY: bảng SheetJS thiếu mục này
ERR_CODES = {0x00: '#NULL!', 0x07: '#DIV/0!', 0x0F: '#VALUE!', 0x17: '#REF!',
             0x1D: '#NAME?', 0x24: '#NUM!', 0x2A: '#N/A'}
BUILTIN_NAMES = {0: 'Consolidate_Area', 1: 'Auto_Open', 2: 'Auto_Close', 3: 'Extract',
                 4: 'Database', 5: 'Criteria', 6: 'Print_Area', 7: 'Print_Titles',
                 8: 'Recorder', 9: 'Data_Form', 10: 'Auto_Activate',
                 11: 'Auto_Deactivate', 12: 'Sheet_Title', 13: '_FilterDatabase'}
BINOPS = {0x03: '+', 0x04: '-', 0x05: '*', 0x06: '/', 0x07: '^', 0x08: '&', 0x09: '<',
          0x0A: '<=', 0x0B: '=', 0x0C: '>=', 0x0D: '>', 0x0E: '<>', 0x0F: ' ',
          0x10: ',', 0x11: ':'}
R_BOF, R_EOF, R_CONTINUE = 0x809, 0x0A, 0x3C
MERGE_CONT = {0x18, 0x06, 0x4BC, 0x221, 0x23, 0x1AE, 0x17, 0x207}


class DecodeError(Exception):
    pass


# --------------------------------------------------------------- helpers
def col_letters(c):
    s = ''
    c += 1
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def cell_a1(r, c):
    return '%s%d' % (col_letters(c), r + 1)


def read_ustr(b, pos, cch):
    """Đọc chuỗi (flag byte + ký tự). Trả (str, vị trí mới)."""
    flag = b[pos]
    pos += 1
    if flag & 0x08:
        pos += 2
    if flag & 0x04:
        pos += 4
    if flag & 0x01:
        s = b[pos:pos + 2 * cch].decode('utf-16le', 'replace')
        pos += 2 * cch
    else:
        s = b[pos:pos + cch].decode('latin-1')
        pos += cch
    return s, pos


def jsnum(x):
    """Số dạng JS (rút gọn tối thiểu, không '.0')."""
    if x != x or x in (float('inf'), float('-inf')):
        return '#NUM!'
    if x == int(x) and abs(x) < 1e21:
        return str(int(x))
    s = repr(x)
    if 'e' in s or 'E' in s:
        from decimal import Decimal
        d = Decimal(s)
        if -7 < d.adjusted() < 21:
            s = format(d, 'f')
    return s


def quote_sheet(n):
    if re.fullmatch(r"[^\W\d][\w.]*", n) and not re.fullmatch(r'[A-Za-z]{1,3}\d+', n) \
            and not re.fullmatch(r'R\d*C\d*', n, re.I):
        return n
    return "'" + n.replace("'", "''") + "'"


def fmt_col(c, rel):
    return ('' if rel else '$') + col_letters(c)


def fmt_row(r, rel):
    return ('' if rel else '$') + str(r + 1)


def fmt_ref(r, c, rrel, crel):
    return fmt_col(c, crel) + fmt_row(r, rrel)


def fmt_area(r1, r2, c1, c2, rr1, rr2, cr1, cr2):
    if r1 == 0 and r2 == 0xFFFF:
        return '%s:%s' % (fmt_col(c1, cr1), fmt_col(c2, cr2))
    if c1 == 0 and c2 == 0xFF:
        return '%s:%s' % (fmt_row(r1, rr1), fmt_row(r2, rr2))
    return '%s:%s' % (fmt_ref(r1, c1, rr1, cr1), fmt_ref(r2, c2, rr2, cr2))


def s8(x):
    return x - 256 if x >= 128 else x


def s16(x):
    return x - 65536 if x >= 32768 else x


# --------------------------------------------------------------- workbook context
class Ctx:
    def __init__(self):
        self.sheets = []        # dict(name, state, offset, dt)
        self.supbooks = []      # dict(kind, path, sheets, names)
        self.xti = []           # (isup, first, last)
        self.names = []         # dict(name, builtin, hidden, itab, rgce, rgcb)

    def sheetname_of(self, isup, itab):
        if itab in (0xFFFE, 0xFFFF):
            return '#REF!'
        sb = self.supbooks[isup] if isup < len(self.supbooks) else None
        if sb is None:
            return '#REF!'
        if sb['kind'] == 'internal':
            return self.sheets[itab]['name'] if itab < len(self.sheets) else '#REF!'
        if sb['kind'] == 'external':
            return sb['sheets'][itab] if itab < len(sb['sheets']) else '#REF!'
        return '#REF!'


# --------------------------------------------------------------- ptg decoder
def decode_rgce(rgce, rgcb, base_r, base_c, ctx, shared=False):
    """Giải rgce thành chuỗi A1. Trả (chuỗi, [ext_refs])."""
    st = []
    ext = []
    lost = []
    tidx = -1
    i = 0
    n = len(rgce)
    pend_pre = ''
    pend_post = ''
    cb = 0

    def push(s):
        nonlocal pend_pre
        if pend_pre:
            s = pend_pre + s
            pend_pre = ''
        st.append(s)

    def loc(row, colf, relmode):
        rrel = bool(colf & 0x8000)
        crel = bool(colf & 0x4000)
        c = colf & 0x00FF
        r = row
        if relmode:
            if rrel:
                r = (base_r + s16(row)) & 0xFFFF
            if crel:
                c = (base_c + s8(c)) & 0xFF
        return r, c, rrel, crel

    def sheetref(ixti):
        if ixti >= len(ctx.xti):
            raise DecodeError('ixti ngoài EXTERNSHEET: %d' % ixti)
        isup, first, last = ctx.xti[ixti]
        if isup >= len(ctx.supbooks):
            raise DecodeError('SUPBOOK %d không có' % isup)
        sb = ctx.supbooks[isup]
        f = ctx.sheetname_of(isup, first)
        la = ctx.sheetname_of(isup, last)
        if sb['kind'] == 'internal' and f == '#REF!':
            return None, isup, sb
        if first == last:
            txt = quote_sheet(f)
        else:
            txt = quote_sheet(f + ':' + la) if not re.fullmatch(r"[^\W\d][\w.:]*", f + ':' + la) else f + ':' + la
        pre = '[ext:%d]' % isup if sb['kind'] == 'external' else ''
        return pre + txt, isup, sb

    def add3d(ixti, body, ptg=None, pos=None):
        sp, isup, sb = sheetref(ixti)
        tok = (sp + '!' + body) if sp is not None else '#REF!'
        if sp is None and ptg is not None:
            # sheet đã xóa (itab 0xFFFE/0xFFFF): chuỗi ghi #REF! như Excel, giữ địa chỉ trong byte
            ent = {'token_index': tidx, 'ptg': '0x%02X' % ptg, 'row': pos[0], 'col': pos[1], 'ref_a1': body}
            if len(pos) > 2:
                ent['row_last'] = pos[2]
                ent['col_last'] = pos[3]
            lost.append(ent)
        if sb['kind'] == 'external':
            ext.append({'supbook': isup, 'token': tok})
        push(tok)

    while i < n:
        t = rgce[i]
        tidx += 1
        i += 1
        if t in BINOPS:
            b = st.pop()
            a = st.pop()
            push(a + BINOPS[t] + b)
            continue
        if t == 0x12:
            push('+' + st.pop())
        elif t == 0x13:
            push('-' + st.pop())
        elif t == 0x14:
            push(st.pop() + '%')
        elif t == 0x15:
            x = st.pop() + pend_post
            pend_post = ''
            push('(' + x + ')')
        elif t == 0x16:
            st.append('')
        elif t == 0x17:
            cch = rgce[i]
            flag = rgce[i + 1]
            i += 2
            if flag & 1:
                s = rgce[i:i + 2 * cch].decode('utf-16le', 'replace')
                i += 2 * cch
            else:
                s = rgce[i:i + cch].decode('latin-1')
                i += cch
            push('"' + s.replace('"', '""') + '"')
        elif t == 0x19:
            g = rgce[i]
            w = struct.unpack_from('<H', rgce, i + 1)[0]
            i += 3
            if g & 0x04:
                i += 2 * (w + 1)
            elif g & 0x40:
                typ = w & 0xFF
                cnt = (w >> 8) & 0xFF
                sp = (' ' if typ in (0, 2, 4, 6) else '\r') * cnt
                if typ in (0, 1, 2, 3):
                    pend_pre += sp
                elif typ in (4, 5):
                    pend_post += sp
            elif g & 0x10:
                push('SUM(' + st.pop() + ')')
            # attrIf/Goto/Volatile/Baxcel: không ảnh hưởng chuỗi
        elif t == 0x1C:
            push(ERR_CODES.get(rgce[i], '#ERR%d!' % rgce[i]))
            i += 1
        elif t == 0x1D:
            push('TRUE' if rgce[i] else 'FALSE')
            i += 1
        elif t == 0x1E:
            push(str(struct.unpack_from('<H', rgce, i)[0]))
            i += 2
        elif t == 0x1F:
            push(jsnum(struct.unpack_from('<d', rgce, i)[0]))
            i += 8
        elif t >= 0x20:
            base = (t & 0x1F) | 0x20
            if base == 0x20:    # array
                i += 7
                cols = rgcb[cb] + 1
                rows = struct.unpack_from('<H', rgcb, cb + 1)[0] + 1
                cb += 3
                vals = []
                for _r in range(rows):
                    row = []
                    for _c in range(cols):
                        ty = rgcb[cb]
                        cb += 1
                        if ty == 0:
                            row.append('')
                            cb += 8
                        elif ty == 1:
                            row.append(jsnum(struct.unpack_from('<d', rgcb, cb)[0]))
                            cb += 8
                        elif ty == 2:
                            cch = struct.unpack_from('<H', rgcb, cb)[0]
                            s, cb = read_ustr(rgcb, cb + 2, cch)
                            row.append('"' + s.replace('"', '""') + '"')
                        elif ty == 4:
                            row.append('TRUE' if rgcb[cb] else 'FALSE')
                            cb += 8
                        elif ty == 0x10:
                            row.append(ERR_CODES.get(rgcb[cb], '#ERR!'))
                            cb += 8
                        else:
                            raise DecodeError('array const type %d' % ty)
                    vals.append(','.join(row))
                push('{' + ';'.join(vals) + '}')
            elif base == 0x21:   # func
                ift = struct.unpack_from('<H', rgce, i)[0]
                i += 2
                if ift not in FTAB_ARGC or ift not in FTAB:
                    raise DecodeError('hàm iftab %d thiếu bảng' % ift)
                k = FTAB_ARGC[ift]
                args = st[len(st) - k:] if k else []
                if k:
                    del st[len(st) - k:]
                push(FTAB[ift] + '(' + ','.join(args) + ')')
            elif base == 0x22:   # funcvar
                k = rgce[i] & 0x7F
                ift = struct.unpack_from('<H', rgce, i + 1)[0] & 0x7FFF
                i += 3
                args = st[len(st) - k:] if k else []
                if k:
                    del st[len(st) - k:]
                if ift == 255:
                    fname = args[0] if args else '?'
                    args = args[1:]
                else:
                    if ift not in FTAB:
                        raise DecodeError('hàm iftab %d thiếu tên' % ift)
                    fname = FTAB[ift]
                push(fname + '(' + ','.join(args) + ')')
            elif base == 0x23:   # name
                idx = struct.unpack_from('<H', rgce, i)[0]
                i += 4
                if idx < 1 or idx > len(ctx.names):
                    raise DecodeError('name idx %d' % idx)
                push(ctx.names[idx - 1]['name'])
            elif base == 0x24:   # ref
                row, colf = struct.unpack_from('<HH', rgce, i)
                i += 4
                r, c, rr, cr = loc(row, colf, False)
                push(fmt_ref(r, c, rr, cr))
            elif base == 0x25:   # area
                r1, r2, c1f, c2f = struct.unpack_from('<HHHH', rgce, i)
                i += 8
                a = loc(r1, c1f, False)
                b = loc(r2, c2f, False)
                push(fmt_area(a[0], b[0], a[1], b[1], a[2], b[2], a[3], b[3]))
            elif base in (0x26, 0x27, 0x28):   # memArea/memErr/memNoMem
                i += 6
                if base == 0x26:
                    cnt = struct.unpack_from('<H', rgcb, cb)[0]
                    cb += 2 + 8 * cnt
            elif base == 0x29:   # memFunc
                i += 2
            elif base == 0x2A:
                i += 4
                push('#REF!')
            elif base == 0x2B:
                i += 8
                push('#REF!')
            elif base == 0x2C:   # refN
                row, colf = struct.unpack_from('<HH', rgce, i)
                i += 4
                r, c, rr, cr = loc(row, colf, True)
                push(fmt_ref(r, c, rr, cr))
            elif base == 0x2D:   # areaN
                r1, r2, c1f, c2f = struct.unpack_from('<HHHH', rgce, i)
                i += 8
                a = loc(r1, c1f, True)
                b = loc(r2, c2f, True)
                push(fmt_area(a[0], b[0], a[1], b[1], a[2], b[2], a[3], b[3]))
            elif base == 0x39:   # nameX
                ixti, idx = struct.unpack_from('<HH', rgce, i)
                i += 6
                if ixti >= len(ctx.xti):
                    raise DecodeError('nameX ixti %d' % ixti)
                isup = ctx.xti[ixti][0]
                sb = ctx.supbooks[isup]
                nm = sb['names'][idx - 1] if 0 < idx <= len(sb['names']) else '?'
                if sb['kind'] == 'external':
                    tok = '[ext:%d]%s' % (isup, nm)
                    ext.append({'supbook': isup, 'token': tok})
                    push(tok)
                else:
                    push(nm)
            elif base == 0x3A:   # ref3d
                ixti, row, colf = struct.unpack_from('<HHH', rgce, i)
                i += 6
                r, c, rr, cr = loc(row, colf, shared)
                add3d(ixti, fmt_ref(r, c, rr, cr), 0x3A, (r, c))
            elif base == 0x3B:   # area3d
                ixti, r1, r2, c1f, c2f = struct.unpack_from('<HHHHH', rgce, i)
                i += 10
                a = loc(r1, c1f, shared)
                b = loc(r2, c2f, shared)
                add3d(ixti, fmt_area(a[0], b[0], a[1], b[1], a[2], b[2], a[3], b[3]), 0x3B, (a[0], a[1], b[0], b[1]))
            elif base == 0x3C:
                ixti = struct.unpack_from('<H', rgce, i)[0]
                i += 6
                add3d(ixti, '#REF!')
            elif base == 0x3D:
                ixti = struct.unpack_from('<H', rgce, i)[0]
                i += 10
                add3d(ixti, '#REF!')
            else:
                raise DecodeError('ptg 0x%02X' % t)
        else:
            raise DecodeError('ptg 0x%02X' % t)
    if len(st) != 1:
        raise DecodeError('stack cuối %d' % len(st))
    return st[0] + pend_pre, ext, lost


# --------------------------------------------------------------- BIFF records
def iter_records(d):
    """Trả danh sách (type, data, offset) đã gộp CONTINUE cho loại cần thiết."""
    out = []
    pos = 0
    n = len(d)
    while pos + 4 <= n:
        t, l = struct.unpack_from('<HH', d, pos)
        b = d[pos + 4:pos + 4 + l]
        if t == R_CONTINUE and out and out[-1][0] in MERGE_CONT:
            out[-1] = (out[-1][0], out[-1][1] + b, out[-1][2])
        else:
            out.append((t, b, pos))
        pos += 4 + l
    return out


def parse_supbook(b):
    ctab = struct.unpack_from('<H', b, 0)[0]
    if len(b) == 4 and b[2:4] == b'\x01\x04':
        return {'kind': 'internal', 'path': '', 'sheets': [], 'names': [], 'ctab': ctab}
    if len(b) == 4 and b[2:4] == b'\x01\x3a':
        return {'kind': 'addin', 'path': '', 'sheets': [], 'names': [], 'ctab': ctab}
    cch = struct.unpack_from('<H', b, 2)[0]
    path, p = read_ustr(b, 4, cch)
    sheets = []
    for _ in range(ctab):
        c2 = struct.unpack_from('<H', b, p)[0]
        s, p = read_ustr(b, p + 2, c2)
        sheets.append(s)
    return {'kind': 'external', 'path': path, 'sheets': sheets, 'names': [], 'ctab': ctab}


def pretty_path(p):
    out = []
    for ch in p:
        if ch == '\x01' or ch == '\x02' or ch == '\x04':
            continue
        out.append('/' if ch == '\x03' else ch)
    return ''.join(out)


def parse_workbook_stream(d):
    recs = iter_records(d)
    ctx = Ctx()
    depth = 0
    sheet_by_off = {}
    cur = None
    sub_order = 0
    globals_done = False
    cells = []     # (sheet_idx, type, data, row, col)
    per_sheet = {}  # idx -> list of (type, data)
    last_sup = None
    seq_sheet = -1
    for t, b, off in recs:
        if t == R_BOF:
            depth += 1
            if depth == 1:
                if not globals_done and sub_order == 0:
                    cur = 'globals'
                else:
                    cur = sheet_by_off.get(off)
                    if cur is None:
                        seq_sheet += 1
                        cur = seq_sheet if seq_sheet < len(ctx.sheets) else None
                    per_sheet.setdefault(cur, [])
                sub_order += 1
            continue
        if t == R_EOF:
            depth -= 1
            if depth == 0:
                if cur == 'globals':
                    globals_done = True
                cur = None
            continue
        if cur == 'globals' and depth == 1:
            if t == 0x85:   # BOUNDSHEET
                lb, grbit = struct.unpack_from('<IH', b, 0)
                cch = b[6]
                name, _ = read_ustr(b, 7, cch)
                idx = len(ctx.sheets)
                ctx.sheets.append({'name': name, 'state': grbit & 3, 'dt': (grbit >> 8) & 0xFF, 'offset': lb})
                sheet_by_off[lb] = idx
            elif t == 0x1AE:
                ctx.supbooks.append(parse_supbook(b))
                last_sup = ctx.supbooks[-1]
            elif t == 0x23 and last_sup is not None:
                cch = b[6]
                nm, _ = read_ustr(b, 7, cch)
                last_sup['names'].append(nm)
            elif t == 0x17:
                cx = struct.unpack_from('<H', b, 0)[0]
                ctx.xti = [struct.unpack_from('<Hhh', b, 2 + 6 * k) for k in range(cx)]
                ctx.xti = [(a, f & 0xFFFF, l & 0xFFFF) for a, f, l in ctx.xti]
            elif t == 0x18:
                grbit, chkey, cch, cce, ixals, itab = struct.unpack_from('<HBBHHH', b, 0)
                ccm, cdesc, chlp, cst = b[10], b[11], b[12], b[13]
                builtin = bool(grbit & 0x20)
                s, p = read_ustr(b, 14, cch)
                if builtin:
                    nm = BUILTIN_NAMES.get(ord(s[0]) if s else -1, 'Builtin%d' % (ord(s[0]) if s else -1))
                else:
                    nm = s
                ctx.names.append({'name': nm, 'builtin': builtin, 'hidden': bool(grbit & 1),
                                  'itab': itab, 'rgce': b[p:p + cce], 'rgcb': b[p + cce:], 'grbit': grbit})
        elif isinstance(cur, int) and depth == 1:
            per_sheet.setdefault(cur, []).append((t, b))
    return ctx, per_sheet


def process_biff(path, data):
    """Trả dict kết quả cho 1 tệp BIFF; ném lỗi nếu không phải workbook."""
    ole = olefile.OleFileIO(data)
    sname = 'Workbook' if ole.exists('Workbook') else ('Book' if ole.exists('Book') else None)
    if sname is None:
        ole.close()
        raise ValueError('không có luồng Workbook/Book')
    d = ole.openstream(sname).read()
    ole.close()
    ctx, per_sheet = parse_workbook_stream(d)
    errors = []
    formulas = []
    stats = {'normal': 0, 'shared': 0, 'array': 0, 'shrfmla_records': 0, 'array_records': 0}
    for si, recs in sorted(per_sheet.items(), key=lambda kv: (kv[0] is None, kv[0] if kv[0] is not None else 0)):
        if si is None:
            continue
        sh = ctx.sheets[si]
        shr = {}     # anchor(r,c) -> (ref, rgce, rgcb)
        arr = {}
        fcells = []  # (r,c,rgce,rgcb)
        prev = None
        for t, b in recs:
            if t == 0x06:
                r, c = struct.unpack_from('<HH', b, 0)
                cce = struct.unpack_from('<H', b, 20)[0]
                rg = b[22:22 + cce]
                rcb = b[22 + cce:]
                fcells.append((r, c, rg, rcb))
                prev = (r, c)
            elif t == 0x4BC:
                r1, r2, c1, c2 = struct.unpack_from('<HHBB', b, 0)
                cce = struct.unpack_from('<H', b, 8)[0]
                key = prev if prev is not None else (r1, c1)
                shr[key] = ('%s:%s' % (cell_a1(r1, c1), cell_a1(r2, c2)), b[10:10 + cce], b[10 + cce:])
                stats['shrfmla_records'] += 1
            elif t == 0x221:
                r1, r2, c1, c2 = struct.unpack_from('<HHBB', b, 0)
                cce = struct.unpack_from('<H', b, 12)[0]
                arr[(r1, c1)] = ('%s:%s' % (cell_a1(r1, c1), cell_a1(r2, c2)), b[14:14 + cce], b[14 + cce:], (r1, r2, c1, c2))
                stats['array_records'] += 1
        # nếu SHRFMLA gắn nhầm anchor, lập thêm chỉ mục theo góc trái-trên vùng
        shr_by_corner = {}
        for k, v in shr.items():
            m = re.match(r'([A-Z]+)(\d+):', v[0])
            shr_by_corner.setdefault((int(m.group(2)) - 1, m.group(1)), v)
        for r, c, rg, rcb in fcells:
            kind = 'normal'
            ref = None
            try:
                if len(rg) == 5 and rg[0] == 0x01:
                    er, ec = struct.unpack_from('<HH', rg, 1)
                    if (er, ec) in arr:
                        ref, arg, acb, _rect = arr[(er, ec)]
                        kind = 'array'
                        f, ex, lo = decode_rgce(arg, acb, _rect[0], _rect[2], ctx, False)
                    elif (er, ec) in shr:
                        ref, arg, acb = shr[(er, ec)]
                        kind = 'shared'
                        f, ex, lo = decode_rgce(arg, acb, r, c, ctx, True)
                    else:
                        raise DecodeError('ptgExp (%d,%d) không có SHRFMLA/ARRAY' % (er, ec))
                else:
                    f, ex, lo = decode_rgce(rg, rcb, r, c, ctx, False)
            except (DecodeError, IndexError, struct.error) as e:
                errors.append({'sheet': sh['name'], 'cell': cell_a1(r, c), 'error': str(e)})
                f, ex, lo, kind = None, [], [], kind
            stats[kind] += 1
            ent = {'sheet': sh['name'], 'hidden': sh['state'] != 0, 'cell': cell_a1(r, c),
                   'formula': f, 'kind': kind, 'ext_refs': ex}
            if lo:
                ent['lost_refs'] = lo
            if ref is not None:
                ent['shared_ref' if kind == 'shared' else 'array_ref'] = ref
            formulas.append(ent)
    names = []
    for ni, nm in enumerate(ctx.names):
        try:
            f, ex, lo = decode_rgce(nm['rgce'], nm['rgcb'], 0, 0, ctx, False) if nm['rgce'] else ('', [], [])
        except (DecodeError, IndexError, struct.error) as e:
            errors.append({'sheet': None, 'cell': 'name:' + nm['name'], 'error': str(e)})
            f, ex, lo = None, [], []
        scope = 'workbook' if nm['itab'] == 0 else (ctx.sheets[nm['itab'] - 1]['name'] if nm['itab'] - 1 < len(ctx.sheets) else '?')
        names.append({'index': ni + 1, 'name': nm['name'], 'builtin': nm['builtin'], 'hidden': nm['hidden'],
                      'scope': scope, 'formula': f, 'ext_refs': ex, **({'lost_refs': lo} if lo else {})})
    supbooks = [{'index': i, 'kind': s['kind'], 'path': pretty_path(s['path']), 'path_raw': s['path'],
                 'sheets': s['sheets'], 'names': s['names']} for i, s in enumerate(ctx.supbooks)]
    xti = [{'index': i, 'supbook': a, 'sheet_first': f, 'sheet_last': l} for i, (a, f, l) in enumerate(ctx.xti)]
    sheets = [{'index': i, 'name': s['name'], 'state': ['visible', 'hidden', 'veryHidden', '?'][s['state']],
               'type': s['dt']} for i, s in enumerate(ctx.sheets)]
    return {'format': 'biff8', 'sheets': sheets, 'supbooks': supbooks, 'externsheet': xti,
            'defined_names': names, 'formulas': formulas, 'errors': errors, 'stats': stats}


# --------------------------------------------------------------- OOXML (.xlsx)
_STR_OR_QUOTE = re.compile(r'("(?:[^"]|"")*"|\'(?:[^\']|\'\')*\')')
_REF_RE = re.compile(r'(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\w(])')


def col_to_idx(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n - 1


def shift_formula(f, dr, dc):
    parts = _STR_OR_QUOTE.split(f)
    for k in range(0, len(parts), 2):
        def rep(m):
            cabs, col, rabs, row = m.groups()
            c = col_to_idx(col)
            r = int(row)
            if not cabs:
                c += dc
            if not rabs:
                r += dr
            return '%s%s%s%d' % ('$' if cabs else '', col_letters(c), '$' if rabs else '', r)
        parts[k] = _REF_RE.sub(rep, parts[k])
    return ''.join(parts)


def split_a1(a1):
    m = re.fullmatch(r'([A-Z]+)(\d+)', a1)
    return int(m.group(2)) - 1, col_to_idx(m.group(1))


def process_xlsx(path, data):
    import io
    z = zipfile.ZipFile(io.BytesIO(data))
    ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rmap = {r.get('Id'): r.get('Target') for r in rels}
    sheets = []
    for s in wb.find('m:sheets', ns):
        tgt = rmap[s.get('{%s}id' % ns['r'])]
        tgt = tgt.lstrip('/')
        if not tgt.startswith('xl/'):
            tgt = 'xl/' + tgt
        sheets.append({'index': len(sheets), 'name': s.get('name'), 'state': s.get('state', 'visible'),
                       'type': 0, 'part': tgt})
    names = []
    dn = wb.find('m:definedNames', ns)
    if dn is not None:
        for k, d in enumerate(dn):
            nm = d.get('name')
            ls = d.get('localSheetId')
            scope = 'workbook' if ls is None else sheets[int(ls)]['name']
            builtin = nm.startswith('_xlnm.')
            names.append({'index': k + 1, 'name': nm[6:] if builtin else nm, 'builtin': builtin,
                          'hidden': d.get('hidden') == '1', 'scope': scope, 'formula': d.text or '', 'ext_refs': []})
    formulas = []
    stats = {'normal': 0, 'shared': 0, 'array': 0, 'shrfmla_records': 0, 'array_records': 0}
    errors = []
    mq = '{%s}' % ns['m']
    for sh in sheets:
        root = ET.fromstring(z.read(sh['part']))
        masters = {}
        for c in root.iter(mq + 'c'):
            f = c.find(mq + 'f')
            if f is None:
                continue
            a1 = c.get('r')
            t = f.get('t')
            txt = f.text or ''
            kind = 'normal'
            ref = None
            if t == 'shared':
                si = f.get('si')
                kind = 'shared'
                if txt:
                    masters[si] = (a1, txt)
                    ref = f.get('ref')
                    stats['shrfmla_records'] += 1
                else:
                    if si not in masters:
                        errors.append({'sheet': sh['name'], 'cell': a1, 'error': 'shared master thiếu'})
                        continue
                    ma, mt = masters[si]
                    r0, c0 = split_a1(ma)
                    r1, c1 = split_a1(a1)
                    txt = shift_formula(mt, r1 - r0, c1 - c0)
            elif t == 'array':
                kind = 'array'
                ref = f.get('ref')
                stats['array_records'] += 1
            ex = []
            for m in re.finditer(r'\[(\d+)\]', _STR_OR_QUOTE.sub('', txt)):
                ex.append({'supbook': int(m.group(1)), 'token': m.group(0)})
            stats[kind] += 1
            ent = {'sheet': sh['name'], 'hidden': sh['state'] != 'visible', 'cell': a1,
                   'formula': txt, 'kind': kind, 'ext_refs': ex}
            if ref and kind != 'normal':
                ent['shared_ref' if kind == 'shared' else 'array_ref'] = ref
            formulas.append(ent)
    for sh in sheets:
        sh.pop('part')
    return {'format': 'xlsx', 'sheets': sheets, 'supbooks': [], 'externsheet': [],
            'defined_names': names, 'formulas': formulas, 'errors': errors, 'stats': stats}


# --------------------------------------------------------------- khóa theo menu.xml
ALIASES = {}   # FolderName (Validity) khác MaFile -> MaFile


def load_menu(path):
    mf = {}
    root = ET.parse(path).getroot()
    for m in root.iter('Menu'):
        k = m.get('MaFile')
        if k:
            mf.setdefault(k, set()).add(m.get('VersionXML', ''))
            for v in m.iter('Validity'):
                fn = v.get('FolderName')
                if fn and fn != k:
                    ALIASES[fn] = k
    return mf


class KeyMapper:
    FORM_RE = re.compile(r'^_?(?P<stem>.+?)(?:_xml)?_(?P<ver>\d{2,5})$')

    def __init__(self, menu):
        self.stem_keys = {}
        for mf in menu:
            k = mf.lower()
            self.stem_keys.setdefault(k, mf)
            if len(k) > 1 and k[0] == '0' and k[1].isdigit():
                self.stem_keys.setdefault(k[1:], mf)
            if mf.startswith('200_'):
                self.stem_keys.setdefault('00_' + k[4:], mf)
        for fn, mf in ALIASES.items():
            k = fn.lower()
            self.stem_keys.setdefault(k, mf)
            if len(k) > 1 and k[0] == '0' and k[1].isdigit():
                self.stem_keys.setdefault(k[1:], mf)
        self.menu = menu

    def vers_digits(self, mf):
        return {v.replace('.', ''): v for v in self.menu[mf] if v}

    def resolve(self, stem, ver):
        if stem in self.stem_keys:
            return self.stem_keys[stem]
        c = sorted(set(v for k, v in self.stem_keys.items() if k and k[0].isdigit() and k[1:] == stem))
        if len(c) == 1:
            return c[0]
        if len(c) > 1:
            c2 = [m for m in c if ver in self.vers_digits(m)]
            if len(c2) == 1:
                return c2[0]
        return None

    def map(self, name):
        disp = name.rstrip('.')
        disp = os.path.splitext(disp)[0].rstrip('_') + os.path.splitext(disp)[1]
        stem = os.path.splitext(disp)[0].lower()
        stem = re.sub(r'^(bang_ke|bangke|khbs|pl)_', '', stem)
        m = self.FORM_RE.match(stem)
        if not m:
            return None
        mf = self.resolve(m.group('stem'), m.group('ver'))
        if mf is None:
            return None
        vd = self.vers_digits(mf)
        ver = m.group('ver')
        if ver in vd:
            return mf, vd[ver], 'menu'
        if len(ver) in (2, 3):   # phiên bản suy từ hậu tố số: 251 -> 2.5.1, 23 -> 2.3
            return mf, '.'.join(ver), 'ten_tep'
        return None


def safe_name(s):
    return re.sub(r'[^A-Za-z0-9._-]', '_', s.rstrip('.'))


# --------------------------------------------------------------- ghi JSON
def dumps(o):
    return json.dumps(o, ensure_ascii=False, separators=(',', ':'))


def write_result(outdir, key, header, res):
    p = os.path.join(outdir, key + '.json')
    parts = ['{\n"header":' + dumps(header),
             '"key":' + dumps(key),
             '"key_origin":' + dumps(res['key_origin']),
             '"format":' + dumps(res['format']),
             '"stats":' + dumps(res['stats']),
             '"errors":' + dumps(res['errors']),
             '"sheets":' + dumps(res['sheets']),
             '"supbooks":' + dumps(res['supbooks']),
             '"externsheet":' + dumps(res['externsheet']),
             '"defined_names":' + dumps(res['defined_names'])]
    body = ',\n'.join(parts) + ',\n"formulas":[\n' + ',\n'.join(dumps(e) for e in res['formulas']) + '\n]\n}\n'
    with open(p, 'w', encoding='utf-8', errors='backslashreplace', newline='\n') as f:
        f.write(body)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--src', required=True)
    ap.add_argument('--menu', required=True)
    ap.add_argument('--out', required=True)
    a = ap.parse_args()
    menu = load_menu(a.menu)
    km = KeyMapper(menu)
    os.makedirs(a.out, exist_ok=True)
    index = []
    used = {}
    for name in sorted(os.listdir(a.src)):
        sp = os.path.join(a.src, name)
        if not os.path.isfile(sp) or name == 'menu.xml':
            continue
        with open(sp, 'rb') as f:
            data = f.read()
        magic = data[:4]
        if magic == b'\xd0\xcf\x11\xe0':
            kind = 'ole'
        elif magic[:2] == b'PK':
            kind = 'zip'
        else:
            continue
        sha = hashlib.sha256(data).hexdigest()
        mp = km.map(name)
        key = ('%s@v%s@%s' % (mp[0], mp[1], safe_name(name))) if mp else safe_name(name)
        origin = mp[2] if mp else 'ten_nguon'
        if key.lower() in used:
            raise SystemExit('trùng khóa: %s (%s, %s)' % (key, name, used[key.lower()]))
        used[key.lower()] = name
        entry = {'key': key, 'source_name': name, 'source_path': sp, 'source_sha256': sha,
                 'mafile': mp[0] if mp else None, 'version_xml': mp[1] if mp else None,
                 'key_origin': origin}
        try:
            res = process_xlsx(sp, data) if kind == 'zip' else process_biff(sp, data)
            res['key_origin'] = origin
        except Exception as e:
            entry.update({'status': 'skipped', 'reason': '%s: %s' % (type(e).__name__, e)})
            index.append(entry)
            continue
        header = {'generator': GENERATOR, 'version': VERSION, 'source_path': sp, 'source_sha256': sha}
        write_result(a.out, key, header, res)
        st = res['stats']
        entry.update({'status': 'ok', 'format': res['format'], 'json': key + '.json',
                      'formulas': len(res['formulas']),
                      'formulas_ext': sum(1 for e in res['formulas'] if e['ext_refs']),
                      'shared': st['shared'], 'array': st['array'],
                      'shrfmla_records': st['shrfmla_records'], 'array_records': st['array_records'],
                      'defined_names_user': sum(1 for n in res['defined_names'] if not n['builtin']),
                      'defined_names_builtin': sum(1 for n in res['defined_names'] if n['builtin']),
                      'supbooks_external': sum(1 for s in res['supbooks'] if s['kind'] == 'external'),
                      'errors': len(res['errors'])})
        index.append(entry)
    idx = {'generator': GENERATOR, 'version': VERSION, 'files': index,
           'totals': {'files_ok': sum(1 for e in index if e['status'] == 'ok'),
                      'files_skipped': sum(1 for e in index if e['status'] != 'ok'),
                      'formulas': sum(e.get('formulas', 0) for e in index),
                      'errors': sum(e.get('errors', 0) for e in index)}}
    with open(os.path.join(a.out, 'index.json'), 'w', encoding='utf-8', newline='\n') as f:
        json.dump(idx, f, ensure_ascii=False, indent=1)
        f.write('\n')
    print(json.dumps(idx['totals']))


if __name__ == '__main__':
    main()
