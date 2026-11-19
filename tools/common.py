"""Shared styling and reference data for the Treaty Performance Tracker generator.

Copied from the Aggregate Tracker generator (iwadaya/cresta tools/build_workbook.py) so both
tools carry the same Universe palette, countries, currencies, treaty types, classes and cedants.
"""
BRAND = "TREATY PERFORMANCE"
import sys
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, GradientFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule, DataBarRule, CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as L
from openpyxl.comments import Comment

# ---- Universe palette (shared/universeXlsxTheme.js) ---------------------------
TEAL, GREEN, HEADER, INK, INK_SOFT = "0D5C75", "2F8F6B", "1E7A5A", "16323B", "5B7079"
TOTAL_BG, TOTAL_INK, GRID, SUB_BG = "CDE7DD", "0D4A3A", "D8E2E6", "F4F9FA"
CALC_BG, INPUT_BLUE = "EEF3F5", "0000FF"
FONT = "Arial"
MONEY = '#,##0;[Red]-#,##0;"-"'
RATE = '#,##0.000000'
DATE = 'dd-mmm-yyyy'
PCT = '0.00%'
STAMP = 'dd-mmm-yyyy hh:mm'

FIRST, LAST_AGG, LAST_CED, LAST_CTRY, LAST_RATE, LAST_LIST = 6, 5005, 505, 205, 105, 105


thin = Side(style="thin", color=GRID)
BOX = Border(top=thin, left=thin, bottom=thin, right=thin)


def f(size=10, bold=False, color=INK, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(c):
    return PatternFill("solid", fgColor=c)


def title_bar(ws, ncols, title, subtitle):
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(1, 1, f"{BRAND}    {title}")
    c.font = Font(name=FONT, size=15, bold=True, color="FFFFFF")
    c.alignment = Alignment(vertical="center", indent=1)
    c.fill = GradientFill(degree=0, stop=(TEAL, GREEN))
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 14
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=ncols)
    s = ws.cell(3, 1, subtitle)
    s.font = f(9, italic=True, color=INK_SOFT)
    s.fill = fill(SUB_BG)
    s.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[3].height = 16


def header(ws, row, labels, start_col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row, start_col + i, lab)
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        c.fill = fill(HEADER)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[row].height = 30


def widths(ws, ws_widths):
    for col, w in ws_widths.items():
        ws.column_dimensions[col].width = w


def style_range(ws, ref, font=None, number_format=None, bg=None, align=None, border=True):
    for row in ws[ref]:
        for c in row:
            if font: c.font = font
            if number_format: c.number_format = number_format
            if bg: c.fill = fill(bg)
            if align: c.alignment = align
            if border: c.border = BOX


def name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


# ---- Reference data (from Universe startup/ensureReferenceData.js, plus extra
#      African markets and currencies the Universe seed does not carry) -----------
REGION = {}
for codes, reg in [
    ("AE SA KW BH OM QA JO LB", "Middle East"),
    ("EG MA TN DZ ZA NG KE ZM ZW GH TZ UG ET RW MZ BW NA MW SN CI CM AO CD MU", "Africa"),
    ("GB FR DE IT ES CH NL BE TR", "Europe"),
    ("IN PK LK BD MY SG JP CN KR TH ID PH", "Asia"),
    ("AU NZ", "Oceania"), ("US CA MX BR CL CO AR", "Americas")]:
    for c in codes.split():
        REGION[c] = reg

COUNTRIES = [
    ('United Arab Emirates','AE','ARE'),('Saudi Arabia','SA','SAU'),('Kuwait','KW','KWT'),('Bahrain','BH','BHR'),
    ('Oman','OM','OMN'),('Qatar','QA','QAT'),('Jordan','JO','JOR'),('Lebanon','LB','LBN'),('Egypt','EG','EGY'),
    ('Morocco','MA','MAR'),('Tunisia','TN','TUN'),('Algeria','DZ','DZA'),('South Africa','ZA','ZAF'),
    ('Nigeria','NG','NGA'),('Kenya','KE','KEN'),('United Kingdom','GB','GBR'),('France','FR','FRA'),
    ('Germany','DE','DEU'),('Italy','IT','ITA'),('Spain','ES','ESP'),('Switzerland','CH','CHE'),
    ('Netherlands','NL','NLD'),('Belgium','BE','BEL'),('Turkey','TR','TUR'),('India','IN','IND'),
    ('Pakistan','PK','PAK'),('Sri Lanka','LK','LKA'),('Bangladesh','BD','BGD'),('Malaysia','MY','MYS'),
    ('Singapore','SG','SGP'),('Japan','JP','JPN'),('China','CN','CHN'),('South Korea','KR','KOR'),
    ('Thailand','TH','THA'),('Indonesia','ID','IDN'),('Philippines','PH','PHL'),('Australia','AU','AUS'),
    ('New Zealand','NZ','NZL'),('United States','US','USA'),('Canada','CA','CAN'),('Mexico','MX','MEX'),
    ('Brazil','BR','BRA'),('Chile','CL','CHL'),('Colombia','CO','COL'),('Argentina','AR','ARG'),
    # additions for Africa coverage
    ('Zambia','ZM','ZMB'),('Zimbabwe','ZW','ZWE'),('Ghana','GH','GHA'),('Tanzania','TZ','TZA'),
    ('Uganda','UG','UGA'),('Ethiopia','ET','ETH'),('Rwanda','RW','RWA'),('Mozambique','MZ','MOZ'),
    ('Botswana','BW','BWA'),('Namibia','NA','NAM'),('Malawi','MW','MWI'),('Senegal','SN','SEN'),
    ("Cote d'Ivoire",'CI','CIV'),('Cameroon','CM','CMR'),('Angola','AO','AGO'),
    ('DR Congo','CD','COD'),('Mauritius','MU','MUS'),
]
COUNTRIES.sort(key=lambda r: r[0])

CURRENCIES = [
    ('USD','US Dollar'),('EUR','Euro'),('GBP','British Pound'),('AED','UAE Dirham'),
    ('SAR','Saudi Riyal'),('KWD','Kuwaiti Dinar'),('BHD','Bahraini Dinar'),
    ('OMR','Omani Rial'),('QAR','Qatari Riyal'),('JOD','Jordanian Dinar'),
    ('EGP','Egyptian Pound'),('MAD','Moroccan Dirham'),('ZAR','South African Rand'),
    ('NGN','Nigerian Naira'),('KES','Kenyan Shilling'),('INR','Indian Rupee'),
    ('PKR','Pakistani Rupee'),('LKR','Sri Lankan Rupee'),('MYR','Malaysian Ringgit'),
    ('SGD','Singapore Dollar'),('JPY','Japanese Yen'),('CNY','Chinese Yuan'),
    ('KRW','South Korean Won'),('THB','Thai Baht'),('IDR','Indonesian Rupiah'),
    ('AUD','Australian Dollar'),('NZD','New Zealand Dollar'),('CAD','Canadian Dollar'),
    ('CHF','Swiss Franc'),('TRY','Turkish Lira'),('BRL','Brazilian Real'),('MXN','Mexican Peso'),
    # additions for Africa coverage
    ('ZMW','Zambian Kwacha'),('ZWG','Zimbabwe Gold'),('GHS','Ghanaian Cedi'),('TZS','Tanzanian Shilling'),
    ('UGX','Ugandan Shilling'),('ETB','Ethiopian Birr'),('RWF','Rwandan Franc'),('MZN','Mozambican Metical'),
    ('BWP','Botswana Pula'),('NAD','Namibian Dollar'),('MWK','Malawian Kwacha'),('XOF','West African CFA Franc'),
    ('XAF','Central African CFA Franc'),('AOA','Angolan Kwanza'),('CDF','Congolese Franc'),('MUR','Mauritian Rupee'),
]
CURRENCIES.sort()
# Only USD and fixed USD pegs are pre-filled (USD per 1 unit). Every floating
# currency is left blank for the user to enter - no rate is guessed.
PEGS = {'USD': (1.0, 'Base currency'), 'SAR': (1/3.75, 'Peg 3.75 SAR/USD'),
        'AED': (1/3.6725, 'Peg 3.6725 AED/USD'), 'BHD': (1/0.376, 'Peg 0.376 BHD/USD'),
        'QAR': (1/3.64, 'Peg 3.64 QAR/USD'), 'JOD': (1/0.709, 'Peg 0.709 JOD/USD'),
        'OMR': (1/0.3845, 'Peg 0.3845 OMR/USD')}
PEG_DATE = date(2026, 1, 1)

TREATY_TYPES = [('ALL','UNSPECIFIED'),
    ('Quota Share','PROPORTIONAL'),('Quota Share & Surplus','PROPORTIONAL'),
    ('First Surplus','PROPORTIONAL'),('Second Surplus','PROPORTIONAL'),
    ('Third Surplus','PROPORTIONAL'),('Fac Oblig','PROPORTIONAL'),
    ('Risk XL','NON_PROPORTIONAL'),('CAT XL','NON_PROPORTIONAL'),
    ('Risk & CAT XL','NON_PROPORTIONAL'),('Stop Loss','NON_PROPORTIONAL'),
    ('Aggregate XL','NON_PROPORTIONAL')]
CLASSES = [('ALL','ALL'),
    ('Property','PROP'),('Motor','MOT'),('Marine','MAR'),('Engineering','ENG'),
    ('Liability','LIA'),('Medical','MED'),('Aviation','AVI'),('Energy','ENE'),
    ('Agriculture','AGR'),('Credit & Surety','CS'),('Miscellaneous','MISC'),
    ('Life','LIFE'),('Group Life','GL'),('Workers Compensation','WC')]
PERILS = [('EQ','Earthquake'),('WS','Windstorm / tropical cyclone'),('Flood','Flood (riverine, flash, surge)'),
          ('SRCC','Strikes, riots & civil commotion'),('Others','Any other peril tracked')]
MATRIX_COUNTRIES = ['Saudi Arabia','United Arab Emirates','Kuwait','Bahrain','Oman','Qatar','Jordan','Egypt',
                    'Morocco','South Africa','Nigeria','Kenya','Zambia','Zimbabwe','Ghana','Tanzania','Uganda',
                    'Ethiopia','Pakistan','India']

# Universe cedant seed (ensureReferenceData.js + seeds/002_reference_data.sql,
# merged; ADNIC / Oman Insurance / Kuwait Insurance dropped as duplicates).
# Must match vba/AggSeed.bas.
UNIVERSE_CEDANTS = [
    ('Tawuniya','Saudi Arabia'),('MEDGULF','Saudi Arabia'),('Al Rajhi Takaful','Saudi Arabia'),
    ('Bupa Arabia','Saudi Arabia'),('Walaa Insurance','Saudi Arabia'),('SALAMA','Saudi Arabia'),
    ('Malath Insurance','Saudi Arabia'),('Gulf Union Insurance','Saudi Arabia'),
    ('Orient Insurance','United Arab Emirates'),('Oman Insurance Co','United Arab Emirates'),
    ('Abu Dhabi National Insurance','United Arab Emirates'),('Sukoon Insurance','United Arab Emirates'),
    ('Dubai Insurance','United Arab Emirates'),('Salama Islamic Insurance','United Arab Emirates'),
    ('GIG Bahrain','Bahrain'),('Solidarity Bahrain','Bahrain'),
    ('Kuwait Insurance Co','Kuwait'),('Gulf Insurance Group','Kuwait'),('Warba Insurance','Kuwait'),
    ('Misr Insurance','Egypt'),('GIG Egypt','Egypt'),('Allianz Egypt','Egypt'),
    ('Dhofar Insurance','Oman'),('National Life & General','Oman'),
    ('Aviva','United Kingdom'),('AXA UK','United Kingdom'),('RSA Insurance','United Kingdom'),('Zurich UK','United Kingdom'),
]

