import openpyxl
import argparse

DEFAULT_2M_OFFSET = '600 khZ'
DEFAULT_70CM_OFFSET = '5.00 MHz'

# insert a column with a default value
def add_filled_column(sheet, col, name, value):
    sheet.insert_cols(idx=col, amount=1)
    sheet.cell(row=1, column=col).value = name
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=col, max_col=col):
        row[0].value = value

def populate_anytone(workbook, anytone_sheet_name='Anytone', source_sheet_name='Import'):
    "populate the specified sheet in anytone format from a sheet in ft70 format"

    source_sheet = workbook[source_sheet_name]
    anytone_sheet = workbook.create_sheet(anytone_sheet_name)
    
    # populate channel numbers with a counter
    anytone_sheet['A1'] = 'No.'
    rowctr = 1
    for row in anytone_sheet.iter_rows(min_row=2, max_row=source_sheet.max_row, min_col=1, max_col=1):
        row[0].value = rowctr
        rowctr += 1

    # now add name, rx freq, tx freq, channel type, tx pwr, bw, ctcss dec, ctcss enc, everything else defaulted

def translate_repeaterbook(workbook, sheet_name):
    "translate the named sheet from repeaterbook to ft70 format"
    sheet = workbook[sheet_name]

    # change the offset direction column to two columns: 'Offset Frequency', and 'Offset Direction'
    # 'Offset Direction' is Plus, Minus, Simplex instead of +, -, blank
    if sheet['D1'].value == 'Offset Direction':
        sheet.insert_cols(idx=4, amount=1)
        sheet['D1'].value = 'Offset Frequency'
        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=3, max_col=5):
            # transmit_freq, offset_freq, offset_dir
            if row[0].value < 200:
                offset_freq = DEFAULT_2M_OFFSET
            else:
                offset_freq = DEFAULT_70CM_OFFSET
            if row[2].value == '+':
                row[1].value = offset_freq
                row[2].value = 'Plus'
            elif row[2].value == '-':
                row[1].value = offset_freq
                row[2].value = 'Minus'
            else:
                row[2].value = 'Simplex'

        # now add two columns after 'Offset Direction'
        # 'Operating Mode', and AMS
        add_filled_column(sheet, 6, 'Operating Mode', 'FM')
        add_filled_column(sheet, 7, 'AMS', 'On')

        # add the 'Show Name' column and set all the 'On'
        add_filled_column(sheet, 9, 'Show Name', 'On')

        # Change CTCSS to an fp string with an Hz qualifier
        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=11, max_col=11):
            tfreq = float(row[0].value)
            tfreq_string = f"{tfreq:.1f} Hz"
            row[0].value = tfreq_string

        # delete the Rx CTCSS column
        sheet.delete_cols(idx=12)

        # delete the RX DCS column
        sheet.delete_cols(idx=13)

        # add remaining defaulted columns before comment field
        add_filled_column(sheet, 13, 'DCS Polarity', 'RN-TN')
        add_filled_column(sheet, 14, 'PR FREQ', '1600 Hz')
        add_filled_column(sheet, 15, 'Tx Power', 'High')
        add_filled_column(sheet, 16, 'Skip', 'Off')
        add_filled_column(sheet, 17, 'Step', 'Auto')
        add_filled_column(sheet, 18, 'Mask', 'Off')
        add_filled_column(sheet, 19, 'Attenuator', 'Off')
        add_filled_column(sheet, 20, 'S-Meter Squelch', 'Off')
        add_filled_column(sheet, 21, 'Bell', 'Off')
        add_filled_column(sheet, 22, 'Half Dev', 'Off')
        add_filled_column(sheet, 23, 'Clock Shift', 'Off')

        # insert the memory bank columns
        for i in range(1, 25):
            add_filled_column(sheet, i + 23, f'BANK {i}', 'Off')

    else:
        print(f"Unexpected value in D1: {sheet['D1'].value}, exiting without modifying")
        exit(1)


def main():
    parser = argparse.ArgumentParser(description='Convert code plugs between formats')
    parser.add_argument('-i', '--input', help='Input file', required=True)
    parser.add_argument('-o', '--output', help='Output file', required=True)
    parser.add_argument('-s', '--sheet', help='sheet to modify or translate', default='Import')
    parser.add_argument('-a', '--anytone', help='sheet name to create in anytone cps format')
    parser.add_argument('-y', '--yaesu', help='sheet in repeaterbook format to modify to RTSystems FT70 format')
    args = parser.parse_args()

    if not args.anytone and not args.yaesu:
        print("Error, must specify either anytone or yaesu")
        exit(1)

    workbook = openpyxl.load_workbook(args.input)

    if args.anytone:
        populate_anytone(workbook, 'Anytone', args.sheet)
    else:
        translate_repeaterbook(workbook, args.sheet)
    
    workbook.save(args.output)

if __name__ == '__main__':
    main()
