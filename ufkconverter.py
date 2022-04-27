"""
UfkConverter from XLS to BD0 & VT0 files
modified by Feliks Nabiullin @feel2code
"""

import contextlib
from random import randint

import pandas as pd

# File sources
NEW_FILE_NAME = randint(1, 9)
SOURCEFILE = 'VT_BD_SF.xls'
BDOUTPUT = f'1200{NEW_FILE_NAME}ABC.BD0'
VTOUT = f'1200{NEW_FILE_NAME}ABC.VT0'


def panda_to_date(pandate):
    """makes pddate to string date."""
    date_list = str(pandate).split()[0].split('-')
    date_list.reverse()
    return '.'.join(date_list)


# Create VT report
VTHEADER = "FK|TXBD120101|Converter_PAY_UFK|1.5.4| ТЗ|"
vtsheet = pd.read_excel(
    SOURCEFILE, sheet_name="12000ABC.VT0", convert_float=False)

VTFROM_LIST = "|".join(
    map(lambda x: "" if str(x) == "nan" else str(
        x), vtsheet.iloc[1:3, 6].tolist())
)
VTTO_LIST = vtsheet.iloc[4:7, 6].tolist()
VTTO_LIST[0] = int(VTTO_LIST[0])
VTVT_LIST = list(map(lambda x: "" if pd.isna(x) else x,
                 vtsheet.iloc[8:30, 6].tolist()))
# VTVT_LIST = vtsheet.iloc[8:30, 6].tolist()
VTVT_LIST[2] = VTVT_LIST[2].strftime('%d.%m.%Y')
VTVT_LIST[3] = VTVT_LIST[3].strftime('%d.%m.%Y')
VTVT_LIST[16] = VTVT_LIST[16].strftime('%d.%m.%Y')
VTVT_LIST[17] = "{:.2f}".format(VTVT_LIST[17])
VTVT_LIST[18] = "{:.2f}".format(VTVT_LIST[18])
VTVT_LIST[19] = "{:.2f}".format(VTVT_LIST[19])
VTVT_LIST[20] = "{:.2f}".format(VTVT_LIST[20])
VTVT_LIST[21] = "{:.2f}".format(VTVT_LIST[21])
VTSUM_LIST = list(
    map(lambda x: '{:.2f}'.format(x), vtsheet.iloc[31:41, 6].tolist()))

VTOPER = vtsheet.iloc[42:60, 6:]
VTNOPER = vtsheet.iloc[61:, 6:]
with open(VTOUT, "w") as vtout:
    with contextlib.redirect_stdout(vtout):
        print(VTHEADER)
        print(f"FROM|{VTFROM_LIST}|")
        print(f'TO|{"|".join(map(str, VTTO_LIST))}|')
        print(f'VT|{"|".join(map(str, VTVT_LIST))}|')
        print(f'VTSUM|{"|".join(map(str, VTSUM_LIST))}|')

        for cnt in range(len(VTOPER.columns)):
            CURR_VTOPER = VTOPER.iloc[0:, cnt: cnt + 1]
            CURR_VTNOPER = VTNOPER.iloc[0:, cnt: cnt + 1]
            print(
                f"VTOPER|{CURR_VTOPER.iloc[0,0]}|"
                f"{CURR_VTOPER.iloc[1,0]}|{CURR_VTOPER.iloc[2,0]}|",
                end=""
            )
            print(
                f"{CURR_VTOPER.iloc[3,0].strftime('%d.%m.%Y')}|",
                end=""
            )
            print(f"{CURR_VTOPER.iloc[4,0]}|", end="")
            print(f"{CURR_VTOPER.iloc[5,0]:.0f}|", end="")
            print(
                f"{'' if pd.isna(CURR_VTOPER.iloc[6,0]) else CURR_VTOPER.iloc[6,0].strftime('%d.%m.%Y')}|",
                end=""
            )
            print(
                f'{"|".join(map(lambda x: str("{:.2f}".format(x)), CURR_VTOPER.iloc[7:10, 0]))}|',
                end=""
            )
            print(f"{CURR_VTOPER.iloc[10,0]}|", end="")
            print(
                f"{CURR_VTOPER.iloc[12,0]:.0f}|{CURR_VTOPER.iloc[13,0]}|", end="")
            print(
                f"{'' if pd.isna(CURR_VTOPER.iloc[14,0]) else str(CURR_VTOPER.iloc[14,0])}|", end="")
            print(
                f'{"|".join(map(lambda x: "" if pd.isna(x) else str("{:.0f}".format(x)), CURR_VTOPER.iloc[15:18,0]))}|', end="")
            print()
            print(
                f'VTNOPER|{CURR_VTNOPER.iloc[0,0]}|{"|".join(map(lambda x: str("{:.0f}".format(x)), CURR_VTNOPER.iloc[1:3, 0]))}|', end="")
            print(
                f"{'' if pd.isna(CURR_VTNOPER.iloc[4,0]) else CURR_VTNOPER.iloc[3,0].strftime('%d.%m.%Y')}|", end="")
            print(
                f'{"|".join(map(lambda x: str("{:.2f}".format(x)), CURR_VTNOPER.iloc[4:6, 0]))}|', end="")
            print(f"{CURR_VTNOPER.iloc[6,0]:.0f}|")
vtout.close()

# Create BD report
BDHEADER = "FK|TXBD120101|Converter_PAY_UFK|1.5.4| ТЗ|"
bdsheet = pd.read_excel(
    SOURCEFILE, sheet_name="12000ABC.BD0", convert_float=False)
bdout = open(BDOUTPUT, "w")
print(BDHEADER, file=bdout)
from_list = "|".join(
    map(lambda x: "" if str(x) == "nan" else str(
        x), bdsheet.iloc[1:3, 6].tolist())
)

start_to_list = bdsheet.iloc[4:9, 6].tolist()
start_to_list[2] = int(start_to_list[2])
to_list = "|".join(
    map(lambda x: "" if str(x) == "nan" else str(
        x), start_to_list)
)

block = bdsheet.iloc[10:16, 6].tolist()
block[0] = str(block[0])
block[1] = block[1].strftime("%d.%m.%Y")
block[4] = int(block[4])
block[5] = "{:.2f}".format(block[5])
block.insert(4, '')
bd_list = "|".join(map(lambda x: "" if str(x) == "nan" else str(x), block))

print(f"FROM|{from_list}|", file=bdout)
print(f"TO|{to_list}|", file=bdout)
print(f"BD|{bd_list}|", file=bdout)

bdpd = bdsheet.iloc[17:57, 6:]
bdpdst = bdsheet.iloc[58:, 6:]

for cnt in range(len(bdpd.columns)):
    curr_bdpd = bdpd.iloc[0:, cnt: cnt + 1]
    curr_bdpdst = bdpdst.iloc[0:, cnt: cnt + 1]

    # format bdpd str
    cell8 = curr_bdpd.iloc[8, 0]
    if f'{cell8:.0f}' == 'nan' or f'{cell8:.0f}' == '0':
        nan_replace = ''
    else:
        nan_replace = f'{curr_bdpd.iloc[8,0]:.0f}'

    print(
        f"BDPD|{curr_bdpd.iloc[0,0]}|"
        f"{panda_to_date(curr_bdpd.iloc[1,0])}|"
        f"{curr_bdpd.iloc[2,0]:.2f}|"
        f"{curr_bdpd.iloc[3,0]:.0f}|"
        f"{panda_to_date(curr_bdpd.iloc[4,0])}|"
        f"{panda_to_date(curr_bdpd.iloc[5,0])}|"
        f"{curr_bdpd.iloc[6,0]}|"
        f"{curr_bdpd.iloc[7,0]}|"
        f"{nan_replace}|"
        f"{curr_bdpd.iloc[9,0]}|"
        f"{curr_bdpd.iloc[10,0]}|"
        f"{curr_bdpd.iloc[11,0]:.0f}|"
        f"{curr_bdpd.iloc[12,0]}|"
        f"{curr_bdpd.iloc[13,0]}|"
        f"{curr_bdpd.iloc[14,0]:.0f}|"
        f"{curr_bdpd.iloc[15,0]:.0f}|"
        f"{curr_bdpd.iloc[16,0]}|"
        f"{curr_bdpd.iloc[17,0]}|"
        f"{curr_bdpd.iloc[18,0]:.0f}|"
        f"{curr_bdpd.iloc[19,0]}|"
        f"{curr_bdpd.iloc[20,0]}|"
        f"{panda_to_date(curr_bdpd.iloc[21,0])}|"
        f"{curr_bdpd.iloc[22,0]}|"
        f"{curr_bdpd.iloc[23,0]}|"
        f"{curr_bdpd.iloc[24,0]}|"
        f"{curr_bdpd.iloc[25,0]}|"
        f"{curr_bdpd.iloc[26,0]:.0f}|"
        f"{curr_bdpd.iloc[27,0]}|"
        f"{curr_bdpd.iloc[28,0]}|"
        f"{curr_bdpd.iloc[29,0]}|"
        f"{curr_bdpd.iloc[30,0]}|"
        f"{curr_bdpd.iloc[31,0]}|",
        end="",
        file=bdout,
    )
    if pd.isna(curr_bdpd.iloc[32, 0]):
        print("|", end="", file=bdout)
    else:
        print(f"{curr_bdpd.iloc[32,0]:.0f}|", end="", file=bdout)

    cell35 = curr_bdpd.iloc[35, 0]
    print(
        f"{curr_bdpd.iloc[33,0]}|"
        f"{'' if pd.isna(curr_bdpd.iloc[34,0]) else curr_bdpd.iloc[34,0]}|"
        f"{'' if pd.isna(cell35) else cell35.strftime('%d.%m.%Y')}|",
        end="",
        file=bdout
        # f"{'' if pd.isna(curr_bdpd.iloc[36,0])
        #  else type(curr_bdpd.iloc[36,0])}|"
    )
    if pd.isna(curr_bdpd.iloc[36, 0]):
        print("|", end="", file=bdout)
    else:
        print(f"{curr_bdpd.iloc[36,0]:.2f}|", end="", file=bdout)
    print(f"{curr_bdpd.iloc[37,0]}|", end="", file=bdout)
    print(
        f"{panda_to_date(curr_bdpd.iloc[38,0])}|",
        end="",
        file=bdout
    )
    print(f"{curr_bdpd.iloc[39,0]}|", end="", file=bdout)
    print(file=bdout)
    # format bdpdst str
    print("BDPDST|", end="", file=bdout)
    print(
        f"{curr_bdpdst.iloc[0,0]}|" f"{curr_bdpdst.iloc[1,0]:.0f}|",
        end="",
        file=bdout
    )
    if pd.isna(curr_bdpdst.iloc[2, 0]):
        print("|", end="", file=bdout)
    else:
        print(f"{curr_bdpdst.iloc[2,0]}|", end="", file=bdout)
    if pd.isna(curr_bdpdst.iloc[3, 0]):
        print("|", end="", file=bdout)
    else:
        print(f"{curr_bdpdst.iloc[3,0]:.0f}|", end="", file=bdout)
    print(f"{curr_bdpdst.iloc[4,0]:.0f}|", end="", file=bdout)
    print(f"{curr_bdpdst.iloc[5,0]:.2f}|", end="", file=bdout)
    print(f"{curr_bdpdst.iloc[6,0]}|", end="", file=bdout)
    if pd.isna(curr_bdpdst.iloc[7, 0]):
        print("|", end="", file=bdout)
    else:
        print(
            f"{curr_bdpdst.iloc[7,0]}|", end="", file=bdout
        )
        # нужно проверить, возможно необходим формат
        # (:.0f) для отрезания дробной части
    if pd.isna(curr_bdpdst.iloc[7, 0]):
        print("|", end="", file=bdout)
    else:
        print(f"{curr_bdpdst.iloc[7, 0]}", end="", file=bdout)
    print(file=bdout)
bdout.close()


def remove_line(file_name, lineskip):
    """ Removes a given line from a file """
    with open(file_name, 'r') as read_file:
        lines = read_file.readlines()

    currentline = 1
    with open(file_name, 'w') as write_file:
        for line in lines:
            if currentline == lineskip:
                pass
            else:
                write_file.write(line)
            currentline += 1


remove_line(BDOUTPUT, 9)
remove_line(BDOUTPUT, 9)
remove_line(BDOUTPUT, 9)
