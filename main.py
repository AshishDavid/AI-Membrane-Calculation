from openpyxl import load_workbook
import math
import argparse
from scipy.optimize import fsolve
from scipy.stats import linregress
import numpy as np
import json
import csv

# defining constants
pfeed = 1.01
qsweep = 50
lmembrane = 0.1
dmembrane = 2.7


def calc_avg_bg():
    sum_bg_n = 0
    for k in range(2, sheet.max_row + 1):
        if sheet["H" + str(k)].value == "switch time":
            count = k - 1
            break
    for k in range(count - 30, count + 1):
        sum_bg_n = sum_bg_n + int(sheet[gas_column + str(k)].value)
    avg_bg_n = sum_bg_n / 31
    return avg_bg_n


def calc_rsi(x):
    diction = {'CH4': [79.21, 233.34], 'N2': [-41.31, 91.58], 'He': [-58.92, 267.25]}
    if args.gas == 'CO2':
        return 366.542 * math.exp(0.18068 / (
                (pfeed * (io / x) / ((io / x) + pmsar)) / (pfeed * pmsar / ((io / x) + pmsar)) * 100 + 0.08308)) - x
    else:
        return diction[args.gas][0] * math.log10(
            (pfeed * (io / x) / ((io / x) + pmsar)) / (pfeed * pmsar / ((io / x) + pmsar)) * 100) \
            + diction[args.gas][1] - x


def calc_switch_row():
    for i in range(2, sheet.max_row + 1):
        if sheet["H" + str(i)].value == "switch time":
            return i


parser = argparse.ArgumentParser()
parser.add_argument('--filename', type=argparse.FileType('r'), help='Location of File')
parser.add_argument('--filename1', type=argparse.FileType('r'), help='Location of 1st File')
parser.add_argument('--filename2', type=argparse.FileType('r'), help='Location of 2nd File')
parser.add_argument('--filename3', type=argparse.FileType('r'), help='Location of 3rd File')
parser.add_argument('--filename4', type=argparse.FileType('r'), help='Location of 4th File')
parser.add_argument('--filename5', type=argparse.FileType('r'), help='Location of 5th File')
parser.add_argument('fileformat')
parser.add_argument('--rso', help='Value of RS-0')
parser.add_argument('--rso1', help='Value of 1st RS-0')
parser.add_argument('--rso2', help='Value of 2nd RS-0')
parser.add_argument('--rso3', help='Value of 3rd RS-0')
parser.add_argument('--rso4', help='Value of 4th RS-0')
parser.add_argument('--rso5', help='Value of 5th RS-0')
parser.add_argument('--amembrane', help='Value of amembrane')
parser.add_argument('--amembrane1', help='Value of 1st amembrane')
parser.add_argument('--amembrane2', help='Value of 2nd amembrane')
parser.add_argument('--amembrane3', help='Value of 3rd amembrane')
parser.add_argument('--amembrane4', help='Value of 4th amembrane')
parser.add_argument('--amembrane5', help='Value of 5th amembrane')
parser.add_argument('--gas', help='Which gas is used')
parser.add_argument('--gas1', help='Which gas is used for 1st')
parser.add_argument('--gas2', help='Which gas is used for 2nd')
parser.add_argument('--gas3', help='Which gas is used for 3rd')
parser.add_argument('--gas4', help='Which gas is used for 4th')
parser.add_argument('--gas5', help='Which gas is used for 5th')
args = parser.parse_args()

filenames = [args.filename1]
rso = [args.rso1]
amembrane = [args.amembrane1]
gases = [args.gas1]

if args.filename2 is not None:
    filenames.append(args.filename2)
    rso.append(args.rso2)
    amembrane.append(args.amembrane2)
    gases.append(args.gas2)
if args.filename3 is not None:
    filenames.append(args.filename3)
    rso.append(args.rso3)
    amembrane.append(args.amembrane3)
    gases.append(args.gas3)
if args.filename4 is not None:
    filenames.append(args.filename4)
    rso.append(args.rso4)
    amembrane.append(args.amembrane4)
    gases.append(args.gas4)
if args.filename5 is not None:
    filenames.append(args.filename5)
    rso.append(args.rso5)
    amembrane.append(args.amembrane5)
    gases.append(args.gas5)

for tr in range(len(filenames)):
    args.filename = filenames[tr]
    args.rso = rso[tr]
    args.gas = gases[tr]
    args.amembrane = amembrane[tr]
    workbook = load_workbook(filename=args.filename.name)
    sheet = workbook.active

    if args.gas == 'H2':
        gas_column = 'D'
    if args.gas == 'He':
        gas_column = 'C'
    if args.gas == 'CH4' or args.gas == 'N2':
        gas_column = 'E'
    if args.gas == 'CO2':
        gas_column = 'F'

    avg_bg = calc_avg_bg()

    switch_row = calc_switch_row()
    raw = {}
    prms = {}
    sum_ = 0
    c = 0

    for j in range(switch_row, sheet.max_row + 1):
        if sheet[gas_column + str(j)].value in raw:
            raw[sheet[gas_column + str(j)].value] += 1
        else:
            if sheet[gas_column + str(j)].value != 0:
                raw[sheet[gas_column + str(j)].value] = 1
        if sheet["G" + str(j)].value in prms:
            prms[sheet["G" + str(j)].value] += 1
        else:
            prms[sheet["G" + str(j)].value] = 1
        sum_ = sum_ + sheet["G" + str(j)].value
        c = c + 1

    h2_raw = max(zip(raw.values(), raw.keys()))[1]
    h2_prms = max(zip(prms.values(), prms.keys()))[1]
    pmsar = sum_ / c
    real = h2_raw - avg_bg
    io = real * float(args.rso)
    rsi = [1171] if args.gas == 'H2' else fsolve(calc_rsi, 1000)
    pmsi = io / rsi[0]
    pmstot = pmsar + pmsi
    ppi = 1.01 * pmsi / pmstot
    ppar = 1.01 * pmsar / pmstot
    qpiqar = ppi / ppar * 100
    qpermeate = qsweep * qpiqar / 100
    permeancebar = 0.6 * qpermeate / (pfeed - ppi) / float(args.amembrane)
    permeancegpu = 370 * permeancebar
    permeancepa = permeancegpu * 3.35 / 1000
    permeability = permeancegpu * lmembrane
    gas_list = [sheet[gas_column + str(i)].value for i in range(switch_row, sheet.max_row + 1)]
    processed_signal = [((gas_list[i] - avg_bg) * float(args.rso)) / rsi[0] for i in range(len(gas_list))]
    argon_list = [sheet["G" + str(i)].value for i in range(switch_row - 1, sheet.max_row + 1)]
    sum_argon_processed = [processed_signal[i] + argon_list[i] for i in range(len(processed_signal))]
    gas_in_bar = [1.01 * processed_signal[i] / sum_argon_processed[i] for i in range(len(processed_signal))]
    argon_in_bar = [1.01 * argon_list[i] / sum_argon_processed[i] for i in range(len(processed_signal))]
    gas_mln_min = [qsweep * gas_in_bar[i] / argon_in_bar[i] for i in range(len(gas_in_bar))]
    gas_cum_vol = [0]

    for i in range(1, len(gas_list)):
        gas_cum_vol.append(0.5 * (gas_mln_min[i] + gas_mln_min[i - 1]))
    gas_cum_vol_mln = [0]
    sum_gas_cum_vol = 0

    for i in range(1, len(gas_list)):
        sum_gas_cum_vol = sum_gas_cum_vol + gas_cum_vol[i]
        gas_cum_vol_mln.append(1 / 60 * sum_gas_cum_vol)

    x_axis = [i for i in range(len(gas_list))]
    slope, intercept, r_value, p_value, std_err = linregress(x_axis[23:], gas_cum_vol_mln[23:])
    diff_coefficient = -intercept / slope

    print(f"Gas: {args.gas}")
    print(f"BG: {avg_bg:}")
    print(f"raw: {h2_raw}")
    print(f"p-ms-ar:  {pmsar}")

    print(f"rs-0: {args.rso}")

    print(f"Real: {real}")

    print(f"I-0: {io}")

    print(f"rs-i: {rsi[0]}")

    print(f"P-MS-I: {pmsi}")

    print(f"P-MS-Tot: {pmstot}")

    print(f"P-p,i: {ppi}")

    print(f"P-p,Ar,i: {ppar}")

    print(f"Q-p,i/Q-Ar,i: {qpiqar}")

    print(f"Q-Sweep(Ar): {qsweep} mln/min")

    print(f"Q-permeate: {qpermeate} mln/min")

    print(f"P-feed: {pfeed} bar")

    print(f"d-membrane: {dmembrane} cm")

    print(f"A-membrane: {args.amembrane} cm\u00b2")

    print(f"Permeance in m\u00b3(STP)/(m\u00b2 h bar): {permeancebar}")

    print(f"Permeance in GPU: {permeancegpu}")

    print(f"Permeance in 10\u207B\u2077 mol/(m\u00b2 s Pa): {permeancepa}")

    print(f"L-membrane in μm: {lmembrane}")

    print(f"Permeability in Barrer: {permeability}")

    print(f'Diffusion coefficient: {diff_coefficient} \n \n')

    data = {"Gas": args.gas, "BG": avg_bg, "Raw": h2_raw, "p-ms-ar": pmsar, "rs-0": args.rso, "Real": real, "I-O": io,
            "rs-i"
            : str(rsi[0]), "P-MS-I": pmsi, "P-MS-Tot": pmstot, "P-p,i": ppi, "P-p,Ar,i": ppar, "Q-p,i/Q-Ar,i": qpiqar,
            "Q-Sweep(Ar)":
                qsweep, "Q-permeate": qpermeate, "P-feed": pfeed, "d-membrane": dmembrane, "A-membrane": args.amembrane,
            "Permeance in bar": permeancebar, "Permeance in GPU": permeancegpu, "Permeance in Pa": permeancepa,
            "L-membrane": lmembrane, "Permeability in Barrer": permeability, "Diffusion coefficient": diff_coefficient}

    if args.fileformat == "json":
        if tr == 0:
            with open("output_in_json.json", "w") as outfile:
                json.dump(data, outfile)
        else:
            with open("output_in_json.json", "a") as outfile:
                json.dump(data, outfile)
    elif args.fileformat == "csv":
        fieldnames = ['Gas', 'BG', 'Raw', 'p-ms-ar', 'rs-0', 'Real', 'I-0', 'rs-i', 'P-MS-I', 'P-MS-Tot', 'P-p,i', 'P-p,Ar,i',
                      'Q-p,i/Q-Ar,i',
                      'Q-Sweep(Ar)', 'Q-permeate', 'P-feed', 'd-membrane', 'A-membrane', 'Permeance in bar',
                      'Permeance in GPU',
                      'Permeance in Pa', 'Permeability in Barrer', 'L-membrane', 'I-O', 'Diffusion coefficient']
        if tr == 0:
            with open("output_in_csv.csv", 'w') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows([data])
        else:
            with open("output_in_csv.csv", 'a') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows([data])
