from openpyxl import load_workbook
import math
import argparse
from scipy.optimize import fsolve
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
parser.add_argument('filename', type=argparse.FileType('r'))
parser.add_argument('fileformat')
parser.add_argument('--rso', help='Value of RS-0')
parser.add_argument('--amembrane', help='Value of amembrane')
parser.add_argument('--gas', help='Which gas is used')
args = parser.parse_args()
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
rsi = [602] if args.gas == 'H2'else fsolve(calc_rsi, 1000)
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

data = {"BG": avg_bg, "Raw": h2_raw, "p-ms-ar": pmsar, "rs-0": args.rso, "Real": real, "I-O": io,
        "rs-i"
        : str(rsi[0]), "P-MS-I": pmsi, "P-MS-Tot": pmstot, "P-p,i": ppi, "P-p,Ar,i": ppar, "Q-p,i/Q-Ar,i": qpiqar,
        "Q-Sweep(Ar)":
            qsweep, "Q-permeate": qpermeate, "P-feed": pfeed, "d-membrane": dmembrane, "A-membrane": args.amembrane,
        "Permeance in bar": permeancebar, "Permeance in GPU": permeancegpu, "Permeance in Pa": permeancepa,
        "L-membrane": lmembrane, "Permeability in Barrer": permeability}

gas_list = [sheet[gas_column + str(i)].value for i in range(switch_row - 1, sheet.max_row + 1)]
processed_signal = [(gas_list[i] - avg_bg * float(args.rso)) / rsi[0] for i in range(len(gas_list))]
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

if args.fileformat == "json":
    with open("output_in_json.json", "w") as outfile:
        json.dump(data, outfile)
elif args.fileformat == "csv":
    fieldnames = ['BG', 'Raw', 'p-ms-ar', 'rs-0', 'Real', 'I-0', 'rs-i', 'P-MS-I', 'P-MS-Tot', 'P-p,i', 'P-p,Ar,i',
                  'Q-p,i/Q-Ar,i',
                  'Q-Sweep(Ar)', 'Q-permeate', 'P-feed', 'd-membrane', 'A-membrane', 'Permeance in bar',
                  'Permeance in GPU',
                  'Permeance in Pa', 'Permeability in Barrer', 'L-membrane', 'I-O']
    with open("output_in_csv.csv", 'w') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows([data])
