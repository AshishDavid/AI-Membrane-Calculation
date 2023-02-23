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


def calc_avg_bg():
    count = 0
    sum_bg_n = 0
    for k in range(2, sheet.max_row + 1):
        if sheet["H" + str(k)].value == "switch time":
            break
        count += 1
    for k in range(count - 30, count + 1):
        sum_bg_n = sum_bg_n + int(sheet["D" + str(k)].value)
    avg_bg_n = sum_bg_n / 30
    return avg_bg_n


def calc_rsi(x):
    return 355 * math.log10(
        (pfeed * (io / x) / ((io / x) + pmsar)) / (pfeed * pmsar / ((io / x) + pmsar)) * 100) - 32 - x


parser = argparse.ArgumentParser()
parser.add_argument('filename', type=argparse.FileType('r'))
parser.add_argument('fileformat')
parser.add_argument('rso')
parser.add_argument('dmembrane')
args = parser.parse_args()
workbook = load_workbook(filename=args.filename.name)
sheet = workbook.active

avg_bg = calc_avg_bg()
print("BG:   ", end='')
print(avg_bg)
for i in range(1, sheet.max_row + 1):
    if sheet["H" + str(i)].value == "switch time":
        switch_row = i
        break

raw = {}
prms = {}
sum_ = 0
c = 0
for j in range(switch_row, sheet.max_row + 1):
    if sheet["D" + str(j)].value in raw:
        raw[sheet["D" + str(j)].value] += 1
    else:
        raw[sheet["D" + str(j)].value] = 1
    if sheet["G" + str(j)].value in prms:
        prms[sheet["G" + str(j)].value] += 1
    else:
        prms[sheet["G" + str(j)].value] = 1
    sum_ = sum_ + sheet["G" + str(j)].value
    c = c + 1
h2_raw = max(zip(raw.values(), raw.keys()))[1]
h2_prms = max(zip(prms.values(), prms.keys()))[1]
print("raw:  ", end='')
print(h2_raw)
print("p-ms-ar:  ", end='')
pmsar = sum_ / c
print(pmsar)

print("rs-0: ", end='')
print(args.rso)

real = h2_raw - avg_bg
print("Real: ", end='')
print(real)

io = real * 126.9
print("I-0: ", end='')
print(io)

print("rs-i: ", end='')
rsi = fsolve(calc_rsi, 1)
print(rsi[0])

pmsi = io / rsi[0]
print("P-MS-I: ", end='')
print(pmsi)

pmstot = pmsar + pmsi
print("P-MS-Tot: ", end='')
print(pmstot)

ppi = 1.01 * pmsi / pmstot
print("P-p,i: ", end='')
print(ppi)

ppar = 1.01 * pmsar / pmstot
print("P-p,Ar,i: ", end='')
print(ppar)

qpiqar = ppi / ppar * 100
print("Q-p,i/Q-Ar,i: ", end='')
print(qpiqar)

print("Q-Sweep(Ar): ", end='')
print(qsweep, end='')
print(" mln/min")

qpermeate = qsweep * qpiqar / 100
print("Q-permeate: ", end='')
print(qpermeate, end='')
print(" mln/min")

amembrane = 3.141593 / 4 * float(args.dmembrane) * float(args.dmembrane)

print("P-feed: ", end='')
print(pfeed, end='')
print(" bar")

print("d-membrane: ", end='')
print(args.dmembrane, end='')
print(" cm")

print("A-membrane: ", end='')
print(amembrane, end='')
print(" cm^2")

permeancebar = 0.6 * qpermeate / (pfeed - ppi) / amembrane
print("Permeance in m^3(STP)/(m^2 h bar): ", end='')
print(permeancebar)

permeancegpu = 370 * permeancebar
print("Permeance in GPU: ", end='')
print(permeancegpu)

permeancepa = permeancegpu * 3.35 / 1000
permeability = permeancegpu * lmembrane

print("Permeance in 10^-7 mol/(m^2 s Pa): ", end='')
print(permeancepa)

print("L-membrane in μm: ", end='')
print(lmembrane)

print("Permeability in Barrer: ", end='')
print(permeability)
data = {"BG": avg_bg, "Raw": h2_raw, "p-ms-ar": pmsar, "rs-0": args.rso, "Real": real, "I-O": io,
        "rs-i"
        : str(rsi[0]), "P-MS-I": pmsi, "P-MS-Tot": pmstot, "P-p,i": ppi, "P-p,Ar,i": ppar, "Q-p,i/Q-Ar,i": qpiqar,
        "Q-Sweep(Ar)":
            qsweep, "Q-permeate": qpermeate, "P-feed": pfeed, "d-membrane": args.dmembrane, "A-membrane": amembrane,
        "Permeance in bar": permeancebar, "Permeance in GPU": permeancegpu, "Permeance in Pa": permeancepa,
        "L-membrane": lmembrane, "Permeability in Barrer": permeability}
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
