from openpyxl import load_workbook
import math
from scipy.optimize import fsolve


def func(x):
    return 355 * math.log10(
        (1.01 * (18654.3 / x) / ((18654.3 / x) + 2213.2)) / (1.01 * 2213.2 / ((18654.3 / x) + 2213.2)) * 100) - 32 - x


file_location = input("Enter the location of raw data")
file_location = file_location.replace("\\", "//")
workbook = load_workbook(filename=file_location)
sheet = workbook.active

sum_bg = 0
count = 0
for i in range(2, sheet.max_row + 1):
    if sheet["H" + str(i)].value == "switch time":
        break
    count += 1
for i in range(count - 30, count + 1):
    sum_bg = sum_bg + int(sheet["D" + str(i)].value)
avg_bg = sum_bg / 30
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
print(126.9)

print("rs-i: ", end='')
rsi = fsolve(func, 1)
print(rsi[0])

real = h2_raw - avg_bg
print("Real: ", end='')
print(real)
io = real * 126.9
print("I-0: ", end='')
print(io)
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
qsweep = 50
print("Q-Sweep(Ar): ", end='')
print(qsweep, end='')
print(" mln/min")
qpermeate = qsweep * qpiqar / 100
print("Q-permeate: ", end='')
print(qpermeate, end='')
print(" mln/min")
pfeed = 1.01
dmembrane = 2.7
amembrane = 3.141593 / 4 * dmembrane * dmembrane
print("P-feed: ", end='')
print(pfeed, end='')
print(" bar")
print("d-membrane: ", end='')
print(dmembrane, end='')
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
lmembrane = 0.1
permeability = permeancegpu * lmembrane
print("Permeance in 10^-7 mol/(m^2 s Pa): ", end='')
print(permeancepa)
print("L-membrane in μm: ", end='')
print(lmembrane)
print("Permeability in Barrer: ", end='')
print(permeability)
