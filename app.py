from openpyxl import load_workbook

path = r"C:\Users\satya\Desktop\datasheet\12600637.xlsx"

wb = load_workbook(path, data_only=True)
ws = wb["Daily Log"]


def get_values(start, end, col):
    data = []
    # as ending row is exclusive to have to add + 1
    for row in range(start, end + 1):
        value = ws[f"{col}{row}"].value
         
        if value is not None:
            data.append(value)

    return data


def average(data):
    if len(data) == 0:
        return 0

    return sum(data) / len(data)


def correlation(x, y):
    if len(x) < 2 or len(x) != len(y):
        return 0

    avg_x = average(x)
    avg_y = average(y)

    numerator = 0
    sum_x = 0
    sum_y = 0

    for i in range(len(x)):
        numerator += (x[i] - avg_x) * (y[i] - avg_y)
        sum_x += (x[i] - avg_x) ** 2
        sum_y += (y[i] - avg_y) ** 2

    denominator = (sum_x * sum_y) ** 0.5

    if denominator == 0:
        return 0

    return numerator / denominator

#on excel sheet  13th row -> 17 aug
#48th row -> 21 sept
start_row = 13
end_row = 48

sleep_data = get_values(start_row, end_row, "B")
coding_data = get_values(start_row, end_row, "E")
study_data = get_values(start_row, end_row, "D")
print("avg study :" , round(sum(study_data) / len(study_data),2))
class_data = get_values(start_row, end_row, "F")
print("avg class :" , round(sum(class_data) / len(class_data),2))
fitness_data = get_values(start_row, end_row, "C")
other_data = get_values(start_row, end_row, "H")
total_data = get_values(start_row, end_row, "I")
free_data = get_values(start_row, end_row, "J")

SRI = average(sleep_data)
TPI = average(coding_data)
PhAI = average(fitness_data)
otherA = average(other_data)
freeA = average(free_data)
TUI = average(total_data)

ABI_data = []

for value in total_data:
    #in a day total mins is 24 * 60  =  1440 mins
    ABI_data.append(1440 - value)

ABI = average(ABI_data)

study_class = []

for row in range(start_row, end_row + 1):
    study = ws[f"D{row}"].value
    classes = ws[f"F{row}"].value

    if study is not None and classes is not None:
        study_class.append(study + classes)

AAI = average(study_class)

feeling = {
    "excellent": 5,
    "very good": 5,
    "good": 4,
    "normal": 3,
    "neutral": 3,
    "low": 2,
    "bad": 1
}

satisfaction = {
    "very satisfied": 5,
    "satisfied": 4,
    "satishfied": 4,
    "normal": 3,
    "neutral": 3,
    "unsatisfied": 2,
    "very unsatisfied": 1
}

energy = {
    "high": 3,
    "medium": 2,
    "low": 1
}

experience_data = []

for row in range(start_row, end_row + 1):
    feel = ws[f"K{row}"].value #my feelings
    satisfy = ws[f"L{row}"].value 

    scores = []

    if feel is not None and feel.lower().strip() in feeling:
        scores.append(feeling[feel.lower().strip()])

    if satisfy is not None and satisfy.lower().strip() in satisfaction:
        scores.append(satisfaction[satisfy.lower().strip()])

    if len(scores) > 0:
        experience_data.append(average(scores))

EI = average(experience_data)

dci_data = get_values(9, 52, "B")
expected_days = 44 # from 13 aug to 25 sept as mentioned in the pdf 
DCI = (len(dci_data) / expected_days) * 100
PAI = ((0.15 * TPI) + (0.20 * AAI) + (0.15 * PhAI) + (0.20 * SRI) + (0.15 * TUI) + (0.10 * EI) + (0.05 * DCI))

print("Sleep Recovery Index (SRI):", round(SRI, 2), "minutes")
print("Technical Productivity Index (TPI):", round(TPI, 2), "minutes")
print("Academic Activity Index (AAI):", round(AAI, 2), "minutes")
print("Physical Activity Index (PhAI):", round(PhAI, 2), "minutes")
print("Average Other Activity:", round(otherA, 2), "minutes")
print("Average Free Time:", round(freeA, 2), "minutes")
print("Accounted Busy Index (ABI):", round(ABI, 2), "minutes")
print("Total Utilization Index (TUI):", round(TUI, 2), "minutes")
print("Experience Index (EI):", round(EI, 2), "/ 5")
print("Data Continuity Index (DCI):", round(DCI, 2), "%")
print("Personal Activity Index (PAI):", round(PAI, 2))


coding_energy_x = []
coding_energy_y = []

sleep_energy_x = []
sleep_energy_y = []

study_satisfaction_x = []
study_satisfaction_y = []

for row in range(start_row, end_row + 1):
    coding = ws[f"E{row}"].value  #coding time
    sleep = ws[f"B{row}"].value # sleeping time
    study = ws[f"D{row}"].value # how much i have studied
    energy_value = ws[f"M{row}"].value # energy levels
    satisfy = ws[f"L{row}"].value # satishfaction level

    if coding is not None and energy_value is not None:
        key = energy_value.lower().strip()

        if key in energy:
            coding_energy_x.append(coding)
            coding_energy_y.append(energy[key])

    if sleep is not None and energy_value is not None:
        key = energy_value.lower().strip()

        if key in energy:
            sleep_energy_x.append(sleep)
            sleep_energy_y.append(energy[key])

    if study is not None and satisfy is not None:
        key = satisfy.lower().strip()

        if key in satisfaction:
            study_satisfaction_x.append(study)
            study_satisfaction_y.append(satisfaction[key])


print("----- Relationship  ANALYSIS -----")
print("Coding <-> Energy:", round(correlation(coding_energy_x, coding_energy_y), 2))
print("Sleep <-> Energy:", round(correlation(sleep_energy_x, sleep_energy_y), 2))
print("Study <-> Satisfaction:", round(correlation(study_satisfaction_x, study_satisfaction_y), 2))