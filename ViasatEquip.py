import pandas as pd
from texttable import Texttable as tt

df = pd.read_excel('viasatjan.1.8.xlsx')
elist = pd.read_excel('Equip.xlsx')

size = len(df)

# X order 0=Kevin, 1David, 2Kendall, 3Johnny

def checkout(tech):
    global df
    global elist
    if tech == "Kevin":
        x = 0
    elif tech == "David":
        x = 1
    elif tech == "Kendall":
        x = 2
    elif tech == "Johnny":
        x = 3
    elif tech == "James":
        x = 4
    elif tech == "Crabtree":
        x = 5
    Dish_check = input("Dish:")
    VS2_check = input("Viasat 2:")
    Ptria_check = input("Ptria:")
    SB2plus_check = input("SB2+:")
    SB2_check = input("SB2:")
    Etria_check = input("Etria:")

    vs2_amount = elist.iloc[x,2] + int(VS2_check)
    elist.at[x,'Viasat 2'] = vs2_amount

    ptria_amount = elist.iloc[x, 3] + int(Ptria_check)
    elist.at[x, 'Ptria'] = ptria_amount

    etria_amount = elist.iloc[x, 6] + int(Etria_check)
    elist.at[x, 'Etria'] = etria_amount

    dish_amount = elist.iloc[x, 1] + int(Dish_check)
    elist.at[x, 'Dish'] = dish_amount

    sb2_amount = elist.iloc[x, 5] + int(SB2_check)
    elist.at[x, 'SB2'] = sb2_amount

    sb2plus_amount = elist.iloc[x, 4] + int(SB2plus_check)
    elist.at[x, 'SB2+'] = sb2plus_amount


equipment_used = []
fsm_num = []
date_completed = []
job_type = []

def equip_used(tech):
    global size
    global elist
    global equipment_used
    x = 0
    while x != size:
        tech1 = df.iloc[x,26]
        if tech1 == tech:
            equipment_used.append(df.iloc[x,33])
            fsm_num.append(df.iloc[x,0])
            date_completed.append(df.iloc[x,5])
            job_type.append(df.iloc[x,7])
        x = x+1

def subtract(tech):
    global elist
    global equipment_used
    if tech == "Kevin":
        y = 0
    elif tech == "David":
        y = 1
    elif tech == "Kendall":
        y = 2
    elif tech == "Johnny":
        y = 3
    elif tech == "James":
        y = 4
    elif tech == "Crabtree":
        y = 5
    for a in equipment_used:
        if a == 'SB2':
            sb2_amount = elist.iloc[y, 2] - 1
            elist.at[y,'SB2'] = sb2_amount
            dish_amount = elist.iloc[y, 1] - 1
            elist.at[y, 'Dish'] = dish_amount
            etria_amount = elist.iloc[y, 6] - 1
            elist.at[y, 'Etria'] = etria_amount
        elif a == 'VS2SPKWIFI':
            vs2_amount = (elist.iloc[y, 2]) - 1
            elist.at[y, 'Viasat 2'] = vs2_amount
            dish_amount = elist.iloc[y, 1] - 1
            elist.at[y, 'Dish'] = dish_amount
            ptria_amount = elist.iloc[y, 3] - 1
            elist.at[y, 'Ptria'] = ptria_amount
        elif a == 'SB2PLUS':
            sb2plus_amount = elist.iloc[y, 4] - 1
            elist.at[y, 'SB2+'] = sb2plus_amount
            dish_amount = elist.iloc[y, 1] - 1
            elist.at[y, 'Dish'] = dish_amount
            etria_amount = elist.iloc[y, 6] - 1
            elist.at[y, 'Etria'] = etria_amount
    for a in job_type:
        if a == 'Upgrade':
            dish_amount = elist.iloc[y, 1] + 1
            elist.at[y, 'Dish'] = dish_amount
    print(elist.iloc[y])

checkout("Johnny")

equip_used("Johnny")

subtract("Johnny")


tab = tt()
headings = ['Equipment','FSM#','Date Completed','Job Type']
tab.header(headings)

for row in zip(equipment_used,fsm_num,date_completed,job_type):
    tab.add_row(row)

s = tab.draw()
print(s)