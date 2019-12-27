import pandas as pd

df = pd.read_excel('ViaDec.xlsx')
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
    VS2_check = input("Viasat 2:")
    Ptria_check = input("Ptria:")
    Etria_check = input("Etria:")
    Dish_check = input("Dish:")
    SB2_check = input("SB2:")
    SB2plus_check = input("SB2+:")

    vs2_amount = elist.iloc[x,1] + int(VS2_check)
    elist.at[x,'Viasat 2'] = vs2_amount

    ptria_amount = elist.iloc[x, 2] + int(Ptria_check)
    elist.at[x, 'Ptria'] = ptria_amount

    etria_amount = elist.iloc[x, 3] + int(Etria_check)
    elist.at[x, 'Etria'] = etria_amount

    dish_amount = elist.iloc[x, 4] + int(Dish_check)
    elist.at[x, 'Dish'] = dish_amount

    sb2_amount = elist.iloc[x, 5] + int(SB2_check)
    elist.at[x, 'SB2'] = sb2_amount

    sb2plus_amount = elist.iloc[x, 6] + int(SB2plus_check)
    elist.at[x, 'SB2+'] = sb2plus_amount


equipment_used = []

def equip_used(tech):
    global size
    global elist
    global equipment_used
    x = 0
    while x != size:
        tech1 = df.iloc[x,26]
        if tech1 == tech:
            equipment_used.append(df.iloc[x,33])
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
    for a in equipment_used:
        if a == 'SB2':
            sb2_amount = elist.iloc[y, 5] - 1
            elist.at[y,'SB2'] = sb2_amount
            dish_amount = elist.iloc[y, 4] - 1
            elist.at[y, 'Dish'] = dish_amount
            etria_amount = elist.iloc[y, 3] - 1
            elist.at[y, 'Etria'] = etria_amount
        elif a == 'VS2SPKWIFI':
            vs2_amount = (elist.iloc[y, 1]) - 1
            elist.at[y, 'Viasat 2'] = vs2_amount
            dish_amount = elist.iloc[y, 4] - 1
            elist.at[y, 'Dish'] = dish_amount
            ptria_amount = elist.iloc[y, 2] - 1
            elist.at[y, 'Ptria'] = ptria_amount
        elif a == 'SB2PLUS':
            sb2plus_amount = elist.iloc[y, 6] - 1
            elist.at[y, 'SB2+'] = sb2plus_amount
            dish_amount = elist.iloc[y, 4] - 1
            elist.at[y, 'Dish'] = dish_amount
            etria_amount = elist.iloc[y, 3] - 1
            elist.at[y, 'Etria'] = etria_amount

    print(elist)

checkout("Johnny")

print(elist)

#equip_used("Johnny")

#subtract("Johnny")

print(equipment_used)