import random


def values_for_paystub(year):
    if year == "2022":
        all_vars = {
            "federal_first": 15,
            "province_first": 5.05,
            "federal_second": 20.5,
            "province_second": 9.15,
            "federal_three": 26,
            "province_three": 11.16,
            "federal_four": 29,
            "province_four": 12.16,
            "federal_five": 33,
            "province_five": 13.16,
            
            "EI_Rate": 1.58,
            "CPP_Rate": 5.70,
            "EI_Maximum_Deduction": 952.74,
            "CPP_Maximum_Deduction": 3499.80,
            "last_year_to_date": 0,
        }
        return all_vars
    
    elif year == "2023":
        all_vars = {
            "federal_first": 15,
            "province_first": 5.05,
            "federal_second": 20.5,
            "province_second": 9.15,
            "federal_three": 26,
            "province_three": 11.16,
            "federal_four": 29,
            "province_four": 12.16,
            "federal_five": 33,
            "province_five": 13.16,
            
            "EI_Rate": 1.63,
            "CPP_Rate": 5.95,
            "EI_Maximum_Deduction": 1002.45,
            "CPP_Maximum_Deduction": 3754.45,
            "last_year_to_date": 0,
        }
        return all_vars


EI_Rates = [
    {"year": 2020, "ei_rate": 1.58, "max_deduction": 856.36},
    {"year": 2021, "ei_rate": 1.58, "max_deduction": 889.54},
    {"year": 2022, "ei_rate": 1.58, "max_deduction": 952.74},
    {"year": 2023, "ei_rate": 1.63, "max_deduction": 1002.45},
]


CPP_Rates = [
    {"year": 2020, "cpp_rate": 5.25, "max_deduction": 2898.00},
    {"year": 2021, "cpp_rate": 5.45, "max_deduction": 3166.45},
    {"year": 2022, "cpp_rate": 5.70, "max_deduction": 3499.80},
    {"year": 2023, "cpp_rate": 5.95, "max_deduction": 3754.45},
]

Max_EI_Insureable_Income = [
    {"year": 2020, "income": 54200},
    {"year": 2021, "income": 56300},
    {"year": 2022, "income": 60300},
    {"year": 2023, "income": 61500},
]


Max_CPP_Pensionable_Income = [
    {"year": 2020, "income": 58700},
    {"year": 2021, "income": 61600},
    {"year": 2022, "income": 64900},
    {"year": 2023, "income": 66600},
]


withdrawls = [
    {"description": "LONDON GUEST HOUSE ON LONDON", "withdraw": "35.19"},
    {"description": "MCDONALD'S #4040 LONDON", "withdraw": "8.12"},
    {"description": "APPLE.COM/BILL 866-712-7753", "withdraw": "25.34"},
    {"description": "TIM HORTONS #1048 LONDON", "withdraw": "1.67"},
    {"description": "REXALL PHARMACY #1768 LONDON", "withdraw": "44.06"},
    {"description": "DAIRY QUEEN #11882 LONDON", "withdraw": "15.86"},
    {"description": "SOUTHSIDE FAMILY RESTAURA", "withdraw": "5.33"},
    {"description": "RCSS OXFORD #2812 LONDON", "withdraw": "20.97"},
    {"description": "BABYLON PIZZA & SHAWARMA", "withdraw": "5.99"},
    {"description": "BABYLON PIZZA & SHAWARMA", "withdraw": "68.48"},
    {"description": "APPLE.COM/BILL 866-712-7753", "withdraw": "23.72"},
    {"description": "TIM HORTONS #2406 LONDON", "withdraw": "1.67"},
    {"description": "A&W #4201 SOUTHDALE LONDON", "withdraw": "8.80"},
    {"description": "APPLE.COM/BILL 866-712-7753", "withdraw": "2.23"},
    {"description": "CASH INTEREST", "withdraw": "1.82"},
    {"description": "METRO 150 LONDON", "withdraw": "13.98"},
    {"description": "TIM HORTONS #2406 LONDON", "withdraw": "1.92"},
    {"description": "MCDONALD'S #1501 LONDON", "withdraw": "3.35"},
    {"description": "SUMMERS HOME HDWE 1560-7 LONDON", "withdraw": "3.94"},
    {"description": "TIM HORTONS #0042 LONDON", "withdraw": "1.92"},
    {"description": "TIM HORTONS #2506 LONDON", "withdraw": "1.67"},
    {"description": "RCSS OXFORD #2812 LONDON", "withdraw": "3.98"},
    {"description": "A&W #4201 SOUTHDALE LONDON", "withdraw": "8.80"},
    {"description": "GOODNESS ME NATURAL FOOD London", "withdraw": "10.25"},
    {"description": "BWW 0356 LONDON", "withdraw": "40.21"},
    {"description": "ROGERS ******8443 888-764-3771", "withdraw": "47.46"},
    {"description": "TIM HORTONS #0049 LONDON", "withdraw": "1.97"},
    {"description": "SOONIES LONDON", "withdraw": "29.23"},
    {"description": "TIM HORTONS #1764 LONDON", "withdraw": "3.25"},
    {"description": "BWW 0356 LONDON", "withdraw": "7.27"},
    {"description": "THE HOME DEPOT #7033 LONDON", "withdraw": "9.45"},
    {"description": "MCDONALD'S #29156 LONDON", "withdraw": "26.25"},
    {"description": "ZAATARZ BAKERY SWEETS A LONDON", "withdraw": "14.25"},
    {"description": "ZAATARZ BAKERY SWEETS A LONDON", "withdraw": "11.23"},
    {"description": "TIM HORTONS #0047 LONDON", "withdraw": "3.57"},
    {"description": "WALIMA MISSISSAUGA", "withdraw": "100.23"},
    {"description": "TIM HORTONS #0770 CAMBRIDGE", "withdraw": "1.68"},
    {"description": "WHOLE HEALTH NATUROPATHIC", "withdraw": "132.24"},
    {"description": "FINGERPRINTING CGL OAKVILLE", "withdraw": "64.23"},
    {"description": "FINGERPRINTING CGL OAKVILLE", "withdraw": "61.24"},
    {"description": "APPLE.COM/BILL 866-712-7753", "withdraw": "12.42"},
    {"description": "APPLE.COM/BILL TORONTO", "withdraw": "23.71"},
    {"description": "FLEETWAY BOWLING CENTR LONDON", "withdraw": "41.24"},
    {"description": "CHUCKS ROADHOUSE BAR & GR London", "withdraw": "84.21"},
    {"description": "APPLE.COM/BILL 866-712-7753", "withdraw": "1.46"},
    {"description": "BWW 0356 LONDON", "withdraw": "23.25"},
    {"description": "TIM HORTONS #0047 LONDON", "withdraw": "5.55"},
    # {"description" : "", "deposit" :  ""},
]

deposits = [
    {"description": "E-TRANSFER ***Myy", "deposit": "150.00"},
    {"description": "E-TRANSFER ***Fjv", "deposit": "80.00"},
    {"description": "E-TRANSFER ***MEY", "deposit": "250.00"},
    {"description": "E-TRANSFER ***mZs ", "deposit": "50.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "70.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "80.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "155.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "205.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "210.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "170.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "135.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "30.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "35.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "210.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "40.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "40.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "180.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "170.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "160.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "40.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "50.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "90.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "220.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "250.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "110.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "100.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "130.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "70.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "50.00"},
    {"description": f"TD ATM DEP  00{random.randint(1000,9999)}", "deposit": "80.00"},
]
