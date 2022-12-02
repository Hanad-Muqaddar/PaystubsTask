from __future__ import print_function
from mailmerge import MailMerge
import dateparser
import os
import random
from random import randrange
import sys
import os
import comtypes.client

###########################################################################################################################################
###########################################################################################################################################
federal_first = 15
province_first = 5.05
federal_second = 20.5
province_second = 9.15
federal_three = 26
province_three = 11.16
federal_four = 29
province_four = 12.16
federal_five = 33
province_five = 13.16

EI_Rate = 952.74
CPP_Rate = 3499.80
last_year_to_date = 0
############################################################################################################################################
############################################################################################################################################


def calculate_year_to_date(rate, hours, period_date):
    d = dateparser.parse(period_date)
    current_month = d.strftime('%m')
    current_day = d.strftime('%d')
    if int(current_day) > 15:
        periods = int(current_month) * 2
    else:
        periods = (int(current_month) - 1) * 2 + 1
    
    year_t_date = rate * periods * hours
    y_t_d = f"{year_t_date:,}"
    return y_t_d

def parse_and_make_date(date):
    am = dateparser.parse(date)
    month = am.strftime('%m')
    day = am.strftime('%d')
    year = am.strftime('%y')
    date_format = f"{day}/{month}/20{year}"
    return date_format

def percentage(percent, whole):
  return (percent * whole) / 100.0


def convert_to_pdf(filename):
    wdFormatPDF = 17
    in_file = os.path.abspath("Output_File.docx")
    out_file = os.path.abspath(filename)
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()


def federal_income_tax_calculator(gross_pay):
    if gross_pay < 50197:
        amount = percentage(federal_first, gross_pay)
        return amount
    elif gross_pay > 50197 and gross_pay < 100392:
        amount = percentage(federal_second, gross_pay)
        return amount
    elif gross_pay > 100392 and gross_pay < 155625:
        amount = percentage(federal_three, gross_pay)
        return amount
    elif gross_pay > 155625 and gross_pay < 221708:
        amount = percentage(federal_four, gross_pay)
        return amount
    elif gross_pay > 221708:
        amount = percentage(federal_five, gross_pay)
        return amount


def province_income_tax_calculator(gross_pay):
    if gross_pay < 46226:
        amount = percentage(province_first, gross_pay)
        return amount
    elif gross_pay >= 46227 and gross_pay <= 92454:
        amount = percentage(province_second, gross_pay)
        return amount
    elif gross_pay >= 92455 and gross_pay <= 150000:
        amount = percentage(province_three, gross_pay)
        return amount
    elif gross_pay >= 150001 and gross_pay <= 220000:
        amount = percentage(province_four, gross_pay)
        return amount
    elif gross_pay > 220000:
        amount = percentage(province_five, gross_pay)
        return amount

def total_incom_tax_calculator_period(gross_pay):
    fed_in_tax = federal_income_tax_calculator(gross_pay)
    prov_in_tax = province_income_tax_calculator(gross_pay)
    total_income_tax = fed_in_tax + prov_in_tax
    return total_income_tax

def total_incom_tax_calculator_year_to_date(y_to_d):
    fed_in_tax = federal_income_tax_calculator(y_to_d)
    prov_in_tax = province_income_tax_calculator(y_to_d)
    total_income_tax = fed_in_tax + prov_in_tax
    return total_income_tax

def EI_calculator_year_to_date(period_date):
    d = dateparser.parse(period_date)
    current_month = d.strftime('%m')
    periods = int(current_month) * 2
    amount =  EI_Rate / 24
    final_amount = amount * periods
    return final_amount

def CPP_Calculator_year_to_date(period_date):
    d = dateparser.parse(period_date)
    current_month = d.strftime('%m')
    periods = int(current_month) * 2
    amount =  CPP_Rate / 24
    final_amount = amount * periods
    return final_amount

def jaugard_function(pay):
    netpay = pay.split(".")
    mod = int(netpay[-1]) - 1
    netpay[-1] = str(mod)
    return ".".join(netpay)

def EI_calculator_Period():
    amount =  EI_Rate / 24
    return amount

def CPP_Calculator_Period():
    amount = CPP_Rate / 24
    return amount

def making_two_zer_dec(num):
    a = num.split(".")
    if len(a[-1]) > 1:
        return ".".join(a)
    elif len(a[-1]) == 1:
        a[-1] = a[-1] + "0"
        return ".".join(a)

def making_two_zer_dec1(num):
    ret_num = format(num, ".2f")
    return ret_num

def return_float(number):
    d = number.replace(",","")
    return float(d)

def comma_seprated(number):
    return f"{number:,}"

def making_pdf_file(name_i, employee_address_i, hours_i, rate_i, employer_name_i, employer_address_1_i, 
                                employer_address_2_i, g_total_i, account_number_i, year_to_date,period_ending_date_i, pay_date_i, i,
                                income_tax_i, Ei_tax_i, cpp_tax_i, net_pay_i, year_to_date_incom_tax_i,year_to_date_ei_i,year_to_date_cpp_i):
    template = "Hanad-ADP-PAYSTUBS.docx"
    document = MailMerge(template)
    # print(document.get_merge_fields())
    document.merge(employee_name_1=name_i, 
        emp_2 = name_i, 
        hours = str(making_two_zer_dec1(hours_i)), 
        rate = str(making_two_zer_dec1(rate_i)),
        total = str(comma_seprated((round(g_total_i, 2)))),
        gp_total = str(comma_seprated((round(g_total_i,2)))),
        employee_addr = str(employee_address_i),
        employer_name = str(employer_name_i),
        employer_addr_1 = str(employer_address_1_i),
        employer_addr_2 = str(employer_address_2_i),
        employer_name_1 = str(employer_name_i),
        employer_addr_1_1 = str(employer_address_1_i),
        employer_addr_2_2 = str(employer_address_2_i),
        acn = str(account_number_i),
        y_to_d_1 = str(year_to_date),
        y_to_d_2 = str(year_to_date),
        p_end_date = str(period_ending_date_i),
        inc_tax = str(making_two_zer_dec(comma_seprated(round(income_tax_i, 2)))),
        ei_tax = str(making_two_zer_dec(comma_seprated(round(Ei_tax_i, 2)))),
        cpp_tax = str(making_two_zer_dec(comma_seprated(round(cpp_tax_i, 2)))),
        net_pay_1 = str(making_two_zer_dec(net_pay_i)),
        net_pay_2 = str(making_two_zer_dec(net_pay_i)),
        net_pay_3 = str(making_two_zer_dec(net_pay_i)),
        y_t_d_it =  str(making_two_zer_dec(comma_seprated(round(year_to_date_incom_tax_i, 2)))),
        y_t_d_ei = str(making_two_zer_dec(comma_seprated(round(year_to_date_ei_i, 2)))),
        y_t_d_cpp = str(making_two_zer_dec(comma_seprated(round(year_to_date_cpp_i, 2)))),
        pay_date = str(pay_date_i),
        pay_date_2 = str(pay_date_i),
    )
    document.write('Output_File.docx')
    convert_to_pdf(f"PDF-ADP-PAYSTUB_{i}")
    try:
        os.remove(os.path.abspath("Output_File.docx"))
    except:
        print("Error in Removing File.")
    return


if __name__ == '__main__':
    print("***************************")
    print("***************************")
    name = input('Please enter Employee name: ')
    print("***************************")
    print("***************************")
    employee_address = input("Please enter Employee address: ")
    print("***************************")
    print("***************************")
    employer_name = input("Please enter Employer name: ")
    print("***************************")
    print("***************************")
    employer_address_1 = input("Please enter Employer Address 1: ")
    print("***************************")
    print("***************************")
    employer_address_2 = input("Please enter Employer Adress 2: ")
    rate = float(input("Please enter the rate which you decided: "))
    print("***************************")
    print("***************************")
    
    account_number = randrange(1000, 9999)

    number_of_pay_stubs = input("Please enter, how many number of paystubs you want to create: ")
    if int(number_of_pay_stubs) == 0:
        print("You have enterd 0. So i am not creating any paystub. Thanks")
        sys.exit()
    elif int(number_of_pay_stubs) > 0:
        for i in range(int(number_of_pay_stubs)):
            period_ending_date = input("Please Enter Period for Paystub: ")
            f_period_ending_date = parse_and_make_date(period_ending_date)
            print("***************************")
            print("***************************")
            hours = random.randint(75,80)
            gross_total = hours * rate
            if i == 0:
                year_to_date = calculate_year_to_date(rate, hours, period_ending_date)
                last_year_to_date = year_to_date
            elif i > 0:
                year_to_date = return_float(last_year_to_date) + gross_total
                year_to_date = comma_seprated(year_to_date)
                last_year_to_date = year_to_date

            y_t_date_input = return_float(year_to_date)
            year_to_date_incom_tax = total_incom_tax_calculator_year_to_date(y_t_date_input)
            year_to_date_ei = EI_calculator_year_to_date(period_ending_date)
            year_to_date_cpp = CPP_Calculator_year_to_date(period_ending_date)
            
            pay_date = input("Please enter pay date: ")
            print("***************************")
            print("***************************")

            income_tax = total_incom_tax_calculator_period(gross_total)
            Ei_tax = EI_calculator_Period()
            cpp_tax = CPP_Calculator_Period()
            net_pay = gross_total - income_tax - Ei_tax - cpp_tax
            round_pay = round(net_pay, 2)
            f_net_pay = f"{round_pay:,}"
            f_net_pay = jaugard_function(f_net_pay)
            
            making_pdf_file(name, employee_address, hours, rate, employer_name, employer_address_1, 
                employer_address_2, gross_total, account_number, year_to_date, f_period_ending_date, pay_date, i, income_tax,
                 Ei_tax, cpp_tax, f_net_pay, year_to_date_incom_tax, year_to_date_ei, year_to_date_cpp)

