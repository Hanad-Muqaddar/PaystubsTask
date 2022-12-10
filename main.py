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

EI_Rate = 1.58
CPP_Rate = 5.70
EI_Maximum_Deduction = 952.74
CPP_Maximum_Deduction = 3499.80
last_year_to_date = 0
############################################################################################################################################
############################################################################################################################################

class PayStubs:

    Old_ei_Calculation = 0
    Max_Val_EI = False
    Old_cpp_Calculation = 0
    Max_Val_CPP = False


    def calculate_year_to_date(self, rate, hours, period_date):
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

    def parse_and_make_date(self, date):
        am = dateparser.parse(date)
        month = am.strftime('%m')
        day = am.strftime('%d')
        year = am.strftime('%y')
        date_format = f"{day}/{month}/20{year}"
        return date_format

    def percentage(self, percent, whole):
        return (percent * whole) / 100.0


    def convert_to_pdf(self, filename):
        wdFormatPDF = 17
        in_file = os.path.abspath("Output_File.docx")
        out_file = os.path.abspath(filename)
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()


    def federal_income_tax_calculator(self, gross_pay):
        if gross_pay < 50197:
            amount = self.percentage(federal_first, gross_pay)
            return amount , federal_first
        elif gross_pay > 50197 and gross_pay < 100392:
            amount = self.percentage(federal_second, gross_pay)
            return amount , federal_second
        elif gross_pay > 100392 and gross_pay < 155625:
            amount = self.percentage(federal_three, gross_pay)
            return amount , federal_three
        elif gross_pay > 155625 and gross_pay < 221708:
            amount = self.percentage(federal_four, gross_pay)
            return amount , federal_four
        elif gross_pay > 221708:
            amount = self.percentage(federal_five, gross_pay)
            return amount , federal_five


    def province_income_tax_calculator(self, gross_pay):
        if gross_pay < 46226:
            amount = self.percentage(province_first, gross_pay)
            return amount , province_first
        elif gross_pay >= 46227 and gross_pay <= 92454:
            amount = self.percentage(province_second, gross_pay)
            return amount , province_second
        elif gross_pay >= 92455 and gross_pay <= 150000:
            amount = self.percentage(province_three, gross_pay)
            return amount , province_three
        elif gross_pay >= 150001 and gross_pay <= 220000:
            amount = self.percentage(province_four, gross_pay)
            return amount , province_four
        elif gross_pay > 220000:
            amount = self.percentage(province_five, gross_pay)
            return amount , province_five
    
    def federal_income_tax_calculator_y_t_d(self, gross_pay):
        if gross_pay < 50197:
            amount = self.percentage(federal_first, gross_pay)
            return amount , federal_first
        elif gross_pay > 50197 and gross_pay < 100392:
            first_amount = 50197
            second_amount = gross_pay - 50197
            first_amount_cal = self.percentage(federal_first, first_amount)
            second_amount_cal = self.percentage(federal_second, second_amount)
            total_amount = first_amount_cal + second_amount_cal
            return total_amount , federal_second
        elif gross_pay > 100392 and gross_pay < 155625:
            first_amount = 50197
            second_amount = 50195
            third_amount = gross_pay - 100392
            first_amount_cal = self.percentage(federal_first, first_amount)
            second_amount_cal = self.percentage(federal_second, second_amount)
            third_amount_cal = self.percentage(federal_three, third_amount)
            total_amount = first_amount_cal + second_amount_cal + third_amount_cal
            return total_amount , federal_three
        elif gross_pay > 155625 and gross_pay < 221708:
            first_amount = 50197
            second_amount = 50195
            third_amount = 55233
            fourth_amount = gross_pay - 155625
            first_amount_cal = self.percentage(federal_first, first_amount)
            second_amount_cal = self.percentage(federal_second, second_amount)
            third_amount_cal = self.percentage(federal_three, third_amount)
            fourth_amount_cal = self.percentage(federal_four, fourth_amount)
            total_amount = first_amount_cal + second_amount_cal + third_amount_cal + fourth_amount_cal
            return total_amount , federal_four
        elif gross_pay > 221708:
            first_amount = 50197
            second_amount = 50195
            third_amount = 55233
            fourth_amount = 66083
            fifth_amount = gross_pay - 221708
            first_amount_cal = self.percentage(federal_first, first_amount)
            second_amount_cal = self.percentage(federal_second, second_amount)
            third_amount_cal = self.percentage(federal_three, third_amount)
            fourth_amount_cal = self.percentage(federal_four, fourth_amount)
            fifth_amount_cal = self.percentage(federal_five, fifth_amount)
            total_amount = first_amount_cal + second_amount_cal + third_amount_cal + fourth_amount_cal + fifth_amount_cal
            return total_amount , federal_five
            
    def province_income_tax_calculator_y_t_d(self, gross_pay):
        if gross_pay < 46226:
            amount = self.percentage(province_first, gross_pay)
            return amount , province_first
        elif gross_pay >= 46227 and gross_pay <= 92454:
            first_amount = 46226
            second_amount = gross_pay - first_amount
            first_amount_cal = self.percentage(province_first, first_amount)
            second_amount_cal = self.percentage(province_second, second_amount)
            total_amount = first_amount_cal + second_amount_cal
            return total_amount , province_second
        elif gross_pay >= 92455 and gross_pay <= 150000:
            first_amount = 46226
            second_amount = 46229
            third_amount = gross_pay - 92455
            first_amount_cal = self.percentage(province_first, first_amount)
            second_amount_cal = self.percentage(province_second, second_amount)
            third_amount_cal = self.percentage(province_three, third_amount)
            total_amount = first_amount_cal + second_amount_cal + third_amount_cal
            return total_amount , province_three
        elif gross_pay >= 150001 and gross_pay <= 220000:
            first_amount = 46226
            second_amount = 46229
            third_amount = 57546
            fourth_amount = gross_pay - 150001
            first_amount_cal = self.percentage(province_first, first_amount)
            second_amount_cal = self.percentage(province_second, second_amount)
            third_amount_cal = self.percentage(province_three, third_amount)
            fourth_amount_cal = self.percentage(province_four , fourth_amount)
            total_amount =  first_amount_cal + second_amount_cal + third_amount_cal + fourth_amount_cal
            return total_amount , province_four
        elif gross_pay > 220000:
            first_amount = 46226
            second_amount = 46229
            third_amount = 57546
            fourth_amount = 69999
            fifth_amount = gross_pay - 220000
            first_amount_cal = self.percentage(province_first, first_amount)
            second_amount_cal = self.percentage(province_second, second_amount)
            third_amount_cal = self.percentage(province_three, third_amount)
            fourth_amount_cal = self.percentage(province_four , fourth_amount)
            fifth_amount_cal = self.percentage(province_five, fifth_amount)
            total_amount =  first_amount_cal + second_amount_cal + third_amount_cal + fourth_amount_cal + fifth_amount_cal
            return total_amount , province_five

    def total_incom_tax_calculator_period(self, gross_pay, total_percentage_for_monthly):
        total_income_tax = self.percentage(total_percentage_for_monthly, gross_pay)
        return total_income_tax

    def total_incom_tax_calculator_year_to_date(self, y_to_d):
        fed_in_tax, percentage_fed = self.federal_income_tax_calculator_y_t_d(y_to_d)
        prov_in_tax, percentage_prov = self.province_income_tax_calculator_y_t_d(y_to_d)
        total_income_tax = fed_in_tax + prov_in_tax
        total_percentage_for_monthly =  percentage_fed + percentage_prov
        return total_income_tax , total_percentage_for_monthly

    def EI_calculator_year_to_date(self, y_t_d_pay):
        if self.percentage(EI_Rate, y_t_d_pay) >= EI_Maximum_Deduction:
            return EI_Maximum_Deduction
        else:
            amount =  self.percentage(EI_Rate, y_t_d_pay)
            return amount

    def CPP_Calculator_year_to_date(self, y_t_d_pay):
        if self.percentage(CPP_Rate, y_t_d_pay) >= CPP_Maximum_Deduction:
            return CPP_Maximum_Deduction
        else:
            amount =  self.percentage(CPP_Rate, y_t_d_pay)
            return amount

    def jaugard_function(self, pay):
        netpay = pay.split(".")
        mod = int(netpay[-1]) - 1
        netpay[-1] = str(mod)
        return ".".join(netpay)

    def EI_calculator_Period(self, gross_total, Ei_calculator_y_t_d):
        if Ei_calculator_y_t_d >= EI_Maximum_Deduction:
            amount = EI_Maximum_Deduction - self.Old_ei_Calculation
            if self.Max_Val_EI == True:
                return 0
            else:
                self.Max_Val_EI = True    
                return amount
        else:
            self.Old_ei_Calculation = Ei_calculator_y_t_d
            amount =  self.percentage(EI_Rate, gross_total)
            return amount

    def CPP_Calculator_Period(self, gross_total, CPP_Calculator_y_t_d):
        if CPP_Calculator_y_t_d >= CPP_Maximum_Deduction:
            amount = CPP_Maximum_Deduction - self.Old_cpp_Calculation
            if self.Max_Val_CPP == True:
                return 0
            else:
                self.Max_Val_CPP = True    
                return amount
        else:
            self.Old_cpp_Calculation = CPP_Calculator_y_t_d
            amount = self.percentage(CPP_Rate, gross_total)
            return amount

    def making_two_zer_dec(self, num):
        a = num.split(".")
        if len(a[-1]) > 1:
            return ".".join(a)
        elif len(a[-1]) == 1:
            a[-1] = a[-1] + "0"
            return ".".join(a)

    def making_two_zer_dec1(self, num):
        ret_num = format(num, ".2f")
        return ret_num

    def return_float(self, number):
        d = number.replace(",","")
        return float(d)

    def comma_seprated(self, number):
        return f"{number:,}"

    def making_pdf_file(self, name_i, employee_address_i, hours_i, rate_i, employer_name_i, employer_address_1_i, 
                                    employer_address_2_i, g_total_i, account_number_i, year_to_date,period_ending_date_i, pay_date_i, i,
                                    income_tax_i, Ei_tax_i, cpp_tax_i, net_pay_i, year_to_date_incom_tax_i,year_to_date_ei_i,year_to_date_cpp_i):
        template = "Hanad-ADP-PAYSTUBS.docx"
        document = MailMerge(template)
        # print(document.get_merge_fields())
        document.merge(employee_name_1=name_i, 
            emp_2 = name_i, 
            hours = str(self.making_two_zer_dec1(hours_i)), 
            rate = str(self.making_two_zer_dec1(rate_i)),
            total = str(self.comma_seprated((round(g_total_i, 2)))),
            gp_total = str(self.comma_seprated((round(g_total_i,2)))),
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
            inc_tax = str(self.making_two_zer_dec(self.comma_seprated(round(income_tax_i, 2)))),
            ei_tax = str(self.making_two_zer_dec(self.comma_seprated(round(Ei_tax_i, 2)))),
            cpp_tax = str(self.making_two_zer_dec(self.comma_seprated(round(cpp_tax_i, 2)))),
            net_pay_1 = str(self.making_two_zer_dec(net_pay_i)),
            net_pay_2 = str(self.making_two_zer_dec(net_pay_i)),
            net_pay_3 = str(self.making_two_zer_dec(net_pay_i)),
            y_t_d_it =  str(self.making_two_zer_dec(self.comma_seprated(round(year_to_date_incom_tax_i, 2)))),
            y_t_d_ei = str(self.making_two_zer_dec(self.comma_seprated(round(year_to_date_ei_i, 2)))),
            y_t_d_cpp = str(self.making_two_zer_dec(self.comma_seprated(round(year_to_date_cpp_i, 2)))),
            pay_date = str(pay_date_i),
            pay_date_2 = str(pay_date_i),
        )
        document.write('Output_File.docx')
        self.convert_to_pdf(f"PDF-ADP-PAYSTUB_{i}")
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

    pay_sub_object = PayStubs() 
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
            f_period_ending_date = pay_sub_object.parse_and_make_date(period_ending_date)
            print("***************************")
            print("***************************")
            hours = random.randint(75,80)
            gross_total = hours * rate
            if i == 0:
                year_to_date = pay_sub_object.calculate_year_to_date(rate, hours, period_ending_date)
                last_year_to_date = year_to_date
            elif i > 0:
                year_to_date = pay_sub_object.return_float(last_year_to_date) + gross_total
                year_to_date = pay_sub_object.comma_seprated(year_to_date)
                last_year_to_date = year_to_date
            y_t_date_input = pay_sub_object.return_float(year_to_date)
            year_to_date_incom_tax , total_percentage_for_monthly = pay_sub_object.total_incom_tax_calculator_year_to_date(y_t_date_input)
            year_to_date_ei = pay_sub_object.EI_calculator_year_to_date(y_t_date_input)
            year_to_date_cpp = pay_sub_object.CPP_Calculator_year_to_date(y_t_date_input)
            
            pay_date = input("Please enter pay date: ")
            print("***************************")
            print("***************************")

            income_tax = pay_sub_object.total_incom_tax_calculator_period(gross_total , total_percentage_for_monthly)
            Ei_tax = pay_sub_object.EI_calculator_Period(gross_total, year_to_date_ei)
            cpp_tax = pay_sub_object.CPP_Calculator_Period(gross_total, year_to_date_cpp)
            net_pay = gross_total - income_tax - Ei_tax - cpp_tax
            round_pay = round(net_pay, 2)
            f_net_pay = f"{round_pay:,}"
            # f_net_pay = pay_sub_object.jaugard_function(f_net_pay)
            
            pay_sub_object.making_pdf_file(name, employee_address, hours, rate, employer_name, employer_address_1, 
                employer_address_2, gross_total, account_number, year_to_date, f_period_ending_date, pay_date, i, income_tax,
                 Ei_tax, cpp_tax, f_net_pay, year_to_date_incom_tax, year_to_date_ei, year_to_date_cpp)

