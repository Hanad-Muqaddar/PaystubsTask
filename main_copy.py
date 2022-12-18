from __future__ import print_function
from mailmerge import MailMerge
import dateparser
import os
import random
from random import randrange
import sys
import os
from docx2pdf import convert
# import comtypes.client

# PayStub Variables
###########################################################################################################################################
###########################################################################################################################################
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
        new_filename = filename + ".pdf"
        # wdFormatPDF = 17
        in_file = os.path.abspath("Output_File.docx")
        out_file = os.path.abspath("Results/"+new_filename)
        convert(in_file, out_file)
        # word = comtypes.client.CreateObject('Word.Application')
        # doc = word.Documents.Open(in_file)
        # doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        # doc.Close()
        # word.Quit()


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
            if self.Max_Val_EI == True or amount == EI_Maximum_Deduction:
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
            if self.Max_Val_CPP == True or amount == CPP_Maximum_Deduction:
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
        if Ei_tax_i == 0:
            ei_tax_mod = "00.00"
        else:
            ei_tax_mod = str(self.making_two_zer_dec(self.comma_seprated(round(Ei_tax_i, 2))))
        
        if cpp_tax_i == 0:
            cpp_tax_mod = "00.00"
        else:
            cpp_tax_mod = str(self.making_two_zer_dec(self.comma_seprated(round(cpp_tax_i, 2))))

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

            ei_tax = ei_tax_mod,
            cpp_tax = cpp_tax_mod,

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
    
    def paystub_wrapper(self, name , employee_address):
        # pay_sub_object = PayStubs() 
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
                f_period_ending_date = self.parse_and_make_date(period_ending_date)
                print("***************************")
                print("***************************")
                hours = random.randint(75,80)
                gross_total = hours * rate
                if i == 0:
                    year_to_date = self.calculate_year_to_date(rate, hours, period_ending_date)
                    last_year_to_date = year_to_date
                elif i > 0:
                    year_to_date = self.return_float(last_year_to_date) + gross_total
                    year_to_date = self.comma_seprated(year_to_date)
                    last_year_to_date = year_to_date
                y_t_date_input = self.return_float(year_to_date)
                year_to_date_incom_tax , total_percentage_for_monthly = self.total_incom_tax_calculator_year_to_date(y_t_date_input)
                year_to_date_ei = self.EI_calculator_year_to_date(y_t_date_input)
                year_to_date_cpp = self.CPP_Calculator_year_to_date(y_t_date_input)
                
                pay_date = input("Please enter pay date: ")
                print("***************************")
                print("***************************")

                income_tax = self.total_incom_tax_calculator_period(gross_total , total_percentage_for_monthly)
                Ei_tax = self.EI_calculator_Period(gross_total, year_to_date_ei)
                cpp_tax = self.CPP_Calculator_Period(gross_total, year_to_date_cpp)
                net_pay = gross_total - income_tax - Ei_tax - cpp_tax
                round_pay = round(net_pay, 2)
                f_net_pay = f"{round_pay:,}"
                # f_net_pay = pay_sub_object.jaugard_function(f_net_pay)
                
                self.making_pdf_file(name, employee_address, hours, rate, employer_name, employer_address_1, 
                    employer_address_2, gross_total, account_number, year_to_date, f_period_ending_date, pay_date, i, income_tax,
                    Ei_tax, cpp_tax, f_net_pay, year_to_date_incom_tax, year_to_date_ei, year_to_date_cpp)

#######################################################################################################################################
#######################################################################################################################################
#######################################################################################################################################
#######################################################################################################################################

class Proof_Of_SIN:
    
    def making_address(self, address):
        address_list = address.split(" ")
        middle = int(len(address_list)/2)
        address_1 = address_list[:middle]
        address_2 = address_list[middle:]
        address_1_f = " ".join(address_1)
        address_2_f = " ".join(address_2)
        return address_1_f, address_2_f

    def convert_to_pdf(self, filename):
        new_filename = filename + ".pdf"
        # wdFormatPDF = 17
        in_file = os.path.abspath("Output_File_SIN.docx")
        out_file = os.path.abspath("Results/"+new_filename)
        convert(in_file, out_file)
        # word = comtypes.client.CreateObject('Word.Application')
        # doc = word.Documents.Open(in_file)
        # doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        # doc.Close()
        # word.Quit()
    
    def making_fist_last_name(self, name):
        name_list = name.split(" ")
        first_name = name_list[:1]
        last_name = name_list[1:]
        first_name = "".join(first_name)
        last_name = " ".join(last_name)
        return first_name, last_name
    
    def making_sin(self, sin_number):
        sin_1 = sin_number[:3]
        sin_2 = sin_number[3:6]
        sin_3 = sin_number[6:]
        return sin_1, sin_2, sin_3

    def making_sin_pdf_file(self, name_i, employee_address_i, sin_number_i):
        address_1, address_2 = self.making_address(employee_address_i)
        first_name, last_name = self.making_fist_last_name(name_i)
        sin1, sin2, sin3 = self.making_sin(sin_number_i)
        template = "Proof_Of_SIN.docx"
        document = MailMerge(template)
        document.merge(sin_name = str(name_i),
        address_1_sin = str(address_1),
        address_2_sin = str(address_2),
        first_name = str(first_name),
        last_name = str(last_name),
        sin_no_1 = str(sin1),
        sin_no_2 = str(sin2),
        sin_no_3 = str(sin3), 
        )
        document.write('Output_File_SIN.docx')
        self.convert_to_pdf(f"Proof_Of_SIN")
        try:
            os.remove(os.path.abspath("Output_File_SIN.docx"))
        except:
            print("Error in Removing File.")
        return

    def SIN_Wrapper(self, name, employee_address):
        print("*************************************")
        print("*************************************")
        sin_number = str(input("Please Enter SIN Number: "))
        print("*************************************")
        print("*************************************")
        self.making_sin_pdf_file(name, employee_address, sin_number)


class TFour:

    def making_address(self, address):
        address_list = address.split(" ")
        middle = int(len(address_list)/2)
        address_1 = address_list[:middle]
        address_2 = address_list[middle:]
        address_1_f = " ".join(address_1)
        address_2_f = " ".join(address_2)
        return address_1_f, address_2_f
    
    def convert_to_pdf(self, filename):
        new_filename = filename + ".pdf"
        in_file = os.path.abspath("Output_File_T4.docx")
        out_file = os.path.abspath("Results/" + new_filename)
        convert(in_file, out_file)
    
    def making_sin(self, sin_number):
        sin_1 = sin_number[:3]
        sin_2 = sin_number[3:6]
        sin_3 = sin_number[6:]
        return sin_1, sin_2, sin_3
    
    def making_fist_last_name(self, name):
        name_list = name.split(" ")
        first_name = name_list[:1]
        last_name = name_list[1:]
        first_name = "".join(first_name)
        last_name = " ".join(last_name)
        return first_name, last_name
    
    def comma_seprated(self, number):
        return f"{number:,}"
    
    def breaking_number(self, num):
        a = str(num).split(".")
        if len(a) == 1:
            before_point = a[0]
            after_point = "00"
            return before_point, after_point
        before_point = a[0]
        after_point = a[-1][:2]
        return before_point, after_point
    
    def percentage(self, percent, whole):
        return (percent * whole) / 100.0
    
    def EI_calculator_year_to_date(self, y_t_d_pay, EI_Rate, EI_Maximum_Deduction):
        if self.percentage(EI_Rate, y_t_d_pay) >= EI_Maximum_Deduction:
            return EI_Maximum_Deduction
        else:
            amount =  self.percentage(EI_Rate, y_t_d_pay)
            return amount
    
    def CPP_Calculator_year_to_date(self, y_t_d_pay, CPP_Rate, CPP_Maximum_Deduction ):
        if self.percentage(CPP_Rate, y_t_d_pay) >= CPP_Maximum_Deduction:
            return CPP_Maximum_Deduction
        else:
            amount =  self.percentage(CPP_Rate, y_t_d_pay)
            return amount
    
    def getting_eirate_and_max_deductions(self, year):
        from Constants import EI_Rates
        for i in EI_Rates:
            if i['year'] == int(year):
                return i['ei_rate'], i['max_deduction']
    
    def getting_cpprate_and_max_deductions(slef, year):
        from Constants import CPP_Rates
        for i in CPP_Rates:
            if i['year'] == int(year):
                return i['cpp_rate'], i['max_deduction'] 
    
    def getting_maximum_EI_insurable_amount(self, salary, year):
        from Constants import Max_EI_Insureable_Income
        year_max_income = 0
        for i in Max_EI_Insureable_Income:
            if i['year'] == int(year):
                year_max_income = i['income']
        if salary > year_max_income:
            return year_max_income
        elif salary <= year_max_income:
            return salary


    def getting_maximum_CPP_insurable_amount(self, salary, year):
        from Constants import Max_CPP_Pensionable_Income
        year_max_income = 0
        for i in Max_CPP_Pensionable_Income:
            if i['year'] == int(year):
                year_max_income = i['income']
        if salary > year_max_income:
            return year_max_income
        elif salary <= year_max_income:
            return salary


    def making_t4_pdf_file(self, name_i, employee_address_i, sin_number_i, t4_year_input_i, i_i, gross_salary_i ):
        template = "T4_2021_Creations.docx"
        document = MailMerge(template)
        ##############################
        address_1, address_2 = self.making_address(employee_address_i)
        first_name, last_name = self.making_fist_last_name(name_i)
        sin1, sin2, sin3 = self.making_sin(sin_number_i)
        year_number = str(t4_year_input_i)[2:]

        paystub_object = PayStubs()
        income_tax, percntage = paystub_object.total_incom_tax_calculator_year_to_date(gross_salary_i)

        rate, max = self.getting_eirate_and_max_deductions(t4_year_input_i)
        t4_EI = self.EI_calculator_year_to_date(gross_salary_i, rate, max)

        cpp_rate, cpp_max = self.getting_cpprate_and_max_deductions(t4_year_input_i)
        t4_CPP = self.CPP_Calculator_year_to_date(gross_salary_i, cpp_rate, cpp_max)
        
        maximum_EI_amount = self.getting_maximum_EI_insurable_amount(gross_salary_i, t4_year_input_i)

        maximum_cpp_amount = self.getting_maximum_CPP_insurable_amount(gross_salary_i, t4_year_input_i)
        # befor_point_sal, after_point_sal = self.breaking_number(gross_salary_i)

        document.merge(
            t4_name_1 = str(name_i),
            t4_name_2 = str(name_i),
            ###################
            t4_year = str(t4_year_input_i),
            t4_year_1 = str(t4_year_input_i),
            ###################
            t4_address_1_1 = str(address_1),
            t4_address_1_2 = str(address_1),
            ###################
            t4_address_2_1 = str(address_2),
            t4_address_2_2 = str(address_2),
            ###################
            f_nm_1 = str(first_name),
            f_nm_2 = str(first_name),
            ###################
            l_nm_1 = str(last_name),
            l_nm_2 = str(last_name),
            ###################
            sn1 = str(sin1),
            sn2 = str(sin2),
            sn3 = str(sin3),
            sn4 = str(sin1),
            s5 = str(sin2),
            s6 = str(sin3),  
            ###################
            yn = str(year_number),
            y2 = str(year_number),
            ###################
            gross_sal = str(self.comma_seprated(round(gross_salary_i, 2))),
            inc_tax = str(self.comma_seprated(round(income_tax, 2))),
            t4_ei = str(self.comma_seprated(round(t4_EI, 2))),
            t4_cpp = str(self.comma_seprated(round(t4_CPP, 2))),
            max_ei = str(self.comma_seprated(round(maximum_EI_amount, 2))),
            max_cpp = str(self.comma_seprated(round(maximum_cpp_amount, 2))),
            ###################
            gs_sal_1 = str(self.comma_seprated(round(gross_salary_i, 2))),
            inc_tx_1 = str(self.comma_seprated(round(income_tax, 2))),
            t4_ei_1 = str(self.comma_seprated(round(t4_EI, 2))),
            t4_cp_1 = str(self.comma_seprated(round(t4_CPP, 2))),
            max_ei1 = str(self.comma_seprated(round(maximum_EI_amount, 2))),
            max_cpp1 = str(self.comma_seprated(round(maximum_cpp_amount, 2))),
            
        )
        document.write(f'Results/Output_File_T4_{i_i}.docx')
        # self.convert_to_pdf(f"T4_Document_{i_i}")
        # try:
        #     os.remove(os.path.abspath("Output_File_T4.docx"))
        # except:
        #     print("Error in Removing File.")
        # return
    
    def T4_Wrapper(self, name, employee_address):
        print("*************************************")
        print("*************************************")
        number_of_documents = int(input("Please enter How many documents you want to create: "))
        print("*************************************")
        print("*************************************")
        if number_of_documents == 0:
            print("You have Enter 0, We are not creating any document. Thanks")
        else:
            for i in range(number_of_documents):
                t4_year_input = input("Please Enter Year for T4: ")
                print("*************************************")
                print("*************************************")
                sin_number = str(input("Please Enter SIN Number: "))
                print("*************************************")
                print("*************************************")
                gross_salary = float(input("Please enter your gross Salary: "))
                print("*************************************")
                print("*************************************")
                self.making_t4_pdf_file(name, employee_address, sin_number, t4_year_input, i, gross_salary)
    

#######################################################################################################################################
#######################################################################################################################################
#######################################################################################################################################
#######################################################################################################################################

if __name__ == '__main__':

    print("***************************")
    print("***************************")
    name = input('Please enter Employee name: ')
    print("***************************")
    print("***************************")
    employee_address = input("Please enter Employee address: ")

    print("***************************")
    print("***************************")
    document_type = input(''' Please Enter Which Document You want to Create, Select Options 
                         1 ) PayStub 
                         2 ) Proof Of SIN 
                         3 ) PayStub and Proof Of SIN  
                         4 ) T4 Document
                    ''')
    print("***************************")
    print("***************************")
    
    if int(document_type) == 1:
        pay_sub_object = PayStubs()
        pay_sub_object.paystub_wrapper(name, employee_address)
    elif int(document_type) == 2:
        sin_object = Proof_Of_SIN()
        sin_object.SIN_Wrapper(name, employee_address)
    elif int(document_type) == 3:
        pay_sub_object = PayStubs()
        sin_object = Proof_Of_SIN()
        pay_sub_object.paystub_wrapper(name, employee_address)
        sin_object.SIN_Wrapper(name, employee_address)
    elif int(document_type) == 4:
        TFour_object = TFour()
        TFour_object.T4_Wrapper(name, employee_address)

    