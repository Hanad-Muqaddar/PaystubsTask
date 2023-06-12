from mailmerge import MailMerge
import dateparser
import os
import random
from random import randrange
import sys
import os
from docx2pdf import convert
from Constants import withdrawls
from Constants import deposits
from Constants import values_for_paystub
from datetime import datetime, timedelta
import pandas as pd
import json

# Variables For Paystub
###########################################################################################################################################
###########################################################################################################################################
###########################################################################################################################################
###########################################################################################################################################

global_testing_var = 0


federal_first = ""
province_first = ""
federal_second = ""
province_second = ""
federal_three = ""
province_three = ""
federal_four = ""
province_four = ""
federal_five = ""
province_five = ""
EI_Rate = ""
CPP_Rate = ""
EI_Maximum_Deduction = ""
CPP_Maximum_Deduction = ""


# Global Variables
###########################################################################################################################################
###########################################################################################################################################
###########################################################################################################################################
###########################################################################################################################################

last_year_to_date = 0

global_name = ""
global_employee_address = ""
global_employer_name = ""
global_employer_address_1 = ""
global_employer_address_2 = ""
global_sin_number = ""


def making_address(address):
    address_list = address.split(" ")
    middle = int(len(address_list) / 2)
    address_1 = address_list[:middle]
    address_2 = address_list[middle:]
    address_1_f = " ".join(address_1)
    address_2_f = " ".join(address_2)
    return address_1_f, address_2_f


def making_folder(folder_name):
    current_directory = os.getcwd()
    folder_path = f"{current_directory}\\Results\\{folder_name}"
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    return


# New Feature
###########################################################################################################################################
###########################################################################################################################################
###########################################################################################################################################
###########################################################################################################################################


def options_feature(name, employee_address, row_i):
    selected_option = input(
        """Do you want to create Selected documents (1,2,3), Exit : Select Options,
                    1 ) Selected
                    2 ) Exit    
                        """
    )
    if int(selected_option) == 1:
        if pd.isna(row_i["doc_options"]):
            document_type = input(
                """ Please Enter Which Document You want to Create, Select Options like this : (1,2,3)
                                1 ) PayStub 
                                2 ) Proof Of SIN 
                                3 ) T4 Document
                                4 ) Proof Of Enrollment
                                5 ) TD Document
                                6 ) ALL
                                """
            )
            int_numbers = list(map(int, document_type.split(",")))
        else:
            int_split = row_i["doc_options"].split(",")
            int_numbers = list(map(int, int_split))
        for number in int_numbers:
            if number == 1:
                slection_number = input(
                    """Which Document You want to Create 
                a) Paystub 1
                b) Paystub 2
                c) Paystub 3
                """
                )
                if str(slection_number) == "a":
                    pay_sub_object = PayStubs()
                    pay_sub_object.paystub_wrapper(name, employee_address, row_i)
                elif str(slection_number) == "b":
                    child_obj = PayStubChild()
                    child_obj.paystub_child_wrapper(name, employee_address)
                elif str(slection_number) == "c":
                    pay_stub_three_obj = PayStubChildONE()
                    pay_stub_three_obj.paystub_child_wrapper(name, employee_address)
            elif number == 2:
                sin_object = Proof_Of_SIN()
                sin_object.SIN_Wrapper(name, employee_address)
            elif number == 3:
                TFour_object = TFour()
                TFour_object.T4_Wrapper(name, employee_address)
            elif number == 4:
                poof_of_enrl = Proof_Of_Enrollment()
                poof_of_enrl.pof_wrapper(name)
            elif number == 5:
                td_clas = TD_Document()
                td_clas.TD_wrapper(name, employee_address)
            elif number == 6:
                pay_sub_object = PayStubs()
                sin_object = Proof_Of_SIN()
                TFour_object = TFour()
                poof_of_enrl = Proof_Of_Enrollment()
                td_clas = TD_Document()

                pay_sub_object.paystub_wrapper(name, employee_address)
                sin_object.SIN_Wrapper(name, employee_address)
                TFour_object.T4_Wrapper(name, employee_address)
                poof_of_enrl.pof_wrapper(
                    name,
                )
                td_clas.TD_wrapper(name, employee_address)

    if int(selected_option) == 2:
        import sys

        print("Thanks for using this. GoodBye")
        sys.exit()


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
        current_month = d.strftime("%m")
        current_day = d.strftime("%d")
        if int(current_day) > 15:
            periods = int(current_month) * 2
        else:
            periods = (int(current_month) - 1) * 2 + 1

        year_t_date = rate * periods * hours
        y_t_d = f"{year_t_date:,}"
        return y_t_d

    def parse_and_make_date(self, date):
        am = dateparser.parse(date)
        month = am.strftime("%m")
        day = am.strftime("%d")
        year = am.strftime("%y")
        date_format = f"{day}/{month}/20{year}"
        return date_format

    def percentage(self, percent, whole):
        return (percent * whole) / 100.0

    def convert_to_pdf(self, filename):
        new_filename = filename + ".pdf"
        in_file = os.path.abspath("Output_File.docx")
        making_folder(global_name)
        out_file = os.path.abspath(f"Results/{global_name}/" + new_filename)
        convert(in_file, out_file)

    def federal_income_tax_calculator(self, gross_pay):
        if gross_pay < 50197:
            amount = self.percentage(federal_first, gross_pay)
            return amount, federal_first
        elif gross_pay > 50197 and gross_pay < 100392:
            amount = self.percentage(federal_second, gross_pay)
            return amount, federal_second
        elif gross_pay > 100392 and gross_pay < 155625:
            amount = self.percentage(federal_three, gross_pay)
            return amount, federal_three
        elif gross_pay > 155625 and gross_pay < 221708:
            amount = self.percentage(federal_four, gross_pay)
            return amount, federal_four
        elif gross_pay > 221708:
            amount = self.percentage(federal_five, gross_pay)
            return amount, federal_five

    def province_income_tax_calculator(self, gross_pay):
        if gross_pay < 46226:
            amount = self.percentage(province_first, gross_pay)
            return amount, province_first
        elif gross_pay >= 46227 and gross_pay <= 92454:
            amount = self.percentage(province_second, gross_pay)
            return amount, province_second
        elif gross_pay >= 92455 and gross_pay <= 150000:
            amount = self.percentage(province_three, gross_pay)
            return amount, province_three
        elif gross_pay >= 150001 and gross_pay <= 220000:
            amount = self.percentage(province_four, gross_pay)
            return amount, province_four
        elif gross_pay > 220000:
            amount = self.percentage(province_five, gross_pay)
            return amount, province_five

    def federal_income_tax_calculator_y_t_d(self, gross_pay):
        if gross_pay < 50197:
            amount = self.percentage(federal_first, gross_pay)
            return amount, federal_first
        elif gross_pay > 50197 and gross_pay < 100392:
            first_amount = 50197
            second_amount = gross_pay - 50197
            first_amount_cal = self.percentage(federal_first, first_amount)
            second_amount_cal = self.percentage(federal_second, second_amount)
            total_amount = first_amount_cal + second_amount_cal
            return total_amount, federal_second
        elif gross_pay > 100392 and gross_pay < 155625:
            first_amount = 50197
            second_amount = 50195
            third_amount = gross_pay - 100392
            first_amount_cal = self.percentage(federal_first, first_amount)
            second_amount_cal = self.percentage(federal_second, second_amount)
            third_amount_cal = self.percentage(federal_three, third_amount)
            total_amount = first_amount_cal + second_amount_cal + third_amount_cal
            return total_amount, federal_three
        elif gross_pay > 155625 and gross_pay < 221708:
            first_amount = 50197
            second_amount = 50195
            third_amount = 55233
            fourth_amount = gross_pay - 155625
            first_amount_cal = self.percentage(federal_first, first_amount)
            second_amount_cal = self.percentage(federal_second, second_amount)
            third_amount_cal = self.percentage(federal_three, third_amount)
            fourth_amount_cal = self.percentage(federal_four, fourth_amount)
            total_amount = (
                first_amount_cal
                + second_amount_cal
                + third_amount_cal
                + fourth_amount_cal
            )
            return total_amount, federal_four
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
            total_amount = (
                first_amount_cal
                + second_amount_cal
                + third_amount_cal
                + fourth_amount_cal
                + fifth_amount_cal
            )
            return total_amount, federal_five

    def province_income_tax_calculator_y_t_d(self, gross_pay):
        if gross_pay < 46226:
            amount = self.percentage(province_first, gross_pay)
            return amount, province_first
        elif gross_pay >= 46227 and gross_pay <= 92454:
            first_amount = 46226
            second_amount = gross_pay - first_amount
            first_amount_cal = self.percentage(province_first, first_amount)
            second_amount_cal = self.percentage(province_second, second_amount)
            total_amount = first_amount_cal + second_amount_cal
            return total_amount, province_second
        elif gross_pay >= 92455 and gross_pay <= 150000:
            first_amount = 46226
            second_amount = 46229
            third_amount = gross_pay - 92455
            first_amount_cal = self.percentage(province_first, first_amount)
            second_amount_cal = self.percentage(province_second, second_amount)
            third_amount_cal = self.percentage(province_three, third_amount)
            total_amount = first_amount_cal + second_amount_cal + third_amount_cal
            return total_amount, province_three
        elif gross_pay >= 150001 and gross_pay <= 220000:
            first_amount = 46226
            second_amount = 46229
            third_amount = 57546
            fourth_amount = gross_pay - 150001
            first_amount_cal = self.percentage(province_first, first_amount)
            second_amount_cal = self.percentage(province_second, second_amount)
            third_amount_cal = self.percentage(province_three, third_amount)
            fourth_amount_cal = self.percentage(province_four, fourth_amount)
            total_amount = (
                first_amount_cal
                + second_amount_cal
                + third_amount_cal
                + fourth_amount_cal
            )
            return total_amount, province_four
        elif gross_pay > 220000:
            first_amount = 46226
            second_amount = 46229
            third_amount = 57546
            fourth_amount = 69999
            fifth_amount = gross_pay - 220000
            first_amount_cal = self.percentage(province_first, first_amount)
            second_amount_cal = self.percentage(province_second, second_amount)
            third_amount_cal = self.percentage(province_three, third_amount)
            fourth_amount_cal = self.percentage(province_four, fourth_amount)
            fifth_amount_cal = self.percentage(province_five, fifth_amount)
            total_amount = (
                first_amount_cal
                + second_amount_cal
                + third_amount_cal
                + fourth_amount_cal
                + fifth_amount_cal
            )
            return total_amount, province_five

    def total_incom_tax_calculator_period(
        self, gross_pay, total_percentage_for_monthly
    ):
        total_income_tax = self.percentage(total_percentage_for_monthly, gross_pay)
        return total_income_tax

    def total_incom_tax_calculator_year_to_date(self, y_to_d):
        fed_in_tax, percentage_fed = self.federal_income_tax_calculator_y_t_d(y_to_d)
        prov_in_tax, percentage_prov = self.province_income_tax_calculator_y_t_d(y_to_d)
        total_income_tax = fed_in_tax + prov_in_tax
        total_percentage_for_monthly = percentage_fed + percentage_prov
        return total_income_tax, total_percentage_for_monthly

    def EI_calculator_year_to_date(self, y_t_d_pay):
        if self.percentage(EI_Rate, y_t_d_pay) >= EI_Maximum_Deduction:
            return EI_Maximum_Deduction
        else:
            amount = self.percentage(EI_Rate, y_t_d_pay)
            return amount

    def CPP_Calculator_year_to_date(self, y_t_d_pay):
        if self.percentage(CPP_Rate, y_t_d_pay) >= CPP_Maximum_Deduction:
            return CPP_Maximum_Deduction
        else:
            amount = self.percentage(CPP_Rate, y_t_d_pay)
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
            amount = self.percentage(EI_Rate, gross_total)
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
        d = number.replace(",", "")
        return float(d)

    def comma_seprated(self, number):
        return f"{number:,}"

    def making_pdf_file(
        self,
        name_i,
        employee_address_i,
        hours_i,
        rate_i,
        g_total_i,
        account_number_i,
        year_to_date,
        period_ending_date_i,
        pay_date_i,
        i,
        income_tax_i,
        Ei_tax_i,
        cpp_tax_i,
        net_pay_i,
        year_to_date_incom_tax_i,
        year_to_date_ei_i,
        year_to_date_cpp_i,
    ):
        template = "Hanad-ADP-PAYSTUBS.docx"
        document = MailMerge(template)
        if Ei_tax_i == 0:
            ei_tax_mod = "00.00"
        else:
            ei_tax_mod = str(
                self.making_two_zer_dec(self.comma_seprated(round(Ei_tax_i, 2)))
            )

        if cpp_tax_i == 0:
            cpp_tax_mod = "00.00"
        else:
            cpp_tax_mod = str(
                self.making_two_zer_dec(self.comma_seprated(round(cpp_tax_i, 2)))
            )

        document.merge(
            employee_name_1=name_i.upper(),
            emp_2=name_i.upper(),
            hours=str(self.making_two_zer_dec1(hours_i)),
            rate=str(self.making_two_zer_dec1(rate_i)),
            total=str(self.comma_seprated((round(g_total_i, 2)))),
            gp_total=str(self.comma_seprated((round(g_total_i, 2)))),
            employee_addr=str(employee_address_i).upper(),
            employer_name=str(global_employer_name).upper(),
            employer_addr_1=str(global_employer_address_1).upper(),
            employer_addr_2=str(global_employer_address_2).upper(),
            employer_name_1=str(global_employer_name).upper(),
            employer_addr_1_1=str(global_employer_address_1).upper(),
            employer_addr_2_2=str(global_employer_address_2).upper(),
            acn=str(account_number_i),
            y_to_d_1=str(year_to_date),
            y_to_d_2=str(year_to_date),
            p_end_date=str(period_ending_date_i),
            inc_tax=str(
                self.making_two_zer_dec(self.comma_seprated(round(income_tax_i, 2)))
            ),
            ei_tax=ei_tax_mod,
            cpp_tax=cpp_tax_mod,
            net_pay_1=str(self.making_two_zer_dec(net_pay_i)),
            net_pay_2=str(self.making_two_zer_dec(net_pay_i)),
            net_pay_3=str(self.making_two_zer_dec(net_pay_i)),
            y_t_d_it=str(
                self.making_two_zer_dec(
                    self.comma_seprated(round(year_to_date_incom_tax_i, 2))
                )
            ),
            y_t_d_ei=str(
                self.making_two_zer_dec(
                    self.comma_seprated(round(year_to_date_ei_i, 2))
                )
            ),
            y_t_d_cpp=str(
                self.making_two_zer_dec(
                    self.comma_seprated(round(year_to_date_cpp_i, 2))
                )
            ),
            pay_date=str(pay_date_i),
            pay_date_2=str(pay_date_i),
        )
        document.write("Output_File.docx")
        abcd = period_ending_date_i.replace(" ", "-")
        am = dateparser.parse(abcd, settings={"DATE_ORDER": "DMY"})
        s = am.strftime("%d %b %y")
        fin = s.replace(" ", "-")
        self.convert_to_pdf(f"PAYSTUB-{fin}_{i}")
        try:
            os.remove(os.path.abspath("Output_File.docx"))
        except:
            print("Error in Removing File.")
        return

    def making_start_end_date(self, input_date):
        # convert input string to datetime object
        date_obj = datetime.strptime(input_date, "%b %d %Y")

        # calculate start and end dates based on input date
        last_day = datetime(date_obj.year, date_obj.month, 1) + timedelta(days=31)
        if last_day.month != date_obj.month:
            end_day = last_day - timedelta(days=last_day.day - 15)
        else:
            end_day = datetime(date_obj.year, date_obj.month, 15)

        if date_obj.day > 15:
            start_day = datetime(date_obj.year, date_obj.month, 16)
            next_month = date_obj.replace(day=28) + timedelta(days=4)
            if next_month.month != date_obj.month:
                end_day = datetime(next_month.year, next_month.month, 1) - timedelta(
                    days=1
                )
        else:
            start_day = datetime(date_obj.year, date_obj.month, 1)
            end_day = datetime(date_obj.year, date_obj.month, 15)
            if start_day.weekday() >= 4:
                end_day = datetime(date_obj.year, date_obj.month, 15)

        # format dates as strings in desired format
        # start_str = start_day.strftime("%Y-%m-%d")
        end_str = end_day.strftime("%b %d %Y")

        return f"{end_str}"

    def find_business_day(self, start_date, gap):
        import calendar

        # start_date = datetime.strptime(start_date, '%Y-%m-%d').date()
        start_date = datetime.strptime(start_date, "%b %d %Y").date()

        target_date = start_date + timedelta(days=gap)

        while True:
            if target_date.weekday() < 5 and not calendar.isleap(target_date.year):
                break
            target_date += timedelta(days=1)  # Increment the target_date by one day

        return str(start_date), str(target_date)

    def generate_period_date_list(self, start_date, num_dates):
        date_list = []
        current_date = start_date

        for _ in range(num_dates):
            res = self.making_start_end_date(current_date)
            st_dt = datetime.strptime(res, "%b %d %Y").date()
            final_date = st_dt + timedelta(days=2)
            end_date_str = final_date.strftime("%b %d %Y")
            current_date = end_date_str
            # date_list.append(datetime.strptime(res, "%b %d %Y").strftime("%d/%m/%Y"))
            date_list.append(res)

        return date_list

    def paystub_wrapper(self, name, employee_address, row_i):
        print(
            "************************************************************************************"
        )
        print(
            "********************** We are Doing Paystub A Document *******************************"
        )
        print(
            "************************************************************************************"
        )
        # From here we are implementing the functionality of paystub parameteres To Excel.

        if pd.isna(row_i["paystub_A_options"]):
            excel_rate = ""
            excel_account_number = ""
            excel_no_f_paystubs = ""
            excel_period_ending_date = ""
        else:
            paystub_excel_data = json.loads(row_i["paystub_A_options"])
            excel_rate = paystub_excel_data["Rate"]
            excel_account_number = paystub_excel_data["4_Digit_Account_Number"]
            excel_no_f_paystubs = paystub_excel_data["Numbe of Paystubs"]
            excel_period_ending_date = paystub_excel_data["Period"]

        if excel_rate == "":
            rate = float(input("Please enter the rate which you decided: "))
        else:
            rate = excel_rate
        print("***************************")
        print("***************************")
        # account_number = input("Last digits of bank account number XXXX : Yes or No :")
        # if account_number.lower() == "yes":
        #     account_number = randrange(1000, 9999)
        # else:
        if excel_account_number == "":
            account_number = input("Please Enter 4 digit Account Number : ")
        else:
            account_number = excel_account_number

        if excel_no_f_paystubs == "":
            number_of_pay_stubs = input(
                "Please enter, how many number of paystubs you want to create: "
            )
        else:
            number_of_pay_stubs = excel_no_f_paystubs

        if excel_period_ending_date == "":
            dates_lst = []
        else:
            dates_lst = self.generate_period_date_list(
                excel_period_ending_date, excel_no_f_paystubs
            )

        if int(number_of_pay_stubs) == 0:
            print("You have enterd 0. So i am not creating any paystub. Thanks")
            sys.exit()
        elif int(number_of_pay_stubs) > 0:
            for i in range(int(number_of_pay_stubs)):
                if dates_lst == [] or dates_lst == "":
                    period_ending_date = input("Please Enter Period for Paystub: ")
                else:
                    period_ending_date, check_date = self.find_business_day(
                        dates_lst[i], 2
                    )

                f_period_ending_date = self.parse_and_make_date(period_ending_date)

                new_year_to_send = f_period_ending_date.split("/")[-1]
                important_values_for_paystub = values_for_paystub(new_year_to_send)

                global federal_first
                federal_first = important_values_for_paystub["federal_first"]
                global province_first
                province_first = important_values_for_paystub["province_first"]
                global federal_second
                federal_second = important_values_for_paystub["federal_second"]
                global province_second
                province_second = important_values_for_paystub["province_second"]
                global federal_three
                federal_three = important_values_for_paystub["federal_three"]
                global province_three
                province_three = important_values_for_paystub["province_three"]
                global federal_four
                federal_four = important_values_for_paystub["federal_four"]
                global province_four
                province_four = important_values_for_paystub["province_four"]
                global federal_five
                federal_five = important_values_for_paystub["federal_five"]
                global province_five
                province_five = important_values_for_paystub["province_five"]
                global EI_Rate
                EI_Rate = important_values_for_paystub["EI_Rate"]
                global CPP_Rate
                CPP_Rate = important_values_for_paystub["CPP_Rate"]
                global EI_Maximum_Deduction
                EI_Maximum_Deduction = important_values_for_paystub[
                    "EI_Maximum_Deduction"
                ]
                global CPP_Maximum_Deduction
                CPP_Maximum_Deduction = important_values_for_paystub[
                    "CPP_Maximum_Deduction"
                ]
                # last_year_to_date = important_values_for_paystub['last_year_to_date']

                print("***************************")
                print("***************************")
                # hours = random.randint(75, 80)
                hours = int(
                    input(
                        "Please enter the number of hours employee has worked like (50 0r 60):"
                    )
                )
                gross_total = hours * rate
                if i == 0:
                    year_to_date = self.calculate_year_to_date(
                        rate, hours, period_ending_date
                    )
                    last_year_to_date = year_to_date
                elif i > 0:
                    year_to_date = self.return_float(last_year_to_date) + gross_total
                    year_to_date = self.comma_seprated(year_to_date)
                    last_year_to_date = year_to_date
                y_t_date_input = self.return_float(year_to_date)
                (
                    year_to_date_incom_tax,
                    total_percentage_for_monthly,
                ) = self.total_incom_tax_calculator_year_to_date(y_t_date_input)
                year_to_date_ei = self.EI_calculator_year_to_date(y_t_date_input)
                year_to_date_cpp = self.CPP_Calculator_year_to_date(y_t_date_input)

                if check_date == "":
                    pay_date = input("Please enter pay date: ")
                else:
                    pay_date = datetime.strptime(check_date, "%Y-%m-%d").strftime(
                        "%d/%m/%Y"
                    )
                print("***************************")
                print("***************************")

                income_tax = self.total_incom_tax_calculator_period(
                    gross_total, total_percentage_for_monthly
                )
                Ei_tax = self.EI_calculator_Period(gross_total, year_to_date_ei)
                cpp_tax = self.CPP_Calculator_Period(gross_total, year_to_date_cpp)
                net_pay = gross_total - income_tax - Ei_tax - cpp_tax
                round_pay = round(net_pay, 2)
                f_net_pay = f"{round_pay:,}"

                self.making_pdf_file(
                    name,
                    employee_address,
                    hours,
                    rate,
                    gross_total,
                    account_number,
                    year_to_date,
                    f_period_ending_date,
                    pay_date,
                    i,
                    income_tax,
                    Ei_tax,
                    cpp_tax,
                    f_net_pay,
                    year_to_date_incom_tax,
                    year_to_date_ei,
                    year_to_date_cpp,
                )
                
#######################################################################################################################################
#######################################################################################################################################
#######################################################################################################################################
#######################################################################################################################################


class Proof_Of_SIN:
    def making_address(self, address):
        address_list = address.split(" ")
        middle = int(len(address_list) / 2)
        address_1 = address_list[:middle]
        address_2 = address_list[middle:]
        address_1_f = " ".join(address_1)
        address_2_f = " ".join(address_2)
        return address_1_f, address_2_f

    def convert_to_pdf(self, filename):
        new_filename = filename + ".pdf"
        in_file = os.path.abspath("Output_File_SIN.docx")
        making_folder(global_name)
        out_file = os.path.abspath(f"Results/{global_name}/" + new_filename)
        convert(in_file, out_file)

    def making_fist_last_name(self, name):
        name_list = name.split(" ")
        first_name = name_list[:1]
        last_name = name_list[1:]
        first_name = "".join(first_name)
        last_name = " ".join(last_name)
        return first_name, last_name

    def making_sin(self, sin_number):
        sin_1 = str(int(sin_number[:3]))
        sin_2 = str(int(sin_number[3:6]))
        sin_3 = str(int(sin_number[6:]))
        return sin_1, sin_2, sin_3

    def making_sin_pdf_file(self, name_i, employee_address_i, sin_number_i):
        address_1, address_2 = self.making_address(employee_address_i)
        first_name, last_name = self.making_fist_last_name(name_i)
        sin1, sin2, sin3 = self.making_sin(sin_number_i)
        template = "Proof_Of_SIN.docx"
        document = MailMerge(template)
        document.merge(
            sin_name=str(name_i).upper(),
            address_1_sin=str(address_1).upper(),
            address_2_sin=str(address_2).upper(),
            first_name=str(first_name).upper(),
            last_name=str(last_name).upper(),
            sin_no_1=str(sin1).upper(),
            sin_no_2=str(sin2).upper(),
            sin_no_3=str(sin3).upper(),
        )
        document.write("Output_File_SIN.docx")
        self.convert_to_pdf(f"Proof_Of_SIN")
        try:
            os.remove(os.path.abspath("Output_File_SIN.docx"))
        except:
            print("Error in Removing File.")
        return

    def SIN_Wrapper(self, name, employee_address):
        print(
            "************************************************************************************"
        )
        print(
            "********************** We are Doing SIN Document *******************************"
        )
        print(
            "************************************************************************************"
        )

        self.making_sin_pdf_file(name, employee_address, global_sin_number)


class Proof_Of_Enrollment:
    def convert_to_pdf(self, filename):
        new_filename = filename + ".pdf"
        in_file = os.path.abspath("Output_File_POE.docx")
        making_folder(global_name)
        out_file = os.path.abspath(f"Results/{global_name}/" + new_filename)
        convert(in_file, out_file)

    def making_POE_pdf_file(
        self,
        enrollment_date_i,
        student_name_i,
        student_number_i,
        career_i,
        term_i,
        term_start_date_i,
        term_ending_date_i,
        faculty_i,
        plan_of_study_i,
        term_status_i,
        year_in_program_i,
        program_length_i,
    ):
        template = "POE.docx"
        document = MailMerge(template)
        document.merge(
            enrol_date=str(enrollment_date_i).capitalize(),
            student_name=str(student_name_i).title(),
            student_number=str(student_number_i),
            std_career=str(career_i).title(),
            std_term=str(term_i).title(),
            term_start_date=str(term_start_date_i).capitalize(),
            term_end_date=str(term_ending_date_i).capitalize(),
            faculty=str(faculty_i).title(),
            plan_of_study=str(plan_of_study_i).title(),
            term_status=str(term_status_i).title(),
            year_in_program=str(year_in_program_i).upper(),
            length=str(program_length_i).upper(),
        )
        document.write("Output_File_POE.docx")
        self.convert_to_pdf(f"Proof_of_Enrollment")
        try:
            os.remove(os.path.abspath("Output_File_POE.docx"))
        except:
            print("Error in Removing File.")
        return

    def pof_wrapper(self, name):
        print(
            "************************************************************************************"
        )
        print(
            "****************** We are Doing Proof Of Enrollment Document **********************"
        )
        print(
            "************************************************************************************"
        )

        print("*************************************")
        print("*************************************")
        enrollment_date = str(
            input("Please Enter Enrollment Date Like (June 12, 2021): ")
        )
        student_name = name

        student_number_default = "2512" + str(random.randint(10000, 99999))
        student_option_input = str(input("Student Number Default : Yes or NO : "))
        if student_option_input.lower() == "yes":
            student_number = student_number_default
        else:
            student_number = "2512" + str(input("Please Enter Student Number : "))
        print("*************************************")
        print("*************************************")
        career_default_value = "Undergraduate"
        career_option_input = str(input("Career Undergraduate: Yes or NO : "))
        if career_option_input.lower() == "yes":
            career = career_default_value
        else:
            career = str(input("Please Enter Career : "))
        print("*************************************")
        print("*************************************")
        term_year = str(input("Please Enter the year like (2020 or 2021) :"))
        term_first_val = str(
            input(
                """
                            Please Select any option from these. Select Number like 1 or 2
                            1 )  Fall/Winter
                            2 )  Summer                             
                            """
            )
        )
        if int(term_first_val) == 1:
            term_first_val = "Fall/Winter"
        elif int(term_first_val) == 2:
            term_first_val = "Summer"
        term = term_year + " " + term_first_val
        print("*************************************")
        print("*************************************")
        term_start_date = str(
            input("Please Enter Term Starting Date Like (September 9, YYYY): ")
        )
        print("*************************************")
        print("*************************************")
        term_ending_date = str(
            input("Please Enter Term ending date Like (April 30, YYYY): ")
        )
        print("*************************************")
        print("*************************************")
        faculty_default_value = "Faculty of Science"
        faculty_option_input = str(
            input("Faculty/Program of Study : Science  Yes or NO : ")
        )
        if faculty_option_input.lower() == "yes":
            faculty = faculty_default_value
        else:
            faculty = str(input("Please Enter Faculty/Program of Study :"))
        print("*************************************")
        print("*************************************")
        plan_of_study_default_value = "Bachelor of Science Honours (4 Year)"
        plan_of_study_input_option = str(
            input("Plan of Study Bachelor of Science Honours (4 Year) : YES or NO ")
        )
        if plan_of_study_input_option.lower() == "yes":
            plan_of_study = plan_of_study_default_value
        else:
            plan_of_study = str(input("Please Enter Plan of Study : "))
        print("*************************************")
        print("*************************************")
        term_status_default_value = "Full-time"
        term_status_input_option = str(input("Term Status Full-time : YES or NO "))
        if term_status_input_option.lower() == "yes":
            term_status = term_status_default_value
        else:
            term_status = str(input("Please Enter Term Status : "))
        print("*************************************")
        print("*************************************")
        year_in_program = str(input("Please Enter Year in Program : "))
        print("*************************************")
        print("*************************************")
        program_length_default_value = "4"
        program_length_input_value = str(input("Program Length is 4 : YES or NO  "))
        if program_length_input_value.lower() == "yes":
            program_length = program_length_default_value
        else:
            program_length = str(input("Please Enter Program Length : "))
        print("*************************************")
        print("*************************************")
        self.making_POE_pdf_file(
            enrollment_date,
            student_name,
            student_number,
            career,
            term,
            term_start_date,
            term_ending_date,
            faculty,
            plan_of_study,
            term_status,
            year_in_program,
            program_length,
        )


class TFour:
    def making_address(self, address):
        address_list = address.split(" ")
        middle = int(len(address_list) / 2)
        address_1 = address_list[:middle]
        address_2 = address_list[middle:]
        address_1_f = " ".join(address_1)
        address_2_f = " ".join(address_2)
        return address_1_f, address_2_f

    def convert_to_pdf(self, filename):
        new_filename = filename + ".pdf"
        in_file = os.path.abspath("Output_File_T4.docx")
        making_folder(global_name)
        out_file = os.path.abspath(f"Results/{global_name}/" + new_filename)
        convert(in_file, out_file)

    def making_sin(self, sin_number):
        sin_1 = str(int(sin_number[:3]))
        sin_2 = str(int(sin_number[3:6]))
        sin_3 = str(int(sin_number[6:]))
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
        if len(a[-1]) == 1:
            before_point = a[0]
            after_point = str(a[-1]) + "0"
            return before_point, after_point
        elif len(a) == 1:
            before_point = a[0]
            after_point = "00"
            return before_point, after_point
        else:
            before_point = a[0]
            after_point = a[-1][:2]
            return before_point, after_point

    def percentage(self, percent, whole):
        return (percent * whole) / 100.0

    def EI_calculator_year_to_date(self, y_t_d_pay, EI_Rate, EI_Maximum_Deduction):
        if self.percentage(EI_Rate, y_t_d_pay) >= EI_Maximum_Deduction:
            return EI_Maximum_Deduction
        else:
            amount = self.percentage(EI_Rate, y_t_d_pay)
            return amount

    def CPP_Calculator_year_to_date(self, y_t_d_pay, CPP_Rate, CPP_Maximum_Deduction):
        if self.percentage(CPP_Rate, y_t_d_pay) >= CPP_Maximum_Deduction:
            return CPP_Maximum_Deduction
        else:
            amount = self.percentage(CPP_Rate, y_t_d_pay)
            return amount

    def getting_eirate_and_max_deductions(self, year):
        from Constants import EI_Rates

        for i in EI_Rates:
            if i["year"] == int(year):
                return i["ei_rate"], i["max_deduction"]

    def getting_cpprate_and_max_deductions(slef, year):
        from Constants import CPP_Rates

        for i in CPP_Rates:
            if i["year"] == int(year):
                return i["cpp_rate"], i["max_deduction"]

    def getting_maximum_EI_insurable_amount(self, salary, year):
        from Constants import Max_EI_Insureable_Income

        year_max_income = 0
        for i in Max_EI_Insureable_Income:
            if i["year"] == int(year):
                year_max_income = i["income"]
        if salary > year_max_income:
            return year_max_income
        elif salary <= year_max_income:
            return salary

    def getting_maximum_CPP_insurable_amount(self, salary, year):
        from Constants import Max_CPP_Pensionable_Income

        year_max_income = 0
        for i in Max_CPP_Pensionable_Income:
            if i["year"] == int(year):
                year_max_income = i["income"]
        if salary > year_max_income:
            return year_max_income
        elif salary <= year_max_income:
            return salary

    def making_t4_pdf_file(
        self,
        name_i,
        employee_address_i,
        sin_number_i,
        t4_year_input_i,
        i_i,
        gross_salary_i,
        employer_name_i,
    ):
        template = "T4_2021_Creations.docx"
        document = MailMerge(template)
        ##############################
        address_1, address_2 = self.making_address(employee_address_i)
        first_name, last_name = self.making_fist_last_name(name_i)
        sin1, sin2, sin3 = self.making_sin(sin_number_i)
        # year_number = str(t4_year_input_i)[2:]
        year_number = str(t4_year_input_i)

        important_values_for_paystub = values_for_paystub(year_number)

        global federal_first
        federal_first = important_values_for_paystub["federal_first"]
        global province_first
        province_first = important_values_for_paystub["province_first"]
        global federal_second
        federal_second = important_values_for_paystub["federal_second"]
        global province_second
        province_second = important_values_for_paystub["province_second"]
        global federal_three
        federal_three = important_values_for_paystub["federal_three"]
        global province_three
        province_three = important_values_for_paystub["province_three"]
        global federal_four
        federal_four = important_values_for_paystub["federal_four"]
        global province_four
        province_four = important_values_for_paystub["province_four"]
        global federal_five
        federal_five = important_values_for_paystub["federal_five"]
        global province_five
        province_five = important_values_for_paystub["province_five"]
        global EI_Rate
        EI_Rate = important_values_for_paystub["EI_Rate"]
        global CPP_Rate
        CPP_Rate = important_values_for_paystub["CPP_Rate"]
        global EI_Maximum_Deduction
        EI_Maximum_Deduction = important_values_for_paystub["EI_Maximum_Deduction"]
        global CPP_Maximum_Deduction
        CPP_Maximum_Deduction = important_values_for_paystub["CPP_Maximum_Deduction"]

        paystub_object = PayStubs()
        income_tax, percntage = paystub_object.total_incom_tax_calculator_year_to_date(
            gross_salary_i
        )

        rate, max = self.getting_eirate_and_max_deductions(t4_year_input_i)
        t4_EI = self.EI_calculator_year_to_date(gross_salary_i, rate, max)

        cpp_rate, cpp_max = self.getting_cpprate_and_max_deductions(t4_year_input_i)
        t4_CPP = self.CPP_Calculator_year_to_date(gross_salary_i, cpp_rate, cpp_max)

        maximum_EI_amount = self.getting_maximum_EI_insurable_amount(
            gross_salary_i, t4_year_input_i
        )

        maximum_cpp_amount = self.getting_maximum_CPP_insurable_amount(
            gross_salary_i, t4_year_input_i
        )

        ################################################################################################

        befor_point_gross_sal, after_point_gross_sal = self.breaking_number(
            gross_salary_i
        )
        before_point_income_tax, after_point_income_tax = self.breaking_number(
            income_tax
        )
        before_point_cpp, after_point_cpp = self.breaking_number(t4_CPP)
        befor_point_ei, after_point_ei = self.breaking_number(t4_EI)
        befor_point_max_ei, after_point_max_ei = self.breaking_number(maximum_EI_amount)
        befor_point_max_cpp, after_point_max_cpp = self.breaking_number(
            maximum_cpp_amount
        )

        document.merge(
            emp_name_1=str(employer_name_i).upper(),
            emp_name_2=str(employer_name_i).upper(),
            ###################
            t4_year=str(t4_year_input_i),
            t4_year_1=str(t4_year_input_i),
            ###################
            t4_address_1_1=str(address_1).upper(),
            t4_address_1_2=str(address_1).upper(),
            ###################
            t4_address_2_1=str(address_2).upper(),
            t4_address_2_2=str(address_2).upper(),
            ###################
            f_nm_1=str(first_name).upper(),
            f_nm_2=str(first_name).upper(),
            ###################
            l_nm_1=str(last_name).upper(),
            l_nm_2=str(last_name).upper(),
            ###################
            sn1=str(sin1),
            sn2=str(sin2),
            sn3=str(sin3),
            sn4=str(sin1),
            sn5=str(sin2),
            sn6=str(sin3),
            ###################
            yn=str(year_number),
            y2=str(year_number),
            ###################
            grs_sal_bfr=str(self.comma_seprated(int(befor_point_gross_sal))),
            g_af=str(after_point_gross_sal),
            inc_tax_bfr=str(self.comma_seprated(int(before_point_income_tax))),
            i_af=str(after_point_income_tax),
            t4_cpp_bfr=str(self.comma_seprated(int(before_point_cpp))),
            c_af=str(after_point_cpp),
            t4_ei_bfr=str(self.comma_seprated(int(befor_point_ei))),
            e_af=str(after_point_ei),
            mx_ei_bfr=str(self.comma_seprated(int(befor_point_max_ei))),
            d_af=str(after_point_max_ei),
            mx_cp_bfr=str(self.comma_seprated(int(befor_point_max_cpp))),
            h_af=str(after_point_max_cpp),
            ###################
            grs_sal_bfr_1=str(self.comma_seprated(int(befor_point_gross_sal))),
            u_af=str(after_point_gross_sal),
            inc_tax_bfr_1=str(self.comma_seprated(int(before_point_income_tax))),
            w_af=str(after_point_income_tax),
            t4_cpp_bfr_1=str(self.comma_seprated(int(before_point_cpp))),
            v_af=str(after_point_cpp),
            t4_ei_bfr_1=str(self.comma_seprated(int(befor_point_ei))),
            x_af=str(after_point_ei),
            max_ei_bfr_1=str(self.comma_seprated(int(befor_point_max_ei))),
            y_af=str(after_point_max_ei),
            max_cp_bfr_1=str(self.comma_seprated(int(befor_point_max_cpp))),
            z_af=str(after_point_max_cpp),
        )
        document.write(f"Output_File_T4.docx")
        self.convert_to_pdf(f"T4-{t4_year_input_i}-{i_i}")
        try:
            os.remove(os.path.abspath("Output_File_T4.docx"))
        except:
            print("Error in Removing File.")
        return

    def T4_Wrapper(self, name, employee_address):
        print(
            "************************************************************************************"
        )
        print(
            "************************* We are Doing T4 Document *********************************"
        )
        print(
            "************************************************************************************"
        )

        number_of_documents = int(
            input("Please enter How many documents you want to create: ")
        )
        print("*************************************")
        print("*************************************")
        if number_of_documents == 0:
            print("You have Enter 0, We are not creating any document. Thanks")
        else:
            for i in range(number_of_documents):
                print("*************************************")
                print("*************************************")
                t4_year_input = input("Please Enter Year for T4: ")
                print("*************************************")
                print("*************************************")
                gross_salary = float(input("Please enter your gross Salary: "))
                print("*************************************")
                print("*************************************")
                self.making_t4_pdf_file(
                    name,
                    employee_address,
                    global_sin_number,
                    t4_year_input,
                    i,
                    gross_salary,
                    global_employer_name,
                )


class TD_Document:
    global_starting_balance = 0
    global_account_number = ""

    def ret_bank_name(self, name):
        name = name.split(" ")
        try:
            bnk_1 = " ".join(name[:3])
        except:
            bnk_1 = ""
        try:
            bnk_2 = " ".join(name[3:7])
        except:
            bnk_2 = ""
        try:
            bnk_3 = " ".join(name[7:])
        except:
            bnk_3 = ""
        return bnk_1, bnk_2, bnk_3

    def making_address(self, address):
        address_list = address.split(" ")
        middle = int(len(address_list) / 2)
        address_1 = address_list[:middle]
        address_2 = address_list[middle:]
        address_1_f = " ".join(address_1)
        address_2_f = " ".join(address_2)
        return address_1_f, address_2_f

    def convert_to_pdf(self, filename):
        new_filename = filename + ".pdf"
        in_file = os.path.abspath("Output_File_TD.docx")
        making_folder(global_name)
        out_file = os.path.abspath(f"Results/{global_name}/" + new_filename)
        convert(in_file, out_file)

    def making_account_number(self, number):
        num_1 = random.randint(100, 999)
        num_2 = random.randint(100, 999)
        account_number = f"{num_1}-{num_2}{number}"
        return account_number

    def making_statement_from(self, input):
        from calendar import monthrange
        from time import strptime

        d = input.split(" ")
        month = d[0][:3]
        year = d[1]
        month_number = strptime(month, "%b").tm_mon
        num_days = monthrange(int(year), month_number)[1]
        ret_string = (
            f"{month.upper()} 1/{year[2:]} - {month.upper()} {num_days}/{year[2:]}"
        )
        return ret_string, num_days

    def starting_blnc_date(self, text):
        data = text.split("-")[0].split("/")[0]
        data1 = data.split(" ")
        if len(data1[-1]) > 1:
            return data
        else:
            date = f"{0}{data1[-1]}"
            date2 = f"{data1[0]} {date}"
            return date2

    def making_two_zer_dec(self, num):
        a = num.split(".")
        if len(a[-1]) == 2:
            return ".".join(a)
        elif len(a[-1]) == 1:
            a[-1] = a[-1] + "0"
            return ".".join(a)
        elif len(a[-1]) > 2 and a[-1] != a[0]:
            a[-1] = a[-1][:2]
            return ".".join(a)
        elif a[-1] == a[0]:
            a[0] = f"{a[0]}.00"
            return ".".join(a)

    def comma_seprated(self, number):
        return f"{number:,}"

    def calPercent(self, x):
        percent2 = 33 / 100 * x
        percent1 = 67 / 100 * x
        return round(percent1), round(percent2)

    def myFunc(self, e):
        return e["Date"]

    def solving_balance_clean_function(self, dict):
        u_values = list(set([i["Date"] for i in dict]))
        testings = []
        for j in u_values:
            lst = []
            for k in dict:
                if k["Date"] == j:
                    lst.append(k)
            testings.append(lst)
        for test in testings:
            if len(test) > 1:
                for ent in range(len(test) - 1):
                    test[ent]["balance"] = ""
        flat_list = [item for sublist in testings for item in sublist]
        return flat_list

    def adding_month(self, dict, mon):
        for i in dict:
            if len(str(i["Date"])) > 1:
                old_val = i["Date"]
                i["Date"] = f"{mon} {old_val}"
            elif len(str(i["Date"])) == 1:
                old_val = i["Date"]
                i["Date"] = f"{mon} 0{old_val}"
        return dict

    def calculate_total_wth_drawl(self, dict):
        total = 0
        for j in dict:
            if j["withdraw"] != "":
                total = total + float(j["withdraw"])
        return total

    def calculate_total_deposit(self, dict):
        total = 0
        for j in dict:
            if j["deposit"] != "":
                total = total + float(j["deposit"])
        return total

    def convrt_val(self, text):
        first_step = text.split(".")
        if len(first_step) == 2:
            after_point = len(first_step[-1])
            if after_point == 1:
                first_step[-1] = first_step[-1] + "0"
                final_ret_val = first_step[0] + "." + first_step[-1]
                return final_ret_val
            elif after_point == 2:
                return first_step[0] + "." + first_step[-1]
        elif len(first_step) == 1:
            send_val = "00"
            final_ret_val = first_step[0] + "." + send_val
            return final_ret_val

    def final_update_on_td(self, trans):
        new_trans = []
        for i in trans:
            try:
                i["withdraw"] = self.convrt_val(i["withdraw"])
            except:
                pass
            try:
                i["deposit"] = self.convrt_val(i["deposit"])
            except:
                pass
            new_trans.append(i)
        return new_trans

    def another_final_update_on_td(self, trans):
        new_trans = []
        for i in trans:
            try:
                if i["withdraw"] == ".00":
                    i["withdraw"] = ""
            except:
                pass
            try:
                if i["deposit"] == ".00":
                    i["deposit"] = ""
            except:
                pass
            new_trans.append(i)
        return new_trans

    def making_all_transactions(
        self,
        trans,
        global_balance,
        incoming_deposits,
        month_days,
        date,
    ):
        withdrawls_trans, deposits_trans = self.calPercent(trans)
        wth_drws = random.sample(withdrawls, withdrawls_trans)
        depos = random.sample(deposits, deposits_trans)

        all_transactions = []
        for i in wth_drws:
            i["Date"] = random.randint(2, month_days)
            try:
                del i["balance"]
            except:
                pass
            try:
                del i["deposit"]
            except:
                pass
            all_transactions.append(i)

        for j in depos:
            j["Date"] = random.randint(2, month_days)
            try:
                del j["balance"]
            except:
                pass
            try:
                del j["withdraw"]
            except:
                pass
            all_transactions.append(j)

        for k in incoming_deposits:
            all_transactions.append(k)

        all_transactions.sort(key=self.myFunc)

        for tran in all_transactions:
            if "withdraw" not in tran and "deposit" in tran:
                payment = global_balance + float(tran["deposit"])
                global_balance = payment
                tran["balance"] = self.making_two_zer_dec(
                    self.comma_seprated(float(global_balance))
                )
                tran["withdraw"] = ""
            elif "deposit" not in tran and "withdraw" in tran:
                payment = global_balance - float(tran["withdraw"])
                global_balance = payment
                tran["balance"] = self.making_two_zer_dec(
                    self.comma_seprated(float(global_balance))
                )
                tran["deposit"] = ""

        all_transactions = self.solving_balance_clean_function(all_transactions)

        all_transactions = self.adding_month(all_transactions, date)

        for i in range(30 - trans):
            all_transactions.append(
                {
                    "description": "",
                    "deposit": "",
                    "balance": "",
                    "withdraw": "",
                    "Date": "",
                }
            )

        return all_transactions, global_balance

    def making_TD_pdf_file_for_thirty_trans(
        self,
        b_1_i,
        b_2_i,
        b_3_i,
        name_i,
        address_i,
        branch_number_i,
        account_type_i,
        statement_from_i,
        starting_balance_i,
        i_i,
        total_deposits_i,
        total_transactions_i,
    ):
        template = "TD_Document_Final.docx"
        document = MailMerge(template)
        # "*************************************"

        ad_1, ad_2 = self.making_address(address_i)
        statement_date, month_days = self.making_statement_from(statement_from_i)
        starting_balance_dat = self.starting_blnc_date(statement_date)
        date_to_send = starting_balance_dat[:3]
        trans_after_final_mod, ret_balance = self.making_all_transactions(
            total_transactions_i,
            self.global_starting_balance,
            total_deposits_i,
            month_days,
            date_to_send,
        )
        self.global_starting_balance = ret_balance
        total_with_drawl = self.calculate_total_wth_drawl(trans_after_final_mod)
        total_depos = self.calculate_total_deposit(trans_after_final_mod)

        for f_trans in trans_after_final_mod:
            try:
                f_trans["deposit"] = self.comma_seprated(float(f_trans["deposit"]))
            except:
                pass
            try:
                f_trans["withdraw"] = self.comma_seprated(float(f_trans["withdraw"]))
            except:
                pass

        trans_after_final_mod = self.final_update_on_td(trans_after_final_mod)
        trans_after_final_mod = self.another_final_update_on_td(trans_after_final_mod)

        # "*************************************"
        document.merge(
            bnk_1=str(b_1_i).upper(),
            bnk_2=str(b_2_i).upper(),
            bnk_3=str(b_3_i).upper(),
            emp_name=str(name_i).upper(),
            adr_1=str(ad_1).upper(),
            adr_2=str(ad_2).upper(),
            br_no=str(branch_number_i),
            ac_no=str(self.global_account_number),
            acc_type=str(account_type_i),
            stmnt_date=str(statement_date),
            strt_bl=str(
                self.making_two_zer_dec(self.comma_seprated(float(starting_balance_i)))
            ),
            st_date=str(starting_balance_dat),
            ttl_wth=str(self.making_two_zer_dec(self.comma_seprated(total_with_drawl))),
            ttl_dep=str(self.making_two_zer_dec(self.comma_seprated(total_depos))),
            # ******************************************
            des_1=str(trans_after_final_mod[0]["description"][:15]),
            wth_1=str(trans_after_final_mod[0]["withdraw"]),
            dep_1=str(trans_after_final_mod[0]["deposit"]),
            dt_1=str(trans_after_final_mod[0]["Date"]),
            blnc_1=str(trans_after_final_mod[0]["balance"]),
            # *******************************************
            # ******************************************
            des_2=str(trans_after_final_mod[1]["description"][:15]),
            wth_2=str(trans_after_final_mod[1]["withdraw"]),
            dep_2=str(trans_after_final_mod[1]["deposit"]),
            dt_2=str(trans_after_final_mod[1]["Date"]),
            blnc_2=str(trans_after_final_mod[1]["balance"]),
            # *******************************************
            # ******************************************
            des_3=str(trans_after_final_mod[2]["description"][:15]),
            wth_3=str(trans_after_final_mod[2]["withdraw"]),
            dep_3=str(trans_after_final_mod[2]["deposit"]),
            dt_3=str(trans_after_final_mod[2]["Date"]),
            blnc_3=str(trans_after_final_mod[2]["balance"]),
            # *******************************************
            # ******************************************
            des_4=str(trans_after_final_mod[3]["description"][:15]),
            wth_4=str(trans_after_final_mod[3]["withdraw"]),
            dep_4=str(trans_after_final_mod[3]["deposit"]),
            dt_4=str(trans_after_final_mod[3]["Date"]),
            blnc_4=str(trans_after_final_mod[3]["balance"]),
            # *******************************************
            # ******************************************
            des_5=str(trans_after_final_mod[4]["description"][:15]),
            wth_5=str(trans_after_final_mod[4]["withdraw"]),
            dep_5=str(trans_after_final_mod[4]["deposit"]),
            dt_5=str(trans_after_final_mod[4]["Date"]),
            blnc_5=str(trans_after_final_mod[4]["balance"]),
            # *******************************************
            # ******************************************
            des_6=str(trans_after_final_mod[5]["description"][:15]),
            wth_6=str(trans_after_final_mod[5]["withdraw"]),
            dep_6=str(trans_after_final_mod[5]["deposit"]),
            dt_6=str(trans_after_final_mod[5]["Date"]),
            blnc_6=str(trans_after_final_mod[5]["balance"]),
            # *******************************************
            # ******************************************
            des_7=str(trans_after_final_mod[6]["description"][:15]),
            wth_7=str(trans_after_final_mod[6]["withdraw"]),
            dep_7=str(trans_after_final_mod[6]["deposit"]),
            dt_7=str(trans_after_final_mod[6]["Date"]),
            blnc_7=str(trans_after_final_mod[6]["balance"]),
            # *******************************************
            # ******************************************
            des_8=str(trans_after_final_mod[7]["description"][:15]),
            wth_8=str(trans_after_final_mod[7]["withdraw"]),
            dep_8=str(trans_after_final_mod[7]["deposit"]),
            dt_8=str(trans_after_final_mod[7]["Date"]),
            blnc_8=str(trans_after_final_mod[7]["balance"]),
            # *******************************************
            # ******************************************
            des_9=str(trans_after_final_mod[8]["description"][:15]),
            wth_9=str(trans_after_final_mod[8]["withdraw"]),
            dep_9=str(trans_after_final_mod[8]["deposit"]),
            dt_9=str(trans_after_final_mod[8]["Date"]),
            blnc_9=str(trans_after_final_mod[8]["balance"]),
            # *******************************************
            # ******************************************
            des_10=str(trans_after_final_mod[9]["description"][:15]),
            wth_10=str(trans_after_final_mod[9]["withdraw"]),
            dep_10=str(trans_after_final_mod[9]["deposit"]),
            dt_10=str(trans_after_final_mod[9]["Date"]),
            blnc_10=str(trans_after_final_mod[9]["balance"]),
            # *******************************************
            # ******************************************
            des_11=str(trans_after_final_mod[10]["description"][:15]),
            wth_11=str(trans_after_final_mod[10]["withdraw"]),
            dep_11=str(trans_after_final_mod[10]["deposit"]),
            dt_11=str(trans_after_final_mod[10]["Date"]),
            blnc_11=str(trans_after_final_mod[10]["balance"]),
            # *******************************************
            # ******************************************
            des_12=str(trans_after_final_mod[11]["description"][:15]),
            wth_12=str(trans_after_final_mod[11]["withdraw"]),
            dep_12=str(trans_after_final_mod[11]["deposit"]),
            dt_12=str(trans_after_final_mod[11]["Date"]),
            blnc_12=str(trans_after_final_mod[11]["balance"]),
            # *******************************************
            # ******************************************
            des_13=str(trans_after_final_mod[12]["description"][:15]),
            wth_13=str(trans_after_final_mod[12]["withdraw"]),
            dep_13=str(trans_after_final_mod[12]["deposit"]),
            dt_13=str(trans_after_final_mod[12]["Date"]),
            blnc_13=str(trans_after_final_mod[12]["balance"]),
            # *******************************************
            # ******************************************
            des_14=str(trans_after_final_mod[13]["description"][:15]),
            wth_14=str(trans_after_final_mod[13]["withdraw"]),
            dep_14=str(trans_after_final_mod[13]["deposit"]),
            dt_14=str(trans_after_final_mod[13]["Date"]),
            blnc_14=str(trans_after_final_mod[13]["balance"]),
            # *******************************************
            # ******************************************
            des_15=str(trans_after_final_mod[14]["description"][:15]),
            wth_15=str(trans_after_final_mod[14]["withdraw"]),
            dep_15=str(trans_after_final_mod[14]["deposit"]),
            dt_15=str(trans_after_final_mod[14]["Date"]),
            blnc_15=str(trans_after_final_mod[14]["balance"]),
            # *******************************************
            # ******************************************
            des_16=str(trans_after_final_mod[15]["description"][:15]),
            wth_16=str(trans_after_final_mod[15]["withdraw"]),
            dep_16=str(trans_after_final_mod[15]["deposit"]),
            dt_16=str(trans_after_final_mod[15]["Date"]),
            blnc_16=str(trans_after_final_mod[15]["balance"]),
            # *******************************************
            # ******************************************
            des_17=str(trans_after_final_mod[16]["description"][:15]),
            wth_17=str(trans_after_final_mod[16]["withdraw"]),
            dep_17=str(trans_after_final_mod[16]["deposit"]),
            dt_17=str(trans_after_final_mod[16]["Date"]),
            blnc_17=str(trans_after_final_mod[16]["balance"]),
            # *******************************************
            # ******************************************
            des_18=str(trans_after_final_mod[17]["description"][:15]),
            wth_18=str(trans_after_final_mod[17]["withdraw"]),
            dep_18=str(trans_after_final_mod[17]["deposit"]),
            dt_18=str(trans_after_final_mod[17]["Date"]),
            blnc_18=str(trans_after_final_mod[17]["balance"]),
            # *******************************************
            # ******************************************
            des_19=str(trans_after_final_mod[18]["description"][:15]),
            wth_19=str(trans_after_final_mod[18]["withdraw"]),
            dep_19=str(trans_after_final_mod[18]["deposit"]),
            dt_19=str(trans_after_final_mod[18]["Date"]),
            blnc_19=str(trans_after_final_mod[18]["balance"]),
            # *******************************************
            # ******************************************
            des_20=str(trans_after_final_mod[19]["description"][:15]),
            wth_20=str(trans_after_final_mod[19]["withdraw"]),
            dep_20=str(trans_after_final_mod[19]["deposit"]),
            dt_20=str(trans_after_final_mod[19]["Date"]),
            blnc_20=str(trans_after_final_mod[19]["balance"]),
            # *******************************************
            # ******************************************
            des_21=str(trans_after_final_mod[20]["description"][:15]),
            wth_21=str(trans_after_final_mod[20]["withdraw"]),
            dep_21=str(trans_after_final_mod[20]["deposit"]),
            dt_21=str(trans_after_final_mod[20]["Date"]),
            blnc_21=str(trans_after_final_mod[20]["balance"]),
            # *******************************************
            # ******************************************
            des_22=str(trans_after_final_mod[21]["description"][:15]),
            wth_22=str(trans_after_final_mod[21]["withdraw"]),
            dep_22=str(trans_after_final_mod[21]["deposit"]),
            dt_22=str(trans_after_final_mod[21]["Date"]),
            blnc_22=str(trans_after_final_mod[21]["balance"]),
            # *******************************************
            # ******************************************
            des_23=str(trans_after_final_mod[22]["description"][:15]),
            wth_23=str(trans_after_final_mod[22]["withdraw"]),
            dep_23=str(trans_after_final_mod[22]["deposit"]),
            dt_23=str(trans_after_final_mod[22]["Date"]),
            blnc_23=str(trans_after_final_mod[22]["balance"]),
            # *******************************************
            # ******************************************
            des_24=str(trans_after_final_mod[23]["description"][:15]),
            wth_24=str(trans_after_final_mod[23]["withdraw"]),
            dep_24=str(trans_after_final_mod[23]["deposit"]),
            dt_24=str(trans_after_final_mod[23]["Date"]),
            blnc_24=str(trans_after_final_mod[23]["balance"]),
            # *******************************************
            # ******************************************
            des_25=str(trans_after_final_mod[24]["description"][:15]),
            wth_25=str(trans_after_final_mod[24]["withdraw"]),
            dep_25=str(trans_after_final_mod[24]["deposit"]),
            dt_25=str(trans_after_final_mod[24]["Date"]),
            blnc_25=str(trans_after_final_mod[24]["balance"]),
            # *******************************************
            # ******************************************
            des_26=str(trans_after_final_mod[25]["description"][:15]),
            wth_26=str(trans_after_final_mod[25]["withdraw"]),
            dep_26=str(trans_after_final_mod[25]["deposit"]),
            dt_26=str(trans_after_final_mod[25]["Date"]),
            blnc_26=str(trans_after_final_mod[25]["balance"]),
            # *******************************************
            # ******************************************
            des_27=str(trans_after_final_mod[26]["description"][:15]),
            wth_27=str(trans_after_final_mod[26]["withdraw"]),
            dep_27=str(trans_after_final_mod[26]["deposit"]),
            dt_27=str(trans_after_final_mod[26]["Date"]),
            blnc_27=str(trans_after_final_mod[26]["balance"]),
            # *******************************************
            # ******************************************
            des_28=str(trans_after_final_mod[27]["description"][:15]),
            wth_28=str(trans_after_final_mod[27]["withdraw"]),
            dep_28=str(trans_after_final_mod[27]["deposit"]),
            dt_28=str(trans_after_final_mod[27]["Date"]),
            blnc_28=str(trans_after_final_mod[27]["balance"]),
            # *******************************************
            # ******************************************
            des_29=str(trans_after_final_mod[28]["description"][:15]),
            wth_29=str(trans_after_final_mod[28]["withdraw"]),
            dep_29=str(trans_after_final_mod[28]["deposit"]),
            dt_29=str(trans_after_final_mod[28]["Date"]),
            blnc_29=str(trans_after_final_mod[28]["balance"]),
            # *******************************************
            # ******************************************
            des_30=str(trans_after_final_mod[29]["description"][:15]),
            wth_30=str(trans_after_final_mod[29]["withdraw"]),
            dep_30=str(trans_after_final_mod[29]["deposit"]),
            dt_30=str(trans_after_final_mod[29]["Date"]),
            blnc_30=str(trans_after_final_mod[29]["balance"]),
            # *******************************************
        )
        document.write("Output_File_TD.docx")
        self.convert_to_pdf(f"TD_Trust_{i_i}")
        try:
            os.remove(os.path.abspath("Output_File_TD.docx"))
        except:
            print("Error in Removing File.")
        return

    def TD_wrapper(self, emp_name, address):
        print(
            "************************************************************************************"
        )
        print(
            "************************** We are Doing TD Document ********************************"
        )
        print(
            "************************************************************************************"
        )

        number_of_months = input("How many months you want to create for : ")
        b_1 = ""
        b_2 = ""
        b_3 = ""
        branch_number = ""
        account_type = ""
        name = ""
        for i in range(int(number_of_months)):
            if i == 0:
                print("*************************************")
                print("*************************************")
                default_bank_name = (
                    "LONDON POND MILLS 1086 COMMISSIONERS ROAD EAST LONDON, ON N5Z 4W8"
                )
                bank_option_input = str(
                    input("Branch info will remain same : Yes or No : ")
                )
                if bank_option_input.lower() == "yes":
                    b_1, b_2, b_3 = self.ret_bank_name(default_bank_name)
                else:
                    bank_name = str(input("Please Enter Bank Name: "))
                    b_1, b_2, b_3 = self.ret_bank_name(bank_name)
            if i == 0:
                print("*************************************")
                print("*************************************")
                default_branch_no = "005110"
                branch_number = input("Branch Number Will remain Same : Yes or No : ")
                if branch_number.lower() == "yes":
                    branch_number = default_branch_no
                else:
                    branch_number = input("Please Enter Branch Number : ")
            if i == 0:
                print("*************************************")
                print("*************************************")
                account_number = input("Please Enter 4 Digit Account Number : ")
                account_num = self.making_account_number(account_number)
                self.global_account_number = account_num
            if i == 0:
                print("*************************************")
                print("*************************************")
                default_account_type = "STUDENT"
                account_type = input("Account type will remain same : Yes or No : ")
                if account_type.lower() == "yes":
                    account_type = default_account_type
                else:
                    account_type = input("Please Enter account type : ")
            print("*************************************")
            print("*************************************")
            statement_from = input("Please enter year and month like(feb, 2023) : ")
            if i == 0:
                print("*************************************")
                print("*************************************")
                starting_balance = str(input("Please Enter Starting Balanace : "))
                self.global_starting_balance = float(starting_balance)
            elif i > 0:
                starting_balance = self.global_starting_balance
            print("*************************************")
            print("*************************************")
            total_deposits = []
            number_of_deposits = int(input("How many deposits you want to make : "))
            for j in range(number_of_deposits):
                if j == 0 and i == 0:
                    name = str(input("Please enter Employer Name : "))
                date = int(input("Please enter date : "))
                amount = float(input("Please enter amount : "))
                print("*************************************")
                total_deposits.append(
                    {"description": name.upper(), "deposit": amount, "Date": date}
                )
            print("*************************************")
            print("*************************************")
            total_transactions = int(
                input(
                    "Please Enter the total number of transactions you want to make : "
                )
            )
            total_transactions = total_transactions - len(total_deposits)

            self.making_TD_pdf_file_for_thirty_trans(
                b_1,
                b_2,
                b_3,
                emp_name,
                address,
                branch_number,
                account_type,
                statement_from,
                starting_balance,
                i,
                total_deposits,
                total_transactions,
            )


class PayStubChild(PayStubs):
    def date_range(self, input_date):
        # convert input string to datetime object
        date_obj = datetime.strptime(input_date, "%b %d %Y")

        # calculate start and end dates based on input date
        last_day = datetime(date_obj.year, date_obj.month, 1) + timedelta(days=31)
        if last_day.month != date_obj.month:
            end_day = last_day - timedelta(days=last_day.day - 15)
        else:
            end_day = datetime(date_obj.year, date_obj.month, 15)

        if date_obj.day > 15:
            start_day = datetime(date_obj.year, date_obj.month, 16)
            next_month = date_obj.replace(day=28) + timedelta(days=4)
            if next_month.month != date_obj.month:
                end_day = datetime(next_month.year, next_month.month, 1) - timedelta(
                    days=1
                )
        else:
            start_day = datetime(date_obj.year, date_obj.month, 1)
            end_day = datetime(date_obj.year, date_obj.month, 15)
            if start_day.weekday() >= 4:
                end_day = datetime(date_obj.year, date_obj.month, 15)

        # format dates as strings in desired format
        start_str = start_day.strftime("%Y-%m-%d")
        end_str = end_day.strftime("%Y-%m-%d")

        return f"{start_str} - {end_str}"

    def checque_date(self, input_str):
        date_obj = datetime.strptime(input_str, "%b %d %Y")
        return date_obj.strftime("%Y-%m-%d")

    def format_number(self, num):
        return "{:,.2f}".format(float(num))

    def making_paystub_two_document(
        self,
        e_name_i,
        e_address_i,
        occupation_i,
        pay_period_i,
        cheque_date_i,
        number_of_hours_i,
        rate_per_hour_i,
        gross_total_i,
        year_to_date_for_paystub2_i,
        year_to_date_cpp_i,
        year_to_date_ei_i,
        year_to_date_incom_tax_i,
        cpp_tax_i,
        Ei_tax_i,
        income_tax_i,
        current_total_i,
        total_y_t_d_calculations_i,
        cur_tot1_i,
        y_t_d_net_cal1_i,
        paystub_number_i,
    ):
        template = "PaystubTwo.docx"
        document = MailMerge(template)

        td_object = TD_Document()
        ad_1, ad_2 = td_object.making_address(e_address_i)
        processed_pay_period = self.date_range(pay_period_i)
        final_cheque_date = self.checque_date(cheque_date_i)

        document.merge(
            employer_name=str(global_employer_name).upper(),
            employer_ad_1=str(global_employer_address_1).upper(),
            employer_ad_2=str(global_employer_address_2).upper(),
            employer_name_2=str(global_employer_name).upper(),
            employer_add_3=str(global_employer_address_1).upper(),
            employer_add_4=str(global_employer_address_2).upper(),
            employee_name=str(e_name_i).upper(),
            employee_name_1=str(e_name_i).upper(),
            employee_add_1=str(ad_1).upper(),
            employee_add_2=str(ad_2).upper(),
            emp_addr_3=str(e_address_i).upper(),
            occupation=str(occupation_i).upper(),
            pay_period=str(processed_pay_period),
            cheque=str(final_cheque_date),
            qty=self.format_number(str(number_of_hours_i)),
            rate=self.format_number(str(rate_per_hour_i)),
            curr=self.format_number(str(gross_total_i)),
            y_t_d=self.format_number(
                self.return_float(str(year_to_date_for_paystub2_i))
            ),
            ytd_cpp=self.format_number(str(year_to_date_cpp_i)),
            ytd_ei=self.format_number(str(year_to_date_ei_i)),
            ytd_in=self.format_number(str(year_to_date_incom_tax_i)),
            cpp=self.format_number(str(cpp_tax_i)),
            ei=self.format_number(str(Ei_tax_i)),
            inc=self.format_number(str(income_tax_i)),
            cur_net=self.format_number(current_total_i),
            y_td_net=self.format_number(total_y_t_d_calculations_i),
            cur_total=self.format_number(str(cur_tot1_i)),
            y_td_tot=self.format_number(str(y_t_d_net_cal1_i)),
        )

        document.write("Output_File.docx")
        self.convert_to_pdf(f"PayStubTwo-{paystub_number_i}")
        try:
            os.remove(os.path.abspath("Output_File.docx"))
        except:
            print("Error in Removing File.")
        return

    def paystub_child_wrapper(self, e_name, e_address):
        print(
            "************************************************************************************"
        )
        print(
            "************************** We are Doing PayStub 2 ********************************"
        )
        print(
            "************************************************************************************"
        )

        rate_per_hour = int(input("Please Enter Rate per Hour : "))
        print("*" * 100)
        print("*" * 100)

        occupation = input("Please Enter Occupation : ")
        print("*" * 100)
        print("*" * 100)

        number_of_pay_stubs = input(
            "Please enter, how many number of paystubs you want to create: "
        )
        if int(number_of_pay_stubs) == 0:
            print("You have enterd 0. So i am not creating any paystub. Thanks")
            sys.exit()
        elif int(number_of_pay_stubs) > 0:
            for paystub_number in range(int(number_of_pay_stubs)):
                pay_period = input("Please enter date for Pay Period (Mar 01 2023) : ")
                print("*" * 100)
                print("*" * 100)
                cheque_date = input("Please Enter Cheque date (Feb 11 2023): ")
                print("*" * 100)
                print("*" * 100)
                number_of_hours = int(input("Please Enter Number of Hours : "))
                print("*" * 100)
                print("*" * 100)

                new_year_to_send = pay_period.split(" ")[-1]
                important_values_for_paystub = values_for_paystub(new_year_to_send)

                global federal_first
                federal_first = important_values_for_paystub["federal_first"]
                global province_first
                province_first = important_values_for_paystub["province_first"]
                global federal_second
                federal_second = important_values_for_paystub["federal_second"]
                global province_second
                province_second = important_values_for_paystub["province_second"]
                global federal_three
                federal_three = important_values_for_paystub["federal_three"]
                global province_three
                province_three = important_values_for_paystub["province_three"]
                global federal_four
                federal_four = important_values_for_paystub["federal_four"]
                global province_four
                province_four = important_values_for_paystub["province_four"]
                global federal_five
                federal_five = important_values_for_paystub["federal_five"]
                global province_five
                province_five = important_values_for_paystub["province_five"]
                global EI_Rate
                EI_Rate = important_values_for_paystub["EI_Rate"]
                global CPP_Rate
                CPP_Rate = important_values_for_paystub["CPP_Rate"]
                global EI_Maximum_Deduction
                EI_Maximum_Deduction = important_values_for_paystub[
                    "EI_Maximum_Deduction"
                ]
                global CPP_Maximum_Deduction
                CPP_Maximum_Deduction = important_values_for_paystub[
                    "CPP_Maximum_Deduction"
                ]

                gross_total = number_of_hours * rate_per_hour

                if paystub_number == 0:
                    year_to_date_for_paystub2 = self.calculate_year_to_date(
                        number_of_hours, rate_per_hour, pay_period
                    )
                    last_year_to_date = year_to_date_for_paystub2
                elif paystub_number > 0:
                    year_to_date_for_paystub2 = (
                        self.return_float(last_year_to_date) + gross_total
                    )
                    year_to_date_for_paystub2 = self.comma_seprated(
                        year_to_date_for_paystub2
                    )
                    last_year_to_date = year_to_date_for_paystub2

                y_t_date_input = self.return_float(year_to_date_for_paystub2)

                (
                    year_to_date_incom_tax,
                    total_percentage_for_monthly,
                ) = self.total_incom_tax_calculator_year_to_date(y_t_date_input)
                year_to_date_ei = self.EI_calculator_year_to_date(y_t_date_input)
                year_to_date_cpp = self.CPP_Calculator_year_to_date(y_t_date_input)

                income_tax = self.total_incom_tax_calculator_period(
                    gross_total, total_percentage_for_monthly
                )
                Ei_tax = self.EI_calculator_Period(gross_total, year_to_date_ei)
                cpp_tax = self.CPP_Calculator_Period(gross_total, year_to_date_cpp)

                # This is the total of current calculations
                cur_tot1 = float(income_tax) + float(Ei_tax) + float(cpp_tax)
                current_total = float(gross_total) - cur_tot1

                # This is year to date net calculations
                y_t_d_net_cal1 = (
                    float(year_to_date_incom_tax)
                    + float(year_to_date_cpp)
                    + float(year_to_date_ei)
                )
                total_y_t_d_calculations = (
                    float(self.return_float(year_to_date_for_paystub2)) - y_t_d_net_cal1
                )

                self.making_paystub_two_document(
                    e_name,
                    e_address,
                    occupation,
                    pay_period,
                    cheque_date,
                    number_of_hours,
                    rate_per_hour,
                    gross_total,
                    year_to_date_for_paystub2,
                    year_to_date_cpp,
                    year_to_date_ei,
                    year_to_date_incom_tax,
                    cpp_tax,
                    Ei_tax,
                    income_tax,
                    current_total,
                    total_y_t_d_calculations,
                    cur_tot1,
                    y_t_d_net_cal1,
                    paystub_number,
                )


class PayStubChildONE(PayStubs):
    def date_range(self, input_date):
        # convert input string to datetime object
        date_obj = datetime.strptime(input_date, "%b %d %Y")

        # calculate start and end dates based on input date
        last_day = datetime(date_obj.year, date_obj.month, 1) + timedelta(days=31)
        if last_day.month != date_obj.month:
            end_day = last_day - timedelta(days=last_day.day - 15)
        else:
            end_day = datetime(date_obj.year, date_obj.month, 15)

        if date_obj.day > 15:
            start_day = datetime(date_obj.year, date_obj.month, 16)
            next_month = date_obj.replace(day=28) + timedelta(days=4)
            if next_month.month != date_obj.month:
                end_day = datetime(next_month.year, next_month.month, 1) - timedelta(
                    days=1
                )
        else:
            start_day = datetime(date_obj.year, date_obj.month, 1)
            end_day = datetime(date_obj.year, date_obj.month, 15)
            if start_day.weekday() >= 4:
                end_day = datetime(date_obj.year, date_obj.month, 15)

        # format dates as strings in desired format
        start_str = start_day.strftime("%Y-%m-%d")
        end_str = end_day.strftime("%Y-%m-%d")

        def manipulate_date(date):
            date_obj = datetime.strptime(date, "%Y-%m-%d")
            formatted_date = date_obj.strftime("%B %d, %Y")
            return formatted_date

        start_str = manipulate_date(start_str)
        end_str = manipulate_date(end_str)

        return start_str, end_str

    def checque_date(self, input_str):
        date_obj = datetime.strptime(input_str, "%b %d %Y")
        date_obj = date_obj.strftime("%Y-%m-%d")

        def manipulate_date(date):
            date_obj = datetime.strptime(date, "%Y-%m-%d")
            formatted_date = date_obj.strftime("%B %d, %Y")
            return formatted_date

        cheque_date = manipulate_date(date_obj)
        return cheque_date

    def format_number(self, num):
        return "{:,.2f}".format(float(num))

    def split_name(self, name):
        e_name = name.split(" ")
        first_name = e_name[0]
        last_name = " ".join(e_name[1:])
        return first_name, last_name

    def making_paystub_two_document(
        self,
        e_name_i,
        e_address_i,
        # occupation_i,
        pay_period_i,
        cheque_date_i,
        number_of_hours_i,
        rate_per_hour_i,
        gross_total_i,
        year_to_date_for_paystub2_i,
        year_to_date_cpp_i,
        year_to_date_ei_i,
        year_to_date_incom_tax_i,
        cpp_tax_i,
        Ei_tax_i,
        income_tax_i,
        current_total_i,
        total_y_t_d_calculations_i,
        cur_tot1_i,
        y_t_d_net_cal1_i,
        paystub_number_i,
    ):
        template = "PayStubThree.docx"
        document = MailMerge(template)

        td_object = TD_Document()
        ad_1, ad_2 = td_object.making_address(e_address_i)
        processed_pay_period_1, processed_pay_period_2 = self.date_range(pay_period_i)
        final_cheque_date = self.checque_date(cheque_date_i)
        first_name, last_name = self.split_name(e_name_i)

        ytd_hours_cal = float(self.return_float(year_to_date_for_paystub2_i)) / float(
            rate_per_hour_i
        )

        document.merge(
            employer_name=str(global_employer_name).upper(),
            employer_ad_1=str(global_employer_address_1).upper(),
            empad2=str(global_employer_address_2).upper(),
            employer_name_2=str(global_employer_name).upper(),
            # employer_add_3=str(global_employer_address_1).upper(),
            # employer_add_4=str(global_employer_address_2).upper(),
            f_name=str(first_name).upper(),
            l_name=str(last_name).upper(),
            # employee_name_1=str(e_name_i).upper(),
            employee_add_1=str(ad_1).upper(),
            employee_add_2=str(ad_2).upper(),
            # emp_addr_3=str(e_address_i).upper(),
            # occupation=str(occupation_i).upper(),
            pay_period_1=str(processed_pay_period_1),
            pay_period_2=str(processed_pay_period_2),
            cheque=str(final_cheque_date),
            qty=self.format_number(str(number_of_hours_i)),
            rate=self.format_number(str(rate_per_hour_i)),
            curr=self.format_number(str(gross_total_i)),
            curr1=self.format_number(str(gross_total_i)),
            y_t_d=self.format_number(
                self.return_float(str(year_to_date_for_paystub2_i))
            ),
            y_t_d_1=self.format_number(
                self.return_float(str(year_to_date_for_paystub2_i))
            ),
            ytd_cpp=self.format_number(str(year_to_date_cpp_i)),
            ytd_ei=self.format_number(str(year_to_date_ei_i)),
            ytd_in=self.format_number(str(year_to_date_incom_tax_i)),
            cpp=self.format_number(str(cpp_tax_i)),
            ei=self.format_number(str(Ei_tax_i)),
            inc=self.format_number(str(income_tax_i)),
            cur_net=self.format_number(current_total_i),
            cur_net1=self.format_number(current_total_i),
            cur_net2=self.format_number(current_total_i),
            # y_td_net = self.format_number(total_y_t_d_calculations_i),
            cur_total=self.format_number(str(cur_tot1_i)),
            y_td_tot=self.format_number(str(y_t_d_net_cal1_i)),
            ytd_hours=self.format_number(str(ytd_hours_cal)),
        )

        document.write("Output_File.docx")
        self.convert_to_pdf(f"PayStubThree-{paystub_number_i}")
        try:
            os.remove(os.path.abspath("Output_File.docx"))
        except:
            print("Error in Removing File.")
        return

    def paystub_child_wrapper(self, e_name, e_address):
        print(
            "************************************************************************************"
        )
        print(
            "************************** We are Doing PayStub 3 ********************************"
        )
        print(
            "************************************************************************************"
        )

        rate_per_hour = int(input("Please Enter Rate per Hour : "))
        print("*" * 100)
        print("*" * 100)

        number_of_pay_stubs = input(
            "Please enter, how many number of paystubs you want to create: "
        )
        if int(number_of_pay_stubs) == 0:
            print("You have enterd 0. So i am not creating any paystub. Thanks")
            sys.exit()
        elif int(number_of_pay_stubs) > 0:
            for paystub_number in range(int(number_of_pay_stubs)):
                # occupation = input("Please Enter Occupation : ")
                # print("*" * 100)
                # print("*" * 100)
                pay_period = input("Please enter date for Pay Period (Mar 01 2023) : ")
                print("*" * 100)
                print("*" * 100)
                cheque_date = input("Please Enter Cheque date (Feb 11 2023): ")
                print("*" * 100)
                print("*" * 100)
                number_of_hours = int(input("Please Enter Number of Hours : "))
                print("*" * 100)
                print("*" * 100)

                new_year_to_send = pay_period.split(" ")[-1]
                important_values_for_paystub = values_for_paystub(new_year_to_send)

                global federal_first
                federal_first = important_values_for_paystub["federal_first"]
                global province_first
                province_first = important_values_for_paystub["province_first"]
                global federal_second
                federal_second = important_values_for_paystub["federal_second"]
                global province_second
                province_second = important_values_for_paystub["province_second"]
                global federal_three
                federal_three = important_values_for_paystub["federal_three"]
                global province_three
                province_three = important_values_for_paystub["province_three"]
                global federal_four
                federal_four = important_values_for_paystub["federal_four"]
                global province_four
                province_four = important_values_for_paystub["province_four"]
                global federal_five
                federal_five = important_values_for_paystub["federal_five"]
                global province_five
                province_five = important_values_for_paystub["province_five"]
                global EI_Rate
                EI_Rate = important_values_for_paystub["EI_Rate"]
                global CPP_Rate
                CPP_Rate = important_values_for_paystub["CPP_Rate"]
                global EI_Maximum_Deduction
                EI_Maximum_Deduction = important_values_for_paystub[
                    "EI_Maximum_Deduction"
                ]
                global CPP_Maximum_Deduction
                CPP_Maximum_Deduction = important_values_for_paystub[
                    "CPP_Maximum_Deduction"
                ]

                gross_total = number_of_hours * rate_per_hour

                if paystub_number == 0:
                    year_to_date_for_paystub2 = self.calculate_year_to_date(
                        number_of_hours, rate_per_hour, pay_period
                    )
                    last_year_to_date = year_to_date_for_paystub2
                elif paystub_number > 0:
                    year_to_date_for_paystub2 = (
                        self.return_float(last_year_to_date) + gross_total
                    )
                    year_to_date_for_paystub2 = self.comma_seprated(
                        year_to_date_for_paystub2
                    )
                    last_year_to_date = year_to_date_for_paystub2

                y_t_date_input = self.return_float(year_to_date_for_paystub2)

                (
                    year_to_date_incom_tax,
                    total_percentage_for_monthly,
                ) = self.total_incom_tax_calculator_year_to_date(y_t_date_input)
                year_to_date_ei = self.EI_calculator_year_to_date(y_t_date_input)
                year_to_date_cpp = self.CPP_Calculator_year_to_date(y_t_date_input)

                income_tax = self.total_incom_tax_calculator_period(
                    gross_total, total_percentage_for_monthly
                )
                Ei_tax = self.EI_calculator_Period(gross_total, year_to_date_ei)
                cpp_tax = self.CPP_Calculator_Period(gross_total, year_to_date_cpp)

                # This is the total of current calculations
                cur_tot1 = float(income_tax) + float(Ei_tax) + float(cpp_tax)
                current_total = float(gross_total) - cur_tot1

                # This is year to date net calculations
                y_t_d_net_cal1 = (
                    float(year_to_date_incom_tax)
                    + float(year_to_date_cpp)
                    + float(year_to_date_ei)
                )
                total_y_t_d_calculations = (
                    float(self.return_float(year_to_date_for_paystub2)) - y_t_d_net_cal1
                )

                self.making_paystub_two_document(
                    e_name,
                    e_address,
                    # occupation,
                    pay_period,
                    cheque_date,
                    number_of_hours,
                    rate_per_hour,
                    gross_total,
                    year_to_date_for_paystub2,
                    year_to_date_cpp,
                    year_to_date_ei,
                    year_to_date_incom_tax,
                    cpp_tax,
                    Ei_tax,
                    income_tax,
                    current_total,
                    total_y_t_d_calculations,
                    cur_tot1,
                    y_t_d_net_cal1,
                    paystub_number,
                )


#######################################################################################################################################
#######################################################################################################################################
#######################################################################################################################################
#######################################################################################################################################

if __name__ == "__main__":
    df = pd.read_excel("Global_Variables.xlsx")
    for index, row in df.iterrows():
        print("***************************")
        print("***************************")
        global_name = row["Employee_Name"]
        if pd.isna(row["Employee_Name"]):
            global_name = input("Please enter Employee name: ").upper()
        print(f"We are doing document of {global_name}")
        print("***************************")
        print("***************************")
        global_employee_address = row["Employee_Address"]
        if pd.isna(row["Employee_Address"]):
            global_employee_address = input("Please enter Employee address: ").upper()
        # print("***************************")
        # print("***************************")
        global_employer_name = row["Employer_Name"]
        if pd.isna(row["Employer_Name"]):
            global_employer_name = input("Please Enter Employer Name : ").upper()
        # print("***************************")
        # print("***************************")
        employer_address = row["Employer_Address"]
        global_employer_address_1, global_employer_address_2 = making_address(
            employer_address
        )
        if pd.isna(row["Employer_Address"]):
            employer_address = input("Please Enter Employer Address: ").upper()
            global_employer_address_1, global_1employer_address_2 = making_address(
                employer_address
            )
        # print("***************************")
        # print("***************************")
        global_sin_number = str(row["SIN_Number"])
        if pd.isna(row["SIN_Number"]):
            global_sin_number = str(input("Please Enter SIN number : "))

        # print(row['doc_options'])
        # print(type(row['doc_options']))
        options_feature(global_name, global_employee_address, row)


# 1680 Richmond St
# LONDON, ON N6G 3Y9


# 57-1478 ADELAIDE ST N
# LONDON, ON N5X 3Y1
