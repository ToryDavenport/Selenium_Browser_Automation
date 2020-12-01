from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import *
import time
from time import sleep

driver = webdriver.Chrome()

driver.get("http://www.ascomwireless.com/rma")

assert "Return Material Authorization Form" in driver.title

#####################################################################################
# Fill form information                          									#
#####################################################################################
first_name = driver.find_element_by_name("txtReqFirstName")
first_name.send_keys("Tory")

last_name = driver.find_element_by_name("txtReqLastName")
last_name.send_keys("Davenport")

company = driver.find_element_by_name("txtReqCompany")
company.send_keys("First Light")

customer_number = driver.find_element_by_name("txtAscomCustomerNumber")
customer_number.send_keys("AC3219")

end_user = driver.find_element_by_name("txtReqEndUserName")
end_user.send_keys("Wegmans")

phone_num = driver.find_element_by_name("txtReqPhone")
phone_num.send_keys("5854336649")

email = driver.find_element_by_name("txtReqEmail")
email.send_keys("wegmanstickets@firstlight.net")

protection_plan = driver.find_element_by_name("txtReqPPP")
protection_plan.send_keys("PN-002484")

ship_company_name = driver.find_element_by_name("txtShipCompanyName")
ship_company_name.send_keys("First Light")

ship_street = driver.find_element_by_name("txtShipStreet1")
ship_street.send_keys("7890 Lehigh Crossing")

ship_city = driver.find_element_by_name("txtShipCity")
ship_city.send_keys("Victor")

ship_state = driver.find_element_by_name("txtShipState")
ship_state.send_keys("NY")

ship_zip = driver.find_element_by_name("txtShipZip")
ship_zip.send_keys("14564")

#####################################################################################
# Fill in Form DATA (This will eventually be automated by pulling data from excel)  #
#####################################################################################


# MODEL TYPE 1 - 10
row_1_model_1 = driver.find_element_by_name("ddlModelType1")
row_1_model_1.send_keys("D63")

row_2_model_2 = driver.find_element_by_name("ddlModelType2")
row_2_model_2.send_keys("D63")

row_3_model_3 = driver.find_element_by_name("ddlModelType3")
row_3_model_3.send_keys("D63")

row_4_model_4 = driver.find_element_by_name("ddlModelType4")
row_4_model_4.send_keys("D63")

row_5_model_5 = driver.find_element_by_name("ddlModelType5")
row_5_model_5.send_keys("D63")

row_6_model_6 = driver.find_element_by_name("ddlModelType6")
row_6_model_6.send_keys("D63")

row_7_model_7 = driver.find_element_by_name("ddlModelType7")
row_7_model_7.send_keys("D63")

row_8_model_8 = driver.find_element_by_name("ddlModelType8")
row_8_model_8.send_keys("D63")

row_9_model_9 = driver.find_element_by_name("ddlModelType9")
row_9_model_9.send_keys("D63")

row_10_model_10 = driver.find_element_by_name("ddlModelType10")
row_10_model_10.send_keys("D63")

# Fill Serial's
serial_1 = driver.find_element_by_name("txtSerial1")
serial_1.send_keys("")

serial_2 = driver.find_element_by_name("txtSerial2")
serial_2.send_keys("")

serial_3 = driver.find_element_by_name("txtSerial3")
serial_3.send_keys("")

serial_4 = driver.find_element_by_name("txtSerial4")
serial_4.send_keys("")

serial_5 = driver.find_element_by_name("txtSerial5")
serial_5.send_keys("")

serial_6 = driver.find_element_by_name("txtSerial6")
serial_6.send_keys("")

serial_7 = driver.find_element_by_name("txtSerial7")
serial_7.send_keys("")

serial_8 = driver.find_element_by_name("txtSerial8")
serial_8.send_keys("")

serial_9 = driver.find_element_by_name("txtSerial9")
serial_9.send_keys("")

serial_10 = driver.find_element_by_name("txtSerial10")
serial_10.send_keys("")


# Required Variables
# This allows you to enter two problem descriptions per item
""" 
MAP THESE LIST ITEMS TO CELLS ON THE SPREAD SHEET
"""
problem_1 = ['','']
problem_2 = ['','']
problem_3 = ['','']
problem_4 = ['','']
problem_5 = ['','']
problem_6 = ['','']
problem_7 = ['','']
problem_8 = ['','']
problem_9 = ['','']
problem_10 = ['','']

# Fill Problem Descriptions 1 and 2 for each of the 10 rows
problem_descr1_row1 = driver.find_element_by_name("ddlProblemDescr1_row1")
problem_descr1_row1.send_keys(problem_1[0])

problem_descr2_row1 = driver.find_element_by_name("ddlProblemDescr2_row1")
problem_descr2_row1.send_keys(problem_1[1])

problem_descr1_row2 = driver.find_element_by_name("ddlProblemDescr1_row2")
problem_descr1_row2.send_keys(problem_2[0])

problem_descr2_row2 = driver.find_element_by_name("ddlProblemDescr2_row2")
problem_descr2_row2.send_keys(problem_2[1])

problem_descr1_row3 = driver.find_element_by_name("ddlProblemDescr1_row3")
problem_descr1_row3.send_keys(problem_3[0])

problem_descr2_row3 = driver.find_element_by_name("ddlProblemDescr2_row3")
problem_descr2_row3.send_keys(problem_3[1])

problem_descr1_row4 = driver.find_element_by_name("ddlProblemDescr1_row4")
problem_descr1_row4.send_keys(problem_4[0])

problem_descr2_row4 = driver.find_element_by_name("ddlProblemDescr2_row4")
problem_descr2_row4.send_keys(problem_4[1])

problem_descr1_row5 = driver.find_element_by_name("ddlProblemDescr1_row5")
problem_descr1_row5.send_keys(problem_5[0])

problem_descr2_row5 = driver.find_element_by_name("ddlProblemDescr2_row5")
problem_descr2_row5.send_keys(problem_5[1])

problem_descr1_row6 = driver.find_element_by_name("ddlProblemDescr1_row6")
problem_descr1_row6.send_keys(problem_6[0])

problem_descr2_row6 = driver.find_element_by_name("ddlProblemDescr2_row6")
problem_descr2_row6.send_keys(problem_6[1])

problem_descr1_row7 = driver.find_element_by_name("ddlProblemDescr1_row7")
problem_descr1_row7.send_keys(problem_7[0])

problem_descr2_row7 = driver.find_element_by_name("ddlProblemDescr2_row7")
problem_descr2_row7.send_keys(problem_7[1])

problem_descr1_row8 = driver.find_element_by_name("ddlProblemDescr1_row8")
problem_descr1_row8.send_keys(problem_8[0])

problem_descr2_row8 = driver.find_element_by_name("ddlProblemDescr2_row8")
problem_descr2_row8.send_keys(problem_8[1])

problem_descr1_row9 = driver.find_element_by_name("ddlProblemDescr1_row9")
problem_descr1_row9.send_keys(problem_9[0])

problem_descr2_row9 = driver.find_element_by_name("ddlProblemDescr2_row9")
problem_descr2_row9.send_keys(problem_9[1])

problem_descr1_row10 = driver.find_element_by_name("ddlProblemDescr1_row10")
problem_descr1_row10.send_keys(problem_10[0])

problem_descr2_row10 = driver.find_element_by_name("ddlProblemDescr2_row10")
problem_descr2_row10.send_keys(problem_10[1])
input("Press enter to continue")

""" Check boxes for Out of Box Warrenty Column
all_checked = False

if all_checked == True:
	chkBoxFail1 = driver.find_element_by_name("chkBoxFail1").click()
	chkBoxFail2 = driver.find_element_by_name("chkBoxFail2").click()
	chkBoxFail3 = driver.find_element_by_name("chkBoxFail3").click()
	chkBoxFail4 = driver.find_element_by_name("chkBoxFail4").click()
	chkBoxFail5 = driver.find_element_by_name("chkBoxFail5").click()
	chkBoxFail6 = driver.find_element_by_name("chkBoxFail6").click()
	chkBoxFail7 = driver.find_element_by_name("chkBoxFail7").click()
	chkBoxFail8 = driver.find_element_by_name("chkBoxFail8").click()
	chkBoxFail9 = driver.find_element_by_name("chkBoxFail9").click()
	chkBoxFail10 = driver.find_element_by_name("chkBoxFail10").click()

else:
	pass"""

# QUIT WHEN DONE
driver.quit()


