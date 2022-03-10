import time
import sys
import csv
import os
import tkinter as tk

from tkinter import messagebox
from selenium import webdriver
from datetime import date
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select


def CheckForSubmitErrors(prod_type,firstname,lastname,saledate):

    if ("The contracts listed above have the same Contract ID as the one trying to be written." in driver.page_source):
        msg = "%s Duplicate entry: Replace with %s [%s %s] ?" % (prod_type,saledate,firstname,lastname)
        print(msg)
        MsgBox = tk.messagebox.askyesno('Duplicate Entry Detected',msg, icon='warning')
        print("Message Box Answer = [",MsgBox,"]")
        if (MsgBox == True):
            print("Clicking on Continue with New App")
            driver.find_element_by_link_text('Continue With New App').click()
            return "OK"
        else:
            return "DUPLICATE ENTRY"

    if ("There are no rates available for this vehicle." in driver.page_source):
        msg = "%s Manual entry required: %s [%s %s]" % (prod_type, saledate, firstname, lastname)
        return "MANUAL_ENTRY"

    return "OK"

def LookUpStateFromAbbreviation(abbev):
    if (abbev == "AL"):
        return "Alabama"
    elif (abbev == "AK"):
        return "Alaska"
    elif (abbev == "AZ"):
        return "Arizona"
    elif (abbev == "AR"):
        return "Arkansas"
    elif (abbev == "CA"):
        return "California"
    elif (abbev == "CO"):
        return "Colorado"
    elif (abbev == "CT"):
        return "Connecticut"
    elif (abbev == "DE"):
        return "Delaware"
    elif (abbev == "FL"):
        return "Florida"
    elif (abbev == "GA"):
        return "Georgia"
    elif (abbev == "HI"):
        return "Hawaii"
    elif (abbev == "ID"):
        return "Idaho"
    elif (abbev == "IL"):
        return "Illinois"
    elif (abbev == "IN"):
        return "Indiana"
    elif (abbev == "IA"):
        return "Iowa"
    elif (abbev == "KS"):
        return "Kansas"
    elif (abbev == "KY"):
        return "Kentucky"
    elif (abbev == "LA"):
        return "Louisiana"
    elif (abbev == "ME"):
        return "Maine"
    elif (abbev == "MD"):
        return "Maryland"
    elif (abbev == "MA"):
        return "Massachusetts"
    elif (abbev == "MI"):
        return "Michigan"
    elif (abbev == "MN"):
        return "Minnesota"
    elif (abbev == "MS"):
        return "Mississippi"
    elif (abbev == "MO"):
        return "Missouri"
    elif (abbev == "MT"):
        return "Montana"
    elif (abbev == "NE"):
        return "Nebraska"
    elif (abbev == "NV"):
        return "Nevada"
    elif (abbev == "NH"):
        return "New Hampshire"
    elif (abbev == "NJ"):
        return "New Jersey"
    elif (abbev == "NM"):
        return "New Mexico"
    elif (abbev == "NY"):
        return "New York"
    elif (abbev == "NC"):
        return "North Carolina"
    elif (abbev == "ND"):
        return "North Dakota"
    elif (abbev == "OK"):
        return "Oklahoma"
    elif (abbev == "PA"):
        return "Pennsylvania"
    elif (abbev == "RI"):
        return "Rhode Island"
    elif (abbev == "SC"):
        return "South Carolina"
    elif (abbev == "SD"):
        return "South Dakota"
    elif (abbev == "TN"):
        return "Tennessee"
    elif (abbev == "TX"):
        return "Texas"
    elif (abbev == "UT"):
        return "Utah"
    elif (abbev == "VT"):
        return "Vermont"
    elif (abbev == "VA"):
        return "Virgina"
    elif (abbev == "WA"):
        return "Washington"
    elif (abbev == "WV"):
        return "West Virgina"
    elif (abbev == "WI"):
        return "Wisconsin"
    elif (abbev == "WY"):
        return "Wyoming"
    else:
        return "ERROR"

def EnterProductInfo(vin,prod_type,saledate,firstname,lastname,address,city,state,zip,phone,saleprice):
    print("%s:Entering customer information: %s" % (prod_type,vin))
    state = LookUpStateFromAbbreviation(state)
    driver.find_element_by_id('txtFirstName').send_keys(firstname)
    driver.find_element_by_id('txtLastName').send_keys(lastname)
    driver.find_element_by_id('txtPurchaseDate').click()
    driver.find_element_by_id('txtPurchaseDate').send_keys(saledate)
    driver.find_element_by_id('txtPurchaseDate').send_keys(Keys.TAB)
    driver.find_element_by_id('txtClientAddress').send_keys(address)
    driver.find_element_by_id('txtCity').send_keys(city)
    Select(driver.find_element_by_id('ddlState')).select_by_visible_text(state)
    #Zip code has difficulties / enter 3x times
    driver.implicitly_wait(5)
    driver.find_element_by_id('txtZip').send_keys(Keys.TAB)
    driver.implicitly_wait(3);
    driver.find_element_by_id('txtZip').send_keys(zip)
    driver.implicitly_wait(10);
    driver.find_element_by_id('txtZip').send_keys(zip)
    driver.implicitly_wait(10);
    driver.find_element_by_id('txtZip').send_keys(zip)
    driver.implicitly_wait(3);
    driver.find_element_by_id('txtPhone').send_keys(phone)
    driver.find_element_by_id('txtVehicleCost2').send_keys(saleprice)
    
    # DEBUG - comment this out to not submit
    driver.find_element_by_id('btnLookupSender').click()
    print("%s product: [%s %s] submitted!" % (prod_type,firstname,lastname))
    
    #if ("VALIDATION ERROR" in driver.page_source):
    #    msg = "Detected Validation Error"
    #    print(msg)
    #    MsgBox = tk.messagebox.askyesno('Validation Error Detected',msg, icon='warning')

    HTML_OUTPUT_FILE.write("<td align=\"center\">OK</td>\n")
    HTML_OUTPUT_FILE.write("</tr>\n")
    return


def EnterBasicWarranty(vin,odometer,firstname,lastname,saledate):
    driver.find_element_by_id('GridView1_lnkDealerID_0').click()
    driver.find_element_by_id('ContentPlaceHolder1_txtVIN').send_keys(vin)
    driver.find_element_by_id('ContentPlaceHolder1_txtMileage').send_keys(odometer)
    Select(driver.find_element_by_id('ContentPlaceHolder1_ddlNewUsed')).select_by_visible_text('Used')
    driver.find_element_by_id('ContentPlaceHolder1_btnSubmit').click()
    driver.implicitly_wait(3);

    # Check for submit errors
    answer = CheckForSubmitErrors("WARRANTY",firstname,lastname,saledate)
    print("WARRANTY:CheckForSubmitErrors:ANSWER -> [%s]" % (answer))

    if (answer != "OK"):
        HTML_OUTPUT_FILE.write("<td align=\"center\"><font color=red>%s</font></td>\n" % (answer))
        HTML_OUTPUT_FILE.write("</tr>\n")
        driver.find_element_by_link_text('Contracts').click()
        driver.implicitly_wait(3);
        driver.find_element_by_xpath('//li/span').click()
        driver.implicitly_wait(3);
        return "ERROR"

    driver.find_element_by_xpath('//td/input').click()  # 3 Months / 3000 Miles
    driver.find_element_by_id('lnkNext').click()
    driver.implicitly_wait(3)
    return "OK"

def EnterBasicGap(vin,odometer,prod_months,firstname,lastname,saledate):
    driver.find_element_by_id('GridView1_lnkDealerID_1').click()
    driver.find_element_by_id('ContentPlaceHolder1_txtVIN').send_keys(vin)
    driver.find_element_by_id('ContentPlaceHolder1_txtMileage').send_keys(odometer)
    Select(driver.find_element_by_id('ContentPlaceHolder1_ddlNewUsed')).select_by_visible_text('Used')
    driver.implicitly_wait(3)
    driver.find_element_by_id('ContentPlaceHolder1_btnSubmit').click()
    

    # Check for submit errors
    answer = CheckForSubmitErrors("GAP",firstname,lastname,saledate)
    print("GAP:EnterBasicGap:CheckForSubmitErrors:ANSWER -> [%s]" % (answer))

    if (answer != "OK"):
        HTML_OUTPUT_FILE.write("<td align=\"center\"><font color=red>%s</font></td>\n" % (answer))
        HTML_OUTPUT_FILE.write("</tr>\n")
        driver.find_element_by_link_text('Contracts').click()
        driver.implicitly_wait(3);
        driver.find_element_by_xpath('//li/span').click()
        driver.implicitly_wait(3);
        return "ERROR"

    driver.implicitly_wait(3)
    print("GAP:EnterBasicGap: Product Months -> [%s]" % (prod_months))

    if (prod_months == "12"):
        driver.find_element_by_xpath('//td/input').click()
    elif( prod_months == "24" ):
        driver.find_element_by_xpath('//tr[3]/td/input').click()
    elif( prod_months == "36" ):
        driver.find_element_by_xpath('//tr[4]/td/input').click()
    elif (prod_months == "48"):
        driver.find_element_by_xpath('//tr[5]/td/input').click()
    else:
        msg = "GAP:EnterBasicGap: Product Months Invalid -> [%s months] " % (prod_months)
        print(msg)
        tk.messagebox.showinfo('Error',msg)
        
    driver.find_element_by_id('lnkNext').click()
    driver.implicitly_wait(3)
    return "OK"


def EnterHighMileageWarranty(vin,odometer,prod_months,firstname,lastname,saledate):
    driver.find_element_by_id('GridView1_lnkDealerID_6').click()
    driver.find_element_by_id('ContentPlaceHolder1_txtVIN').send_keys(vin)
    driver.find_element_by_id('ContentPlaceHolder1_txtMileage').send_keys(odometer)
    Select(driver.find_element_by_id('ContentPlaceHolder1_ddlNewUsed')).select_by_visible_text('Used')
    driver.find_element_by_id('ContentPlaceHolder1_btnSubmit').click()
    driver.implicitly_wait(3);

    # Check for submit errors
    answer = CheckForSubmitErrors("WARRANTY",firstname,lastname,saledate)
    print("WARRANTY:CheckForSubmitErrors:ANSWER -> [%s]" % (answer))
    driver.implicitly_wait(3);

    if (answer != "OK"):
        HTML_OUTPUT_FILE.write("<td align=\"center\"><font color=red>%s</font></td>\n" % (answer))
        HTML_OUTPUT_FILE.write("</tr>\n")
        driver.find_element_by_link_text('Contracts').click()
        driver.implicitly_wait(3);
        driver.find_element_by_xpath('//li/span').click()
        driver.implicitly_wait(3);
        return "ERROR"

    print("HIGH MILEAGE WARRANTY - MONTHS = %d\n" % (prod_months))
    if (prod_months == "12"):
        driver.find_element_by_xpath('//td/input').click()
    elif( prod_months == "24" ):
        driver.find_element_by_xpath('//tr[3]/td/input').click()
    elif( prod_months == "36" ):
        driver.find_element_by_xpath('//tr[5]/td/input').click()

    driver.find_element_by_id('lnkNext').click()
    driver.implicitly_wait(3);
    return "OK"

def EnterAVP3Warranty(vin,odometer,prod_months,prod_miles,prod_deductible,firstname,lastname,saledate):
    custom_warranty = 0
    driver.find_element_by_id('GridView1_lnkDealerID_5').click()
    driver.find_element_by_id('ContentPlaceHolder1_txtVIN').send_keys(vin)
    driver.find_element_by_id('ContentPlaceHolder1_txtMileage').send_keys(odometer)

    # For warranties with coverage mileage >= 50k , pick CUSTOM
    if (int(prod_miles) >= 50000 ):
        Select(driver.find_element_by_id('ContentPlaceHolder1_ddlNewUsed')).select_by_visible_text('Program')
        custom_warranty = 1
    # For warranties with coverage miles < 50k, pick USED
    else:
        Select(driver.find_element_by_id('ContentPlaceHolder1_ddlNewUsed')).select_by_visible_text('Used')
        custom_warranty = 0

    driver.find_element_by_id('ContentPlaceHolder1_btnSubmit').click()
    driver.implicitly_wait(3);

    # Check for submit errors
    answer = CheckForSubmitErrors("WARRANTY",firstname,lastname,saledate)
    print("WARRANTY:CheckForSubmitErrors:ANSWER -> [%s]" % (answer))
    driver.implicitly_wait(3);

    if (answer != "OK"):
        HTML_OUTPUT_FILE.write("<td align=\"center\"><font color=red>%s</font></td>\n" % (answer))
        HTML_OUTPUT_FILE.write("</tr>\n")
        driver.find_element_by_link_text('Contracts').click()
        driver.implicitly_wait(3);
        driver.find_element_by_xpath('//li/span').click()
        driver.implicitly_wait(3);
        return "ERROR"

    print("WARRANTY - DEDUCTIBLE = %s\n" % (prod_deductible))
    if ( prod_deductible == "0" ):
        Select(driver.find_element_by_id('ddlDeductibleOption')).select_by_visible_text('Zero Deductible (Add $100 to Price)')
    elif ( prod_deductible == "50" ):
        Select(driver.find_element_by_id('ddlDeductibleOption')).select_by_visible_text('$50 Deductible (Add $50 to Price)')
    elif ( prod_deductible == "100" ):
        Select(driver.find_element_by_id('ddlDeductibleOption')).select_by_visible_text('$100 Disappearing Deductible (Add $75 to Price)')
    elif ( prod_deductible == "200" ):
        Select(driver.find_element_by_id('ddlDeductibleOption')).select_by_visible_text('$200 Deductible ( NO CHARGE )')

    if ( custom_warranty > 0 ):
        print("CUSTOM WARRANTY - MONTHS = %s\n" % (prod_months))
        if ( prod_miles == "50000" and prod_months == "36" ):
            driver.find_element_by_xpath('//tr[1]/td/input').click()
        elif ( prod_miles == "50000" and prod_months == "48" ):
            driver.find_element_by_xpath('//tr[2]/td/input').click()
        elif ( prod_miles == "50000" and prod_months == "60" ):
            driver.find_element_by_xpath('//tr[3]/td/input').click()
        elif ( prod_miles == "60000" and prod_months == "36" ):
            driver.find_element_by_xpath('//tr[13]/td/input').click()
        elif ( prod_miles == "60000" and prod_months == "48" ):
            driver.find_element_by_xpath('//tr[14]/td/input').click()
        elif ( prod_miles == "60000" and prod_months == "60" ):
            driver.find_element_by_xpath('//tr[15]/td/input').click()
    else:
        print("WARRANTY - MONTHS = %s\n" % (prod_months))
        if ( int(odometer) >= 85000 ):
            if ( prod_months == "12" ):
                driver.find_element_by_xpath('//td/input').click()
            elif ( prod_months == "18" ):
                driver.find_element_by_xpath('//tr[4]/td/input').click()
            elif ( prod_months == "24" ):
                driver.find_element_by_xpath('//tr[6]/td/input').click()
            elif ( prod_months == "36" ):
                driver.find_element_by_xpath('//tr[7]/td/input').click()
        else:
            if ( prod_months == "12" ):
                driver.find_element_by_xpath('//td/input').click();
            elif ( prod_months == "18" ):
                driver.find_element_by_xpath('//tr[5]/td/input').click()
            elif ( prod_months == "24" ):
                driver.find_element_by_xpath('//tr[8]/td/input').click()
            elif ( prod_months == "36" ):
                driver.find_element_by_xpath('//tr[11]/td/input').click()

    driver.find_element_by_id('lnkNext').click()
    driver.implicitly_wait(3);
    return "OK"


def main(inputFilename):
    print(inputFilename);
    HTML_OUTPUT_FILE.write("<html>\n");
    HTML_OUTPUT_FILE.write("<head><title>Pro Fin AVP eXpress</title></head>\n");
    HTML_OUTPUT_FILE.write("<body>\n");
    HTML_OUTPUT_FILE.write("<div style=\"display:block;text-align:left\"><a href=\"https://login.assuredvehicleprotection.com/Login.aspx\" imageanchor=1><img align=\"left\" src=\"avp.jpg\" border=0></a><h1><I>ProFin AVP eXpress</I></h1>");
    HTML_OUTPUT_FILE.write("<head><style>\n");
    HTML_OUTPUT_FILE.write("table  { width:80%;}\n");
    HTML_OUTPUT_FILE.write("th, td { padding: 10px;}\n");
    HTML_OUTPUT_FILE.write("table#table01 tr:nth-child(even) { background-color: #eee; }\n");
    HTML_OUTPUT_FILE.write("table#table01 tr:nth-child(odd)  { background-color: #fff; }\n");
    HTML_OUTPUT_FILE.write("table#table01 th { background-color: #084B8A; color: white; }\n");
    HTML_OUTPUT_FILE.write("</style></head>\n");
    HTML_OUTPUT_FILE.write("<table border=5 id=\"table01\" >\n");
    HTML_OUTPUT_FILE.write("<tr><th>Index</th><th>Product</th><th>Months<th>Mileage</th><th>Deductible</th><th>Sale Date</th><th>VIN</th><th>Odometer</th><th>Name</th><th>Result</th></tr>\n");

    #driver = webdriver.Chrome();
    driver.get('https://login.assuredvehicleprotection.com/Login.aspx');
    driver.implicitly_wait(3);
    driver.find_element_by_name('btnCookieConsent').click();
    driver.find_element_by_name('txtUserName').send_keys('jmcqueenJTR');
    driver.find_element_by_name('txtPassword').send_keys('jmcqueenJTR');
    driver.find_element_by_name('btnLogin').click();
    driver.implicitly_wait(3);
    driver.switch_to.frame(0);
    driver.find_element_by_id('newContract').click();
    loop_iteration = 0;

    with open(sys.argv[1], newline='') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='|')
        for row in spamreader:
            prod_type = row[0];
            prod_months = row[1];
            prod_miles = row[2];
            prod_deductible = row[3];
            vin = row[4];
            odometer = row[5];
            firstname = row[6];
            lastname = row[7];
            saledate = row[8];
            address = row[9];
            city = row[10];
            state = row[11];
            zip = row[12];
            phone = row[13];
            price = row[14];
            high_mileage = row[15];

            prod_type = prod_type.strip().upper()
            firstname = firstname.strip().upper()
            lastname  = lastname.strip().upper()
            high_mileage = high_mileage.strip().upper()
            month,day,year = saledate.split('/')
            year = "20%s" % year
            saledate = "%02d/%02d/%s" % (int(month),int(day),year)

            # DEBUG
            print(prod_type,prod_months,prod_miles,prod_deductible,vin,odometer,firstname,lastname,saledate,address,city,state,zip,phone,price,high_mileage)
            loop_iteration = loop_iteration + 1

            ##
            ## WARRANTY
            ##
            if (prod_type == "WARRANTY") :
                HTML_OUTPUT_FILE.write("<tr>\n");
                HTML_OUTPUT_FILE.write("<td align=\"center\">%s</td>\n" % (loop_iteration));
                HTML_OUTPUT_FILE.write("<td align=\"center\">WARRANTY</td>\n");
                HTML_OUTPUT_FILE.write("<td align=\"center\">%s</td>\n" % (prod_months));
                HTML_OUTPUT_FILE.write("<td align=\"center\">%s</td>\n" % (prod_miles));
                HTML_OUTPUT_FILE.write("<td align=\"center\">%s</td>\n" % (prod_deductible));
                HTML_OUTPUT_FILE.write("<td align=\"center\">%s</td>\n" % (saledate));
                HTML_OUTPUT_FILE.write("<td align=\"center\">%s</td>\n" % (vin));
                HTML_OUTPUT_FILE.write("<td align=\"center\">%s</td>\n" % (odometer));
                HTML_OUTPUT_FILE.write("<td align=\"center\">%s %s</td>\n" % (firstname,lastname));

                driver.implicitly_wait(5)
                    
                # Pro Fin Lending Solutions Warranty 3/3
                if (prod_months == "3") :
                    print("WARRANTY 3/3")
                    result = EnterBasicWarranty(vin,odometer,firstname,lastname,saledate)

                # Pro Fin Lending Solutions AVP3
                elif(high_mileage == "NO") :
                    print("WARRANTY - AVP3")
                    result = EnterAVP3Warranty(vin,odometer,prod_months,prod_miles,prod_deductible,firstname,lastname,saledate);

                # Pro Fin Lending Solutions SPL - HIGH MILEAGE
                elif(high_mileage == "YES") :
                    print("WARRANTY - HIGH MILEAGE")
                    result = EnterHighMileageWarranty(vin,odometer,prod_months,firstname,lastname,saledate)

                if (result == "OK"):
                    EnterProductInfo(vin,prod_type,saledate,firstname,lastname,address,city,state,zip,phone,price)
            ##
            ## End of Warranty
            ##

            ##
            ## GAP
            ##
            if (prod_type == "GAP") :
                HTML_OUTPUT_FILE.write("<tr>\n")
                HTML_OUTPUT_FILE.write("<td align=\"center\">%s</td>\n" % (loop_iteration))
                HTML_OUTPUT_FILE.write("<td align=\"center\">GAP</td>\n")
                HTML_OUTPUT_FILE.write("<td align=\"center\">%s</td>\n" % (prod_months))
                HTML_OUTPUT_FILE.write("<td align=\"center\">%s</td>\n" % (prod_miles))
                HTML_OUTPUT_FILE.write("<td align=\"center\">%s</td>\n" % (prod_deductible))
                HTML_OUTPUT_FILE.write("<td align=\"center\">%s</td>\n" % (saledate))
                HTML_OUTPUT_FILE.write("<td align=\"center\">%s</td>\n" % (vin))
                HTML_OUTPUT_FILE.write("<td align=\"center\">%s</td>\n" % (odometer))
                HTML_OUTPUT_FILE.write("<td align=\"center\">%s %s</td>\n" % (firstname,lastname))
                
                driver.implicitly_wait(5)

                # Pro Fin Lending Solutions GAP
                result = EnterBasicGap(vin,odometer,prod_months,firstname,lastname,saledate)

                driver.implicitly_wait(5)
                
                if (result == "OK"):
                    EnterProductInfo(vin, prod_type, saledate, firstname, lastname, address, city, state, zip, phone, price)
            # End of GAP

            # Setup for the next row entry...
            if (result == "OK"):
                driver.implicitly_wait(5)
                driver.find_element_by_link_text('Contracts').click()
                driver.implicitly_wait(3)
                driver.find_element_by_xpath('//li/span').click()
                
    driver.implicitly_wait(5)
    print("Entries complete, show Pending Contracts...\n");
    driver.find_element_by_link_text('Home').click()
    driver.implicitly_wait(3)
    driver.find_element_by_xpath('//form[@id=\'form1\']/div[6]/a/div/div[2]').click()
    driver.implicitly_wait(3)
    driver.find_element_by_link_text('+ Pending Contracts').click()
    driver.implicitly_wait(3)
    driver.find_element_by_id('btnGetDate').click();
    driver.implicitly_wait(3)

    HTML_OUTPUT_FILE.write("<br><br><br>\n")
    HTML_OUTPUT_FILE.write("</table>\n")
    HTML_OUTPUT_FILE.write("</form>\n")
    HTML_OUTPUT_FILE.write("</body>\n")
    HTML_OUTPUT_FILE.write("</html>\n")

    print("Closing up program...")
    HTML_OUTPUT_FILE.close()

    # Kick off chrome browser with results
    command = "start chrome %s" % (htmlOutputFilename)
    os.system(command);
    time.sleep(30)
    driver.quit()
    exit()



##
##  Program Start
##
today = date.today()
htmlOutputFilename = "%s\ProvidentAvpFinal.html" % (today.strftime("%Y_%m_%d"))

HTML_OUTPUT_FILE = open(htmlOutputFilename, "w")
driver = webdriver.Chrome()
main(sys.argv[1]) # ProvidentAvp_output.csv

time.sleep(60)  # Let the user actually see something!
driver.quit()




