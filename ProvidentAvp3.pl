#use strict;
#use warnings;
use Time::Piece;
use Time::HiRes qw(sleep);
use Test::WWW::Selenium;
use Test::More "no_plan";
use Test::Exception;
use Tk;

# This script is known to work with C:\Selenium>java -jar selenium-server-standalone-2.50.1.jar
# and Firefox version 47.0.2
# Firefox has Adds-On for Debugging Help: 
# - Selenium IDE
# - Selenium IDE Button
# - Selenium IDE: Perl Formatter
# Date: 2/3/2017

$AVP_WEBSITE = "https://login.assuredvehicleprotection.com/Login.aspx";

$SLOW_SPEED = 250;
$DUPLICATE_ENTRY = "DUPLICATE ENTRY";
$MANUAL_ENTRY = "MANUAL ENTRY REQUIRED";
$PROVIDENT = "Provident";
$PROFIN = "ProFin";

if (open(AM_INPUT_FILE,$ARGV[0]) == 0) {
   print "Error opening input AutoManager report file: ",$ARGV[0],"\n";
   exit;  
}

my $myCompany = $ARGV[1];

if (length($myCompany) eq 0)
{
	print "Error input company missing: ",$ARGV[1],"\n";
	exit;
}

if ( $myCompany eq $PROVIDENT )
{
	$myLogin = "jmcqueen";
}
elsif ($myCompany eq $PROFIN)
{
	$myLogin = "jmcqueenJTR";	
}
else 
{
	print "Error - no company name match: ", $myCompany, "\n";
	exit;
}

print "Company: ", $myCompany,"\n";


my $myTodayFormat = localtime->strftime('%Y_%m_%d');
my $avpHtmlFilename = sprintf("%s\\ProvidentAvpFinal.html",$myTodayFormat);
my $filename = sprintf(">%s",$avpHtmlFilename);

if (open(HTML_OUTPUT_FILE,$filename) == 0) {
   print "Error opening html: %s",$filename,"\n";
   exit;  
}

print HTML_OUTPUT_FILE  "<html>\n";
print HTML_OUTPUT_FILE  "<head><title>Pro Fin AvpXpress</title></head>\n";
print HTML_OUTPUT_FILE  "<body>\n";
print HTML_OUTPUT_FILE "<div style=\"display:block;text-align:left\"><a href=\"https://login.assuredvehicleprotection.com/Login.aspx\" imageanchor=1><img align=\"left\" src=\"avp.jpg\" border=0></a><h1><I>Provident Financial AvpXpress</I></h1>";
print HTML_OUTPUT_FILE "<head><style>\n";
print HTML_OUTPUT_FILE "table  { width:80%;}\n";
print HTML_OUTPUT_FILE "th, td { padding: 10px;}\n";
print HTML_OUTPUT_FILE "table#table01 tr:nth-child(even) { background-color: #eee; }\n";
print HTML_OUTPUT_FILE "table#table01 tr:nth-child(odd)  { background-color: #fff; }\n";
print HTML_OUTPUT_FILE "table#table01 th { background-color: #084B8A; color: white; }\n";
print HTML_OUTPUT_FILE "</style></head>\n";
print HTML_OUTPUT_FILE  "<table border=5 id=\"table01\" >\n";
print HTML_OUTPUT_FILE  "<tr><th>Index</th><th>Product</th><th>Months<th>Mileage</th><th>Deductible</th><th>Sale Date</th><th>VIN</th><th>Odometer</th><th>Name</th><th>Result</th></tr>\n";

my $answer;
my $sel = Test::WWW::Selenium->new( host => "localhost", 
                                    port => 4444, 
                                    browser => "*firefox", 
                                    browser_url => "https://login.assuredvehicleprotection.com/Login.aspx");

$sel->open_ok($AVP_WEBSITE);
$sel->click_ok("id=errorTryAgain");
#$sel->window_maximize();

while ( $sel->is_element_present("id=txtUserName") eq 0 )
{
	print "Refreshing AVP login page.\n";
	$sel->open($AVP_WEBSITE);
	sleep(2);
}

# New cookie consent button added to site - 4/2/20
$sel->click_ok("id=btnCookieConsent", $myLogin);
sleep(1);


$sel->type_ok("id=txtUserName", $myLogin);
$sel->type_ok("id=txtPassword", $myLogin);
$sel->click_ok("id=btnLogin");
$sel->wait_for_page_to_load_ok("30000");
#$sel->click_ok("css=#item_contracts > li.miOpen > span.miOpen");
$sel->select_frame_ok("frmContent");
$sel->click_ok("id=newContract");
#$sel->click_ok("css=li > span");
$sel->wait_for_page_to_load_ok("30000");

my $months;
my $mileage;
my $deductible;
my $vin;
my $odometer;
my $lastname;
my $firstname;
my $address;
my $city;
my $state;
my $zip;
my $saledate;
my $phone;
my $price;
my $high_mileage;
my $loopIteration = 0;

while (<AM_INPUT_FILE>) 
{
	chomp;
	($product,$months,$mileage,$deductible,$vin,$odometer,$firstname,$lastname,$saledate,$address,$city,$state,$zip,$phone,$price,$high_mileage) = split(",");
	
	$loopIteration++;
	
	$firstname =~ s/^ *//;
	$firstname = sprintf("%s",ucfirst(lc($firstname)));
	$lastname  = sprintf("%s",ucfirst(lc($lastname)));
	$high_mileage = uc($high_mileage);
		
	my ($sd_month, $sd_day, $sd_year) = split /\//, $saledate;
	
	if (length($sd_year) eq 2 )
	{
		$saledate = sprintf("%s/%s/20%s/",$sd_month,$sd_day,$sd_year);
	}
	
	print "PRODUCT: ",$product, " MONTHS/MILEAGE:",$months, "/",$mileage," DEDUCTIBLE:",$deductible," HIGH MILEAGE:",$high_mileage," VIN:",$vin," MILEAGE:",$odometer," NAME:",$firstname," ",$lastname, " SALEDATE:",$saledate,"\n";
	
	
	## DEBUG, uncomment this 'next' call to see the while loop debug variables without entering in anything to webstite
	#sleep(10);
	#next;
	
	
	##
	## WARRANTY
	## 
	if ((uc($product) eq "WARRANTY"))
	{

		print  HTML_OUTPUT_FILE "<tr>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$loopIteration,"</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">","WARRANTY","</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$months,"</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$mileage,"</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$deductible,"</td>\n";		
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$saledate,"</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$vin,"</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$odometer,"</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",uc($firstname)," ",uc($lastname),"</td>\n";
		
		print "WARRANTY: Clicking on warranty control..\n";
		
		# The Provident Login still have both companies.. need to use (2-warranty), but (0) for ProFin login
		if ($myCompany eq $PROFIN )
		{
			print "PROFIN WARRANTY: MONTHS/MILEAGE:",$months,"/",$mileage, " DEDUCTIBLE:",$deductible," HIGH MILEAGE:",$high_mileage, "\n";
			
			# Pro Fin Lending Solutions 3/3
			if ($months eq "3")
			{
				print "WARRANTY 3/3\"n";
				EnterBasicWarranty();
			}
			
			# Pro Fin Lending Solutions AVP3
			elsif ($high_mileage eq "NO")
			{
				print "WARRANTY - AVP3\n";
				EnterAVP3Warranty();
			}
			
			# Pro Fin Lending Solutions SPL - HIGH MILEAGE
			elsif ($high_mileage eq "YES")
			{
				print "WARRANTY - HIGH MILEAGE\n";
				EnterHighMileageWarranty();
			}			
		}
		else
		{
			# Provident Financial 3/3
			print "PROVIDENT 3/3 WARRANTY\n";
			EnterBasicWarranty();
		}
		
		## ------------------
		EnterCommonWarrantyInfo();		
					
	}
	
	##
	## GAP
	##
	if ((uc($product) eq "GAP"))	
	{
		print  HTML_OUTPUT_FILE "<tr>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$loopIteration,"</td>\n";		
		print  HTML_OUTPUT_FILE "<td align=\"center\">","GAP","</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$months,"</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$mileage,"</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$deductible,"</td>\n";		
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$saledate,"</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$vin,"</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$odometer,"</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",uc($firstname)," ",uc($lastname),"</td>\n";
		
		print "GAP: Clicking on warranty control..\n";
		##$sel->click_ok("id=GridView1_ctl03_lnkDealerID");
		
		# Pro Fin Lending Solutions GAP
		$sel->click_ok("id=GridView1_lnkDealerID_1");
				
		$sel->wait_for_page_to_load_ok("30000");
		$sel->type_ok("id=ContentPlaceHolder1_txtVIN", $vin);
		$sel->type_ok("id=ContentPlaceHolder1_txtMileage", $odometer);
		$sel->select_ok("id=ContentPlaceHolder1_ddlNewUsed", "label=Used");
		$sel->click_ok("id=ContentPlaceHolder1_btnSubmit");
		$sel->wait_for_page_to_load_ok("30000");
		print "GAP:Submitting VIN:",$vin,"\n";
		
		$answer = CheckForSubmitErrors($product);		
		print "GAP:CheckForSubmitErrors: ANSWER -> [", $answer,"]\n";
		
		if ( $answer eq $DUPLICATE_ENTRY )
		{
			$name = sprintf("%s, %s",$lastname,$firstname);
			$popup_answer = PopupBoxAnswer($product,$saledate,$name);
			print "GAP:PopupBoxAnswer:ANSWER -> [", $popup_answer,"]\n";
			
			if ($popup_answer eq "Yes" )
			{
				$sel->click_ok("link=Continue With New App");
				$sel->wait_for_page_to_load_ok("30000");
				$answer = "OK";
			}
		}
		
		if ( $answer ne "OK" )
		{
			print  HTML_OUTPUT_FILE "<td align=\"center\"><font color=red>",$answer,"</font></td>\n";
			print  HTML_OUTPUT_FILE "</tr>\n";
			$sel->click_ok("link=Contracts");
			$sel->wait_for_page_to_load_ok("30000");
			#$sel->select_frame_ok("frmContent");
			$sel->click_ok("css=li > span");
			$sel->wait_for_page_to_load_ok("30000");
			next
		}
		
		print uc($product),"- Entering customer information.\n";		
		
		if ($months eq 60)
		{
			$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[5]");		
		}
		
		if ($months eq 48)
		{
			$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[4]");		
		}
		
		if ($months eq 36)
		{
			$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[3]");		
		}
		
		if ($months eq 24)
		{
			$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[2]");		
		}
		
		if ($months eq 12)
		{
			$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[1]");		
		}
		
		## Try to see if there is only one GAP offered
		#$sel->click_ok("name=OptContractRate");
		
		## However, if there is 2, select the second one
		#$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[1]");		
		
		
		$sel->click_ok("id=lnkNext");
		$sel->wait_for_page_to_load_ok("30000");
		EnterProductInformation();
		print uc($product),"-entry submitted!\n";
		
		print  HTML_OUTPUT_FILE "<td align=\"center\">","OK","</td>\n";
		print  HTML_OUTPUT_FILE "</tr>\n";
	}
	
	print "Setting up for next entry <START>\n";
	
	# For the sake of the next loop...
	$sel->click_ok("link=Contracts");
	$sel->wait_for_page_to_load_ok("30000");
	#$sel->select_frame_ok("frmContent");
	$sel->click_ok("css=li > span");
	$sel->wait_for_page_to_load_ok("30000");
	
	print "Setting up for next entry <END>\n";	
}

print "Entries complete, load up Pending Contracts...\n";
	
$sel->click_ok("link=Home");
$sel->wait_for_page_to_load_ok("5000");
$sel->click_ok("css=a > div > div.boxContent");
$sel->wait_for_page_to_load_ok("5000");
$sel->click_ok("link=+ Pending Contracts");
$sel->wait_for_page_to_load_ok("5000");
$sel->click_ok("id=btnGetDate");
$sel->wait_for_page_to_load_ok("5000");

print HTML_OUTPUT_FILE "<br><br><br>\n";
print HTML_OUTPUT_FILE "</table>\n";
print HTML_OUTPUT_FILE "</form>\n";
print HTML_OUTPUT_FILE "</body>\n";
print HTML_OUTPUT_FILE "</html>\n";
print HTML_OUTPUT_FILE "<br><br>\n";

close(AM_INPUT_FILE); 
close(HTML_OUTPUT_FILE);

my $command = sprintf("start chrome %s",$avpHtmlFilename);
my $status = system($command);
sleep(30000);

sub EnterBasicWarranty
{
	$sel->click_ok("id=GridView1_lnkDealerID_0");
	$sel->wait_for_page_to_load_ok("30000");
	$sel->set_speed($SLOW_SPEED);
	$sel->type_ok("id=ContentPlaceHolder1_txtVIN", $vin);
	$sel->type_ok("id=ContentPlaceHolder1_txtMileage", $odometer);
	$sel->select_ok("id=ContentPlaceHolder1_ddlNewUsed", "label=Used");		
	$sel->click_ok("id=ContentPlaceHolder1_btnSubmit");
	$sel->wait_for_page_to_load_ok("30000");	
	$sel->click_ok("name=OptContractRate");		
}

sub EnterAVP3Warranty
{
	my $custom_warranty = 0;
	
	$sel->click_ok("id=GridView1_lnkDealerID_2");
	$sel->wait_for_page_to_load_ok("30000");
	$sel->set_speed($SLOW_SPEED);
	$sel->type_ok("id=ContentPlaceHolder1_txtVIN", $vin);
	$sel->type_ok("id=ContentPlaceHolder1_txtMileage", $odometer);
	
	# For warranites with coverage mileage >= 50k , pick CUSTOM
	if ( int($mileage) >= 50000 )
	{
		$sel->select_ok("id=ContentPlaceHolder1_ddlNewUsed", "label=Program");
		$custom_warranty = 1;
	}
	# For warranties with coverage miles < 50k, pick USED
	else
	{
		$sel->select_ok("id=ContentPlaceHolder1_ddlNewUsed", "label=Used");
		$custom_warranty = 0;
	}
	
	$sel->click_ok("id=ContentPlaceHolder1_btnSubmit");
	$sel->wait_for_page_to_load_ok("30000");
	
	print "WARRANTY - DEDUCTIBLE = ",$deductible,"\n";
				
	if ( $deductible eq "0" )
	{
		$sel->select_ok("id=ddlDeductibleOption", "label=Zero Deductible (Add \$100 to Price)");		
	}
	elsif ( $deductible eq "50" )
	{
		$sel->select_ok("id=ddlDeductibleOption", "label=\$50 Deductible (Add \$50 to Price)");					
	}		
	elsif ( $deductible eq "100" )
	{
		$sel->select_ok("id=ddlDeductibleOption", "label=\$100 Disappearing Deductible (Add \$75 to Price)");		
	}
	elsif ( $deductible eq "200" )
	{
		$sel->select_ok("id=ddlDeductibleOption", "label=\$200 Deductible ( NO CHARGE )");
	}
	
	
	if ( $custom_warranty > 0 )
	{
		print "CUSTOM WARRANTY - MONTHS = ",$months,"\n";
			
		if ( $mileage eq "50000" && $months eq "36" )
		{
			$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[1]");						
		}
		elsif ( $mileage eq "50000" && $months eq "48" )
		{
			$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[2]");
		}
		elsif ( $mileage eq "50000" && $months eq "60" )
		{
			$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[3]");
		}
		
		elsif ( $mileage eq "60000" && $months eq "36" )
		{
			$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[13]");						
		}
		elsif ( $mileage eq "60000" && $months eq "48" )
		{
			$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[14]");
		}
		elsif ( $mileage eq "60000" && $months eq "60" )
		{
			$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[15]");
		}
	}
	else 
	{
		print "WARRANTY - MONTHS = ",$months,"\n";
					
		if ( int($odometer) >= 85000 )
		{
			if ( $months eq "12" )
			{
				$sel->click_ok("name=OptContractRate");						
			}
			elsif ( $months eq "18" )
			{
				$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[4]");
			}
			elsif ( $months eq "24" )
			{
				$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[5]");
			}
			elsif ( $months eq "36" )
			{
				$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[7]");				
			}
		}
		else
		{
			if ( $months eq "12" )
			{
				$sel->click_ok("name=OptContractRate");						
			}
			elsif ( $months eq "18" )
			{
				$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[4]");
			}
			elsif ( $months eq "24" )
			{
				$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[7]");				
			}
			elsif ( $months eq "36" )
			{
				$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[10]");
			}
		}
	
	}	
}

sub EnterHighMileageWarranty
{
	$sel->click_ok("id=GridView1_lnkDealerID_3");
	$sel->wait_for_page_to_load_ok("30000");
	$sel->set_speed($SLOW_SPEED);
	$sel->type_ok("id=ContentPlaceHolder1_txtVIN", $vin);
	$sel->type_ok("id=ContentPlaceHolder1_txtMileage", $odometer);
	$sel->select_ok("id=ContentPlaceHolder1_ddlNewUsed", "label=Used");		
	$sel->click_ok("id=ContentPlaceHolder1_btnSubmit");
	$sel->wait_for_page_to_load_ok("30000");
	
	print "HIGH MILEAGE WARRANTY - MONTHS = ",$months,"\n";		
	
	if ( $months eq "12" )
	{
		$sel->click_ok("name=OptContractRate");						
	}
	elsif ( $months eq "24" )
	{
		$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[3]");
	}
	elsif ( $months eq "36" )
	{
		$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[5]");
	}
}

sub EnterCommonWarrantyInfo
{
	print "WARRANTY:Submitting VIN:",$vin,"\n";
	$answer = CheckForSubmitErrors($product);
	print "WARRANTY:CheckForSubmitErrors:ANSWER -> [", $answer,"]\n";
	
	if ( $answer eq $DUPLICATE_ENTRY )
	{
		$name = sprintf("%s, %s",$lastname,$firstname);
		$popup_answer = PopupBoxAnswer($product,$saledate,$name);
		print "WARRANTY:PopupBoxAnswer:ANSWER -> [", $popup_answer,"]\n";
		
		if ($popup_answer eq "Yes" )
		{
			$sel->click_ok("link=Continue With New App");
			$sel->wait_for_page_to_load_ok("30000");
			$answer = "OK";
		}
	}
			
	if ( $answer ne "OK" )
	{
		print  HTML_OUTPUT_FILE "<td align=\"center\"><font color=red>",$answer,"</font></td>\n";
		print  HTML_OUTPUT_FILE "</tr>\n";
		$sel->click_ok("link=Contracts");
		$sel->wait_for_page_to_load_ok("30000");
		$sel->select_frame_ok("frmContent");
		$sel->click_ok("css=li > span");
		$sel->wait_for_page_to_load_ok("30000");
		next;
	}
	
	print uc($product),":Entering customer information.\n";
	$sel->click_ok("id=lnkNext");
	$sel->wait_for_page_to_load_ok("30000");
	# DEBUG - sleep(10000);
	EnterProductInformation();
	print uc($product),":entry submitted!\n";
	
	print  HTML_OUTPUT_FILE "<td align=\"center\">","OK","</td>\n";
	print  HTML_OUTPUT_FILE "</tr>\n";
	
}

sub EnterProductInformation
{
	$sel->type_ok("id=txtFirstName", $firstname);
	$sel->type_ok("id=txtLastName", $lastname);
	
	$sel->set_speed(0);
	$sel->type_keys("id=txtZip", substr($zip, 0, 1));
	$sel->type_keys("id=txtZip", substr($zip, 1, 1));
	$sel->type_keys("id=txtZip", substr($zip, 2, 1));
	$sel->type_keys("id=txtZip", substr($zip, 3, 1));
	$sel->type_keys("id=txtZip", substr($zip, 4, 1));
	$sel->double_click("id=txtPurchaseDate");
	$sel->type_keys("id=txtPurchaseDate", "\e");
	
	my @values = split('/', $saledate);
	my $newsaledate;
	my $firsttime = 1;
	foreach my $val (@values) 
	{
	   $val = sprintf("%02d",$val);
	   
	   if ($firsttime eq 1 )
	   {
		   $firsttime = 0;       
		   $newsaledate = $newsaledate.$val;
	   }
	   else
	   {
		   $newsaledate = $newsaledate."/".$val;
	   }
	}
	
	$saledate = $newsaledate;
	
	for (my $i=0; $i <= length($saledate); $i++) 
	{
		$sel->type_keys("id=txtPurchaseDate", substr($saledate, $i, 1));
	}
	$sel->set_speed($SLOW_SPEED);
	
	$sel->type_ok("id=txtClientAddress", $address);
	$sel->type_ok("id=txtCity", $city);	
	
	my $myState = sprintf("label=%s",LookUpStateFromAbbreviation($state));
	$sel->select_ok("id=ddlState", $myState);

	$sel->set_speed(0);	
	$sel->double_click("id=txtPhone");
	
	for (my $i=0; $i <= length($phone); $i++) 
	{
		$sel->type_keys("id=txtPhone", substr($phone, $i, 1));
	}	
	$sel->set_speed($SLOW_SPEED);	
	
	$sel->type_ok("id=txtVehicleCost2", $price);
	# DEBUG: Comment out this line to not submit the warranty/gap 
	$sel->click_ok("id=btnLookupSender");
	$sel->wait_for_page_to_load_ok("30000");
}

sub PopupBoxAnswer
{
	my $mw = MainWindow->new();
	$mw->withdraw();
	
	my $product   = $_[0];
	my $saledate  = $_[1];
	my $name      = $_[2];
	
	my $message = sprintf("Client Name: %s  Purchase Date: %s, replace?\n",$name,$saledate);
	my $title   = sprintf("%s: Duplicate Entry Detected",uc($product));
	
	my $answer = $mw->messageBox(
	  -title   => $title,
	  -message => $message,
	  -type    => 'YesNo',
	  -icon    => 'question',
	);

	return $answer;	
}


sub CheckForSubmitErrors
{
	my $answer = 0;
	my $product = $_[0];
	
	$answer = $sel->is_text_present("The contracts listed above have the same Contract ID as the one trying to be written.");
	
	if ($answer eq 1 )
	{
		print uc($product),": Found duplicate entry.\n";
		return $DUPLICATE_ENTRY;
	}

	$answer = $sel->is_text_present("There are no rates available for this vehicle.");
	
	if ($answer eq 1 )
	{
		print uc($product),": Manual entry required.\n";
		return $MANUAL_ENTRY;
	}	
	
	return "OK";
}

sub LookUpStateFromAbbreviation
{
	my $abbev = $_[0];
	
	if ($abbev eq "AL" )
	{
		return "Alabama";
	}
	elsif ($abbev eq "AK")
	{
		return "Alaska";
	}
	elsif ($abbev eq "AZ")
	{
		return "Arizona";
	}
	elsif ($abbev eq "AR")
	{		
		return "Arkansas";
	}
	elsif ($abbev eq "CA")
	{
		return "California";
	}
	elsif ($abbev eq "CO")
	{
		return "Colorado";
	}
	elsif ($abbev eq "DE")
	{
		return "Delaware";
	}
	elsif ($abbev eq "FL")
	{
		return "Florida";
	}
	elsif ($abbev eq "GA")
	{
		return "Georgia";
	}
	elsif ($abbev eq "HI")
	{
		return "Hawaii";
	}
	elsif ($abbev eq "ID")
	{
		return "Idaho";
	}
	elsif ($abbev eq "IL")
	{
		return "Illinois";
	}
	elsif ($abbev eq "IN")
	{
		return "Indiana";
	}
	elsif ($abbev eq "IA")
	{
		return "Iowa";
	}
	elsif ($abbev eq "KS")
	{
		return "Kansas";
	}
	elsif ($abbev eq "KY")
	{
		return "Kentucky";
	}
	elsif ($abbev eq "LA")
	{
		return "Louisiana";
	}
	elsif ($abbev eq "ME")
	{
		return "Maine";
	}
	elsif ($abbev eq "MD")
	{
		return "Maryland";
	}
	elsif ($abbev eq "MA")
	{
		return "Massachusetts";
	}
	elsif ($abbev eq "MI")
	{
		return "Michigan";
	}
	elsif ($abbev eq "MN")
	{
		return "Minnesota";
	}
	elsif ($abbev eq "MS")
	{
		return "Mississippi";
	}
	elsif ($abbev eq "MO")
	{
		return "Missouri";
	}
	elsif ($abbev eq "MT")
	{
		return "Montana";
	}
	elsif ($abbev eq "NE")
	{
		return "Nebraska";
	}
	elsif ($abbev eq "NV")
	{
		return "Nevada";
	}
	elsif ($abbev eq "NH")
	{
		return "New Hampshire";
	}
	elsif ($abbev eq "NJ")
	{
		return "New Jersey";
	}
	elsif ($abbev eq "NM")
	{
		return "New Mexico";
	}
	elsif ($abbev eq "NY")
	{
		return "New York";
	}
	elsif ($abbev eq "NC")
	{
		return "North Carolina";
	}
	elsif ($abbev eq "ND")
	{
		return "North Dakota";
	}
	elsif ($abbev eq "OH")
	{
		return "Ohio";
	}
	elsif ($abbev eq "OK")
	{
		return "Oklahoma";
	}
	elsif ($abbev eq "OR")
	{
		return "Oregon";
	}
	elsif ($abbev eq "PA")
	{
		return "Pennsylvania";
	}
	elsif ($abbev eq "RI")
	{
		return "Rhode Island";
	}
	elsif ($abbev eq "SC")
	{
		return "South Carolina";
	}
	elsif ($abbev eq "SD")
	{
		return "South Dakota";
	}
	elsif ($abbev eq "TN")
	{
		return "Tennessee";
	}
	elsif ($abbev eq "TX")
	{
		return "Texas";
	}
	elsif ($abbev eq "UT")
	{
		return "Utah";
	}
	elsif ($abbev eq "VT")
	{
		return "Vermont";
	}
	elsif ($abbev eq "VA")
	{
		return "Virginia";
	}
	elsif ($abbev eq "WA")
	{
		return "Washington";
	}
	elsif ($abbev eq "WV")
	{
		return "West Virginia";
	}
	elsif ($abbev eq "WI")
	{
		return "Wisconsin";
	}
	elsif ($abbev eq "WY")
	{
		return "Wyoming";
	}
	else
	{
		return "Error";
	}
	
}

