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

$AVP_WEBSITE = "http://new.assuredvehicleprotection.com/";

$SLOW_SPEED = 250;
$DUPLICATE_ENTRY = "DUPLICATE ENTRY";
$MANUAL_ENTRY = "MANUAL ENTRY REQUIRED";

if (open(AM_INPUT_FILE,$ARGV[0]) == 0) {
   print "Error opening input AutoManager report file: ",$ARGV[0],"\n";
   exit;  
}

my $myTodayFormat = localtime->strftime('%Y_%m_%d');
my $avpHtmlFilename = sprintf("%s\\ProvidentAvpFinal.html",$myTodayFormat);
my $filename = sprintf(">%s",$avpHtmlFilename);

if (open(HTML_OUTPUT_FILE,$filename) == 0) {
   print "Error opening: %s",$filename,"\n";
   exit;  
}

print HTML_OUTPUT_FILE  "<html>\n";
print HTML_OUTPUT_FILE  "<head><title>Provident Financial AVP eXpress</title></head>\n";
print HTML_OUTPUT_FILE  "<body>\n";
print HTML_OUTPUT_FILE "<div style=\"display:block;text-align:left\"><a href=\"http://new.assuredvehicleprotection.com\" imageanchor=1><img align=\"left\" src=\"avp.jpg\" border=0></a><h1><I>Provident Financial AVP eXpress</I></h1>";
print HTML_OUTPUT_FILE "<head><style>\n";
print HTML_OUTPUT_FILE "table  { width:80%;}\n";
print HTML_OUTPUT_FILE "th, td { padding: 10px;}\n";
print HTML_OUTPUT_FILE "table#table01 tr:nth-child(even) { background-color: #eee; }\n";
print HTML_OUTPUT_FILE "table#table01 tr:nth-child(odd)  { background-color: #fff; }\n";
print HTML_OUTPUT_FILE "table#table01 th { background-color: #084B8A; color: white; }\n";
print HTML_OUTPUT_FILE "</style></head>\n";
print HTML_OUTPUT_FILE  "<table border=5 id=\"table01\" >\n";
print HTML_OUTPUT_FILE  "<tr><th>Index</th><th>Product</th><th>Sale Date</th><th>VIN</th><th>Mileage</th><th>Name</th><th>Result</th></tr>\n";

my $answer;
my $sel = Test::WWW::Selenium->new( host => "localhost", 
                                    port => 4444, 
                                    browser => "*firefox", 
                                    browser_url => "http://new.assuredvehicleprotection.com");

$sel->open_ok($AVP_WEBSITE);
$sel->click_ok("id=errorTryAgain");
$sel->window_maximize();

while ( $sel->is_element_present("id=txtUserName") eq 0 )
{
	print "Refreshing AVP login page.\n";
	$sel->open($AVP_WEBSITE);
	sleep(2);
}

$sel->type_ok("id=txtUserName", "tbaer");
$sel->type_ok("id=txtPassword", "tbaer");
$sel->click_ok("id=btnLogin");
$sel->wait_for_page_to_load_ok("30000");
#$sel->click_ok("css=#item_contracts > li.miOpen > span.miOpen");
$sel->select_frame_ok("frmContent");
$sel->click_ok("id=newContract");
#$sel->click_ok("css=li > span");
$sel->wait_for_page_to_load_ok("30000");

my $vin;
my $mileage;
my $lastname;
my $firstname;
my $address;
my $city;
my $state;
my $zip;
my $saledate;
my $phone;
my $loopIteration = 0;

while (<AM_INPUT_FILE>) 
{
	chomp;
	($product,$vin,$mileage,$firstname,$lastname,$saledate,$address,$city,$state,$zip,$phone,$price) = split(",");
	
	$loopIteration++;
	
	$firstname =~ s/^ *//;
	$firstname = sprintf("%s",ucfirst(lc($firstname)));
	$lastname  = sprintf("%s",ucfirst(lc($lastname)));
		
	my ($sd_month, $sd_day, $sd_year) = split /\//, $saledate;
	##print "Month: ",$sd_month," Day: ",$sd_day, " Year: ",$sd_year, "\n";

	if (length($sd_year) eq 2 )
	{
		$saledate = sprintf("%s/%s/20%s/",$sd_month,$sd_day,$sd_year);
	}
	
	print "PRODUCT: ",$product, " VIN: ",$vin," MILEAGE: ",$mileage," LASTNAME:",$lastname, " SALEDATE: ",$saledate,"\n";
	
		
	##
	## WARRANTY
	##
	if (($product eq "warranty"))
	{
		print  HTML_OUTPUT_FILE "<tr>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$loopIteration,"</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">","WARRANTY","</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$saledate,"</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$vin,"</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$mileage,"</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",uc($firstname)," ",uc($lastname),"</td>\n";
		
		print "WARRANTY: Clicking on warranty control..\n";
		##$sel->click_ok("id=GridView1_ctl02_lnkDealerID");
		$sel->click_ok("id=GridView1_lnkDealerID_0");
		$sel->wait_for_page_to_load_ok("30000");

		$sel->set_speed($SLOW_SPEED);
		$sel->type_ok("id=ContentPlaceHolder1_txtVIN", $vin);
		$sel->type_ok("id=ContentPlaceHolder1_txtMileage", $mileage);
		$sel->select_ok("id=ContentPlaceHolder1_ddlNewUsed", "label=Used");
		$sel->click_ok("id=ContentPlaceHolder1_btnSubmit");
		
		##$sel->type_ok("id=ctl00_ContentPlaceHolder1_txtVIN", $vin);
		##$sel->type_ok("id=ctl00_ContentPlaceHolder1_txtMileage", $mileage);
		##$sel->select_ok("id=ctl00_ContentPlaceHolder1_ddlNewUsed", "label=Used");
		##$sel->click_ok("id=ctl00_ContentPlaceHolder1_btnSubmit");
		$sel->wait_for_page_to_load_ok("30000");
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
			#$sel->select_frame_ok("frmContent");
			$sel->click_ok("css=li > span");
			$sel->wait_for_page_to_load_ok("30000");
			next;
		}
		
		print uc($product),":Entering customer information.\n";
		$sel->click_ok("name=OptContractRate");
		$sel->click_ok("id=lnkNext");
		$sel->wait_for_page_to_load_ok("30000");
		EnterProductInformation();
		print uc($product),":entry submitted!\n";
		
		print  HTML_OUTPUT_FILE "<td align=\"center\">","OK","</td>\n";
		print  HTML_OUTPUT_FILE "</tr>\n";			
	}
	
	##
	## GAP
	##
	if (($product eq "gap"))	
	{
		print  HTML_OUTPUT_FILE "<tr>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$loopIteration,"</td>\n";		
		print  HTML_OUTPUT_FILE "<td align=\"center\">","GAP","</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$saledate,"</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$vin,"</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",$mileage,"</td>\n";
		print  HTML_OUTPUT_FILE "<td align=\"center\">",uc($firstname)," ",uc($lastname),"</td>\n";
		
		print "GAP: Clicking on warranty control..\n";
		##$sel->click_ok("id=GridView1_ctl03_lnkDealerID");
		$sel->click_ok("id=GridView1_lnkDealerID_1");
		$sel->wait_for_page_to_load_ok("30000");
		##$sel->type_ok("id=ctl00_ContentPlaceHolder1_txtVIN", $vin);
		##$sel->type_ok("id=ctl00_ContentPlaceHolder1_txtMileage", $mileage);
		##$sel->select_ok("id=ctl00_ContentPlaceHolder1_ddlNewUsed", "label=Used");
		##$sel->click_ok("id=ctl00_ContentPlaceHolder1_btnSubmit");
		$sel->type_ok("id=ContentPlaceHolder1_txtVIN", $vin);
		$sel->type_ok("id=ContentPlaceHolder1_txtMileage", $mileage);
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
		##$sel->click_ok("document.form1.OptContractRate[1]");
		$sel->click_ok("xpath=(//input[\@name='OptContractRate'])[2]");
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
	$sel->click_ok("id=btnLookupSender");
	#$sel->click_ok("id=btnSave");
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
		return "Virgina";
	}
	elsif ($abbev eq "WA")
	{
		return "Washington";
	}
	elsif ($abbev eq "WV")
	{
		return "West Virgina";
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
		return 0;
	}
	
}

