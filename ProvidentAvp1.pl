#!/usr/bin/perl -w

use warnings;
use strict;
use 5.010;

use Time::Piece;
use IO::Handle;
use Locale::Currency::Format;

 
# Works with Firefox version 43
# https://ftp.mozilla.org/pub/firefox/releases/43.0/win32/en-US/
  
# ProvidentAvp1.pl 
# Input file: ARGV[0] = Sales Report in .CSV format
# Output file: ProvidentAvpFile.html - the selection of products for each customer
#

if (open(AM_INPUT_FILE,$ARGV[0]) == 0) {
   print "Error opening input AutoManager report file: ",$ARGV[0],"\n";
   exit -1;  
}

my $vin;
my $mileage;
my $lastname;
my $firstname;
my $address;
my $city;
my $zip;
my $saledate;
my $homephone;
my $cellphone;
my $vehicle;
my $state;
my $price;


while (<AM_INPUT_FILE>) 
{
	chomp;
	($vin,$mileage,$firstname,$lastname,$address,$city,$state,$zip,$saledate,$homephone,$cellphone,$price,$vehicle) = split(",");
	
	if (length($vin) ne 17)
	{
		last;
	}
	
	if (!length($mileage) || !length($firstname) || !length($lastname) || !length($address) || !length($city) || !length($state) || !length($zip) || !length($saledate) || !length($price) || !length($vehicle))
	{
		print "Error found in entry: VIN: ",$vin, " Firstname: ",$firstname, " Lastname: ",$lastname,"\n";
		exit;
	}
}

close (AM_INPUT_FILE);



if (open(AM_INPUT_FILE,$ARGV[0]) == 0) {
   print "Error opening input AutoManager report file: ",$ARGV[0],"\n";
   exit -1;  
}

my $myTodayFormat = localtime->strftime('%Y_%m_%d');
mkdir $myTodayFormat;

my $avpHtmlFilename = sprintf("%s\\ProvidentAvpFile.html",$myTodayFormat);
my $filename = sprintf(">%s",$avpHtmlFilename);
if (open(HTML_OUTPUT_FILE,$filename) == 0) {
   print "Error opening: %s",$filename,"\n";
   exit -1;  
}


print HTML_OUTPUT_FILE  "<html>\n";
print HTML_OUTPUT_FILE  "<body>\n";
print HTML_OUTPUT_FILE  "<title>AVP eXpress</title>\n";
print HTML_OUTPUT_FILE  "<div style=\"display:block;text-align:left\">\n";
print HTML_OUTPUT_FILE  "<a href=\"http://new.assuredvehicleprotection.com/Login.aspx\" imageanchor=1>\n";
print HTML_OUTPUT_FILE  "<img align=\"left\" src=\"ProvidentAvp.png\" border=0></a><h1><I>AVP eXpress</I></h1><br>\n";
print HTML_OUTPUT_FILE  "<form method=\"GET\" action=\"http://localhost/ProvidentAvpExpress/ProvidentAvp2.pl\">\n";

print HTML_OUTPUT_FILE  "<head><style>\n";
print HTML_OUTPUT_FILE  "table { width:100%;}\n";
print HTML_OUTPUT_FILE  "th, td { padding: 10px;}\n";
print HTML_OUTPUT_FILE  "table#table01 tr:nth-child(even) { background-color: #eee; }\n";
print HTML_OUTPUT_FILE  "table#table01 tr:nth-child(odd)  { ba084B8Ackground-color: #fff; }\n";
print HTML_OUTPUT_FILE  "table#table01 th { background-color: #084B8A; color: white; }\n";
print HTML_OUTPUT_FILE  "</style></head>\n";

print HTML_OUTPUT_FILE "<style TYPE=\"text/css\">";
print HTML_OUTPUT_FILE "<!--\n";
print HTML_OUTPUT_FILE "TD{font-family: Arial; font-size: 10pt;}\n";
print HTML_OUTPUT_FILE "TH{font-family: Arial; font-size: 10pt;}\n";
print HTML_OUTPUT_FILE "--->\n";
print HTML_OUTPUT_FILE "</style>\n";

print HTML_OUTPUT_FILE  "<td width=\"3%\" align=\"center\" bgcolor=\"#F3F781\"><input type=\"radio\" name=\"cust_product\" value=\"Warranty\" checked>Warranty</td>   "; 
print HTML_OUTPUT_FILE  "<td width=\"4%\" align=\"center\" bgcolor=\"#F3F781\"><input type=\"radio\" name=\"cust_product\" value=\"GAP\">GAP</td><br>\n"; 
print HTML_OUTPUT_FILE  "<br><br>";	

print HTML_OUTPUT_FILE  "<table border=5 id=\"table01\" >\n";
print HTML_OUTPUT_FILE  "<tr><th>Entry</th><th>Select</th><th>Warranty (months)</th><th>Warranty (mileage)</th><th>Warranty (deductible)</th><th>High Mileage</th><th>Sales Date</th><th>VIN</th><th>Odometer</th><th>Name</th><th>Sales Price</th><th>Vehicle</th></tr>\n";

my $loopIteration = 0;

while (<AM_INPUT_FILE>) 
{
	chomp;
	($vin,$mileage,$firstname,$lastname,$address,$city,$state,$zip,$saledate,$homephone,$cellphone,$price,$vehicle) = split(",");
	
	print $loopIteration, " ", $vin, "\n";	
	
	if (length($vin) ne 17)
	{
		last;
	}
	
	$firstname =~ s/^ *//;
	$firstname = sprintf("%s",ucfirst(lc($firstname)));
	$lastname  = sprintf("%s",ucfirst(lc($lastname)));
	$vehicle   = uc($vehicle);
	
	my $phone = $cellphone;
	
	if (length $phone < 10)
	{
		$phone = $homephone;		
	}
	
	if (!length($mileage) || !length($firstname) || !length($lastname) || !length($address) || !length($city) || !length($state) || !length($zip) || !length($saledate) || !length($price) || !length($vehicle))
	{
		print "Error found in entry: VIN: ",$vin, " Firstname: ",$firstname, " Lastname: ",$lastname,"\n";
		exit;
	}
	
	
	my $formatted_price   = currency_format('usd',$price,FMT_SYMBOL);	
	my $formatted_mileage = $mileage; 
	$formatted_mileage =~ s/(?<=\d)(?=(?:\d\d\d)+\b)/,/g;	
	
	print HTML_OUTPUT_FILE  "<tr>\n";	
	
	print HTML_OUTPUT_FILE  "<td width=\"1%\" align=\"center\">",$loopIteration+1,"</td>\n";
		
	print HTML_OUTPUT_FILE  "<td width=\"1%\" align=\"center\" bgcolor=\"#F3F781\"><input type=\"checkbox\" name=\"cust_",$loopIteration,"_cb\" value=\"yes\" checked></td>\n"; 
	
	print HTML_OUTPUT_FILE  "<td width=\"7%\" align=\"center\" bgcolor=\"#04B431\"><label for=\"months\"></label><select id=\"months\" name=\"cust_",$loopIteration,"_product_months\"><option value=\"3\">3</option><option value=\"12\">12</option><option value=\"18\">18</option><option value=\"24\">24</option><option value=\"36\">36</option><option value=\"48\">48</option><option value=\"60\">60</option></select></td>\n";
	
	print HTML_OUTPUT_FILE  "<td width=\"7%\" align=\"center\" bgcolor=\"#04B431\"><label for=\"mileage\"></label><select id=\"mileage\" name=\"cust_",$loopIteration,"_product_mileage\"><option value=\"12000\">12,000</option><option value=\"18000\">18,000</option><option value=\"24000\">24,000</option><option value=\"36000\">36,000</option><option value=\"48000\">48,000</option><option value=\"50000\">50,000</option><option value=\"60000\">60,000</option></select></td>\n";
		
	print HTML_OUTPUT_FILE  "<td width=\"7%\" align=\"center\" bgcolor=\"#04B431\"><label for=\"deductible\"></label><select id=\"deductible\" name=\"cust_",$loopIteration,"_product_deductible\"><option value=\"0\">\$0</option><option value=\"50\">\$50</option><option value=\"100\">\$100</option><option value=\"200\">\$200</option></select></td>\n";
	
	print HTML_OUTPUT_FILE  "<td width=\"7%\" align=\"center\" bgcolor=\"#04B431\"><label for=\"high_mileage\"></label><select id=\"high_mileage\" name=\"cust_",$loopIteration,"_high_mileage\"><option value=\"No\">No</option><option value=\"Yes\">Yes</option></select></td>\n";
	
	#print HTML_OUTPUT_FILE  "<td width=\"4%\" align=\"center\" bgcolor=\"#C89600\"><input type=\"checkbox\" name=\"cust_",$loopIteration,"_high_mileage\" value=\"yes\"></td>\n";
		
	print HTML_OUTPUT_FILE  "<td width=\"7%\" align=\"center\">",$saledate,"</td>\n";
	print HTML_OUTPUT_FILE  "<td align=\"center\">",$vin,"</td>\n";
	print HTML_OUTPUT_FILE  "<td align=\"left\">",$formatted_mileage,"</td>\n";	
	print HTML_OUTPUT_FILE  "<td align=\"left\">",uc($firstname)," ",uc($lastname),"</td>\n";
	print HTML_OUTPUT_FILE  "<td align=\"center\">",$formatted_price,"</td>\n";
	print HTML_OUTPUT_FILE  "<td align=\"left\">",$vehicle,"</td>\n";
	print HTML_OUTPUT_FILE  "</tr>\n";
		
	print HTML_OUTPUT_FILE  "<input type=\"hidden\" name=\"cust_",$loopIteration,"_vin\" value=\"",$vin,"\">\n";
	print HTML_OUTPUT_FILE  "<input type=\"hidden\" name=\"cust_",$loopIteration,"_mileage\" value=\"",$mileage,"\">\n";
	print HTML_OUTPUT_FILE  "<input type=\"hidden\" name=\"cust_",$loopIteration,"_firstname\" value=\"",$firstname,"\">\n";
	print HTML_OUTPUT_FILE  "<input type=\"hidden\" name=\"cust_",$loopIteration,"_lastname\" value=\"",$lastname,"\">\n";
	print HTML_OUTPUT_FILE  "<input type=\"hidden\" name=\"cust_",$loopIteration,"_address\" value=\"",$address,"\">\n";
	print HTML_OUTPUT_FILE  "<input type=\"hidden\" name=\"cust_",$loopIteration,"_city\" value=\"",$city,"\">\n";
	print HTML_OUTPUT_FILE  "<input type=\"hidden\" name=\"cust_",$loopIteration,"_state\" value=\"",$state,"\">\n";	
	print HTML_OUTPUT_FILE  "<input type=\"hidden\" name=\"cust_",$loopIteration,"_zip\" value=\"",$zip,"\">\n";
	print HTML_OUTPUT_FILE  "<input type=\"hidden\" name=\"cust_",$loopIteration,"_saledate\" value=\"",$saledate,"\">\n";
	print HTML_OUTPUT_FILE  "<input type=\"hidden\" name=\"cust_",$loopIteration,"_phone\" value=\"",$phone,"\">\n";
	print HTML_OUTPUT_FILE  "<input type=\"hidden\" name=\"cust_",$loopIteration,"_price\" value=\"",$price,"\">\n";
	print HTML_OUTPUT_FILE  "<input type=\"hidden\" name=\"cust_",$loopIteration,"_vehicle\" value=\"",$vehicle,"\">\n";
			
	$loopIteration++;
}

print HTML_OUTPUT_FILE "</table>\n";
print HTML_OUTPUT_FILE "<b>Number of Sales: <input type=\"text\" name=\"num_accounts\" value=\"",$loopIteration,"\" size=4 style=\"border:none\" readonly></b><br>\n";   

print HTML_OUTPUT_FILE "<br><input type=\"submit\" value=\"Submit Accounts\" style=\"height:30px; width:150px\"><br><br><br>\n";
print HTML_OUTPUT_FILE "</form>\n";
print HTML_OUTPUT_FILE "</body>\n";
print HTML_OUTPUT_FILE "</html>\n";

close(AM_INPUT_FILE); 
close(HTML_OUTPUT_FILE);

