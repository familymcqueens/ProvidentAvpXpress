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

my $myTodayFormat = localtime->strftime('%Y_%m_%d');
mkdir $myTodayFormat;

my $avpHtmlFilename = sprintf("%s\\ProvidentAvpFile.html",$myTodayFormat);
my $filename = sprintf(">%s",$avpHtmlFilename);
if (open(HTML_OUTPUT_FILE,$filename) == 0) {
   print "Error opening: %s",$filename,"\n";
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

print HTML_OUTPUT_FILE  "<td width=\"4%\" align=\"center\" bgcolor=\"#F3F781\"><input type=\"radio\" name=\"cust_product\" value=\"Warranty\" checked>Warranty</td>   "; 
print HTML_OUTPUT_FILE  "<td width=\"4%\" align=\"center\" bgcolor=\"#F3F781\"><input type=\"radio\" name=\"cust_product\" value=\"GAP\">GAP</td><br>\n"; 
print HTML_OUTPUT_FILE  "<br><br>";	


print HTML_OUTPUT_FILE  "<table border=5 id=\"table01\" >\n";
print HTML_OUTPUT_FILE  "<tr><th>Select</th><th>12 mo</th><th>24 mo</th><th>36 mo</th><th>48 mo</th><th>Sales Date</th><th>VIN</th><th>Mileage</th><th>Name</th><th>Sales Price</th><th>Vehicle</th></tr>\n";

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
	
	my $formatted_price   = currency_format('usd',$price,FMT_SYMBOL);	
	my $formatted_mileage = $mileage; 
	$formatted_mileage =~ s/(?<=\d)(?=(?:\d\d\d)+\b)/,/g;	
	
	print HTML_OUTPUT_FILE  "<tr>\n";	
	
	print HTML_OUTPUT_FILE  "<td width=\"4%\" align=\"center\" bgcolor=\"#F3F781\"><input type=\"checkbox\" name=\"cust_",$loopIteration,"_cb\" value=\"yes\" checked></td>\n"; 
	print HTML_OUTPUT_FILE  "<td width=\"4%\" align=\"center\" bgcolor=\"#04B431\"><input type=\"radio\" name=\"cust_",$loopIteration,"_gap_product\" value=\"12\"checked></td>\n";
	print HTML_OUTPUT_FILE  "<td width=\"4%\" align=\"center\" bgcolor=\"#04B431\"><input type=\"radio\" name=\"cust_",$loopIteration,"_gap_product\" value=\"24\"></td>\n";
	print HTML_OUTPUT_FILE  "<td width=\"4%\" align=\"center\" bgcolor=\"#04B431\"><input type=\"radio\" name=\"cust_",$loopIteration,"_gap_product\" value=\"36\"></td>\n";
	print HTML_OUTPUT_FILE  "<td width=\"4%\" align=\"center\" bgcolor=\"#04B431\"><input type=\"radio\" name=\"cust_",$loopIteration,"_gap_product\" value=\"48\"></td>\n";
	
	print HTML_OUTPUT_FILE  "<td align=\"center\">",$saledate,"</td>\n";
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

