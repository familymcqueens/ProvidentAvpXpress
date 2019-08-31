#!/usr/bin/perl -w
use Time::Piece;
use Locale::Currency::Format;
 
use Data::Dumper;
use CGI;
$Data::Dumper::Indent   = 1;
$Data::Dumper::Sortkeys = 1;
my $q = CGI->new;
my %data;
print $q->header;

# GAP or Warranty
$data{cust_product} = $q->param("cust_product");
my $product  = $data{cust_product};


my $myTodayFormat = localtime->strftime('%Y_%m_%d');

# ***************************************************************************
# *** CHANGE THIS IF LOCATION OF PROVIDENT AVP EXPRESS DIRECTORY CHANGES  ***
# ***************************************************************************
my $csv_output_filename = sprintf("c:\\jmcqueen\\Provident\\Wwebserver\\ProvidentAvpExpress\\%s\\ProvidentAvp_output.csv",$myTodayFormat);

print   "<html>\n";
print   "<title>AVP eXpress</title>\n";
print   "<body>\n";
print   "<div style=\"display:block;text-align:left\"><a href=\"http://new.assuredvehicleprotection.com\" imageanchor=1><img align=\"left\" src=\"ProvidentAvp.png\" border=0></a><h1><I>Provident Financial AVP eXpress</I></h1>";
print   "<head><style>\n";
print   "table  { width:80%;}\n";
print   "th, td { padding: 10px;}\n";
print   "table#table01 tr:nth-child(even) { background-color: #eee; }\n";
print   "table#table01 tr:nth-child(odd)  { background-color: #fff; }\n";
print   "table#table01 th { background-color: #084B8A; color: white; }\n";
print   "</style></head>\n";

print  "<style TYPE=\"text/css\">";
print  "<!--\n";
print  "TD{font-family: Arial; font-size: 10pt;}\n";
print  "TH{font-family: Arial; font-size: 10pt;}\n";
print  "--->\n";
print  "</style>\n";

print   "<table border=5 id=\"table01\" >\n";
print   "<tr><th>Index</th><th>Product</th><th>VIN</th><th>Mileage</th><th>Customer Name</th><th>Sales Date</th><th>Price</th><th>Vehicle</th></tr>\n";

print   "<form method=\"GET\" action=\"http://localhost/ProvidentAvpExpress/ProvidentAvp3.pl\">\n";

my $filename = sprintf(">%s",$csv_output_filename);
if (open(CSV_OUTPUT_FILE,$filename) == 0) {
   print "Error opening: %s",$filename,"\n";
   exit -1;  
}

$data{num_accounts}  = $q->param('num_accounts');

my $num_accounts_checked=0;

for (my $i=0; $i < $data{num_accounts}; $i++) 
{
	my $cust_checked    = sprintf("cust_%d_cb", $i);	
	$data{cust_cb} = $q->param($cust_checked);
	
	if ( $data{cust_cb} )
	{
		$num_accounts_checked++;
		
		my $cust_vin = sprintf("cust_%d_vin", $i);
	    $data{cust_vin} = $q->param($cust_vin);
		
		my $cust_mileage = sprintf("cust_%d_mileage", $i);
	    $data{cust_mileage} = $q->param($cust_mileage);
		
		my $cust_lastname = sprintf("cust_%d_lastname", $i);
	    $data{cust_lastname} = $q->param($cust_lastname);
		
		my $cust_firstname = sprintf("cust_%d_firstname", $i);
	    $data{cust_firstname} = $q->param($cust_firstname);
		
		my $cust_address = sprintf("cust_%d_address", $i);
	    $data{cust_address} = $q->param($cust_address);

		my $cust_city = sprintf("cust_%d_city", $i);
	    $data{cust_city} = $q->param($cust_city);
		
		my $cust_state = sprintf("cust_%d_state", $i);
	    $data{cust_state} = $q->param($cust_state);

		my $cust_zip = sprintf("cust_%d_zip", $i);	
	    $data{cust_zip} = $q->param($cust_zip);

		my $cust_saledate = sprintf("cust_%d_saledate", $i);	
	    $data{cust_saledate} = $q->param($cust_saledate);
		
		my $cust_phone = sprintf("cust_%d_phone", $i);	
	    $data{cust_phone} = $q->param($cust_phone);
		
		my $cust_price = sprintf("cust_%d_price", $i);	
	    $data{cust_price} = $q->param($cust_price);
		
		my $cust_vehicle = sprintf("cust_%d_vehicle", $i);	
	    $data{cust_vehicle} = $q->param($cust_vehicle);

		my $vin = $data{cust_vin};
		my $mileage = $data{cust_mileage};	
		my $firstname = sprintf("%s",ucfirst(lc($data{cust_firstname})));
		my $lastname  = sprintf("%s",ucfirst(lc($data{cust_lastname})));
		my $address = $data{cust_address};
		my $city = $data{cust_city};
		my $state = $data{cust_state};
		my $zip = $data{cust_zip};
		my $saledate = $data{cust_saledate};
		my $phone = $data{cust_phone};
		my $price = $data{cust_price};
		my $vehicle = $data{cust_vehicle};
		
		my $formatted_price   = currency_format('usd',$price,FMT_SYMBOL);	
		my $formatted_mileage = $mileage; 
		$formatted_mileage =~ s/(?<=\d)(?=(?:\d\d\d)+\b)/,/g;	

		print   "<td align=\"center\">",$i+1,"</td>\n";
		print   "<td width=\"4%\" align=\"center\" bgcolor=\"#F3F781\">",uc($product),"</td>\n";
		print   "<td align=\"center\">",$vin,"</td>\n";
		print   "<td align=\"center\">",$formatted_mileage,"</td>\n";
		print   "<td align=\"left\">",uc($firstname)," ",uc($lastname),"</td>\n";
		print   "<td align=\"center\">",$saledate,"</td>\n";
		print   "<td align=\"center\">",$formatted_price,"</td>\n";
		print   "<td align=\"left\">",$vehicle,"</td>\n";
		print   "</tr>\n";
		
		print CSV_OUTPUT_FILE $product,",",$vin,",",$mileage,",",$firstname,",",$lastname,",",$saledate,",",$address,",",$city,",",$state,",",$zip,",",$phone,",",$price,"\n";
		
	}
}

print   "<br><br><br>\n";
print   "</table>\n";
print   "<b>Number of Accounts: <input type=\"text\" name=\"num_accounts_checked\" value=\"",$num_accounts_checked,"\" size=4 style=\"border:none\" readonly></b><br>\n";   
print   "<b>Output Filename: </b><a href=\"",$csv_output_filename, "\">",$csv_output_filename,"</a><br>\n";
print   "</form>\n";
print   "</body>\n";
print   "</html>\n";
print   "<br><br>\n";

close(CSV_OUTPUT_FILE);



