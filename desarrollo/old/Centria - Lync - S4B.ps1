$S4B_TST = @("sip.","lyncdiscover.")
$S4B_SOA = @("_sip._tls.","_sipfederationtls._tcp.")

$DOM_TST = @(
"Centria.net",
"Urbanova.com.pe",
"estrategica.com.pe",
"ahr.com.pe",
"aesa.com.pe",
"constructoraaesa.com.pe",
"protepersa.com.pe",
"aporta.org.pe",
"bvo.com.pe",
"cmpiura.com")


foreach($DOM in $DOM_TST){

    Resolve-DnsName ($S4B_SOA[0] + $DOM) -server 8.8.8.8 -Type SRV -DnsOnly -ErrorAction SilentlyContinue
    Resolve-DnsName ($S4B_SOA[1] + $DOM) -server 8.8.8.8 -Type SRV -DnsOnly -ErrorAction SilentlyContinue 

}


foreach($DOM in $DOM_TST){

    Resolve-DnsName ($S4B_TST[0] + $DOM) -server 8.8.8.8 -Type CNAME -DnsOnly 
    Resolve-DnsName ($S4B_TST[1] + $DOM) -server 8.8.8.8 -Type CNAME -DnsOnly
}
