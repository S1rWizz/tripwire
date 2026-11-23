First step to a functioning honeypot server.
Self hosted tripwire, that can send email alerts when a bait file has been accessed.

setup requires intermediate knowledge of apache2 server setup, http tunneling thru cloudflare and how to setup env variables. 

clone the repo.
From the project directory, create and activate the virtual environment, then install the dependencies:

```powershell
py -3.13 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

In VS Code, select `.venv\Scripts\python.exe` as the Python interpreter. (or run '.\.venv\Scripts\Activate.ps1' in the terminal)

initialise the env file, add the following details:  
SENDER_EMAIL=example@gmail.com  
APP_PASSWORD= (this needs to be generated on the website)  
ALERT_TO_EMAIL=example.com  
SMTP_HOST=smtp.gmail.com  
SMTP_PORT=587  
 
initialise venv and install all dependencies. 
boot up ur vm, or just anywhere to run an apache2 server.
start your apache2 server.   
ENSURE THE SERVER IS ON SAME LAN AS PYTHON SCRIPT. 

Use cloudflare to tunnel the localhost.
on the apache2 server system - activate a python server to access the logs endpoint from the server. 
use this one liner: (or something similar depending on ur system)  

sudo python3 -m http.server 8080 --bind 0.0.0.0 --directory /var/log/apache2


run python script. 
itll ask for details:  
[?] Select format >> ( choose ur file format )  
[?] Tunnel link >> the tunnel link tht u got from cloudflare  
[?] Server IP  >> ur server ip (apache2)  
[?] Log HTTP port >> port number  
[?] Alert message >> (custom message tht u get on ur email)  

AND TA-DA its all done 
enjoy !

