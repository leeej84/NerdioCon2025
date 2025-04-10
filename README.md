# NerdioCon2025
This repository contains all the scripts used in my presentation session to be shared with the community.

The topic of the presentation was cloud based performance testing. There are many solutions that are commercially available that enables the collection of user experience, application response times and performance metrics. In this session I dig into a homegrown solution that enable the capture of application response times and performance data.

Using PowerShell, Telegraf, Influx and Grafana all metrics are collected in Influx and then visualised using Grafana.

<img src=".\Dashboard.png">

The following items are included:

- Setup Scripts
  - Setup a Windows Influx Server also hosting Grafana
    - Download_Install_Influx_Grafana.ps1
  - Install Telegraf
    - Used to install telegraf on the target host machine
      - Install_Telegraf.ps1
  - Prepare a terminal server for multiple users - Install Office, Microsoft Edge and the necessary Windows Roles, also configures user accounts
    - Setup_TS_Server.ps1
  Configure Microsoft Office with a Test Activation Key for 5 days of usable Office usage
    - Activate_Office_Test_Mode.ps1
- Automation Scripts
  - A workload for each application
    - MSEdge - bbc.co.uk website
    - Excel
    - PowerPoint
    - Word
  - There is a config file for the Automation scripts that will dictate the number of test runs, this is stored in the repo and should be kept alongside the Manager.ps1 file that handles testing iterations.
- Launcher
  - An RDP session launcher automation script is provided that will create RDP sessions using user credentials in the Users.csv file.
    - Launcher.ps1

# Setup Instructions
You firstly need to build two virtual machines, one of the virtual machines will host Influx and Grafana, the second machine will be a Terminal Server host. The the sake of this readme - INFLUX01 and TSHOST01 respectively.

- Run <i>Download_Install_Influx_Grafana.ps1</i> on the INFLUX01 VM.
- Login in to Influx (INFLUX01:8086) using credentials - influx_admin / Password100
- Create an API key for Influx
  - Select the upload icon on the left > API Tokens
  - Select Generate an API Token
  - All Access
  - Telegraf is the description
  - Take a note of the API Token - we will need it later for Telegraf and the Automation Scripts.
- Run <i>Setup_TS_Server.ps1</i> on TSHOST01
- Run <i>Install_Telegraf.ps1</i> on TSHOST01
  - Navigate to C:\Program Files\Telegraf on TSHOST01 and you will see a telegraf.conf file. Edit this as an admin and navigate to [[outputs.influxdb_v2]] - there is a token parameter here - place in your API token from earlier.
  - Save the file and restart the Service "Telegraf"
  - Navigate to C:\Test_Scripts and create a new text file - Influx_Token.txt. 
  - Paste your influx token from earlier into this file. Save the file.
- Office 365 Activation
  - Run <i>Activate_Office_Test_Mode.ps1</i> on TSHOST01, this will put office into a 5 day grace period where the program will work even with activation prompts. The script will prompt to rearm if ran more than once.
- Launcher
  - From any windows machine you can now run <i>Launcher.ps1</i> with the User.csv file locally and it will launch as many RDP sessions as specified. When a user logs in with a name like "testuser" the Manager.ps1 file will automatically run and start the application simulations.
- Grafana Dashboard
  - The last step is to login to Grafana and import the Dashboard
  - navigate to INFLUX01:3000 and you'll be prompted to login to Grafana. admin/admin is the default and you'll be prompted to change this.
  - Select Dashboards on the left, Select "New" and then Import.
  - Browse to the Grafana Dashboard folder of the repo and import the dashboard.

I'm not officially supporting this project but please feel free to adapt and modify as fits your needs. 