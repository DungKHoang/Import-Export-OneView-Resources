## Import and Export OneView resources

Import-OVResources.ps1 and Export-OVResources.ps1 are PowerShell scripts that leverage HPE OneView PowerShell library and Excel to automate configuration of OneView OVResources
Export-OVresources.ps1 queries OneView to collect OV resources settings and save them in Excel spreadsheets.
Import-OVResources.ps1 reads Excel sheets for values of OV resources and generate PowerShell scripts to create OV resources in OneView destination. 




## Prerequisites
Both scripts require:
   * HPE OneView PowerShell library 5.40
   * ImportExcel module

Those modules can be pulled from the Microsoft PowerShell Gallery.

## Excel spreadsheet
OV-Template.xlsx is **required** and  used by Export-OVResources.ps1.
It must reside in the same folder as the scripts

## OneView environment
The scripts have been tested on
   * OneView appliances 5.2 -5.3 -5.4
   * OneView composers 5.2 -5.3 -5.4



## Export-OVResources.PS1 

Export-OVResources.ps1 is a PowerShell script that exports OV resources into Excel spreadsheets including:
   * Address Pool
   * Ethernet newtorks
   * Network set
   * FC / FCOE networks
   * SAN Manager
   * Storage Systems: 3PAR
   * Storage Volume templates
   * Storage Volumes
   * Logical InterConnect Groups
   * Uplink Sets
   * Enclosure Groups
   * Enclosures
   * DL Servers 
   * Network connections
   * Local Storage connections
   * SAN Storage connections
   * iLO settings
   * Server Profile Templates
   * Server Profiles

   * IP addresses used by Synergy components
   * WWN when there are FC networks in profile 
   * Firmware bundles
   * Time & Locale Settings
   * SNMP Settings
   * Scope Settings
   * Users 
   * Firmware Bundles
   * Backup Settings
   * Remote Support Settings
   * Proxy settings



## Syntax

### To export OneView resources

```
    .\Export-OVResources.ps1 -jsonConfigFiles <list-of-jsonfiles>

```
where jsonfile uses the follwoing syntax:
```
{                                         
     "ip":              "<OV-IP>",  
     "loginAcknowledge": "true",      
     "credentials" :    {               
         "userName":    "<admin-account>",         
         "password":    "<admin-password>",   
         "authDomain":  "<LDAP-domain> or <LOCAL>"       
      },                                  
     "api_version" :     "1200"         
}
```
The script will read the OV-Template located in the same folder and jsonfiles to connect to multiple Oneview instances, if needed.
It will then write down vlaues to Excel spreadsheets with names ExportFrom-<OV-IP>.xlsx
For example:
.\Export-OVResources.ps1 -jsonConfigFiles 192.168.1.51.json, 192.168.1.175.json 
will generate Excel files named as : ExportFrom-192.168.1.51.xlsx and ExportFrom-192.168.1.175.xlsx

**Note:** Ensure that you have OV-Template.xlsx in the same folder as the script




## Import-OVResources.PS1 

Import-OVResources.ps1 is a PowerShell script that configures OV resources based on Excel sheets including:
   * Address Pool
   * Ethernet networks
   * Network set
   * FC / FCOE networks
   * Logical InterConnect Groups
   * Uplink Sets
   * Enclosure Groups
   * Enclosures
   * Servers 
   * Network connections
   * Local Storage connections
   * iLO settings
   * Server Profile Templates
   * Server Profiles

   * Backup configuration
   * Firmware bundles
   * Time & Locale Settings
   * SNMP Settings
   * Alert Settings
   * Scope Settings
   * Users 
   * Firmware Bundles
   * Backup Settings
   * Proxy settings


## Syntax

There are 2 use cases:
   * Using an existing Excel file 
It assumes that you have an Excel sheet filled out with values of resources exported using the Export-OVresource.ps1 described above.
In this case, open the Excel file, and go to the sheet 'OV Destination'. Fill out with OV IP address and credentials for the OneView instance at destination.

   * Starting from scratch
 Copy the OV-Templates.xlsx to a new file.
 Open the Excel file, and go to the sheet 'OV Destination'. Fill out with OV IP address and credentials for the OneView instance at destination.
 Fill out other sheets with values for your new OneView resources.

The Import-OVResources.ps1 will read the 'OV destination' to generate commands to connect to the OneView instance at destination

 To generate the scripts for importing, run the command 


```
    .\Import-OVResources.ps1 -workbook < Excel file>

```

The script will:
* create sub-folders ( info existed): Appliance - Facilities - Hypervisors - Networking - Servers - Settings - Storage
* read the corresponding sheeet in Excel file and genearet PowerShell script. The script will be located in the appropriate folder. For example fcnetwork.ps1 will be located under networking
* create an AllScripts file that points to each individual script

## Actions
* You can review each script and modify values in parameters to match with your new environment
* You can execute each individual script to create corresponding OV resources
* All passwords are set to **REDACTED**, so ensure to update them with new values


Enjoy!



### To import  SAN Manager

```
    .\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVSANManagerCSV c:\SANManager.csv

```

### To import SAN Storage Systems

```
    .\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVSANStorageSystemCSV c:\SANStorageSystem.csv

```

### To import  FC networks

```
    .\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVFCNetworksCSV c:\FCNetworks.csv

```

### To import  Volume Templates

```
    .\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVStorageVolumeTemplateCSV c:\StorageVolumeTemplate.csv

```

### To import  Storage Volume

```
    .\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVStorageVolumeCSV c:\StorageVolume.csv

```

### To import  Logical Interconnect Group

```
    .\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVLogicalInterconnectGroup c:\LogicalInterconnectGroup.csv

```
### To import  UplinkSet

```
    .\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVuplinkSetCSV c:\Uplinkset.csv

```

### To import  Enclosure Group

```
.\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVEnclosureGroupCSV c:\EnclosureGroup.csv

```

### To import Enclosure

```

.\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVEnclosureCSV c:\Enclosure.csv

```

### To import DL servers

```

.\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVServerCSV c:\Server.csv

```


### To import  Server profile Template

```

.\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVProfileTemplateCSV c:\ProfileTemplate.csv -OVProfileLOCALStorageCSV c:\ProfileLOCALStorage.csv -OVProfileSANStorageCSV c:\ProfileSANStorage.csv -OVProfileConnectionCSV c:\ProfileConnection.csv

```

### To import  Server profile

```

.\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVProfileCSV c:\Profile.csv -OVProfileLOCALStorageCSV c:\ProfileLOCALStorage.csv -OVProfileSANStorageCSV c:\ProfileSANStorage.csv -OVProfileConnectionCSV c:\ProfileConnection.csv

```

### To import Backup Configurations
The login credential to the remote backup server needs to included in the CSV file.
```
.\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVBackupConfig c:\BackupConfig.csv

```

### To Import Backup Configurations

```

.\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVBackupConfig c:\BackupConfig.csv

```
### To import Firmware Repositories

```

.\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVFWReposCSV c:\FirmwareRepositories.csv

```

### To Import Time Locale

```

.\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVTimeLocaleCSV c:\TimeLocale.csv

```

### To import Proxy

```

.\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVProxyCSV c:\Proxy.csv

```
### To import SMTP

```

.\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVsmtpCSV c:\smtp.csv

```
### To import LDAP

```

.\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVldapCSV c:\ldap.csv

```
### To import alerts

```

.\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVAlertsCSV c:\alerts.csv

```

### To import OSDeployment

```

.\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVOSDeploymentCSV c:\OSDeployment.csv

```

### To import LDAPgroups

```

.\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVldapgroupsCSV c:\ldapgroups.csv

```

### To import scopes

```

.\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVscopesCSV c:\scopes.csv

```

### To import users

```

.\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVusersCSV c:\scopes.csv

```

### To Import RemoteSupport Config

```

.\Import-OVResources.ps1 -OVApplianceIP <OV-IP-Address> -OVAdminName <Admin-name> -OVAdminPassword <password> -OVRSConfigCSV c:\RSConfig.csv

```

-----------------------------------------------------------------------------------------------------------------------------------------------------------------------

