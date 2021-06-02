##############################################################################
#
#   Export-HPOVResources.ps1
#
#   - Export resources from OneView instances or Synergy Composers to Excel file
#
#   VERSION 5.0
#
# (C) Copyright 2013-2020 Hewlett Packard Enterprise Development LP
##############################################################################
<#
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

  .SYNOPSIS
     Export resources from OneView appliance or composers

  .DESCRIPTION
	 Export resources from OneView appliance or composers

  .EXAMPLE


    .\ Export-HPOVResources.ps1 -jsonConfigFiles 192.168.1.51.json, 192.168.1.175.json
   Export all OnevIew resources in Excel files:
   - ExportFrom-192.168.1.51.xlsx
   - ExportFrom-192.168.1.175.xlsx


  .PARAMETER jsonConfigFiles
	List of json files tp provide OneView credential and IP to connect to. Eamples of json file":
	{                                         
     "ip":              "192.168.1.51",  
     "loginAcknowledge": "true",      
     "credentials" :    {               
         "userName":    "administrator",         
         "password":    "password",   
         "authDomain":  "local"       
      },                                  
     "api_version" :     "1200"         
}                                         

#>




# ------------------ Parameters
Param ( 
        [string[]]$jsonConfigFiles              = [System.Collections.ArrayList]::new()

)


$DoubleQuote    = '"'
$CRLF           = "`r`n"
$Delimiter      = "\"   # Delimiter for CSV profile file
$SepHash        = ";"   # Use for multiple values fields
$SepChar        = '|'
$CRLF           = "`r`n"
$OpenDelim      = "={"
$CloseDelim     = "}"
$CR             = "`n"
$Comma          = ','
$HexPattern     = "^[0-9a-fA-F][0-9a-fA-F]:"
$SYN12K         = 'SY12000'




#------------------- Interconnect Types
[HashTable]$ICTypes         = @{
    "571956-B21" =  "FlexFabric" ;
    "455880-B21" =  "Flex10"     ;
    "638526-B21" =  "Flex1010D"  ;
    "691367-B21" =  "Flex2040f8" ;
    "572018-B21" =  "VCFC20"     ;
    "466482-B21" =  "VCFC24"     ;
    "641146-B21" =  "FEX"
}

# ---------------------- FC networks
[HashTable]$Global:FCNetworkFabricTypeEnum = @{

        FA           = 'FabricAttach';
        FabricAttach = 'FabricAttach';
        DA           = 'DirectAttach';
        DirectAttach = 'DirectAttach'

    }

[Hashtable]$global:GetUplinkSetPortSpeeds    = @{
        Speed0M   = "0";
        Speed100M = "100Mb";
        Speed10G  = "10Gb";
        Speed10M  = "10Mb";
        Speed1G   = "1Gb";
        Speed1M   = "1Mb";
        Speed20G  = "20Gb";
        Speed2G   = "2Gb";
        Speed2_5G = "2.5Gb";
        Speed40G  = "40Gb";
        Speed4G   = "4Gb";
        Speed8G   = "8Gb";
        Auto       = "Auto"
	}
	
[Hashtable]$global:SetUplinkSetPortSpeeds    = @{
        '0'    = "Speed0M";
        '100M' = "Speed100M";
        '100'  = "Speed100M";
        '10G'  = "Speed10G";
        '10'   = "Speed10G";
        '10M'  = "Speed10M";
        '1G'   = "Speed1G";
        '1'    = "Speed1G";
        '1M'   = "Speed1M";
        '20G'  = "Speed20G";
        '2G'   = "Speed2G";
        '2'    = "Speed2G";
        '2.5G' = "Speed2_5G";
        '40G'  = "Speed40G";
        '4G'   = "Speed4G";
        '8G'   = "Speed8G";
        '4'    = "Speed4G";
        '8'    = "Speed8G";
        'Auto' = "Auto"
    }

#-------------------- Resource type

[HashTable]$ResourceCategoryEnum = @{
    Baseline                    = 'firmware-drivers';
    ServerHardware              = 'server-hardware';
    ServerHardwareType          = 'server-hardware-types';
    ServerProfile               = 'server-profiles';
    ServerProfileTemplate       = 'server-profile-templates';
    Enclosure                   = 'enclosures';
    LogicalEnclosure            = 'logical-enclosures';
    EnclosureGroup              = 'enclosure-groups';
    Interconnect                = 'interconnects';
    LogicalInterconnect         = 'logical-interconnects';
    LogicalInterconnectGroup    = 'logical-interconnect-groups';
    LogicalSwitch               = 'logical-switches';
    LogicalSwitchGroup          = 'logical-switch-groups';
    UplinkSet                   = 'uplink-sets';
    SasInterconnect             = 'sas-interconnects';
    SasLogicalInterconnect      = 'sas-logical-interconnects';
    SasLogicalInterconnectGroup = 'sas-logical-interconnect-groups';
    ClusterProfile              = 'hypervisor-cluster-profiles';
    ClusterNode                 = 'hypervisor-host-profiles';
    HypervisorManager           = 'hypervisor-managers';
    HypervisorCluster           = 'hypervisor-cluster-profiles';
    FabricManager               = 'fabric-managers';
    FabricManagerTenant         = 'tenants';
    RackManager                 = 'rack-managers';
    NetworkSet                  = 'network-sets';
    EthernetNetwork             = 'ethernet-networks';
    StorageVolumeSet            = 'storage-volume-sets';
    StorageVolume               = 'storage-volumes';
    StorageVolumeTemplate       = 'storage-volume-templates';
    StoragePool                 = 'storage-pools';
    IPv4Subnet                  = 'id-range-IPv4-subnet';
    IPv4Range                   = 'id-range-IPv4';
    IPv6Subnet                  = 'id-range-IPv6-subnet';
    IPv6Range                   = 'id-range-IPv6';
    LogicalJBOD                 = 'sas-logical-jbods';
    Drive                       = 'drives';
    DriveEnclosure              = 'drive-enclosures'
}

[Hashtable]$ServerProfileConnectionBootPriorityEnum = @{
	none           = 'NotBootable';
	NotBootable    = 'NotBootable';
	Primary        = 'Primary';
	Secondary      = 'Secondary';
	IscsiPrimary   = 'Primary';
	IscsiSecondary = 'Secondary';
	LoadBalanced   = 'LoadBalanced'
}
[Hashtable]$ServerProfileSanManageOSType            = @{
	CitrixXen  = "Citrix Xen Server 5.x/6.x";
	CitrisXen7 = "Citrix Xen Server 7.x";
	AIX        = "AIX";
	IBMVIO     = "IBM VIO Server";
	RHEL4      = "RHE Linux (Pre RHEL 5)";
	RHEL3      = "RHE Linux (Pre RHEL 5)";
	RHEL       = "RHE Linux (5.x, 6.x, 7.x)";
	RHEV       = "RHE Virtualization (5.x, 6.x)";
	RHEV7      = "RHE Virtualization 7.x";
	VMware     = "VMware (ESXi)";
	Win2k3     = "Windows 2003";
	Win2k8     = "Windows 2008/2008 R2";
	Win2k12    = "Windows 2012 / WS2012 R2";
	Win2k16    = "Windows Server 2016";
	OpenVMS    = "OpenVMS";
	Egenera    = "Egenera";
	Exanet     = "Exanet";
	Solaris9   = "Solaris 9/10";
	Solaris10  = "Solaris 9/10";
	Solaris11  = "Solaris 11";
	ONTAP      = "NetApp/ONTAP";
	OEL        = "OE Linux UEK (5.x, 6.x)";
	OEL7       = "OE Linux UEK 7.x";
	HPUX11iv1  = "HP-UX (11i v1, 11i v2)"
	HPUX11iv2  = "HP-UX (11i v1, 11i v2)";
	HPUX11iv3  = "HP-UX (11i v3)";
	SUSE       = "SuSE (10.x, 11.x, 12.x)";
	SUSE9      = "SuSE Linux (Pre SLES 10)";
	Inform     = "InForm"
}

[Hashtable]$SnmpAuthLevelEnum                    = @{
	None        = "noauthnopriv";
	AuthOnly    = "authnopriv";
	AuthAndPriv = "authpriv"
}
[Hashtable]$Snmpv3UserAuthLevelEnum              = @{
	None        = "None";
	AuthOnly    = "Authentication";
	AuthAndPriv = "Authentication and privacy"
}
[Hashtable]$SnmpAuthProtocolEnum                 = @{

	'none'   = 'none';
	'md5'    = 'MD5';
	'SHA'    = 'SHA';
	'sha1'   = 'SHA1';
	'sha2'   = 'SHA2';
	'sha256' = 'SHA256';
	'sha384' = 'SHA384';
	'sha512' = 'SHA512'

}
[Hashtable]$SnmpPrivProtocolEnum                 = @{
	'none'    = 'none';
	'aes'     = "AES128";
	'aes-128' = "AES128";
	'aes-192' = "AES192";
	'aes-256' = "AES256";
	'aes128'  = "AES128";
	'aes192'  = "AES192";
	'aes256'  = "AES256";
	'des56'   = "DES56";
	'3des'    = "3DES";
	'tdea'    = 'TDEA'
}
[Hashtable]$ApplianceSnmpV3PrivProtocolEnum      = @{
	'none'   = 'none';
	"des56"  = 'DES';
	'3des'   = '3DES';
	'aes128' = 'AES-128';
	'aes192' = 'AES-192';
	'aes256' = 'AES-256'
}

[Hashtable]$consistencyCheckingEnum 			= @{
	'NotChecked'	= 'None';
	'Not Checked'	= 'None';
	'UnChecked'		= 'None';
	'None'			= 'None';

	'Checked'		= 'Exact';
	'ExactMatch'	= 'Exact';
	'Exact'			= 'Exact';

	'ExactMinimum'	= 'Minimum';
	'Minimum'		= 'Minimum'
}

[HashTable]$iLOPrivilgeParamEnum     = @{
    'userConfigPriv'			= ' -AdministerUserAccounts $True '  ;
    'remoteConsolePriv'			= ' -RemoteConsole $True '			 ;
    'virtualMediaPriv'			= ' -VirtualMedia $True '			 ;
    'virtualPowerAndResetPriv'	= ' -VirtualPowerAndReset $True '	 ;
    'iloConfigPriv'				= ' -ConfigureIloSettings $True '	 ;
    'loginPriv'					= ' -Login $True '					 ;
    'hostBIOSConfigPriv'		= ' -HostBIOS $True '				 ;
    'hostNICConfigPriv'			= ' -HostNIC $True '				 ;
    'hostStorageConfigPriv'		= ' -HostStorage $True '			 
}

# ----------------------------------------------------
#
# 		OneView configuration
#
# ----------------------------------------------------

class OVnetwork
{
	[string]$hostName	
	[string]$domainName	
	[string]$ipV4	
	[string]$app1Ipv4	
	[string]$app2IpV4	
	[string]$ipv4Subnet	
	[string]$ipv4Gateway	
	[string]$ipv4Dns	
	[string]$ipV6	
	[string]$app1Ipv6	
	[string]$app2Ipv6	
	[string]$ipv6Subnet	
	[string]$ipv6Gateway	
	[string]$ipv6Dns	

}

class OVsec
{
	[string]$TLSname	
	[string]$mode	
	[string]$modeIsEnabled	
	[string]$enabled	
	[string]$cipherSuites	
}	

class OVAuth
{
	[string]$enable2FactorAuthentication	
	[string]$strictEnforcement	
	[string]$allowLocalLogin	
	[string]$allowEmergencyLogin	
	[string]$emergencyLoginType	
	[string]$directoryDomain	
	[string]$directoryDomainType	
	[string]$smartCardLoginOnly	
	[string]$validationOIDs							

}

# ----------------------------------------------------
#
# 		OneView settings
#
# ----------------------------------------------------

Class BackupConfig
{
	[boolean]$enabled
	[string]$remoteServerName
	[string]$remoteServerDir
	[string]$protocol
	[string]$port
	[string]$userName
	[string]$password 		= '***REDACTED***'
	[string]$scheduleInterval					# DAILY - WEEKLY
	[string]$scheduleDays						# MON|FRI|WED
	[string]$scheduleTime
	[string]$remoteServerPublicKey


}

Class Scope
{
	[string]$Name
	[string]$Description
	[string]$ResourceName
	[string]$ResourceType
}

Class ApplianceTimeLocale
{	
	[string]$locale          
	[string]$timezone        
	[string]$ntpServers      
	[string]$pollingInterval 
	[string]$syncWithHost
}

Class SmtpConfig
{             
	[string]$senderEmailAddress 
	[string]$password 		= '***REDACTED***'          
	[string]$smtpServer         
	[string]$smtpPort           
	[string]$smtpProtocol       
	[string]$alertEmailDisabled
	[string]$alertEmailFilters  
}

class firmwareBundle 
{
	[string]$name
	[string]$isofile
}

class repository
{
	[string]$name
	[string]$repositoryUrl
	[string]$directory
	[string]$username
	[string]$password 		= '***REDACTED***'
}

class proxy
{
	[string]$server
	[string]$port
	[string]$protocol
	[string]$username
}

class addressPool
{
	[string]$name
	[string]$poolType
	[boolean]$enabled
	[string]$rangeCategory
	[string]$startAddress
	[string]$endAddress
	[string]$networkId
	[string]$subnetmask
	[string]$gateway
	[string]$dnsServers
	[string]$domain
}

# ----------------------------------------------------
#
# 		OneView resources
#
# ----------------------------------------------------
class storageSystem
{
	[string]$name
	[string]$hostName
	[string]$familyName
	[string]$userName
	[string]$password		= '***REDACTED***'	
	[string]$systemPorts
	[string]$domainName
	[string]$vips
	[string]$model
	[string]$serialNumber
	[string]$wwnn
	[string]$firmware
}

class storagePool
{
	[string]$name	
	[string]$description	
	[string]$storageSystem	
	[string]$storageDomain	
	[string]$state	
	[string]$totalCapacity	
	[string]$allocatedCapacity	
	[string]$freeCapacity	
	[string]$driveType	
	[string]$RAID				
}

class storageVolumeTemplate
{
	[string]$name	
	[string]$description	
	[string]$familyName
	[string]$storageSystem	
	[string]$storagePool
	[string]$lockStoragePool
	[string]$snapshotStoragePool
	[string]$lockSnapshotStoragePool
	[string]$capacity
	[string]$lockCapacity
	[string]$provisioningType
	[string]$lockProvisioningType
	[string]$enableCompression	
	[string]$lockEnableCompression
	[string]$shared	
	[string]$lockProvisionMode
	[string]$enableAdaptiveOptimization	
	[string]$lockAdaptiveOptimization
	[string]$dataProtectionLevel
	[string]$lockDataProtectionLevel
	[string]$enableDeduplication	
	[string]$lockEnableDeduplication	
	[string]$lockPerformancePolicy
	[string]$enableEncryption		
	[string]$lockEnableEncryption
	[string]$cachePinning	
	[string]$lockCachePinning
	[string]$volumeSet
	[string]$lockVolumeSet
	[string]$enableIOPSLimit		
	[string]$lockEnableIOPSLimit
	[string]$enableDataTransferLimit		
	[string]$lockDataTransferLimit
	[string]$folder
	[string]$lockFolder
	[string]$scopes

}

class storageVolume
{
	[string]$name	
	[string]$description
	[string]$volumeTemplate
	[string]$storageSystem	
	[string]$storagePool
	[string]$snapshotStoragePool
	[string]$capacity
	[string]$provisioningType
	[string]$usedBy							
	[string]$enableCompression	
	[string]$shared	
	[string]$enableAdaptiveOptimization	
	[string]$dataProtectionLevel
	[string]$enableDeduplication	
	[string]$enableEncryption		
	[string]$cachePinning	
	[string]$volumeSet
	[string]$enableIOPSLimit		
	[string]$enableDataTransferLimit		
	[string]$folder
	[string]$scopes

}

class LogicalJBOD
{
	[string]$name	
	[string]$description	
	[string]$driveType	
	[int]$NumberofDrives	
	[string]$minDriveSize	
	[string]$maxDriveSize	
	[string]$driveEnclosure
	[boolean]$eraseDataonDelete	
	[string]$scopes
}


class ethernetNetwork 					
{
	[string]$name
	[string]$type
	[string]$vlanId   
	[string]$ethernetNetworkType
	[string]$subnetID	
	[string]$ipV6subnetID
	[string]$typicalBandwidth
	[string]$maximumBandwidth
	[string]$smartLink
	[string]$privateNetwork
	[string]$purpose  
	[string]$scopes
}

class fcFcoeNetwork 					
{
	[string]$name
	[string]$type
	[string]$fabricType
	[string]$managedSan
	[string]$vlanId   
	[string]$typicalBandwidth
	[string]$maximumBandwidth
	[string]$autoLoginRedistribution
	[string]$linkStabilityTime
	[string]$scopes
}

class networkSet 					
{
	[string]$name
	[string]$typicalBandwidth
	[string]$maximumBandwidth
	[string]$networkSetType
	[string]$networks
	[string]$nativeNetwork
	[string]$scopes
}

class lig
{
	[string]$name      
	[string]$frameCount
	[string]$interconnectBaySet
	[string]$enclosureType
	[string]$fabricModuleType
	[string]$bayConfig
	[string]$redundancyType
	[string]$internalNetworks
	[string]$consistencyCheckingForInternalNetworks
	[string]$interconnectConsistencyChecking
	[string]$enableIgmpSnooping
	[string]$igmpIdleTimeoutInterval
	[string]$enableFastMacCacheFailover
	[string]$macRefreshInterval
	[string]$enableNetworkLoopProtection
	[string]$enablePauseFloodProtection
	[string]$enableRichTLV 
	[string]$enableTaggedLldp
	[string]$lldpIpAddressMode
	[string]$lldpIpv4Address
	[string]$lldpIpv6Address
	[string]$enableStormControl
	[string]$stormControlPollingInterval
	[string]$stormControlThreshold
	[string]$qosconfigType
	[string]$downlinkClassificationType
	[string]$uplinkClassificationType
	[string]$scopes
	[string]$snmpConsistencyChecking


}


class uplinkset
{
	[string]$ligName
	[string]$name  
	[string]$networkType
	[string]$networks
	[string]$networkSets
	[string]$nativeNetwork
	[Boolean]$Trunking 
	[string]$fabricModuleName
	[string]$logicalPortConfigInfos
	[string]$fcUplinkSpeed 
	[string]$loadBalancingMode
	[string]$lacpTimer
	[string]$primaryPort
	[string]$privateVLanDomains
	[string]$consistencyChecking
} 


# -- snmp 
class snmpConfiguration
{
	[string]$source						= 'Appliance'			# either 'Appliance'or lig_name
	[string]$communityString	
	[string]$contact	
	[string]$accessList	
	[string]$engineId
}

class snmpV3User
{
	#[boolean]$applianceSnmpUser		= $True 
	[string]$source 					= 'Appliance'			# either 'Appliance'or lig_name
	[string]$userName 					
	[string]$securityLevel 										# None - AuthOnly - AuthAndPriv
	[string]$authProtocol										# None - MD5 - SHA - SHA1 - SHA256 - SHA384 - SHA512
	[string]$authPassword 				=  '***REDACTED***'
	[string]$privacyProtocol									# None - des56 -3des - aes128 - aes192 -aes 256
	[string]$privacyPassword			=  '***REDACTED***'
}


class snmpTrap
{ 		
	[string]$source 					= 'Appliance'			# either 'Appliance'or lig_name	
	[string]$format						= 'SnmpV3'				# SnmpV1 or SnmpV3	
	[string]$destinationAddress
	[string]$port	
	[string]$communityString
	[string]$userName
	[string]$trapType	 										# trap or Inform
	[string]$engineId											# prefix with 0x - followed by even number of 10 to 64 Hex digits
	[string]$trapSeverities	
	[string]$vcmTrapCategories	
	[string]$enetTrapCategories	
	[string]$fcTrapCategories

}


class eg 
{
	[string]$name
	[string]$logicalInterConnectGroupMapping
	[int]$enclosureCount
	[string]$IPv4AddressingMode
	[string]$IPv4Range
	[string]$IPv6AddressingMode
	[string]$IPv6Range
	[string]$powerMode
	[string]$deploymentMode
	[string]$deploymentNetwork
	[string]$scopes					

}

class le
{
	[string]$name
	[string]$enclosureSerialNumber
	[string]$enclosureName
	[string]$enclosureNewName
	[string]$enclosureGroup
	[string]$manualAddresses
	[string]$firmwareBaseline	
	[Boolean]$forceInstallFirmware	
	[string]$logicalInterconnectUpdateMode	
	[Boolean]$updateFirmwareOnUnmanagedInterconnect	
	[Boolean]$validateIfLIFirmwareUpdateIsNonDisruptive
	[string]$scopes									

}

class server
{
	[string]$name	
	[string]$serverName	
	[string]$description	
	[string]$model	
	[int]$processorCount	
	[int]$processorCoreCount	
	[string]$processorSpeed	
	[string]$processorType	
	[int]$memory	
	[string]$serialNumber	
	[string]$virtualSerialNumber	
	[string]$profileName	
	[string]$scopes	
}

class sht
{
	[string]$name	
	[string]$model
	[string]$formFactor	
	[string]$adapterModel
	[string]$adapterSlot
	
}


class spt
{
	[string]$name
	[string]$description
	[string]$serverProfileDescription

	[string]$sht
	[string]$enclosureGroupName
	[string]$affinity

	[Boolean]$manageFirmware
	[string]$firmwareConsistencyChecking
	[string]$firmwareBaselineName
	[string]$firmwareInstallType
	[Boolean]$forceInstallFirmware
	[string]$firmwareActivationType
	[string]$firmwareSchedule
	
	[Boolean]$manageConnections
	[string]$connectionConsistencyChecking

	[string]$localStorageConsistencyChecking 	

	[string]$manageSanStorage						= $False
	[string]$sanStorageConsistencyChecking			
	[string]$hostOSType
	

	[Boolean]$manageBootMode
	[string]$bootModeConsistencyChecking
	[string]$mode	
	[string]$pxeBootPolicy
	[string]$secureBoot


	[Boolean]$manageBootOrder
	[string]$bootOrderConsistencyChecking
	[string]$order
	

	[Boolean]$manageBios
	[string]$biosConsistencyChecking
	[string]$overriddenSettings

	[Boolean]$manageIlo
	[string]$iloConsistencyChecking

	[string]$macType
	[string]$wwnType
	[string]$serialNumberType
	[string]$iscsiInitiatorNameType
	[Boolean]$hideUnusedFlexNics
	[string]$scopes

	
}


class sp
{
	[string]$name
	[string]$description
	[string]$serverProfileTemplate
	[string]$serverHardware
	[string]$empty

	[string]$sht	
	[string]$enclosureGroupName

	[string]$affinity

	[Boolean]$manageFirmware
	[string]$firmwareBaselineName
	[string]$firmwareInstallType
	[Boolean]$forceInstallFirmware
	[string]$firmwareActivationType
	[string]$firmwareSchedule

	[Boolean]$manageLocalStorage
	[Boolean]$manageSanStorage			= $False
	[string]$hostOSType

	[Boolean]$manageBootMode
	[string]$mode	
	[string]$pxeBootPolicy
	[string]$secureBoot

	[Boolean]$manageBootOrder
	[string]$order
	
	[Boolean]$manageBios
	[string]$overriddenSettings

	[Boolean]$manageIlo

	[string]$macType
	[string]$wwnType
	[string]$serialNumberType
	[string]$serialNumber
	[string]$iscsiInitiatorNameType
	[string]$iscsiInitiatorName

	[Boolean]$hideUnusedFlexNics
	[string]$scopes
}

class connection
{
	[string]$profileName
	[string]$name
	[string]$id
	[string]$functionType
	[string]$network
	[string]$portId
	[string]$requestedMbps
	[Boolean]$boot
	[string]$priority
	[string]$bootVolumeSource
	[string]$targets				#HKD03

	
	[Boolean]$userDefined 			= $False			
	[string]$macType
	[string]$mac
	[string]$wwpnType
	[string]$wwnn
	[string]$wwpn
	[string]$lagName
	[string]$requestedVFs
	[boolean]$isolatedTrunk
	[string]$privateVlanPortType

	
}

class ilosetting
{
	[string]$profileName
	[string]$profileType  						# SPT or SP
	[string]$settingType
	[Boolean]$deleteAdministratorAccount
	[string]$adminPassword					=  '***REDACTED***'

	[string]$userName
	[string]$displayName
	[string]$userPassword					=  '***REDACTED***'
	[string]$userPrivileges

	[string]$directoryAuthentication
	[Boolean]$directoryGenericLDAP
	[string]$iloObjectDistinguishedName
	[string]$directoryPassword				
	[string]$directoryServerAddress
	[string]$directoryServerPort
	[string]$directoryUserContext
	[Boolean]$kerberosAuthentication
	[string]$kerberosRealm
	[string]$kerberosKDCServerAddress
	[string]$kerberosKDCServerPort
	[string]$kerberosKeytab
	
	[string]$groupDN
	[string]$groupSID
	[string]$groupPrivileges
	
	[string]$hostName
	
	[string]$primaryServerAddress
	[string]$primaryServerPort
	[string]$secondaryServerAddress
	[string]$secondaryServerPort
	[Boolean]$redundancyRequired
	[string]$groupName
	[string]$certificateName
	[string]$loginName
	[string]$keyManagerpassword				=  '***REDACTED***'
}

class spLocalStorage 
{
	[string]$profileName
	[string]$deviceSlot
	[string]$mode 							# RAID or HBA
	[string]$initialize
	[string]$driveWriteCache
	#[string]$predictiveSpareRebuild
	[string]$logicalDrives
	[string]$id 
	[string]$description
	[string]$raidLevel
	[string]$bootable
	[string]$driveTechnology
	[string]$numPhysicalDrives
	[string]$driveMinSize 
	[string]$driveMaxSize 
	[string]$accelerator
	[string]$eraseData
	[string]$persistent
}

class spSANStorage 
{
	[string]$profileName
	[string]$volumeName	
	[string]$volumeID
	[string]$volumeLUN	
	[string]$volumeLUNType	
	[string]$volumeStorageSystem	
	[string]$volumeBootPriority	
	[string]$volumeStoragePaths

}

#------------------- Functions
function Get-Header-Values([PSCustomObject[]]$ObjList)
{
    ForEach ($obj in $ObjList)
        {
            # --------
            # Get Properties name out PSCustomObject
            $Properties   = $obj.psobject.Properties
            $PropNames    = [System.Collections.ArrayList]::new()
            $PropValues   = [System.Collections.ArrayList]::new()

            ForEach ($p in $Properties)
            {
                $PropNames    += $p.Name
                $PropValues   += $p.Value
            }

           $header         = $PropNames -join $Comma
           $ValuesArray   += $($PropValues -join $Comma) + $CR
        }

    return $header, $ValuesArray
}


function Get-NamefromUri([string]$uri, $hostconnection)
{
    $name = ""

    if ($Uri)
    {
        try
        {
            $name   = (Send-HPOVRequest -uri $Uri  -hostName $hostconnection).name 
        }
        catch
        {
            $name = ""
        }
    }

    return $name
}


function Get-TypefromUri([string]$uri, $hostconnection)
{
    $type = ""

    if ($Uri)
    {
        try
        {
            $type   = (Send-HPOVRequest -uri $Uri -hostName $hostconnection).Type
        }
        catch
        {
            $type = ""
        }
    }

    return $type
}







Function rebuild-fwISO($BaselineObj)
{

	# ----------------------- Rescontruct FW ISO filename
	# When uploading the FW ISO file into OV, all the '.' chars are replaced with "_"
	# so if the ISO filename is:        SPP_2018.06.20180709_for_HPE_Synergy_Z7550-96524.iso
	# OV will show $fw.ISOfilename ---> SPP_2018_06_20180709_for_HPE_Synergy_Z7550-96524.iso
	# 
	# This helper function will try to re-build the original ISO filename

	$newstr = $null

	switch ($BaselineObj.GetType().Fullname)
	{

		'HPOneView.Appliance.Baseline'
		{

			$arrList = New-Object System.Collections.ArrayList

			$StrArray = $BaselineObj.ResourceId.Split($Underscore)

			ForEach ($string in $StrArray)
			{

				[void]$arrList.Add($string.Replace($dot, $Underscore))

			}
			
			$newstr = "{0}.iso" -f [String]::Join($Underscore, $arrList.ToArray())                

		}

		'HPOneView.Appliance.BaselineHotfix'
		{

			$newStr     = $BaselineObj.FwComponents.Filename

		}

		default
		{

			$newstr = $null

		}

	}

	return $newStr
		
}

Function get-scopes($scopesUri,$hostconnection)
{
	$scopeNames	  				= [System.Collections.ArrayList]::new()
	$scopes 					= ""
	if ($scopesUri)
	{
		$scopedResource 	= send-HPOVRequest -uri $scopesUri
		$scopeUris 			= $scopedResource.scopeUris
		if ($scopeUris)
		{
			$scopeNames	 	= $scopeUris | % { Get-NamefromUri -uri $_ -hostconnection $hostconnection} 
		}
	}
	if ($scopeNames)
	{
		$scopes			   		= $scopeNames -join $SepChar  
	} 
	return $scopes
}

# ================================================================================================
#
#   Function Connect-Composers
#
# ================================================================================================
function Connect-Composers([string[]]$jsonConfigs)
{
	$connectionList = @()
    foreach($jsonFile in $jsonConfigs)
    {
        if (test-path $jsonFile)
        {
            $json           = type $jsonFile | convertfrom-Json 

            $userName       = $json.credentials.userName
            $securePassword = $json.credentials.password | ConvertTo-SecureString -AsPlainText -Force
            $authDomain     = if ($json.credentials.authDomain) {$json.credentials.authDomain} else {'local'}

            $ip             = $json.ip
            $xApi           = $json.api_version
            $loginAck       = [Boolean]$json.loginAcknowledge

            $cred           = New-Object System.Management.Automation.PSCredential  -ArgumentList $userName, $securePassword

            if ($ip)
            {
                $conn = Connect-HPOVMgmt -Hostname $ip -loginAcknowledge:$loginAck -AuthLoginDomain $authDomain -Credential $cred
				$connectionList 	+= $conn
            }
        } 
        else
        {
            write-host -foreground Yellow " cannot find json file for OneView ---> $jsonFile"
        }
	} 
	
	return $connectionList
}

# ================================================================================================
##
##      Function writeto-Excel
##
# ================================================================================================
function writeto-Excel($data, $sheetName, $destWorkbook)
{
	if ($destWorkBook)
	{
		if ($data -and (test-path -path $destWorkbook) )
		{
			
			$data | Export-Excel -path $destWorkBook -noHeader -StartRow $startRow -WorksheetName $sheetName
		}
	}
}                                                                                                                      


# --- Scope
function Export-Scopes($connection, $sheetName, $destWorkbook)
{
	
	$ValuesArray 	 = [System.Collections.ArrayList]::new()
	$namesArray 	 = $typesArray = [System.Collections.ArrayList]::new()

	$inputObject 	= Get-HPOVScope -ApplianceConnection $connection
	foreach ($_scope in $inputObject)
	{
		$scopeElement 				= New-Object -typeName Scope
		$_scope.members | % { $namesArray += $_.name; $typesArray += $_.type}
		$scopeElement.Name			=	$_scope.Name
		$scopeElement.Description 	=	$_scope.Description
		$scopeElement.resourceName	=	$namesArray -join $sepChar    # Transform array to string
		$scopeElement.resourceType	=	$typesArray -join $sepChar 

		$ValuesArray 				+= $scopeElement
	}
	writeto-Excel -data $ValuesArray -sheetName $sheetName -destworkBook $destWorkBook
    
}

# ------- Time locale
function Export-TimeLocale ($connection,$sheetName, $destWorkbook)
{
	$ValuesArray        		= [System.Collections.ArrayList]::new()
	
    $TimeLocale					= Get-HPOVApplianceDateTime			

	$_timelocale				= new-object -type ApplianceTimeLocale
	$_timelocale.locale     	= $TimeLocale.Locale.Split(".")[0]
	$_timelocale.timeZone       = $TimeLocale.TimeZone
	$_timelocale.ntpServers     = if ($TimeLocale.NtpServers) {$TimeLocale.NtpServers -join $SepChar} else {''}
	$_timelocale.pollingInterval = $TimeLocale.pollingInterval
	$_timelocale.syncWithHost    = $TimeLocale.SyncWithHost

	$ValuesArray				+= $_timelocale

	writeto-Excel -data $ValuesArray -sheetName $sheetName -destworkBook $destWorkBook

}

# --------------------------------------------------------
#
#             OneView configuration
#
# ---------------------------------------------------------

# ---- OV network
function export-HPOVnetwork ($connection, $sheetName, $destWorkbook)
{
	$ValuesArray        	= [System.Collections.ArrayList]::new()
	$InputObject 			= (Get-HPOVApplianceNetworkConfig -ApplianceConnection $connection).appliancenetworks

	$_net 					= new-object -typeName OVnetwork
	$_net.hostName			= $InputObject.hostname
	$_net.domainName		= $InputObject.domainName
	$_net.ipV4				= $InputObject.virtIpv4Addr
	$_net.app1Ipv4			= $InputObject.app1IpV4Addr
	$_net.app2Ipv4			= $InputObject.app2IpV4Addr
	$_net.ipv4Subnet		= $InputObject.ipv4Subnet 
	$_net.ipv4Gateway		= $InputObject.ipv4Gateway
	$_net.ipv4Dns			= if ($InputObject.ipv4NameServers) {$InputObject.ipv4NameServers -join '|' } else {''}
	$_net.ipV6				= $InputObject.virtIpv6Addr
	$_net.app1Ipv6			= $InputObject.app1IpV6Addr
	$_net.app2Ipv6			= $InputObject.app2IpV6Addr
	$_net.ipv6Subnet		= $InputObject.ipv6Subnet 
	$_net.ipv6Gateway		= $InputObject.ipv6Gateway
	$_net.ipv6Dns			= if ($InputObject.ipv6NameServers) {$InputObject.ipv6NameServers -join '|' } else {''}

	$ValuesArray			+= $_net

	writeto-Excel -data $ValuesArray -sheetName $sheetName -destworkBook $destWorkBook
}

# ---- OV security protocol
function export-HPOVsecurityProtocol ($connection, $sheetName, $destWorkbook)
{
	$ValuesArray        	= [System.Collections.ArrayList]::new()
	$secList	 			= Get-HPOVApplianceSecurityProtocol -ApplianceConnection $connection

	foreach ($InputObject in $secList)
	{
		$_sec 					= new-object -typeName OVsec
		$_sec.TLSname			= $InputObject.name
		$_sec.mode 				= $InputObject.mode
		$_sec.modeIsEnabled		= $InputObject.modeIsEnabled	
		$_sec.enabled			= $InputObject.enabled

		if ($InputObject.cipherSuites) 
		{
			$_cipherArr			= [System.Collections.ArrayList]::new()
			$_cipherArr			= $InputObject.cipherSuites | % { "{0}:{1}" -f $_.CipherSuiteName, $_.Enabled } 
			$_sec.cipherSuites	= $_cipherArr -join '|' 
		}



		$ValuesArray			+= $_sec
	}

	writeto-Excel -data $ValuesArray -sheetName $sheetName -destworkBook $destWorkBook
}

# ---- OV authentication
function export-HPOVauthentication ($connection, $sheetName, $destWorkbook)
{
	$ValuesArray        		= [System.Collections.ArrayList]::new()
	$twoFactorAuthentication	= Get-HPOVApplianceTwoFactorAuthentication -ApplianceConnection $connection					

	foreach ($InputObject in $twoFactorAuthentication)
	{
		$_auth 								= new-object -typeName OVAuth
		$_auth.enable2FactorAuthentication	= $InputObject.Enabled
		$_auth.strictEnforcement			= $InputObject.strictEnforcement
		$_auth.allowLocalLogin				= $InputObject.allowLocalLogin	
		$_auth.allowEmergencyLogin			= $InputObject.allowEmergencyLogin
		$_auth.emergencyLoginType			= $InputObject.emergencyLoginType

		$ValuesArray						+= $_auth
	}

	writeto-Excel -data $ValuesArray -sheetName $sheetName -destworkBook $destWorkBook
}

# --------------------------------------------------------
#
#             OneView settings
#
# ---------------------------------------------------------

# --- SMTP
function Export-SMTP ($connection, $sheetName, $destWorkbook)
{
	[Hashtable]$SmtpConnectionSecurityEnum = @{

		None     = 'PLAINTEXT';
		Tls      = 'TLS';
		StartTls = 'STARTTLS'

	}

	$ValuesArray        	= [System.Collections.ArrayList]::new()

	$Smtp					= Get-HPOVSMTPConfig -ApplianceConnection $connection
	$s						= new-object -TypeName smtpConfig
	$s.senderEmailAddress   = $Smtp.senderEmailAddress
	$s.smtpServer	        = $Smtp.smtpServer
	$s.smtpPort          	= $Smtp.smtpPort
	$s.smtpProtocol 		= ($SmtpConnectionSecurityEnum.GetEnumerator() | ? Value -eq $Smtp.smtpProtocol).Name

	$ValuesArray			+= $s

	writeto-Excel -data $ValuesArray -sheetName $sheetName -destworkBook $destWorkBook

}


#----------- backup 
Function Export-BackupConfig($connection, $sheetName, $destWorkbook)
{
	$valuesArray 		= [System.Collections.ArrayList]::new()
	$InputObject		= Get-HPOVAutomaticBackupConfig -ApplianceConnection $connection
	foreach ($bkp in $InputObject)
	{
		$remoteBackupEnabled		= $bkp.enabled
		if ($remoteBackupEnabled)
		{
			$_bkp 					= new-object -TypeName BackupConfig
			$_bkp.enabled 			= $bkp.enabled
			$_bkp.remoteServerName	= $bkp.remoteServerName 
			$_bkp.remoteServerDir	= $bkp.remoteServerDir
			$_bkp.protocol 			= $bkp.protocol
			$_bkp.port 				= $bkp.port
			$_bkp.userName			= $bkp.userName
			$_bkp.scheduleInterval	= $bkp.scheduleInterval
			if ($bkp.scheduleDays)
			{
				$_bkp.scheduleDays	= $bkp.scheduleDays -join '|'
			}
			$_bkp.scheduleTime		= $bkp.scheduleTime
			$_bkp.remoteServerPublicKey 	= $bkp.remoteServerPublicKey

			$valuesArray			+= $_bkp
		} 	
	}
	writeto-Excel -data $ValuesArray -sheetName $sheetName -destworkBook $destWorkbook
}


#----------- fw Baseline 
Function Export-fwBaseline($connection, $sheetName, $destWorkbook)
{
	$valuesArray 		= [System.Collections.ArrayList]::new()
	$InputObject		= get-HPOVbaseline -ApplianceConnection $connection

	foreach ($fwBase in $InputObject)
	{
		$fwElement 			= new-object -type firmwareBundle
		# - OV strips the dot from the ISOfilename, so we have to re-construct it
		$filename   		= rebuild-fwISO -BaselineObj $fwBase	
		$fwElement          = new-object -TypeName firmwareBundle
		$fwElement.name     = $fwBase.name
		$fwElement.isofile  = $filename

		$valuesArray        += $fwElement

	}
	writeto-Excel -data $ValuesArray -sheetName $sheetName -destworkBook $destWorkbook
}

#----------- Repository
Function Export-repository($connection, $sheetName, $destWorkbook)
{
	$valuesArray 	= [System.Collections.ArrayList]::new()

	$inputObject 			= Get-HPOVBaselineRepository -ApplianceConnection $connection
	foreach ($repo in $inputObject)
	{
		$_repo          		= new-object -type repository
		
		$_repo.name 			= $repo.name
		$_repo.repositoryUrl	= $repo.repositoryUrl
		$_repo.username 		= $repo.username
		$_repo.password 		= '***REDACTED***'

		$directory 				= $repo.directory
		if ($directory)
		{
			$_repo.directory 	= $directory
		}
		$valuesArray 		 	+= $_repo
	}

	writeto-Excel -data $ValuesArray -sheetName $sheetName -destworkBook $destWorkbook

	
}


#----------- proxy
Function Export-proxy($connection, $sheetName, $destWorkbook)
{

	$valuesArray 	= [System.Collections.ArrayList]::new()

	$inputObject 	= Get-HPOVApplianceProxy -ApplianceConnection $connection
	$proxy          = new-object -type proxy
	$server         = $InputObject.Server

	if ($server)
	{
		$proxy.protocol       = $InputObject.protocol 
		$proxy.port           = $InputObject.port 
		$proxy.username       = $InputObject.username 
		$proxy.server 		  = $server
		$valuesArray 		 += $proxy
	} 

	writeto-Excel -data $ValuesArray -sheetName $sheetName -destworkBook $destWorkBook
}

# ---------- Address Pool Range
function Export-AddressPoolRange($connection, $sheetName, $destWorkbook)
{

	$ValuesArray					= [System.Collections.ArrayList]::new()
	$inputObject 					= Get-HPOVAddressPoolRange -ApplianceConnection $connection
	foreach ($range in $inputObject)
	{
		$cat            			= $range.category
		$poolType       			= $cat.Split('-')[-1] 
		$_range 					= new-object -TypeName addressPool

		$_range.name			= $range.Name
		$_range.rangeCategory 	= $range.rangeCategory
		$_range.Enabled			= $range.Enabled 

		if ($poolType -eq 'IPv4')
		{
			$subnet 				= send-HPOVRequest -uri $range.subnetUri
			$_range.poolType	 	= $poolType
			$_range.startAddress 	= $range.StartStopFragments.startAddress
			$_range.endAddress 		= $range.StartStopFragments.endAddress
			$_range.networkId		= $subnet.networkId
			$_range.subnetmask		= $subnet.subnetmask
			$_range.gateway			= $subnet.gateway
			$dns 					= $subnet.dnsServers
			$_range.dnsServers		= if ($dns) { $dns -join $sepChar} else {''}

		}
		else 	# MAC - WWN - SN 
		{
			$_range.poolType	 	= $range.Name 
			$_range.startAddress 	= $range.startAddress
			$_range.endAddress 		= $range.endAddress


		}

		$ValuesArray				+= $_range

	}

	if ($valuesArray)
	{
		writeto-Excel -data $ValuesArray -sheetName $sheetName -destworkBook $destWorkBook
	}

}

# ---------- snmp  Users
function Export-snmpConfiguration($connection, $sheetName, $destWorkbook)
{

	$ValuesArray					= [System.Collections.ArrayList]::new()

	# Appliance snmp 
	$_c 							= new-object -TypeName snmpConfiguration
	$_c.communityString 			= Get-HPOVSnmpReadCommunity 		-ApplianceConnection $connection
	$_c.engineId 					= (Get-HPOVApplianceSnmpV3EngineId 	-ApplianceConnection $connection).EngineID

	$ValuesArray					+= $_c

	# --- Extract snmp settings from lig-snmp if existed
	$ligs					= Get-HPOVLogicalInterconnectGroup 	-ApplianceConnection $connection
	foreach ($l in $ligs)
	{
		$ligName 					= $l.name
		$settingsList 				= $l.snmpConfiguration
		foreach ($InputObject in $settingsList)
		{ 
			$_c 					= new-object -TypeName  snmpConfiguration
			$_c.source 				= $ligName
			$_c.communityString		= $InputObject.readCommunity
			$_c.contact				= $InputObject.systemContact
			if ($InputObject.accessList)
			{
				$_c.accessList		= $InputObject.accessList -join '|'
			}
			$isSnmpConfigurationEmpty = ($null -eq $InputObject.systemContact) -and ($null -eq $InputObject.readCommunity) -and ($null -eq $InputObject.trapDestinations) -and ($null -eq $inputObject.snmpUsers) -and ($null -eq $InputObject.accessList)
			
			if (-not $isSnmpConfigurationEmpty)
			{ 
				$ValuesArray			+= $_c
			}

		}
	}
	if ($valuesArray)
	{
		writeto-Excel -data $ValuesArray -sheetName $sheetName -destworkBook $destWorkBook
	}

}


# ---------- snmp  Users
function Export-snmpUsers($connection, $sheetName, $destWorkbook)
{

	$ValuesArray					= [System.Collections.ArrayList]::new()


	# --- Get appliance snmpv3 users
	$usersList					= Get-HPOVsnmpV3User 				-ApplianceConnection $connection
	foreach ($InputObject in $usersList)
	{
		$_u 						= new-object -TypeName  snmpV3User
		$_u.userName				= $InputObject.UserName
		$_u.securityLevel			= ($Snmpv3UserAuthLevelEnum.GetEnumerator() 		| where Value -eq $InputObject.securityLevel).Name
		$_u.authProtocol			= ($SnmpAuthProtocolEnum.GetEnumerator() 			| where Value -eq $InputObject.authenticationProtocol).Name
		$_u.privacyProtocol			= ($ApplianceSnmpV3PrivProtocolEnum.GetEnumerator() | where Value -eq $InputObject.privacyProtocol).Name 
		
		$ValuesArray				+= $_u

	}
	# --- Extract snmp v3 users from lig-snmp if existed
	$ligs					= Get-HPOVLogicalInterconnectGroup 	-ApplianceConnection $connection
	foreach ($l in $ligs)
	{
		$ligName 					= $l.name
		$usersList 					= $l.snmpConfiguration.snmpUsers
		foreach ($InputObject in $usersList)
		{
			$_u 					= new-object -TypeName  snmpV3User
			$_u.source 				= $ligName
			$_u.userName			= $InputObject.snmpV3UserName
			$_u.securityLevel		= ($Snmpv3UserAuthLevelEnum.GetEnumerator() | where Value -eq $InputObject.securityLevel).Name
			$_u.authProtocol		= ($SnmpAuthProtocolEnum.GetEnumerator() 	| where Value -eq $InputObject.v3authProtocol).Name
			$_u.privacyProtocol		= ($SnmpPrivProtocolEnum.GetEnumerator() 	| where Value -eq $InputObject.v3privacyProtocol).Name
			

			$ValuesArray				+= $_u

		}

	}

	if ($valuesArray)
	{
		writeto-Excel -data $ValuesArray -sheetName $sheetName -destworkBook $destWorkBook
	}



}

# ---------- snmp  Users
function Export-snmpTraps($connection, $sheetName, $destWorkbook)
{

	$ValuesArray					= [System.Collections.ArrayList]::new()

	# --- Get appliance snmpv3 traps
	$trapsList						= Get-HPOVApplianceTrapDestination 	-ApplianceConnection $connection
	foreach ($InputObject in $trapsList)
	{
		$_t 						= new-object -TypeName  snmpTrap
		$_t.format 					= $InputObject.type -replace 'TrapDestination', '' # type can be SnmpV1TrapDestination or SnmpV3TrapDestination
		$_t.destinationAddress		= $InputObject.destinationAddress
		$_t.port  					= $InputObject.port

		$isSnmpV3 					= $_t.format -eq 'SnmpV3'
		if ($isSnmpV3)
		{
			$_t.userName			= $InputObject.SnmpV3User
			$_t.trapType 			= $InputObject.TrapType
		}
		else #snmpV1 
		{
			$_t.communityString		= $InputObject.communityString
		}

		$ValuesArray				+= $_t

	}
	# --- Extract snmp traps from lig-snmp if existed
	$ligs					= Get-HPOVLogicalInterconnectGroup 	-ApplianceConnection $connection
	foreach ($l in $ligs)
	{
		$ligName 					= $l.name
		$trapsList 					= $l.snmpConfiguration.trapDestinations
		foreach ($InputObject in $trapsList)
		{
			$_t 					= new-object -TypeName  snmpTrap
			$_t.source 				= $ligName
			$_t.format 				= $InputObject.trapFormat 
			$_t.destinationAddress	= $InputObject.trapDestination
			$_t.port  				= $InputObject.port

			$isSnmpV3 				= $_t.format -eq 'SnmpV3'
			if ($isSnmpV3)
			{
				$_t.userName		= $InputObject.userName
				$_t.trapType 		= if ($InputObject.Inform) { 'TRAP'} else {'Inform'}
				$_t.engineId 		= $InputObject.engineId 
				if ($InputObject.trapSeverities)
				{
					$_t.trapSeverities 		= $InputObject.trapSeverities -join '|' 
				}
				if ($InputObject.vcmTrapCategories)
				{
					$_t.vcmTrapCategories 	= $InputObject.vcmTrapCategories -join '|' 
				}
				if ($InputObject.enetTrapCategories)
				{
					$_t.enetTrapCategories 	= $InputObject.enetTrapCategories -join '|' 
				}
				if ($InputObject.fcTrapCategories)
				{
					$_t.fcTrapCategories 	= $InputObject.fcTrapCategories -join '|' 
				}
			}
			else #snmpV1 
			{
				$_t.communityString		= $InputObject.communityString
			}

			$ValuesArray			+= $_t
		}

	}

	if ($valuesArray)
	{
		writeto-Excel -data $ValuesArray -sheetName $sheetName -destworkBook $destWorkBook
	}



}

# ------- Networks
function Export-Network ($connection,$sheetName, $destWorkbook)
{

	$ethValuesArray   = $fcValuesArray = [System.Collections.ArrayList]::new()

	$sheets  		  = $sheetName.Split($SepChar)
	$ethSheet 		  = $sheets[0]
	$fcSheet 		  = $sheets[1]

	$ListofNetworks   = Get-HPOVNetwork -ApplianceConnection $connection -ErrorAction Stop


    foreach ($net in $ListofNetworks )
    {
		$type        				= $net.type.Split("-")[0]   		# Value is like ethernet-networkVx or FC-networkVz

		
        # ----------------------- Construct Network information
		$name						= $net.name
        $tBandwidth					= (1/1000 * $net.DefaultTypicalBandwidth).ToString()
        $mBandwidth  				= (1/1000 * $net.DefaultMaximumBandwidth).ToString()
		
		# Scopes
		$scopes 					= get-scopes -scopesUri $net.scopesUri
		
		if ($type -eq 'Ethernet')
		{
			$_net   	 			= new-object -type ethernetNetwork
			$_net.name 				= $name
			$_net.type 				= $type
			$_net.typicalBandwidth 	= $tBandwidth
			$_net.maximumBandwidth 	= $mBandwidth
			$_net.scopes 			= $scopes
			$vLANType    			= $_net.ethernetNetworkType = $net.ethernetNetworkType
			$_net.vlanId     		= if ($vLANType -eq 'Tagged') { $net.vLanId } else { ''}
			$_net.smartlink   		= $net.SmartLink
			$_net.privateNetwork    = $net.PrivateNetwork
			$_net.purpose     		= $net.purpose
			# Valid only for Synergy Composer
			$subnetUri    			= $net.subnetUri	
			$ipV6subnetUri 			= $net.ipV6subnetUri

			$subnet 	  			= ""		
			if ( $subnetUri -and ($connection.ApplianceType -eq 'Composer') )
			{
				$ThisSubnet 		= Get-HPOVAddressPoolSubnet | Where-Object URI -eq $subnetURI
				if ($ThisSubnet)
					{ $subnet 		= $ThisSubnet.NetworkID }
			}
			$_net.subnetID 			= $subnet

			$subnet 	  			= ""
			if ( $ipV6subnetUri -and ($connection.ApplianceType -eq 'Composer') )
			{
				$ThisSubnet 		= Get-HPOVAddressPoolSubnet | Where-Object URI -eq $ipV6subnetURI
				if ($ThisSubnet)
					{ $subnet 		= $ThisSubnet.NetworkID }
			}
			$_net.ipV6subnetID 		= $subnet


			$ethValuesArray 	+= $_net
			
		}
		else    # Type = FC or FCOE
		{
			$_net   	 			= new-object -type fcFcoeNetwork
			$_net.name 				= $name
			$_net.typicalBandwidth 	= $tBandwidth
			$_net.maximumBandwidth 	= $mBandwidth
			$_net.scopes 			= $scopes
			$_net.type 				= $type
			$sanUri 				= $net.managedSanUri
			$_net.managedSan    	= if ($SanUri) { Get-NamefromUri -uri $sanUri -hostconnection $connection} else {''}
			if ($type -eq 'FC') 
			{ 
				$_net.fabricType 	= $net.fabricType
				$_net.autoLoginRedistribution = $net.autoLoginRedistribution
				$_net.linkStabilityTime	= $net.linkStabilityTime
			}
			else
			{
				$_net.vlanId      	=  $net.vLanId 
			}

			$fcValuesArray 			+= $_net
		}

    }

    if ($ethValuesArray)
    {
		writeto-Excel -data $ethValuesArray -sheetName $ethSheet -destworkBook $destWorkBook

	}
	
	if ($fcValuesArray)
    {
		writeto-Excel -data $fcValuesArray -sheetName $fcSheet -destworkBook $destWorkBook

    }
	

}


# ------- Network set
function Export-NetworkSet ($connection,$sheetName, $destWorkbook)
{
	$valuesArray 							= [System.Collections.ArrayList]::new()
	$ListofNetworkSet 						= Get-HPOVNetworkSet -ApplianceConnection $connection| Sort-Object Name
    # ---------------------- Construct Network Set Names
	foreach ($netset in $ListofNetworkSet)
	{
		$_netset 							= new-object -type networkSet
		$_netset.name						= $netset.name
        $_netset.typicalBandwidth			= (1/1000 * $netset.typicalBandwidth).ToString()
		$_netset.maximumBandwidth  			= (1/1000 * $netset.MaximumBandwidth).ToString()
		$_netset.networkSetType				= $netset.networkSetType

		$networkUris 						= $netset.networkUris
		$nativeNetworkUri 					= $netset.nativeNetworkUri

		if ($networkUris)
		{
			$networkNames 					= $networkUris | % { Get-NamefromUri -uri $_ -hostconnection $connection}
			$_netset.networks	 			= $networkNames -join $SepChar
		}

		if ($nativeNetworkUri)
		{
			$_netset.nativeNetwork			= Get-NamefromUri -uri $nativeNetworkUri -hostconnection $connection

		}

		$_netset.scopes 					= get-scopes -scopesUri $netset.scopesUri -hostconnection $connection

		$valuesArray					   += $_netset 
		
	}

	if ($ValuesArray)
    {
		writeto-Excel -data $ValuesArray -sheetName $SheetName -destworkBook $destWorkBook

	}
}


# ------- LogicalInterConnectGroup 
function Export-LogicalInterConnectGroup ($connection,$sheetName, $destWorkbook)
{

	$ICModuleTypes               = @{
		"VirtualConnectSE40GbF8ModuleforSynergy"   =  "SEVC40f8" ;
		"VirtualConnectSE100GbF32ModuleforSynergy" =  "SEVC100f32" ;
		"Synergy50GbInterconnectLinkModule"        =  "SE50ILM";
		"Synergy20GbInterconnectLinkModule"        =  "SE20ILM";
		"Synergy10GbInterconnectLinkModule"        =  "SE10ILM";
		"VirtualConnectSE16GbFCModuleforSynergy"   =  "SEVC16GbFC";
		"VirtualConnectSE32GbFCModuleforSynergy"   =  "SEVC32GbFC";
		"Synergy12GbSASConnectionModule"           =  "SE12SAS";
		"571956-B21"                               =  "FlexFabric";
		"455880-B21"                               =  "Flex10";
		"638526-B21"                               =  "Flex1010D";
		"691367-B21"                               =  "Flex2040f8";
		"572018-B21"                               =  "VCFC20";
		"466482-B21"                               =  "VCFC24";
		"641146-B21"                               =  "FEX"
	}

	$FabricModuleTypes           = @{
		"VirtualConnectSE40GbF8ModuleforSynergy"    =  "SEVC40f8" ;
		"VirtualConnectSE100GbF32ModuleforSynergy"  =  "SEVC100f32" ;
		"Synergy12GbSASConnectionModule"            =  "SAS";
		"VirtualConnectSE16GbFCModuleforSynergy"    =  "SEVCFC";
		"VirtualConnectSE32GbFCModuleforSynergy"    =  "SEVCFC";
	}

	$ICModuleToFabricModuleTypes = @{
		"SEVC40f8"                                  = "SEVC40f8" ;
		"SEVC100f32"                                = "SEV100f32" ;
		'SE50ILM'                                   = "SEV100f32" ;
		'SE20ILM'                                   = "SEVC40f8" ;
		'SE10ILM'                                   = "SEVC40f8" ;
		"SEVC16GbFC"                                = "SEVCFC" ;
		"SEVC32GbFC"                                = "SEVCFC" ;
		"SE12SAS"                                   = "SAS"
	}

	$ligValuesArray 			= [System.Collections.ArrayList]::new()
	$uplValuesArray 			= [System.Collections.ArrayList]::new()
	$snmpValuesArray 			= [System.Collections.ArrayList]::new()
	$snmpValuesArray			= @()
	$sheets 					= $sheetName.Split($SepChar)
	$ligSheetName 				= $sheets[0]
	$uplSheetName 				= $sheets[1]
	$snmpSheetName				= $sheets[2]

	$UnsupportedLigTypes 		= 'FEX', 'SAS'
	$LigType					= '' 

	$LIGs  						= Get-HPOVLogicalInterconnectGroup -ApplianceConnection $connection 

	foreach ($lig in $LIGs)
	{
		$_lig					= 	new-object -TypeName lig		
		$_lig.name          	= 	$lig.Name
		$enclosureType 	        = 	$lig.enclosureType
        $_lig.enclosureType     =   $enclosureType
        $_lig.interconnectBaySet=   $lig.interconnectBaySet


		switch ($lig.category)
		{

			'sas-logical-interconnect-groups'
			{

				$LigType = 'SAS'

			}

			'logical-interconnect-groups'
			{

				$LigType = 'EthernetFC'
				$snmp                   = $lig.snmpConfiguration
				$Telemetry              = $lig.telemetryConfiguration
					$sampleCount            = $Telemetry.sampleCount
					$sampleInterval         = $Telemetry.sampleInterval

				# The following is only applicable to Ethernet LIG, not FC or SAS
				$intNetworks				= [System.Collections.ArrayList]@()
				if ($lig.internalNetworkUris)
				{
					$intNetworks 			+= $lig.internalNetworkUris | % { Get-NamefromUri -uri $_ -hostconnection $connection}
					$_lig.internalNetworks 	= $intNetworks -join $SepChar
					

					$_lig.consistencyCheckingForInternalNetworks = if ($lig.consistencyCheckingForInternalNetworks) { $consistencyCheckingEnum.Item($lig.consistencyCheckingForInternalNetworks) } else {'None'}
				}
				$ethernetSettings					= $lig.ethernetSettings
				$_lig.enableIgmpSnooping   			= $ethernetSettings.enableIGMPSnooping
				$_lig.igmpIdleTimeoutInterval		= $ethernetSettings.igmpIdleTimeoutInterval
				$_lig.enableFastMacCacheFailover   	= $ethernetSettings.enableFastMacCacheFailover
				$_lig.macrefreshInterval     		= $ethernetSettings.macRefreshInterval
				$_lig.enableNetworkLoopProtection   = $ethernetSettings.enableNetworkLoopProtection
				$_lig.enablePauseFloodProtection    = $ethernetSettings.enablePauseFloodProtection
				$_lig.enableRichTLV          		= $ethernetSettings.enableRichTLV
				$_lig.enableTaggedLldp       		= $ethernetSettings.enableTaggedLldp 
				$_lig.lldpIpAddressMode				= $ethernetSettings.lldpIpAddressMode
				$_lig.lldpIpv4Address				= $ethernetSettings.lldpIpv4Address
				$_lig.lldpIpv6Address				= $ethernetSettings.lldpIpv6Address
				$_lig.enableStormControl 		    = $ethernetSettings.enableStormControl   
				$_lig.stormControlPollingInterval   = $ethernetSettings.stormControlPollingInterval 
				$_lig.stormControlThreshold		    = $ethernetSettings.stormControlThreshold 
				$_lig.interconnectConsistencyChecking = if ($ethernetSettings.consistencyChecking) { $consistencyCheckingEnum.Item($ethernetSettings.consistencyChecking) } else {'None'}


				$_lig.redundancyType         		= $lig.redundancyType


				# --- uplinkset
				$interconnects 						= $lig.interconnectMaptemplate.interconnectMapEntrytemplates
				$uplinkSets    						= $lig.uplinksets | Sort-Object Name
				foreach ($upl in $uplinkSets)
				{
					$_upl 						= new-object -TypeName uplinkset
					$_upl.ligName 				= $lig.name
					$_upl.name 					= $upl.name
					$networkType	 			= $upl.networkType
					$_upl.networkType 			= $networkType
					$_upl.ConsistencyChecking 	= if ($upl.ConsistencyChecking) { $consistencyCheckingEnum.Item($upl.ConsistencyChecking) } else {'None'}
					$networkUris 				= $upl.networkUris
					$uplLogicalPorts  			= $upl.logicalportconfigInfos

					switch ($networkType)
					{
						'Ethernet'
						{
							$ethType 					= $upl.ethernetNetworkType
							if ('Tagged' -ne $ethType)
							{
								$_upl.networkType   	= $upl.ethernetNetworkType
							}

							$_upl.loadBalancingMode		= $upl.loadBalancingMode
							$_upl.lacpTimer 			= $upl.lacpTimer
							$networkSetUris 			= $upl.networkSetUris
							$nativeNetworkUri 			= $upl.nativeNetworkUri
							$privateVLanDomains 		= $upl.privateVLanDomains 


							# Networksets
							if ($Null -ne $networkSetUris)
							{
								$networkSetUris     	= $networkSetUris | % { get-NamefromUri -uri $_ -hostconnection $connection}
								$_upl.networkSets 		= $networkSetUris -join $SepChar
							}

							# NativeNetwork 
							if ($Null -ne $nativeNetworkUri )
							{
								$_upl.nativeNetwork 	= Get-NamefromUri -uri $nativeNetworkUri -hostconnection $connection
							}

							#PrivateVLAN Domains
							if ($Null -ne $privateVLanDomains )
							{
								$isolatedNetworkNames 	= @()
								$isolatedNetworks 		= $privateVLanDomains.isolatedNetworks
								foreach ($isolatedNet in $isolatedNetworks)
								{
									$isolatedNetworkNames += $isolatedNet.name
								}		
								$_upl.privateVLanDomains  = $isolatedNetworkNames -join $SepChar
							}
						}

						'FibreChannel'
						{
							$_upl.Trunking  	= if ( $upl.fcMode -eq 'TRUNK') {$True} else {$False}
						}

					}

					# Networks
					if ($Null -ne $networkUris)
					{
						$networkUris 			= $networkUris | % { get-NamefromUri -uri $_ -hostconnection $connection}
						$_upl.networks 			= $networkUris -join $SepChar
					}

					# Uplink Ports
					$LocationPropName 			= 'logicalLocation'
                	$ValuePropName    			= 'relativeValue' 


					$UpLinkArray 		= [System.Collections.ArrayList]::new()
					$UpLinkSpeedArray	= [System.Collections.ArrayList]::new()
					$fcUplinkSpeed 		= ''


					foreach ($logicalPort in $uplLogicalPorts)
					{


	
						$Speed          = $UpLinkLocation = $Port = $icm = $null

						$ThisBayNumber  = ($logicalPort.$LocationPropName.locationEntries | Where-Object Type -eq 'Bay').$ValuePropName
						$ThisPortNumber = ($logicalPort.$LocationPropName.locationEntries | Where-Object Type -eq 'Port').$ValuePropName
						$ThisEnclosure  = ($logicalPort.$LocationPropName.locationEntries | Where-Object Type -eq 'Enclosure').$ValuePropName
						$entry 			= $GetUplinkSetPortSpeeds.GetEnumerator() | where Name -eq $logicalPort.desiredSpeed
		
						$UpLinkSpeedArray += $entry.value # Not sure about Arry
						$fcUplinkSpeed 	= $entry.value	

						# Loop through Interconnect Map Entry Template items looking for the provided Interconnet Bay number
						ForEach ($l in $interconnects) 
						{

							$ThisIcmEnclosureLocation = ($l.$LocationPropName.locationEntries | Where-Object { $_.type -eq "Enclosure" -and $_.$ValuePropName -eq $ThisEnclosure}).$ValuePropName
							$ThisIcmBayLocation       = ($l.$LocationPropName.locationEntries | Where-Object { $_.type -eq "Bay" -and $_.$ValuePropName -eq $ThisBayNumber}).$ValuePropName

							if ($enclosureType -eq $Syn12K) 
							{

								if ($ThisIcmBayLocation -and $l.enclosureIndex -eq $ThisEnclosure) 
								{
											
									$permittedInterconnectTypeUri = $l.permittedInterconnectTypeUri

								}

							}

							else
							{
							
								if ($l.$LocationPropName.locationEntries | Where-Object { $_.type -eq "Bay" -and $_.$ValuePropName -eq $ThisBayNumber }) 
								{
											
									$permittedInterconnectTypeUri = $l.permittedInterconnectTypeUri

								}
														
							}
							
						} 


						$PermittedInterConnectType 	= Send-HPOVRequest $permittedInterconnectTypeUri -Hostname $connection
						#Find Module Name
						$ICtypeName   				= $PermittedInterConnectType.Name
						$_upl.fabricModuleName		= $ICtypeName

						
						# 1. Find port numbers and port names from permittedInterconnectType
						$PortInfos     = $PermittedInterConnectType.PortInfos

						# 2. Find Bay number and Port number on uplinksets
						$ICLocation    = $icm.$RootLocationInfos.$subLocationInfo
						$ICBay         = ($ICLocation | Where-Object Type -eq "Bay").$ValuePropName
						$ICEnclosure   = ($IClocation | Where-Object Type -eq "Enclosure").$ValuePropName

						# 3. Get faceplate port name
						$ThisPortName   = ($PortInfos    | Where-Object PortNumber -eq $ThisPortNumber).PortName

						if ($ThisEnclosure -eq -1)    # FC module
						{

							$UpLinkLocation = "Bay{0}:{1}" -f $ThisBayNumber, $ThisPortName   # Bay1:1
						
						}

						else  # Synergy Frames or C7000
						{

							if ($enclosureType -eq $Syn12K) 
							{

								$UpLinkLocation = "Enclosure{0}:Bay{1}:{2}" -f $ThisEnclosure, $ThisBayNumber, $ThisPortName.Replace(":", ".")   # Enclosure#:Bay#:Q1.3; In $PortInfos, format is Q1:4, output expects Q1.4
							
							}

							else # C7000
							{

								$UpLinkLocation = "Bay{0}:{1}" -f $ThisBayNumber, $ThisPortName.Replace(":", ".")   # Bay#:Q1.3; In $PortInfos, format is Q1:4, output expects Q1.4
							
							}

						}

						[void]$UpLinkArray.Add($UpLinkLocation)

					}

					$UpLinkArray.Sort()

					# Uplink Ports
					$uplinkPortParam    = $uplinkPortCode    = $null

					if ($UplinkArray) 
					{
						$_upl.logicalPortConfigInfos 	= $UplinkArray -join $SepChar

					}

					#if ($UplinkSpeedArray) 
					#{
						#$_upl.fcUplinkSpeed 	= $UplinkSpeedArray -join $SepChar

					#}

					$_upl.fcUplinkSpeed = $fcUplinkSpeed -replace 'Gb',''


	

					# Add to Array
					$uplValuesArray				+= $_upl

				}

				# --- snmp
				if ($snmp)
				{
					$_lig.snmpConsistencyChecking	= if ($snmp.consistencyChecking)  { $consistencyCheckingEnum.Item($snmp.consistencyChecking) } else {'None'}
				}

			}





		} # end switch

		
		$FrameCount = $InterconnectBaySet = $frameCountParam = $null
		$intnetParam = $null

		# ----------------------------
		#     Find Interconnect devices
		$Bays         	= [System.Collections.ArrayList]::new()
		$UpLinkPorts 	= [System.Collections.ArrayList]::new()
		$Frames         = [System.Collections.ArrayList]::new()

		$LigInterconnects = $lig.interconnectMapTemplate.interconnectMapEntryTemplates | Where-Object { -not [String]::IsNullOrWhiteSpace($_.permittedInterconnectTypeUri) }

		$BayHashtable = New-Object System.Collections.Specialized.OrderedDictionary

		foreach ($ligIC in $LigInterconnects)
		{

			# -----------------
			# Locate the Interconnect device and its position
			$ICTypeuri  = $ligIC.permittedInterconnectTypeUri

			if ($enclosureType -eq $Syn12K)
			{

				$ICtypeName   		= (Get-NamefromUri -Uri $ICTypeUri -hostconnection $connection).Replace(' ','') # remove Spaces
				$ICmoduleName 		= $ICModuleTypes[$ICtypeName]
				$BayNumber   		= ($ligIC.logicalLocation.locationEntries | Where-Object Type -eq "Bay").RelativeValue
				$FrameNumber 		= [math]::abs(($ligIC.logicalLocation.locationEntries | Where-Object Type -eq "Enclosure").RelativeValue)

				$fabricModuleType 	= $ICModuleToFabricModuleTypes[$ICmoduleName] 

				if (-not ($BayHashtable.GetEnumerator() | Where-Object Name -eq "Frame$FrameNumber"))
				{

					$BayHashtable.Add("Frame$FrameNumber", (New-Object Hashtable))

				}

				# Use this hashtable to build the final string value for scriptCode
				$BayHashtable."Frame$FrameNumber".Add("Bay$BayNumber", $ICmoduleName)

			}

			else # C7K
			{
				$PartNumber   = (Send-HPOVRequest -Uri $ICTypeuri -Hostname $Connection).partNumber

				


				if ("FEX" -eq $ICModuleTypes[$PartNumber])
				{

					$LigType = 'FEX'

				}

				$ICmoduleName = $ICModuleTypes[$PartNumber]
				$BayNumber    = ($ligIC.logicalLocation.locationEntries | Where-Object Type -eq "Bay").RelativeValue

				[void]$Bays.Add(('Bay{0} = "{1}"' -f $BayNumber, $ICmoduleName)) # Format is xx=Flex Fabric

			}
		
		}

		## BayConfig

		[Array]::Sort($Bays)

		$BayConfig    = [System.Collections.ArrayList]::new()

		$_bayConfig = ""

		if ($enclosureType -eq $Syn12K)  # Synergy
		{

			# $BayConfigperFrame = [System.Collections.ArrayList]::new()
			$SynergyCode       = [System.Collections.ArrayList]::new()
			$CurrentFrame      = $null


			$f = 1

			# Process Bays parameter
			foreach ($b in ($BayHashtable.GetEnumerator() | Sort-Object Name))
			{

				$_bayConfig += "{0}={1}" -f $b.Name, '{'     # Framexx= {}

				$endDelimiter = $SepChar

				if ($f -eq $BayHashtable.Count)
				{

					$endDelimiter = $null

				}

				$_b = 1

				# Loop through ports
				ForEach ($l in ($b.Value.GetEnumerator() | Sort-Object Name))
				{

					$subEndDelimiter = '|'

					if ($_b -eq $b.Value.Count)
					{

						$subEndDelimiter = $null

					}

					$_bayConfig += "{0}='{1}'{2}" -f $l.Name, $l.Value, $subEndDelimiter   # Bayxx ='SEC40GB'

					$_b++

				}

				$_bayConfig += "{0}{1}" -f $CloseDelim, $CR

				$f++

			}

		}
		$_lig.bayConfig =  $_bayConfig
		
		# Fabric Module Type 
		$_lig.fabricModuleType  	= $fabricModuleType

		# Frame Count
		if ($enclosureType -eq $Syn12K )
		{
			$_lig.frameCount = $lig.EnclosureIndexes.Count

		}

		## Scopes


		$ligValuesArray += $_lig


	} # end foreach lig

	if ($ligValuesArray)
    {
		writeto-Excel -data $ligValuesArray -sheetName $ligSheetName -destworkBook $destWorkBook

	}

	if ($uplValuesArray)
    {
		writeto-Excel -data $uplValuesArray -sheetName $uplSheetName -destworkBook $destWorkBook

	}

	if ($snmpValuesArray)
    {

		writeto-Excel -data $snmpValuesArray -sheetName $snmpSheetName -destworkBook $destWorkBook

	}
}


# ------- StorageSystem 
function Export-StorageSystem ($connection,$sheetName, $destWorkbook)
{
	$ValuesArray 					= [System.Collections.ArrayList]::new()

	$storageSystemList 				= get-HPOVStorageSystem -ApplianceConnection $connection
   
	foreach ($InputObject in $storageSystemList)
	{
			$_sts					= new-object -TypeName storageSystem
			$_sts.name	 			= $InputObject.name
			$_sts.hostName 			= $InputObject.hostName
			$_sts.familyName 		= $InputObject.family
			$_sts.userName 			= $InputObject.credentials.username

			$_attributes			= $InputObject.deviceSpecificAttributes
			$_sts.model 			= $_attributes.model
			$_sts.serialNumber 		= $_attributes.serialNumber
			$_sts.wwnn				= $_attributes.wwn
			$_sts.firmware 			= $_attributes.firmware
			$_sts.domainName		= $_attributes.managedDomain

			$_portArray 			= [System.Collections.ArrayList]::new() 				
			foreach ($p in $InputObject.ports )
			{
				$_portArray 		+= "{0}:{1}" -f $p.name, $p.expectedNetworkName
			}
			if ($_portArray)
			{
				$_sts.SystemPorts  	= $_portArray -join '|'
			}

			##TBD - vips for Nimble - Primera

			$ValuesArray 			+= $_sts
	}

	if ($ValuesArray)
    {
		writeto-Excel -data $ValuesArray -sheetName $SheetName -destworkBook $destWorkBook

	}

}

			
# ------- StoragePool
function Export-StoragePool ($connection,$sheetName, $destWorkbook)
{
	$ValuesArray 					= [System.Collections.ArrayList]::new()

	$storagePoolList 				= get-HPOVStoragePool -ApplianceConnection $connection
   
	foreach ($InputObject in $storagePoolList)
	{
			$_stp					= new-object -TypeName storagePool
			$_stp.name	 			= $InputObject.name
			$_stp.description		= $InputObject.description
			$_stp.state				= $InputObject.state
			$_stp.storageSystem		= get-NamefromUri  -uri $InputObject.storageSystemUri -hostconnection $connection
			$_stp.totalCapacity		= "{0:n2}" -f ($InputObject.totalCapacity/1GB)
			$_stp.allocatedCapacity	= "{0:n2}" -f ($InputObject.allocatedCapacity/1GB)
			$_stp.freeCapacity		= "{0:n2}" -f ($InputObject.freeCapacity/1GB)

			$_attributes			= $InputObject.deviceSpecificAttributes
			$_stp.storageDomain		= $_attributes.domain
			$_stp.driveType 		= $_attributes.deviceType
			$_stp.RAID 				= $_attributes.SupportedRAIDLevel

			$ValuesArray 			+= $_stp
	}

	if ($ValuesArray)
    {
		writeto-Excel -data $ValuesArray -sheetName $SheetName -destworkBook $destWorkBook

	}
}

# ------- StorageVolumeTemplate
function get-ValueandLockProperty ($property)
{
	$_val 		= $_lock 	= ''
	if ($property)
	{
		$_val 	= $property.default
		$_lock 	= $property.meta.locked
	}
	return $_val, $_lock

}
function Export-StorageVolumeTemplate ($connection,$sheetName, $destWorkbook)
{
	$ValuesArray 					= [System.Collections.ArrayList]::new()

	$storageVolumeTemplateList 		= get-HPOVStorageVolumeTemplate -ApplianceConnection $connection
   
	foreach ($InputObject in $storageVolumeTemplateList)
	{
			$_svt							= new-object -TypeName storageVolumeTemplate
			$_svt.name	 					= $InputObject.name
			$_svt.description				= $InputObject.description
			$_svt.familyName 				= $InputObject.family

			$_prop 							= $InputObject.properties
			$_svt.capacity					= "{0}" -f ($_prop.size.default/1GB)

			$_svt.shared, $_svt.lockProvisionMode					= get-ValueandLockProperty -property $_prop.isShareable
			$_svt.provisioningType, $_svt.lockProvisioningType		= get-ValueandLockProperty -property $_prop.provisioningType
			
			
			$_sPoolUri, $_svt.lockStoragePool						= get-ValueandLockProperty -property $_prop.storagePool
			$_sPool 												= send-HPOVrequest -uri $_sPoolUri -hostName $connection

			$_svt.storagePool										= Get-NamefromUri -uri $_sPoolUri  -hostConnection $connection
			$_svt.storageSystem 									= Get-NamefromUri -uri $_sPool.storageSystemUri  -hostConnection $connection
			

			switch ($InputObject.family)
			{
				'StoreServ'
					{
						$_sPoolUri, $_svt.lockSnapshotStoragePool				= get-ValueandLockProperty -property $_prop.snapshotPool
						$_svt.snapshotStoragePool								= Get-NamefromUri -uri $_sPoolUri -hostConnection $connection
						$_svt.enableDeduplication,$_svt.lockEnableDeduplication = get-ValueandLockProperty -property $_prop.isDeduplicated

					}
				'StoreVirtual'
					{
						$_svt.dataProtectionLevel, $_svt.lockDataProtectionLevel			= get-ValueandLockProperty -property $_prop.dataProtectionLevel
						$_svt.enableAdaptiveOptimization, $_svt.lockAdaptiveOptimization	= get-ValueandLockProperty -property $_prop.isAdaptiveOptimizationEnabled

					}
			}		

					
			$ValuesArray 			+= $_svt
	}

	if ($ValuesArray)
    {
		writeto-Excel -data $ValuesArray -sheetName $SheetName -destworkBook $destWorkBook

	}
}

# ------- StorageVolume
function Export-StorageVolume ($connection,$sheetName, $destWorkbook)
{
	$ValuesArray 			= [System.Collections.ArrayList]::new()
	$storageVolumeList 		= get-HPOVStorageVolume -ApplianceConnection $connection

   
	foreach ($InputObject in $storageVolumeList)
	{
			$_sv							= new-object -TypeName storageVolume
			$_sv.name	 					= $InputObject.name
			$_sv.description				= $InputObject.description
			
			$_voltemplate	 				= Get-NamefromUri -uri $InputObject.volumeTemplateUri -hostConnection $connection
			$_sv.volumeTemplate				= if ($_voltemplate -notlike '*root*template*') {$_voltemplate} else {''}

			$_sPoolUri 						= $InputObject.storagePoolUri 
			$_sPool 						= send-HPOVrequest -uri $_sPoolUri -hostName $connection
			$_sv.storagePool				= Get-NamefromUri -uri $_sPoolUri  -hostConnection $connection
			$_sv.storageSystem 				= Get-NamefromUri -uri $_sPool.storageSystemUri  -hostConnection $connection

			$_sv.capacity					= "{0:n2}" -f ($InputObject.provisionedCapacity/ 1GB)
			$_sv.shared						= $InputObject.isShareable
			$_sv.ProvisioningType			= $InputObject.provisioningType
			$_prop 		 					= $InputObject.deviceSpecificAttributes
			$_sv.enableDeduplication		= $_prop.isDeduplicated
			$_sv.snapshotStoragePool		= Get-NamefromUri -uri $_prop.snapshotPoolUri -hostConnection $connection
			$_sv.dataProtectionLevel		= $_prop.dataProtectionLevel
			$_sv.enableAdaptiveOptimization = $_prop.isAdaptiveOptimizationEnabled												
				
			# Add 'usedBy for profiles'
			$_volUri 						= $InputObject.uri
			$_volArray						= [System.Collections.ArrayList]::new() 

			$uri 							= "/rest/index/associations?childUri={0}&name=server_profiles_to_storage_volumes" -f $_volUri
			$_members 						= (Send-HPOVRequest -uri $uri -hostName $connection).members
			foreach ($_m in $_members)
			{
				$_volArray 					+= Get-NamefromUri -uri $_m.parentUri -hostconnection $connection		
			}

			if ($_volArray)
			{
				$_sv.usedBy 			= $_volArray -join '|'
			}

			$ValuesArray 			+= $_sv
	}

	if ($ValuesArray)
    {
		writeto-Excel -data $ValuesArray -sheetName $SheetName -destworkBook $destWorkBook

	}
}


# ------- logical JBOD
function Export-logicalJBOD ($connection,$sheetName, $destWorkbook)
{
	$ValuesArray 					= [System.Collections.ArrayList]::new()

	$JBODList 						= get-HPOVlogicalJBOD -ApplianceConnection $connection

   
	foreach ($InputObject in $JBODList)
	{
			$_jbod							= new-object -TypeName logicalJBOD
			$_jbod.name	 					= $InputObject.name
			$_jbod.description				= $InputObject.description
			$_interface 					= $InputObject.interface
			$_media							= $InputObject.media
			$_jbod.driveType 				= if ($_media -eq 'SSD') { "$_interface$_media"} else {$_interface} 
			$_jbod.numberofDrives 			= $InputObject.numberofDrives
			$_jbod.minDriveSize 			= "{0:n2}" -f $InputObject.MinSize
			$_jbod.maxDriveSize 			= "{0:n2}" -f $InputObject.MaxSize
			$_drives 						= $InputObject.drives
			$_jbod.eraseDataOnDelete 		= $InputObject.eraseDataOnDelete
					
			if ($_drives)
			{
				$_drive 					= [string]$_drives[0]
				$_driveEnclosure			= $_drive.Split(':')[0] 		# "F1-CN75140CR5, bay 1
				$_driveEnclosure 			= $_driveEnclosure -replace '"',''	# Reove double quotes
				$_jbod.driveEnclosure		= $_driveEnclosure
			}

			$ValuesArray 					+= $_jbod
	}

	if ($ValuesArray)
    {
		writeto-Excel -data $ValuesArray -sheetName $SheetName -destworkBook $destWorkBook

	}
}


# ------- EnclosureGroup 
function Export-EnclosureGroup ($connection,$sheetName, $destWorkbook)
{
	$ValuesArray 					= [System.Collections.ArrayList]::new()

	$EGlist 						= get-HPOVEnclosureGroup -ApplianceConnection $connection
   
	foreach ($EG in $EGlist)
	{
			$_eg					= new-object -TypeName eg #enclosureGroup
            $_eg.name               = $EG.name
            $_eg.enclosureCount     = $EG.enclosureCount
            $_eg.powerMode          = $EG.powerMode
            $scopesUri              = $EG.scopesUri

            $manageOSDeploy     	= $EG.osDeploymentSettings.manageOSDeployment
            $deploySettings         = $EG.osDeploymentSettings.deploymentModeSettings
            $_eg.deploymentMode     = $deploySettings.deploymentMode

            
            $enclosureType          = $EG.enclosureTypeUri.split('/')[-1]
            $ICnetBayMappings       = $EG.interConnectBayMappings | Where-Object EnclosureIndex -eq $Null | Where-Object  logicalInterconnectGroupUri -ne $Null 
            $ICbayMappings          = $EG.interConnectBayMappings | Where-Object EnclosureIndex -ne $Null  | Where-Object logicalInterconnectGroupUri -ne $Null | Sort-Object enclosureIndex, interconnectBay
			$enclosureCount         = $EG.enclosureCount
			
			$_eg.ipV4AddressingMode = $EG.ipAddressingMode.replace('IpPool', 'AddressPool')
            $ipRangeUris            = $EG.ipRangeUris
			$ipV4AddressType		= $_eg.ipV4AddressingMode 

			$_eg.ipV6AddressingMode = $EG.ipV6AddressingMode.replace('IpPool', 'AddressPool')
            $ipV6RangeUris          = $EG.ipV6RangeUris
			$ipV6AddressType		= $_eg.ipV6AddressingMode 

            # --- Find Enclosure Bay Mapping
            ###

            $BayHashtable           = New-Object System.Collections.Specialized.OrderedDictionary
            $CurrentLigName         = ""
            $CurrentLigVarName      = $NULL
            $BayHashtable           = New-Object System.Collections.Specialized.OrderedDictionary            

            # Collect IC for Ethernet first
            if ($ICnetBayMappings)
            {
                

                if ($enclosureType -eq $Syn12K)
                {
                    $thisLIGName        = Get-NamefromUri -Uri $ICNetBayMappings[0].logicalInterconnectGroupURI -hostconnection $connection
                    $LigVarName         = $thisLIGName

                    for ($i=1;$i -le $EnclosureCount;$i++) 
					{
                        $FrameID        = 'Frame{0}' -f $i
                        $BayHashtable.Add($FrameID,$LigVarName)
                    }
                }
                else   # C7000
                {
                    foreach ($LIG in $ICnetBayMappings)
                    {  
                        $thisLIGName            = Get-NamefromUri -Uri $LIG.logicalInterconnectGroupURI -hostconnection $connection
                        $icBayID                = '{0}' -f $LIG.InterconnectBay
                        if ($thisLIGName -ne $CurrentLigName)
                        {
							$CurrentLigName     = $thisLIGName
							$BayHashtable.Add($ICbayID,$CurrentLigName)
                        }                        
                    }
                } 
                
            }

            $netBayInterconnect     = 3
            if ($ICbayMappings)
            {
                
                ForEach ($LIG in $ICBayMappings)
                {

                    if ($LIG.InterconnectBay -lt $netBayInterconnect)
                    {
                        $thisLIGName    = Get-NamefromUri -Uri $LIG.logicalInterconnectGroupURI -hostconnection $connection
                        $LigVarName     = $thisLIGName
                        
                        # Multi or specific frame configuration
                        $FrameID        = 'Frame{0}' -f $LIG.enclosureIndex
                        
                        $entry          = $BayHashtable.GetEnumerator() | Where-Object Name -eq $FrameID
                        $eValue         = $entry.Value
                        if ($eValue -notlike "*$LigVarName*")   
                        {
                            $eValue     += ",$LigVarName"
                            $BayHashtable.set_item($frameID,$eValue)
                        }

                    }

                }
            }

            if ($BayHashtable)
            {
                $c = 1

                ForEach ($l in $BayHashtable.GetEnumerator())
                {
                    $endDelimiter = $SepChar

                    if ($c -eq $BayHashtable.Count)
                    {

                        $endDelimiter = $null

                    }

					$_eg.logicalInterConnectGroupMapping += "{0} = {1}{2}" -f $l.Name, $l.Value, $endDelimiter
                    $c++

                }

            }
            
        

            if ($enclosuretype -eq $SYN12K)
            {

                #---- IP V4 Address Pool

                if($ipV4AddressType -eq 'AddressPool')
                {

                    $RangeNames = [System.Collections.ArrayList]::new()

                    foreach ($uri in $ipRangeUris)
                    {

                        $rangeName          = Get-NamefromUri -Uri $uri -hostconnection $connection
                        [void]$RangeNames.Add($rangeName)
                    
                    }
					$_eg.IPv4Range 	= $RangeNames -join $Comma

				}
				
				#---- IP V6 Address Pool

				if($ipV6AddressType -eq 'AddressPool')
				{

					$RangeNames = [System.Collections.ArrayList]::new()

					foreach ($uri in $ipV6RangeUris)
					{

						$rangeName          = Get-NamefromUri -Uri $uri -hostconnection $connection
						[void]$RangeNames.Add($rangeName)
					
					}
					$_eg.IPv6Range 	= $RangeNames -join $Comma

				}

            }

            # --- OS Deployment with IS
            $OSdeploymentParam           = $null

            if ($manageOSDeploy)
            {

                if ($deploymentMode -eq 'External')
                {

                    $_eg.deploymentNetwork	= Get-NamefromUri -Uri $deploySettings.deploymentNetworkUri -hostconnection $connection

                }

            }

            # EG script
            #$uri = $EG.uri + '/script'
            #$egScript = Send-HPOVRequest -Uri $uri -Hostname $connection


            # Scopes

            
            $ResourceScope = Send-HPOVRequest -Uri $scopesUri -Hostname $Connection


            $n = 1
			$ScopeNamesArray 					= [System.Collections.ArrayList]::new()
            
            if (-not [String]::IsNullOrEmpty($ResourceScope.scopeUris))
            {

                ForEach ($scopeUri in $ResourceScope.scopeUris)
                {
                    $scopeName = Get-NamefromUri -Uri $scopeUri -hostconnection $connection
					[void]$ScopeNamesArray.Add($scopeName)
                }

                $_eg.scopes 	= $ScopeNamesArray -join $SepChar 

            }
	
		[void]$ValuesArray.add($_eg)

	}


	if ($ValuesArray)
    {
		writeto-Excel -data $ValuesArray -sheetName $SheetName -destworkBook $destWorkBook

	}
}


# ------- Logical Enclosure 
function Export-LogicalEnclosure ($connection,$sheetName, $destWorkbook)
{
	$ValuesArray 					= [System.Collections.ArrayList]::new()
	$enclNames 						= [System.Collections.ArrayList]::new()
	$enclSerialNumbers 				= [System.Collections.ArrayList]::new()
	$frameAddresses 				= [System.Collections.ArrayList]::new()

	
	$LElist 						= get-HPOVLogicalEnclosure -ApplianceConnection $connection

	
	foreach ($LE in $LElist)
	{
		$_le 						= new-object -TypeName le
		$_le.name					= $LE.name
		$enclUris      				= $LE.enclosureUris
		$EncGroupUri   				= $LE.enclosuregroupUri
		$FWbaselineUri 				= $LE.firmware.firmwareBaselineUri
		$_le.forceInstallFirmware	= $LE.firmware.forceInstallFirmware
		$scopesUri     				= $LE.scopesUri

		$EGName        				= Get-NamefromUri -Uri $EncGroupUri -hostconnection $connection
		foreach ($uri in $enclUris)
		{
			$obj 					= Send-HPOVRequest -Hostname $connection -uri $uri 
			[void]$enclNames.Add($obj.name)
			[void]$enclSerialNumbers.Add($obj.serialNumber) 				
		}

		$_le.enclosureGroup 	   	= $EGName
		$_le.enclosureSerialNumber  = $enclSerialNumbers -join $sepChar
		$_le.enclosureName		   	= $enclNames -join $sepChar

		$fwparam = $null
		
		if ($FWbaselineUri)
		{
			$fwName  				= Get-NamefromUri -Uri $FWbaselineUri -hostconnection $connectiom
			$_le.firmwareBaseline 	= $fwName

		}

		# Scopes
		$ResourceScope = Send-HPOVRequest -Uri $scopesUri -Hostname $connection

		if (-not [String]::IsNullOrEmpty($ResourceScope.scopeUris))
		{
			$scopeNames 	= [System.Collections.ArrayList]::new()
			ForEach ($scopeUri in $ResourceScope.scopeUris)
			{
				$scopeName = Get-NamefromUri -Uri $scopeUri -hostconnection $connection
				[void]$scopeNames.Add($scopeName)
			}

		}
		if ($scopeNames)
		{
			$_le.scopes 	= $scopeNames -join $sepChar
		}

		# ------- EBPIA region 
	    $manualAddresses	 			= [System.Collections.ArrayList]::new()
		$enclosures 					= $LE.enclosures		#[]

		for($i=0; $i -lt $enclUris.Count; $i++)
		{
			$_uri 					= $enclUris[$i] 		# get enclosure Uri
            $_encl					= $enclosures.$_uri 

			$deviceBays 			= $_encl.deviceBays
			$interconnectBays 		= $_encl.interconnectBays

			if ($deviceBays -or $interconnectBays)
			{


				# 1 - Get manual address in DeviceBays
				foreach ($_bay in $deviceBays)
				{
					$_ipAddr		= [System.Collections.ArrayList]::new()	
					$_deviceName 		= "Device{0}"	-f $_bay.bayNumber
					$addressArr 	= $_bay.manualAddresses
					foreach ($_addr in $addressArr)
					{
						$_ip 		= "{0}Address='{1}'" -f $_addr.type, $_addr.ipAddress 
						[void]$_ipAddr.add($_ip)
					    $_item 			= "$_deviceName={" + ($_ipAddr -join "$SepHash$CR") + "}"
					    [void]$manualAddresses.add($_item)
					}

				}

				# 2 - Get manual address in InterconnectBays
				foreach ($_ic in $interconnectBays)
				{
                    $_ipAddr		= [System.Collections.ArrayList]::new()	
					$_icName 		= "Interconnect{0}"	-f $_ic.bayNumber
					$addressArr 	= $_ic.manualAddresses
					foreach ($_addr in $addressArr)
					{
						$_ip 		= "{0}Address='{1}'" -f $_addr.type, $_addr.ipAddress 
						[void]$_ipAddr.add($_ip)
					    $_item 			= "$_icName={" + ($_ipAddr -join "$SepHash$CR") + "}"
					    [void]$manualAddresses.add($_item)
					}

				}



				# 3 - Build the ebpia
				if ($manualAddresses)
				{
					$_frame 			= "Frame{0}" -f ($i +1)
					$_addressString 	= ($manualAddresses | sort) -join "$SepHash$CR"
					$_item 				= "$_frame=@{$_addressString}"
					[void]$frameAddresses.Add($_item)

                    
	                $manualAddresses	 			= [System.Collections.ArrayList]::new()
				}

			}

		}

        $frameAddresses = $frameAddresses -join "$Delimiter$CR"


		$_le.manualAddresses		= $frameAddresses

		$valuesArray 		+= $_le

	}


	##
	if ($ValuesArray)
    {
		writeto-Excel -data $ValuesArray -sheetName $SheetName -destworkBook $destWorkBook

	}
}	


Function Export-Server($connection,$sheetName, $destWorkbook)
{
	$InputObject 					= Get-HPOVServer -ApplianceConnection $connection
	$valuesArray 					= [System.Collections.ArrayList]::new()
	foreach ( $s in $InputObject)
	{
		$_s 						= New-Object -TypeName server
		$_s.name					= $s.name
		$_s.serverName				= $s.serverName
		$_s.description				= $s.description
		$_s.model					= $s.model	
		$_s.processorCount			= $s.processorCount	
		$_s.processorCoreCount		= $s.processorCoreCount
		$_s.processorSpeed			= "{0:n2} " -f ($s.processorSpeedMhz	/1000)
		$_s.processorType			= $s.processorType	
		$_s.memory					= $s.memoryMb / 1KB
		$_s.serialNumber			= $s.serialNumber
		$_s.virtualSerialNumber		= $s.virtualSerialNumber
		if ($s.serverProfileUri)
		{
			$_s.profileName			= Get-NamefromUri -uri  $s.serverProfileUri -hostconnection $connection
		}
		#$_s.scopes	
		
		$valuesArray				+= $_s

	}

	##
	if ($ValuesArray)
	{
		writeto-Excel -data $ValuesArray -sheetName $SheetName -destworkBook $destWorkBook

	}
}

Function Export-ServerHardwareType($connection,$sheetName, $destWorkbook)
{
	$InputObject 					= Get-HPOVServerHardwareType -ApplianceConnection $connection
	$valuesArray 					= [System.Collections.ArrayList]::new()
	foreach ( $s in $InputObject)
	{
		$_sht 						= New-Object -TypeName sht
		$_sht.name					= $s.name
		$_sht.formFactor			= $s.formFactor
		$_sht.model					= $s.model
	
		$aModelArr					= [System.Collections.ArrayList]::new()
		$aSlotArr					= [System.Collections.ArrayList]::new()

		$_adapterList 				= $s.adapters
		foreach ($_ad in $_adapterList)
		{
			[void]$aModelArr.Add($_ad.model)
			[void]$aSlotArr.Add($_ad.slot)
		}

		if ($aModelArr)
		{
			$_sht.adapterModel 		= $aModelArr -join $SepChar
		}
		if ($aSlotArr)
		{
			$_sht.adapterSlot 		= $aSlotArr -join $SepChar
		}


		$valuesArray				+= $_sht

	}

	##
	if ($ValuesArray)
	{
		writeto-Excel -data $ValuesArray -sheetName $SheetName -destworkBook $destWorkBook

	}

}

Function Export-Profile($connection,$sheetName, $destWorkbook)
{

	$sptList 						= Get-HPOVServerProfile -ApplianceConnection $connection

	if ($sptList)
	{
		export-profileorTemplate -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook -profList $sptList
	}
}

Function Export-profileTemplate($connection,$sheetName, $destWorkbook)
{

	$sptList 						= Get-HPOVServerProfileTemplate -ApplianceConnection $connection

	if ($sptList)
	{
		export-profileorTemplate -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook -profList $sptList
	}
}

function Export-ProfileorTemplate($connection,$sheetName, $destWorkbook,$profList)
{
	$sheets 						= $sheetname.split($SepChar) 			#profile|Connections|LocalStorage|SANStorage
	$profileSheetName 				= $sheets[0]
	$connectionSheetName 			= $sheets[1]
	$localStorageSheetName			= $sheets[2]
	$SANSheetName					= $sheets[3]
	$ILOSheetName 					= $sheets[4]

	$profileValuesArray 			= [System.Collections.ArrayList]::new()
	$connectionValuesArray 			= [System.Collections.ArrayList]::new()
	$localStorageValuesArray 		= [System.Collections.ArrayList]::new()
	$sanStorageValuesArray 			= [System.Collections.ArrayList]::new()
	$profileILOValuesArray 			= [System.Collections.ArrayList]::new()




	foreach ($InputObject in $profList)
	{
		$isSpt 	= $InputObject.category -eq 'server-profile-templates'
		$isSp	= $InputObject.category -eq 'server-profiles'

		if ($isSpt)
			{ $_sp 			= new-object -TypeName spt }
		else 
			{ $_sp  		= new-object -TypeName sp  }

		$name        	    = $InputObject.Name   
		$description    	= $InputObject.Description 
		$spDescription      = $InputObject.serverprofileDescription 
		$scopesUri          = $InputObject.scopesUri
		
		$shtUri             = $InputObject.serverHardwareTypeUri
		$egUri              = $InputObject.enclosureGroupUri
		$sptUri             = $InputObject.serverProfileTemplateUri
		$serverUri          = $InputObject.serverHardwareUri
		$enclosureUri       = $InputObject.enclosureUri
		$enclosureBay       = $InputObject.enclosureBay
		$affinity           = $InputObject.affinity 

		$fw          		= $InputObject.firmware
		$isFwManaged 		= $fw.manageFirmware
		$fwInstallType  	= $fw.firmwareInstallType
		$fwForceInstall 	= $fw.forceInstallFirmware
		$fwActivation   	= $fw.firmwareActivationType
		$fwSchedule     	= $fw.firmwareScheduleDateTime
		$fwBaseUri      	= $fw.firmwareBaselineUri

		$ConnectionSettings = $InputObject.connectionSettings

		$localStorage 		= $InputObject.localStorage
		$controllerList 	= $localStorage.controllers
		$JBODList 			= $localStorage.sasLogicalJBODs

		$sanStorage 		= $InputObject.sanStorage


		$bm                 = $InputObject.bootMode
		$isManageBootMode 	= $bm.manageMode
		$bo                 = $InputObject.boot
		$isManageBootOrder  = $bo.manageBoot
		
		$bios 				= $InputObject.bios
		$isManageBios		= $bios.manageBios
		$biosSettings 		= $bios.overriddenSettings

		$ilo 				= $InputObject.managementProcessor
		$isManageILO 		= $ilo.manageMp
		$iloSettings 		= $ilo.mpSettings

		$hideFlexNics       = $InputObject.hideUnusedFlexNics

		$macType            = $InputObject.macType
		$wwnType            = $InputObject.wwnType
		$snType             = $InputObject.serialNumberType       
		$iscsiType          = $InputObject.iscsiInitiatorNameType 

		if ($isSp)
		{
			$sn 			= $InputObject.serialNumber
			$iscsiName 		= $InputObject.iscsiInitiatorName
		}

		$osdeploysetting    = $InputObject.osDeploymentSettings
		
		# --------- Consistency Check only on SPT
		if ($isSpt)
		{
			$_sp.firmwareConsistencyChecking  	= if ($fw.complianceControl) 					{$consistencyCheckingEnum.Item($fw.complianceControl) } else {'None'}
			$_sp.bootOrderConsistencyChecking  	= if ($bo.complianceControl) 					{$consistencyCheckingEnum.Item($bo.complianceControl) } else {'None'}
			$_sp.bootModeConsistencyChecking  	= if ($bm.complianceControl) 					{$consistencyCheckingEnum.Item($bm.complianceControl) } else {'None'}
			$_sp.biosConsistencyChecking  		= if ($bios.complianceControl)					{$consistencyCheckingEnum.Item($bios.complianceControl) } else {'None'}
			$_sp.iloConsistencyChecking			= if ($ilo.complianceControl)					{$consistencyCheckingEnum.Item($ilo.complianceControl) } else {'None'}
			$_sp.connectionConsistencyChecking  = if ($connectionSettings.complianceControl)	{$consistencyCheckingEnum.Item($connectionSettings.complianceControl) } else {'None'}
			$_sp.localStorageConsistencyChecking = if ($localStorage.complianceControl)			{$consistencyCheckingEnum.Item($localStorage.complianceControl) } else {'None'}
			$_sp.sanStorageConsistencyChecking   = if ($sanStorage.complianceControl)			{$consistencyCheckingEnum.Item($sanStorage.complianceControl) } else {'None'}
		
		}
		# --------- General Section
		$_sp.name 			= $name
		$_sp.description	= $description

		if (-not [String]::IsNullOrEmpty($scopesUri) )
		{
			$ResourceScope 	= Send-HPOVRequest -Uri $scopesUri -Hostname $connection
		}
		if (-not [String]::IsNullOrEmpty($ResourceScope.scopeUris))
		{
			$scopeNames 	= [System.Collections.ArrayList]::new()
			ForEach ($scopeUri in $ResourceScope.scopeUris)
			{
				$scopeName = Get-NamefromUri -Uri $scopeUri -hostconnection $connection
				[void]$scopeNames.Add($ScopeName)
			}
			$_sp.Scopes 	= if ($ScopeNames) {$scopeNames -join $SepChar } else {$NULL}
		}


		# ---------- Server Profile Section
		if ($isSpt)
			{ $_sp.serverProfileDescription = $spDescription}
		# Get SHT
		$_sp.sht 				= Get-NamefromUri -uri $shtUri -hostconnection $connection
		
		# Get enclosure group
		$_sp.enclosureGroupName	= Get-NamefromUri -uri $egUri -hostconnection $connection

		# Get affinity
		$_sp.affinity 			= $affinity


		# ------- Server hardware
		if (-not [String]::IsNullOrWhiteSpace($serverUri))
			{ $_sp.serverHardware = Get-NamefromUri -uri $serverUri -hostconnection $connection }

		# --------- Server profile Template
		if (  $isSp -and (-not [String]::IsNullOrWhiteSpace($sptUri)) )
			{ $_sp.serverProfileTemplate = Get-NamefromUri -uri $sptUri -hostconnection $connection }



		# ----------- firmware section
		$_sp.manageFirmware		=  $isFwManaged
		if ($isFWmanaged)
		{
			$_sp.firmwareBaselineName			= Get-NamefromUri -Uri $fwBaseUri -hostconnection $connection
			$_sp.firmwareInstallType			= $fwInstallType
			$_sp.forceInstallFirmware			= $fwForceInstall
			$_sp.firmwareActivationType			= $fwActivation
			if ($fwActivation -eq 'Scheduled')
				{ $_sp.firmwareSchedule			= ([DateTime]$fwSchedule).ToString() }
		
		}

		# ----------- Connections section
		if ($isSpt)
		{
			$_sp.manageConnections 	= $ConnectionSettings.manageConnections
		}
	
		$connections 			= $connectionSettings.Connections
		foreach ($conn in $connections)
		{
			
			$_conn				= new-object -TypeName connection
			$netUri             = $conn.networkUri
			$bootSettings       = $conn.boot
			
			$_conn.profileName 	= $_sp.name
			$_conn.name 		= $conn.name
			$_conn.id           = $conn.id
			$_conn.functionType = $conn.functionType
			$_conn.network 		= Get-NamefromUri -uri $netUri -hostconnection $connection
			$_conn.portId       = $conn.portID
			$_conn.requestedMbps= $conn.requestedMbps
			$_conn.requestedVFs = $Conn.requestedVFs
			$_conn.lagName 		= $conn.lagName

			$priority 			= $bootSettings.priority
			$bootVolumeSource 	= $bootSettings.bootVolumeSource

			if ($bootVolumeSource -eq 'UserDefined')		#HKD02
			{
				$targetArray 		= [System.Collections.ArrayList]::new()	#HKD03
				foreach ($t in $bootSettings.targets)
				{   
					$targetWWpn 	= $t.arrayWwpn -replace '(..)(?=.)','$1:' # convert to WWN format
					$s 				= '@{ ' + 'arrayWwpn = {0} ; lun = {1} '  -f $targetWWpn , $t.lun + '}'
					[void]$targetArray.Add($s)
				}
				if ($targetArray)
				{
					$_conn.targets 	= $targetArray -join $SepChar
				}

			}
			$_conn.bootVolumeSource	= $bootVolumeSource

			$_conn.boot 		= -not($priority -eq 'NotBootable')
			$_conn.priority 	= $priority 

			

			if ($isSp)
			{
				$_conn.macType 	= $Conn.macType
				$_conn.wwpnType	= $Conn.wwpnType
				$_conn.mac	 	= $Conn.mac
				$_conn.wwpn		= $Conn.wwpn
				$_conn.wwnn		= $Conn.wwnn
			}

			[void]$connectionValuesArray.Add($_conn)

		}		
					
		
		# ----------- local Storage section
		foreach ($c in $controllerList)
		{
			$_spc 					= new-object -TypeName spLocalStorage
			$_spc.profileName		= $_sp.name
			$_spc.deviceSlot 		= $c.deviceSlot
			$_spc.mode				= $c.mode
			$_spc.initialize 		= $c.initialize
			$_spc.driveWriteCache	= $c.driveWriteCache
					
			$ldDrives 				= $c.logicalDrives
			$ldDriveNames = $ldBoot = $ldRaidLevel = $ldDriveTechnology = $ldPhysicalDrives = $ldAccelerator = [System.Collections.ArrayList]::new()
			
			foreach ($ld in $ldDrives)
			{
				$ldDriveNames 		+= $ld.name
				$ldRaidLevel		+= $ld.raidLevel
				$ldBoot				+= $ld.bootable
				$ldPhysicalDrives	+= $ld.numPhysicalDrives
				$ldDriveTechnology	+= if ($ld.driveTechnology ) { $ld.driveTechnology.replace('Hdd','') } else {'Auto'}
				$ldAccelerator		+= $ld.accelerator
			}
			$_spc.logicalDrives		= if ($ldDriveNames.Count -gt 0)		{$ldDriveNames 		-join '|'}		else {''}
			$_spc.raidLevel 		= if ($ldRaidLevel.Count -gt 0)			{$ldRaidLevel 		-join '|'}		else {''}
			$_spc.bootable 			= if ($ldBoot.Count -gt 0)				{$ldBoot			-join '|'}		else {''}
			$_spc.numPhysicalDrives = if ($ldPhysicalDrives.Count -gt 0)	{$ldPhysicalDrives	-join '|'}		else {''}	
			$_spc.driveTechnology 	= if ($ldDriveTechnology.Count -gt 0)	{$ldDriveTechnology	-join '|'}		else {''}
			$_spc.accelerator 		= if ($ldaccelerator.Count -gt 0)		{$ldaccelerator		-join '|'}		else {''}
				
			[void]$localStorageValuesArray.Add($_spc)
		}
		
		foreach ($sas in $JBODList)
		{
			$_jbod							= new-object -TypeName spLocalStorage
			$_jbod.profileName				= $_sp.name 
			$_jbod.deviceSlot 				= $sas.deviceSlot
			$_jbod.id 						= $sas.id
			$_jbod.logicalDrives	 		= $sas.name
			$_jbod.description				= $sas.description
			$_jbod.driveTechnology 			= if ($sas.driveTechnology ) { $sas.driveTechnology.replace('Hdd','') } else {'Auto'}
			$_jbod.numPhysicalDrives		= $sas.numPhysicalDrives
			$_jbod.driveMinSize 			= "{0:n2}" -f $sas.driveMinSizeGB
			$_jbod.driveMaxSize	 			= "{0:n2}" -f $sas.driveMaxSizeGB
			$_jbod.eraseData 				= $sas.eraseData
			$_jbod.persistent				= $sas.persistent
															
			[void]$localStorageValuesArray.Add($_jbod)
		}

		# ----------- SAN Storage section
		if ($sanStorage.manageSanStorage)
		{
			$_sp.hostOStype 						= $ServerProfileSanManageOSType.GetEnumerator().where({$_.value -eq $sanStorage.hostOStype}).Name
			$_sp.manageSanStorage 					= $sanStorage.manageSanStorage
			$_volAttachments 						= $sanStorage.volumeAttachments
			foreach ($_vol in $_volAttachments)
			{
				$_sanStorage 						= new-object -type spSANStorage
				$_sanStorage.profileName			= $_sp.name		
				
				$_sanStorage.volumeID 				= $_vol.id
				$_sanStorage.volumeName 			= get-NamefromUri -uri  $_vol.volumeUri -hostconnection $connection
				$_lunType 							= $_vol.LunType 
				$_sanStorage.volumeLUNType			= $_lunType
				if ($_lunType -eq 'Manual')
				{
					$_sanStorage.volumeLUN 			= $_vol.lun
				}
				$_sanStorage.volumeStorageSystem	= get-NamefromUri -uri  $_vol.volumeStorageSystemUri -hostconnection $connection
				$_sanStorage.volumeBootPriority		= $_vol.bootVolumePriority	
				$storagePaths 						= $_vol.storagepaths
				$pathsArray							= [System.Collections.ArrayList]::new()

				foreach ($_path in $storagePaths)
				{
					$_netname 						= Get-NamefromUri -uri $_path.networkUri -hostconnection $connection
					#$_pathString 					= '@{' + 'connectionId = {0}; network = "{1}"; isEnabled = {2}' -f $_path.connectionId, $_netname, $_path.isEnabled.ToString()
					$_pathString 					= '@{' + 'connectionId = {0}; isEnabled = ${1}' -f $_path.connectionId, $_path.isEnabled.ToString()
					$_pathString					+= '}'
					[void]$pathsArray.Add($_pathString)
				}
				if ($pathsArray)
				{
					$_sanStorage.volumeStoragePaths	= $pathsArray -join $SepChar 
				}

				[void]$sanStorageValuesArray.Add($_sanStorage)
			}

			
			
		}



		# ----------- boot  section

		$_sp.manageBootMode = $ismanageBootMode
		$_sp.manageBootOrder = $isManageBootOrder

		if ($isManageBootMode)
		{
			$_sp.mode 			= $bm.mode
			$_sp.pxeBootPolicy  = $bm.pxeBootPolicy
			$_sp.secureBoot     = $bm.secureBoot
			if ($isManageBootOrder)
			{
				$_sp.order 		= $bo.order -join $SepChar
			}
		}

		# ---------------- BIOS section
		$_sp.manageBios 	= $IsManageBios
		$settingArray 		= [System.Collections.ArrayList]::new()

		if ($isManageBios -and $biosSettings)
		{
			foreach ($setting in $biosSettings)
			{
				#$s 			= $setting.id + '=' + $setting.value
				$s 			= '@{ ' + 'id = "{0}"; value = "{1}"' -f $setting.id , $setting.value + '}'
				[void]$settingArray.Add($s)
			}
			$_sp.overriddenSettings	= $settingArray -join $SepChar
		}

		# ---------------- iLO section
		$_sp.manageIlo 	= $isManageiLO
		if ($isManageILO)
		{                                                   																																																																																														
			foreach ($s in $iloSettings)
			{

				switch ($s.settingType)
				{
					'AdministratorAccount'
						{ 
							$_ilo 								= new-object -TypeName ilosetting
							$_ilo.profileName 					= $_sp.name
							$_ilo.profileType 					= if ( $isSpt ) {'ProfileTemplate'} else {'Profile'}
							$_ilo.settingType 					= $s.settingType
							$_ilo.deleteAdministratorAccount	= $s.args.deleteAdministratorAccount
							$_ilo.adminPassword 				= $s.args.password

							$profileILOValuesArray				+= $_ilo
						}

					'LocalAccounts'
						{
							foreach ($account in  $s.args.localAccounts)
							{
								$privs 							= [System.Collections.ArrayList]::new()

								$_ilo 							= new-object -TypeName ilosetting
								$_ilo.profileName 				= $_sp.name
								$_ilo.profileType 				= if ( $isSpt ) {'ProfileTemplate'} else {'Profile'}
								$_ilo.settingType 				= $s.settingType

								$_ilo.userName 					= $account.userName
								$_ilo.displayName 				= $account.displayName
								$_ilo.userPassword 				= '***REDACTED***'
								
								if ($account.userConfigPriv)
									{	[void]$privs.Add('userConfigPriv') }
								
								if ($account.remoteConsolePriv)
									{	[void]$privs.Add('remoteConsolePriv') }
								
								if ($account.virtualMediaPriv)
									{	[void]$privs.Add('virtualMediaPriv') }
								
								if ($account.iLOConfigPriv)
									{	[void]$privs.Add('iLOConfigPriv') }
								
								if ($account.virtualPowerAndResetPriv)
									{	[void]$privs.Add('virtualPowerAndResetPriv') }
								
								if ($account.loginPriv)
									{	[void]$privs.Add('loginPriv') }

								if ($account.hostBIOSConfigPriv)
									{	[void]$privs.Add('hostBIOSConfigPriv') }

								if ($account.hostNICConfigPriv)
									{	[void]$privs.Add('hostNICConfigPriv') }

								if ($account.hostStorageConfigPriv)
									{	[void]$privs.Add('hostStorageConfigPriv') }

								$_ilo.userPrivileges 					= $privs -join $SepChar

								$profileILOValuesArray				+= $_ilo
							}
						}

					'Directory'
						{
							$_ilo 								= new-object -TypeName ilosetting
							$_ilo.profileName 					= $_sp.name
							$_ilo.profileType 					= if ( $isSpt ) {'ProfileTemplate'} else {'Profile'}
							$_ilo.settingType 					= $s.settingType

							$args  								= $s.args
							
							$dirAuth 							= $args.directoryAuthentication -replace 'defaultSchema',  'DirectoryDefault'
							$dirAuth 							= $dirAuth -replace 'extendedSchema', 'HPEExtended' 
							$dirAuth 							= $dirAuth -replace 'disabledSchema' , 'Disabled'
							$_ilo.directoryAuthentication		= $dirAuth
							if ($dirAuth -ne 'Disabled')
							{
								$_ilo.iloObjectDistinguishedName	= $args.iloObjectDistinguishedName
								$_ilo.directoryPassword 			= '***REDACTED***'
								$_ilo.directoryServerAddress		= $args.directoryServerAddress
								$_ilo.directoryServerPort			= $args.directoryServerPort
								$_ilo.directoryUserContext			= $args.directoryUserContext -join $SepChar
							}

							$_ilo.kerberosAuthentication		= $args.kerberosAuthentication
							$_ilo.kerberosRealm					= $args.kerberosRealm
							$_ilo.kerberosKDCServerAddress		= $args.kerberosKDCServerAddress
							$_ilo.kerberosKDCServerPort			= $args.kerberosKDCServerPort
							$_ilo.kerberosKeytab				= $args.kerberosKeytab -replace $CR, '`n'

							$profileILOValuesArray				+= $_ilo
						}

					'DirectoryGroups'
						{
							foreach ($g in $s.args.directoryGroupAccounts)
							{
								$privs 							= [System.Collections.ArrayList]::new()

								$_ilo 							= new-object -TypeName ilosetting
								$_ilo.profileName 				= $_sp.name
								$_ilo.profileType 				= if ( $isSpt ) {'ProfileTemplate'} else {'Profile'}
								$_ilo.settingType 				= $s.settingType

								$_ilo.groupDN 					= $g.groupDN
								$_ilo.groupSID 					= $g.groupSID

								if ($g.userConfigPriv)
									{	[void]$privs.Add('userConfigPriv') }
								
								if ($g.remoteConsolePriv)
									{	[void]$privs.Add('remoteConsolePriv') }
								
								if ($g.virtualMediaPriv)
									{	[void]$privs.Add('virtualMediaPriv') }
								
								if ($g.iLOConfigPriv)
									{	[void]$privs.Add('iLOConfigPriv') }
								
								if ($g.virtualPowerAndResetPriv)
									{	[void]$privs.Add('virtualPowerAndResetPriv') }
								
								$_ilo.groupPrivileges 			= $privs -join $SepChar

								$profileILOValuesArray			+= $_ilo
							}
						}

					'HostName'
						{
							$_ilo 								= new-object -TypeName ilosetting
							$_ilo.profileName 					= $_sp.name
							$_ilo.profileType 					= if ( $isSpt ) {'ProfileTemplate'} else {'Profile'}
							$_ilo.settingType 					= $s.settingType
							$args  								= $s.args
							$_ilo.HostName						= $args.HostName

							$profileILOValuesArray				+= $_ilo
						}

					'KeyManager'
						{
							$_ilo 								= new-object -TypeName ilosetting
							$_ilo.profileName 					= $_sp.name
							$_ilo.profileType 					= if ( $isSpt ) {'ProfileTemplate'} else {'Profile'}
							$_ilo.settingType 					= $s.settingType
							$args  								= $s.args

							$_ilo.primaryServerAddress			= $args.primaryServerAddress
							$_ilo.primaryServerPort				= $args.primaryServerPort
							$_ilo.secondaryServerAddress		= $args.secondaryServerAddress
							$_ilo.secondaryServerPort			= $args.secondaryServerPort
							$_ilo.redundancyRequired			= $args.redundancyRequired
							$_ilo.groupName						= $args.groupName
							$_ilo.certificateName				= $args.certificateName
							$_ilo.loginName						= $args.loginName
							$_ilo.keyManagerPassword 			= '***REDACTED***'

							$profileILOValuesArray				+= $_ilo
						}

				} # end switch

			}
		}

	

		# ---------------- Advanced section
		if ($isSp )
		{
			$_sp.serialNumber		= $sn
			$_sp.iscsiInitiatorName	= $iscsiName
		}

		$_sp.macType 				= $macType
		$_sp.wwnType 				= $wwnType
		$_sp.serialNumberType		= $snType
		$_sp.iscsiInitiatorNameType	= $iscsiType

		$_sp.hideUnusedFlexNics 	= $hideUnusedFlexNics

		# Add to list
		[void]$profileValuesArray.Add($_sp)
    }



	

		### write to Excel
	if ($profileValuesArray)
    {
		writeto-Excel -data $profileValuesArray -sheetName $profileSheetName -destworkBook $destWorkBook

	}

	if ($connectionValuesArray)
    {
		writeto-Excel -data $connectionValuesArray -sheetName $connectionSheetName -destworkBook $destWorkBook

	}

	if ($localstorageValuesArray)
    {
		writeto-Excel -data $localstorageValuesArray -sheetName $localStorageSheetName -destworkBook $destWorkBook

	}

	if ($SANstorageValuesArray)
    {
		writeto-Excel -data $SANstorageValuesArray -sheetName $SANSheetName -destworkBook $destWorkBook

	}

	if ($profileILOValuesArray)
    {
		writeto-Excel -data $profileILOValuesArray -sheetName $ILOSheetName -destworkBook $destWorkBook

	}


}



# -------------------------------------------------------------------------------------------------------------
#
#       Main Entry
#
# -------------------------------------------------------------------------------------------------------------


# ---------------- Connect to Synergy Composer
#
if ($jsonConfigFiles)
{
	$connectionList = Connect-Composers -jsonConfigs $jsonConfigFiles 
}
else 
{
	write-host -ForegroundColor Yellow 'No OV config file specified. Exiting...'	
	exit
}

# ---------------- Define Excel files
#
$startRow 				= 15
$ExcelTemplate			= "OV-Template.xlsx"
if (test-path $ExcelTemplate)
{
	foreach ($connection in $connectionList)
	{
		$ip 			= $connection.name
		write-host -ForegroundColor Cyan "----- Connecting to OneView --> $ip"
		$dest 			= $connection.Name
		$destWorkbook 	= "ExportFrom-$dest.xlsx"
		Copy-Item -path $ExcelTemplate -destination $destWorkbook 

		#----------------------------------------------
		#              OV appliance Settings
		#----------------------------------------------

		# ---- Export OV network
		write-host -ForegroundColor Cyan "--------- Exporting OneView networking"
		$sheetName  	 = 'OVnetwork'
		export-HPOVnetwork -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook

		# ---- Export OV security protocol
		write-host -ForegroundColor Cyan "--------- Exporting OneView security protocol"
		$sheetName  	 = 'OVsecurityProtocol'
		export-HPOVsecurityProtocol -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook

		# ---- Export OV security authentication
		write-host -ForegroundColor Cyan "--------- Exporting OneView security authentication"
		$sheetName  	 = 'OVauthentication'
		export-HPOVauthentication  -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook
		
		#----------------------------------------------
		#              OV Settings
		#----------------------------------------------

		# ---- Export scopes
		write-host -ForegroundColor Cyan "--------- Exporting Scopes"
		$sheetName  	 = 'scope'
		Export-Scopes -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook

		# ---- Export time and locale
		write-host -ForegroundColor Cyan "--------- Exporting Time and Locale"
		$sheetName  	= 'timeLocale'
		Export-TimeLocale -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook		

		# ---- Export SMTP
		write-host -ForegroundColor Cyan "--------- Exporting SMTP"
		$sheetName  	= 'smtp'
		Export-SMTP -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook
		
		# ---- Export firmware baseline
		write-host -ForegroundColor Cyan "--------- Exporting Firmware Bundle"
		$sheetName  	= 'firmwareBundle'
		Export-fwBaseline -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook	

		# ---- Export Backup Config
		write-host -ForegroundColor Cyan "--------- Exporting backup Configuration"
		$sheetName  	= 'backupConfig'
		Export-BackupConfig -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook	
		
		# ---- Export baseline repository
		write-host -ForegroundColor Cyan "--------- Exporting baseline Repository"
		$sheetName  	= 'repository'
		Export-repository -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook

		# ---- Export Address pool range
		write-host -ForegroundColor Cyan "--------- Exporting Address Pool"
		$sheetName  	= 'addressPool'
		Export-AddressPoolRange -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook	
		
		# ---- Export snmp Configuration
		write-host -ForegroundColor Cyan "--------- Exporting snmp Configuration"
		$sheetName  	= 'snmpConfiguration'
		Export-snmpConfiguration -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook	

		# ---- Export snmp V3 Users
		write-host -ForegroundColor Cyan "--------- Exporting snmp V3 users"
		$sheetName  	= 'snmpV3User'
		Export-snmpUsers -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook	
		
		# ---- Export snmp traps
		write-host -ForegroundColor Cyan "--------- Exporting snmp traps"
		$sheetName  	= 'snmpTrap'
		Export-snmpTraps -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook	
		
		#----------------------------------------------
		#              OV Resources
		#----------------------------------------------

		# ---- Export Network
		write-host -ForegroundColor Cyan "--------- Exporting Ethernet/FC/FCOE Networks"
		$sheetName  	= 'ethernetNetwork|fcNetwork'
		Export-Network -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook	

		# ---- Export NetworkSet
		write-host -ForegroundColor Cyan "--------- Exporting Network Sets"
		$sheetName  	= 'networkSet'
		Export-NetworkSet -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook	

		# ---- Export Logical Interconnect Group
		write-host -ForegroundColor Cyan "--------- Exporting Logical Interconnect Group"
		$sheetName  	= 'logicalInterconnectGroup|UplinkSet|ligSnmp'
		Export-LogicalInterconnectGroup -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook	
		


		# ------------------Storage region
		#
		# ---- Export StorageSystem
		write-host -ForegroundColor Cyan "--------- Exporting Storage System"
		$sheetName  	= 'storageSystem'
		Export-StorageSystem -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook	

		# ---- Export StoragePool
		write-host -ForegroundColor Cyan "--------- Exporting Storage Pool"
		$sheetName  	= 'storagePool'
		Export-StoragePool -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook	
				
		# ---- Export Storage Volume Template
		write-host -ForegroundColor Cyan "--------- Exporting Storage volume template"
		$sheetName  	= 'storageVolumeTemplate'
		Export-StorageVolumeTemplate -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook	

		# ---- Export Storage Volume
		write-host -ForegroundColor Cyan "--------- Exporting Storage Volume"
		$sheetName  	= 'storageVolume'
		Export-StorageVolume -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook	

		# ---- Export logical JBOD
		write-host -ForegroundColor Cyan "--------- Exporting logical JBOD"
		$sheetName  	= 'logicalJBOD'
		Export-logicalJBOD -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook	
		
		# ------------------Enclosure region
		#
		write-host -ForegroundColor Cyan "--------- Exporting Enclosure Group"
		$sheetName  	= 'enclosureGroup'
		Export-EnclosureGroup -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook
		
		# ---- Export Logical Enclosure
        write-host -ForegroundColor Cyan "--------- Exporting Logical Enclosure"
		$sheetName  	= 'logicalEnclosure'
		Export-LogicalEnclosure -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook

		# ---- Export Server
		write-host -ForegroundColor Cyan "--------- Exporting server hardware"
		$sheetName  	= 'server'
		Export-Server -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook

		# ---- Export ServerHardwareType (SHT)
		write-host -ForegroundColor Cyan "--------- Exporting server hardware type"
		$sheetName  	= 'serverHardwareType'
		Export-ServerHardwareType -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook

		# ---- Export profile template
        write-host -ForegroundColor Cyan "--------- Exporting server profile template"
		$sheetName  	= 'profileTemplate|profileTemplateConnection|profileTemplatelocalStorage|profileTemplateSANStorage|profileTemplateILO'
		Export-profileTemplate -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook

		# ---- Export profile
        write-host -ForegroundColor Cyan "--------- Exporting server profile"
		$sheetName  	= 'profile|profileConnection|profilelocalStorage|profileSANStorage|profileILO'
		Export-profile -connection $connection -sheetName $sheetName -destWorkBook $destWorkbook

		Disconnect-HPOVMgmt -ApplianceConnection $connection
	}
}
else
{
	write-host -foreground Yellow "Cannot find $ExcelTemplate --> Exiting script"
	exit
}

